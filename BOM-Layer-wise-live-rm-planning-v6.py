import requests
import pandas as pd
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.styles import Border, Side, Alignment, PatternFill, Font
import logging
import warnings
import time
import os
import sys
from openpyxl import load_workbook
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import msal

if sys.platform == "win32":
    import os
    os.environ["PYTHONIOENCODING"] = "utf-8"

warnings.filterwarnings('ignore')

# --- Logging Setup ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(message)s',
    handlers=[
        logging.FileHandler('bom_layerwise.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# --- Configuration ---
SAP_API_URL = "https://my409486-api.s4hana.cloud.sap/sap/opu/odata/sap/API_BILL_OF_MATERIAL_SRV;v=0002"
USERNAME = (os.environ.get("SAP_USERNAME") or "").strip()
PASSWORD = (os.environ.get("SAP_PASSWORD") or "").strip()
if not USERNAME or not PASSWORD:
    raise RuntimeError(
        "SAP_USERNAME / SAP_PASSWORD not set - add them as GitHub repository secrets"
    )
VERIFY_SSL = True
PREFERRED_FORMAT = 'json'
BATCH_SIZE = 500
MAX_LAYERS = 4

# --- v6: keep 'RM Planning SAP New.xlsx' small enough for Excel Online ---
# Until v5 the stock sheet held every plant / storage location / batch row (~176k rows), which
# together with BOM Layer-wise pushed the workbook past what Excel Online can open within its
# timeout (createSession and even single-row writes started returning 504 on 24-Oct-26).
# Stock is now filtered in SAP to the plants below (1000 for RM planning, 2000 so FG-RM can
# later see it too) and summed per line, with Batch dropped. Each plant keeps its own rows.
# The column names in 'MB52 Raw Data' stay the same so the 'MB52 1000 Stock' pivot (filtered
# to Plant 1000) and anything else reading the sheet keep working.
STOCK_PLANTS = ["1000", "2000"]
STOCK_GROUP_BY = ['Material', 'Plant', 'StorageLocation', 'InventorySpecialStockType',
                  'InventoryStockType', 'MaterialBaseUnit']
STOCK_BATCH_SIZE = 5000
PRODUCT_API_URL = "https://my409486-api.s4hana.cloud.sap/sap/opu/odata/sap/API_PRODUCT_SRV"

import re

# --- Material group code -> description (static list supplied by the plant) ---
_MATERIAL_GROUP_TEXT = """
1101|Brass parts/fittings
1102|Copper tube
1103|Aluminium tube
1104|Aluminium foil
1105|Sheet Metal
1106|Rubber Plug
1107|Distributor
1108|Components
1109|Tripod, U Bend
1110|Brazing rings
1111|Header Aluminium
1201|Tooling spares
1202|Mech-Spares
1203|Elect-Spares
1204|Hydraulic Spares
1205|Pnuematic Spares
1206|Maintenance consumab
1207|Jigs & fixtures
1301|Chemicals, Paints
1302|Lubricants all types
1303|Industrial Gases
1304|Brazing rings, rods
1305|Stationary
1306|IT Consumables
1307|Maintenance consumab
1308|Factory consumables
1401|Plant Machinery, Equ
1402|Furniture & Fixtures
1403|Vehicles
1404|Office Equipement
1405|Building
1406|Software
1501|Admin & IT Service
1502|Material Handling Se
1503|Waste disposal servi
1504|Technical Services
1505|Professional fees
1506|Freight & forwarding
1601|Aluminum Scrap
1602|Copper Scrap
1603|GI and MS Scrap
1604|Brass scrap
1605|X
1606|Empty Drums
1607|SS Scrap
1701|Endplate
1702|Copper hair pins
1703|Aluminium hair pins
1704|Expanded Block
1705|Header Copper
1706|Header Aluminium
1707|Shroud and Bracket
1708|Return Bend
1709|Finished Component
1710|Subcontracted
1711|Combi EB
1712|Combi FC
1801|Outer Packing Materi
1802|Wooden Material
1803|Packing consumables
1804|Filler material
1901|Evaporator
1902|Condenser
1903|Heater
1904|Combi
1905|Assembly
1906|Sales Kit
2001|Product Inquiry
2101|Header Assy Copper
2102|Header Assy Aluminiu
2201|Coil scrap internal
2202|Coil scrap external
A001|On Account Billed
A002|On Account Utilized
B001|One Time Credit
B002|Cloud Revenue
E001|Accommodation
E002|Airfare
E003|Ground Transp.
E004|Travel Expenses misc
L001|Trading Materials
L002|Raw Materials
L003|Semi-Finished Goods
L004|Finished Goods
L005|Unvaluated Materials
L006|Returnable Packaging
L007|Licenses
L008|Hardware
P000|Contract Type
P001|Services
P002|Expenses
P003|Periodic Service
YBFA01|Real Estate (Land)
YBFA02|Buildings
YBFA03|Land Improvements
YBFA04|Leasehold Improvemen
YBFA05|Machinery Equipment
YBFA06|Fixtures Fittings
YBFA07|Vehicles
YBFA08|Computer Hardware
YBFA09|Computer Software
YBFA10|Low Value Assets
YBFA11|Other Intangibles
YBFA12|Office Equipment
YBMM00|Non-Sto Mat. w.o ID
YBMM01|Non-Sto Mat. w. ID
YBPM01|Spare parts
YBPM02|Spare parts services
"""
MATERIAL_GROUP_DESC = {}
for _line in _MATERIAL_GROUP_TEXT.strip().splitlines():
    _code, _desc = _line.split("|", 1)
    MATERIAL_GROUP_DESC[_code.strip().upper()] = _desc.strip()

def material_group_desc(code):
    """Description for a material group code; blank if the code is blank or not in the list."""
    return MATERIAL_GROUP_DESC.get(str(code).strip().upper(), "")

def is_foil_component(desc, material_group):
    """True for an aluminium foil line (material group 1104 or an 'AL FOIL' description), not the cut-piece scrap line."""
    d = str(desc).upper().strip()
    if re.search(r"CUT ?PIECE", d):
        return False
    return str(material_group).strip() == "1104" or bool(re.match(r"^(AL|ALUMINI?UM)\s+FOIL", d))

def _fmt_num(v):
    if v is None:
        return ""
    return str(int(v)) if float(v).is_integer() else str(round(float(v), 3))

# --- Fin-pack text parser (handles the messy typed text on the foil item text) ---

def _num(m):
    if not m:
        return None
    v = next((g for g in m.groups() if g), None)
    try:
        return float(v) if v is not None else None
    except ValueError:
        return None

_RX_FPI = re.compile(r"FPI[-\s]*0?(\d+(?:\.\d+)?)|(\d+(?:\.\d+)?)\.?FPI", re.I)
_RX_FIN = re.compile(r"(?<![A-Z])FINS?[-\s]*(\d+(?:\.\d+)?)|(\d+(?:\.\d+)?)FINS?", re.I)
_RX_FL = re.compile(r"F/?L[-\s]*(\d+(?:\.\d+)?)", re.I)

_EDGE_RIPPLE = re.compile(r"RI+P\w*")
_EDGE_STRAIGHT = re.compile(r"STR[AI]*G?[IH]*T\w*|STRAIGH\w*|STRAGIHT|STRIGHT")


def _classify_surface(s):
    s = re.sub(r"[^A-Z0-9 ]", " ", s.upper())
    s = re.sub(r"\s+", " ", s).strip()
    if not s:
        return ""
    if "TSUNAMI" in s:
        return "Tsunami"
    lanced = re.search(r"L[AE]NC", s)
    raised = re.search(r"RAI", s) or re.search(r"(^|\s)R\s?(?=L[AE]NC)", s) or (re.search(r"RAIS", s))
    louvered = re.search(r"L[O0]U?V|LOUER|LOVER", s)
    flat = re.search(r"\bFLAT\b|\bFLA\b|FLAT", s)
    if raised and (lanced or re.search(r"\bL\b|\bLA\b", s)):
        return "Raised lanced"
    if lanced:
        return "Flat lanced" if flat else "Lanced"
    if louvered:
        return "Flat louvered" if flat else "Louvered"
    if re.search(r"SIN[EW]|\bSW\b|WAVY", s):
        return "Sinewave"
    if re.search(r"COR+U|\bCR\b", s):
        return "Corrugated"
    if flat:
        return "Flat"
    return ""


def parse_fin_spec(text):
    """Return dict: fpi, fins, length_mm, surface, edge, check. Never raises."""
    out = {"fpi": None, "fins": None, "length_mm": None, "surface": "", "edge": "", "check": ""}
    raw = "" if text is None else str(text).strip()
    if not raw:
        out["check"] = "No fin text"
        return out
    try:
        s = raw.upper()
        s = re.sub(r"\.{2,}", ".", s)                       # 122..16 -> 122.16
        s = re.sub(r"\([^)]*\)", " ", s)                     # (-0.5)
        out["fpi"] = _num(_RX_FPI.search(s))
        out["fins"] = _num(_RX_FIN.search(s))
        out["length_mm"] = _num(_RX_FL.search(s))

        # remove the numeric spec so only the pattern words are left
        t = _RX_FPI.sub(" ", s)
        t = _RX_FIN.sub(" ", t)
        t = re.sub(r"F/?L[-\s]*\d+(?:\.\d+)?\s*(?:MM)?", " ", t)
        t = re.sub(r",\s*[\d\.]+\s*(?:\"|''|\"\"|IN\b|INCH\b)", " ", t)   # , 25"  or , 7''
        t = re.sub(r"[\d\.]+\s*(?:\"|'')", " ", t)
        t = re.sub(r"HYDROPH\w*|HYDR\b|HYDRO\b|\bBARE\b", " ", t)
        t = re.sub(r"\bMM\b", " ", t)
        t = re.sub(r"[,;]", " ", t)
        t = re.sub(r"\s+", " ", t).strip(" -")

        parts = [p.strip() for p in re.split(r"\s*/\s*", t) if p.strip()]
        surface_parts, edge = [], ""
        for idx, p in enumerate(parts):
            m_r, m_s = _EDGE_RIPPLE.search(p), _EDGE_STRAIGHT.search(p)
            if m_r or m_s:
                m = m_r or m_s
                edge = edge or ("Ripple" if m_r else "Straight")
                rest = (p[:m.start()] + " " + p[m.end():]).strip()
                if rest:
                    surface_parts.append(rest)
            elif re.fullmatch(r"R|RI", p):
                edge = edge or "Ripple"
            elif "PLAIN" in p:
                edge = edge or "Plain"
            elif idx > 0 and surface_parts and re.fullmatch(r"FLAT(\s+EDGE)?|FLA", p):
                edge = edge or "Flat"
            elif idx > 0 and surface_parts and re.search(r"SIN[EW]|\bSW\b", p) and \
                    _classify_surface(" ".join(surface_parts)) not in ("", "Sinewave"):
                edge = edge or "Sinewave"
            else:
                surface_parts.append(p)
        surface = _classify_surface(" ".join(surface_parts))
        if not surface and re.fullmatch(r"S", " ".join(surface_parts).strip()) and edge == "Ripple":
            surface = "Sinewave"
        out["surface"], out["edge"] = surface, edge

        if out["fpi"] is None or out["fins"] is None or out["length_mm"] is None:
            out["check"] = "Could not read FPI / fins / length"
        else:
            calc = out["fpi"] * out["length_mm"] / 25.4
            if abs(calc - out["fins"]) > max(2.0, 0.03 * out["fins"]):
                out["check"] = "Fins do not match FPI x length"
        if not out["check"] and not surface:
            out["check"] = "Surface pattern not stated"
    except Exception:
        out["check"] = "Could not parse"
    return out


# === SHAREPOINT-ONLY GRAPH API INTEGRATION ===
import msal
import json

class AuthConfig:
    CLIENT_ID = (os.environ.get("AZURE_CLIENT_ID") or "").strip()
    CLIENT_SECRET = (os.environ.get("AZURE_CLIENT_SECRET") or "").strip()
    AUTHORITY = 'https://login.microsoftonline.com/ed97e9bb-e119-4bd4-ab00-307d64bdf908'
    if not CLIENT_ID or not CLIENT_SECRET:
        raise RuntimeError(
            "AZURE_CLIENT_ID / AZURE_CLIENT_SECRET not set - add them as GitHub repository secrets"
        )

def get_access_token():
    """Get Microsoft Graph API access token (retries on transient network errors)"""
    auth_config = AuthConfig()
    last_error = None
    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            app = msal.ConfidentialClientApplication(
                auth_config.CLIENT_ID,
                authority=auth_config.AUTHORITY,
                client_credential=auth_config.CLIENT_SECRET
            )
            result = app.acquire_token_silent(scopes=['https://graph.microsoft.com/.default'], account=None)
            if not result:
                result = app.acquire_token_for_client(scopes=['https://graph.microsoft.com/.default'])
            if 'access_token' in result:
                return result['access_token']
            # A bad secret / client ID will not fix itself, so fail immediately
            raise Exception(f"Failed to acquire access token: {result.get('error_description', result)}")
        except requests.exceptions.RequestException as e:
            last_error = e
            if attempt < MAX_ATTEMPTS:
                wait = BACKOFF_SECONDS[attempt - 1]
                logger.warning(f"Token request network error ({e}); retrying in {wait}s (attempt {attempt}/{MAX_ATTEMPTS})")
                time.sleep(wait)
    raise Exception(f"Failed to acquire access token after {MAX_ATTEMPTS} attempts: {last_error}")

def get_excel_column_letter(col_num):
    """Convert column number to Excel column letter (e.g., 1 -> A, 27 -> AA)"""
    letter = ""
    while col_num > 0:
        col_num -= 1
        letter = chr(col_num % 26 + ord('A')) + letter
        col_num //= 26
    return letter

# ============================================================================
# GRAPH API RETRY SETTINGS
# ============================================================================
# SharePoint/Graph occasionally returns 429 (throttled) or 5xx (temporarily
# unavailable). These clear on their own, so every critical Graph call below
# is retried with an increasing wait before giving up.
RETRY_STATUS_CODES = {429, 500, 502, 503, 504}
MAX_ATTEMPTS = 6                        # 1 try + 5 retries
BACKOFF_SECONDS = [10, 20, 40, 60, 90]  # waits between attempts (~3.5 min total)


class GraphCallError(Exception):
    """Raised when a critical Graph call still fails after all retries."""
    pass


def graph_call(session, method, url, description, **kwargs):
    """
    Make one Graph API call, retrying on 429/5xx and connection errors.
    Returns the response on success (2xx).
    Returns "INVALID_SESSION" if the server reports the workbook session is gone
    (the caller already knows how to recover from that: open a new session and
    rewrite the sheet).
    Raises GraphCallError if it still fails after all attempts, or immediately
    on a non-retryable error (e.g. 400 for a reason other than a dead session,
    401, 403, 404).
    """
    last_problem = None
    for attempt in range(1, MAX_ATTEMPTS + 1):
        retry_after = None
        try:
            resp = session.request(method, url, timeout=kwargs.pop("timeout", 120), **kwargs)
        except requests.exceptions.RequestException as e:
            last_problem = f"connection error: {e}"
        else:
            if 200 <= resp.status_code < 300:
                if attempt > 1:
                    logger.info(f"{description} succeeded on attempt {attempt}")
                return resp
            if "InvalidSession" in resp.text:
                return "INVALID_SESSION"
            if resp.status_code not in RETRY_STATUS_CODES:
                raise GraphCallError(f"{description} failed with {resp.status_code} (not retryable): {resp.text[:400]}")
            last_problem = f"HTTP {resp.status_code}"
            retry_after = resp.headers.get("Retry-After")
        if attempt < MAX_ATTEMPTS:
            wait = BACKOFF_SECONDS[attempt - 1]
            if retry_after and str(retry_after).isdigit():
                wait = max(wait, int(retry_after))
            logger.warning(f"{description}: {last_problem}; retrying in {wait}s (attempt {attempt}/{MAX_ATTEMPTS})")
            time.sleep(wait)
    raise GraphCallError(f"{description} failed after {MAX_ATTEMPTS} attempts ({last_problem})")


def graph_headers(access_token, session_id=None):
    """
    Build the headers for a Graph Excel API call.

    When session_id (or ACTIVE_SESSION) is set, the call joins that explicit workbook
    session. When it is not - because no session could be opened on this workbook, see
    below - the 'workbook-session-id' header is simply omitted: Microsoft Graph then
    treats the call as its own one-off transaction, which still reads/writes the real
    file and still persists the change, it just does not share state with other calls.
    This is the documented fallback for exactly the case where an explicit session
    cannot be opened (the normal case for very large workbooks, where createSession
    itself can time out even though individual range reads/writes succeed fine).
    """
    headers = {"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"}
    sid = ACTIVE_SESSION["id"] or session_id
    if sid:
        headers["workbook-session-id"] = sid
    return headers

# --- Excel session handling: if SharePoint drops the session (for example because someone else edited the
# --- workbook at the same moment), wait, open a new session and write the whole sheet again.
ACTIVE_SESSION = {"id": None}

def _create_workbook_session(site_id, file_server_relative_url, access_token):
    url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/createSession"
    session = requests.Session()
    try:
        r = graph_call(session, "POST", url, "Create Excel session",
                       headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
                       json={"persistChanges": True})
        if r not in (None, "INVALID_SESSION"):
            return r.json().get("id")
    except GraphCallError as e:
        logger.error(f"Could not open a new Excel session: {e}")
    return None

def create_sharepoint_worksheet_optimized(site_id, file_server_relative_url, worksheet_name, df, access_token, session_id, *args, **kwargs):
    """Write a worksheet. If the Excel session is dropped mid-write, open a fresh session and rewrite the whole sheet."""
    max_sheet_attempts = 4
    current = ACTIVE_SESSION["id"] or session_id
    for attempt in range(1, max_sheet_attempts + 1):
        result = _write_worksheet_once(site_id, file_server_relative_url, worksheet_name, df, access_token, current, *args, **kwargs)
        if result != "INVALID_SESSION":
            return result
        logger.warning(f"Excel session was dropped while writing '{worksheet_name}' (attempt {attempt}/{max_sheet_attempts}). "
                       f"This usually means someone else has the workbook open and is editing it.")
        if attempt == max_sheet_attempts:
            break
        time.sleep(30 * attempt)
        new_id = _create_workbook_session(site_id, file_server_relative_url, access_token)
        if new_id:
            current = new_id
            ACTIVE_SESSION["id"] = new_id
            logger.info(f"Opened a new Excel session for '{worksheet_name}' and starting that sheet again")
    logger.error(f"Giving up on '{worksheet_name}' after {max_sheet_attempts} attempts - close the workbook everywhere and run again")
    return False

def _write_worksheet_once(site_id, file_server_relative_url, worksheet_name, df, access_token, session_id, chunk_size=1000, max_retries=5):
    """
    Robust creation/update for large SharePoint sheets with data validation and exponential backoff.

    Order of operations: write the new data first, THEN clear whatever old rows/columns are
    left over beyond it. A sheet is never blanked before the new data is confirmed written, so a
    failure partway through leaves the PREVIOUS data in place instead of an empty sheet.
    """
    session = requests.Session()
    api_headers = graph_headers(access_token, session_id)
    base = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook"

    logger.info(f"Creating/updating SharePoint worksheet: {worksheet_name}")
    try:
        # Check for worksheet existence, create if needed
        ws_list_resp = graph_call(session, "GET", f"{base}/worksheets", f"List worksheets ({worksheet_name})",
                                  headers=api_headers)
        if ws_list_resp == "INVALID_SESSION":
            return "INVALID_SESSION"
        ws_names = [ws['name'] for ws in ws_list_resp.json().get('value', [])]
        if worksheet_name not in ws_names:
            r = graph_call(session, "POST", f"{base}/worksheets/add", f"Create worksheet '{worksheet_name}'",
                           headers=api_headers, json={"name": worksheet_name})
            if r == "INVALID_SESSION":
                return "INVALID_SESSION"
            old_rows, old_cols = 0, 0
        else:
            # Remember how big the PREVIOUS data was, so we know what to clear after the new
            # data lands (and never have to blank the sheet up front to find out).
            old_rows, old_cols = 0, 0
            try:
                # $select keeps Graph from sending every cell value back just to report the size;
                # if it is ever refused, fall back to the plain v5 request
                size_url = f"{base}/worksheets('{worksheet_name}')/usedRange(valuesOnly=true)"
                try:
                    used = graph_call(session, "GET", f"{size_url}?$select=address,rowCount,columnCount",
                                      f"Read current size of '{worksheet_name}'", headers=api_headers)
                except GraphCallError as e:
                    logger.warning(f"Size-only read of '{worksheet_name}' refused ({e}); reading the full used range instead")
                    used = graph_call(session, "GET", size_url,
                                      f"Read current size of '{worksheet_name}'", headers=api_headers)
                if used == "INVALID_SESSION":
                    return "INVALID_SESSION"
                rc = used.json()
                old_rows, old_cols = rc.get("rowCount", 0), rc.get("columnCount", 0)
            except GraphCallError as e:
                logger.warning(f"Could not read previous size of '{worksheet_name}' ({e}); "
                               f"will skip clearing leftovers for this sheet")
    except GraphCallError as e:
        logger.error(f"Could not prepare worksheet '{worksheet_name}': {e}")
        return False

    # Write headers first
    headers = df.columns.tolist()
    data_conv = df.astype(str).where(~df.isna(), "")  # Ensures all strings
    num_cols = len(headers)
    header_range = f"A1:{get_excel_column_letter(num_cols)}1"
    header_url = f"{base}/worksheets('{worksheet_name}')/range(address='{header_range}')"
    try:
        r = graph_call(session, "PATCH", header_url, f"Write header ({worksheet_name})",
                       headers=api_headers, json={"values": [headers]})
        if r == "INVALID_SESSION":
            return "INVALID_SESSION"
    except GraphCallError as e:
        logger.error(f"Could not write header for '{worksheet_name}': {e}")
        return False

    total_rows = len(df)
    total_chunks = (total_rows + chunk_size - 1) // chunk_size
    for chunk_idx in range(0, total_rows, chunk_size):
        chunk_end = min(chunk_idx + chunk_size, total_rows)
        chunk_data = data_conv.iloc[chunk_idx:chunk_end].values.tolist()
        chunk_range = f"A{chunk_idx+2}:{get_excel_column_letter(num_cols)}{chunk_end+1}"
        chunk_url = f"{base}/worksheets('{worksheet_name}')/range(address='{chunk_range}')"
        try:
            r = graph_call(session, "PATCH", chunk_url,
                           f"Write chunk {(chunk_idx//chunk_size)+1}/{total_chunks} ({worksheet_name})",
                           headers=api_headers, json={"values": chunk_data})
            if r == "INVALID_SESSION":
                logger.warning(f"Excel session dropped on chunk {(chunk_idx//chunk_size)+1} of '{worksheet_name}'")
                return "INVALID_SESSION"
            logger.info(f"Chunk {(chunk_idx//chunk_size)+1}/{total_chunks} successful for '{worksheet_name}' (rows {chunk_idx+2}-{chunk_end+1})")
        except GraphCallError as e:
            logger.error(f"Giving up on chunk {(chunk_idx//chunk_size)+1} for '{worksheet_name}': {e}")
            return False

    new_rows, new_cols = total_rows + 1, num_cols  # +1 for the header row
    logger.info(f"'{worksheet_name}' data upload complete ({new_rows} rows x {new_cols} cols).")

    # Now that the new data is confirmed in, clear whatever old rows/columns are left over
    # beyond it. A failure here is logged but does NOT fail the write - the new data already
    # landed; worst case a stale tail of old rows/columns lingers until the next run.
    # v6: leftovers are cleared in blocks of 20k rows (one 170k-row clear can time out on its own)
    # and with applyTo "All", so the old per-cell borders go with the old values.
    leftovers = []
    last_col = get_excel_column_letter(max(old_cols, new_cols))
    for start in range(new_rows + 1, old_rows + 1, 20000):
        leftovers.append(f"A{start}:{last_col}{min(start + 19999, old_rows)}")
    if old_cols > new_cols:
        leftovers.append(f"{get_excel_column_letter(new_cols + 1)}1:{get_excel_column_letter(old_cols)}{new_rows}")
    for address in leftovers:
        try:
            graph_call(session, "POST", f"{base}/worksheets('{worksheet_name}')/range(address='{address}')/clear",
                      f"Clear leftover range {address} ({worksheet_name})",
                      headers=api_headers, json={"applyTo": "All"})
        except GraphCallError as e:
            logger.warning(f"Could not clear leftover range {address} in '{worksheet_name}' ({e}); "
                           f"the new data is in place, some old rows/columns beyond it may remain")
    if leftovers:
        logger.info(f"Cleared leftover rows/columns in '{worksheet_name}'")

    apply_sharepoint_formatting_fixed(site_id, file_server_relative_url, worksheet_name, session_id, access_token, num_cols)
    return True
def apply_sharepoint_formatting_fixed(site_id, file_server_relative_url, worksheet_name, session_id, access_token, num_cols):
    """Format the header row only.
    v5 also put borders, right alignment and autofit on every data cell. On sheets of 40k-176k
    rows that stored formatting for millions of cells in the file and was one of the slowest
    calls of the run, so v6 leaves the data cells alone."""
    session = requests.Session()
    api_headers = graph_headers(access_token, session_id)
    header_range = f"A1:{get_excel_column_letter(num_cols)}1"
    range_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/range(address='{header_range}')"
    try:
        session.patch(f"{range_url}/format/fill", headers=api_headers, json={"color": "#366092"}, timeout=60)
        session.patch(f"{range_url}/format/font", headers=api_headers, json={"bold": True, "color": "#FFFFFF"}, timeout=60)
        session.post(f"{range_url}/format/autofitColumns", headers=api_headers, timeout=60)
        logger.info(f"Formatted header row of '{worksheet_name}' ({header_range})")
    except Exception as e:
        logger.warning(f"Error formatting header of '{worksheet_name}': {e}")

def create_sharepoint_bom_layerwise_sheet(site_id, file_server_relative_url, session_id, access_token, df_bom, all_paths, max_layers, header_mapping, header_qty_mapping):
    """Create BOM Layer-wise sheet directly in SharePoint"""
    logger.info("Creating BOM Layer-wise sheet directly in SharePoint...")

    # Same processing logic as local export_tree_paths
    columns = ["Main_FG", "Main_FG_Description"]
    for lvl in range(1, max_layers+1):
        columns.extend([
            f"BOM_Layer_{lvl}",
            f"BOM_Layer_{lvl}_Desc",
            f"BOM_Layer_{lvl}_Quantity",
            f"BOM_Layer_{lvl}_Unit"
        ])

    # Item Text (line 1) columns go at the far right, after all existing columns
    for lvl in range(1, max_layers+1):
        columns.append(f"BOM_Layer_{lvl}_ItemText")
    for lvl in range(1, max_layers+1):
        columns.extend([f"BOM_Layer_{lvl}_MaterialGroup", f"BOM_Layer_{lvl}_MaterialGroupDesc"])
    columns.extend(["Fin_FPI", "Fin_Fins_Per_Pack", "Fin_Length_mm", "Fin_Surface", "Fin_Edge", "Fin_Spec_Check"])

    records = []
    for path in all_paths:
        row = dict()
        if path:
            first = path[0]
            row["Main_FG"] = first[0]
            row["Main_FG_Description"] = header_mapping.get(first[0], "")
        else:
            row["Main_FG"], row["Main_FG_Description"] = "", ""

        for i, comp in enumerate(path):
            if i < max_layers:
                row[f"BOM_Layer_{i+1}"] = comp[4]
                row[f"BOM_Layer_{i+1}_Desc"] = comp[1]
                row[f"BOM_Layer_{i+1}_Quantity"] = comp[2]
                row[f"BOM_Layer_{i+1}_Unit"] = comp[3]
                row[f"BOM_Layer_{i+1}_ItemText"] = comp[6]
                row[f"BOM_Layer_{i+1}_MaterialGroup"] = comp[7]
                row[f"BOM_Layer_{i+1}_MaterialGroupDesc"] = material_group_desc(comp[7])

        # fin-pack spec: read from the item text of the aluminium foil line in this path
        for i, comp in enumerate(path):
            if i < max_layers and is_foil_component(comp[1], comp[7]):
                fin = parse_fin_spec(comp[6])
                row["Fin_FPI"] = _fmt_num(fin["fpi"])
                row["Fin_Fins_Per_Pack"] = _fmt_num(fin["fins"])
                row["Fin_Length_mm"] = _fmt_num(fin["length_mm"])
                row["Fin_Surface"] = fin["surface"]
                row["Fin_Edge"] = fin["edge"]
                row["Fin_Spec_Check"] = fin["check"]
                break
        records.append(row)

    df_out = pd.DataFrame(records, columns=columns)

    # Apply same filtering as local
    filtered_df_out = df_out[df_out["Main_FG"].str.startswith(("A", "C", "I"))].copy()
    logger.info(f"Filtered rows (A/C/I only), remaining count: {len(filtered_df_out)}")

    # Apply BOM header quantity division
    def divide_qtys_by_divisor(qty_str, divisor):
        try:
            if not isinstance(qty_str, str):
                qty_str = str(qty_str) if qty_str is not None else ''
            qtys = [float(q) if q.strip() else 0 for q in qty_str.split(" | ")]
            divided = [q / divisor if divisor != 0 else q for q in qtys]
            result = " | ".join([str(round(q, 10)).rstrip('0').rstrip('.') if q != 0 else '0' for q in divided])
            return result
        except Exception as e:
            return qty_str

    for main_fg, group in filtered_df_out.groupby("Main_FG"):
        divisor = header_qty_mapping.get(main_fg, 1)
        try:
            divisor = float(divisor)
        except Exception:
            divisor = 1

        if pd.notna(divisor) and divisor != 0:
            mask = filtered_df_out["Main_FG"] == main_fg
            # ensure column is object/string before assigning string results
            filtered_df_out["BOM_Layer_1_Quantity"] = filtered_df_out["BOM_Layer_1_Quantity"].astype("object")
            filtered_df_out.loc[mask, "BOM_Layer_1_Quantity"] = (
                filtered_df_out.loc[mask, "BOM_Layer_1_Quantity"]
                .apply(divide_qtys_by_divisor, divisor=divisor)
            )

    # Convert G to KG for all layers
    for lvl in range(1, max_layers + 1):
        qty_col = f"BOM_Layer_{lvl}_Quantity"
        unit_col = f"BOM_Layer_{lvl}_Unit"

        if qty_col in filtered_df_out.columns and unit_col in filtered_df_out.columns:
            mask = filtered_df_out[unit_col].astype(str).str.strip().str.upper() == 'G'
            n_g = mask.sum()

            if n_g > 0:
                def convert_and_scale_g_to_kg(row):
                    qty_value = row[qty_col]
                    try:
                        q = float(qty_value)
                    except Exception:
                        q = 0.0
                    return q / 1000

                filtered_df_out.loc[mask, qty_col] = filtered_df_out.loc[mask].apply(convert_and_scale_g_to_kg, axis=1)
                filtered_df_out.loc[mask, unit_col] = 'KG'

    # Writing 8 columns that are blank on every row only made the workbook bigger
    mg_cols = [c for c in filtered_df_out.columns if "_MaterialGroup" in c]
    if mg_cols and filtered_df_out[mg_cols].fillna("").astype(str).eq("").all().all():
        logger.warning("No material groups were found for any component - leaving the MaterialGroup columns out")
        filtered_df_out = filtered_df_out.drop(columns=mg_cols)

    ok = create_sharepoint_worksheet_optimized(site_id, file_server_relative_url, "BOM Layer-wise", filtered_df_out, access_token, session_id)
    # The caller builds FG-RM Required from this in-memory copy instead of reading the
    # 1.5M-cell sheet back from SharePoint; blanks match what the read-back used to return.
    return ok, filtered_df_out.fillna("")

def create_sharepoint_fg_rm_sheet(site_id, file_server_relative_url, session_id, access_token, main_sheet_df, max_layers):
	"""Create FG-RM Required sheet directly in SharePoint with correct cascading (no scaling for Layer-1)."""

	logger.info("Creating FG-RM Required sheet directly in SharePoint...")

	summary_records = []

	for main_fg in main_sheet_df["Main_FG"].unique():
		if not isinstance(main_fg, str) or not main_fg:
			continue

		fg_rows = main_sheet_df[main_sheet_df["Main_FG"] == main_fg]
		rm_dict = {}

		for _, row in fg_rows.iterrows():
			# Base Layer-1 quantity (used only as multiplier for deeper layers)
			layer_1_qty_str = str(row.get("BOM_Layer_1_Quantity", "1")).strip()
			try:
				layer_1_qtys = [float(q) if q.strip() else 1 for q in layer_1_qty_str.split(" | ")]
				layer_1_total_qty = sum(layer_1_qtys) if layer_1_qtys else 1
			except Exception:
				layer_1_total_qty = 1

			for lvl in range(1, max_layers + 1):
				mat_col = f"BOM_Layer_{lvl}"
				desc_col = f"BOM_Layer_{lvl}_Desc"
				qty_col = f"BOM_Layer_{lvl}_Quantity"
				unit_col = f"BOM_Layer_{lvl}_Unit"

				mat = str(row.get(mat_col, "")).strip()

				# Only real raw materials (start with 5)
				if not mat.startswith("5"):
					continue

				desc = str(row.get(desc_col, "")).strip()
				qty_str = str(row.get(qty_col, "")).strip()
				unit = str(row.get(unit_col, "")).strip().upper()

				# Skip packaging / non-RM items
				if any(k in desc.lower() for k in ("pallet", "box", "label", "paint")):
					continue

				# Parse component quantity (sum pipe-separated if present)
				try:
					qtys = [float(q) if q.strip() else 0 for q in qty_str.split(" | ")]
					component_qty = sum(qtys)
				except Exception:
					component_qty = 0

				if component_qty <= 0:
					continue

				# --- FIXED LOGIC ---
				# Layer-1 RM: use its BOM quantity as FINAL (no multiplication at all)
				if lvl == 1:
					final_qty = component_qty

				# Deeper layers: cascade using Layer-1 and intermediate layers
				else:
					cascading_multiplier = layer_1_total_qty

					# Multiply by quantities of all superior intermediate layers (2 .. lvl-1)
					for superior_lvl in range(2, lvl):
						superior_qty_col = f"BOM_Layer_{superior_lvl}_Quantity"
						superior_qty_str = str(row.get(superior_qty_col, "1")).strip()
						try:
							superior_qtys = [float(q) if q.strip() else 1 for q in superior_qty_str.split(" | ")]
							superior_total_qty = sum(superior_qtys) if superior_qtys else 1
						except Exception:
							superior_total_qty = 1
						cascading_multiplier *= superior_total_qty

					final_qty = component_qty * cascading_multiplier
				# --- END FIX ---

				if mat in rm_dict:
					rm_dict[mat]["qty"] += final_qty
				else:
					rm_dict[mat] = {
						"desc": desc,
						"qty": final_qty,
						"unit": unit,
					}

		for rm_code, details in rm_dict.items():
			if details["qty"] > 0:
				summary_records.append(
					{
						"Main_FG": main_fg,
						"Raw Material": rm_code,
						"Raw Material Description": details["desc"],
						"Raw Material qty": details["qty"],
						"Raw Material UoM": details["unit"],
					}
				)

	summary_df = pd.DataFrame(
		summary_records,
		columns=[
			"Main_FG",
			"Raw Material",
			"Raw Material Description",
			"Raw Material qty",
			"Raw Material UoM",
		],
	)

	return create_sharepoint_worksheet_optimized(
		site_id,
		file_server_relative_url,
		"FG-RM Required",
		summary_df,
		access_token,
		session_id,
		chunk_size=1000,
	)

def create_sharepoint_mb52_sheet(site_id, file_server_relative_url, session_id, access_token, stock_df):
    """Create MB52 Raw Data sheet directly in SharePoint"""
    logger.info("Creating MB52 Raw Data sheet directly in SharePoint...")
    
    return create_sharepoint_worksheet_optimized(site_id, file_server_relative_url, "MB52 Raw Data", stock_df, access_token, session_id)

def read_sharepoint_worksheet_data(site_id, file_server_relative_url, worksheet_name, session_id, access_token):
    """Generic function to read data from any SharePoint worksheet"""
    session = requests.Session()
    
    api_headers = graph_headers(access_token, session_id)
    
    try:
        # Get used range from specified worksheet (retried: a transient 429/5xx here would
        # otherwise look identical to "the sheet is genuinely empty")
        range_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/usedRange"
        range_response = graph_call(session, "GET", range_url, f"Read {worksheet_name}", headers=api_headers)
        if range_response == "INVALID_SESSION":
            logger.error(f"Excel session dropped while reading {worksheet_name}")
            return pd.DataFrame()

        range_data = range_response.json()
        values = range_data.get('values', [])
        
        if not values:
            logger.warning(f"No data found in {worksheet_name}")
            return pd.DataFrame()
        
        # Convert to DataFrame
        headers = values[0]
        data_rows = values[1:]
        df = pd.DataFrame(data_rows, columns=headers)
        
        logger.info(f"Read {len(df)} rows from SharePoint {worksheet_name}")
        return df
        
    except Exception as e:
        logger.error(f"Error reading SharePoint {worksheet_name}: {e}")
        return pd.DataFrame()

def set_sharepoint_timestamp(site_id, file_server_relative_url, worksheet_name, session_id, access_token, cell_address="R1"):
    """Set timestamp in specified cell"""
    session = requests.Session()
    
    api_headers = graph_headers(access_token, session_id)
    
    update_datetime = datetime.now().strftime('Last Updated: %d-%m-%Y %H:%M:%S')
    payload = {"values": [[update_datetime]]}
    
    update_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/range(address='{cell_address}')"
    
    try:
        r = session.patch(update_url, headers=api_headers, json=payload, timeout=60)
        if r.status_code in [200, 201]:
            logger.info(f"Timestamp written to cell {cell_address} in '{worksheet_name}'")
            return True
        else:
            logger.error(f"Failed to update timestamp: {r.status_code}")
            return False
    except Exception as e:
        logger.error(f"Error setting timestamp: {e}")
        return False

# --- BOM Item Data Fetch (SAME AS ORIGINAL) ---
def fetch_bom_data(api_url, username, password, verify_ssl=True, preferred_format='json', batch_size=500, max_workers=15) -> pd.DataFrame:
    logger.info("Starting BOM ITEM data fetch with enhanced concurrency...")
    session = requests.Session()
    session.auth = (username, password)
    session.verify = verify_ssl
    session.headers.update({'Accept': f'application/{preferred_format}', 'Content-Type': f'application/{preferred_format}'})
    
    select_fields = [
        'Material', 'ComponentDescription', 'BillOfMaterialComponent',
        'IsAssembly', 'BillOfMaterialItemQuantity', 'BillOfMaterialItemUnit',
        'BOMItemDescription', 'MaterialGroup'
    ]
    
    endpoint = f"{api_url}/MaterialBOMItem"
    
    def fetch_batch(skip):
        """Returns the batch's records ([] past the end), or None if it still fails after 3 tries."""
        params = {'$select': ','.join(select_fields), '$top': batch_size, '$orderby': 'Material', '$skip': skip}
        for attempt in range(1, 4):
            try:
                logger.info(f"Fetching BOM items batch starting at {skip}...")
                response = session.get(endpoint, params=params, timeout=180)
                response.raise_for_status()

                content_type = response.headers.get('content-type', '')
                if 'xml' in content_type or preferred_format == 'xml':
                    return parse_xml_bom(response.content)
                else:
                    return parse_json_bom(response.text)
            except Exception as e:
                logger.warning(f"BOM items batch at skip={skip} failed (attempt {attempt}/3): {e}")
                time.sleep(5 * attempt)
        return None

    # Total item count ($inlinecount is needed for SAP to return __count at all)
    params_count = {'$top': 1, '$inlinecount': 'allpages'}
    resp_count = session.get(endpoint, params=params_count, timeout=30)
    resp_count.raise_for_status()
    data_count = resp_count.json()
    total_items = int(data_count.get('d', {}).get('__count', '0') or 0)
    if total_items == 0:
        total_items = 50000
    logger.info(f"SAP reports {total_items} BOM items")

    total_batches = (total_items // batch_size) + 1
    all_records = []
    failed_batches = []

    # requests' default pool keeps 10 connections; match it to the number of threads
    adapter = requests.adapters.HTTPAdapter(pool_connections=max_workers, pool_maxsize=max_workers)
    session.mount('https://', adapter)

    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = {executor.submit(fetch_batch, batch_num * batch_size): batch_num for batch_num in range(total_batches)}
        for future in as_completed(futures):
            records = future.result()
            if records is None:
                failed_batches.append(futures[future])
            elif records:
                all_records.extend(records)
                logger.info(f"Fetched batch size: {len(records)} (total so far: {len(all_records)})")

    # A missing batch means missing BOM lines, which would silently understate RM requirements
    if failed_batches:
        raise RuntimeError(f"BOM item batches {sorted(failed_batches)} could not be fetched from SAP - "
                           f"stopping rather than writing an incomplete BOM")

    logger.info(f"Total BOM ITEM records fetched: {len(all_records)}")
    return pd.DataFrame(all_records)

def fetch_material_bom_headers(api_url, username, password, verify_ssl=True, preferred_format='json', batch_size=500):
    logger.info("Starting BOM HEADER data fetch...")
    session = requests.Session()
    session.auth = (username, password)
    session.verify = verify_ssl
    session.headers.update({'Accept': f'application/{preferred_format}', 'Content-Type': f'application/{preferred_format}'})
    
    select_fields = ['Material', 'ProductDescription', 'BOMHeaderQuantityInBaseUnit']
    endpoint = f"{api_url}/MaterialBOM"
    params = {'$select': ','.join(select_fields), '$top': batch_size}
    
    all_records = []
    skip = 0
    
    while True:
        params['$skip'] = skip
        logger.info(f"Fetching BOM headers batch starting at {skip}...")
        response = session.get(endpoint, params=params, timeout=180)
        response.raise_for_status()
        
        content_type = response.headers.get('content-type', '')
        if 'xml' in content_type or preferred_format == 'xml':
            ns = {'atom': 'http://www.w3.org/2005/Atom', 'd': 'http://schemas.microsoft.com/ado/2007/08/dataservices', 'm': 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'}
            tree = ET.fromstring(response.content)
            entries = tree.findall('.//atom:entry', ns)
            records = []
            for entry in entries:
                properties = entry.find('.//m:properties', ns)
                if properties is not None:
                    m = properties.find('.//d:Material', ns)
                    d = properties.find('.//d:ProductDescription', ns)
                    h = properties.find('.//d:BOMHeaderQuantityInBaseUnit', ns)
                    rec = {
                        'Material': m.text.strip() if m is not None and m.text else '',
                        'ProductDescription': d.text.strip() if d is not None and d.text else '',
                        'BOMHeaderQuantityInBaseUnit': h.text.strip() if h is not None and h.text else ''
                    }
                    records.append(rec)
        else:
            import json
            dct = response.json()
            records = [{'Material': i.get('Material', ''), 'ProductDescription': i.get('ProductDescription', ''), 'BOMHeaderQuantityInBaseUnit': i.get('BOMHeaderQuantityInBaseUnit', '')}
                       for i in dct.get('d', {}).get('results', [])]
        
        if not records:
            logger.warning("No more BOM header records found.")
            break
        
        all_records.extend(records)
        logger.info(f"Fetched header batch size: {len(records)} (total so far: {len(all_records)})")
        
        if len(records) < batch_size:
            break
        skip += batch_size
    
    logger.info(f"Total BOM HEADER records fetched: {len(all_records)}")
    return pd.DataFrame(all_records)

def fetch_material_stock_data(api_url, username, password, verify_ssl=True, preferred_format='json', batch_size=500):
    logger.info("Starting Material Stock data fetch...")
    session = requests.Session()
    session.auth = (username, password)
    session.verify = verify_ssl
    session.headers.update({'Accept': f'application/{preferred_format}', 'Content-Type': f'application/{preferred_format}'})
    
    select_fields = [
        'Material', 'Plant', 'StorageLocation', 'Batch',
        'InventorySpecialStockType',
        'InventoryStockType', 'MaterialBaseUnit', 'MatlWrhsStkQtyInMatlBaseUnit'
    ]
    
    stock_api_url = "https://my409486-api.s4hana.cloud.sap/sap/opu/odata/sap/API_MATERIAL_STOCK_SRV"
    endpoint = f"{stock_api_url}/A_MatlStkInAcctMod"
    # Only the plants in STOCK_PLANTS are needed; filtering in SAP skips the other plants' rows
    plant_filter = ' or '.join(f"Plant eq '{plant}'" for plant in STOCK_PLANTS)
    params = {'$select': ','.join(select_fields), '$top': batch_size,
              '$filter': f"({plant_filter})", '$orderby': 'Plant,Material,StorageLocation,Batch'}

    all_records = []
    skip = 0

    while True:
        params['$skip'] = skip
        logger.info(f"Fetching Material Stock batch starting at {skip}...")

        for attempt in range(1, 4):
            try:
                response = session.get(endpoint, params=params, timeout=180)
                response.raise_for_status()
                break
            except Exception as e:
                if attempt == 3:
                    # Stopping here used to leave a partial stock list that looked complete
                    raise RuntimeError(f"Material Stock batch at skip={skip} failed after 3 attempts: {e}")
                logger.warning(f"Material Stock batch at skip={skip} failed (attempt {attempt}/3): {e}")
                time.sleep(5 * attempt)

        content_type = response.headers.get('content-type', '')
        if 'xml' in content_type or preferred_format == 'xml':
            records = parse_xml_material_stock(response.content)
        else:
            records = parse_json_material_stock(response.text)

        if not records:
            break

        all_records.extend(records)
        logger.info(f"Fetched Material Stock batch size: {len(records)} (total so far: {len(all_records)})")

        if len(records) < batch_size:
            break
        skip += batch_size

    logger.info(f"Total Material Stock records fetched for plants {', '.join(STOCK_PLANTS)}: {len(all_records)}")
    return pd.DataFrame(all_records, columns=select_fields)

def summarise_plant_stock(stock_df):
    """Stock for STOCK_PLANTS summed per plant / material / storage location / stock type (Batch dropped).
    Keeps the same 8 column names as before so the 'MB52 1000 Stock' pivot still finds its fields."""
    columns = ['Material', 'Plant', 'StorageLocation', 'Batch', 'InventorySpecialStockType',
               'InventoryStockType', 'MaterialBaseUnit', 'MatlWrhsStkQtyInMatlBaseUnit']
    if stock_df.empty:
        return pd.DataFrame(columns=columns)
    df = stock_df.copy()
    for col in STOCK_GROUP_BY:
        df[col] = df[col].fillna('').astype(str).str.strip()
    df = df[df['Plant'].isin(STOCK_PLANTS)]
    df['MatlWrhsStkQtyInMatlBaseUnit'] = pd.to_numeric(df['MatlWrhsStkQtyInMatlBaseUnit'], errors='coerce').fillna(0)
    summary = df.groupby(STOCK_GROUP_BY, as_index=False)['MatlWrhsStkQtyInMatlBaseUnit'].sum()
    summary['MatlWrhsStkQtyInMatlBaseUnit'] = summary['MatlWrhsStkQtyInMatlBaseUnit'].round(3)
    summary['Batch'] = ''
    summary = summary.sort_values(['Plant', 'Material', 'StorageLocation']).reset_index(drop=True)
    for plant, rows in summary.groupby('Plant'):
        logger.info(f"Plant {plant} stock summarised: {(df['Plant'] == plant).sum()} SAP rows -> {len(rows)} rows "
                    f"({rows['Material'].nunique()} materials)")
    return summary[columns]

def fetch_product_groups(username, password, verify_ssl=True, batch_size=5000):
    """Material -> material group from the product master. MaterialBOMItem.MaterialGroup comes back
    blank on this system, which left the BOM_Layer_n_MaterialGroup columns empty in v4/v5.
    Returns {} (and the run carries on) if the product API is not available to this SAP user."""
    logger.info("Starting product material-group fetch...")
    session = requests.Session()
    session.auth = (username, password)
    session.verify = verify_ssl
    session.headers.update({'Accept': 'application/json'})
    endpoint = f"{PRODUCT_API_URL}/A_Product"
    params = {'$select': 'Product,ProductGroup', '$top': batch_size, '$orderby': 'Product'}
    groups = {}
    skip = 0
    try:
        while True:
            params['$skip'] = skip
            response = session.get(endpoint, params=params, timeout=180)
            response.raise_for_status()
            results = response.json().get('d', {}).get('results', [])
            for item in results:
                product = str(item.get('Product', '')).strip().upper()
                group = str(item.get('ProductGroup', '') or '').strip().upper()
                if product and group:
                    groups[product] = group
            if len(results) < batch_size:
                break
            skip += batch_size
    except Exception as e:
        logger.warning(f"Could not read material groups from {PRODUCT_API_URL} ({e}); "
                       f"the material-group columns will be left out of BOM Layer-wise this run")
        return {}
    logger.info(f"Material groups fetched for {len(groups)} products")
    return groups

def parse_json_bom(text):
    import json
    try:
        data = json.loads(text)
        results = []
        for item in data.get('d', {}).get('results', []):
            r = {}
            for key in ['Material', 'BillOfMaterialComponent', 'ComponentDescription', 'IsAssembly', 'BillOfMaterialItemQuantity', 'BillOfMaterialItemUnit', 'BOMItemDescription', 'MaterialGroup']:
                value = item.get(key, '')
                if isinstance(value, dict): value = str(value)
                r[key] = value
            results.append(r)
        return results
    except Exception as ex:
        logger.error(f"JSON BOM parse failed: {str(ex)}")
        return []

def parse_xml_bom(xml_content):
    ns = {'atom': 'http://www.w3.org/2005/Atom', 'd': 'http://schemas.microsoft.com/ado/2007/08/dataservices', 'm': 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'}
    tree = ET.fromstring(xml_content)
    records = []
    entries = tree.findall('.//atom:entry', ns)
    
    for entry in entries:
        properties = entry.find('.//m:properties', ns)
        if properties is not None:
            record = {}
            for field in ['Material', 'BillOfMaterialComponent', 'ComponentDescription', 'IsAssembly', 'BillOfMaterialItemQuantity', 'BillOfMaterialItemUnit', 'BOMItemDescription', 'MaterialGroup']:
                elem = properties.find(f'.//d:{field}', ns)
                record[field] = elem.text.strip() if elem is not None and elem.text else ''
            records.append(record)
    return records

def parse_json_material_stock(text):
    import json
    try:
        data = json.loads(text)
        results = []
        for item in data.get('d', {}).get('results', []):
            r = {}
            for key in ['Material', 'Plant', 'StorageLocation', 'Batch', 'InventorySpecialStockType', 'InventoryStockType', 'MaterialBaseUnit', 'MatlWrhsStkQtyInMatlBaseUnit']:
                value = item.get(key, '')
                if isinstance(value, dict):
                    value = str(value)
                r[key] = value
            results.append(r)
        return results
    except Exception as ex:
        logger.error(f"JSON Material Stock parse failed: {str(ex)}")
        return []

def parse_xml_material_stock(xml_content):
    ns = {'atom': 'http://www.w3.org/2005/Atom', 'd': 'http://schemas.microsoft.com/ado/2007/08/dataservices', 'm': 'http://schemas.microsoft.com/ado/2007/08/dataservices/metadata'}
    tree = ET.fromstring(xml_content)
    records = []
    entries = tree.findall('.//atom:entry', ns)
    
    for entry in entries:
        properties = entry.find('.//m:properties', ns)
        if properties is not None:
            record = {}
            for field in ['Material', 'Plant', 'StorageLocation', 'Batch', 'InventorySpecialStockType', 'InventoryStockType', 'MaterialBaseUnit', 'MatlWrhsStkQtyInMatlBaseUnit']:
                elem = properties.find(f'.//d:{field}', ns)
                record[field] = elem.text.strip() if elem is not None and elem.text else ''
            records.append(record)
    return records

def standardize_bom_data(df, product_groups=None):
    logger.info("Standardizing BOM item dataframe...")
    df = df.fillna('')
    
    for col in ['Material', 'BillOfMaterialComponent', 'ComponentDescription', 'BillOfMaterialItemUnit']:
        df[col] = df[col].astype(str).str.strip().str.upper()
    
    if 'BOMItemDescription' not in df.columns:
        df['BOMItemDescription'] = ''
    df['BOMItemDescription'] = df['BOMItemDescription'].astype(str).str.strip()

    if 'MaterialGroup' not in df.columns:
        df['MaterialGroup'] = ''
    df['MaterialGroup'] = df['MaterialGroup'].astype(str).str.strip().str.upper()
    # a component's material group is the same everywhere, so fill blanks from any line that has it
    _mg_known = df[df['MaterialGroup'] != ''].groupby('BillOfMaterialComponent')['MaterialGroup'].first()
    df['MaterialGroup'] = df['MaterialGroup'].where(df['MaterialGroup'] != '', df['BillOfMaterialComponent'].map(_mg_known).fillna(''))
    # then from the product master
    if product_groups:
        df['MaterialGroup'] = df['MaterialGroup'].where(df['MaterialGroup'] != '', df['BillOfMaterialComponent'].map(product_groups).fillna(''))
    _mg_filled = df.drop_duplicates('BillOfMaterialComponent')
    logger.info(f"Material group filled for {(_mg_filled['MaterialGroup'] != '').sum()} of {len(_mg_filled)} components; "
                f"{(_mg_filled['MaterialGroup'].map(material_group_desc) != '').sum()} found in the description list")

    df['BillOfMaterialItemQuantity'] = pd.to_numeric(df['BillOfMaterialItemQuantity'], errors='coerce').fillna(0)
    df['IsAssembly'] = df['IsAssembly'].astype(str).str.upper()
    df = df.drop_duplicates(subset=['Material', 'BillOfMaterialComponent'], keep='first')
    
    logger.info(f"BOM item dataframe standardized, total rows: {len(df)}")
    return df

def tree_explode(df, start_mat, max_depth, parents_seen=None, path=None):
    if parents_seen is None:
        parents_seen = set()
    if path is None:
        path = []
    
    children = df[df["Material"] == start_mat]
    
    if len(children) == 0:
        yield path
    else:
        for idx, row in children.iterrows():
            child = row["BillOfMaterialComponent"]
            if child == "" or child in parents_seen:
                continue
            
            row_tuple = (
                start_mat,
                row["ComponentDescription"],
                row["BillOfMaterialItemQuantity"],
                row["BillOfMaterialItemUnit"],
                child,
                row["IsAssembly"],
                row["BOMItemDescription"],
                row["MaterialGroup"]
            )
            
            next_path = path + [row_tuple]
            
            if row["IsAssembly"] == 'X' and len(next_path) < max_depth:
                yield from tree_explode(df, child, max_depth, parents_seen | {start_mat}, next_path)
            else:
                yield next_path

def create_mb52_stock_dict(mb52_df):
    """Create stock dictionary from MB52 data for Plant 1000 ONLY"""
    logger.info("Creating MB52 stock dictionary for Plant 1000...")
    
    plant_1000_stock = mb52_df[
        mb52_df['Plant'].astype(str).str.strip() == '1000'
    ].copy()
    
    logger.info(f"Plant 1000 stock records: {len(plant_1000_stock)}")
    
    qty_col = 'MatlWrhsStkQtyInMatlBaseUnit'
    if qty_col in plant_1000_stock.columns:
        plant_1000_stock[qty_col] = pd.to_numeric(plant_1000_stock[qty_col], errors='coerce').fillna(0)
    else:
        logger.error(f"Column {qty_col} not found in stock data")
        return {}
    
    # Aggregate by Material
    stock_dict = {}
    for _, row in plant_1000_stock.iterrows():
        material = str(row['Material']).strip().upper()
        qty = float(row[qty_col])
        
        if material in stock_dict:
            stock_dict[material] += qty
        else:
            stock_dict[material] = qty
    
    logger.info(f"MB52 stock dictionary created: {len(stock_dict)} unique materials")
    return stock_dict

def get_rm_requirements_for_fg(fg, fg_rm_df):
    """Get RM requirements for a given FG"""
    rm_req = {}
    fg_norm = str(fg).strip().upper()
    
    rows = fg_rm_df[fg_rm_df['Main_FG'].astype(str).str.strip().str.upper() == fg_norm]
    
    for _, row in rows.iterrows():
        rm_code = str(row['Raw Material']).strip()
        rm_desc = str(row['Raw Material Description']).strip()
        rm_qty = float(row['Raw Material qty']) if pd.notna(row['Raw Material qty']) else 0.0
        rm_uom = str(row['Raw Material UoM']).strip()
        
        if rm_code and rm_qty > 0:
            rm_req[rm_code] = (rm_desc, rm_qty, rm_uom)
    
    return rm_req

def set_fixed_header_timestamp(site_id, file_server_relative_url, worksheet_name, cell_address, session_id, access_token):
    """
    Writes timestamp ONLY to the given header cell (e.g., 'R1') and applies header formatting.
    Usage:
      set_fixed_header_timestamp(SITE_ID, SHAREPOINT_FILE, "Production Tracker", "R1", session_id, access_token)
    """

    session = requests.Session()
    api_headers = graph_headers(access_token, session_id)

    # 1) Write timestamp value
    ts_text = "Last Updated: " + datetime.now().strftime('%d-%m-%Y %H:%M:%S')
    put_url = (
        f"https://graph.microsoft.com/v1.0/sites/{site_id}"
        f"/drive/root:/{file_server_relative_url}:/workbook/worksheets('{worksheet_name}')/range(address='{cell_address}')"
    )
    try:
        r = graph_call(session, "PATCH", put_url, f"Write timestamp ({worksheet_name} {cell_address})",
                       headers=api_headers, json={"values": [[ts_text]]})
        if r == "INVALID_SESSION":
            logger.error(f"Excel session dropped while writing timestamp for {worksheet_name} {cell_address}")
            return False
    except GraphCallError as e:
        logger.error(f"Timestamp write failed for {worksheet_name} {cell_address}: {e}")
        return False

    # 2) Apply header formatting (fill + font) to that exact cell
    fill_url = f"{put_url}/format/fill"
    font_url = f"{put_url}/format/font"

    # Header background color
    session.patch(fill_url, headers=api_headers, json={"color": "#366092"}, timeout=30)
    # Header font style
    session.patch(font_url, headers=api_headers, json={"bold": True, "color": "#FFFFFF"}, timeout=30)

    logger.info(f"Timestamp written and formatted at {worksheet_name} {cell_address}")
    return True

def validate_data_linkage(df_tracker, fg_rm_df):
    """Enhanced data validation with detailed logging"""
    logger.info("=== VALIDATING DATA LINKAGE ===")
    
    try:
        tracker_parts = set(df_tracker['Indus Part #'].astype(str).str.strip().str.upper().unique())
        fg_rm_parts = set(fg_rm_df['Main_FG'].astype(str).str.strip().str.upper().unique())
        
        logger.info(f"Production Tracker unique parts: {len(tracker_parts)}")
        logger.info(f"FG-RM Required unique parts: {len(fg_rm_parts)}")
        
        # Find matches and mismatches
        matches = tracker_parts.intersection(fg_rm_parts)
        tracker_only = tracker_parts - fg_rm_parts
        fg_rm_only = fg_rm_parts - tracker_parts
        
        logger.info(f"Matching parts: {len(matches)}")
        logger.info(f"Tracker-only parts: {len(tracker_only)}")
        logger.info(f"FG-RM-only parts: {len(fg_rm_only)}")
        
        if matches:
            logger.info(f"Sample matches: {list(matches)[:5]}")
        if tracker_only:
            logger.warning(f"Parts in Tracker but not in FG-RM Required: {list(tracker_only)[:5]}")
        
        return len(matches) > 0, matches
        
    except Exception as e:
        logger.error(f"Data validation error: {e}")
        return False, set()

# --- MAIN FUNCTION: SHAREPOINT-ONLY PROCESSING ---
def main():
    """COMPLETE SHAREPOINT-ONLY BOM layer-wise FIFO allocation system"""
    logger.info("Starting SHAREPOINT-ONLY BOM layer-wise FIFO allocation system...")
    start_time = time.time()
    
    # SharePoint configuration
    SITE_ID = "201c767f-23cc-4eff-b76c-47a275706531,094b21f9-6ed4-47c3-a751-3b217a18b98e"
    SHAREPOINT_FILE = "RM Planning SAP New.xlsx"
    
    try:
        # === PHASE 0: SHAREPOINT REACHABILITY PRE-CHECK ===
        # Confirm the Azure credentials work and the file is reachable BEFORE spending
        # 5-7 minutes pulling SAP data. This is a plain metadata read (not a workbook
        # session), so it is fast and cheap, and it catches the class of error that
        # createSession failing later cannot tell apart from "file just too big to open
        # a session on": a bad/expired Azure secret, wrong site or file path, the file
        # renamed or deleted, or no permission on it. A handful of quick attempts only -
        # a slow metadata read strongly suggests the same underlying size/load issue that
        # the session-less fallback already handles, so this check does not try to be
        # exhaustive about transient errors; MAX_ATTEMPTS later covers the calls that
        # actually matter.
        logger.info("=== SHAREPOINT REACHABILITY PRE-CHECK ===")
        precheck_token = get_access_token()
        precheck_session = requests.Session()
        precheck_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive/root:/{SHAREPOINT_FILE}"
        PRECHECK_ATTEMPTS, PRECHECK_BACKOFF = 2, [10]  # fast: ~10s worst case, not the full retry table
        last_problem = None
        for attempt in range(1, PRECHECK_ATTEMPTS + 1):
            try:
                resp = precheck_session.get(precheck_url, timeout=15,
                                            headers={"Authorization": f"Bearer {precheck_token}"})
                if resp.status_code == 200:
                    logger.info(f"'{SHAREPOINT_FILE}' is reachable - proceeding to SAP data collection")
                    break
                last_problem = f"HTTP {resp.status_code}: {resp.text[:200]}"
            except requests.exceptions.RequestException as e:
                last_problem = f"connection error: {e}"
            if attempt < PRECHECK_ATTEMPTS:
                logger.warning(f"Reachability check failed ({last_problem}); retrying once in {PRECHECK_BACKOFF[attempt - 1]}s")
                time.sleep(PRECHECK_BACKOFF[attempt - 1])
        else:
            raise RuntimeError(
                f"Cannot reach '{SHAREPOINT_FILE}' before even starting SAP data collection: {last_problem}. "
                f"Check the Azure credentials, the site/file path, and that the file still exists "
                f"at that location before running again - this is a basic connectivity problem, "
                f"not the workbook-session timeout the rest of the script already works around."
            )

        # === PHASE 1: DATA COLLECTION ===
        logger.info("=== DATA COLLECTION PHASE ===")
        
        # Enhanced concurrent data fetching
        df_bom = fetch_bom_data(SAP_API_URL, USERNAME, PASSWORD, VERIFY_SSL, PREFERRED_FORMAT, BATCH_SIZE, max_workers=15)

        # Fetch other data concurrently
        with ThreadPoolExecutor(max_workers=3) as executor:
            header_future = executor.submit(fetch_material_bom_headers, SAP_API_URL, USERNAME, PASSWORD, VERIFY_SSL, PREFERRED_FORMAT, BATCH_SIZE)
            stock_future = executor.submit(fetch_material_stock_data, SAP_API_URL, USERNAME, PASSWORD, VERIFY_SSL, PREFERRED_FORMAT, STOCK_BATCH_SIZE)
            groups_future = executor.submit(fetch_product_groups, USERNAME, PASSWORD, VERIFY_SSL)

            header_df = header_future.result()
            stock_df = summarise_plant_stock(stock_future.result())
            product_groups = groups_future.result()

        df_bom = standardize_bom_data(df_bom, product_groups)

        header_mapping = dict(zip(header_df["Material"], header_df["ProductDescription"]))
        header_qty_mapping = dict(zip(header_df["Material"], pd.to_numeric(header_df["BOMHeaderQuantityInBaseUnit"], errors="coerce").fillna(1)))
        
        logger.info(f"Data collection completed - BOM: {len(df_bom)}, Headers: {len(header_df)}, Stock: {len(stock_df)}")
        
        # === PHASE 2: BOM TREE STRUCTURE CREATION ===
        logger.info("=== BOM TREE STRUCTURE CREATION ===")
        
        # Create BOM tree structure (in memory, no local files)
        logger.info("Creating BOM layer-wise structure...")
        all_paths = []
        for fg in df_bom["Material"].unique():
            fg_paths = list(tree_explode(df_bom, fg, MAX_LAYERS))
            all_paths.extend(fg_paths)
        
        logger.info(f"BOM tree structure created with {len(all_paths)} paths")
        
        # === PHASE 3: SHAREPOINT EXCEL SESSION CREATION ===
        logger.info("=== SHAREPOINT EXCEL SESSION CREATION ===")
        
        # Get access token and create session
        access_token = get_access_token()

        # Try to open an explicit SharePoint Excel session (a short, bounded number of
        # attempts - not the full MAX_ATTEMPTS table - because for a workbook this large
        # createSession can time out on the server side every single time, and waiting
        # longer between identical attempts does not change that outcome).
        #
        # If it still fails, proceed WITHOUT a session: every call below then runs as its
        # own one-off transaction instead (see graph_headers()), which Microsoft Graph
        # supports natively and which still reads and writes the real file. This is the
        # expected, working path for very large workbooks where createSession itself is
        # unreliable but individual range reads/writes are not.
        session_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive/root:/{SHAREPOINT_FILE}:/workbook/createSession"
        setup_session = requests.Session()
        session_id = None
        SESSION_SETUP_ATTEMPTS, SESSION_SETUP_BACKOFF = 3, [10, 20]
        for attempt in range(1, SESSION_SETUP_ATTEMPTS + 1):
            try:
                resp = setup_session.post(
                    session_url, timeout=60,
                    headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
                    json={"persistChanges": True}
                )
                if resp.status_code in (200, 201):
                    session_id = resp.json().get("id")
                    logger.info(f"Created SharePoint Excel session: {session_id}")
                    break
                logger.warning(f"Create SharePoint Excel session: HTTP {resp.status_code} (attempt {attempt}/{SESSION_SETUP_ATTEMPTS})")
            except requests.exceptions.RequestException as e:
                logger.warning(f"Create SharePoint Excel session: {e} (attempt {attempt}/{SESSION_SETUP_ATTEMPTS})")
            if attempt < SESSION_SETUP_ATTEMPTS:
                time.sleep(SESSION_SETUP_BACKOFF[attempt - 1])
        if not session_id:
            logger.warning(f"Could not open an explicit Excel session for '{SHAREPOINT_FILE}' after "
                           f"{SESSION_SETUP_ATTEMPTS} attempts - this is expected for a workbook this "
                           f"size. Continuing WITHOUT a session: each write below runs as its own "
                           f"one-off transaction instead.")


        try:
            # === PHASE 4: SHAREPOINT SHEET CREATION ===
            logger.info("=== SHAREPOINT SHEET CREATION ===")
            
            # Each sheet must land before the next one is touched. v5 carried on after a failed
            # BOM Layer-wise write and built FG-RM Required from the stale sheet it read back.
            # The 'Last Updated' cell of a sheet is only written once that sheet is in.
            #
            # MB52 goes first: it does not depend on the BOM, and shrinking it from ~176k rows
            # takes most of the load off the workbook before the big BOM Layer-wise write.
            logger.info("Creating MB52 Raw Data sheet in SharePoint...")
            success_stock = create_sharepoint_mb52_sheet(SITE_ID, SHAREPOINT_FILE, session_id, access_token, stock_df)
            if not success_stock:
                raise RuntimeError("MB52 Raw Data failed to update - stopping before BOM Layer-wise and FG-RM Required")
            set_fixed_header_timestamp(SITE_ID, SHAREPOINT_FILE, "MB52 Raw Data", "I1", session_id, access_token)

            logger.info("Creating BOM Layer-wise sheet in SharePoint...")
            success_main, df_main_sheet = create_sharepoint_bom_layerwise_sheet(SITE_ID, SHAREPOINT_FILE, session_id, access_token, df_bom, all_paths, MAX_LAYERS, header_mapping, header_qty_mapping)
            if not success_main:
                raise RuntimeError("BOM Layer-wise failed to update (MB52 Raw Data was updated) - "
                                   "stopping before FG-RM Required so it stays in line with the previous BOM")
            # timestamp sits one empty column to the right of the data (AL1 with the full 36 columns)
            set_fixed_header_timestamp(SITE_ID, SHAREPOINT_FILE, "BOM Layer-wise",
                                       f"{get_excel_column_letter(len(df_main_sheet.columns) + 2)}1", session_id, access_token)

            logger.info("Creating FG-RM Required sheet in SharePoint...")
            success_fg_rm = create_sharepoint_fg_rm_sheet(SITE_ID, SHAREPOINT_FILE, session_id, access_token, df_main_sheet, MAX_LAYERS)
            if not success_fg_rm:
                raise RuntimeError("FG-RM Required failed to update (MB52 Raw Data and BOM Layer-wise were updated)")
            set_fixed_header_timestamp(SITE_ID, SHAREPOINT_FILE, "FG-RM Required", "F1", session_id, access_token)

            logger.info("Refresh the 'MB52 1000 Stock' pivot (Data > Refresh All) to pick up the new stock")

        finally:
            # Close the session, if one was ever opened. Running sessionless (see above)
            # means there is nothing to close - every call already closed itself.
            active = ACTIVE_SESSION["id"] or session_id
            if active:
                try:
                    close_session_url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drive/root:/{SHAREPOINT_FILE}:/workbook/closeSession"
                    close_headers = {"Authorization": f"Bearer {access_token}", "workbook-session-id": active}
                    close_response = requests.post(close_session_url, headers=close_headers, timeout=15)
                    if close_response.status_code in [200, 204]:
                        logger.info("SharePoint Excel session closed successfully")
                except Exception as close_error:
                    logger.warning(f"Error closing session: {close_error}")
        
        elapsed_time = time.time() - start_time
        logger.info(f"SHAREPOINT-ONLY FIFO allocation system completed successfully in {elapsed_time:.2f} seconds")
        
    except Exception as e:
        logger.error(f"Critical error in main process: {e}")
        import traceback
        logger.error(f"Full traceback: {traceback.format_exc()}")
        raise

if __name__ == "__main__":
    main()
