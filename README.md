# BOM Layer-wise Manufacturing Plan — Indus International

Automated GitHub Actions pipeline that reads live Bill of Materials (BOM) data and current SAP inventory from SharePoint, performs layer-by-layer raw material requirement planning, and produces a formatted Excel manufacturing plan report.

## What It Does

- Reads multi-level BOM structure and current stock levels from SharePoint Excel via Microsoft Graph API
- Explodes BOM layer by layer (L1 finished goods → L2 sub-assemblies → L3 raw materials) to compute net material requirements
- Cross-references against live SAP inventory (MB52) to identify shortfalls
- Generates a formatted Excel output with colour-coded layers and requirement breakdowns
- Two parallel tracks: **live** (reads from SharePoint in real time) and **rm-planning** (raw material planning variant with procurement quantities)

## Script Versions

| File | Description |
|------|-------------|
| `BOM-Layer-wise-live.py` | v1 — live BOM explosion, inventory comparison |
| `BOM-Layer-wise-live-v2.py` | v2 — improved data handling and output formatting |
| `BOM-Layer-wise-live-rm-planning.py` | v1 — adds raw material procurement planning output |
| `BOM-Layer-wise-live-rm-planning-v2.py` | v2 — full RM planning |
| `BOM-Layer-wise-live-rm-planning-v6.py` | v6 — current active version (run by `rm_plan_bom_automation.yml`); Plant 1000 + 2000 stock summarised per plant (batches added together), no read-backs, header-only formatting, stops on the first failed sheet |

Use `BOM-Layer-wise-live-rm-planning-v6.py` for production runs.

## Workflow

```
SharePoint Excel (BOM + Stock data)
        ↓ Graph API
   Layer-wise BOM Explosion
        ↓
  Net Requirements Calculation
        ↓
  Formatted Excel Report Output
```

## Tech Stack

- **Python 3.11**
- `msal` — Microsoft OAuth2
- `requests` — Graph API
- `pandas` — BOM explosion logic
- `openpyxl` — Excel report generation
- **GitHub Actions** — scheduled execution

## Setup

```bash
pip install -r requirements.txt
```

Required GitHub Secrets:

```
CLIENT_ID
CLIENT_SECRET
TENANT_ID
SITE_ID
```

---
*Indus International FZC — Internal Automation Tool*
