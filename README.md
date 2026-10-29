# Marketplace audit tools

Two independent Python scripts recovered as a source snapshot, without seller data, credentials or private reports.

`ozon_cpc_campaign_audit.py` reads an explicitly selected Ozon CPC Excel export and writes an Excel analysis and text summary beside it (or under `--output-prefix`). It does not connect to Ozon or change campaigns. Its thresholds and suggestions are heuristics, not validated estimates of profit or campaign performance. Existing output files with the same name can be replaced; select a fresh prefix.

`wb_concurents.py` is an experimental optional live competitor lookup. It needs an explicitly configured `WB_TOKEN`, a CSV containing an `nmID` column (or `--fetch-all`), and network access to WB endpoints. No token file is read. Endpoint compatibility, authorization and live results were not tested. Matching scores are heuristics. It writes a CSV and does not edit listings or prices.

## Setup and offline verification

Python 3.10+; Tkinter is needed only for the optional Ozon GUI.

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python -m unittest -v test_offline
python ozon_cpc_campaign_audit.py synthetic.xlsx --output-prefix analysis
```

Supply your own Excel export or CSV locally. Neither seller inputs nor generated outputs belong in this repository. Configuration is intentionally empty; keep credentials outside source control. See VALIDATION.md for the actual check scope.
