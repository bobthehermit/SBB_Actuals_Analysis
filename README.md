# Actuals Analysis & Compliance

A Streamlit application for automating quarterly actuals reviews of New Mexico school district and charter school financial reports. Built for the NM Public Education Department, School Budget Bureau.

## What It Does

This app streamlines the quarterly compliance review process by:

- **Automated validation** of Revenue, Expenditure, and Cash reports against a 60+ step checklist
- **Direct OBMS data pull** — Revenue and Expenditure reports load straight from the OBMS parquet registry (the same Google Drive data store the [OBMS Financial Explorer](https://huggingface.co/spaces/bobthehermit/OBMS-Financial-Explorer) reads), no CSV round trip. File uploads remain available as a fallback source.
- **Batch Portfolio Scan** — run the full check suite across an entire portfolio of entities in one pass: submission status, revenue/expenditure/cash totals, and flag counts in a single attention-sorted table, with per-entity drill-down and one-click handoff into a full single-entity review
- **Cross-report reconciliation** — Cash Line 2 vs Revenue YTD, Cash Line 5 vs Expenditure YTD, with per-fund rollup logic
- **Flagging issues** — negative balances, FTE mismatches, forbidden object codes, budget overruns, burn rate outliers, encumbrance risk
- **Revenue ratio checks** — Impact Aid, Ad Valorem, Forest Reserve fund distribution compliance
- **Enrollment Projection Outlook** — compare the funded growth projection against the 40-day actual count; shortfalls flag with mid-year SEG adjustment and cash flow implications
- **Input guardrails** — entity/fiscal-year consistency checks across reports, a cash-report filename vs. review-entity mismatch warning, per-period actuals row counts shown before pulling, and refusal to pull empty (unsubmitted) periods rather than generating noise findings
- **Interactive checklist** — track progress, add notes per step, save/resume sessions
- **Export options**:
  - Word memo with findings, detail tables, and an OpenBooks public-record notice
  - Excel checklist tracker
  - Batch portfolio summary (Excel)
  - HTML visual dashboard with Chart.js charts (revenue vs expenditure, function breakdown, salary by job class, program spending, etc.)

## Data Sources

**Revenue & Expenditure** come from either source, selected in the sidebar:

1. **Pull directly from OBMS** (default) — reads `gdrive_manifest.json` for fiscal-year → Google Drive file IDs, downloads the actuals and budget parquet for the selected year (cached ~1 hour), and builds reports for the selected entity and reporting period. The manifest is fetched from the OBMS Data Explorer repo first (so new fiscal years appear automatically), with the local copy as fallback.
2. **Upload CSV/Excel files** — exports from the OBMS Financial Explorer's Actuals tab, as before.

**Cash Reports** are always uploaded (Excel from the district's quarterly submission; the app reads the "Summary" tab). In batch mode, multiple cash reports can be uploaded at once and are matched to entities by filename.

| Report | Format | Key Fields |
|--------|--------|------------|
| Cash Report | Excel (.xlsx) with a "Summary" tab | Fund, Lines 1-12 |
| Revenue Report | Pulled from OBMS, or CSV/Excel | Fund, Object, Function, Period Amount, YTD, Budget |
| Expenditure Report | Pulled from OBMS, or CSV/Excel | Fund, Object, Function, JobClass, Program, Period, YTD, FTE, Budget, Encumbrance |

## Running Locally

```bash
pip install -r requirements.txt
streamlit run Actuals_Analysis_v2.py
```

### Environment notes (important)

- **`pyarrow` must stay below 25** (pinned in `requirements.txt`). pyarrow 25.0.0 has a native bug that segfaults ("Python quit unexpectedly", `zsh: segmentation fault`) under Streamlit's dataframe serialization — reproduced across pandas 2.x and 3.x. pyarrow 24.0.0 is stable. Note that `pip install -U pyarrow` will happily reinstall the broken newest version; use `pip install "pyarrow==24.0.0"` if the local venv drifts.
- The app also sets `pd.set_option("mode.string_storage", "python")` at startup, which sidesteps a class of pandas 3.x Arrow-backed-string crashes inside Streamlit's script-runner thread. Leave it in place.

## Deployment

Deployed on Streamlit Community Cloud from this repo — any push to `main` redeploys automatically. Parquet files must be publicly shared (view access) on Google Drive. To add a new fiscal year, add its file IDs to `gdrive_manifest.json` in the OBMS Data Explorer repo (this app picks it up automatically) or to the local copy here.

## File Structure

```
SBB_Actuals_Analysis/
├── Actuals_Analysis_v2.py    # Main application
├── Actuals_Checklist.csv     # 60+ step review checklist
├── gdrive_manifest.json      # FY → Google Drive parquet file IDs (fallback copy)
├── requirements.txt          # Includes the pyarrow<25 pin — see Environment notes
├── 300 DPI NM PED Logo JPEG.jpg
├── .gitignore
├── .streamlit/
│   └── config.toml           # Streamlit theme config (NMPED palette)
└── README.md
```

## Notes

- Data files (CSVs, Excel reports) are excluded from the repo via `.gitignore`
- Session progress can be saved/loaded as `.pkl` files from the sidebar
- The Review Period (Q1 vs Q2–Q4) auto-sets from the OBMS pull's reporting period
- Approved actuals become public record on New Mexico's Sunshine Portal (OpenBooks); the exported memo carries a standing notice
- The checklist steps align with SBB's quarterly review procedures per NMAC 6.20.2