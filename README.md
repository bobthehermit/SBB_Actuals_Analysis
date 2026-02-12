# Actuals Analysis & Compliance

A Streamlit application for automating quarterly actuals reviews of New Mexico school district and charter school financial reports. Built for the NM Public Education Department, School Budget Bureau.

## What It Does

This app streamlines the quarterly compliance review process by:

- **Automated validation** of Revenue, Expenditure, and Cash reports against a 60+ step checklist
- **Cross-report reconciliation** — Cash Line 2 vs Revenue YTD, Cash Line 5 vs Expenditure YTD, with per-fund rollup logic
- **Flagging issues** — negative balances, FTE mismatches, forbidden object codes, budget overruns, burn rate outliers, encumbrance risk
- **Revenue ratio checks** — Impact Aid, Ad Valorem, Forest Reserve fund distribution compliance
- **Interactive checklist** — track progress, add notes per step, save/resume sessions
- **Export options**:
  - Word memo with findings and detail tables
  - Excel checklist tracker
  - HTML visual dashboard with Chart.js charts (revenue vs expenditure, function breakdown, salary by job class, program spending, etc.)

## Reports Used

The app expects three OBMS reports per district review:

| Report | Format | Key Fields |
|--------|--------|------------|
| Cash Report | Excel (.xlsx) with a "Summary" tab | Fund, Lines 1-12 |
| Revenue Report | CSV or Excel | Fund, Object, Function, Period Amount, YTD, Budget |
| Expenditure Report | CSV or Excel | Fund, Object, Function, JobClass, Program, Period, YTD, FTE, Budget, Encumbrance |

## Running Locally

```bash
pip install streamlit pandas openpyxl python-docx pillow xlsxwriter
streamlit run Actuals_Analysis_v2.py
```

## File Structure

```
Actuals_Analysis/
├── Actuals_Analysis_v2.py    # Main application
├── Actuals_Checklist.csv     # 60+ step review checklist
├── 300 DPI NM PED Logo JPEG.jpg
├── .gitignore
├── .streamlit/
│   └── config.toml           # Streamlit theme config
└── README.md
```

## Notes

- Data files (CSVs, Excel reports) are excluded from the repo via `.gitignore`
- Session progress can be saved/loaded as `.pkl` files from the sidebar
- The checklist steps align with SBB's quarterly review procedures per NMAC 6.20.2
