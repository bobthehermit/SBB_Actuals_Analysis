# Actuals Analysis & Compliance

A Streamlit application for automating quarterly actuals reviews of New Mexico school district and charter school financial reports. Built for the NM Public Education Department, School Budget Bureau.

## What It Does

- **Automated validation** of Revenue, Expenditure, and Cash reports against a 60+ step checklist
- **Direct OBMS data pull** — Revenue and Expenditure reports load straight from the OBMS parquet registry (the same Google Drive data store the [OBMS Financial Explorer](https://huggingface.co/spaces/bobthehermit/OBMS-Financial-Explorer) reads). File uploads remain available as a fallback source.
- **Batch Portfolio Scan** — run the full check suite across an entire portfolio in one pass, with per-entity drill-down and one-click handoff into a single-entity review
- **Cross-report reconciliation** — Cash Line 2 vs Revenue YTD, Cash Line 5 vs Expenditure YTD, with per-fund rollup logic
- **Flagging** — negative balances, FTE mismatches, forbidden object codes, budget overruns, burn rate outliers, encumbrance risk
- **Revenue ratio checks** — Impact Aid, Ad Valorem, Forest Reserve fund distribution compliance
- **Enrollment Projection Outlook** — funded growth projection vs the 40-day actual count
- **Input guardrails** — entity/fiscal-year consistency checks, cash-report filename vs. entity mismatch warning, per-period row counts before pulling, refusal to pull unsubmitted periods
- **Interactive checklist** with progress tracking and per-step notes, autosaved as you work
- **Exports** — Word findings memo, Excel checklist tracker, HTML visual dashboard, district email text, and a filled district Feedback-Checklist workbook

## Data Sources

**Revenue & Expenditure** come from either source, selected in the sidebar:

1. **Pull directly from OBMS** (default) — reads `gdrive_manifest.json` for fiscal-year → Google Drive file IDs, loads the actuals and budget parquet for the selected year, and builds reports for the selected entity and reporting period. The manifest is fetched from the OBMS Data Explorer repo first (so new fiscal years appear automatically), with the local copy as fallback.
2. **Upload CSV/Excel files** — exports from the OBMS Financial Explorer's Actuals tab.

**Cash Reports** are always uploaded (Excel from the district's quarterly submission; the app reads the "Summary" tab). In batch mode, multiple cash reports can be uploaded at once and are matched to entities by filename.

| Report | Format | Key Fields |
|--------|--------|------------|
| Cash Report | Excel (.xlsx) with a "Summary" tab | Fund, Lines 1-12 |
| Revenue Report | Pulled from OBMS, or CSV/Excel | Fund, Object, Function, Period Amount, YTD, Budget |
| Expenditure Report | Pulled from OBMS, or CSV/Excel | Fund, Object, Function, JobClass, Program, Period, YTD, FTE, Budget, Encumbrance |

## Saving and Resuming a Review

There are two layers, designed so you never have to remember to save.

**Autosave (automatic, per browser).** Every time the review changes — a box ticked, a note typed, an input entered — the app writes a compact snapshot to the browser's `localStorage`: checklist status, notes, typed inputs, entity/FY/period, and the cash report Summary tab. Revenue and Expenditure are *not* stored; they are re-pulled from OBMS when you resume. Up to 12 reviews are kept (oldest dropped), so several schools can be in progress at once. After a crash, refresh, or redeploy, the sidebar offers the saved reviews under **Save / Resume Progress → Autosaved reviews in this browser**, with Resume and Delete. Autosave is tied to the browser and host it was created on; it does not follow you to another machine. Reviews built from *uploaded* Revenue/Expenditure files restore everything except those two reports (re-upload them).

**Portable save file (manual).** Click **Prepare save file**, then **Download Progress** to get a `.pkl` that contains the full review including the DataFrames. Use this to move a review between machines or file it alongside the district's folder. Restore it with **Resume a Previous Review**. The download button label tells you if the review has changed since the file was prepared.

**Start a new review** clears the loaded data, checklist, and notes. Pulling a different entity or period from OBMS also starts clean — the previous review remains in autosave.

## Exports

Click **Prepare download files** to build the Word memo, Excel tracker, and HTML report from the current state; the download buttons then appear. Exports are built only on request (not on every interaction), and the app warns when prepared files are older than the review.

## Running Locally

```bash
pip install -r requirements.txt
streamlit run Actuals_Analysis_v2.py
```

### Environment notes (important)

- **`pyarrow` must stay below 25** (pinned in `requirements.txt`). pyarrow 25.0.0 has a native bug that segfaults under Streamlit's dataframe serialization. Use `pip install "pyarrow==24.0.0"` if the local venv drifts.
- The app sets `pd.set_option("mode.string_storage", "python")` at startup to sidestep pandas 3.x Arrow-backed-string crashes inside Streamlit. Leave it in place.
- `streamlit-js-eval` provides the browser autosave. If it is missing the app still runs, with autosave disabled and a note in the sidebar.
- Streamlit ≥ 1.37 is required for `st.fragment`.

## Performance notes

The app is structured to keep per-interaction cost low, which matters on small hosts (e.g. Streamlit Community Cloud's ~1 GB memory limit):

- Full-year OBMS parquet files are held with `st.cache_resource` (one shared object, not a copy per rerun). Treat them as read-only.
- Fiscal-year period/entity lists and per-entity row counts are cached by Drive file ID, so the big frames are not scanned on reruns.
- `run_all_validations` and `generate_analysis_summary` are `st.cache_data`-cached on their inputs; they only rerun when data, entity, period, or a typed input changes.
- The checklist is an `st.fragment`: ticking a box or typing a note reruns only the checklist, not the sidebar, validations, or dashboard. Inputs that feed validations (Step 47, enrollment) escalate to a full rerun when changed.
- The cash report Excel is parsed once per uploaded file, and exports/pickles are built on demand.

## Deployment

Deployed on Streamlit Community Cloud from this repo — any push to `main` redeploys automatically. Parquet files must be publicly shared (view access) on Google Drive. To add a new fiscal year, add its file IDs to `gdrive_manifest.json` in the OBMS Data Explorer repo (this app picks it up automatically) or to the local copy here.

The app has no host-specific dependencies and runs unchanged on Hugging Face Spaces (Streamlit SDK). Autosave data is per-origin, so moving hosts starts with an empty autosave list.

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

- District data files (CSV/Excel), saved sessions (`.pkl`), and the local `venv/` are excluded from the repo via `.gitignore`
- The Review Period (Q1 vs Q2–Q4) auto-sets from the OBMS pull's reporting period
- Approved actuals become public record on New Mexico's Sunshine Portal (OpenBooks); the exported memo carries a standing notice
- The checklist steps align with SBB's quarterly review procedures per NMAC 6.20.2