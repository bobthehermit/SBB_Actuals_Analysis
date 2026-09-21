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
- **Interactive checklist** with progress tracking and per-step notes, autosaved to the browser as you work, plus a one-click save file
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

Everything is built on one small **review snapshot** (JSON): checklist ticks, notes, typed inputs (Steps 6/7/47, enrollment), entity/FY/period, and the cash report Summary tab. Revenue/Expenditure pulled from OBMS are re-pulled on restore rather than stored.

**Autosave (automatic, per browser).** Every tick, note, or typed input writes the snapshot to the browser's `localStorage`. The bar pinned above the checklist shows **Progress · ✓ autosaved HH:MM:SS** — the checkmark appears only after the browser confirms the write. If the browser storage isn't responding, the bar and the sidebar say so plainly instead of claiming autosave is on. Up to 12 reviews are kept (oldest dropped); each tab updates only its own review's slot, so two tabs reviewing different schools don't overwrite each other. After a crash, refresh, or redeploy, pick the review under **Save / Resume Progress → Autosaved reviews in this browser** and click **Resume**. Autosave is tied to the browser and host; it does not follow you to another machine.

**No silent overwrites.** If you load a review (OBMS pull, upload, or save file) that this browser already has an autosave for, and the two differ, autosave pauses and asks: *Use the autosave* or *Keep what's on screen*. Nothing is overwritten until you choose.

**Save file (manual, one click).** **Download save file** sits next to the progress bar and always contains exactly what's on screen — it's rebuilt on every checklist change. Each download is a new file stamped with date and time (`Review_<Entity>_<FY>_<Period>_<YYYYMMDD-HHMMSS>.json`), so saving often is safe and the newest file sorts last. Restore with **Resume from a save file** in the sidebar. Reviews built from *uploaded* Revenue/Expenditure files carry those reports inside the save file too. A note you're still typing is applied when you click out of the box (or press Ctrl+Enter).

Older `.pkl` progress files still load through the same uploader. JSON is the new format because unpickling a file runs code from it, and JSON is readable in any text editor.

**Start a new review** clears the loaded data, checklist, and notes (including the upload boxes). Pulling a different entity or period from OBMS also starts clean — the previous review remains in autosave.

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
- `streamlit-js-eval` provides the browser autosave. If it is missing the app still runs, with autosave disabled and a warning in the save bar and sidebar.
- Streamlit ≥ 1.45 is required (`st.fragment`, `download_button(on_click="ignore")`).

## Performance notes

The app is structured to keep per-interaction cost low, which matters on small hosts (e.g. Streamlit Community Cloud's ~1 GB memory limit):

- Full-year OBMS parquet files are held with `st.cache_resource` (one shared object, not a copy per rerun). Treat them as read-only.
- Fiscal-year period/entity lists and per-entity row counts are cached by Drive file ID, so the big frames are not scanned on reruns.
- `run_all_validations` and `generate_analysis_summary` are `st.cache_data`-cached on their inputs; they only rerun when data, entity, period, or a typed input changes.
- The checklist is an `st.fragment`: ticking a box or typing a note reruns only the checklist, not the sidebar, validations, or dashboard. Inputs that feed validations (Step 47, enrollment) escalate to a full rerun when changed.
- The cash report Excel is parsed once per uploaded file, and exports are built on demand. The save file is small (no OBMS DataFrames), so it is rebuilt on each checklist render — that's what keeps it current.
- The save bar and autosave live inside the checklist fragment. Anything outside the fragment (the sidebar) does not rerun on a checkbox click, so save state shown there would go stale.

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

- District data files (CSV/Excel), save files (`.json` / legacy `.pkl`), and the local `venv/` are excluded from the repo via `.gitignore`
- The Review Period (Q1 vs Q2–Q4) auto-sets from the OBMS pull's reporting period
- Approved actuals become public record on New Mexico's Sunshine Portal (OpenBooks); the exported memo carries a standing notice
- The checklist steps align with SBB's quarterly review procedures per NMAC 6.20.2