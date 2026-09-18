# UChicago AI-Enabled Investing Competition

## Current deliverables: Lumentum valuation

- [Editable DCF workbook](deliverables/LITE_Bull_Case_DCF.xlsx)
- [Ten-slide DCF section](deliverables/LITE_DCF_Slides.pptx)
- [Presenter Q&A and above-consensus flags](deliverables/LITE_DCF_QA.md)

The current thesis horizon is **3–6 months**: a potential fear-discount reversal if earnings execution holds. The workbook explicitly separates that trading hypothesis from long-term intrinsic value. The conservative component WACC and mature-margin fade do **not** produce a bullish intrinsic value at the reference price; this result is disclosed, not overridden to hit a target. The $1,149 external analyst average is a 12-month cross-check.

Build the final model and slides with `.venv\Scripts\python.exe src\finish_lite_dcf.py`. The engine is `src/lite_dcf_model.py`. `src/recalculate_and_export.ps1` recalculates in Excel and exports the PowerPoint PDF if desktop Office is available.

Earlier FN research, `src/model.py`, `src/build_workbook.py`, and `src/build_lumentum.py` are superseded exploratory drafts. Do not run the earlier builders for the final deliverables. The user's `slideshow/Round1_Pitch.pptx` template is preserved separately.

---

## Earlier research context (superseded)

Investment research, an editable DCF workbook, and a 10-slide presentation for the September 18, 2026 competition. Investment horizon: 3–12 months.

## Build status

- Public filings and research sources are being archived under `research/sources`.
- The initial Fabrinet long hypothesis is being tested against cash generation, tax expense, and capacity spending. Direction and target will be determined by the completed model.
- Final files will be placed in `deliverables/` with an AI-use log and presenter notes.
- These materials are AI-assisted. Human team review and ownership of the investment judgment remain necessary; the log will identify the actual work performed rather than claim human verification that has not occurred.

## Research controls

- Freeze a dated reference price; do not mix live quotes with historical snapshots.
- Distinguish reported facts, management guidance, outside estimates, and model assumptions.
- Cross-check major figures against the original filings.
- Forecast cash taxes, working capital, and capital spending explicitly.
- Preserve contrary evidence and record corrections in the AI-use log.

## Environment

Python 3.12 in `.venv`. Build dependencies: XlsxWriter, openpyxl, python-pptx, PyMuPDF, ReportLab, matplotlib, requests, and Beautiful Soup.

`src/fetch_sources.py` archives the public documents. Failed downloads are recorded in its manifest instead of silently treated as successful.
