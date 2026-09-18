# UChicago AI-Enabled Investing Competition

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
