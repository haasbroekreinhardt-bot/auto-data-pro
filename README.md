# Auto-Data Pro — Excel Automation Dashboard

A Streamlit dashboard that takes any messy `.xlsx`, `.xls` or `.csv` export,
cleans and validates it, and produces a formatted multi-tab Excel workbook plus
a summary PDF.

Drop a file in, press **Process Data File**, review the flagged rows, download
the report.

---

## What it does

**Two engines, picked automatically from the file's columns.**

If the upload resolves to a client-transaction ledger (a money column plus at
least two of date / client ID / status / reference), it runs a 12-rule South
African business engine:

| | Rule |
|---|---|
| R01 | Whitespace and control-character trim |
| R02 | Blank row removal |
| R03 | Date standardisation to `YYYY-MM-DD` |
| R04 | Invalid / unparseable date flagging |
| R05 | Client ID normalisation |
| R06 | Missing Client ID flag |
| R07 | Status vocabulary normalisation |
| R08 | Blank status flag |
| R09 | ZAR currency parsing |
| R10 | Zero / negative amount flag |
| R11 | VAT at 15% and Total Inc VAT |
| R12 | Bank reconciliation and duplicate check |

Anything else falls through to a **universal profiler** that assumes no fixed
schema: it categorises every column as Numeric, Datetime, Categorical, Text/ID
or Empty, then audits for nulls, duplicates, unparseable dates, non-numeric
values in numeric columns, severe outliers, constant columns and inconsistent
label casing — scored into a dataset health percentage.

**Exports.** The workbook ships with a navy 26pt header row, Calibri 10 body at
20pt, zebra striping, currency and integer number formats, soft-red flagged
rows, centred dates and IDs, KPI cards and double-ruled grand totals. The PDF
covers dataset dimensions, column types, a missing-value audit and summary
statistics.

---

## Running it locally

**Windows, one click:** double-click `run_app.bat`. It finds Python, installs
dependencies on first run, and opens the dashboard at `http://localhost:8501`.

**Any platform:**

```bash
pip install -r requirements.txt
streamlit run app.py
```

Requires Python 3.9+.

No file you upload ever leaves your machine when running locally — everything is
processed in-process and the outputs are generated in memory.

---

## Trying it without a file

Press **Load sample M&M Retail dataset** on the Dashboard for a deliberately
messy synthetic ledger. Two sample files are also included:

- `sample_sales.xlsx` — a small ledger that exercises the business-rules engine
- `sample_generic.xlsx` — a non-ledger spreadsheet that exercises the universal
  profiler

Both are synthetic. No real customer data is in this repository.

---

## Configuration

Everything is adjustable under **Rule Settings** in the sidebar: VAT rate,
currency symbol (`R`, `$`, `€`, `£` or none), day-first vs month-first date
parsing, which validations raise flags, duplicate handling, and outlier
sensitivity.

`.streamlit/config.toml` pins the light theme so the dashboard renders
identically regardless of the viewer's dark-mode preference.

---

## Notes

- Streamlit's native data grid and charts need `pyarrow`. Where that is
  unavailable — some locked-down Windows builds block its DLL via Application
  Control policy — the app falls back to self-contained HTML tables and
  inline-SVG charts. Every feature keeps working; only in-grid sorting is lost.
- The PDF writer is dependency-free, so the requirements stay at four packages.

## Licence

No licence has been set. All rights reserved by default — add one if you intend
others to reuse this.
