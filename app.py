"""
AUTO-DATA PRO: EXCEL AUTOMATION DASHBOARD
=========================================
A local Streamlit desktop application that ingests raw monthly sales exports
(.xlsx / .xls / .csv), sanitises them, applies a South African business-rules
engine (15% VAT, bank reconciliation, missing-field flagging) and emits a
multi-tab, fully formatted Excel workbook plus a summary PDF.

Client context : M&M Retail (South Africa)
Currency       : ZAR (South African Rand)
Run locally    : streamlit run app.py   (or double-click run_app.bat)
"""

from __future__ import annotations

import io
import re
import time
from datetime import datetime

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------- #
#  APPLICATION CONSTANTS
# --------------------------------------------------------------------------- #

# Streamlit renders st.dataframe and its native charts through pyarrow. On locked
# down corporate Windows builds the pyarrow DLL is frequently blocked by an
# Application Control / WDAC policy, which would otherwise crash the whole page,
# so every table and chart falls back to a self-contained HTML renderer.
try:
    import pyarrow as _pyarrow  # noqa: F401
    ARROW_AVAILABLE = True
    ARROW_ERROR = ""
except Exception as _arrow_exc:      # pragma: no cover - environment dependent
    ARROW_AVAILABLE = False
    ARROW_ERROR = str(_arrow_exc)


APP_TITLE = "AUTO-DATA PRO: EXCEL AUTOMATION DASHBOARD"
APP_SUBTITLE = ("Streamlining Monthly Sales Data  •  Automated Data Engine • Enterprise Edition"
                "  •  Client: M&M Retail")
BRAND_NAME = "Auto-Data Pro"
VERSION = "1.0.0"

ENGINE_LABEL = "Rules Engine v2.4 (Active)"

VAT_RATE_DEFAULT = 0.15          # South African standard VAT rate

# Currency is a display setting, not a hard-coded assumption. "None" renders
# plain numbers, which is the right choice for non-financial spreadsheets.
CURRENCY_OPTIONS = {
    "R  (South African Rand)": "R",
    "$  (Dollar)": "$",
    "€  (Euro)": "€",
    "£  (Pound)": "£",
    "None (plain numbers)": "",
}
CURRENCY_DEFAULT = "R  (South African Rand)"

NAVY = "0F172A"                  # Slate Navy - Excel header fill / brand primary
RED_FILL_HEX = "FEF2F2"          # soft red fill for flagged rows (text #991B1B)

NAV_ITEMS = [
    ("Dashboard", "Overview, upload, process, export"),
    ("Upload Data", "Stage a raw monthly file"),
    ("Rule Settings", "Configure the business-rules engine"),
    ("Analysis", "Revenue, VAT and status breakdowns"),
    ("History", "Previous automation runs"),
    ("Support", "Docs, diagnostics and contact"),
]

# The 12 business rules applied by the automation engine.
RULE_LIBRARY = [
    ("R01", "Whitespace & control-character trim", True),
    ("R02", "Blank row removal", True),
    ("R03", "Date standardisation (YYYY-MM-DD)", True),
    ("R04", "Invalid / unparseable date flagging", True),
    ("R05", "Client ID normalisation", True),
    ("R06", "Missing Client ID flag", True),
    ("R07", "Status vocabulary normalisation", True),
    ("R08", "Missing / unknown status flag", True),
    ("R09", "ZAR currency parsing (strip R, spaces, commas)", True),
    ("R10", "Zero / negative amount flag", True),
    ("R11", "VAT @ 15% and Total Inc VAT calculation", True),
    ("R12", "Bank reconciliation & duplicate-entry check", True),
]

# Canonical column -> accepted raw header aliases (normalised, lowercase).
COLUMN_ALIASES = {
    "Date": [
        "date", "invoice_date", "transaction_date", "txn_date", "trans_date",
        "doc_date", "document_date", "sale_date", "datum", "posting_date",
    ],
    "Client_ID": [
        "client_id", "clientid", "client", "client_code", "customer_id",
        "customerid", "customer", "customer_code", "account", "account_no",
        "account_number", "acc_no", "debtor", "debtor_code",
    ],
    "Status": [
        "status", "payment_status", "pay_status", "state", "invoice_status",
        "settlement_status", "category", "type", "channel", "industry",
        "segment", "business_type", "customer_type", "sector", "division",
    ],
    "Amount": [
        "amount", "amount_zar", "amount_r", "value", "total", "revenue",
        "sales", "net", "net_amount", "nett", "nett_amount", "amt", "price",
        "line_total", "excl_vat", "amount_excl_vat",
    ],
    "Notes": ["notes", "note", "comment", "comments", "remarks", "memo"],
    "Reference": [
        "reference", "ref", "bank_ref", "bank_reference", "invoice",
        "invoice_no", "invoice_number", "doc_no", "document_no",
        "payment_ref", "payment_reference", "eft_ref",
    ],
}

# Raw status token -> canonical status.
STATUS_MAP = {
    "paid": "PAID", "settled": "PAID", "complete": "PAID", "completed": "PAID",
    "received": "PAID", "cleared": "PAID", "closed": "PAID", "p": "PAID",
    "pending": "PENDING", "outstanding": "PENDING", "unpaid": "PENDING",
    "open": "PENDING", "awaiting": "PENDING", "in_progress": "PENDING",
    "processing": "PENDING",
    "overdue": "OVERDUE", "late": "OVERDUE", "arrears": "OVERDUE",
    "past_due": "OVERDUE", "delinquent": "OVERDUE",
    "cancelled": "CANCELLED", "canceled": "CANCELLED", "void": "CANCELLED",
    "voided": "CANCELLED", "refund": "CANCELLED", "refunded": "CANCELLED",
    "reversed": "CANCELLED", "credit_note": "CANCELLED",
}

# Canonical payment statuses, in report order. Any other non-empty value found
# in the source file is preserved verbatim and sorted after these.
STATUS_ORDER = ["PAID", "PENDING", "OVERDUE", "CANCELLED", "UNASSIGNED"]
# One label for an empty status, used by both the Status badge and the Notes
# tag so the two never disagree about what an unknown status is called.
STATUS_BLANK = "UNASSIGNED"

FLAG_MISSING_ID = "MISSING ID"
FLAG_MISSING_STATUS = "UNASSIGNED"      # matches STATUS_BLANK by design
# Legacy flag: no longer raised (a non-empty status is now always preserved),
# but retained so a previously exported file re-imported here still renders it.
FLAG_UNKNOWN_STATUS = "UNKNOWN STATUS"
FLAG_BAD_DATE = "INVALID DATE"
FLAG_BAD_AMOUNT = "INVALID AMOUNT"
FLAG_NEGATIVE = "NEGATIVE AMOUNT"
FLAG_ZERO = "ZERO AMOUNT"
FLAG_UNRECONCILED = "UNRECONCILED"

# A row carrying any of these words is business-flagged for attention: it is
# not a validation error, but it must not read as CLEAN either.
REVIEW_KEYWORDS = ("action required", "urgent", "query")
REVIEW_LABEL = "NEEDS REVIEW"

RECON_OK = "Reconciled"           # a bank reference is present on the row
RECON_NONE = "Unreconciled"       # reference column exists but this row is blank
RECON_SKIPPED = "Not Checked"     # source file carries no reference column at all

# Internal bookkeeping columns, never written to an export or preview.
HELPER_COLUMNS = ["Flagged", "Flag_Count", "Needs_Review"]
FLAG_RECON_MISMATCH = "RECON MISMATCH"
FLAG_DUPLICATE = "DUPLICATE ENTRY"
FLAG_BELOW_THRESHOLD = "BELOW THRESHOLD"

# Every validation flag the engine can raise - used to render coral pills in the
# preview table and to separate flags from free-text notes carried in from source.
KNOWN_FLAGS = {
    FLAG_MISSING_ID, FLAG_MISSING_STATUS, FLAG_UNKNOWN_STATUS, FLAG_BAD_DATE,
    FLAG_BAD_AMOUNT, FLAG_NEGATIVE, FLAG_ZERO, FLAG_UNRECONCILED,
    FLAG_RECON_MISMATCH, FLAG_DUPLICATE, FLAG_BELOW_THRESHOLD,
}

OUTPUT_COLUMNS = [
    "Date", "Client_ID", "Reference", "Recon_Status", "Status", "Amount",
    "VAT_15", "Total_Inc_VAT", "Notes",
]

st.set_page_config(
    page_title="Auto-Data Pro | Excel Automation Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)


# --------------------------------------------------------------------------- #
#  CUSTOM CSS  —  enterprise SaaS light theme
# --------------------------------------------------------------------------- #

CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

/* ======================================================================
   ENTERPRISE PALETTE
   Slate Navy #0F172A  ·  Off-White #F8FAFC
   Soft Emerald #10B981 ·  Soft Coral #EF4444  ·  Border #E2E8F0
   ====================================================================== */
:root {
    --adp-navy:       #0F172A;   /* Slate Navy - primary */
    --adp-navy-2:     #1E293B;
    --adp-navy-3:     #334155;
    --adp-ink:        #1E293B;   /* high-contrast body text */
    --adp-slate:      #1E293B;
    --adp-muted:      #475569;   /* high-contrast secondary text */
    --adp-line:       #E2E8F0;   /* subtle light grey border */
    --adp-bg:         #F8FAFC;   /* Off-White background */
    --adp-card:       #FFFFFF;
    --adp-green:      #10B981;   /* Soft Emerald - clean status */
    --adp-green-deep: #047857;
    --adp-green-bg:   #D1FAE5;
    --adp-red:        #EF4444;   /* Soft Coral - error flags */
    --adp-red-deep:   #B91C1C;
    --adp-red-bg:     #FEE2E2;
    --adp-red-tint:   #FEF2F2;
    --adp-amber:      #D97706;
    --adp-amber-bg:   #FEF3C7;
    --adp-radius:     10px;
    --adp-shadow:     0 1px 2px rgba(15, 23, 42, 0.04), 0 4px 12px rgba(15, 23, 42, 0.06);
    --adp-shadow-lg:  0 10px 26px rgba(15, 23, 42, 0.18);
}

html, body, [class*="css"], .stApp { font-family: 'Inter', 'Segoe UI', sans-serif; }
.stApp { background: var(--adp-bg); }

/* ---------------- Hide the stock Streamlit chrome ---------------------- */
#MainMenu { visibility: hidden; }
header    { visibility: hidden; height: 0 !important; }
footer    { visibility: hidden; height: 0 !important; }
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
[data-testid="stHeader"] { display: none !important; }
.block-container { padding-top: 1.6rem; padding-bottom: 3rem; max-width: 1500px; }

/* ---------------- Header bar ------------------------------------------ */
.adp-header {
    display: flex; align-items: center; justify-content: space-between;
    background: linear-gradient(105deg, #0F172A 0%, #1E293B 58%, #334155 100%);
    border-radius: var(--adp-radius); padding: 22px 30px; margin-bottom: 22px;
    box-shadow: var(--adp-shadow-lg);
}
.adp-header-title {
    color: #FFFFFF; font-size: 25px; font-weight: 800;
    letter-spacing: 0.6px; line-height: 1.15; margin: 0;
}
.adp-header-sub { color: #CBD5E1; font-size: 13px; font-weight: 500; margin-top: 7px; }
.adp-logo {
    display: flex; align-items: center; gap: 12px;
    background: rgba(255,255,255,0.08); border: 1px solid rgba(255,255,255,0.18);
    border-radius: var(--adp-radius); padding: 12px 18px;
}
.adp-logo-mark {
    width: 40px; height: 40px; border-radius: var(--adp-radius); flex: none;
    background: linear-gradient(135deg, #10B981 0%, #047857 100%);
    color: #fff; font-weight: 800; font-size: 17px;
    display: flex; align-items: center; justify-content: center;
    box-shadow: 0 4px 12px rgba(16, 185, 129, 0.38);
}
.adp-logo-text { color: #FFFFFF; font-weight: 700; font-size: 16px; line-height: 1.1; }
.adp-logo-tag  { color: #94A3B8; font-size: 10.5px; letter-spacing: 1.4px; font-weight: 600; }

/* ---------------- Sidebar --------------------------------------------- */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0F172A 0%, #172033 100%);
    border-right: 1px solid #020617;
}
[data-testid="stSidebar"] * { color: #E2E8F0 !important; }
[data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 { color: #FFFFFF !important; }
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,0.12); }
[data-testid="stSidebar"] [role="radiogroup"] label {
    background: rgba(255,255,255,0.04); border: 1px solid rgba(255,255,255,0.08);
    border-radius: var(--adp-radius); padding: 9px 12px; margin-bottom: 7px; width: 100%;
    transition: all .15s ease-in-out;
}
[data-testid="stSidebar"] [role="radiogroup"] label:hover {
    background: rgba(255,255,255,0.10); border-color: rgba(16, 185, 129, 0.55);
}
.adp-side-brand {
    display: flex; align-items: center; gap: 11px; padding: 4px 2px 16px 2px;
    border-bottom: 1px solid rgba(255,255,255,0.12); margin-bottom: 16px;
}
.adp-side-env {
    background: rgba(16, 185, 129, 0.10); border: 1px solid rgba(16, 185, 129, 0.32);
    border-radius: var(--adp-radius); padding: 11px 13px;
    font-size: 11.5px; line-height: 1.75; margin-top: 14px;
}
.adp-dot { height: 8px; width: 8px; border-radius: 50%; display: inline-block; margin-right: 7px; }
.adp-dot-green { background: #10B981; box-shadow: 0 0 7px #10B981; }
.adp-dot-amber { background: #F59E0B; box-shadow: 0 0 7px #F59E0B; }

/* ---------------- Cards / panels --------------------------------------- */
.adp-panel {
    background: var(--adp-card); border: 1px solid var(--adp-line);
    border-radius: var(--adp-radius); padding: 20px 22px 8px 22px;
    margin-bottom: 18px; box-shadow: var(--adp-shadow);
}
.adp-panel-title {
    font-size: 13px; font-weight: 800; letter-spacing: 1.1px;
    color: var(--adp-navy); text-transform: uppercase;
    border-left: 4px solid var(--adp-green); padding-left: 11px; margin: 0 0 4px 0;
}
.adp-panel-note { font-size: 12.5px; color: var(--adp-muted); margin: 6px 0 14px 15px; }

/* ---------------- Metric cards ----------------------------------------- */
.adp-metric {
    background: var(--adp-card); border: 1px solid var(--adp-line);
    border-left: 5px solid var(--adp-navy); border-radius: var(--adp-radius);
    padding: 17px 19px; height: 100%; box-shadow: var(--adp-shadow);
}
.adp-metric.green { border-left-color: var(--adp-green); }
.adp-metric.red   { border-left-color: var(--adp-red); }
.adp-metric-label {
    font-size: 10.5px; font-weight: 700; letter-spacing: 1.1px;
    color: var(--adp-muted); text-transform: uppercase;
}
.adp-metric-value {
    font-size: 30px; font-weight: 800; color: var(--adp-navy);
    margin-top: 7px; line-height: 1.1; font-variant-numeric: tabular-nums;
}
.adp-metric-value.green { color: var(--adp-green-deep); }
.adp-metric-value.red   { color: var(--adp-red-deep); }
.adp-metric-foot { font-size: 11.5px; color: var(--adp-muted); margin-top: 7px; }

/* ---------------- Pill badges ------------------------------------------ */
.adp-badge {
    display: inline-block; border-radius: 999px; padding: 3px 11px;
    font-size: 10.5px; font-weight: 700; letter-spacing: .5px;
}
.adp-badge-red   { background: var(--adp-red-bg);   color: var(--adp-red-deep); }
.adp-badge-green { background: var(--adp-green-bg); color: var(--adp-green-deep); }
.adp-badge-amber { background: var(--adp-amber-bg); color: #92400E; }
.adp-badge-navy  { background: #E2E8F0; color: var(--adp-navy); }
.adp-badge-slate { background: #F1F5F9; color: var(--adp-muted); }

/* ---------------- File status strip ------------------------------------ */
.adp-file {
    display: flex; align-items: center; justify-content: space-between;
    border-radius: var(--adp-radius); padding: 13px 17px; margin-top: 12px; font-size: 13px;
}
.adp-file-ok   { background: #ECFDF5; border: 1px solid #6EE7B7; color: #065F46; }
.adp-file-idle { background: #F8FAFC; border: 1px dashed #CBD5E1; color: var(--adp-muted); }
.adp-file-name { font-weight: 700; }
.adp-file-meta { font-size: 12px; opacity: .9; margin-top: 2px; }


/* ---------------- Preview table ---------------------------------------- */
.adp-table-wrap {
    border: 1px solid var(--adp-line); border-radius: var(--adp-radius);
    overflow: auto; max-height: 520px; background: var(--adp-card);
    box-shadow: var(--adp-shadow);
}
/* Separates the preview grid from the filter row sitting above it. */
.adp-table-wrap.adp-table-preview { margin-top: 16px; }
table.adp-table { border-collapse: collapse; width: 100%; font-size: 12.6px; }
table.adp-table thead th {
    position: sticky; top: 0; z-index: 2;
    background: var(--adp-navy); color: #FFFFFF; text-align: left;
    padding: 11px 13px; font-weight: 700; font-size: 11px;
    letter-spacing: .8px; text-transform: uppercase; white-space: nowrap;
}
table.adp-table tbody td {
    padding: 9px 13px; color: var(--adp-ink); white-space: nowrap;
    border-bottom: 1px solid var(--adp-line);
    border-right: 1px solid var(--adp-line);
}
table.adp-table tbody td:last-child,
table.adp-table thead th:last-child { border-right: none; }
table.adp-table thead th { border-right: 1px solid rgba(255,255,255,0.12); }
/* Row state is carried by the background tint and the badge pills alone -
   body text stays uniform slate so dates and IDs never read as errors. */
table.adp-table tbody tr.flagged td { background: var(--adp-red-tint); }
table.adp-table tbody tr.review td  { background: #FFFBEB; }
table.adp-table tbody tr:hover td   { background: #F1F5F9; }
table.adp-table tbody tr.flagged:hover td { background: var(--adp-red-bg); }
table.adp-table tbody tr.review:hover td  { background: var(--adp-amber-bg); }
table.adp-table td.num { text-align: right; font-variant-numeric: tabular-nums; font-weight: 600; }

/* Pill-shaped status / condition tags */
.adp-pill {
    display: inline-block; border-radius: 999px; padding: 2px 10px;
    font-size: 10.5px; font-weight: 700; letter-spacing: .4px; white-space: nowrap;
}
.adp-pill + .adp-pill { margin-left: 4px; }
.adp-pill-CLEAN     { background: var(--adp-green-bg); color: var(--adp-green-deep); }
.adp-pill-FLAG      { background: var(--adp-red-bg);   color: var(--adp-red-deep); }
.adp-pill-PAID      { background: var(--adp-green-bg); color: var(--adp-green-deep); }
.adp-pill-PENDING   { background: var(--adp-amber-bg); color: #92400E; }
.adp-pill-OVERDUE   { background: var(--adp-red-bg);   color: var(--adp-red-deep); }
.adp-pill-CANCELLED { background: #E2E8F0;             color: var(--adp-muted); }
.adp-pill-UNASSIGNED { background: #FDE68A;            color: #78350F; }
.adp-pill-CUSTOM    { background: #E2E8F0;             color: var(--adp-navy); }
/* Soft amber - the row is not an error but does need a human to look at it. */
.adp-pill-REVIEW    { background: var(--adp-amber-bg); color: #92400E; }
/* Free-text notes carried in from the source file are plain text, not pills,
   so that badges stay reserved for status and validation flags. */
.adp-note-text {
    color: var(--adp-ink); font-weight: 400; font-size: 12.2px;
    white-space: nowrap;
}
table.adp-table td.notes-cell { text-align: left; }
table.adp-table td.notes-cell .adp-pill + .adp-note-text { margin-left: 6px; }

/* ---------------- Buttons ---------------------------------------------- */
.stButton > button, .stDownloadButton > button {
    border-radius: var(--adp-radius); font-weight: 700; letter-spacing: .4px;
    padding: 0.62rem 1.1rem; border: 1px solid var(--adp-line);
    background: var(--adp-card); color: var(--adp-navy); transition: all .15s ease-in-out;
}
.stButton > button:hover, .stDownloadButton > button:hover {
    border-color: var(--adp-navy); color: var(--adp-navy); background: #F1F5F9;
}
.stButton > button[kind="primary"],
.stButton > button[data-testid="stBaseButton-primary"],
.stDownloadButton > button[kind="primary"],
.stDownloadButton > button[data-testid="stBaseButton-primary"] {
    background: linear-gradient(94deg, #10B981 0%, #047857 100%);
    color: #FFFFFF !important; border: none;
    box-shadow: 0 5px 15px rgba(16, 185, 129, 0.30);
    text-transform: uppercase; font-size: 13px;
}
.stButton > button[kind="primary"]:hover,
.stDownloadButton > button[kind="primary"]:hover {
    background: linear-gradient(94deg, #047857 0%, #065F46 100%);
    box-shadow: 0 7px 19px rgba(16, 185, 129, 0.40);
}
[data-testid="stSidebar"] .stButton > button {
    background: rgba(255,255,255,0.07); border: 1px solid rgba(255,255,255,0.16);
    color: #FFFFFF !important;
}

/* Secondary download button - same height, weight, radius and transition as the
   primary action, but a light surface with dark text so the label stays legible.
   Keyed on st-key-dl_pdf (Streamlit stamps the widget key onto its container). */
.st-key-dl_pdf .stDownloadButton > button {
    background: #FFFFFF;
    color: var(--adp-navy) !important;
    border: none;
    border-radius: var(--adp-radius);
    padding: 0.62rem 1.1rem;
    font-weight: 700; font-size: 13px; letter-spacing: .4px;
    text-transform: uppercase;
    box-shadow: inset 0 0 0 1px #CBD5E1, 0 1px 2px rgba(15, 23, 42, 0.06);
    transition: all .15s ease-in-out;
}
.st-key-dl_pdf .stDownloadButton > button:hover {
    background: #F1F5F9;
    box-shadow: inset 0 0 0 1px var(--adp-navy), 0 4px 12px rgba(15, 23, 42, 0.10);
    color: var(--adp-navy) !important;
}

/* Hover tooltip for the export details, so the caption can stay one line. */
.adp-tip {
    border-bottom: 1px dotted var(--adp-muted); cursor: help;
    color: var(--adp-muted);
}

/* ---------------- Uploader --------------------------------------------- */
[data-testid="stFileUploaderDropzone"] {
    background: #F8FAFC; border: 2px dashed #94A3B8;
    border-radius: var(--adp-radius); padding: 22px;
}
[data-testid="stFileUploaderDropzone"]:hover { border-color: var(--adp-green); background: #ECFDF5; }

/* ---------------- Misc -------------------------------------------------- */
.adp-section-h {
    font-size: 12px; font-weight: 800; letter-spacing: 1px; color: var(--adp-muted);
    text-transform: uppercase; margin: 4px 0 10px 0;
}
.adp-kv { font-size: 13px; color: var(--adp-ink); line-height: 2.0; }
.adp-kv b { color: var(--adp-navy); }
.stTabs [data-baseweb="tab"] { border-radius: 8px 8px 0 0; padding: 8px 16px; font-weight: 600; }
hr { margin: 10px 0 18px 0; border-color: var(--adp-line); }

/* ---------------- Fallback charts (used when pyarrow is unavailable) ----- */
.adp-bars { padding: 6px 2px; }
.adp-bar-row { display: flex; align-items: center; gap: 12px; margin-bottom: 11px; }
.adp-bar-label {
    width: 118px; flex: none; font-size: 12px; font-weight: 700;
    color: var(--adp-navy); text-align: right;
}
.adp-bar-track {
    flex: 1; background: #EEF2F6; border-radius: 6px; height: 22px; overflow: hidden;
}
.adp-bar-fill {
    height: 100%; border-radius: 6px;
    background: linear-gradient(90deg, #334155 0%, #0F172A 100%);
}
.adp-bar-fill.neg { background: linear-gradient(90deg, #F87171 0%, #B91C1C 100%); }
.adp-bar-value {
    width: 132px; flex: none; font-size: 12px; font-weight: 700;
    color: var(--adp-ink); text-align: right; font-variant-numeric: tabular-nums;
}
.adp-spark { width: 100%; height: auto; display: block; }
.adp-fallback-note {
    font-size: 11px; color: var(--adp-muted); margin-top: 8px; font-style: italic;
}

/* ---------------- High-contrast light-mode guard ------------------------ */
/* .streamlit/config.toml pins base="light"; these rules keep every label in  */
/* the main area dark and legible even if the app is launched without it on   */
/* a dark-mode operating system.                                              */
[data-testid="stMain"] { color: var(--adp-ink); }
[data-testid="stMain"] p,
[data-testid="stMain"] li,
[data-testid="stMain"] label,
[data-testid="stMain"] label p,
[data-testid="stMain"] summary,
[data-testid="stMain"] small,
[data-testid="stMain"] [data-testid="stWidgetLabel"],
[data-testid="stMain"] [data-testid="stWidgetLabel"] p,
[data-testid="stMain"] [data-testid="stMarkdownContainer"] p,
[data-testid="stMain"] [data-testid="stFileUploaderDropzoneInstructions"],
[data-testid="stMain"] [data-testid="stFileUploaderDropzoneInstructions"] span,
[data-testid="stMain"] [data-testid="stExpander"] summary,
[data-testid="stMain"] .stTabs [data-baseweb="tab"] {
    color: var(--adp-ink);
}
[data-testid="stMain"] [data-testid="stCaptionContainer"],
[data-testid="stMain"] [data-testid="stCaptionContainer"] p {
    color: var(--adp-muted);
}
[data-testid="stMain"] h1,
[data-testid="stMain"] h2,
[data-testid="stMain"] h3,
[data-testid="stMain"] h4,
[data-testid="stMain"] strong,
[data-testid="stMain"] b { color: var(--adp-navy); }
[data-testid="stMain"] input,
[data-testid="stMain"] textarea,
[data-testid="stMain"] [data-baseweb="input"],
[data-testid="stMain"] [data-baseweb="base-input"],
[data-testid="stMain"] [data-baseweb="select"] > div,
[data-testid="stMain"] [data-testid="stExpander"] details {
    background-color: var(--adp-card) !important;
    color: var(--adp-ink) !important;
    border-color: var(--adp-line) !important;
}
[data-testid="stMain"] [data-testid="stVerticalBlockBorderWrapper"] {
    background-color: transparent;
}
</style>
"""

st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# --------------------------------------------------------------------------- #
#  SESSION STATE
# --------------------------------------------------------------------------- #

def init_state() -> None:
    """Create every session-state key the app relies on exactly once."""
    defaults = {
        "raw_df": None,          # DataFrame as read from disk
        "raw_meta": None,        # dict: name, size_bytes, rows, cols, source
        "clean_df": None,        # processed DataFrame
        "result": None,          # dict of processing metrics
        "history": [],           # list of previous run dicts
        "nav": "Dashboard",
        "vat_rate": VAT_RATE_DEFAULT,
        "opt_flag_missing_id": True,
        "opt_flag_missing_status": True,
        "opt_flag_bad_dates": True,
        "opt_flag_nonpositive": True,
        "opt_drop_duplicates": False,
        "opt_flag_duplicates": True,
        "opt_bank_recon": True,
        "opt_dayfirst": True,
        "opt_exclude_cancelled": True,
        "opt_min_amount": 0.0,
        "currency": CURRENCY_DEFAULT,
        "opt_outlier_sigma": 3.0,
        "opt_flag_nulls": True,
        "opt_flag_outliers": True,
        "boot_time": datetime.now(),
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


init_state()


# --------------------------------------------------------------------------- #
#  SMALL HELPERS
# --------------------------------------------------------------------------- #

def currency_symbol() -> str:
    """The symbol the operator picked under Rule Settings ('' means none)."""
    try:
        return CURRENCY_OPTIONS.get(
            st.session_state.get("currency", CURRENCY_DEFAULT), "R")
    except Exception:
        return "R"


def fmt_money(value: float, decimals: int = 0) -> str:
    """Format a number using the selected currency symbol, e.g. R 1,284,930."""
    symbol = currency_symbol()
    prefix = f"{symbol} " if symbol else ""
    try:
        value = float(value)
    except (TypeError, ValueError):
        return f"{prefix}0"
    return f"{prefix}{value:,.{decimals}f}"


# Backwards-compatible alias - the whole app now routes through fmt_money.
fmt_zar = fmt_money


def fmt_size(num_bytes: int) -> str:
    """Human-readable file size."""
    size = float(num_bytes or 0)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.1f}{unit}" if unit != "B" else f"{int(size)}B"
        size /= 1024
    return f"{size:.1f}GB"


def norm_header(name: object) -> str:
    """Normalise a raw column header for alias matching."""
    text = re.sub(r"[^a-z0-9]+", "_", str(name).strip().lower())
    return text.strip("_")


def esc(value: object) -> str:
    """Minimal HTML escaping for values rendered into the preview table."""
    return (
        str(value)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
    )


def is_blank_value(value: object) -> bool:
    """True for None, NaN/NaT/NA, or a string that is empty once trimmed."""
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip() in ("", "<NA>", "nan", "NaN", "NaT", "None")


def clean_text_series(series: pd.Series) -> pd.Series:
    """RULE R01 - trim whitespace, collapse internal runs, drop control chars."""
    out = series.astype("string")
    out = out.str.replace(r"[\x00-\x1f\x7f]", "", regex=True)
    out = out.str.replace(r"\s+", " ", regex=True).str.strip()
    out = out.replace({"": pd.NA, "nan": pd.NA, "NaN": pd.NA, "None": pd.NA,
                       "null": pd.NA, "NULL": pd.NA, "-": pd.NA, "N/A": pd.NA,
                       "n/a": pd.NA, "#N/A": pd.NA})
    return out


ISO_DATE_RE = r"^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}"


def parse_dates(series: pd.Series, dayfirst: bool = True) -> pd.Series:
    """
    RULE R03 - parse a column of mixed-format dates.

    Year-first values (2025/09/12) are parsed year-month-day regardless of the
    day-first setting; everything else (12/09/2025) honours it. Each element is
    parsed on its own, so one odd value cannot null out the whole column.
    """
    values = series.astype("string")
    parsed = pd.Series(pd.NaT, index=values.index, dtype="datetime64[ns]")

    iso_mask = values.str.match(ISO_DATE_RE).fillna(False).astype(bool)
    if iso_mask.any():
        iso_values = values[iso_mask].str.replace(r"[/.]", "-", regex=True)
        parsed.loc[iso_mask] = pd.to_datetime(
            iso_values, errors="coerce", format="mixed", yearfirst=True, dayfirst=False
        )
    other_mask = ~iso_mask
    if other_mask.any():
        parsed.loc[other_mask] = pd.to_datetime(
            values[other_mask], errors="coerce", format="mixed", dayfirst=dayfirst
        )
    return parsed


def parse_amount(value: object) -> float:
    """RULE R09 - parse a ZAR currency string into a float."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return float("nan")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = str(value).strip()
    if not text:
        return float("nan")
    negative = text.startswith("(") and text.endswith(")")
    text = text.strip("()")
    text = re.sub(r"(?i)\b(zar|rand)\b", "", text)
    text = text.replace("R", "").replace("r", "")
    text = text.replace("\u00a0", "").replace(" ", "")
    # Treat a comma as a decimal separator only when it clearly is one.
    if "," in text and "." not in text and re.search(r",\d{1,2}$", text):
        text = text.replace(",", ".")
    else:
        text = text.replace(",", "")
    text = re.sub(r"[^0-9.\-]", "", text)
    if text in ("", "-", ".", "-."):
        return float("nan")
    try:
        number = float(text)
    except ValueError:
        return float("nan")
    return -number if negative else number


# --------------------------------------------------------------------------- #
#  DATA INGESTION
# --------------------------------------------------------------------------- #

def read_uploaded_file(uploaded) -> pd.DataFrame:
    """Read an uploaded .xlsx / .xls / .csv into a DataFrame."""
    name = uploaded.name.lower()
    data = uploaded.getvalue()
    if name.endswith(".csv"):
        for encoding in ("utf-8-sig", "utf-8", "latin-1"):
            for sep in (None, ",", ";", "\t", "|"):
                try:
                    frame = pd.read_csv(
                        io.BytesIO(data), encoding=encoding, sep=sep,
                        engine="python", dtype=object, skip_blank_lines=True,
                    )
                    if frame.shape[1] > 1 or sep is not None:
                        return frame
                except Exception:
                    continue
        raise ValueError("Unable to parse the CSV file with any known delimiter.")
    engine = "xlrd" if name.endswith(".xls") else "openpyxl"
    try:
        return pd.read_excel(io.BytesIO(data), dtype=object, engine=engine)
    except Exception:
        # .xls files are frequently mislabelled .xlsx exports (and vice versa).
        return pd.read_excel(io.BytesIO(data), dtype=object)


def map_columns(frame: pd.DataFrame) -> tuple[pd.DataFrame, dict, list]:
    """Map arbitrary raw headers onto the canonical schema."""
    normalised = {norm_header(col): col for col in frame.columns}
    mapping: dict[str, str] = {}
    used: set[str] = set()

    for canonical, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in normalised and normalised[alias] not in used:
                mapping[canonical] = normalised[alias]
                used.add(normalised[alias])
                break
    # Second pass: fuzzy contains-match for anything still unmapped.
    for canonical, aliases in COLUMN_ALIASES.items():
        if canonical in mapping:
            continue
        for norm_name, original in normalised.items():
            if original in used:
                continue
            if any(alias in norm_name for alias in aliases):
                mapping[canonical] = original
                used.add(original)
                break

    out = pd.DataFrame(index=frame.index)
    for canonical in ("Date", "Client_ID", "Reference", "Status", "Amount", "Notes"):
        source = mapping.get(canonical)
        out[canonical] = frame[source] if source else pd.NA

    extras = [col for col in frame.columns if col not in used]
    for col in extras:
        label = f"EXTRA_{col}"
        out[label] = frame[col]
    return out, mapping, extras


def build_sample_dataset(rows: int = 220) -> pd.DataFrame:
    """Deterministic, deliberately messy sample export for demo runs."""
    import random

    rng = random.Random(20240915)
    statuses = ["Paid", "paid ", "PENDING", "Pending", "overdue", "Cancelled",
                "settled", "unpaid", "", "Late"]
    branches = ["JHB", "CPT", "DBN", "PTA", "PLZ"]
    records = []
    for i in range(rows):
        day = rng.randint(1, 30)
        amount = round(rng.uniform(350, 48000), 2)
        client = f"MM-{branches[i % len(branches)]}-{1000 + rng.randint(1, 480)}"
        status = statuses[rng.randrange(len(statuses))]
        reference = f"EFT{rng.randint(100000, 999999)}"

        # Inject realistic dirt.
        if i % 17 == 0:
            client = ""                                   # missing ID
        if i % 23 == 0:
            reference = ""                                # unreconciled
        if i % 29 == 0:
            amount = 0                                    # zero amount
        if i % 31 == 0:
            amount = -abs(amount)                         # credit / reversal
        date_value = f"2025/09/{day:02d}"
        if i % 13 == 0:
            date_value = f"{day:02d}-09-2025"
        if i % 41 == 0:
            date_value = "not a date"
        amount_value = amount if i % 7 else f"R {amount:,.2f}"

        records.append({
            "  Date ": date_value,
            "Client ID": f"  {client}  ",
            "Bank Ref": reference,
            "Status": status,
            "Amount (ZAR)": amount_value,
            "Remarks": "Monthly retail sale" if i % 5 else "",
            "Branch": branches[i % len(branches)],
        })
        if i % 37 == 0:                                   # duplicate line
            records.append(records[-1].copy())
        if i % 53 == 0:                                   # fully blank row
            records.append({k: "" for k in records[-1]})
    return pd.DataFrame(records)


def mark_needs_review(frame: pd.DataFrame) -> pd.DataFrame:
    """
    Add a Needs_Review flag for rows whose status or notes ask for attention.

    Looks across the Status and Notes columns (and any other text column in a
    file with an arbitrary schema) for the review keywords.
    """
    search_cols = [c for c in ("Status", "Notes") if c in frame.columns]
    if not search_cols:
        search_cols = [c for c in frame.columns
                       if c not in ("Flagged", "Flag_Count", "Needs_Review")
                       and frame[c].dtype == object]
    if not search_cols:
        frame["Needs_Review"] = False
        return frame

    haystack = frame[search_cols].astype(str).agg(" ".join, axis=1).str.lower()
    hits = pd.Series(False, index=frame.index)
    for word in REVIEW_KEYWORDS:
        hits = hits | haystack.str.contains(word, regex=False, na=False)
    frame["Needs_Review"] = hits.to_numpy()
    return frame


# --------------------------------------------------------------------------- #
#  UNIVERSAL COLUMN PROFILER
#  Categorises any spreadsheet's columns without assuming a fixed schema.
# --------------------------------------------------------------------------- #

TYPE_NUMERIC = "Numeric"
TYPE_DATETIME = "Datetime"
TYPE_CATEGORICAL = "Categorical"
TYPE_TEXT = "Text / ID"
TYPE_EMPTY = "Empty"

PARSE_THRESHOLD = 0.80      # share of non-null values that must parse cleanly
CATEGORICAL_MAX_RATIO = 0.50    # repeats often enough to be a category
CATEGORICAL_SMALL_SET = 12      # or is a small fixed vocabulary in a small file


def _non_null(series: pd.Series) -> pd.Series:
    """Values that are neither null nor an empty/whitespace-only string."""
    return series[~series.map(is_blank_value)]


# A value only counts as numeric for profiling if, once a currency symbol and
# digit separators are removed, nothing alphabetic is left. Without this gate an
# identifier like "ORD-0001" or a note like "note 7" would be read as a number.
NUMERIC_GATE_RE = re.compile(
    r"^\s*[-+(]?\s*(?:ZAR|US\$|R|\$|\u20ac|\u00a3|\u00a5)?\s*"
    r"[0-9][0-9\s\u00a0.,]*\)?\s*%?\s*$",
    re.IGNORECASE,
)


def looks_numeric(value: object) -> bool:
    """True when a raw cell value is safe to interpret as a number."""
    if value is None or is_blank_value(value):
        return False
    if isinstance(value, bool):
        return False
    if isinstance(value, (int, float)):
        return True
    return bool(NUMERIC_GATE_RE.match(str(value)))


def detect_column_type(series: pd.Series) -> tuple[str, pd.Series]:
    """
    Classify one column as Numeric, Datetime, Categorical, Text/ID or Empty.

    Returns the type label plus the coerced series (numeric or datetime values
    where that is what the column turned out to hold, otherwise the original).
    """
    values = _non_null(series)
    if values.empty:
        return TYPE_EMPTY, series

    if pd.api.types.is_bool_dtype(series):
        return TYPE_CATEGORICAL, series
    if pd.api.types.is_numeric_dtype(series):
        return TYPE_NUMERIC, pd.to_numeric(series, errors="coerce")
    if pd.api.types.is_datetime64_any_dtype(series):
        return TYPE_DATETIME, pd.to_datetime(series, errors="coerce")

    # Currency- and thousands-separator aware numeric probe, gated so that text
    # containing digits (identifiers, free-text notes) is not misread as numeric.
    gate = values.map(looks_numeric)
    if gate.mean() >= PARSE_THRESHOLD:
        probe = values[gate].map(parse_amount)
        if len(probe) and probe.notna().mean() >= PARSE_THRESHOLD:
            return TYPE_NUMERIC, series.map(
                lambda v: parse_amount(v) if looks_numeric(v) else float("nan"))

    # Datetime probe - only for values that actually look like dates, so that
    # short numeric-ish codes are not swept up by a permissive parser.
    text = values.astype(str)
    date_ish = text.str.contains(r"[-/.:]", regex=True) | text.str.contains(
        r"(?i)\b(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)", regex=True)
    if date_ish.mean() >= PARSE_THRESHOLD:
        parsed = parse_dates(values)
        if parsed.notna().mean() >= PARSE_THRESHOLD:
            return TYPE_DATETIME, parse_dates(series)

    # A column is categorical when its values repeat (low unique-to-row ratio)
    # or when it holds a small fixed vocabulary. Near-unique strings such as
    # identifiers and free-text notes fall through to Text / ID.
    unique = int(values.nunique())
    ratio = unique / max(len(values), 1)
    if ratio <= CATEGORICAL_MAX_RATIO or unique <= CATEGORICAL_SMALL_SET:
        return TYPE_CATEGORICAL, series
    return TYPE_TEXT, series


def profile_dataframe(frame: pd.DataFrame) -> dict:
    """
    Build a per-column profile of any DataFrame.

    Returns {column: {type, coerced, nulls, null_pct, unique, stats...}} plus a
    "__order__" key holding the original column order.
    """
    profile: dict = {"__order__": list(frame.columns)}
    total = max(len(frame), 1)

    for column in frame.columns:
        series = frame[column]
        col_type, coerced = detect_column_type(series)
        blanks = series.map(is_blank_value)
        nulls = int(blanks.sum())
        values = _non_null(series)

        entry = {
            "type": col_type,
            "coerced": coerced,
            "count": int(len(series)),
            "non_null": int(len(values)),
            "nulls": nulls,
            "null_pct": nulls / total * 100,
            "unique": int(values.nunique()) if len(values) else 0,
        }

        if col_type == TYPE_NUMERIC:
            numeric = pd.to_numeric(coerced, errors="coerce")
            clean = numeric.dropna()
            entry.update({
                "sum": float(clean.sum()) if len(clean) else 0.0,
                "mean": float(clean.mean()) if len(clean) else 0.0,
                "median": float(clean.median()) if len(clean) else 0.0,
                "min": float(clean.min()) if len(clean) else 0.0,
                "max": float(clean.max()) if len(clean) else 0.0,
                "std": float(clean.std()) if len(clean) > 1 else 0.0,
                "unparsed": int(numeric.isna().sum() - nulls),
            })
        elif col_type == TYPE_DATETIME:
            parsed = pd.to_datetime(coerced, errors="coerce")
            good = parsed.dropna()
            entry.update({
                "min_date": good.min().strftime("%Y-%m-%d") if len(good) else "",
                "max_date": good.max().strftime("%Y-%m-%d") if len(good) else "",
                "unparsed": int(parsed.isna().sum() - nulls),
            })
        elif col_type in (TYPE_CATEGORICAL, TYPE_TEXT):
            counts = values.astype(str).value_counts()
            entry["top_values"] = [(str(k), int(v)) for k, v in counts.head(8).items()]
            entry["mode"] = str(counts.index[0]) if len(counts) else ""
            entry["mode_share"] = (float(counts.iloc[0]) / max(len(values), 1) * 100
                                   if len(counts) else 0.0)
        profile[column] = entry
    return profile


def profile_table(profile: dict) -> pd.DataFrame:
    """Flatten a profile into the Column Data Types table shown in the UI."""
    rows = []
    for column in profile["__order__"]:
        entry = profile[column]
        detail = ""
        if entry["type"] == TYPE_NUMERIC:
            detail = (f"min {entry['min']:,.2f} | mean {entry['mean']:,.2f} | "
                      f"max {entry['max']:,.2f}")
        elif entry["type"] == TYPE_DATETIME:
            detail = f"{entry.get('min_date','')} to {entry.get('max_date','')}"
        elif entry.get("top_values"):
            detail = ", ".join(f"{k} ({v})" for k, v in entry["top_values"][:3])
        rows.append({
            "Column": str(column),
            "Detected Type": entry["type"],
            "Non-Null": entry["non_null"],
            "Missing": entry["nulls"],
            "Missing %": round(entry["null_pct"], 1),
            "Unique": entry["unique"],
            "Detail": detail,
        })
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------- #
#  DATA QUALITY & ANOMALY ENGINE
# --------------------------------------------------------------------------- #

SEV_HIGH = "High"
SEV_MEDIUM = "Medium"
SEV_LOW = "Low"


def audit_dataframe(frame: pd.DataFrame, profile: dict, options: dict) -> tuple:
    """
    Inspect any DataFrame for missing values, duplicates, unparseable dates and
    severe numeric outliers.

    Returns (audit_table, row_issues, totals) where row_issues is a list of
    per-row issue labels used to compute the dataset health score.
    """
    total = max(len(frame), 1)
    issues: list[dict] = []
    row_issues: list[list[str]] = [[] for _ in range(len(frame))]
    sigma = float(options.get("outlier_sigma", 3.0) or 3.0)

    def mark(mask, label: str) -> int:
        hits = 0
        for idx, hit in enumerate(pd.Series(mask).fillna(False).to_numpy()):
            if hit and label not in row_issues[idx]:
                row_issues[idx].append(label)
                hits += 1
        return hits

    # ---- Per-column checks -------------------------------------------------
    for column in profile["__order__"]:
        entry = profile[column]
        series = frame[column]
        col_type = entry["type"]

        if entry["nulls"]:
            # A structurally empty column is reported once below; marking every
            # row for it would drive the health score to zero on one bad column.
            if options.get("flag_nulls", True) and col_type != TYPE_EMPTY:
                mark(series.map(is_blank_value), f"NULL: {column}")
            issues.append({
                "Column": str(column), "Type": col_type, "Issue": "Missing values",
                "Rows": entry["nulls"], "% of File": round(entry["null_pct"], 1),
                "Severity": SEV_HIGH if entry["null_pct"] > 20 else
                            (SEV_MEDIUM if entry["null_pct"] > 5 else SEV_LOW),
            })

        if entry["type"] == TYPE_EMPTY:
            issues.append({
                "Column": str(column), "Type": col_type, "Issue": "Column is entirely empty",
                "Rows": len(frame), "% of File": 100.0, "Severity": SEV_HIGH,
            })
            continue

        if entry["unique"] == 1 and entry["non_null"] > 1:
            issues.append({
                "Column": str(column), "Type": col_type,
                "Issue": f"Constant value ('{entry.get('mode', '')}')",
                "Rows": entry["non_null"],
                "% of File": round(entry["non_null"] / total * 100, 1),
                "Severity": SEV_LOW,
            })

        # Values that could not be coerced to the detected type.
        unparsed = int(entry.get("unparsed", 0) or 0)
        if unparsed > 0:
            label = ("Inconsistent date strings" if col_type == TYPE_DATETIME
                     else "Non-numeric values in a numeric column")
            coerced = entry["coerced"]
            bad = coerced.isna() & ~series.map(is_blank_value)
            mark(bad, f"BAD {'DATE' if col_type == TYPE_DATETIME else 'NUMBER'}: {column}")
            issues.append({
                "Column": str(column), "Type": col_type, "Issue": label,
                "Rows": unparsed, "% of File": round(unparsed / total * 100, 1),
                "Severity": SEV_HIGH,
            })

        # Severe numeric outliers, reported by both IQR and standard deviation.
        if col_type == TYPE_NUMERIC and options.get("flag_outliers", True):
            numeric = pd.to_numeric(entry["coerced"], errors="coerce")
            clean = numeric.dropna()
            if len(clean) >= 8:
                q1, q3 = clean.quantile(0.25), clean.quantile(0.75)
                iqr = q3 - q1
                std = clean.std()
                by_iqr = pd.Series(False, index=numeric.index)
                if iqr > 0:
                    by_iqr = (numeric < q1 - 3 * iqr) | (numeric > q3 + 3 * iqr)
                by_sigma = pd.Series(False, index=numeric.index)
                if std and std > 0:
                    by_sigma = (numeric - clean.mean()).abs() > sigma * std
                outliers = (by_iqr | by_sigma).fillna(False)
                count = int(outliers.sum())
                if count:
                    mark(outliers, f"OUTLIER: {column}")
                    issues.append({
                        "Column": str(column), "Type": col_type,
                        "Issue": f"Severe numeric outliers (>{sigma:g} SD or 3x IQR)",
                        "Rows": count, "% of File": round(count / total * 100, 1),
                        "Severity": SEV_MEDIUM,
                    })

        if col_type in (TYPE_CATEGORICAL, TYPE_TEXT):
            text = _non_null(series).astype(str)
            if len(text):
                folded = text.str.strip().str.lower()
                collapsed = int(text.nunique() - folded.nunique())
                if collapsed > 0:
                    issues.append({
                        "Column": str(column), "Type": col_type,
                        "Issue": "Inconsistent casing / spacing in labels",
                        "Rows": collapsed, "% of File": round(collapsed / total * 100, 1),
                        "Severity": SEV_MEDIUM,
                    })

    # ---- Dataset-level duplicate rows --------------------------------------
    dupes = frame.astype(str).duplicated(keep="first")
    dup_count = int(dupes.sum())
    if dup_count:
        mark(dupes, "DUPLICATE ROW")
        issues.append({
            "Column": "(whole row)", "Type": "-", "Issue": "Duplicate rows",
            "Rows": dup_count, "% of File": round(dup_count / total * 100, 1),
            "Severity": SEV_HIGH,
        })

    audit = pd.DataFrame(issues)
    if not audit.empty:
        rank = {SEV_HIGH: 0, SEV_MEDIUM: 1, SEV_LOW: 2}
        audit["__r"] = audit["Severity"].map(rank).fillna(3)
        audit = (audit.sort_values(["__r", "Rows"], ascending=[True, False])
                      .drop(columns="__r").reset_index(drop=True))

    affected = sum(1 for r in row_issues if r)
    totals = {
        "rows": len(frame),
        "columns": frame.shape[1],
        "issue_count": len(issues),
        "rows_affected": affected,
        "rows_clean": len(frame) - affected,
        "health": (len(frame) - affected) / total * 100,
        "duplicate_rows": dup_count,
        "total_cells": int(frame.size),
        "missing_cells": int(sum(profile[c]["nulls"] for c in profile["__order__"])),
    }
    return audit, row_issues, totals


# --------------------------------------------------------------------------- #
#  BUSINESS RULES ENGINE
# --------------------------------------------------------------------------- #

def run_automation_engine(raw: pd.DataFrame, options: dict) -> tuple[pd.DataFrame, dict]:
    """
    Apply the full 12-rule South African automation pipeline.

    Returns the cleaned DataFrame and a metrics dictionary used by the UI,
    the Excel writer and the PDF summary.
    """
    started = time.perf_counter()
    log: list[tuple[str, str]] = []
    applied: list[str] = []

    def note(level: str, message: str) -> None:
        log.append((level, message))

    rows_in = len(raw)
    note("dim", f"{ENGINE_LABEL} | source rows: {rows_in:,}")

    frame, mapping, extras = map_columns(raw)
    note("ok", f"Schema mapped: {', '.join(f'{k} <- {v}' for k, v in mapping.items()) or 'no direct matches'}")
    if extras:
        note("dim", f"Carried through {len(extras)} unmapped column(s) as EXTRA_*")

    # ---- R01  Whitespace & control-character trim -------------------------- #
    text_cols = [c for c in frame.columns if c != "Amount"]
    for col in text_cols:
        frame[col] = clean_text_series(frame[col])
    applied.append("R01")
    note("ok", "R01 Whitespace and control characters trimmed across text columns")

    # ---- R02  Blank row removal -------------------------------------------- #
    if frame.empty:
        raise ValueError("The uploaded file contains no data rows.")

    before = len(frame)
    frame = frame.dropna(how="all")
    if not frame.empty:
        blank_rows = frame.apply(lambda col: col.map(is_blank_value)).all(axis=1)
        frame = frame[~blank_rows]
    blank_dropped = before - len(frame)
    applied.append("R02")
    note("ok" if blank_dropped == 0 else "warn",
         f"R02 Blank rows removed: {blank_dropped:,}")

    if frame.empty:
        raise ValueError("Every row in the uploaded file was blank after sanitisation.")

    frame = frame.reset_index(drop=True)
    frame["Notes"] = frame["Notes"].fillna("").astype(str).str.strip()
    flags: list[list[str]] = [[] for _ in range(len(frame))]

    def add_flag(mask: pd.Series, label: str) -> int:
        """Attach a flag label to every row where mask is True."""
        mask = mask.fillna(False).to_numpy()
        hits = 0
        for idx, hit in enumerate(mask):
            if hit and label not in flags[idx]:
                flags[idx].append(label)
                hits += 1
        return hits

    # ---- R03  Date standardisation ----------------------------------------- #
    parsed_dates = parse_dates(frame["Date"], dayfirst=bool(options.get("dayfirst", True)))
    frame["Date"] = parsed_dates.dt.strftime("%Y-%m-%d")
    applied.append("R03")
    note("ok", "R03 Dates standardised to YYYY-MM-DD")

    # ---- R04  Invalid date flag -------------------------------------------- #
    bad_dates = 0
    if options.get("flag_bad_dates", True):
        bad_dates = add_flag(parsed_dates.isna(), FLAG_BAD_DATE)
        applied.append("R04")
        note("warn" if bad_dates else "ok",
             f"R04 Invalid or unparseable dates flagged: {bad_dates:,}")

    # ---- R05  Client ID normalisation --------------------------------------- #
    frame["Client_ID"] = (
        frame["Client_ID"].astype("string")
        .str.upper()
        .str.replace(r"\s*[-_]\s*", "-", regex=True)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    frame["Reference"] = frame["Reference"].astype("string").str.upper().str.strip()
    applied.append("R05")
    note("ok", "R05 Client IDs normalised (uppercase, delimiter-consistent)")

    # ---- R06  Missing Client ID flag ---------------------------------------- #
    missing_id = 0
    if options.get("flag_missing_id", True):
        missing_id = add_flag(frame["Client_ID"].isna() | (frame["Client_ID"].fillna("") == ""),
                              FLAG_MISSING_ID)
        applied.append("R06")
        note("err" if missing_id else "ok",
             f"R06 Rows missing a Client ID: {missing_id:,}")

    # ---- R07  Status vocabulary normalisation -------------------------------- #
    raw_status = frame["Status"].astype("string").fillna("")
    status_key = raw_status.str.lower().str.strip().str.replace(r"[\s\-]+", "_", regex=True)
    blank_status = raw_status.str.strip() == ""

    # Recognised payment vocabulary is folded onto the canonical set; every other
    # non-empty value (business segment, channel, industry, workflow state) is
    # preserved exactly as written. Only genuinely empty cells become UNKNOWN.
    canonical = status_key.map(STATUS_MAP)
    frame["Status"] = canonical.fillna(raw_status.str.strip()).mask(blank_status,
                                                                   STATUS_BLANK)

    mapped_count = int(canonical.notna().sum())
    preserved = sorted(set(frame["Status"][~blank_status & canonical.isna()].tolist()))
    applied.append("R07")
    note("ok", f"R07 Status normalised | {mapped_count:,} mapped to the payment "
               f"vocabulary, {len(preserved):,} custom value(s) preserved"
               + (f": {', '.join(preserved[:6])}" if preserved else ""))

    # ---- R08  Missing / unknown status flag ---------------------------------- #
    missing_status = 0
    if options.get("flag_missing_status", True):
        # Only a blank / null / empty cell is an exception. An unrecognised but
        # populated value is legitimate business data, not a validation error.
        missing_status = add_flag(blank_status, FLAG_MISSING_STATUS)
        applied.append("R08")
        note("err" if missing_status else "ok",
             f"R08 Rows with a blank status: {missing_status:,}")

    # ---- R09  ZAR currency parsing -------------------------------------------- #
    frame["Amount"] = frame["Amount"].map(parse_amount).astype(float)
    bad_amounts = add_flag(frame["Amount"].isna(), FLAG_BAD_AMOUNT)
    frame["Amount"] = frame["Amount"].fillna(0.0).round(2)
    applied.append("R09")
    note("warn" if bad_amounts else "ok",
         f"R09 ZAR amounts parsed | unparseable values coerced to 0.00: {bad_amounts:,}")

    # ---- R10  Zero / negative amount flag -------------------------------------- #
    negatives = zeros = below = 0
    if options.get("flag_nonpositive", True):
        negatives = add_flag(frame["Amount"] < 0, FLAG_NEGATIVE)
        zeros = add_flag(frame["Amount"] == 0, FLAG_ZERO)
        threshold = float(options.get("min_amount", 0.0) or 0.0)
        if threshold > 0:
            below = add_flag((frame["Amount"] > 0) & (frame["Amount"] < threshold),
                             FLAG_BELOW_THRESHOLD)
        applied.append("R10")
        note("warn" if (negatives or zeros) else "ok",
             f"R10 Amount anomalies | negative: {negatives:,}  zero: {zeros:,}  below threshold: {below:,}")

    # ---- R11  VAT @ 15% and Total Inc VAT --------------------------------------- #
    vat_rate = float(options.get("vat_rate", VAT_RATE_DEFAULT))
    frame["VAT_15"] = (frame["Amount"] * vat_rate).round(2)
    frame["Total_Inc_VAT"] = (frame["Amount"] + frame["VAT_15"]).round(2)
    applied.append("R11")
    note("ok", f"R11 VAT calculated at {vat_rate * 100:.0f}% | Total_Inc_VAT derived")

    # ---- R12  Bank reconciliation & duplicate check ------------------------------ #
    unreconciled = mismatches = duplicates = 0
    dup_dropped = 0
    reconciled = 0
    has_reference_column = "Reference" in mapping
    frame["Recon_Status"] = RECON_SKIPPED

    if options.get("bank_recon", True):
        if has_reference_column:
            no_ref = frame["Reference"].isna() | (frame["Reference"].fillna("") == "")
            # Rows carrying a bank reference are reconciled; only rows that are
            # blank in an existing reference column are exceptions.
            frame["Recon_Status"] = pd.Series(
                [RECON_NONE if missing else RECON_OK for missing in no_ref],
                index=frame.index,
            )
            unreconciled = add_flag(no_ref, FLAG_UNRECONCILED)
            reconciled = len(frame) - unreconciled
            note("warn" if unreconciled else "ok",
                 f"R12a Bank reconciliation | {reconciled:,} reconciled, "
                 f"{unreconciled:,} unassigned")
        else:
            # No reference column in the source at all - flagging every row would
            # be noise, so the per-row check is skipped and reported once.
            note("dim", "R12a Bank reconciliation skipped - the source file has "
                        "no reference column")

        # A settled invoice with no money attached cannot reconcile to the bank.
        mismatches = add_flag((frame["Status"] == "PAID") & (frame["Amount"] <= 0),
                              FLAG_RECON_MISMATCH)
        if mismatches:
            note("warn", f"R12a Reconciliation discrepancies: {mismatches:,}")

    dup_key = frame[["Date", "Client_ID", "Amount", "Reference"]].astype(str)
    dup_mask = dup_key.duplicated(keep="first")
    if options.get("drop_duplicates", False):
        dup_dropped = int(dup_mask.sum())
        keep = ~dup_mask.to_numpy()
        frame = frame[keep].reset_index(drop=True)
        flags = [f for f, k in zip(flags, keep) if k]
        note("warn" if dup_dropped else "ok",
             f"R12b Duplicate entries removed: {dup_dropped:,}")
    elif options.get("flag_duplicates", True):
        duplicates = add_flag(dup_mask, FLAG_DUPLICATE)
        note("warn" if duplicates else "ok",
             f"R12b Duplicate entries flagged (retained): {duplicates:,}")
    applied.append("R12")

    # ---- Compose the Notes column ------------------------------------------------ #
    existing_notes = frame["Notes"].fillna("").astype(str).tolist()
    composed = []
    for original, row_flags in zip(existing_notes, flags):
        parts = [f for f in row_flags]
        original = original.strip()
        if original and original.lower() not in ("nan", "none", "<na>"):
            parts.append(original)
        composed.append(" | ".join(parts))
    frame["Notes"] = composed
    frame["Flagged"] = [bool(f) for f in flags]
    frame["Flag_Count"] = [len(f) for f in flags]
    frame = mark_needs_review(frame)

    # ---- Final ordering ----------------------------------------------------------- #
    helper_cols = ["Flagged", "Flag_Count", "Needs_Review"]
    extra_cols = [c for c in frame.columns
                  if c not in OUTPUT_COLUMNS + helper_cols]
    frame = frame[OUTPUT_COLUMNS + extra_cols + helper_cols]
    frame["Date"] = frame["Date"].fillna("")
    frame["Client_ID"] = frame["Client_ID"].fillna("")
    frame["Reference"] = frame["Reference"].fillna("")

    # ---- Aggregates ----------------------------------------------------------------- #
    if options.get("exclude_cancelled", True):
        revenue_mask = frame["Status"] != "CANCELLED"
    else:
        revenue_mask = pd.Series(True, index=frame.index)
    revenue = float(frame.loc[revenue_mask, "Amount"].sum())
    vat_total = float(frame.loc[revenue_mask, "VAT_15"].sum())
    inc_vat_total = float(frame.loc[revenue_mask, "Total_Inc_VAT"].sum())
    flagged_count = int(frame["Flagged"].sum())

    duration = time.perf_counter() - started
    note("hl", f"Completed in {duration:.2f} seconds | {len(applied)} business "
                f"validation rules applied")

    flag_breakdown: dict[str, int] = {}
    for row_flags in flags:
        for label in row_flags:
            flag_breakdown[label] = flag_breakdown.get(label, 0) + 1

    display = frame.drop(columns=HELPER_COLUMNS, errors="ignore")
    profile = profile_dataframe(display)
    audit, _row_issues, totals = audit_dataframe(display, profile, options)
    health = (len(frame) - flagged_count) / max(len(frame), 1) * 100

    result = {
        "mode": "sales",
        "profile": profile,
        "audit": audit,
        "health": health,
        "columns": display.shape[1],
        "total_cells": totals["total_cells"],
        "missing_cells": totals["missing_cells"],
        "duplicate_rows": totals["duplicate_rows"],
        "numeric_stats": numeric_summary_table(profile),
        "categorical_stats": categorical_summary_table(profile),
        "datetime_stats": datetime_summary_table(profile),
        "key_numeric": key_numeric_columns(profile),
        "numeric_total": revenue,
        "preview_columns": [c for c in PREVIEW_COLUMNS if c in frame.columns
                            and (c != "Recon_Status"
                                 or options.get("bank_recon", True))],
        "rows_in": rows_in,
        "rows_out": len(frame),
        "blank_dropped": blank_dropped,
        "dup_dropped": dup_dropped,
        "flagged": flagged_count,
        "clean": len(frame) - flagged_count,
        "revenue": revenue,
        "vat_total": vat_total,
        "inc_vat_total": inc_vat_total,
        "vat_rate": vat_rate,
        "duration": duration,
        "rules_applied": applied,
        "rule_count": len(applied),
        "log": log,
        "flag_breakdown": dict(sorted(flag_breakdown.items(), key=lambda kv: -kv[1])),
        "mapping": mapping,
        "extras": extras,
        "timestamp": datetime.now(),
        "exclude_cancelled": bool(options.get("exclude_cancelled", True)),
    }
    return frame, result



# --------------------------------------------------------------------------- #
#  UNIVERSAL SUMMARY TABLES
# --------------------------------------------------------------------------- #

def numeric_summary_table(profile: dict) -> pd.DataFrame:
    """Descriptive statistics for every numeric column."""
    rows = []
    for column in profile["__order__"]:
        entry = profile[column]
        if entry["type"] != TYPE_NUMERIC:
            continue
        rows.append({
            "Column": str(column),
            "Count": entry["non_null"],
            "Sum": round(entry["sum"], 2),
            "Mean": round(entry["mean"], 2),
            "Median": round(entry["median"], 2),
            "Min": round(entry["min"], 2),
            "Max": round(entry["max"], 2),
            "Std Dev": round(entry["std"], 2),
            "Missing": entry["nulls"],
        })
    return pd.DataFrame(rows)


def categorical_summary_table(profile: dict, top_n: int = 5) -> pd.DataFrame:
    """Top value distributions for every categorical / text column."""
    rows = []
    for column in profile["__order__"]:
        entry = profile[column]
        if entry["type"] not in (TYPE_CATEGORICAL, TYPE_TEXT):
            continue
        top = entry.get("top_values", [])[:top_n]
        rows.append({
            "Column": str(column),
            "Type": entry["type"],
            "Unique": entry["unique"],
            "Most Common": entry.get("mode", ""),
            "Frequency": top[0][1] if top else 0,
            "Share %": round(entry.get("mode_share", 0.0), 1),
            "Top Values": ", ".join(f"{k} ({v})" for k, v in top),
            "Missing": entry["nulls"],
        })
    return pd.DataFrame(rows)


def datetime_summary_table(profile: dict) -> pd.DataFrame:
    """Range coverage for every detected datetime column."""
    rows = []
    for column in profile["__order__"]:
        entry = profile[column]
        if entry["type"] != TYPE_DATETIME:
            continue
        rows.append({
            "Column": str(column),
            "Earliest": entry.get("min_date", ""),
            "Latest": entry.get("max_date", ""),
            "Distinct Dates": entry["unique"],
            "Unparseable": int(entry.get("unparsed", 0) or 0),
            "Missing": entry["nulls"],
        })
    return pd.DataFrame(rows)


def key_numeric_columns(profile: dict, limit: int = 3) -> list:
    """
    The numeric columns most worth totalling on a KPI card.

    Prefers columns whose name suggests a measure, then falls back to the
    largest magnitude - never assumes a column called 'Amount' exists.
    """
    measure_words = ("amount", "total", "value", "revenue", "sales", "price",
                     "cost", "net", "gross", "sum", "spend", "balance", "qty",
                     "quantity", "units", "count")
    scored = []
    for column in profile["__order__"]:
        entry = profile[column]
        if entry["type"] != TYPE_NUMERIC:
            continue
        name = norm_header(column)
        # An identifier-looking numeric column is a poor thing to total.
        is_id = name.endswith("_id") or name in ("id", "no", "number", "index")
        named = any(word in name for word in measure_words)
        scored.append((0 if is_id else (2 if named else 1),
                       abs(entry.get("sum", 0.0)), column))
    scored.sort(key=lambda t: (-t[0], -t[1]))
    return [c for _, _, c in scored[:limit]]


# --------------------------------------------------------------------------- #
#  UNIVERSAL PROCESSING ENGINE
#  Runs on any spreadsheet whose columns do not resolve to the sales ledger
#  schema. Nothing here assumes a fixed column name.
# --------------------------------------------------------------------------- #

SALES_CANONICAL = ("Date", "Client_ID", "Status", "Amount", "Reference")


def is_sales_schema(mapping: dict) -> bool:
    """
    True when the file resolves to the client-transaction ledger the South
    African rules engine is built for: a money column plus at least two other
    recognised ledger fields. Anything else is handled universally.
    """
    matched = [c for c in SALES_CANONICAL if c in mapping]
    return "Amount" in mapping and len(matched) >= 3


def run_universal_engine(raw: pd.DataFrame, options: dict) -> tuple:
    """
    Profile, sanitise and audit an arbitrary spreadsheet.

    Every original column is preserved. Detected datetime columns are
    standardised to YYYY-MM-DD and detected numeric columns are coerced to
    numbers; everything else is trimmed but otherwise left untouched.
    """
    started = time.perf_counter()
    log: list[tuple[str, str]] = []

    def note(level: str, message: str) -> None:
        log.append((level, message))

    rows_in = len(raw)
    note("dim", f"{ENGINE_LABEL} | universal mode | source rows: {rows_in:,}")

    frame = raw.copy()
    frame.columns = [str(c).strip() or f"Column_{i+1}"
                     for i, c in enumerate(frame.columns)]

    # ---- Sanitisation ------------------------------------------------------
    for column in frame.columns:
        if frame[column].dtype == object:
            frame[column] = clean_text_series(frame[column])
    note("ok", "Whitespace and control characters trimmed across text columns")

    if frame.empty:
        raise ValueError("The uploaded file contains no data rows.")
    before = len(frame)
    frame = frame.dropna(how="all")
    if not frame.empty:
        blank_rows = frame.apply(lambda col: col.map(is_blank_value)).all(axis=1)
        frame = frame[~blank_rows]
    blank_dropped = before - len(frame)
    frame = frame.reset_index(drop=True)
    note("ok" if not blank_dropped else "warn",
         f"Blank rows removed: {blank_dropped:,}")
    if frame.empty:
        raise ValueError("Every row in the uploaded file was blank after sanitisation.")

    # ---- Profile -----------------------------------------------------------
    profile = profile_dataframe(frame)
    counts: dict[str, int] = {}
    for column in profile["__order__"]:
        counts[profile[column]["type"]] = counts.get(profile[column]["type"], 0) + 1
    note("ok", "Column types detected | " + ", ".join(
        f"{v} {k}" for k, v in sorted(counts.items())))

    # ---- Standardise the columns whose type we are confident about ---------
    standardised = 0
    for column in profile["__order__"]:
        entry = profile[column]
        if entry["type"] == TYPE_DATETIME:
            frame[column] = pd.to_datetime(
                entry["coerced"], errors="coerce").dt.strftime("%Y-%m-%d")
            frame[column] = frame[column].fillna("")
            standardised += 1
        elif entry["type"] == TYPE_NUMERIC:
            frame[column] = pd.to_numeric(entry["coerced"], errors="coerce")
            standardised += 1
    if standardised:
        note("ok", f"Standardised {standardised} typed column(s) "
                   "(dates to YYYY-MM-DD, numerics coerced)")
        profile = profile_dataframe(frame)

    # ---- Quality audit -----------------------------------------------------
    audit, row_issues, totals = audit_dataframe(frame, profile, options)
    for _, issue in (audit.head(6).iterrows() if not audit.empty else []):
        note("err" if issue["Severity"] == SEV_HIGH else "warn",
             f"{issue['Column']}: {issue['Issue']} ({issue['Rows']:,} rows)")
    if audit.empty:
        note("ok", "Data quality audit clean - no issues detected")

    frame["Notes"] = [" | ".join(issues) for issues in row_issues]
    frame["Flagged"] = [bool(issues) for issues in row_issues]
    frame["Flag_Count"] = [len(issues) for issues in row_issues]
    frame = mark_needs_review(frame)

    numeric_stats = numeric_summary_table(profile)
    categorical_stats = categorical_summary_table(profile)
    key_columns = key_numeric_columns(profile)
    numeric_total = float(sum(profile[c]["sum"] for c in key_columns)) if key_columns else 0.0

    duration = time.perf_counter() - started
    note("hl", f"Completed in {duration:.2f} seconds | "
               f"{totals['columns']} columns profiled, "
               f"{totals['issue_count']} data quality issue(s) raised")

    breakdown: dict[str, int] = {}
    for issues in row_issues:
        for label in issues:
            breakdown[label] = breakdown.get(label, 0) + 1

    result = {
        "mode": "universal",
        "rows_in": rows_in,
        "rows_out": len(frame),
        "blank_dropped": blank_dropped,
        "dup_dropped": 0,
        "flagged": totals["rows_affected"],
        "clean": totals["rows_clean"],
        "health": totals["health"],
        "columns": totals["columns"],
        "total_cells": totals["total_cells"],
        "missing_cells": totals["missing_cells"],
        "duplicate_rows": totals["duplicate_rows"],
        "duration": duration,
        "rules_applied": [],
        "rule_count": totals["issue_count"],
        "log": log,
        "flag_breakdown": dict(sorted(breakdown.items(), key=lambda kv: -kv[1])),
        "mapping": {},
        "extras": [],
        "timestamp": datetime.now(),
        "profile": profile,
        "audit": audit,
        "numeric_stats": numeric_stats,
        "categorical_stats": categorical_stats,
        "datetime_stats": datetime_summary_table(profile),
        "key_numeric": key_columns,
        "numeric_total": numeric_total,
        "preview_columns": [c for c in profile["__order__"]
                            if c not in HELPER_COLUMNS] + ["Notes"],
        # Sales-only keys, present so the shared UI never has to branch on them.
        "revenue": numeric_total,
        "vat_total": 0.0,
        "inc_vat_total": 0.0,
        "vat_rate": float(options.get("vat_rate", VAT_RATE_DEFAULT)),
        "exclude_cancelled": False,
    }
    return frame, result


def status_summary(frame: pd.DataFrame) -> pd.DataFrame:
    """Pivot totals by Status - the basis of the Executive Summary tab."""
    pivot = frame.groupby("Status", dropna=False).agg(
        Transactions=("Amount", "size"),
        Revenue_Excl_VAT=("Amount", "sum"),
        VAT_15=("VAT_15", "sum"),
        Total_Inc_VAT=("Total_Inc_VAT", "sum"),
        Flagged_Rows=("Flagged", "sum"),
    ).reset_index()
    pivot["__order"] = pivot["Status"].map(
        {s: i for i, s in enumerate(STATUS_ORDER)}).fillna(99)
    # Canonical payment statuses first, then any preserved custom value A-Z.
    pivot = (pivot.sort_values(["__order", "Status"])
                  .drop(columns="__order").reset_index(drop=True))
    for col in ("Revenue_Excl_VAT", "VAT_15", "Total_Inc_VAT"):
        pivot[col] = pivot[col].round(2)
    pivot["Flagged_Rows"] = pivot["Flagged_Rows"].astype(int)
    return pivot


# --------------------------------------------------------------------------- #
#  MULTI-TAB FORMATTED EXCEL OUTPUT  (openpyxl)
# --------------------------------------------------------------------------- #

BODY_FONT_NAME = "Calibri"
BODY_FONT_SIZE = 10
HEADER_ROW_HEIGHT = 26
DATA_ROW_HEIGHT = 20

HEADER_FILL = PatternFill("solid", fgColor=NAVY)          # deep navy #0F172A
HEADER_FONT = Font(name=BODY_FONT_NAME, size=11, bold=True, color="FFFFFF")
ZEBRA_FILL = PatternFill("solid", fgColor="F8FAFC")       # alternating row wash
FLAG_FILL = PatternFill("solid", fgColor=RED_FILL_HEX)    # soft red #FEF2F2
FLAG_FONT = Font(name=BODY_FONT_NAME, size=BODY_FONT_SIZE, color="991B1B")
DATA_FONT = Font(name=BODY_FONT_NAME, size=BODY_FONT_SIZE, color="1E293B")
TITLE_FONT = Font(name=BODY_FONT_NAME, size=16, bold=True, color=NAVY)
SUB_FONT = Font(name=BODY_FONT_NAME, size=9, italic=True, color="64748B")
LABEL_FONT = Font(name=BODY_FONT_NAME, size=11, bold=True, color=NAVY)
TOTAL_FILL = PatternFill("solid", fgColor="E2E8F0")
TOTAL_FONT = Font(name=BODY_FONT_NAME, size=10, bold=True, color=NAVY)
GREEN_FONT = Font(name=BODY_FONT_NAME, size=10, bold=True, color="047857")

KPI_LABEL_FILL = PatternFill("solid", fgColor=NAVY)
KPI_LABEL_FONT = Font(name=BODY_FONT_NAME, size=9, bold=True, color="FFFFFF")
KPI_VALUE_FILL = PatternFill("solid", fgColor="F8FAFC")
KPI_VALUE_FONT = Font(name=BODY_FONT_NAME, size=14, bold=True, color=NAVY)

THIN = Side(style="thin", color="D9D9D9")
MEDIUM = Side(style="medium", color=NAVY)
DOUBLE = Side(style="double", color=NAVY)
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
KPI_BORDER = Border(left=MEDIUM, right=MEDIUM, top=MEDIUM, bottom=MEDIUM)
# Grand-total rows close with a double rule, the accounting convention.
TOTAL_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=DOUBLE)

INT_FMT = '#,##0'

CENTRE = Alignment(horizontal="center", vertical="center")
RIGHT = Alignment(horizontal="right", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")


def money_format() -> str:
    """Excel number format for money, honouring the selected currency symbol."""
    symbol = currency_symbol()
    return f'"{symbol}" #,##0.00' if symbol else '#,##0.00'


# Columns whose values read better centred: dates and identifier-style fields.
CENTRE_HINTS = ("date", "id", "code", "ref", "no", "number", "status", "period")


def is_centre_column(name: object) -> bool:
    """True for date and identifier columns, which centre better than they left-align."""
    key = norm_header(name)
    parts = key.split("_")
    return any(hint in parts or key.endswith("_" + hint) or key == hint
               for hint in CENTRE_HINTS) or key.startswith("date")


def _autofit(worksheet, frame: pd.DataFrame, start_col: int = 1,
             header_row: int = 1, min_width: int = 12, max_width: int = 52,
             padding: int = 7) -> None:
    """Approximate Excel's auto-fit using the longest rendered cell value."""
    for offset, column in enumerate(frame.columns):
        # An all-null column yields NaN from .max(), so coerce before rounding.
        values = frame[column].astype(str)
        longest = 0
        if len(values):
            widest = values.str.len().max()
            longest = 0 if pd.isna(widest) else int(widest)
        width = max(min_width,
                    min(max_width, max(longest, len(str(column))) + padding))
        worksheet.column_dimensions[get_column_letter(start_col + offset)].width = width
    worksheet.row_dimensions[header_row].height = HEADER_ROW_HEIGHT


def _write_table(worksheet, frame: pd.DataFrame, start_row: int = 1,
                 start_col: int = 1, money_cols=(), int_cols=(),
                 zebra: bool = True) -> int:
    """
    Write a DataFrame as an executive-styled table.

    Header: deep navy fill, bold white, 26pt, centred. Body: Calibri 10, 20pt
    rows, alternating zebra wash, money right-aligned in the selected currency,
    counts as integers, dates and identifiers centred.

    Returns the last row index used.
    """
    for offset, column in enumerate(frame.columns):
        cell = worksheet.cell(row=start_row, column=start_col + offset,
                              value=str(column))
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
        cell.border = BORDER
    worksheet.row_dimensions[start_row].height = HEADER_ROW_HEIGHT

    centre_cols = {c for c in frame.columns
                   if c not in money_cols and c not in int_cols
                   and is_centre_column(c)}

    for r_offset, (_, row) in enumerate(frame.iterrows(), start=1):
        excel_row = start_row + r_offset
        worksheet.row_dimensions[excel_row].height = DATA_ROW_HEIGHT
        striped = zebra and (r_offset % 2 == 0)
        for c_offset, column in enumerate(frame.columns):
            value = row[column]
            if pd.isna(value):
                value = ""
            if hasattr(value, "item"):
                value = value.item()
            cell = worksheet.cell(row=excel_row, column=start_col + c_offset,
                                  value=value)
            cell.border = BORDER
            cell.font = DATA_FONT
            if striped:
                cell.fill = ZEBRA_FILL
            if column in money_cols:
                cell.number_format = money_format()
                cell.alignment = RIGHT
            elif column in int_cols:
                cell.number_format = INT_FMT
                cell.alignment = RIGHT
            elif column in centre_cols:
                cell.alignment = CENTRE
            else:
                cell.alignment = LEFT
    return start_row + len(frame)


def _write_kpi_cards(worksheet, row: int, cards: list, span: int = 2) -> int:
    """
    Lay a row of KPI cards across the sheet: a navy label bar above a large
    value cell, each card spanning `span` columns and boxed in a medium border.

    `cards` is a list of (label, value, number_format) tuples. Returns the next
    free row.
    """
    col = 1
    for label, value, fmt in cards:
        end = col + span - 1
        worksheet.merge_cells(start_row=row, start_column=col,
                              end_row=row, end_column=end)
        worksheet.merge_cells(start_row=row + 1, start_column=col,
                              end_row=row + 1, end_column=end)

        head = worksheet.cell(row=row, column=col, value=str(label).upper())
        head.fill = KPI_LABEL_FILL
        head.font = KPI_LABEL_FONT
        head.alignment = CENTRE

        body = worksheet.cell(row=row + 1, column=col, value=value)
        body.fill = KPI_VALUE_FILL
        body.font = KPI_VALUE_FONT
        body.alignment = CENTRE
        if fmt:
            body.number_format = fmt

        # Border every merged cell so the box closes on all sides.
        for c in range(col, end + 1):
            worksheet.cell(row=row, column=c).border = KPI_BORDER
            worksheet.cell(row=row + 1, column=c).border = KPI_BORDER
        col = end + 1

    worksheet.row_dimensions[row].height = 18
    worksheet.row_dimensions[row + 1].height = 30
    return row + 3


def _write_total_row(worksheet, row: int, label: str, values: list,
                     start_col: int = 1) -> None:
    """Bold grand-total row closed with a double bottom border."""
    cell = worksheet.cell(row=row, column=start_col, value=label)
    cell.font = TOTAL_FONT
    cell.fill = TOTAL_FILL
    cell.border = TOTAL_BORDER
    cell.alignment = LEFT
    for offset, (value, fmt) in enumerate(values, start=start_col + 1):
        cell = worksheet.cell(row=row, column=offset, value=value)
        cell.font = TOTAL_FONT
        cell.fill = TOTAL_FILL
        cell.border = TOTAL_BORDER
        cell.alignment = RIGHT
        if fmt:
            cell.number_format = fmt
    worksheet.row_dimensions[row].height = DATA_ROW_HEIGHT


def highlight_mask(frame: pd.DataFrame) -> list:
    """Rows to tint soft red: validation failures plus urgent / review rows."""
    flagged = (frame["Flagged"] if "Flagged" in frame.columns
               else pd.Series(False, index=frame.index))
    review = (frame["Needs_Review"] if "Needs_Review" in frame.columns
              else pd.Series(False, index=frame.index))
    return (flagged.fillna(False) | review.fillna(False)).tolist()


def build_excel_report(frame: pd.DataFrame, result: dict, source_name: str) -> bytes:
    """Produce the two-tab, fully formatted .xlsx deliverable."""
    export = frame.drop(columns=HELPER_COLUMNS, errors="ignore").copy()
    # Missing IDs, unassigned statuses and urgent/review rows all highlight.
    flagged_flags = highlight_mask(frame)

    workbook = Workbook()

    # ------------------------- TAB 1: Cleaned Data ------------------------- #
    sheet = workbook.active
    sheet.title = "Cleaned Data"
    sheet.sheet_properties.tabColor = NAVY

    money_cols = ("Amount", "VAT_15", "Total_Inc_VAT")
    last_row = _write_table(sheet, export, start_row=1, start_col=1, money_cols=money_cols)

    # Highlight every flagged / missing-ID row in light red.
    for offset, is_flagged in enumerate(flagged_flags, start=2):
        if not is_flagged:
            continue
        for col_idx in range(1, len(export.columns) + 1):
            cell = sheet.cell(row=offset, column=col_idx)
            cell.fill = FLAG_FILL
            cell.font = FLAG_FONT

    sheet.freeze_panes = "A2"
    if last_row >= 1:
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(export.columns))}{last_row}"
    _autofit(sheet, export)

    # ------------------------ TAB 2: Executive Summary ---------------------- #
    summary = workbook.create_sheet("Executive Summary")
    summary.sheet_properties.tabColor = "10B981"
    summary["A1"] = "M&M RETAIL - EXECUTIVE SUMMARY"
    summary["A1"].font = TITLE_FONT
    summary["A2"] = (f"Source: {source_name}   |   Generated: "
                     f"{result['timestamp']:%Y-%m-%d %H:%M:%S}   |   "
                     f"Auto-Data Pro v{VERSION}")
    summary["A2"].font = SUB_FONT

    # KPI cards across the top of the sheet.
    _write_kpi_cards(summary, 4, [
        ("Rows Processed", result["rows_out"], INT_FMT),
        ("Flagged Rows", result["flagged"], INT_FMT),
        ("Health Score", f"{float(result.get('health', 0.0)):.1f}%", None),
        ("Total Revenue", round(result["revenue"], 2), money_format()),
        (f"VAT @ {result['vat_rate'] * 100:.0f}%",
         round(result["vat_total"], 2), money_format()),
    ])

    pivot = status_summary(frame)
    pivot_display = pivot.rename(columns={
        "Revenue_Excl_VAT": "Revenue (Excl VAT)",
        "VAT_15": f"VAT @ {result['vat_rate'] * 100:.0f}%",
        "Total_Inc_VAT": "Total (Incl VAT)",
        "Flagged_Rows": "Flagged Rows",
    })
    summary["A7"] = "REVENUE PIVOT BY STATUS"
    summary["A7"].font = LABEL_FONT
    vat_label = f"VAT @ {result['vat_rate'] * 100:.0f}%"
    money_display = ("Revenue (Excl VAT)", vat_label, "Total (Incl VAT)")
    pivot_last = _write_table(summary, pivot_display, start_row=8, start_col=1,
                              money_cols=money_display,
                              int_cols=("Transactions", "Flagged Rows"))

    total_row = pivot_last + 1
    money_fmt = money_format()
    _write_total_row(summary, total_row, "GRAND TOTAL", [
        (int(pivot_display["Transactions"].sum()), INT_FMT),
        (round(float(pivot_display["Revenue (Excl VAT)"].sum()), 2), money_fmt),
        (round(float(pivot_display[vat_label].sum()), 2), money_fmt),
        (round(float(pivot_display["Total (Incl VAT)"].sum()), 2), money_fmt),
        (int(pivot_display["Flagged Rows"].sum()), INT_FMT),
    ])
    _autofit(summary, pivot_display, header_row=8)
    return _finish_workbook(workbook, summary, result, total_row, pivot, source_name)


def _finish_workbook(workbook, summary, result: dict, total_row: int,
                     pivot: pd.DataFrame, source_name: str) -> bytes:
    """Append the run-metrics block and data-quality block, then serialise."""
    row = total_row + 3
    summary.cell(row=row, column=1, value="AUTOMATION RUN METRICS").font = LABEL_FONT
    row += 1

    vat_pct = f"{result['vat_rate'] * 100:.0f}%"
    metrics = [
        ("Source file", source_name),
        ("Rows received", result["rows_in"]),
        ("Rows processed", result["rows_out"]),
        ("Blank rows removed", result["blank_dropped"]),
        ("Duplicate rows removed", result["dup_dropped"]),
        ("Clean rows", result["clean"]),
        ("Flagged rows", result["flagged"]),
        (f"Total revenue (excl VAT)", round(result["revenue"], 2)),
        (f"Total VAT @ {vat_pct}", round(result["vat_total"], 2)),
        ("Total inclusive of VAT", round(result["inc_vat_total"], 2)),
        ("Business rules applied", result["rule_count"]),
        ("Processing time (seconds)", round(result["duration"], 2)),
        ("Cancelled excluded from revenue", "Yes" if result["exclude_cancelled"] else "No"),
        ("Generated", f"{result['timestamp']:%Y-%m-%d %H:%M:%S}"),
    ]
    for label, value in metrics:
        label_cell = summary.cell(row=row, column=1, value=label)
        label_cell.font = Font(bold=True, color="334155")
        value_cell = summary.cell(row=row, column=2, value=value)
        if isinstance(value, float):
            value_cell.number_format = money_format()
        elif isinstance(value, int):
            value_cell.number_format = INT_FMT
        value_cell.alignment = Alignment(horizontal="left")
        row += 1

    # Data-quality breakdown.
    row += 2
    summary.cell(row=row, column=1, value="DATA QUALITY - FLAG BREAKDOWN").font = LABEL_FONT
    row += 1
    header_a = summary.cell(row=row, column=1, value="Flag")
    header_b = summary.cell(row=row, column=2, value="Rows")
    for cell in (header_a, header_b):
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.border = BORDER
    row += 1
    if result["flag_breakdown"]:
        for label, count in result["flag_breakdown"].items():
            name_cell = summary.cell(row=row, column=1, value=label)
            count_cell = summary.cell(row=row, column=2, value=int(count))
            count_cell.number_format = INT_FMT
            for cell in (name_cell, count_cell):
                cell.fill = FLAG_FILL
                cell.font = FLAG_FONT
                cell.border = BORDER
            row += 1
    else:
        clean_cell = summary.cell(row=row, column=1,
                                  value="No exceptions raised - all rows passed validation")
        clean_cell.font = GREEN_FONT
        row += 1

    summary.column_dimensions["A"].width = max(34, summary.column_dimensions["A"].width or 0)
    summary.column_dimensions["B"].width = max(22, summary.column_dimensions["B"].width or 0)
    summary.freeze_panes = "A9"

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


# --------------------------------------------------------------------------- #
#  SUMMARY PDF  —  minimal, dependency-free PDF 1.4 writer
# --------------------------------------------------------------------------- #

PAGE_W, PAGE_H = 595, 842          # A4 at 72 dpi
MARGIN_X = 48
BODY_TOP = 760
LEADING = 15.5
BODY_BOTTOM = 60


def _pdf_escape(text: str) -> str:
    """Escape the three characters that are special inside a PDF string."""
    return (str(text).replace("\\", r"\\").replace("(", r"\(").replace(")", r"\)"))


def _pdf_text(text: str) -> str:
    """Coerce to a Latin-1-safe string (the base-14 fonts are single-byte)."""
    return str(text).encode("latin-1", "replace").decode("latin-1")


def _content_stream(lines: list[tuple], page_no: int, total_pages: int) -> str:
    """Build one page's content stream from (text, size, bold, colour) tuples."""
    parts = [
        # Navy masthead band.
        "0.059 0.090 0.165 rg 0 792 595 50 re f",
        # Emerald accent rule beneath it.
        "0.063 0.725 0.506 rg 0 788 595 4 re f",
        f"BT /F2 15 Tf 1 1 1 rg {MARGIN_X} 812 Td ({_pdf_escape('AUTO-DATA PRO')}) Tj ET",
        f"BT /F1 8.5 Tf 0.796 0.835 0.882 rg {MARGIN_X} 799 Td "
        f"({_pdf_escape('Excel Automation Dashboard  -  Executive Summary')}) Tj ET",
    ]
    y = BODY_TOP
    for text, size, bold, colour in lines:
        if text == "__RULE__":
            parts.append(f"0.886 0.910 0.941 rg {MARGIN_X} {y + 5} "
                         f"{PAGE_W - 2 * MARGIN_X} 0.8 re f")
            y -= LEADING
            continue
        font = "F2" if bold else "F1"
        r, g, b = colour
        parts.append(
            f"BT /{font} {size} Tf {r} {g} {b} rg {MARGIN_X} {y} Td "
            f"({_pdf_escape(_pdf_text(text))}) Tj ET"
        )
        y -= LEADING if size <= 11 else LEADING + 4

    footer = (f"Auto-Data Pro v{VERSION}  |  Client: M&M Retail  |  "
              f"Page {page_no} of {total_pages}")
    parts.append("0.886 0.910 0.941 rg 48 52 499 0.8 re f")
    parts.append(f"BT /F1 8 Tf 0.392 0.455 0.545 rg {MARGIN_X} 38 Td "
                 f"({_pdf_escape(footer)}) Tj ET")
    return "\n".join(parts)


def _assemble_pdf(pages: list[list[tuple]]) -> bytes:
    """Assemble page content streams into a valid single-file PDF."""
    total = len(pages)
    objects: dict[int, bytes] = {}
    page_obj_nums = [5 + 2 * i for i in range(total)]

    objects[1] = b"<< /Type /Catalog /Pages 2 0 R >>"
    kids = " ".join(f"{n} 0 R" for n in page_obj_nums)
    objects[2] = f"<< /Type /Pages /Kids [{kids}] /Count {total} >>".encode("latin-1")
    objects[3] = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding /WinAnsiEncoding >>"
    objects[4] = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold /Encoding /WinAnsiEncoding >>"

    for index, lines in enumerate(pages):
        page_num = page_obj_nums[index]
        content_num = page_num + 1
        stream = _content_stream(lines, index + 1, total).encode("latin-1", "replace")
        objects[page_num] = (
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {PAGE_W} {PAGE_H}] "
            f"/Resources << /Font << /F1 3 0 R /F2 4 0 R >> >> "
            f"/Contents {content_num} 0 R >>"
        ).encode("latin-1")
        objects[content_num] = (
            f"<< /Length {len(stream)} >>\nstream\n".encode("latin-1")
            + stream + b"\nendstream"
        )

    out = bytearray(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
    offsets: dict[int, int] = {}
    for num in sorted(objects):
        offsets[num] = len(out)
        out += f"{num} 0 obj\n".encode("latin-1") + objects[num] + b"\nendobj\n"

    xref_pos = len(out)
    max_num = max(objects)
    out += f"xref\n0 {max_num + 1}\n".encode("latin-1")
    out += b"0000000000 65535 f \n"
    for num in range(1, max_num + 1):
        if num in offsets:
            out += f"{offsets[num]:010d} 00000 n \n".encode("latin-1")
        else:
            out += b"0000000000 65535 f \n"
    out += (f"trailer\n<< /Size {max_num + 1} /Root 1 0 R >>\nstartxref\n"
            f"{xref_pos}\n%%EOF\n").encode("latin-1")
    return bytes(out)


BLACK = (0.118, 0.161, 0.231)      # #1E293B  high-contrast body ink
SLATE = (0.278, 0.333, 0.412)      # #475569  secondary text
NAVY_RGB = (0.059, 0.090, 0.165)   # #0F172A  Slate Navy
RED_RGB = (0.725, 0.110, 0.110)    # #B91C1C  deep coral (print legible)
GREEN_RGB = (0.016, 0.471, 0.341)  # #047857  deep emerald (print legible)


def build_summary_pdf(frame: pd.DataFrame, result: dict, source_name: str) -> bytes:
    """Render the executive summary as a paginated PDF."""
    vat_pct = f"{result['vat_rate'] * 100:.0f}%"
    lines: list[tuple] = []

    def add(text: str = "", size: float = 10, bold: bool = False, colour=BLACK) -> None:
        lines.append((text, size, bold, colour))

    add("MONTHLY SALES AUTOMATION REPORT", 14, True, NAVY_RGB)
    add(f"Client: M&M Retail   |   Source file: {source_name}", 9, False, SLATE)
    add(f"Generated: {result['timestamp']:%Y-%m-%d %H:%M:%S}   |   "
        f"Currency: ZAR   |   VAT rate: {vat_pct}", 9, False, SLATE)
    add("__RULE__")
    add()

    add("1.  PROCESSING RESULTS", 11, True, NAVY_RGB)
    add()
    for label, value in [
        ("Rows received", f"{result['rows_in']:,}"),
        ("Rows processed", f"{result['rows_out']:,}"),
        ("Blank rows removed", f"{result['blank_dropped']:,}"),
        ("Duplicate rows removed", f"{result['dup_dropped']:,}"),
        ("Clean rows", f"{result['clean']:,}"),
        ("Flagged errors", f"{result['flagged']:,}"),
        ("Business rules applied", f"{result['rule_count']}"),
        ("Processing time", f"{result['duration']:.2f} seconds"),
    ]:
        add(f"    {label:<32}{value}", 10, False, BLACK)
    add()

    add("2.  FINANCIAL TOTALS (ZAR)", 11, True, NAVY_RGB)
    add()
    add(f"    {'Total revenue (excl VAT)':<32}{fmt_money(result['revenue'], 2)}", 10, True, GREEN_RGB)
    add(f"    {'VAT @ ' + vat_pct:<32}{fmt_money(result['vat_total'], 2)}", 10, False, BLACK)
    add(f"    {'Total inclusive of VAT':<32}{fmt_money(result['inc_vat_total'], 2)}", 10, True, NAVY_RGB)
    if result["exclude_cancelled"]:
        add("    Cancelled / reversed entries are excluded from revenue totals.", 8.5, False, SLATE)
    add()

    add("3.  REVENUE PIVOT BY STATUS", 11, True, NAVY_RGB)
    add()
    add(f"    {'STATUS':<13}{'TXNS':>7}{'REVENUE':>16}{'VAT':>14}{'INCL VAT':>16}",
        9.5, True, SLATE)
    add("__RULE__")
    for _, row in status_summary(frame).iterrows():
        add(f"    {str(row['Status']):<13}{int(row['Transactions']):>7,}"
            f"{row['Revenue_Excl_VAT']:>16,.2f}{row['VAT_15']:>14,.2f}"
            f"{row['Total_Inc_VAT']:>16,.2f}", 9.5, False, BLACK)
    add("__RULE__")
    pivot = status_summary(frame)
    add(f"    {'TOTAL':<13}{int(pivot['Transactions'].sum()):>7,}"
        f"{pivot['Revenue_Excl_VAT'].sum():>16,.2f}{pivot['VAT_15'].sum():>14,.2f}"
        f"{pivot['Total_Inc_VAT'].sum():>16,.2f}", 9.5, True, NAVY_RGB)
    add()

    add("4.  DATA QUALITY - EXCEPTIONS RAISED", 11, True, NAVY_RGB)
    add()
    if result["flag_breakdown"]:
        for label, count in result["flag_breakdown"].items():
            share = (count / max(result["rows_out"], 1)) * 100
            add(f"    {label:<26}{count:>7,} rows   ({share:.1f}% of file)",
                10, False, RED_RGB)
    else:
        add("    No exceptions raised - every row passed validation.", 10, True, GREEN_RGB)
    add()

    add("5.  BUSINESS RULES APPLIED", 11, True, NAVY_RGB)
    add()
    codes = set(result["rules_applied"])
    for code, label, _ in RULE_LIBRARY:
        mark = "[X]" if code in codes else "[ ]"
        colour = BLACK if code in codes else SLATE
        add(f"    {mark} {code}  {label}", 9.5, False, colour)

    # Paginate.
    pages: list[list[tuple]] = []
    current: list[tuple] = []
    y = BODY_TOP
    for line in lines:
        step = LEADING if line[1] <= 11 else LEADING + 4
        if y - step < BODY_BOTTOM:
            pages.append(current)
            current, y = [], BODY_TOP
        current.append(line)
        y -= step
    pages.append(current)
    return _assemble_pdf(pages)



def build_universal_excel(frame: pd.DataFrame, result: dict, source_name: str) -> bytes:
    """
    Two-tab workbook for any spreadsheet.

    Tab 1 "Cleaned Dataset" - every original column, styled, flagged rows tinted.
    Tab 2 "Executive Summary" - dataset dimensions, column types, missing-value
    audit and statistics for every numeric and categorical column.
    """
    export = frame.drop(columns=HELPER_COLUMNS, errors="ignore").copy()
    flags = highlight_mask(frame)
    profile = result.get("profile") or {}
    money_cols = tuple(result.get("key_numeric", []))

    workbook = Workbook()

    # ------------------------- TAB 1: Cleaned Dataset ---------------------- #
    sheet = workbook.active
    sheet.title = "Cleaned Dataset"
    sheet.sheet_properties.tabColor = NAVY
    last_row = _write_table(sheet, export, start_row=1, start_col=1,
                            money_cols=money_cols)
    for offset, is_flagged in enumerate(flags, start=2):
        if not is_flagged:
            continue
        for col_idx in range(1, len(export.columns) + 1):
            cell = sheet.cell(row=offset, column=col_idx)
            cell.fill = FLAG_FILL
            cell.font = FLAG_FONT
    sheet.freeze_panes = "A2"
    if last_row >= 1 and len(export.columns):
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(export.columns))}{last_row}"
    _autofit(sheet, export)

    # ------------------------ TAB 2: Executive Summary --------------------- #
    summary = workbook.create_sheet("Executive Summary")
    summary.sheet_properties.tabColor = "10B981"
    summary["A1"] = "DATASET EXECUTIVE SUMMARY"
    summary["A1"].font = TITLE_FONT
    summary["A2"] = (f"Source: {source_name}   |   Generated: "
                     f"{result['timestamp']:%Y-%m-%d %H:%M:%S}   |   "
                     f"Auto-Data Pro v{VERSION}")
    summary["A2"].font = SUB_FONT

    numeric_total = float(result.get("numeric_total", 0.0))
    row = _write_kpi_cards(summary, 4, [
        ("Total Rows", result["rows_out"], INT_FMT),
        ("Total Columns", result.get("columns", export.shape[1]), INT_FMT),
        ("Health Score", f"{float(result.get('health', 0.0)):.1f}%", None),
        ("Rows With Issues", result["flagged"], INT_FMT),
        ("Numeric Total", round(numeric_total, 2), money_format()),
    ])

    summary.cell(row=row, column=1, value="DATASET OVERVIEW").font = LABEL_FONT
    row += 1
    overview = [
        ("Source file", source_name),
        ("Total rows", result["rows_out"]),
        ("Total columns", result.get("columns", export.shape[1])),
        ("Total cells", result.get("total_cells", int(export.size))),
        ("Missing cells", result.get("missing_cells", 0)),
        ("Duplicate rows", result.get("duplicate_rows", 0)),
        ("Blank rows removed", result.get("blank_dropped", 0)),
        ("Rows with issues", result["flagged"]),
        ("Clean rows", result["clean"]),
        ("Dataset health score (%)", round(float(result.get("health", 0.0)), 1)),
        ("Data quality issues raised", result.get("rule_count", 0)),
        ("Processing time (seconds)", round(result["duration"], 2)),
    ]
    for label, value in overview:
        head = summary.cell(row=row, column=1, value=label)
        head.font = Font(name=BODY_FONT_NAME, size=BODY_FONT_SIZE,
                         bold=True, color="334155")
        head.alignment = LEFT
        cell = summary.cell(row=row, column=2, value=value)
        cell.font = DATA_FONT
        cell.alignment = RIGHT if isinstance(value, (int, float)) else LEFT
        cell.number_format = INT_FMT if isinstance(value, int) else (
            money_format() if isinstance(value, float) else "General")
        summary.row_dimensions[row].height = DATA_ROW_HEIGHT
        row += 1

    sections = [
        ("COLUMN DATA TYPES", profile_table(profile) if profile else pd.DataFrame(),
         ("Non-Null", "Missing", "Unique"), ()),
        ("NUMERIC COLUMN STATISTICS", result.get("numeric_stats", pd.DataFrame()),
         ("Count", "Missing"), ("Sum", "Mean", "Median", "Min", "Max", "Std Dev")),
        ("CATEGORICAL / TEXT COLUMN BREAKDOWN",
         result.get("categorical_stats", pd.DataFrame()),
         ("Unique", "Frequency", "Missing"), ()),
        ("DATETIME COVERAGE", result.get("datetime_stats", pd.DataFrame()),
         ("Distinct Dates", "Unparseable", "Missing"), ()),
        ("DATA QUALITY AUDIT", result.get("audit", pd.DataFrame()),
         ("Rows",), ()),
    ]
    for heading, table, int_cols, money in sections:
        if table is None or table.empty:
            continue
        row += 2
        summary.cell(row=row, column=1, value=heading).font = LABEL_FONT
        row += 1
        row = _write_table(summary, table, start_row=row, start_col=1,
                           money_cols=money, int_cols=int_cols)
        if heading == "NUMERIC COLUMN STATISTICS" and "Sum" in table.columns:
            row += 1
            money_fmt = money_format()
            _write_total_row(summary, row, "GRAND TOTAL", [
                (int(table["Count"].sum()), INT_FMT),
                (round(float(table["Sum"].sum()), 2), money_fmt),
            ])
        row += 1

    summary.column_dimensions["A"].width = 34
    for idx in range(2, 9):
        summary.column_dimensions[get_column_letter(idx)].width = 20
    summary.freeze_panes = "A7"

    buffer = io.BytesIO()
    workbook.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def build_universal_pdf(frame: pd.DataFrame, result: dict, source_name: str) -> bytes:
    """Executive summary PDF for any dataset."""
    lines: list[tuple] = []

    def add(text: str = "", size: float = 10, bold: bool = False, colour=BLACK) -> None:
        lines.append((text, size, bold, colour))

    profile = result.get("profile") or {}
    health = float(result.get("health", 0.0))

    add("DATASET PROCESSING REPORT", 14, True, NAVY_RGB)
    add(f"Source file: {source_name}", 9, False, SLATE)
    add(f"Generated: {result['timestamp']:%Y-%m-%d %H:%M:%S}   |   "
        f"Engine: {ENGINE_LABEL}", 9, False, SLATE)
    add("__RULE__")
    add()

    add("1.  DATASET DIMENSIONS", 11, True, NAVY_RGB)
    add()
    for label, value in [
        ("Rows received", f"{result['rows_in']:,}"),
        ("Rows processed", f"{result['rows_out']:,}"),
        ("Total columns", f"{result.get('columns', 0):,}"),
        ("Total cells", f"{result.get('total_cells', 0):,}"),
        ("Blank rows removed", f"{result.get('blank_dropped', 0):,}"),
        ("Duplicate rows", f"{result.get('duplicate_rows', 0):,}"),
        ("Processing time", f"{result['duration']:.2f} seconds"),
    ]:
        add(f"    {label:<30}{value}", 10, False, BLACK)
    add()
    add(f"    {'Dataset health score':<30}{health:.1f}%", 10, True,
        GREEN_RGB if health >= 70 else RED_RGB)
    add(f"    {'Clean rows':<30}{result['clean']:,} of {result['rows_out']:,}",
        10, False, BLACK)
    add()

    add("2.  COLUMN DATA TYPES", 11, True, NAVY_RGB)
    add()
    add(f"    {'COLUMN':<26}{'TYPE':<14}{'NON-NULL':>10}{'MISSING':>10}{'UNIQUE':>9}",
        9.5, True, SLATE)
    add("__RULE__")
    if profile:
        for column in profile["__order__"]:
            entry = profile[column]
            add(f"    {str(column)[:25]:<26}{entry['type']:<14}"
                f"{entry['non_null']:>10,}{entry['nulls']:>10,}{entry['unique']:>9,}",
                9.5, False, RED_RGB if entry["nulls"] else BLACK)
    add()

    add("3.  MISSING VALUE AUDIT", 11, True, NAVY_RGB)
    add()
    missing = [(c, profile[c]) for c in profile.get("__order__", [])
               if profile[c]["nulls"]] if profile else []
    if missing:
        for column, entry in missing:
            add(f"    {str(column)[:26]:<28}{entry['nulls']:>7,} missing "
                f"({entry['null_pct']:.1f}% of file)", 10, False, RED_RGB)
    else:
        add("    No missing values - every cell is populated.", 10, True, GREEN_RGB)
    add(f"    {'TOTAL MISSING CELLS':<28}{result.get('missing_cells', 0):>7,} of "
        f"{result.get('total_cells', 0):,}", 10, True, NAVY_RGB)
    add()

    numeric_stats = result.get("numeric_stats")
    if numeric_stats is not None and not numeric_stats.empty:
        add("4.  NUMERIC COLUMN STATISTICS", 11, True, NAVY_RGB)
        add()
        add(f"    {'COLUMN':<20}{'SUM':>16}{'MEAN':>14}{'MIN':>13}{'MAX':>15}",
            9.5, True, SLATE)
        add("__RULE__")
        for _, r in numeric_stats.iterrows():
            add(f"    {str(r['Column'])[:19]:<20}{r['Sum']:>16,.2f}{r['Mean']:>14,.2f}"
                f"{r['Min']:>13,.2f}{r['Max']:>15,.2f}", 9.5, False, BLACK)
        add()

    cat_stats = result.get("categorical_stats")
    if cat_stats is not None and not cat_stats.empty:
        add("5.  CATEGORICAL BREAKDOWN", 11, True, NAVY_RGB)
        add()
        for _, r in cat_stats.iterrows():
            add(f"    {str(r['Column'])[:24]:<26}{int(r['Unique']):>5,} unique   "
                f"top: {str(r['Most Common'])[:22]} ({r['Share %']:.0f}%)",
                9.5, False, BLACK)
        add()

    audit = result.get("audit")
    add("6.  DATA QUALITY AUDIT", 11, True, NAVY_RGB)
    add()
    if audit is not None and not audit.empty:
        add(f"    {'COLUMN':<20}{'ISSUE':<40}{'ROWS':>8}{'SEV':>9}", 9.5, True, SLATE)
        add("__RULE__")
        for _, r in audit.iterrows():
            add(f"    {str(r['Column'])[:19]:<20}{str(r['Issue'])[:39]:<40}"
                f"{int(r['Rows']):>8,}{str(r['Severity']):>9}", 9.5, False, RED_RGB)
    else:
        add("    No data quality issues detected.", 10, True, GREEN_RGB)

    pages, current, y = [], [], BODY_TOP
    for line in lines:
        step = LEADING if line[1] <= 11 else LEADING + 4
        if y - step < BODY_BOTTOM:
            pages.append(current)
            current, y = [], BODY_TOP
        current.append(line)
        y -= step
    pages.append(current)
    return _assemble_pdf(pages)


# --------------------------------------------------------------------------- #
#  UI COMPONENTS
# --------------------------------------------------------------------------- #

def render_header() -> None:
    st.markdown(
        f"""
        <div class="adp-header">
          <div>
            <div class="adp-header-title">{APP_TITLE}</div>
            <div class="adp-header-sub">{APP_SUBTITLE}</div>
          </div>
          <div class="adp-logo">
            <div class="adp-logo-mark">AD</div>
            <div>
              <div class="adp-logo-text">{BRAND_NAME}</div>
              <div class="adp-logo-tag">AUTOMATION SUITE v{VERSION}</div>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar() -> str:
    with st.sidebar:
        st.markdown(
            """
            <div class="adp-side-brand">
              <div class="adp-logo-mark" style="width:34px;height:34px;font-size:14px;">AD</div>
              <div>
                <div class="adp-logo-text" style="font-size:15px;">Auto-Data Pro</div>
                <div class="adp-logo-tag">EXCEL AUTOMATION</div>
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown('<div class="adp-section-h">Navigation</div>', unsafe_allow_html=True)
        labels = [name for name, _ in NAV_ITEMS]
        choice = st.radio(
            "Navigation", labels,
            index=labels.index(st.session_state["nav"]) if st.session_state["nav"] in labels else 0,
            label_visibility="collapsed", key="nav_radio",
        )
        st.session_state["nav"] = choice
        st.caption(dict(NAV_ITEMS)[choice])

        st.markdown("---")
        st.markdown('<div class="adp-section-h">Local Environment</div>', unsafe_allow_html=True)
        engine_state = "RULES ENGINE v2.4 (ACTIVE) - IDLE"
        dot = "adp-dot-amber"
        if st.session_state["result"] is not None:
            engine_state = "RULES ENGINE v2.4 (ACTIVE) - LAST RUN OK"
            dot = "adp-dot-green"
        uptime = datetime.now() - st.session_state["boot_time"]
        minutes = int(uptime.total_seconds() // 60)
        st.markdown(
            f"""
            <div class="adp-side-env">
              <span class="adp-dot adp-dot-green"></span><b>LOCALHOST:8501</b> - CONNECTED<br>
              <span class="adp-dot {dot}"></span>{engine_state}<br>
              <span class="adp-dot adp-dot-green"></span>PANDAS {pd.__version__} / OPENPYXL<br>
              <span class="adp-dot {'adp-dot-green' if ARROW_AVAILABLE else 'adp-dot-amber'}"></span>GRID: {'ARROW NATIVE' if ARROW_AVAILABLE else 'HTML FALLBACK'}<br>
              <span class="adp-dot adp-dot-green"></span>VAT PROFILE: ZA {st.session_state['vat_rate'] * 100:.0f}%<br>
              <span class="adp-dot adp-dot-green"></span>CURRENCY: {currency_symbol() or 'NONE'}<br>
              <span class="adp-dot adp-dot-green"></span>SESSION UPTIME: {minutes} MIN<br>
              <span class="adp-dot adp-dot-green"></span>RUNS THIS SESSION: {len(st.session_state['history'])}
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown("---")
        if st.button("Reset workspace", use_container_width=True):
            for key in ("raw_df", "raw_meta", "clean_df", "result"):
                st.session_state[key] = None
            st.rerun()
        st.caption(f"Session started {st.session_state['boot_time']:%Y-%m-%d %H:%M}")
    return choice


def panel_title(title: str, note: str = "") -> None:
    """Panel heading rendered inside a bordered st.container."""
    html = f'<div class="adp-panel-title">{title}</div>'
    if note:
        html += f'<div class="adp-panel-note">{note}</div>'
    st.markdown(html, unsafe_allow_html=True)


def metric_card(label: str, value: str, foot: str = "", tone: str = "") -> str:
    tone_class = f" {tone}" if tone else ""
    return (
        f'<div class="adp-metric{tone_class}">'
        f'<div class="adp-metric-label">{label}</div>'
        f'<div class="adp-metric-value{tone_class}">{value}</div>'
        f'<div class="adp-metric-foot">{foot}</div></div>'
    )
PREVIEW_COLUMNS = ["Date", "Client_ID", "Status", "Recon_Status", "Amount", "Notes"]


def render_preview_table(frame: pd.DataFrame, limit: int = 250,
                         columns: list | None = None) -> None:
    """Render the cleaned data as an HTML table with conditional formatting.

    Columns default to the sales preview set but any column list may be passed,
    so a file with an arbitrary schema renders its own columns.
    """
    if frame.empty:
        st.info("No rows match the current filter.")
        return

    subset = frame.head(limit)
    wanted = columns if columns else PREVIEW_COLUMNS
    columns = [c for c in wanted if c in subset.columns]
    if not columns:
        columns = [c for c in subset.columns
                   if c not in HELPER_COLUMNS][:12]
    labels = {"Amount": "Amount (ZAR)", "Recon_Status": "Reconciliation"}
    head_cells = "".join(
        f"<th>{labels.get(c, c.replace('_', ' '))}</th>" for c in columns
    )

    rows_html = []
    for _, row in subset.iterrows():
        flagged = bool(row.get("Flagged", False))
        needs_review = bool(row.get("Needs_Review", False))
        cells = []
        for column in columns:
            value = row[column]
            if column == "Amount":
                cells.append(f'<td class="num">{fmt_money(value, 2)}</td>')
            elif isinstance(value, (int, float)) and not isinstance(value, bool):
                cells.append(f'<td class="num">{"" if pd.isna(value) else f"{value:,.2f}"}</td>')
            elif column == "Status":
                # Canonical statuses get their own colour; a preserved custom
                # value renders in a neutral slate pill with its text intact.
                status = str(value or "").strip() or STATUS_BLANK
                if any(w in status.lower() for w in REVIEW_KEYWORDS):
                    pill = "REVIEW"
                else:
                    pill = status if status in STATUS_ORDER else "CUSTOM"
                cells.append(f'<td><span class="adp-pill adp-pill-{pill}">'
                             f'{esc(status)}</span></td>')
            elif column == "Recon_Status":
                recon = str(value or RECON_SKIPPED)
                pill = {RECON_OK: "CLEAN", RECON_NONE: "FLAG"}.get(recon, "CANCELLED")
                cells.append(f'<td><span class="adp-pill adp-pill-{pill}">'
                             f'{esc(recon)}</span></td>')
            elif column == "Notes":
                # Lead badge states the row verdict: coral flags for validation
                # errors, amber NEEDS REVIEW when the row asks for attention,
                # emerald CLEAN only when neither applies. Every trailing note
                # is a pill too, so the column is badges end to end.
                tags = []
                if not flagged:
                    tags.append(
                        f'<span class="adp-pill adp-pill-REVIEW">{REVIEW_LABEL}</span>'
                        if needs_review else
                        '<span class="adp-pill adp-pill-CLEAN">CLEAN</span>'
                    )
                for part in (t.strip() for t in str(value or "").split("|")):
                    if not part:
                        continue
                    if part in KNOWN_FLAGS:
                        tags.append('<span class="adp-pill adp-pill-FLAG">'
                                    f'{esc(part)}</span>')
                    else:
                        # Secondary note text - plain, left aligned, no pill.
                        tags.append(f'<span class="adp-note-text">'
                                    f'{esc(part)}</span>')
                cells.append(f'<td class="notes-cell">{" ".join(tags)}</td>')
            else:
                text = str(value or "")
                cells.append(f"<td>{esc(text) if text else '&mdash;'}</td>")
        css_class = (' class="flagged"' if flagged
                     else (' class="review"' if needs_review else ""))
        rows_html.append(f"<tr{css_class}>{''.join(cells)}</tr>")

    st.markdown(
        f'<div class="adp-table-wrap adp-table-preview"><table class="adp-table">'
        f"<thead><tr>{head_cells}</tr></thead>"
        f'<tbody>{"".join(rows_html)}</tbody></table></div>',
        unsafe_allow_html=True,
    )
    if len(frame) > limit:
        st.caption(f"Showing the first {limit:,} of {len(frame):,} matching rows. "
                   "The downloaded Excel file always contains every row.")
    else:
        st.caption(f"Showing all {len(frame):,} matching rows.")


# --------------------------------------------------------------------------- #
#  RESILIENT TABLE & CHART RENDERERS
#  Streamlit's native grid and charts need pyarrow; when its DLL is blocked by a
#  Windows Application Control policy these HTML renderers keep the app usable.
# --------------------------------------------------------------------------- #

def df_to_html(frame: pd.DataFrame, max_rows: int = 500,
               max_height: int = 430) -> str:
    """Render any DataFrame as a styled, scrollable HTML table."""
    subset = frame.head(max_rows)
    numeric = {c for c in subset.columns
               if pd.api.types.is_numeric_dtype(subset[c])
               and not pd.api.types.is_bool_dtype(subset[c])}
    heads = "".join(f"<th>{esc(str(c).replace('_', ' '))}</th>" for c in subset.columns)

    body = []
    for _, row in subset.iterrows():
        cells = []
        for column in subset.columns:
            value = row[column]
            if is_blank_value(value):
                cells.append("<td>&mdash;</td>")
            elif column in numeric:
                try:
                    number = float(value)
                    text = (f"{int(number):,}" if float(number).is_integer()
                            else f"{number:,.2f}")
                except (TypeError, ValueError):
                    text = esc(value)
                cells.append(f'<td class="num">{text}</td>')
            else:
                cells.append(f"<td>{esc(value)}</td>")
        body.append(f"<tr>{''.join(cells)}</tr>")

    footer = ""
    if len(frame) > max_rows:
        footer = (f'<div class="adp-fallback-note">Showing the first {max_rows:,} '
                  f'of {len(frame):,} rows.</div>')
    return (
        f'<div class="adp-table-wrap" style="max-height:{max_height}px;">'
        f'<table class="adp-table"><thead><tr>{heads}</tr></thead>'
        f'<tbody>{"".join(body)}</tbody></table></div>{footer}'
    )


def render_table(frame: pd.DataFrame, height: int | None = None,
                 max_rows: int = 500) -> None:
    """st.dataframe when pyarrow is available, otherwise an HTML table."""
    if ARROW_AVAILABLE:
        st.dataframe(frame, hide_index=True, use_container_width=True, height=height)
        return
    st.markdown(df_to_html(frame, max_rows=max_rows, max_height=height or 430),
                unsafe_allow_html=True)


def render_bar_chart(series: pd.Series, height: int = 300, money: bool = False) -> None:
    """Horizontal bar chart with a pure-HTML fallback."""
    if ARROW_AVAILABLE:
        st.bar_chart(series, height=height)
        return
    values = series.astype(float)
    peak = max(abs(float(values.max())), abs(float(values.min())), 1.0)
    rows = []
    for label, value in values.items():
        width = abs(float(value)) / peak * 100
        text = fmt_money(value, 2) if money else f"{float(value):,.0f}"
        negative = " neg" if float(value) < 0 else ""
        rows.append(
            f'<div class="adp-bar-row">'
            f'<div class="adp-bar-label">{esc(label)}</div>'
            f'<div class="adp-bar-track">'
            f'<div class="adp-bar-fill{negative}" style="width:{width:.1f}%"></div></div>'
            f'<div class="adp-bar-value">{text}</div></div>'
        )
    st.markdown(f'<div class="adp-bars">{"".join(rows)}</div>', unsafe_allow_html=True)


def render_line_chart(series: pd.Series, height: int = 300) -> None:
    """Time-series line chart with an inline-SVG fallback."""
    if ARROW_AVAILABLE:
        st.line_chart(series, height=height)
        return
    values = [float(v) for v in series.tolist()]
    labels = [str(i) for i in series.index.tolist()]
    if not values:
        st.info("Not enough data to plot a trend.")
        return

    width, pad_l, pad_r, pad_t, pad_b = 900, 74, 18, 18, 34
    plot_w = width - pad_l - pad_r
    plot_h = height - pad_t - pad_b
    low, high = min(values), max(values)
    span = (high - low) or (abs(high) or 1.0)
    step = plot_w / max(len(values) - 1, 1)

    def y_of(value: float) -> float:
        return pad_t + plot_h - ((value - low) / span) * plot_h

    points = " ".join(f"{pad_l + i * step:.1f},{y_of(v):.1f}"
                      for i, v in enumerate(values))
    area = (f"{pad_l:.1f},{pad_t + plot_h:.1f} " + points +
            f" {pad_l + (len(values) - 1) * step:.1f},{pad_t + plot_h:.1f}")

    grid, ticks = [], []
    for frac in (0, 0.25, 0.5, 0.75, 1):
        value = low + span * (1 - frac)
        y = pad_t + plot_h * frac
        grid.append(f'<line x1="{pad_l}" y1="{y:.1f}" x2="{width - pad_r}" '
                    f'y2="{y:.1f}" stroke="#E2E8F0" stroke-width="1"/>')
        ticks.append(f'<text x="{pad_l - 8}" y="{y + 4:.1f}" text-anchor="end" '
                     f'font-size="10" fill="#64748B">{value:,.0f}</text>')

    x_labels = []
    for idx in {0, len(labels) // 2, len(labels) - 1}:
        anchor = "start" if idx == 0 else ("end" if idx == len(labels) - 1 else "middle")
        x_labels.append(
            f'<text x="{pad_l + idx * step:.1f}" y="{height - 10}" '
            f'text-anchor="{anchor}" font-size="10" fill="#64748B">'
            f'{esc(labels[idx])}</text>'
        )

    st.markdown(
        f'<svg class="adp-spark" viewBox="0 0 {width} {height}" '
        f'preserveAspectRatio="xMidYMid meet" role="img">'
        f'{"".join(grid)}'
        f'<polygon points="{area}" fill="#1F3A5F" fill-opacity="0.10"/>'
        f'<polyline points="{points}" fill="none" stroke="#1F3A5F" '
        f'stroke-width="2" stroke-linejoin="round" stroke-linecap="round"/>'
        f'{"".join(ticks)}{"".join(x_labels)}</svg>',
        unsafe_allow_html=True,
    )


def current_options() -> dict:
    """Collect the rule-engine options from session state."""
    return {
        "vat_rate": st.session_state["vat_rate"],
        "flag_missing_id": st.session_state["opt_flag_missing_id"],
        "flag_missing_status": st.session_state["opt_flag_missing_status"],
        "flag_bad_dates": st.session_state["opt_flag_bad_dates"],
        "flag_nonpositive": st.session_state["opt_flag_nonpositive"],
        "drop_duplicates": st.session_state["opt_drop_duplicates"],
        "flag_duplicates": st.session_state["opt_flag_duplicates"],
        "bank_recon": st.session_state["opt_bank_recon"],
        "dayfirst": st.session_state["opt_dayfirst"],
        "exclude_cancelled": st.session_state["opt_exclude_cancelled"],
        "min_amount": st.session_state["opt_min_amount"],
        "outlier_sigma": st.session_state["opt_outlier_sigma"],
        "flag_nulls": st.session_state["opt_flag_nulls"],
        "flag_outliers": st.session_state["opt_flag_outliers"],
    }


def stage_dataframe(frame: pd.DataFrame, name: str, size_bytes: int, source: str) -> None:
    """Register a freshly-loaded DataFrame as the active raw file."""
    st.session_state["raw_df"] = frame
    st.session_state["raw_meta"] = {
        "name": name,
        "size_bytes": size_bytes,
        "rows": len(frame),
        "cols": frame.shape[1],
        "source": source,
        "loaded": datetime.now(),
    }
    st.session_state["clean_df"] = None
    st.session_state["result"] = None


def execute_engine() -> bool:
    """Run the engine over the staged file and record the run in history."""
    raw = st.session_state["raw_df"]
    meta = st.session_state["raw_meta"]
    if raw is None:
        st.warning("Upload a raw monthly file before running the automation engine.")
        return False
    options = current_options()
    try:
        # Nothing is assumed about the incoming schema: the ledger rules only
        # run when the file actually resolves to a client-transaction ledger,
        # otherwise the universal profiler handles whatever columns exist.
        _, mapping, _ = map_columns(raw)
        if is_sales_schema(mapping):
            clean, result = run_automation_engine(raw, options)
        else:
            clean, result = run_universal_engine(raw, options)
    except Exception as exc:                      # surfaced to the operator, not swallowed
        st.error(f"Processing engine failed: {exc}")
        return False

    st.session_state["clean_df"] = clean
    st.session_state["result"] = result
    st.session_state["history"].insert(0, {
        "Run": f"{result['timestamp']:%Y-%m-%d %H:%M:%S}",
        "Source File": meta["name"],
        "Mode": result.get("mode", "sales").title(),
        "Rows In": result["rows_in"],
        "Rows Out": result["rows_out"],
        "Columns": result.get("columns", 0),
        "Flagged": result["flagged"],
        "Health %": round(result.get("health", 0.0), 1),
        "Numeric Total": round(result.get("numeric_total", 0.0), 2),
        "Seconds": round(result["duration"], 2),
    })
    return True


# --------------------------------------------------------------------------- #
#  PANEL 1  —  UPLOAD RAW MONTHLY FILE
# --------------------------------------------------------------------------- #

def panel_upload() -> None:
    with st.container(border=True):
        panel_title(
            "1. Upload Raw Monthly File (.xlsx / .csv)",
            "Drag and drop the raw export straight from the till system, or browse "
            "for it. Column headers are auto-detected.",
        )
        left, right = st.columns([2.1, 1], gap="large")

        with left:
            uploaded = st.file_uploader(
                "Raw monthly sales export",
                type=["xlsx", "xls", "csv"],
                label_visibility="collapsed",
                key="uploader",
            )
            if uploaded is not None:
                meta = st.session_state["raw_meta"]
                signature = (uploaded.name, uploaded.size)
                if meta is None or (meta.get("name"), meta.get("size_bytes")) != signature:
                    try:
                        frame = read_uploaded_file(uploaded)
                        stage_dataframe(frame, uploaded.name, uploaded.size, "upload")
                    except Exception as exc:
                        st.error(f"Could not read **{uploaded.name}**: {exc}")

            meta = st.session_state["raw_meta"]
            if meta:
                st.markdown(
                    f"""
                    <div class="adp-file adp-file-ok">
                      <div>
                        <div class="adp-file-name">{esc(meta['name'])}</div>
                        <div class="adp-file-meta">
                          {meta['rows']:,} rows &nbsp;&bull;&nbsp; {meta['cols']} columns
                          &nbsp;&bull;&nbsp; {fmt_size(meta['size_bytes'])}
                          &nbsp;&bull;&nbsp; loaded {meta['loaded']:%H:%M:%S}
                        </div>
                      </div>
                      <div><span class="adp-badge adp-badge-green">UPLOADED</span></div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )
            else:
                st.markdown(
                    """
                    <div class="adp-file adp-file-idle">
                      <div>
                        <div class="adp-file-name">No file staged</div>
                        <div class="adp-file-meta">
                          Accepted formats: .xlsx, .xls, .csv &nbsp;&bull;&nbsp; max 200MB
                        </div>
                      </div>
                      <div><span class="adp-badge adp-badge-slate">AWAITING FILE</span></div>
                    </div>
                    """,
                    unsafe_allow_html=True,
                )

        with right:
            st.markdown('<div class="adp-section-h">Run Controls</div>', unsafe_allow_html=True)
            ready = st.session_state["raw_df"] is not None
            if st.button("Process Data File", type="primary",
                         use_container_width=True, disabled=not ready,
                         key="btn_run_engine"):
                with st.spinner("Applying the South African business-rules engine..."):
                    if execute_engine():
                        st.toast("Automation complete", icon="✅")
                st.rerun()

            if st.button("Load sample M&M Retail dataset", use_container_width=True,
                         key="btn_sample"):
                sample = build_sample_dataset()
                buffer = io.BytesIO()
                sample.to_csv(buffer, index=False)
                stage_dataframe(sample, "Sept_Raw_Sales.xlsx",
                                buffer.tell(), "sample")
                st.rerun()

            st.caption(
                f"Engine profile: **ZA VAT {st.session_state['vat_rate'] * 100:.0f}%** · "
                f"**{len(RULE_LIBRARY)} rules** · dates "
                f"{'day-first' if st.session_state['opt_dayfirst'] else 'month-first'}"
            )


# --------------------------------------------------------------------------- #
#  PANEL 2  —  PROCESSING RESULTS & SUMMARY
# --------------------------------------------------------------------------- #

def health_tone(health: float) -> str:
    """Card accent for a health score."""
    return "green" if health >= 90 else ("" if health >= 70 else "red")


def panel_results() -> None:
    result = st.session_state["result"]
    with st.container(border=True):
        panel_title("Processing Results & Summary",
                    "Live metrics from the most recent run. Every figure is "
                    "derived from the uploaded file - no fixed schema assumed.")

        if result is None:
            cols = st.columns(4, gap="medium")
            for col, label in zip(cols, ("Total Rows Processed", "Total Columns",
                                         "Dataset Health Score", "Numeric Total")):
                with col:
                    st.markdown(metric_card(label, "—",
                                            "Awaiting a processing run"),
                                unsafe_allow_html=True)
            st.info(f"{ENGINE_LABEL} - idle. Stage a raw data file, then press "
                    "PROCESS DATA FILE.")
            return

        rows = result["rows_out"]
        health = float(result.get("health", 0.0))
        flagged = result["flagged"]
        error_rate = (flagged / max(rows, 1)) * 100
        mode = result.get("mode", "sales")

        c1, c2, c3, c4 = st.columns(4, gap="medium")
        with c1:
            st.markdown(metric_card(
                "Total Rows Processed",
                f"{rows:,}",
                f"{result['rows_in']:,} received &nbsp;·&nbsp; "
                f"{result['blank_dropped']:,} blank removed &nbsp;·&nbsp; "
                f"{result.get('dup_dropped', 0):,} duplicates removed",
            ), unsafe_allow_html=True)
        with c2:
            st.markdown(metric_card(
                "Total Columns",
                f"{result.get('columns', 0):,}",
                f"{result.get('missing_cells', 0):,} missing of "
                f"{result.get('total_cells', 0):,} cells &nbsp;·&nbsp; "
                f"{result.get('duplicate_rows', 0):,} duplicate rows",
            ), unsafe_allow_html=True)
        with c3:
            st.markdown(metric_card(
                "Dataset Health Score",
                f"{health:.1f}%",
                f"{result['clean']:,} complete &amp; un-flagged rows &nbsp;·&nbsp; "
                f"<span class='adp-badge adp-badge-red'>{flagged:,} FLAGGED "
                f"({error_rate:.1f}%)</span>",
                tone=health_tone(health),
            ), unsafe_allow_html=True)
        with c4:
            keys = result.get("key_numeric", [])
            if mode == "sales":
                label, value = "Total Revenue", fmt_money(result["revenue"])
                foot = (f"VAT @ {result['vat_rate'] * 100:.0f}%: "
                        f"<b>{fmt_money(result['vat_total'])}</b> &nbsp;·&nbsp; "
                        f"Incl VAT: <b>{fmt_money(result['inc_vat_total'])}</b>")
            else:
                label = "Total Numeric Sum"
                value = fmt_money(result.get("numeric_total", 0.0))
                foot = ("Summing " + ", ".join(f"<b>{esc(k)}</b>" for k in keys)
                        if keys else "No numeric columns detected")
            st.markdown(metric_card(label, value, foot, tone="green"),
                        unsafe_allow_html=True)

        st.markdown("<div style='height:16px'></div>", unsafe_allow_html=True)
        st.markdown('<div class="adp-section-h">Exception Breakdown</div>',
                    unsafe_allow_html=True)
        if result["flag_breakdown"]:
            breakdown = pd.DataFrame(
                [{"Exception": k, "Rows": v} for k, v in result["flag_breakdown"].items()]
            )
            breakdown["% of File"] = (
                breakdown["Rows"] / max(rows, 1) * 100
            ).map(lambda v: f"{v:.1f}%")
            render_table(breakdown, height=min(340, 48 + 35 * len(breakdown)))
        else:
            st.success("No exceptions raised - every row passed validation.")


def panel_profile_and_audit() -> None:
    """Column typing, data quality audit and categorical distributions."""
    result = st.session_state["result"]
    if result is None:
        return

    with st.container(border=True):
        panel_title("Dataset Profile & Data Quality Audit",
                    "Column types are inferred from the data itself; the audit "
                    "lists every issue found, per column.")
        tab_audit, tab_types, tab_num, tab_cat = st.tabs(
            ["Data Quality Audit", "Column Data Types",
             "Numeric Summary", "Categorical Breakdown"])

        with tab_audit:
            audit = result.get("audit")
            if audit is None or audit.empty:
                st.success("No data quality issues detected in this file.")
            else:
                high = int((audit["Severity"] == SEV_HIGH).sum())
                med = int((audit["Severity"] == SEV_MEDIUM).sum())
                low = int((audit["Severity"] == SEV_LOW).sum())
                st.markdown(
                    f'<div style="margin-bottom:10px;">'
                    f'<span class="adp-badge adp-badge-red">{high} HIGH</span> '
                    f'<span class="adp-badge adp-badge-amber">{med} MEDIUM</span> '
                    f'<span class="adp-badge adp-badge-slate">{low} LOW</span></div>',
                    unsafe_allow_html=True,
                )
                render_table(audit, height=min(430, 48 + 35 * len(audit)))

        with tab_types:
            profile = result.get("profile")
            if profile:
                table = profile_table(profile)
                counts = table["Detected Type"].value_counts()
                chips = " ".join(
                    f'<span class="adp-badge adp-badge-navy">{esc(k)}: {v}</span>'
                    for k, v in counts.items()
                )
                st.markdown(f'<div style="margin-bottom:10px;">{chips}</div>',
                            unsafe_allow_html=True)
                render_table(table, height=min(430, 48 + 35 * len(table)))

        with tab_num:
            stats = result.get("numeric_stats")
            if stats is None or stats.empty:
                st.info("No numeric columns were detected in this file.")
            else:
                shown = stats.copy()
                for col in ("Sum", "Mean", "Median", "Min", "Max", "Std Dev"):
                    shown[col] = shown[col].map(lambda v: fmt_money(v, 2))
                render_table(shown, height=min(430, 48 + 35 * len(shown)))
                dates = result.get("datetime_stats")
                if dates is not None and not dates.empty:
                    st.markdown('<div class="adp-section-h" style="margin-top:14px;">'
                                'Datetime Coverage</div>', unsafe_allow_html=True)
                    render_table(dates)

        with tab_cat:
            cats = result.get("categorical_stats")
            if cats is None or cats.empty:
                st.info("No categorical or text columns were detected in this file.")
            else:
                render_table(cats, height=min(430, 48 + 35 * len(cats)))
                profile = result.get("profile") or {}
                targets = [c for c in cats["Column"].tolist()
                           if profile.get(c, {}).get("type") == TYPE_CATEGORICAL][:2]
                if targets:
                    chart_cols = st.columns(len(targets), gap="large")
                    for col_box, column in zip(chart_cols, targets):
                        with col_box:
                            st.markdown(f'<div class="adp-section-h">Distribution: '
                                        f'{esc(column)}</div>', unsafe_allow_html=True)
                            top = profile[column].get("top_values", [])[:8]
                            if top:
                                series = pd.Series({k: v for k, v in top})
                                render_bar_chart(series, height=260)

# --------------------------------------------------------------------------- #
#  PANEL 3  —  REVIEW CLEANED DATA  +  DOWNLOAD OPTIMISED REPORT
# --------------------------------------------------------------------------- #

def panel_review_and_download() -> None:
    clean = st.session_state["clean_df"]
    result = st.session_state["result"]

    with st.container(border=True):
        panel_title("2. Review Cleaned Data (Preview)",
                    "Flagged rows are tinted coral and tagged with a pill per "
                    "validation flag; clean rows carry an emerald CLEAN badge.")
        if clean is None:
            st.info("Run the processing engine to populate the cleaned-data preview.")
        else:
            preview_cols = [c for c in (result or {}).get("preview_columns", [])
                            if c in clean.columns]
            # The categorical column to filter on is whichever one the file
            # actually has - Status in a ledger, or the first detected
            # categorical column in an arbitrary spreadsheet.
            profile = (result or {}).get("profile") or {}
            if "Status" in clean.columns:
                filter_col = "Status"
            else:
                filter_col = next(
                    (c for c in profile.get("__order__", [])
                     if profile.get(c, {}).get("type") == TYPE_CATEGORICAL
                     and c in clean.columns),
                    None,
                )

            f1, f2, f3, f4 = st.columns([2.2, 1.1, 1.1, 1.2], gap="medium")
            with f1:
                query = st.text_input(
                    "Search", placeholder="Search any column…",
                    label_visibility="collapsed", key="preview_search",
                )
            with f2:
                if filter_col:
                    present = set(clean[filter_col].astype(str))
                    ordered = [v for v in STATUS_ORDER if v in present]
                    ordered += sorted(present - set(STATUS_ORDER))
                    all_label = f"All {filter_col.replace('_', ' ').lower()}"
                    status_pick = st.selectbox(
                        filter_col, [all_label] + ordered,
                        label_visibility="collapsed", key="preview_status",
                    )
                else:
                    all_label, status_pick = "All rows", "All rows"
                    st.caption("No categorical column to filter on")
            with f3:
                view = st.selectbox("View", ["All rows", "Flagged only", "Clean only"],
                                    label_visibility="collapsed", key="preview_view")
            with f4:
                st.markdown(
                    f'<div style="padding-top:6px;">'
                    f'<span class="adp-badge adp-badge-navy">{len(clean):,} ROWS</span> '
                    f'<span class="adp-badge adp-badge-red">'
                    f'{int(clean["Flagged"].sum()):,} FLAGGED</span></div>',
                    unsafe_allow_html=True,
                )

            filtered = clean
            if filter_col and status_pick != all_label:
                filtered = filtered[filtered[filter_col].astype(str) == status_pick]
            if view == "Flagged only":
                filtered = filtered[filtered["Flagged"]]
            elif view == "Clean only":
                filtered = filtered[~filtered["Flagged"]]
            if query:
                needle = query.strip().lower()
                search_cols = [c for c in (preview_cols + ["Reference"])
                               if c in filtered.columns] or [
                    c for c in filtered.columns if c not in HELPER_COLUMNS]
                searchable = filtered[search_cols].astype(str)
                mask = searchable.apply(
                    lambda col: col.str.lower().str.contains(needle, na=False, regex=False)
                ).any(axis=1)
                filtered = filtered[mask]

            render_preview_table(filtered, columns=preview_cols)

            with st.expander("Open the full interactive grid (all columns, sortable)"):
                render_table(filtered.drop(columns=["Flag_Count", "Needs_Review"],
                                           errors="ignore"), height=430)

    with st.container(border=True):
        panel_title("3. Download Optimised Report",
                    "Formatted workbook and summary PDF, ready to send.")
        if clean is None or result is None:
            st.info("The download package is generated once the engine has run.")
            return

        source_name = st.session_state["raw_meta"]["name"]
        stem = re.sub(r"\.(xlsx|xls|csv)$", "", source_name, flags=re.I)
        stamp = f"{result['timestamp']:%Y%m%d_%H%M}"

        try:
            if result.get("mode") == "universal":
                excel_bytes = build_universal_excel(clean, result, source_name)
                pdf_bytes = build_universal_pdf(clean, result, source_name)
            else:
                excel_bytes = build_excel_report(clean, result, source_name)
                pdf_bytes = build_summary_pdf(clean, result, source_name)
        except Exception as exc:
            st.error(f"Report generation failed: {exc}")
            return

        d1, d2, d3 = st.columns([1.5, 1.2, 1.6], gap="medium")
        with d1:
            st.download_button(
                "Download Cleaned Excel File (.xlsx)",
                data=excel_bytes,
                file_name=f"{stem}_CLEANED_{stamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", use_container_width=True, key="dl_excel",
            )
        with d2:
            st.download_button(
                "Download Summary PDF",
                data=pdf_bytes,
                file_name=f"{stem}_SUMMARY_{stamp}.pdf",
                mime="application/pdf",
                use_container_width=True, key="dl_pdf",
            )
        with d3:
            if result.get("mode") == "universal":
                tab_one, tab_two = "Cleaned Dataset", "Executive Summary"
                tip = ("Cleaned Dataset - every source column, navy header band, "
                       "frozen panes, auto-filter, auto-fitted widths, flagged "
                       "rows tinted. | Executive Summary - dataset dimensions, "
                       "column data types, numeric and categorical statistics, "
                       "datetime coverage and the full data quality audit.")
            else:
                tab_one, tab_two = "Cleaned Data", "Executive Summary"
                tip = ("Cleaned Data - navy header band, frozen panes, "
                       "auto-filter, auto-fitted widths, flagged rows tinted. "
                       "| Executive Summary - revenue pivot by status with VAT "
                       "and inclusive totals, run metrics and the exception "
                       "breakdown.")
            st.markdown(
                f'<div class="adp-kv" style="padding-top:4px;">'
                f'<b>{fmt_size(len(excel_bytes))}</b> workbook &nbsp;·&nbsp; '
                f'<b>{fmt_size(len(pdf_bytes))}</b> PDF &nbsp;·&nbsp; '
                f'{len(clean):,} records<br>'
                f'<span class="adp-tip" title="{esc(tip)}">'
                f'Tabs: {tab_one} &amp; {tab_two}</span></div>',
                unsafe_allow_html=True,
            )


# --------------------------------------------------------------------------- #
#  PAGES
# --------------------------------------------------------------------------- #

def page_dashboard() -> None:
    panel_upload()
    panel_results()
    panel_profile_and_audit()
    panel_review_and_download()


def page_upload_data() -> None:
    panel_upload()
    raw = st.session_state["raw_df"]
    with st.container(border=True):
        panel_title("Raw File Inspection",
                    "Exactly what was read from disk, before any sanitisation.")
        if raw is None:
            st.info("Stage a file above to inspect its raw contents.")
            return
        mapped, mapping, extras = map_columns(raw)
        c1, c2 = st.columns([1, 1.4], gap="large")
        with c1:
            st.markdown('<div class="adp-section-h">Detected Schema</div>',
                        unsafe_allow_html=True)
            rows = []
            for canonical in ("Date", "Client_ID", "Reference", "Status", "Amount", "Notes"):
                rows.append({
                    "Canonical Column": canonical,
                    "Source Header": mapping.get(canonical, "— not found —"),
                    "Status": "Mapped" if canonical in mapping else "Missing",
                })
            render_table(pd.DataFrame(rows))
            if extras:
                st.caption(f"Unmapped columns carried through as EXTRA_*: "
                           f"{', '.join(str(c) for c in extras)}")
        with c2:
            st.markdown('<div class="adp-section-h">Raw Preview (first 200 rows)</div>',
                        unsafe_allow_html=True)
            render_table(raw.head(200), height=430, max_rows=200)


def page_rule_settings() -> None:
    with st.container(border=True):
        panel_title("Rule Settings",
                    "Configure the South African business-rules engine. Changes apply "
                    "to the next automation run.")
        c1, c2 = st.columns(2, gap="large")
        with c1:
            st.markdown('<div class="adp-section-h">Display & Currency</div>',
                        unsafe_allow_html=True)
            options_list = list(CURRENCY_OPTIONS.keys())
            current = st.session_state.get("currency", CURRENCY_DEFAULT)
            st.session_state["currency"] = st.selectbox(
                "Currency symbol used for every monetary figure on screen and "
                "in the exports",
                options_list,
                index=options_list.index(current) if current in options_list else 0,
            )
            st.session_state["vat_rate"] = st.number_input(
                "VAT rate (South African standard rate is 15%)",
                min_value=0.0, max_value=0.35,
                value=float(st.session_state["vat_rate"]),
                step=0.01, format="%.2f",
            )
            st.session_state["opt_exclude_cancelled"] = st.checkbox(
                "Exclude CANCELLED / reversed entries from revenue totals",
                value=st.session_state["opt_exclude_cancelled"],
            )
            st.session_state["opt_min_amount"] = st.number_input(
                "Flag transactions below this ZAR value (0 disables the check)",
                min_value=0.0, value=float(st.session_state["opt_min_amount"]), step=50.0,
            )
            st.session_state["opt_dayfirst"] = st.checkbox(
                "Parse ambiguous dates as day-first (SA convention: 09/03 = 9 March)",
                value=st.session_state["opt_dayfirst"],
            )
        with c2:
            st.markdown('<div class="adp-section-h">Validation & Reconciliation</div>',
                        unsafe_allow_html=True)
            st.session_state["opt_flag_missing_id"] = st.checkbox(
                "Flag rows with a missing Client ID (MISSING ID)",
                value=st.session_state["opt_flag_missing_id"])
            st.session_state["opt_flag_missing_status"] = st.checkbox(
                "Flag rows with a blank or unrecognised Status",
                value=st.session_state["opt_flag_missing_status"])
            st.session_state["opt_flag_bad_dates"] = st.checkbox(
                "Flag rows with an invalid or unparseable Date",
                value=st.session_state["opt_flag_bad_dates"])
            st.session_state["opt_flag_nonpositive"] = st.checkbox(
                "Flag zero, negative and below-threshold amounts",
                value=st.session_state["opt_flag_nonpositive"])
            st.session_state["opt_bank_recon"] = st.checkbox(
                "Run the bank reconciliation check (unassigned entries, discrepancies)",
                value=st.session_state["opt_bank_recon"])
            st.session_state["opt_drop_duplicates"] = st.checkbox(
                "Remove duplicate entries instead of flagging them",
                value=st.session_state["opt_drop_duplicates"])
            st.session_state["opt_flag_duplicates"] = st.checkbox(
                "Flag duplicate entries (ignored when removal is enabled)",
                value=st.session_state["opt_flag_duplicates"])

            st.markdown('<div class="adp-section-h" style="margin-top:14px;">'
                        'Universal Data Quality</div>', unsafe_allow_html=True)
            st.session_state["opt_flag_nulls"] = st.checkbox(
                "Flag rows containing null / empty values",
                value=st.session_state["opt_flag_nulls"])
            st.session_state["opt_flag_outliers"] = st.checkbox(
                "Flag severe numeric outliers",
                value=st.session_state["opt_flag_outliers"])
            st.session_state["opt_outlier_sigma"] = st.slider(
                "Outlier sensitivity (standard deviations from the mean)",
                min_value=1.5, max_value=6.0,
                value=float(st.session_state["opt_outlier_sigma"]), step=0.5)

    with st.container(border=True):
        panel_title("Active Rule Library",
                    f"{len(RULE_LIBRARY)} rules ship with the engine.")
        options = current_options()
        enabled_map = {
            "R04": options["flag_bad_dates"],
            "R06": options["flag_missing_id"],
            "R08": options["flag_missing_status"],
            "R10": options["flag_nonpositive"],
            "R12": options["bank_recon"] or options["flag_duplicates"] or options["drop_duplicates"],
        }
        table = pd.DataFrame([
            {"Code": code, "Rule": label,
             "State": "Enabled" if enabled_map.get(code, True) else "Disabled"}
            for code, label, _ in RULE_LIBRARY
        ])
        render_table(table, height=48 + 35 * len(table))


def page_analysis_universal(clean, result) -> None:
    """Analysis views for a file with an arbitrary schema."""
    profile = result.get("profile") or {}
    numeric = [c for c in profile.get("__order__", [])
               if profile[c]["type"] == TYPE_NUMERIC]
    categorical = [c for c in profile.get("__order__", [])
                   if profile[c]["type"] == TYPE_CATEGORICAL]
    datetimes = [c for c in profile.get("__order__", [])
                 if profile[c]["type"] == TYPE_DATETIME]

    with st.container(border=True):
        panel_title("Numeric Analysis",
                    f"{len(numeric)} numeric column(s) detected in "
                    f"{esc(st.session_state['raw_meta']['name'])}.")
        stats = result.get("numeric_stats")
        if stats is None or stats.empty:
            st.info("No numeric columns were detected in this file.")
        else:
            shown = stats.copy()
            for col in ("Sum", "Mean", "Median", "Min", "Max", "Std Dev"):
                shown[col] = shown[col].map(lambda v: fmt_money(v, 2))
            render_table(shown)
            totals = pd.Series({c: profile[c]["sum"] for c in numeric})
            if len(totals):
                st.markdown('<div class="adp-section-h">Column Totals</div>',
                            unsafe_allow_html=True)
                render_bar_chart(totals, height=max(200, 40 * len(totals)), money=True)

    if categorical:
        with st.container(border=True):
            panel_title("Categorical Distributions",
                        "Value counts for each detected categorical column.")
            pick = st.selectbox("Column", categorical, key="analysis_cat")
            top = profile[pick].get("top_values", [])
            if top:
                render_bar_chart(pd.Series({k: v for k, v in top}), height=300)
                breakdown = pd.DataFrame(top, columns=["Value", "Rows"])
                breakdown["Share %"] = (breakdown["Rows"] / max(len(clean), 1) * 100
                                        ).map(lambda v: f"{v:.1f}%")
                render_table(breakdown)

    if datetimes and numeric:
        with st.container(border=True):
            panel_title("Trend Over Time",
                        "Aggregated on a detected date column.")
            c1, c2 = st.columns(2, gap="large")
            with c1:
                date_col = st.selectbox("Date column", datetimes, key="analysis_date")
            with c2:
                value_col = st.selectbox("Value column", numeric, key="analysis_value")
            dated = clean[clean[date_col].astype(str).str.len() == 10]
            if dated.empty:
                st.warning("No parseable dates in that column.")
            else:
                trend = (dated.groupby(date_col)[value_col].sum()
                         .rename(f"{value_col} total"))
                render_line_chart(trend, height=300)


def page_analysis() -> None:
    clean = st.session_state["clean_df"]
    result = st.session_state["result"]
    if clean is None or result is None:
        with st.container(border=True):
            panel_title("Analysis")
            st.info("Run the processing engine on the Dashboard to unlock the "
                    "analysis views.")
        return

    # An arbitrary spreadsheet gets schema-driven analysis; a recognised sales
    # ledger keeps the revenue / VAT / client views.
    if result.get("mode") != "sales":
        page_analysis_universal(clean, result)
        return

    with st.container(border=True):
        panel_title("Revenue & VAT Analysis",
                    f"Based on {len(clean):,} processed records from "
                    f"{esc(st.session_state['raw_meta']['name'])}.")
        pivot = status_summary(clean)
        display = pivot.copy()
        for col in ("Revenue_Excl_VAT", "VAT_15", "Total_Inc_VAT"):
            display[col] = display[col].map(lambda v: fmt_money(v, 2))
        display = display.rename(columns={
            "Revenue_Excl_VAT": "Revenue (Excl VAT)",
            "VAT_15": f"VAT @ {result['vat_rate'] * 100:.0f}%",
            "Total_Inc_VAT": "Total (Incl VAT)",
            "Flagged_Rows": "Flagged Rows",
        })
        render_table(display)

        c1, c2 = st.columns(2, gap="large")
        with c1:
            st.markdown('<div class="adp-section-h">Revenue by Status</div>',
                        unsafe_allow_html=True)
            render_bar_chart(pivot.set_index("Status")["Revenue_Excl_VAT"],
                             height=300, money=True)
        with c2:
            st.markdown('<div class="adp-section-h">Transaction Count by Status</div>',
                        unsafe_allow_html=True)
            render_bar_chart(pivot.set_index("Status")["Transactions"], height=300)

    with st.container(border=True):
        panel_title("Daily Revenue Trend",
                    "Aggregated on the standardised YYYY-MM-DD date column.")
        dated = clean[clean["Date"].astype(str).str.len() == 10].copy()
        if dated.empty:
            st.warning("No parseable dates in this file, so no trend can be plotted.")
        else:
            trend = (dated.groupby("Date", as_index=True)["Amount"].sum()
                     .rename("Revenue (Excl VAT)"))
            render_line_chart(trend, height=300)

    with st.container(border=True):
        panel_title("Top Clients by Revenue")
        clients = clean[clean["Client_ID"].astype(str) != ""]
        if clients.empty:
            st.warning("No identifiable clients in this file.")
        else:
            top = (clients.groupby("Client_ID")
                   .agg(Transactions=("Amount", "size"),
                        Revenue=("Amount", "sum"),
                        VAT=("VAT_15", "sum"),
                        Flagged=("Flagged", "sum"))
                   .sort_values("Revenue", ascending=False).head(15).reset_index())
            top["Revenue"] = top["Revenue"].map(lambda v: fmt_money(v, 2))
            top["VAT"] = top["VAT"].map(lambda v: fmt_money(v, 2))
            top["Flagged"] = top["Flagged"].astype(int)
            render_table(top)

def page_history() -> None:
    with st.container(border=True):
        panel_title("Automation History",
                    "Every run completed in this session, newest first.")
        history = st.session_state["history"]
        if not history:
            st.info("No automation runs recorded yet in this session.")
            return
        frame = pd.DataFrame(history)
        display = frame.copy()
        display["Numeric Total"] = display["Numeric Total"].map(lambda v: fmt_money(v, 2))
        display["Health %"] = display["Health %"].map(lambda v: f"{v:.1f}%")
        render_table(display)

        c1, c2, c3, c4 = st.columns(4, gap="medium")
        with c1:
            st.markdown(metric_card("Runs This Session", f"{len(frame):,}"),
                        unsafe_allow_html=True)
        with c2:
            st.markdown(metric_card("Rows Processed",
                                    f"{int(frame['Rows Out'].sum()):,}"),
                        unsafe_allow_html=True)
        with c3:
            st.markdown(metric_card("Errors Caught", f"{int(frame['Flagged'].sum()):,}",
                                    tone="red"), unsafe_allow_html=True)
        with c4:
            st.markdown(metric_card("Avg. Health Score",
                                    f"{frame['Health %'].mean():.1f}%",
                                    f"Avg. runtime {frame['Seconds'].mean():.2f}s",
                                    tone="green"), unsafe_allow_html=True)

        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
        csv = frame.to_csv(index=False).encode("utf-8")
        st.download_button("Export run history (.csv)", data=csv,
                           file_name="auto_data_pro_run_history.csv", mime="text/csv",
                           key="dl_history")


def page_support() -> None:
    with st.container(border=True):
        panel_title("Support & Documentation",
                    "Everything needed to operate Auto-Data Pro on a local machine.")
        tab1, tab2, tab3 = st.tabs(["Quick start", "Input requirements", "Diagnostics"])

        with tab1:
            st.markdown(
                """
**1.** Double-click `run_app.bat` (or run `streamlit run app.py`). The dashboard
opens at `http://localhost:8501`.

**2.** On the **Dashboard** page, drag the raw monthly export onto the uploader.
`.xlsx`, `.xls` and `.csv` are all accepted. No file at hand? Press
**Load sample M&M Retail dataset** for a realistic, deliberately messy demo file.

**3.** Adjust anything you need under **Rule Settings**, then press
**PROCESS DATA FILE**.

**4.** Review the flagged rows in the preview (light red rows need attention),
then download the formatted workbook and the summary PDF.
                """
            )
        with tab2:
            st.markdown(
                """
Column headers are detected automatically — the raw file does not need to match
a fixed template. The engine looks for these concepts:

| Canonical column | Recognised headers (examples) |
| --- | --- |
| `Date` | Date, Invoice Date, Transaction Date, Posting Date |
| `Client_ID` | Client ID, Customer ID, Account No, Debtor Code |
| `Status` | Status, Payment Status, Invoice Status |
| `Amount` | Amount, Amount (ZAR), Value, Total, Revenue, Net |
| `Reference` | Reference, Bank Ref, Invoice No, EFT Ref |
| `Notes` | Notes, Comments, Remarks, Memo |

Anything that is not recognised is preserved in the output with an `EXTRA_` prefix,
so nothing from the source file is lost.
                """
            )
        with tab3:
            meta = st.session_state["raw_meta"]
            result = st.session_state["result"]
            st.markdown(
                f"""
<div class="adp-kv">
<b>Application</b> — {BRAND_NAME} v{VERSION}<br>
<b>Runtime</b> — Streamlit {st.__version__} · pandas {pd.__version__}<br>
<b>Host</b> — local desktop, http://localhost:8501<br>
<b>VAT profile</b> — South Africa, {st.session_state['vat_rate'] * 100:.0f}%<br>
<b>Rules loaded</b> — {len(RULE_LIBRARY)}<br>
<b>Table / chart engine</b> — {'pyarrow (native Streamlit grid)' if ARROW_AVAILABLE else 'HTML fallback - pyarrow unavailable'}<br>
<b>Staged file</b> — {esc(meta['name']) if meta else 'none'}<br>
<b>Last run</b> — {f"{result['timestamp']:%Y-%m-%d %H:%M:%S}" if result else 'never'}<br>
<b>Runs this session</b> — {len(st.session_state['history'])}
</div>
                """,
                unsafe_allow_html=True,
            )
            st.markdown("---")
            st.markdown(
                """
**Common issues**

- *"Could not read file"* — the export is probably an HTML or PDF file renamed to
  `.xlsx`. Re-export it from the source system as a genuine workbook or CSV.
- *Dates come through blank* — the source column holds free text. Rows are flagged
  `INVALID DATE` rather than dropped, so nothing is silently lost.
- *Amounts read as zero* — values such as `R 1 234,56` are handled, but a column
  mixing text and numbers may need cleaning at source. Affected rows are flagged
  `INVALID AMOUNT`.
                """
            )
            if not ARROW_AVAILABLE:
                st.warning(
                    "**pyarrow is unavailable on this machine**, so tables and "
                    "charts are drawn with the built-in HTML renderer instead of "
                    "Streamlit's native grid. Every feature still works and the "
                    "Excel and PDF exports are unaffected - only sorting inside "
                    "the on-screen grid is lost. This is typically a Windows "
                    "Application Control (WDAC) or Smart App Control policy "
                    "blocking the pyarrow DLL: ask IT to allow it, or reinstall "
                    "with `pip install --force-reinstall pyarrow`."
                )
                st.caption(f"Reported error: {ARROW_ERROR or 'pyarrow not installed'}")
            st.caption("Support: support@autodatapro.co.za  ·  Client account: M&M Retail")


# --------------------------------------------------------------------------- #
#  ENTRY POINT
# --------------------------------------------------------------------------- #

PAGES = {
    "Dashboard": page_dashboard,
    "Upload Data": page_upload_data,
    "Rule Settings": page_rule_settings,
    "Analysis": page_analysis,
    "History": page_history,
    "Support": page_support,
}


def main() -> None:
    render_header()
    choice = render_sidebar()
    PAGES.get(choice, page_dashboard)()


if __name__ == "__main__":
    main()
