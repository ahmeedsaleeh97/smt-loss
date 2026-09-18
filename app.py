"""
SMT Mounter Micro-Loss Analyzer & Task Force Action Tracker — Ver 2.0
==============================================================================
Production-ready single-file Streamlit application.

WHAT'S NEW IN VER 2.0 (on top of every Ver 1.01 feature)
----------------------------------------------------------
1. Local database persistence (SQLite, stdlib `sqlite3` — no server to install).
   Daily per-line metrics and the day's "Hot Issues" snapshot are saved with a
   single button, keyed by a Date Picker in the sidebar. A Historical Trends
   tab lets you pick any date range and review efficiency / loss trends.
2. A full Action Tracker: log a corrective action straight from a Hot Issue,
   manage its lifecycle (Open -> In Progress -> Delayed -> Closed), and:
     - Closing an action is BLOCKED unless a Before AND an After image are
       attached (freshly uploaded or already stored).
     - Moving to "Delayed" is BLOCKED unless a delay justification reason and
       a new Target Date are provided.
     - An Effectiveness Engine checks whether the same stoppage category
       reappears as a Hot Issue on the same line within 7 days of closing —
       flagging the action "Effective" or "Ineffective" automatically (based
       on the daily Hot-Issue snapshots saved to the database).
3. One-click "Generate PPT Report" (python-pptx): a title/summary slide for
   the whole plant, plus one slide per BLOCK4 line with its worst 3 Hot
   Issues, root causes, corrective actions, and Before/After photos pulled
   straight from the Action Tracker database.

NOTE ON THE DATABASE
---------------------
This uses a local SQLite file (see DB_PATH below) — perfectly fine for a
single-machine / local-server deployment. If you deploy to a host with an
ephemeral filesystem (e.g. some cloud PaaS free tiers), point DB_PATH at a
persistent volume, or swap `get_connection()` for a SQLAlchemy engine talking
to your department's PostgreSQL/MySQL server — every function in the
"DATABASE LAYER" section below takes a connection object, so that swap only
touches one function.

Core Methodology (unchanged)
-----------------------------
    Small Loss (min) = Loading Time (min) - [Actual Working Time (min) + Major Loss (>5m) (min)]

Run with:
    streamlit run app.py
"""

import csv
import io
import re
import sqlite3
from datetime import datetime, date, timedelta

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

APP_VERSION = "Ver 2.0"
LINE_NAMES = ["BLOCK4 01", "BLOCK4 02", "BLOCK4 03", "BLOCK4 04"]
DB_PATH = "smt_analyzer.db"
ACTION_STATUSES = ["Open", "In Progress", "Delayed", "Closed"]
EFFECTIVENESS_WINDOW_DAYS = 7

# ----------------------------------------------------------------------------
# PAGE CONFIG
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title=f"SMT Micro-Loss Analyzer & Action Tracker — {APP_VERSION}",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------------------------------------------------------
# DARK-MODE INDUSTRIAL CSS
# ----------------------------------------------------------------------------
CUSTOM_CSS = """
<style>
:root{
    --bg-0:#0b0f14; --bg-1:#111823; --bg-2:#161f2c; --bg-3:#1c2734; --line:#25313f;
    --accent:#00e0c6; --accent-2:#3ea6ff; --warn:#ffb547; --danger:#ff5c5c; --good:#3ddc84;
    --text-0:#eef3f8; --text-1:#a9b7c6; --text-2:#71828f;
}
html, body, [class*="css"]  { font-family: 'Segoe UI', 'Inter', system-ui, sans-serif; }
.stApp { background: radial-gradient(circle at 15% 0%, #0e1621 0%, var(--bg-0) 55%); color: var(--text-0); }
section[data-testid="stSidebar"] { background: linear-gradient(180deg, var(--bg-1) 0%, var(--bg-0) 100%); border-right: 1px solid var(--line); }
h1, h2, h3, h4 { color: var(--text-0) !important; font-weight: 700; letter-spacing: 0.3px; }

.app-header {
    padding: 18px 26px; border-radius: 14px;
    background: linear-gradient(120deg, rgba(0,224,198,0.10), rgba(62,166,255,0.06));
    border: 1px solid var(--line); margin-bottom: 22px;
    display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px;
}
.app-header h1{ margin:0; font-size: 1.6rem; }
.app-header p{ margin: 4px 0 0 0; color: var(--text-1); font-size: 0.92rem; }
.version-badge{
    display:inline-block; padding: 5px 14px; border-radius: 999px;
    background: linear-gradient(120deg, rgba(0,224,198,0.22), rgba(62,166,255,0.18));
    border: 1px solid var(--accent); color: var(--accent); font-weight:800;
    font-size: 0.82rem; letter-spacing: 0.6px; white-space:nowrap;
}
.sidebar-version{
    display:inline-block; padding: 3px 10px; border-radius: 999px;
    background: rgba(0,224,198,0.12); border: 1px solid var(--accent);
    color: var(--accent); font-weight:700; font-size: 0.72rem; margin-bottom: 10px;
}

.kpi-card {
    background: linear-gradient(160deg, var(--bg-2), var(--bg-1));
    border: 1px solid var(--line); border-radius: 16px; padding: 18px 20px;
    box-shadow: 0 6px 18px rgba(0,0,0,0.35); position: relative; overflow: hidden; height: 100%;
}
.kpi-card::before{ content:""; position:absolute; top:0; left:0; right:0; height:3px; background: var(--accent-bar, var(--accent)); }
.kpi-label{ color: var(--text-2); font-size: 0.78rem; text-transform: uppercase; letter-spacing: 1.1px; margin-bottom: 6px; }
.kpi-value{ font-size: 1.85rem; font-weight: 800; color: var(--text-0); line-height:1.1; }
.kpi-sub{ margin-top: 6px; font-size: 0.82rem; color: var(--text-1); }
.kpi-good{ --accent-bar: var(--good); }
.kpi-warn{ --accent-bar: var(--warn); }
.kpi-bad{ --accent-bar: var(--danger); }
.kpi-info{ --accent-bar: var(--accent-2); }

.section-card{ background: var(--bg-1); border: 1px solid var(--line); border-radius: 16px; padding: 18px 22px; margin-bottom: 20px; }
.section-title{ font-size: 1.05rem; font-weight: 700; color: var(--text-0); margin-bottom: 4px; }
.section-caption{ color: var(--text-2); font-size: 0.83rem; margin-bottom: 14px; }

.action-card{ border-radius: 14px; padding: 16px 18px; margin-bottom: 12px; border-left: 4px solid var(--accent); background: var(--bg-2); }
.action-priority-1{ border-left-color: var(--danger); }
.action-priority-2{ border-left-color: var(--warn); }
.action-priority-3{ border-left-color: var(--accent-2); }
.action-title{ font-weight:700; font-size: 0.98rem; color: var(--text-0); }
.action-meta{ color: var(--text-2); font-size: 0.78rem; margin-top:2px;}
.action-body{ color: var(--text-1); font-size: 0.88rem; margin-top:8px; line-height:1.5;}

.badge{ display:inline-block; padding: 2px 10px; border-radius: 999px; font-size: 0.72rem; font-weight:700; letter-spacing:0.4px; margin-right:6px; }
.badge-danger{ background: rgba(255,92,92,0.15); color: var(--danger); border:1px solid rgba(255,92,92,0.4);}
.badge-warn{ background: rgba(255,181,71,0.15); color: var(--warn); border:1px solid rgba(255,181,71,0.4);}
.badge-info{ background: rgba(62,166,255,0.15); color: var(--accent-2); border:1px solid rgba(62,166,255,0.4);}
.badge-good{ background: rgba(61,220,132,0.15); color: var(--good); border:1px solid rgba(61,220,132,0.4);}
.badge-muted{ background: rgba(113,130,143,0.15); color: var(--text-2); border:1px solid rgba(113,130,143,0.4);}

hr.divider{ border: none; border-top: 1px solid var(--line); margin: 22px 0; }
.stDataFrame, .stTable { border-radius: 12px; overflow:hidden; }
.small-note{ color: var(--text-2); font-size: 0.78rem; }
.line-pill{
    display:inline-block; padding: 3px 12px; border-radius: 999px;
    background: rgba(62,166,255,0.14); border: 1px solid rgba(62,166,255,0.4);
    color: var(--accent-2); font-weight:700; font-size: 0.78rem; margin: 2px 4px 2px 0;
}
.hot-card{ background: var(--bg-2); border: 1px solid var(--line); border-radius: 14px; padding: 14px 16px; height:100%; }
.hot-rank{ color: var(--text-2); font-size: 0.72rem; text-transform:uppercase; letter-spacing:0.8px; }
.hot-name{ color: var(--text-0); font-weight:800; font-size: 1.02rem; margin-top:2px; }
.hot-metric{ color: var(--accent-2); font-size: 0.88rem; margin-top:6px; font-weight:700; }
.hot-cause{ color: var(--text-1); font-size: 0.80rem; margin-top:4px; }

.tracker-card{
    background: var(--bg-2); border: 1px solid var(--line); border-radius: 14px;
    padding: 16px 18px; margin-bottom: 14px;
}
.tracker-title{ font-weight:800; font-size: 1.0rem; color: var(--text-0); }
.tracker-meta{ color: var(--text-2); font-size: 0.78rem; margin-top: 2px; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# CONSTANTS & PLAYBOOK
# ----------------------------------------------------------------------------
MAJOR_LOSS_THRESHOLD_MIN = 5.0
UPLOAD_TYPES = ["csv", "xlsx", "xls", "txt"]
IMAGE_TYPES = ["png", "jpg", "jpeg"]

STOPPAGE_PLAYBOOK = {
    "wait previous": {
        "label": "Wait Previous", "root_cause": "Upstream Bottleneck",
        "action": (
            "Line is starving because the upstream process/machine cannot keep pace. "
            "Balance cycle times across the line, check upstream OEE, and consider a small "
            "buffer (WIP) between this mounter and the upstream station."
        ),
    },
    "part supply": {
        "label": "Part Supply", "root_cause": "Material Delay / Feeder Replenishment",
        "action": (
            "Feeders are running out or component kitting/replenishment is late. Implement "
            "a kanban-based feeder refill schedule and pre-stage next-lot reels at the line side."
        ),
    },
    "wait next": {
        "label": "Wait Next", "root_cause": "Downstream Blocking",
        "action": (
            "Machine is blocked because the downstream process cannot accept boards fast "
            "enough. Check downstream cycle time and buffer conveyor capacity."
        ),
    },
    "machine error": {
        "label": "Machine Error", "root_cause": "Equipment Reliability / Quality Fault",
        "action": (
            "Recurring machine faults (recognition error, nozzle/head fault, pick-up miss). "
            "Pull the fault-code frequency log and prioritize top offending heads/nozzles."
        ),
    },
    "operator downtime": {
        "label": "Operator Downtime", "root_cause": "Manning / Operator Response",
        "action": (
            "Losses linked to operator response time or manning gaps. Review manning plan "
            "and response-time SLA for alarms."
        ),
    },
    "maintenance": {
        "label": "Maintenance", "root_cause": "Unplanned/Reactive Maintenance",
        "action": (
            "Time lost to unplanned maintenance intervention. Review PM compliance and "
            "interval, and shift toward condition-based maintenance."
        ),
    },
}
DEFAULT_PLAYBOOK_ENTRY = {
    "label": "Other / Unclassified", "root_cause": "Unclassified Stoppage",
    "action": "Review the raw SW log entries for this status and classify with process engineering.",
}


def get_playbook_entry(status_text):
    key = str(status_text).strip().lower()
    for pk, pv in STOPPAGE_PLAYBOOK.items():
        if pk in key:
            return pv
    return DEFAULT_PLAYBOOK_ENTRY


# ----------------------------------------------------------------------------
# GENERIC HELPERS
# ----------------------------------------------------------------------------
def fmt_min(m):
    if m is None or (isinstance(m, float) and np.isnan(m)):
        return "—"
    return f"{m:,.1f} min"


def fmt_pct(p):
    if p is None or (isinstance(p, float) and np.isnan(p)):
        return "—"
    return f"{p:,.1f}%"


def safe_div(a, b):
    try:
        if b in (0, None) or (isinstance(b, float) and np.isnan(b)):
            return np.nan
        return a / b
    except Exception:
        return np.nan


def find_column(df, must_contain_all=None, must_contain_any=None):
    if df is None or df.empty:
        return None
    for c in df.columns:
        name = str(c).strip().lower()
        ok_all = all(k in name for k in must_contain_all) if must_contain_all else True
        ok_any = any(k in name for k in must_contain_any) if must_contain_any else True
        if ok_all and ok_any:
            return c
    return None


def col_by_letter(df, letter):
    idx = 0
    for ch in letter.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    idx -= 1
    if df is not None and 0 <= idx < len(df.columns):
        return df.columns[idx]
    return None


def parse_hms_to_seconds(value):
    """Robustly parse a duration value into seconds: 'HH:MM:SS', 'MM:SS', Timedelta,
    datetime.time, numeric seconds/minutes, or NaN."""
    if value is None:
        return np.nan
    if isinstance(value, (int, float, np.integer, np.floating)):
        if np.isnan(value):
            return np.nan
        return float(value)
    if isinstance(value, pd.Timedelta):
        return value.total_seconds()
    if isinstance(value, timedelta):
        return value.total_seconds()
    if hasattr(value, "hour") and hasattr(value, "minute") and hasattr(value, "second"):
        try:
            return value.hour * 3600 + value.minute * 60 + value.second
        except Exception:
            pass
    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none", "nat"):
        return np.nan
    match = re.match(r"^(\d+):(\d{1,2}):(\d{1,2}(?:\.\d+)?)$", s)
    if match:
        h, m, sec = match.groups()
        return int(h) * 3600 + int(m) * 60 + float(sec)
    match = re.match(r"^(\d{1,2}):(\d{1,2}(?:\.\d+)?)$", s)
    if match:
        m, sec = match.groups()
        return int(m) * 60 + float(sec)
    try:
        return float(s)
    except ValueError:
        return np.nan


def seconds_to_hms(total_seconds):
    if total_seconds is None or (isinstance(total_seconds, float) and np.isnan(total_seconds)):
        return "—"
    total_seconds = int(round(total_seconds))
    h, rem = divmod(total_seconds, 3600)
    m, s = divmod(rem, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


# ----------------------------------------------------------------------------
# MULTI-FORMAT FILE READING  (.xlsx / .csv / .txt with delimiter auto-detect)
# ----------------------------------------------------------------------------
def _decode_bytes(raw_bytes):
    for enc in ("utf-8-sig", "utf-8", "latin1", "cp1252"):
        try:
            return raw_bytes.decode(enc)
        except Exception:
            continue
    raise ValueError("Could not decode the file using utf-8, utf-8-sig, latin1, or cp1252 encoding.")


def _sniff_delimiter(text):
    sample = "\n".join(text.splitlines()[:25])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=[",", ";", "\t", "|"])
        return dialect.delimiter
    except Exception:
        header_line = text.splitlines()[0] if text.splitlines() else ""
        candidates = [",", ";", "\t", "|"]
        counts = {d: header_line.count(d) for d in candidates}
        best = max(counts, key=counts.get)
        return best if counts[best] > 0 else ","


def read_delimited_text(uploaded_file):
    uploaded_file.seek(0)
    raw_bytes = uploaded_file.read()
    text = _decode_bytes(raw_bytes)
    delimiter = _sniff_delimiter(text)
    df = pd.read_csv(io.StringIO(text), sep=delimiter, engine="python")
    if df.shape[1] <= 1:
        for d in [",", ";", "\t", "|"]:
            if d == delimiter:
                continue
            try:
                trial = pd.read_csv(io.StringIO(text), sep=d, engine="python")
                if trial.shape[1] > df.shape[1]:
                    df = trial
                    delimiter = d
            except Exception:
                continue
    return df, delimiter


def read_all_sheets(uploaded_file):
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")


def load_uploaded_table(uploaded_file):
    name = uploaded_file.name.lower()
    if name.endswith((".xlsx", ".xls")):
        sheets = read_all_sheets(uploaded_file)
        return sheets, "sheets", f"Excel workbook with {len(sheets)} sheet(s)"
    elif name.endswith((".csv", ".txt")):
        df, delimiter = read_delimited_text(uploaded_file)
        delim_name = {",": "comma", ";": "semicolon", "\t": "tab", "|": "pipe"}.get(delimiter, delimiter)
        return df, "table", f"Delimited text file (detected {delim_name}-separated)"
    else:
        raise ValueError(f"Unsupported file type: {uploaded_file.name}")


# ----------------------------------------------------------------------------
# BLOCK4 LINE DETECTION
# ----------------------------------------------------------------------------
BLOCK_PATTERNS = [
    r"block\s*4\s*[-_ ]?\s*0?(\d{1,2})",
    r"\bb4[-_ ]?0?(\d{1,2})\b",
    r"\bblock\s*[-_ ]?0?(\d{1,2})\b",
    r"\bline\s*[-_ ]?0?(\d{1,2})\b",
    r"\bL[-_ ]?0?(\d{1,2})\b",
]


def normalize_to_block(text):
    if text is None:
        return None
    s = str(text).strip()
    if s == "" or s.lower() in ("nan", "none", "nat"):
        return None
    for pat in BLOCK_PATTERNS:
        m = re.search(pat, s, re.IGNORECASE)
        if m:
            try:
                num = int(m.group(1))
            except ValueError:
                continue
            if 1 <= num <= len(LINE_NAMES):
                return f"BLOCK4 0{num}"
    return None


def detect_block_column(df):
    col = find_column(df, must_contain_any=["line", "block"])
    if col is not None:
        return "direct", col
    for kw in ["machine", "equipment", "model", "mounter", "station", "eqp"]:
        c = find_column(df, must_contain_any=[kw])
        if c is not None:
            sample = df[c].dropna().astype(str).head(150)
            if len(sample) > 0 and sample.apply(normalize_to_block).notna().sum() > 0:
                return "extract", c
    for c in df.columns:
        if df[c].dtype == object:
            sample = df[c].dropna().astype(str).head(150)
            if len(sample) == 0:
                continue
            hit_rate = sample.apply(normalize_to_block).notna().sum() / len(sample)
            if hit_rate >= 0.25:
                return "extract", c
    return None, None


def build_block_series(df):
    mode, col = detect_block_column(df)
    if mode is None:
        return pd.Series([None] * len(df), index=df.index), None
    return df[col].astype(str).apply(normalize_to_block), col


# ----------------------------------------------------------------------------
# MES MAJOR LOSS FILE PARSER
# ----------------------------------------------------------------------------
def parse_major_loss_file(uploaded_file, manual_unit_hint="min"):
    warnings, errors = [], []
    try:
        payload, kind, meta = load_uploaded_table(uploaded_file)
    except Exception as e:
        return {"ok": False, "df": None, "total_minutes": np.nan,
                "warnings": warnings, "errors": [f"Could not read the Major Loss file: {e}"]}

    if kind == "sheets":
        sheet_name = None
        for n in payload:
            if "loss" in str(n).lower():
                sheet_name = n
                break
        if sheet_name is None:
            sheet_name = list(payload.keys())[0]
        if len(payload) > 1:
            warnings.append(f"Major Loss workbook has multiple sheets — using sheet '{sheet_name}'.")
        raw = payload[sheet_name]
    else:
        raw = payload

    if raw is None or raw.empty:
        return {"ok": False, "df": None, "total_minutes": np.nan,
                "warnings": warnings, "errors": ["The Major Loss file appears to be empty."]}

    min_col = find_column(raw, must_contain_all=["loss"], must_contain_any=["min"])
    sec_col = find_column(raw, must_contain_all=["loss"], must_contain_any=["sec"])
    duration_col, unit = None, None
    if min_col is not None:
        duration_col, unit = min_col, "min"
    elif sec_col is not None:
        duration_col, unit = sec_col, "sec"
    else:
        k_col = col_by_letter(raw, "K")
        l_col = col_by_letter(raw, "L")
        if k_col is not None and raw[k_col].apply(parse_hms_to_seconds).notna().sum() > 0:
            duration_col, unit = k_col, manual_unit_hint
            warnings.append(f"Loss duration column not found by header — using column K ('{k_col}') by position, assumed unit = {unit}.")
        elif l_col is not None and raw[l_col].apply(parse_hms_to_seconds).notna().sum() > 0:
            duration_col, unit = l_col, manual_unit_hint
            warnings.append(f"Loss duration column not found by header — using column L ('{l_col}') by position, assumed unit = {unit}.")

    if duration_col is None:
        return {"ok": False, "df": None, "total_minutes": np.nan, "warnings": warnings,
                "errors": ["Could not find a 'Loss Time (Min/Sec)' column (expected around column K or L)."]}

    remarks_col = find_column(raw, must_contain_any=["remark", "categor", "reason"])
    if remarks_col is None:
        remarks_col = col_by_letter(raw, "M")
        if remarks_col is not None:
            warnings.append(f"Remarks column not found by header — using column M ('{remarks_col}') by position.")

    line_series, line_col = build_block_series(raw)
    matched_count = line_series.notna().sum()
    if line_col is not None:
        warnings.append(f"Line identifiers scanned in column '{line_col}' — matched {matched_count} of {len(raw)} row(s) to a BLOCK4 line.")
    else:
        warnings.append("Could not find a column with BLOCK4/line identifiers in the Major Loss file — all rows are 'Unmatched'.")

    seconds_series = raw[duration_col].apply(parse_hms_to_seconds)
    n_bad = seconds_series.isna().sum()
    if n_bad > 0:
        warnings.append(f"{n_bad} row(s) in the loss duration column could not be parsed and were ignored.")

    minutes_series = seconds_series / 60.0 if unit == "sec" else seconds_series
    clean = pd.DataFrame({
        "Duration_Min": minutes_series,
        "Remarks": raw[remarks_col].astype(str).str.strip() if remarks_col is not None else "Unclassified",
        "Line": line_series,
    })
    clean = clean.dropna(subset=["Duration_Min"])
    clean = clean[clean["Duration_Min"] >= 0]
    clean["Line"] = clean["Line"].fillna("Unmatched")

    if clean.empty:
        return {"ok": False, "df": None, "total_minutes": np.nan,
                "warnings": warnings, "errors": ["No valid loss-duration rows were found after parsing."]}

    below_threshold = (clean["Duration_Min"] < MAJOR_LOSS_THRESHOLD_MIN).sum()
    if below_threshold > 0:
        warnings.append(f"{below_threshold} row(s) are below the {MAJOR_LOSS_THRESHOLD_MIN:.0f}-minute threshold — still included as provided.")

    unmatched_min = float(clean.loc[clean["Line"] == "Unmatched", "Duration_Min"].sum())
    if unmatched_min > 0:
        warnings.append(f"{fmt_min(unmatched_min)} of Major Loss could not be matched to a BLOCK4 line and is excluded from totals.")

    total_minutes = float(clean.loc[clean["Line"].isin(LINE_NAMES), "Duration_Min"].sum())
    return {"ok": True, "df": clean, "total_minutes": total_minutes, "warnings": warnings, "errors": errors}


# ----------------------------------------------------------------------------
# MACHINE SW STOPPAGE LOG PARSER
# ----------------------------------------------------------------------------
def parse_sw_log_file(uploaded_file):
    warnings, errors = [], []
    try:
        payload, kind, meta = load_uploaded_table(uploaded_file)
    except Exception as e:
        return {"ok": False, "stoppage_df": None, "status_df": None, "module_df": None,
                "warnings": warnings, "errors": [f"Could not read the Machine SW Log file: {e}"]}

    sheets = payload if kind == "sheets" else {"Data": payload}

    event_sheet_name = None
    for n in sheets:
        ln = str(n).lower()
        if "stoppage" in ln and "status" not in ln and "module" not in ln:
            event_sheet_name = n
            break
    if event_sheet_name is None:
        for n, df_ in sheets.items():
            if find_column(df_, must_contain_any=["span", "duration"]) and find_column(df_, must_contain_any=["status"]):
                event_sheet_name = n
                break
    if event_sheet_name is None and len(sheets) == 1:
        event_sheet_name = list(sheets.keys())[0]

    stoppage_df = None
    if event_sheet_name is not None:
        raw1 = sheets[event_sheet_name]
        span_col = find_column(raw1, must_contain_any=["span", "duration"]) or col_by_letter(raw1, "C")
        machine_col = find_column(raw1, must_contain_any=["machine"]) or col_by_letter(raw1, "F")
        module_col = find_column(raw1, must_contain_any=["module"]) or col_by_letter(raw1, "G")
        status_col = find_column(raw1, must_contain_any=["status"]) or col_by_letter(raw1, "H")
        missing = [n for n, c in [("Span", span_col), ("Status", status_col)] if c is None]
        if missing:
            warnings.append(f"Sheet '{event_sheet_name}': could not locate column(s) {', '.join(missing)}.")
        if span_col is not None and status_col is not None:
            line_series, line_col = build_block_series(raw1)
            matched_count = line_series.notna().sum()
            if line_col is not None:
                warnings.append(f"Line identifiers scanned in column '{line_col}' of the SW Log — matched {matched_count} of {len(raw1)} row(s).")
            else:
                warnings.append("Could not find BLOCK4/line identifiers in the SW Log — all events are 'Unmatched'.")
            stoppage_df = pd.DataFrame({
                "Machine": raw1[machine_col].astype(str) if machine_col is not None else "N/A",
                "Module": raw1[module_col].astype(str) if module_col is not None else "N/A",
                "Status": raw1[status_col].astype(str).str.strip(),
                "Duration_Sec": raw1[span_col].apply(parse_hms_to_seconds),
                "Line": line_series,
            })
            stoppage_df = stoppage_df.dropna(subset=["Duration_Sec"])
            stoppage_df["Line"] = stoppage_df["Line"].fillna("Unmatched")
            if stoppage_df.empty:
                stoppage_df = None
                warnings.append(f"No valid duration rows found in sheet '{event_sheet_name}'.")
    else:
        warnings.append("Could not identify an event-level Stoppage sheet/table with Span & Status columns.")

    status_df, module_df = None, None
    if stoppage_df is not None and not stoppage_df.empty:
        status_df = (stoppage_df.groupby("Status")["Duration_Sec"]
                     .agg(Qty="size", Min_Sec="min", Max_Sec="max", Total_Sec="sum").reset_index())
        module_df = stoppage_df.groupby("Module")["Duration_Sec"].agg(Total_Sec="sum").reset_index()
    elif kind == "sheets":
        status_sheet = next((n for n in sheets if "status" in str(n).lower()), None)
        if status_sheet is not None:
            try:
                raw2 = sheets[status_sheet]
                status_col = find_column(raw2, must_contain_any=["status"]) or col_by_letter(raw2, "A")
                qty_col = find_column(raw2, must_contain_any=["qty", "count"]) or col_by_letter(raw2, "B")
                min_col_ = find_column(raw2, must_contain_any=["min"]) or col_by_letter(raw2, "C")
                max_col_ = find_column(raw2, must_contain_any=["max"]) or col_by_letter(raw2, "D")
                total_col = find_column(raw2, must_contain_any=["total"]) or col_by_letter(raw2, "E")
                if status_col is not None and total_col is not None:
                    status_df = pd.DataFrame({
                        "Status": raw2[status_col].astype(str).str.strip(),
                        "Qty": pd.to_numeric(raw2[qty_col], errors="coerce") if qty_col is not None else np.nan,
                        "Min_Sec": raw2[min_col_].apply(parse_hms_to_seconds) if min_col_ is not None else np.nan,
                        "Max_Sec": raw2[max_col_].apply(parse_hms_to_seconds) if max_col_ is not None else np.nan,
                        "Total_Sec": raw2[total_col].apply(parse_hms_to_seconds),
                    })
                    status_df = status_df.dropna(subset=["Total_Sec"])
                    status_df = status_df[status_df["Status"].str.lower() != "nan"]
                    warnings.append(f"Using pre-aggregated sheet '{status_sheet}' — per-line breakdown unavailable in this fallback mode.")
            except Exception as e:
                warnings.append(f"Could not parse fallback sheet '{status_sheet}': {e}")
        module_sheet = next((n for n in sheets if "module" in str(n).lower()), None)
        if module_sheet is not None:
            try:
                raw3 = sheets[module_sheet]
                mod_col = find_column(raw3, must_contain_any=["module"]) or col_by_letter(raw3, "A")
                tot_col = find_column(raw3, must_contain_any=["total"]) or raw3.columns[-1]
                if mod_col is not None and tot_col is not None:
                    module_df = pd.DataFrame({
                        "Module": raw3[mod_col].astype(str).str.strip(),
                        "Total_Sec": raw3[tot_col].apply(parse_hms_to_seconds),
                    })
                    module_df = module_df.dropna(subset=["Total_Sec"])
                    module_df = module_df[module_df["Module"].str.lower() != "nan"]
            except Exception as e:
                warnings.append(f"Could not parse fallback sheet '{module_sheet}': {e}")

    ok = status_df is not None and not status_df.empty
    if not ok:
        errors.append("Could not build a stoppage-by-status summary — need an event-level Stoppage table or a Stoppage(Status) sheet.")

    return {"ok": ok, "stoppage_df": stoppage_df, "status_df": status_df, "module_df": module_df,
            "warnings": warnings, "errors": errors}


def aggregate_status(stoppage_df):
    if stoppage_df is None or stoppage_df.empty:
        return None
    return (stoppage_df.groupby("Status")["Duration_Sec"]
            .agg(Qty="size", Min_Sec="min", Max_Sec="max", Total_Sec="sum").reset_index())


def aggregate_module(stoppage_df):
    if stoppage_df is None or stoppage_df.empty:
        return None
    return stoppage_df.groupby("Module")["Duration_Sec"].agg(Total_Sec="sum").reset_index()


def build_action_plan(status_df, top_n=5):
    if status_df is None or status_df.empty:
        return []
    ranked = status_df.sort_values("Total_Sec", ascending=False).reset_index(drop=True)
    total_all = ranked["Total_Sec"].sum()
    plan = []
    for i, row in ranked.head(top_n).iterrows():
        entry = get_playbook_entry(row["Status"])
        share = safe_div(row["Total_Sec"], total_all) * 100 if total_all else np.nan
        priority = 1 if i == 0 else (2 if i in (1, 2) else 3)
        plan.append({
            "rank": i + 1, "priority": priority, "status_raw": row["Status"],
            "label": entry["label"], "root_cause": entry["root_cause"], "action": entry["action"],
            "total_sec": row["Total_Sec"], "qty": row.get("Qty", np.nan), "share_pct": share,
        })
    return plan


# ============================================================================
# DATABASE LAYER  (SQLite — swap get_connection() for another engine if needed)
# ============================================================================
@st.cache_resource
def get_connection():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL;")
    return conn


def init_db():
    conn = get_connection()
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS daily_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            record_date TEXT NOT NULL,
            line TEXT NOT NULL,
            loading_min REAL, working_min REAL, major_loss_min REAL, small_loss_min REAL,
            efficiency_pct REAL, created_at TEXT,
            UNIQUE(record_date, line)
        );
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS daily_hot_issues (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            record_date TEXT NOT NULL, line TEXT NOT NULL, status TEXT NOT NULL,
            total_min REAL, rank INTEGER, created_at TEXT
        );
    """)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS action_items (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            line TEXT NOT NULL, root_cause TEXT, status_category TEXT,
            corrective_action TEXT, responsible_owner TEXT, target_date TEXT,
            action_status TEXT NOT NULL DEFAULT 'Open',
            delay_reason TEXT,
            before_image BLOB, before_image_name TEXT,
            after_image BLOB, after_image_name TEXT,
            closed_date TEXT, effectiveness TEXT DEFAULT 'NA',
            created_at TEXT, updated_at TEXT
        );
    """)
    conn.commit()


def save_daily_snapshot(record_date_str, per_line_metrics):
    """per_line_metrics: dict[line] -> {loading, working, major, small, eff}"""
    conn = get_connection()
    now = datetime.now().isoformat(timespec="seconds")
    for line, m in per_line_metrics.items():
        conn.execute("""
            INSERT INTO daily_records (record_date, line, loading_min, working_min, major_loss_min, small_loss_min, efficiency_pct, created_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(record_date, line) DO UPDATE SET
                loading_min=excluded.loading_min, working_min=excluded.working_min,
                major_loss_min=excluded.major_loss_min, small_loss_min=excluded.small_loss_min,
                efficiency_pct=excluded.efficiency_pct, created_at=excluded.created_at;
        """, (record_date_str, line, m["loading"], m["working"], m["major"], m["small"], m["eff"], now))
    conn.commit()


def save_hot_issues(record_date_str, line, status_df, top_n=5):
    if status_df is None or status_df.empty:
        return
    conn = get_connection()
    now = datetime.now().isoformat(timespec="seconds")
    conn.execute("DELETE FROM daily_hot_issues WHERE record_date = ? AND line = ?;", (record_date_str, line))
    ranked = status_df.sort_values("Total_Sec", ascending=False).head(top_n).reset_index(drop=True)
    for i, row in ranked.iterrows():
        conn.execute("""
            INSERT INTO daily_hot_issues (record_date, line, status, total_min, rank, created_at)
            VALUES (?, ?, ?, ?, ?, ?);
        """, (record_date_str, line, str(row["Status"]), float(row["Total_Sec"]) / 60.0, int(i + 1), now))
    conn.commit()


def load_daily_records(start_date_str, end_date_str, line=None):
    conn = get_connection()
    if line and line != "All Lines":
        q = "SELECT * FROM daily_records WHERE record_date BETWEEN ? AND ? AND line = ? ORDER BY record_date;"
        params = (start_date_str, end_date_str, line)
    else:
        q = "SELECT * FROM daily_records WHERE record_date BETWEEN ? AND ? ORDER BY record_date, line;"
        params = (start_date_str, end_date_str)
    return pd.read_sql_query(q, conn, params=params)


def add_action_item(line, root_cause, status_category, corrective_action, responsible_owner, target_date_str):
    conn = get_connection()
    now = datetime.now().isoformat(timespec="seconds")
    cur = conn.execute("""
        INSERT INTO action_items (line, root_cause, status_category, corrective_action, responsible_owner,
                                   target_date, action_status, effectiveness, created_at, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, 'Open', 'NA', ?, ?);
    """, (line, root_cause, status_category, corrective_action, responsible_owner, target_date_str, now, now))
    conn.commit()
    return cur.lastrowid


def update_action_item(item_id, fields):
    """fields: dict of column -> value to update."""
    if not fields:
        return
    conn = get_connection()
    fields = dict(fields)
    fields["updated_at"] = datetime.now().isoformat(timespec="seconds")
    set_clause = ", ".join(f"{k} = ?" for k in fields)
    params = list(fields.values()) + [item_id]
    conn.execute(f"UPDATE action_items SET {set_clause} WHERE id = ?;", params)
    conn.commit()


def get_action_items(line=None, status=None):
    conn = get_connection()
    q = "SELECT * FROM action_items WHERE 1=1"
    params = []
    if line and line != "All Lines":
        q += " AND line = ?"
        params.append(line)
    if status and status != "All":
        q += " AND action_status = ?"
        params.append(status)
    q += " ORDER BY CASE action_status WHEN 'Delayed' THEN 0 WHEN 'Open' THEN 1 WHEN 'In Progress' THEN 2 ELSE 3 END, target_date;"
    return pd.read_sql_query(q, conn, params=params)


def check_recurrence(line, status_category, closed_date_str, window_days=EFFECTIVENESS_WINDOW_DAYS):
    """Returns True if `status_category` shows up again in daily_hot_issues for `line`
    within `window_days` days AFTER closed_date_str."""
    conn = get_connection()
    closed_dt = datetime.strptime(closed_date_str, "%Y-%m-%d").date()
    window_end = closed_dt + timedelta(days=window_days)
    q = """
        SELECT COUNT(*) FROM daily_hot_issues
        WHERE line = ? AND status = ? AND record_date > ? AND record_date <= ?;
    """
    row = conn.execute(q, (line, status_category, closed_date_str, window_end.isoformat())).fetchone()
    return (row[0] or 0) > 0


def evaluate_and_persist_effectiveness(action_row, today):
    """Given a row (pandas Series/dict) from action_items with action_status == 'Closed',
    compute effectiveness and persist it if it has changed. Returns the effectiveness string."""
    if action_row["action_status"] != "Closed" or not action_row.get("closed_date"):
        return action_row.get("effectiveness") or "NA"

    current = action_row.get("effectiveness") or "Pending"
    if current in ("Effective", "Ineffective"):
        return current  # already finalized, don't recompute

    closed_dt = datetime.strptime(action_row["closed_date"], "%Y-%m-%d").date()
    recurred = check_recurrence(action_row["line"], action_row["status_category"], action_row["closed_date"])

    if recurred:
        new_status = "Ineffective"
    elif (today - closed_dt).days >= EFFECTIVENESS_WINDOW_DAYS:
        new_status = "Effective"
    else:
        new_status = "Pending"

    if new_status != current:
        update_action_item(int(action_row["id"]), {"effectiveness": new_status})
    return new_status


init_db()


# ----------------------------------------------------------------------------
# PPTX REPORT GENERATOR
# ----------------------------------------------------------------------------
def _pptx_add_title(slide, text, top=Inches(0.3)):
    box = slide.shapes.add_textbox(Inches(0.4), top, Inches(9.2), Inches(0.8))
    tf = box.text_frame
    tf.text = text
    tf.paragraphs[0].font.size = Pt(28)
    tf.paragraphs[0].font.bold = True
    tf.paragraphs[0].font.color.rgb = RGBColor(0x0B, 0x0F, 0x14)
    return box


def _pptx_add_kpi_row(slide, kpis, top=Inches(1.2)):
    """kpis: list of (label, value) tuples, rendered as a simple row of text boxes."""
    n = len(kpis)
    box_w = Inches(9.2 / n)
    for i, (label, value) in enumerate(kpis):
        box = slide.shapes.add_textbox(Inches(0.4) + box_w * i, top, box_w, Inches(1.0))
        tf = box.text_frame
        tf.word_wrap = True
        p1 = tf.paragraphs[0]
        p1.text = str(value)
        p1.font.size = Pt(22)
        p1.font.bold = True
        p2 = tf.add_paragraph()
        p2.text = label
        p2.font.size = Pt(11)


def _pptx_add_table(slide, headers, rows, top=Inches(2.3), height=Inches(2.8)):
    n_rows = len(rows) + 1
    n_cols = len(headers)
    table_shape = slide.shapes.add_table(n_rows, n_cols, Inches(0.4), top, Inches(9.2), height)
    table = table_shape.table
    for j, h in enumerate(headers):
        cell = table.cell(0, j)
        cell.text = str(h)
        cell.text_frame.paragraphs[0].font.bold = True
        cell.text_frame.paragraphs[0].font.size = Pt(12)
    for i, row in enumerate(rows, start=1):
        for j, val in enumerate(row):
            cell = table.cell(i, j)
            cell.text = str(val)
            cell.text_frame.paragraphs[0].font.size = Pt(11)
    return table_shape


def build_ppt_report(plant_metrics, plant_status_df, line_data):
    """
    plant_metrics: dict with loading/working/major/small/eff for the whole plant
    plant_status_df: aggregated status_df for the whole plant (for top losses table)
    line_data: dict[line] -> {"metrics": {...}, "status_df": DataFrame or None, "actions": DataFrame (for that line)}
    Returns a BytesIO containing the .pptx file.
    """
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.63)
    blank_layout = prs.slide_layouts[6]

    # ---- Slide 1: Plant Summary ----
    slide = prs.slides.add_slide(blank_layout)
    _pptx_add_title(slide, "SMD Plant — Efficiency & Micro-Loss Summary")
    kpis = [
        ("Efficiency %", fmt_pct(plant_metrics["eff"])),
        ("Loading (min)", fmt_min(plant_metrics["loading"])),
        ("Major Loss (min)", fmt_min(plant_metrics["major"])),
        ("Small Loss (min)", fmt_min(plant_metrics["small"])),
    ]
    _pptx_add_kpi_row(slide, kpis)

    if plant_status_df is not None and not plant_status_df.empty:
        ranked = plant_status_df.sort_values("Total_Sec", ascending=False).head(5)
        rows = [[r["Status"], f"{r['Total_Sec']/60:,.1f} min", int(r["Qty"]) if pd.notna(r["Qty"]) else "-"]
                for _, r in ranked.iterrows()]
        _pptx_add_table(slide, ["Top Plant-Level Stoppage", "Total Duration", "Events"], rows)
    else:
        box = slide.shapes.add_textbox(Inches(0.4), Inches(2.3), Inches(9), Inches(0.6))
        box.text_frame.text = "No Machine SW Log data uploaded for this session."

    # ---- Slides 2-5: one per line ----
    for line in LINE_NAMES:
        data = line_data.get(line, {})
        metrics = data.get("metrics", {"loading": 0, "working": 0, "major": 0, "small": 0, "eff": np.nan})
        status_df = data.get("status_df")
        actions_df = data.get("actions")

        slide = prs.slides.add_slide(blank_layout)
        _pptx_add_title(slide, f"{line} — Hot Issues & Action Plan")
        kpis = [
            ("Efficiency %", fmt_pct(metrics["eff"])),
            ("Loading (min)", fmt_min(metrics["loading"])),
            ("Major Loss (min)", fmt_min(metrics["major"])),
            ("Small Loss (min)", fmt_min(metrics["small"])),
        ]
        _pptx_add_kpi_row(slide, kpis)

        if status_df is not None and not status_df.empty:
            top3 = status_df.sort_values("Total_Sec", ascending=False).head(3).reset_index(drop=True)
            rows = []
            for _, r in top3.iterrows():
                entry = get_playbook_entry(r["Status"])
                rows.append([r["Status"], entry["root_cause"], f"{r['Total_Sec']/60:,.1f} min"])
            _pptx_add_table(slide, ["Hot Issue", "Root Cause", "Total Duration"], rows,
                             top=Inches(2.3), height=Inches(1.6))
        else:
            box = slide.shapes.add_textbox(Inches(0.4), Inches(2.3), Inches(9), Inches(0.5))
            box.text_frame.text = "No Machine SW Log data matched to this line."

        # Before/After photos from the line's most relevant Closed action item, if any
        photo_top = Inches(4.1)
        if actions_df is not None and not actions_df.empty:
            with_images = actions_df[actions_df["before_image"].notna() & actions_df["after_image"].notna()]
            if not with_images.empty:
                sample = with_images.iloc[0]
                try:
                    before_stream = io.BytesIO(sample["before_image"])
                    slide.shapes.add_picture(before_stream, Inches(0.4), photo_top, height=Inches(1.3))
                    after_stream = io.BytesIO(sample["after_image"])
                    slide.shapes.add_picture(after_stream, Inches(2.1), photo_top, height=Inches(1.3))
                    cap = slide.shapes.add_textbox(Inches(3.8), photo_top, Inches(5.5), Inches(1.3))
                    cap.text_frame.text = f"Before / After — {sample.get('corrective_action', '')[:180]}"
                    cap.text_frame.paragraphs[0].font.size = Pt(11)
                except Exception:
                    pass  # corrupt image bytes shouldn't break the whole report

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


# ----------------------------------------------------------------------------
# SIDEBAR — RECORD DATE, PER-LINE MANUAL INPUTS, FILE UPLOADS
# ----------------------------------------------------------------------------
DEFAULT_LOADING_MIN = 360.0
DEFAULT_WORKING_MIN = 270.0

with st.sidebar:
    st.markdown(f"<span class='sidebar-version'>⚡ {APP_VERSION}</span>", unsafe_allow_html=True)
    st.markdown("## ⚙️ Data Input")
    plant_name = st.text_input("Facility / Plant Name", value="SMD Plant")
    record_date = st.date_input("📅 Record Date (for saving to database)", value=date.today())
    record_date_str = record_date.isoformat()

    st.markdown("#### 1. Manual Time Entry — Per Line (Minutes)")
    line_inputs = {}
    for line in LINE_NAMES:
        with st.expander(f"📍 {line}", expanded=True):
            loading_val = st.number_input(f"Loading Time (Min) — {line}", min_value=0.0,
                                           value=DEFAULT_LOADING_MIN, step=10.0, key=f"loading_{line}")
            working_val = st.number_input(f"Actual Working Time (Min) — {line}", min_value=0.0,
                                           value=DEFAULT_WORKING_MIN, step=10.0, key=f"working_{line}")
            line_inputs[line] = {"loading_min": loading_val, "working_min": working_val}

    st.markdown("---")
    st.markdown("#### 2. MES Major Loss File (> 5 min)")
    st.caption("`LossRegisterAnalysis_YYYYMMDD` — .xlsx, .csv, or .txt.")
    major_loss_file = st.file_uploader("Upload Major Loss Register", type=UPLOAD_TYPES, key="major_loss_upl")
    unit_override = st.selectbox("If unit can't be auto-detected, assume:", options=["Minutes", "Seconds"], index=0)

    st.markdown("---")
    st.markdown("#### 3. Machine SW Stoppage Log")
    st.caption("`Stoppage@...` — .xlsx (multi-sheet), .csv, or .txt event log.")
    sw_log_file = st.file_uploader("Upload Machine SW Log", type=UPLOAD_TYPES, key="sw_log_upl")

    st.markdown("---")
    top_n_pareto = st.slider("Pareto: number of stoppage categories to show", 3, 15, 8)

    st.markdown("---")
    st.caption(f"Database: local SQLite file `{DB_PATH}`. Save daily records below to enable trend tracking and the Effectiveness Engine.")


# ----------------------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------------------
st.markdown(f"""
    <div class="app-header">
        <div>
            <h1>🏭 SMT Micro-Loss Analyzer &amp; Task Force Action Tracker</h1>
            <p>{plant_name} &nbsp;•&nbsp; Record Date: {record_date_str} &nbsp;•&nbsp; Small Loss (min) = Loading − (Working + Major Loss &gt; 5 min)</p>
        </div>
        <span class="version-badge">⚡ {APP_VERSION}</span>
    </div>
""", unsafe_allow_html=True)

pills = "".join(f"<span class='line-pill'>📍 {l}</span>" for l in LINE_NAMES)
st.markdown(f"<div style='margin-bottom:14px;'>{pills}</div>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# PROCESS UPLOADS
# ----------------------------------------------------------------------------
major_loss_result = None
sw_log_result = None

if major_loss_file is not None:
    try:
        major_loss_result = parse_major_loss_file(major_loss_file, manual_unit_hint=("min" if unit_override == "Minutes" else "sec"))
    except Exception as e:
        major_loss_result = {"ok": False, "df": None, "total_minutes": np.nan, "warnings": [],
                              "errors": [f"Unexpected error while parsing Major Loss file: {e}"]}

if sw_log_file is not None:
    try:
        sw_log_result = parse_sw_log_file(sw_log_file)
    except Exception as e:
        sw_log_result = {"ok": False, "stoppage_df": None, "status_df": None, "module_df": None, "warnings": [],
                          "errors": [f"Unexpected error while parsing SW Log file: {e}"]}

for label, result in [("Major Loss file", major_loss_result), ("Machine SW Log file", sw_log_result)]:
    if result is None:
        continue
    for err in result.get("errors", []):
        st.error(f"**{label}:** {err}")
    for warn in result.get("warnings", []):
        st.warning(f"**{label}:** {warn}")

major_df = major_loss_result["df"] if (major_loss_result and major_loss_result.get("ok")) else None
stoppage_df_all = sw_log_result["stoppage_df"] if (sw_log_result and sw_log_result.get("ok")) else None

major_loss_source = (f"From uploaded file ({len(major_df)} loss events)" if major_df is not None
                      else "No file uploaded — treated as 0 min")
st.markdown(f"<p class='small-note'>Major Loss source: {major_loss_source}</p>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# COMPUTE PER-LINE & OVERALL METRICS
# ----------------------------------------------------------------------------
def compute_metrics(loading_min, working_min, major_loss_min):
    loading_min = max(loading_min or 0.0, 0.0)
    working_min = max(working_min or 0.0, 0.0)
    major_loss_min = max(major_loss_min or 0.0, 0.0)
    small_loss_min = loading_min - (working_min + major_loss_min)
    eff = safe_div(working_min, loading_min) * 100 if loading_min else np.nan
    return {"loading": loading_min, "working": working_min, "major": major_loss_min, "small": small_loss_min, "eff": eff}


line_metrics = {}
line_status_dfs = {}
line_module_dfs = {}

for line in LINE_NAMES:
    line_loading_min = line_inputs[line]["loading_min"]
    line_working_min = line_inputs[line]["working_min"]
    line_major_loss_min = (float(major_df.loc[major_df["Line"] == line, "Duration_Min"].sum())
                            if (major_df is not None and (major_df["Line"] == line).any()) else 0.0)
    line_metrics[line] = compute_metrics(line_loading_min, line_working_min, line_major_loss_min)

    if stoppage_df_all is not None and (stoppage_df_all["Line"] == line).any():
        line_stoppage_df = stoppage_df_all[stoppage_df_all["Line"] == line]
        line_status_dfs[line] = aggregate_status(line_stoppage_df)
        line_module_dfs[line] = aggregate_module(line_stoppage_df)
    else:
        line_status_dfs[line] = None
        line_module_dfs[line] = None

overall_loading_min = sum(m["loading"] for m in line_metrics.values())
overall_working_min = sum(m["working"] for m in line_metrics.values())
overall_major_loss_min = sum(m["major"] for m in line_metrics.values())
overall_metrics = compute_metrics(overall_loading_min, overall_working_min, overall_major_loss_min)

if stoppage_df_all is not None:
    matched_stoppage_all = stoppage_df_all[stoppage_df_all["Line"].isin(LINE_NAMES)]
    unmatched_events = int((stoppage_df_all["Line"] == "Unmatched").sum())
    if unmatched_events > 0:
        st.caption(f"⚠️ {unmatched_events} SW stoppage event(s) could not be matched to a BLOCK4 line and are excluded from per-line/plant Pareto & Action Plan views.")
    overall_status_df = aggregate_status(matched_stoppage_all)
    overall_module_df = aggregate_module(matched_stoppage_all)
elif sw_log_result and sw_log_result.get("ok"):
    overall_status_df = sw_log_result["status_df"]
    overall_module_df = sw_log_result["module_df"]
else:
    overall_status_df, overall_module_df = None, None

# ----------------------------------------------------------------------------
# SAVE TO DATABASE
# ----------------------------------------------------------------------------
save_col1, save_col2 = st.columns([1, 3])
with save_col1:
    if st.button("💾 Save Today's Records to Database", use_container_width=True):
        per_line_for_db = {line: line_metrics[line] for line in LINE_NAMES}
        save_daily_snapshot(record_date_str, per_line_for_db)
        for line in LINE_NAMES:
            if line_status_dfs[line] is not None:
                save_hot_issues(record_date_str, line, line_status_dfs[line], top_n=5)
        st.success(f"Saved records for {record_date_str} across {len(LINE_NAMES)} line(s).")
with save_col2:
    st.caption(
        "Saving stores each line's Loading/Working/Major/Small Loss minutes and its top-5 Hot Issues for "
        f"{record_date_str}. This feeds the Historical Trends tab and the 7-day Effectiveness Engine."
    )

st.markdown("<hr class='divider'>", unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# HOT ISSUES SECTION (with inline "quick add to Action Tracker")
# ----------------------------------------------------------------------------
def render_hot_issues(status_df, section_label, line_for_tracker, top_n=3):
    st.markdown(f"""
    <div class="section-card">
        <div class="section-title">🔥 Hot Issues — {section_label}</div>
        <div class="section-caption">Top chronic stoppage causes driving this section's micro-losses</div>
    """, unsafe_allow_html=True)

    if status_df is None or status_df.empty:
        st.info("Upload the Machine SW Log to surface Hot Issues for this section.")
        st.markdown("</div>", unsafe_allow_html=True)
        return

    ranked = status_df.sort_values("Total_Sec", ascending=False).head(top_n).reset_index(drop=True)
    total_all = status_df["Total_Sec"].sum()
    cols = st.columns(len(ranked)) if len(ranked) > 0 else []
    for i, (col, (_, row)) in enumerate(zip(cols, ranked.iterrows())):
        entry = get_playbook_entry(row["Status"])
        share = safe_div(row["Total_Sec"], total_all) * 100 if total_all else np.nan
        with col:
            st.markdown(f"""
            <div class="hot-card">
                <div class="hot-rank">#{i+1} Chronic Issue</div>
                <div class="hot-name">{row['Status']}</div>
                <div class="hot-metric">{row['Total_Sec']/60:,.1f} min &nbsp;•&nbsp; {fmt_pct(share)}</div>
                <div class="hot-cause">→ {entry['root_cause']}</div>
            </div>
            """, unsafe_allow_html=True)
            if line_for_tracker is not None:
                with st.expander("➕ Log corrective action"):
                    ca_key = f"ca_{line_for_tracker}_{i}_{re.sub(r'[^a-zA-Z0-9]', '_', str(row['Status']))}"
                    corrective_action = st.text_area("Corrective Action", value=entry["action"], key=f"{ca_key}_text", height=100)
                    owner = st.text_input("Responsible Owner", key=f"{ca_key}_owner")
                    target = st.date_input("Target Date", value=date.today() + timedelta(days=7), key=f"{ca_key}_target")
                    if st.button("Add to Action Tracker", key=f"{ca_key}_btn"):
                        add_action_item(line_for_tracker, entry["root_cause"], str(row["Status"]),
                                         corrective_action, owner, target.isoformat())
                        st.success(f"Action item added for {line_for_tracker} — {row['Status']}.")
    st.markdown("</div>", unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# REUSABLE DASHBOARD RENDERER
# ----------------------------------------------------------------------------
def render_dashboard(section_label, metrics, status_df, module_df, top_n, key_prefix, line_for_tracker=None):
    loading_min, working_min, major_loss_min, small_loss_min, efficiency_pct = (
        metrics["loading"], metrics["working"], metrics["major"], metrics["small"], metrics["eff"]
    )
    major_loss_pct = safe_div(major_loss_min, loading_min) * 100 if loading_min else np.nan
    small_loss_pct = safe_div(small_loss_min, loading_min) * 100 if loading_min else np.nan

    data_issue = None
    if loading_min == 0:
        data_issue = "Loading Time is 0 — enter a valid Loading Time in the sidebar to see results."
    elif working_min + major_loss_min > loading_min:
        data_issue = "Actual Working Time + Major Loss exceeds Loading Time. Small Loss is negative — please double-check the time entries."

    k1, k2, k3, k4 = st.columns(4)
    eff_class = "kpi-good" if (not np.isnan(efficiency_pct) and efficiency_pct >= 85) else (
        "kpi-warn" if (not np.isnan(efficiency_pct) and efficiency_pct >= 70) else "kpi-bad")
    with k1:
        st.markdown(f"""<div class="kpi-card {eff_class}"><div class="kpi-label">Machine Efficiency</div>
            <div class="kpi-value">{fmt_pct(efficiency_pct)}</div><div class="kpi-sub">Actual Working / Loading Time</div></div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""<div class="kpi-card kpi-info"><div class="kpi-label">Loading Time</div>
            <div class="kpi-value">{fmt_min(loading_min)}</div><div class="kpi-sub">≈ {loading_min/60:,.2f} h total available time</div></div>""", unsafe_allow_html=True)
    with k3:
        st.markdown(f"""<div class="kpi-card kpi-warn"><div class="kpi-label">Major Loss (&gt; 5 min)</div>
            <div class="kpi-value">{fmt_min(major_loss_min)}</div><div class="kpi-sub">{fmt_pct(major_loss_pct)} of Loading Time</div></div>""", unsafe_allow_html=True)
    with k4:
        small_loss_class = "kpi-bad" if (not np.isnan(small_loss_pct) and small_loss_pct > 15) else "kpi-warn"
        st.markdown(f"""<div class="kpi-card {small_loss_class}"><div class="kpi-label">Small Loss (Hidden, &lt; 5 min)</div>
            <div class="kpi-value">{fmt_min(small_loss_min)}</div><div class="kpi-sub">{fmt_pct(small_loss_pct)} of Loading Time</div></div>""", unsafe_allow_html=True)

    if data_issue:
        st.warning(data_issue)
    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    col_donut, col_pareto = st.columns([1, 1.4])
    with col_donut:
        st.markdown(f"""<div class="section-card"><div class="section-title">⏱️ Time Distribution — {section_label}</div>
            <div class="section-caption">Share of Loading Time by category (minutes)</div>""", unsafe_allow_html=True)
        donut_working, donut_major, donut_small = max(working_min, 0), max(major_loss_min, 0), max(small_loss_min, 0)
        if loading_min > 0 and (donut_working + donut_major + donut_small) > 0:
            fig_donut = go.Figure(data=[go.Pie(
                labels=["Actual Working Time", "Major Loss (>5m)", "Small Loss (Hidden)"],
                values=[donut_working, donut_major, donut_small], hole=0.62,
                marker=dict(colors=["#3ddc84", "#ffb547", "#ff5c5c"], line=dict(color="#0b0f14", width=2)),
                textinfo="percent", textfont=dict(color="#eef3f8", size=13), sort=False,
            )])
            fig_donut.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#a9b7c6"), showlegend=True,
                legend=dict(orientation="h", yanchor="bottom", y=-0.18, x=0.5, xanchor="center"),
                margin=dict(t=10, b=10, l=10, r=10), height=340,
                annotations=[dict(text=f"{fmt_pct(efficiency_pct)}<br><span style='font-size:11px;color:#71828f'>Efficiency</span>",
                                   x=0.5, y=0.5, showarrow=False, font=dict(size=22, color="#eef3f8"))],
            )
            st.plotly_chart(fig_donut, use_container_width=True, key=f"{key_prefix}_donut")
        else:
            st.info("Enter a valid Loading Time (and Working Time) to see the distribution chart.")
        st.markdown("</div>", unsafe_allow_html=True)

    with col_pareto:
        st.markdown(f"""<div class="section-card"><div class="section-title">📊 Machine SW Stoppage Pareto — {section_label}</div>
            <div class="section-caption">Total stoppage duration by status, sorted descending</div>""", unsafe_allow_html=True)
        if status_df is not None and not status_df.empty:
            ranked = status_df.sort_values("Total_Sec", ascending=False).head(top_n).copy()
            ranked["Total_Min"] = ranked["Total_Sec"] / 60.0
            ranked["Cumulative_Pct"] = ranked["Total_Sec"].cumsum() / status_df["Total_Sec"].sum() * 100
            fig_pareto = go.Figure()
            fig_pareto.add_trace(go.Bar(x=ranked["Status"], y=ranked["Total_Min"], name="Total Duration (min)",
                                         marker_color="#3ea6ff", text=[f"{v:,.0f}m" for v in ranked["Total_Min"]],
                                         textposition="outside", textfont=dict(color="#eef3f8")))
            fig_pareto.add_trace(go.Scatter(x=ranked["Status"], y=ranked["Cumulative_Pct"], name="Cumulative %",
                                             yaxis="y2", mode="lines+markers", line=dict(color="#ffb547", width=2), marker=dict(size=6)))
            fig_pareto.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#a9b7c6"),
                margin=dict(t=10, b=10, l=10, r=10), height=340, xaxis=dict(tickangle=-25, gridcolor="#1c2734"),
                yaxis=dict(title="Minutes", gridcolor="#1c2734"),
                yaxis2=dict(title="Cumulative %", overlaying="y", side="right", range=[0, 105], showgrid=False),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, xanchor="left"),
            )
            st.plotly_chart(fig_pareto, use_container_width=True, key=f"{key_prefix}_pareto")
        else:
            st.info("Upload the Machine SW Log to see the Pareto chart for this section.")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)
    render_hot_issues(status_df, section_label, line_for_tracker=line_for_tracker, top_n=3)
    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    if module_df is not None and not module_df.empty:
        st.markdown(f"""<div class="section-card"><div class="section-title">🧩 Module-Level Stoppage Breakdown — {section_label}</div>
            <div class="section-caption">Total duration by machine module</div>""", unsafe_allow_html=True)
        mod_ranked = module_df.sort_values("Total_Sec", ascending=False).copy()
        mod_ranked["Total_Min"] = mod_ranked["Total_Sec"] / 60.0
        fig_mod = go.Figure(go.Bar(x=mod_ranked["Total_Min"], y=mod_ranked["Module"], orientation="h",
                                    marker_color="#00e0c6", text=[f"{v:,.0f}m" for v in mod_ranked["Total_Min"]],
                                    textposition="outside", textfont=dict(color="#eef3f8")))
        fig_mod.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", font=dict(color="#a9b7c6"),
                               margin=dict(t=10, b=10, l=10, r=10), height=max(280, 32 * len(mod_ranked)),
                               xaxis=dict(title="Minutes", gridcolor="#1c2734"), yaxis=dict(autorange="reversed"))
        st.plotly_chart(fig_mod, use_container_width=True, key=f"{key_prefix}_module")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown(f"""<div class="section-card"><div class="section-title">🚨 Automated Priority Action Plan — {section_label}</div>
        <div class="section-caption">Ranked by total stoppage duration — mapped to the standard Task Force response playbook.</div>""", unsafe_allow_html=True)
    if status_df is not None and not status_df.empty:
        plan = build_action_plan(status_df, top_n=top_n)
        for item in plan:
            badge_class = "badge-danger" if item["priority"] == 1 else ("badge-warn" if item["priority"] == 2 else "badge-info")
            priority_label = "TOP PRIORITY" if item["priority"] == 1 else f"PRIORITY {item['priority']}"
            qty_txt = f"{int(item['qty'])} events" if pd.notna(item["qty"]) else "qty n/a"
            st.markdown(f"""
            <div class="action-card action-priority-{item['priority']}">
                <span class="badge {badge_class}">{priority_label}</span>
                <span class="action-title">#{item['rank']} — {item['status_raw']} → {item['root_cause']}</span>
                <div class="action-meta">Total: {item['total_sec']/60:,.1f} min ({seconds_to_hms(item['total_sec'])}) &nbsp;•&nbsp; {qty_txt} &nbsp;•&nbsp; {fmt_pct(item['share_pct'])} of total SW stoppage time</div>
                <div class="action-body">{item['action']}</div>
            </div>""", unsafe_allow_html=True)
    else:
        st.info("Upload the Machine SW Log to generate the automated action plan for this section.")
    st.markdown("</div>", unsafe_allow_html=True)

    with st.expander(f"🔍 View Raw / Parsed Stoppage Data — {section_label}"):
        if status_df is not None and not status_df.empty:
            df_show = status_df.copy()
            df_show["Total_Min"] = df_show["Total_Sec"] / 60.0
            df_show["Min_HMS"] = df_show["Min_Sec"].apply(seconds_to_hms)
            df_show["Max_HMS"] = df_show["Max_Sec"].apply(seconds_to_hms)
            df_show["Total_HMS"] = df_show["Total_Sec"].apply(seconds_to_hms)
            st.dataframe(df_show[["Status", "Qty", "Min_HMS", "Max_HMS", "Total_HMS", "Total_Min"]], use_container_width=True)
        else:
            st.caption("No SW stoppage summary available for this section yet.")


# ----------------------------------------------------------------------------
# MAIN TABS
# ----------------------------------------------------------------------------
tab_labels = ["🏭 Overall SMD Plant", *[f"📍 {l}" for l in LINE_NAMES], "📅 Historical Trends", "🛠️ Action Tracker"]
tabs = st.tabs(tab_labels)

# ---- Overall SMD Plant Dashboard ----
with tabs[0]:
    st.caption("Aggregated totals across all 4 BLOCK4 lines.")

    ppt_col1, ppt_col2 = st.columns([1, 3])
    with ppt_col1:
        if PPTX_AVAILABLE:
            if st.button("📤 Generate PPT Report", use_container_width=True):
                line_data_for_ppt = {}
                for line in LINE_NAMES:
                    actions_for_line = get_action_items(line=line)
                    line_data_for_ppt[line] = {
                        "metrics": line_metrics[line],
                        "status_df": line_status_dfs[line],
                        "actions": actions_for_line,
                    }
                ppt_buf = build_ppt_report(overall_metrics, overall_status_df, line_data_for_ppt)
                st.session_state["ppt_buf"] = ppt_buf.getvalue()
            if "ppt_buf" in st.session_state:
                st.download_button(
                    "⬇️ Download PPT Report", data=st.session_state["ppt_buf"],
                    file_name=f"SMT_Efficiency_Report_{record_date_str}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                )
        else:
            st.warning("`python-pptx` isn't installed — add it to requirements.txt to enable PPT report generation.")
    with ppt_col2:
        st.caption("Builds a title/summary slide plus one slide per BLOCK4 line with its worst 3 Hot Issues, root causes, and Before/After photos from Closed action items.")

    render_dashboard(
        section_label="Overall SMD Plant", metrics=overall_metrics,
        status_df=overall_status_df, module_df=overall_module_df, top_n=top_n_pareto,
        key_prefix="plant", line_for_tracker=None,
    )

# ---- Per-line tabs ----
for i, line in enumerate(LINE_NAMES, start=1):
    with tabs[i]:
        if major_df is not None and not (major_df["Line"] == line).any():
            st.caption(f"⚠️ No Major Loss rows were matched to {line} in the uploaded file — showing 0 min.")
        if stoppage_df_all is not None and not (stoppage_df_all["Line"] == line).any():
            st.caption(f"⚠️ No SW stoppage events were matched to {line} — Pareto/Hot Issues/Action Plan unavailable for it.")
        render_dashboard(
            section_label=line, metrics=line_metrics[line],
            status_df=line_status_dfs[line], module_df=line_module_dfs[line], top_n=top_n_pareto,
            key_prefix=re.sub(r"[^a-zA-Z0-9]", "_", line), line_for_tracker=line,
        )

# ---- Historical Trends ----
with tabs[len(LINE_NAMES) + 1]:
    st.markdown("""<div class="section-card"><div class="section-title">📅 Historical Trends</div>
        <div class="section-caption">Pick a date range and line to review saved daily records.</div>""", unsafe_allow_html=True)

    hc1, hc2, hc3 = st.columns(3)
    with hc1:
        hist_start = st.date_input("Start Date", value=date.today() - timedelta(days=30), key="hist_start")
    with hc2:
        hist_end = st.date_input("End Date", value=date.today(), key="hist_end")
    with hc3:
        hist_line = st.selectbox("Line", options=["All Lines"] + LINE_NAMES, key="hist_line")

    if hist_start > hist_end:
        st.error("Start Date must be on or before End Date.")
    else:
        hist_df = load_daily_records(hist_start.isoformat(), hist_end.isoformat(), line=hist_line)
        if hist_df.empty:
            st.info("No saved records found for this range yet. Use '💾 Save Today's Records to Database' above to start building history.")
        else:
            fig_trend = go.Figure()
            if hist_line == "All Lines":
                for ln in LINE_NAMES:
                    sub = hist_df[hist_df["line"] == ln]
                    if not sub.empty:
                        fig_trend.add_trace(go.Scatter(x=sub["record_date"], y=sub["efficiency_pct"], mode="lines+markers", name=ln))
            else:
                fig_trend.add_trace(go.Scatter(x=hist_df["record_date"], y=hist_df["efficiency_pct"], mode="lines+markers", name=hist_line))
            fig_trend.update_layout(
                title="Efficiency % Trend", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#a9b7c6"), height=360, margin=dict(t=40, b=10, l=10, r=10),
                yaxis=dict(title="Efficiency %", gridcolor="#1c2734"), xaxis=dict(gridcolor="#1c2734"),
            )
            st.plotly_chart(fig_trend, use_container_width=True, key="hist_trend_chart")
            st.dataframe(hist_df, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

# ---- Action Tracker ----
with tabs[len(LINE_NAMES) + 2]:
    st.markdown("""<div class="section-card"><div class="section-title">🛠️ Task Force Action Tracker</div>
        <div class="section-caption">Manage corrective actions end-to-end. Closing an item requires Before &amp; After photos; delaying an item requires a justification and a new Target Date.</div>""", unsafe_allow_html=True)

    with st.expander("➕ Add New Action Item (manual entry)"):
        nc1, nc2 = st.columns(2)
        with nc1:
            new_line = st.selectbox("Line", options=LINE_NAMES, key="new_action_line")
            new_root_cause = st.text_input("Root Cause", key="new_action_root_cause")
            new_status_category = st.text_input("Stoppage / Status Category (for recurrence tracking)", key="new_action_status_cat")
        with nc2:
            new_owner = st.text_input("Responsible Owner", key="new_action_owner")
            new_target = st.date_input("Target Date", value=date.today() + timedelta(days=7), key="new_action_target")
        new_corrective_action = st.text_area("Corrective Action", key="new_action_corrective")
        if st.button("Add Action Item", key="new_action_submit"):
            if not new_corrective_action.strip():
                st.error("Corrective Action cannot be empty.")
            else:
                add_action_item(new_line, new_root_cause, new_status_category, new_corrective_action, new_owner, new_target.isoformat())
                st.success(f"Action item added for {new_line}.")

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    fc1, fc2 = st.columns(2)
    with fc1:
        filter_line = st.selectbox("Filter by Line", options=["All Lines"] + LINE_NAMES, key="tracker_filter_line")
    with fc2:
        filter_status = st.selectbox("Filter by Status", options=["All"] + ACTION_STATUSES, key="tracker_filter_status")

    actions_df = get_action_items(line=filter_line, status=filter_status)
    today = date.today()

    if actions_df.empty:
        st.info("No action items match this filter yet.")
    else:
        for _, item in actions_df.iterrows():
            item_id = int(item["id"])
            eff = evaluate_and_persist_effectiveness(item, today) if item["action_status"] == "Closed" else "NA"

            status_badge_class = {"Open": "badge-info", "In Progress": "badge-warn",
                                   "Delayed": "badge-danger", "Closed": "badge-good"}.get(item["action_status"], "badge-muted")
            eff_badge_class = {"Effective": "badge-good", "Ineffective": "badge-danger",
                               "Pending": "badge-warn"}.get(eff, "badge-muted")

            st.markdown(f"""
            <div class="tracker-card">
                <span class="badge {status_badge_class}">{item['action_status'].upper()}</span>
                {"<span class='badge " + eff_badge_class + "'>" + eff.upper() + "</span>" if eff != "NA" else ""}
                <div class="tracker-title">#{item_id} — {item['line']} · {item['status_category'] or item['root_cause'] or 'General'}</div>
                <div class="tracker-meta">Root Cause: {item['root_cause'] or '—'} &nbsp;•&nbsp; Owner: {item['responsible_owner'] or '—'} &nbsp;•&nbsp; Target: {item['target_date'] or '—'}</div>
            </div>
            """, unsafe_allow_html=True)

            with st.expander(f"Manage #{item_id}"):
                st.write(item["corrective_action"] or "_No corrective action text provided._")
                if item["delay_reason"]:
                    st.caption(f"Last delay reason: {item['delay_reason']}")

                new_status = st.selectbox(
                    "Update Status", options=ACTION_STATUSES,
                    index=ACTION_STATUSES.index(item["action_status"]) if item["action_status"] in ACTION_STATUSES else 0,
                    key=f"status_select_{item_id}",
                )

                has_before = pd.notna(item["before_image"])
                has_after = pd.notna(item["after_image"])

                delay_reason_val, new_target_val = None, None
                before_upload, after_upload = None, None

                if new_status == "Delayed":
                    delay_reason_val = st.text_area("Delay Justification Reason (required)", key=f"delay_reason_{item_id}")
                    new_target_val = st.date_input("New Target Date (required)", value=date.today() + timedelta(days=7), key=f"delay_target_{item_id}")

                if new_status == "Closed":
                    st.caption("Both a Before and an After image are required to close this action.")
                    bc1, bc2 = st.columns(2)
                    with bc1:
                        st.write("Before Image" + (" ✅ on file" if has_before else " — not yet uploaded"))
                        if has_before:
                            st.image(item["before_image"], width=160)
                        before_upload = st.file_uploader("Upload/replace Before Image", type=IMAGE_TYPES, key=f"before_upl_{item_id}")
                    with bc2:
                        st.write("After Image" + (" ✅ on file" if has_after else " — not yet uploaded"))
                        if has_after:
                            st.image(item["after_image"], width=160)
                        after_upload = st.file_uploader("Upload/replace After Image", type=IMAGE_TYPES, key=f"after_upl_{item_id}")

                if st.button("💾 Update Item", key=f"update_btn_{item_id}"):
                    update_fields = {"action_status": new_status}

                    if new_status == "Delayed":
                        if not delay_reason_val or not delay_reason_val.strip():
                            st.error("A Delay Justification Reason is required to mark this item Delayed.")
                        elif new_target_val is None:
                            st.error("A new Target Date is required to mark this item Delayed.")
                        else:
                            update_fields["delay_reason"] = delay_reason_val.strip()
                            update_fields["target_date"] = new_target_val.isoformat()
                            update_action_item(item_id, update_fields)
                            st.success("Item updated and marked Delayed.")
                            st.rerun()

                    elif new_status == "Closed":
                        final_before = before_upload.read() if before_upload is not None else (item["before_image"] if has_before else None)
                        final_after = after_upload.read() if after_upload is not None else (item["after_image"] if has_after else None)
                        if final_before is None or final_after is None:
                            st.error("Cannot close this action — both a Before image AND an After image are required.")
                        else:
                            update_fields["before_image"] = final_before
                            update_fields["after_image"] = final_after
                            if before_upload is not None:
                                update_fields["before_image_name"] = before_upload.name
                            if after_upload is not None:
                                update_fields["after_image_name"] = after_upload.name
                            update_fields["closed_date"] = today.isoformat()
                            update_fields["effectiveness"] = "Pending"
                            update_action_item(item_id, update_fields)
                            st.success("Item closed with Before/After images attached. Effectiveness will be evaluated automatically.")
                            st.rerun()

                    else:
                        update_action_item(item_id, update_fields)
                        st.success(f"Item updated to '{new_status}'.")
                        st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)

st.markdown(
    f"<p class='small-note'>SMT Micro-Loss Analyzer &amp; Task Force Action Tracker · {APP_VERSION} · "
    f"Local SQLite database: <code>{DB_PATH}</code>.</p>",
    unsafe_allow_html=True,
)

