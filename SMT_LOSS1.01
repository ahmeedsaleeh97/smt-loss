"""
SMT Mounter Machine Efficiency & Micro-Loss (Small Loss) Analyzer — Ver 1.01
==============================================================================
Production-ready single-file Streamlit application.

WHAT'S NEW IN VER 1.01
-----------------------
1. All manual time entry and every KPI/chart/metric is now expressed primarily
   in MINUTES (percentages are still relative to Total Loading Time).
2. File uploaders for both the MES Major Loss register and the Machine SW Log
   now accept .xlsx / .csv / .txt, with automatic delimiter detection
   (comma, semicolon, tab, pipe) for text-based files.
3. SMT Line identifiers (e.g. "Line 1", "L2", "SMT-3") are auto-detected from
   both uploaded files. When lines are found, the app renders a separate tab
   per line — each with its own KPI cards, time-distribution donut, Pareto
   chart of SW stoppages, and its own Priority Action Plan — in addition to
   an overall Plant Overview tab.
4. The Automated Action Plan engine is now line-segmented: each line gets a
   ranked, root-cause-mapped action list built from that line's own SW
   stoppage events.

Core Methodology (unchanged)
-----------------------------
    Small Loss (min) = Loading Time (min) - [Actual Working Time (min) + Major Loss (>5m) (min)]

Run with:
    streamlit run app.py
"""

import csv
import io
import re
from datetime import timedelta

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

APP_VERSION = "Ver 1.01"

# ----------------------------------------------------------------------------
# PAGE CONFIG
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title=f"SMT Mounter Efficiency & Small Loss Analyzer — {APP_VERSION}",
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
    --bg-0:#0b0f14;
    --bg-1:#111823;
    --bg-2:#161f2c;
    --bg-3:#1c2734;
    --line:#25313f;
    --accent:#00e0c6;
    --accent-2:#3ea6ff;
    --warn:#ffb547;
    --danger:#ff5c5c;
    --good:#3ddc84;
    --text-0:#eef3f8;
    --text-1:#a9b7c6;
    --text-2:#71828f;
}

html, body, [class*="css"]  {
    font-family: 'Segoe UI', 'Inter', system-ui, sans-serif;
}

.stApp {
    background: radial-gradient(circle at 15% 0%, #0e1621 0%, var(--bg-0) 55%);
    color: var(--text-0);
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, var(--bg-1) 0%, var(--bg-0) 100%);
    border-right: 1px solid var(--line);
}

h1, h2, h3, h4 {
    color: var(--text-0) !important;
    font-weight: 700;
    letter-spacing: 0.3px;
}

.app-header {
    padding: 18px 26px;
    border-radius: 14px;
    background: linear-gradient(120deg, rgba(0,224,198,0.10), rgba(62,166,255,0.06));
    border: 1px solid var(--line);
    margin-bottom: 22px;
    display:flex; align-items:center; justify-content:space-between; flex-wrap:wrap; gap:10px;
}
.app-header h1{
    margin:0;
    font-size: 1.6rem;
}
.app-header p{
    margin: 4px 0 0 0;
    color: var(--text-1);
    font-size: 0.92rem;
}
.version-badge{
    display:inline-block;
    padding: 5px 14px;
    border-radius: 999px;
    background: linear-gradient(120deg, rgba(0,224,198,0.22), rgba(62,166,255,0.18));
    border: 1px solid var(--accent);
    color: var(--accent);
    font-weight:800;
    font-size: 0.82rem;
    letter-spacing: 0.6px;
    white-space:nowrap;
}
.sidebar-version{
    display:inline-block;
    padding: 3px 10px;
    border-radius: 999px;
    background: rgba(0,224,198,0.12);
    border: 1px solid var(--accent);
    color: var(--accent);
    font-weight:700;
    font-size: 0.72rem;
    margin-bottom: 10px;
}

/* KPI CARDS */
.kpi-card {
    background: linear-gradient(160deg, var(--bg-2), var(--bg-1));
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 18px 20px;
    box-shadow: 0 6px 18px rgba(0,0,0,0.35);
    position: relative;
    overflow: hidden;
    height: 100%;
}
.kpi-card::before{
    content:"";
    position:absolute; top:0; left:0; right:0; height:3px;
    background: var(--accent-bar, var(--accent));
}
.kpi-label{
    color: var(--text-2);
    font-size: 0.78rem;
    text-transform: uppercase;
    letter-spacing: 1.1px;
    margin-bottom: 6px;
}
.kpi-value{
    font-size: 1.85rem;
    font-weight: 800;
    color: var(--text-0);
    line-height:1.1;
}
.kpi-sub{
    margin-top: 6px;
    font-size: 0.82rem;
    color: var(--text-1);
}
.kpi-good{ --accent-bar: var(--good); }
.kpi-warn{ --accent-bar: var(--warn); }
.kpi-bad{ --accent-bar: var(--danger); }
.kpi-info{ --accent-bar: var(--accent-2); }

.section-card{
    background: var(--bg-1);
    border: 1px solid var(--line);
    border-radius: 16px;
    padding: 18px 22px;
    margin-bottom: 20px;
}
.section-title{
    font-size: 1.05rem;
    font-weight: 700;
    color: var(--text-0);
    margin-bottom: 4px;
}
.section-caption{
    color: var(--text-2);
    font-size: 0.83rem;
    margin-bottom: 14px;
}

.action-card{
    border-radius: 14px;
    padding: 16px 18px;
    margin-bottom: 12px;
    border-left: 4px solid var(--accent);
    background: var(--bg-2);
}
.action-priority-1{ border-left-color: var(--danger); }
.action-priority-2{ border-left-color: var(--warn); }
.action-priority-3{ border-left-color: var(--accent-2); }
.action-title{ font-weight:700; font-size: 0.98rem; color: var(--text-0); }
.action-meta{ color: var(--text-2); font-size: 0.78rem; margin-top:2px;}
.action-body{ color: var(--text-1); font-size: 0.88rem; margin-top:8px; line-height:1.5;}

.badge{
    display:inline-block;
    padding: 2px 10px;
    border-radius: 999px;
    font-size: 0.72rem;
    font-weight:700;
    letter-spacing:0.4px;
    margin-right:6px;
}
.badge-danger{ background: rgba(255,92,92,0.15); color: var(--danger); border:1px solid rgba(255,92,92,0.4);}
.badge-warn{ background: rgba(255,181,71,0.15); color: var(--warn); border:1px solid rgba(255,181,71,0.4);}
.badge-info{ background: rgba(62,166,255,0.15); color: var(--accent-2); border:1px solid rgba(62,166,255,0.4);}

hr.divider{
    border: none;
    border-top: 1px solid var(--line);
    margin: 22px 0;
}

.stDataFrame, .stTable { border-radius: 12px; overflow:hidden; }

.small-note{ color: var(--text-2); font-size: 0.78rem; }

.line-pill{
    display:inline-block;
    padding: 3px 12px;
    border-radius: 999px;
    background: rgba(62,166,255,0.14);
    border: 1px solid rgba(62,166,255,0.4);
    color: var(--accent-2);
    font-weight:700;
    font-size: 0.78rem;
    margin: 2px 4px 2px 0;
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# CONSTANTS
# ----------------------------------------------------------------------------
MAJOR_LOSS_THRESHOLD_MIN = 5.0  # minutes - anything above this is a "Major Loss"
UPLOAD_TYPES = ["csv", "xlsx", "xls", "txt"]

STOPPAGE_PLAYBOOK = {
    "wait previous": {
        "label": "Wait Previous",
        "root_cause": "Upstream Bottleneck",
        "action": (
            "Line is starving because the upstream process/machine cannot keep pace. "
            "Balance cycle times across the line, check upstream OEE, and consider a small "
            "buffer (WIP) between this mounter and the upstream station. Review conveyor "
            "hand-shake / interlock timing for false waits."
        ),
    },
    "part supply": {
        "label": "Part Supply",
        "root_cause": "Material Delay / Feeder Replenishment",
        "action": (
            "Feeders are running out or component kitting/replenishment is late. Implement "
            "a kanban-based feeder refill schedule, pre-stage next-lot reels at the line side, "
            "and review splice/feeder-exchange SOP for cycle-time impact. Check SMD reel "
            "low-stock alerts and material call-off lead time from the warehouse."
        ),
    },
    "wait next": {
        "label": "Wait Next",
        "root_cause": "Downstream Blocking",
        "action": (
            "Machine is blocked because the downstream process cannot accept boards fast "
            "enough. Check downstream machine cycle time, buffer conveyor capacity, and "
            "AOI/reflow throughput. Consider re-balancing line takt time or adding a small "
            "downstream buffer."
        ),
    },
    "machine error": {
        "label": "Machine Error",
        "root_cause": "Equipment Reliability / Quality Fault",
        "action": (
            "Recurring machine faults (recognition error, nozzle/head fault, pick-up miss, "
            "sensor fault). Pull the fault-code frequency log, prioritize top offending "
            "heads/nozzles for cleaning or replacement, and escalate chronic faults to "
            "equipment engineering for root-cause (5-Why) analysis."
        ),
    },
    "operator downtime": {
        "label": "Operator Downtime",
        "root_cause": "Manning / Operator Response",
        "action": (
            "Losses linked to operator response time or manning gaps (break coverage, "
            "changeover assistance, manual intervention delay). Review manning plan, "
            "response-time SLA for alarms, and provide refresher training on quick "
            "changeover / andon response."
        ),
    },
    "maintenance": {
        "label": "Maintenance",
        "root_cause": "Unplanned/Reactive Maintenance",
        "action": (
            "Time lost to unplanned maintenance intervention. Review PM (preventive "
            "maintenance) compliance and interval, analyze MTBF for the affected module, "
            "and shift toward condition-based maintenance for chronic components."
        ),
    },
}

DEFAULT_PLAYBOOK_ENTRY = {
    "label": "Other / Unclassified",
    "root_cause": "Unclassified Stoppage",
    "action": (
        "This stoppage category is not in the standard playbook. Review the raw SW log "
        "entries for this status and classify with the process engineering team so it can "
        "be added to the standard response playbook."
    ),
}


# ----------------------------------------------------------------------------
# GENERIC HELPERS
# ----------------------------------------------------------------------------
def fmt_min(m):
    if m is None or (isinstance(m, float) and np.isnan(m)):
        return "—"
    return f"{m:,.1f} min"


def fmt_min_with_hr(m):
    if m is None or (isinstance(m, float) and np.isnan(m)):
        return "—"
    return f"{m:,.1f} min <span class='small-note'>(≈ {m/60:,.2f} h)</span>"


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
    """Fuzzy-find a column by lower-cased substring matching on header names."""
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
    """Return the column at the Excel-style letter position (A=0, B=1 ... ) if it exists."""
    idx = 0
    for ch in letter.upper():
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    idx -= 1
    if df is not None and 0 <= idx < len(df.columns):
        return df.columns[idx]
    return None


def parse_hms_to_seconds(value):
    """
    Robustly parse a duration value into seconds.
    Accepts: 'HH:MM:SS', 'MM:SS', pandas.Timedelta, datetime.time,
    numeric seconds/minutes, or NaN.
    """
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
    """Read a .csv/.txt file into a DataFrame, auto-detecting comma/semicolon/tab/pipe delimiters."""
    uploaded_file.seek(0)
    raw_bytes = uploaded_file.read()
    text = _decode_bytes(raw_bytes)
    delimiter = _sniff_delimiter(text)

    df = pd.read_csv(io.StringIO(text), sep=delimiter, engine="python")

    # If the guessed delimiter produced a single (probably wrong) column, try the others.
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
    """
    Returns (payload, kind, meta) where:
      kind == 'sheets' -> payload is a dict[str, DataFrame]  (xlsx/xls)
      kind == 'table'  -> payload is a single DataFrame       (csv/txt)
    meta is a short human-readable string describing how it was read (for warnings).
    """
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
# SMT LINE DETECTION
# ----------------------------------------------------------------------------
def extract_line_label(text):
    """Try to pull a normalized 'Line N' label out of a free-text value."""
    if text is None:
        return None
    s = str(text).strip()
    if s == "" or s.lower() in ("nan", "none", "nat"):
        return None

    m = re.search(r"line\s*[-_ ]?\s*(\d{1,3})", s, re.IGNORECASE)
    if m:
        return f"Line {int(m.group(1))}"

    m = re.search(r"\bsmt\s*[-_ ]?\s*(\d{1,3})\b", s, re.IGNORECASE)
    if m:
        return f"Line {int(m.group(1))}"

    m = re.search(r"\bL[-_ ]?(\d{1,3})\b", s)
    if m:
        return f"Line {int(m.group(1))}"

    return None


def detect_line_column(df):
    """
    Returns (mode, column) where mode is:
      'direct'  -> the column already holds line labels directly
      'extract' -> line labels must be regex-extracted from this column's text
      None      -> no usable line-bearing column found
    """
    col = find_column(df, must_contain_any=["line"])
    if col is not None:
        return "direct", col

    for kw in ["machine", "equipment", "model", "mounter", "station", "eqp"]:
        c = find_column(df, must_contain_any=[kw])
        if c is not None:
            sample = df[c].dropna().astype(str).head(80)
            if len(sample) > 0 and sample.apply(extract_line_label).notna().sum() > 0:
                return "extract", c

    # brute-force scan of text-like columns as a last resort
    for c in df.columns:
        if df[c].dtype == object:
            sample = df[c].dropna().astype(str).head(80)
            if len(sample) == 0:
                continue
            hit_rate = sample.apply(extract_line_label).notna().sum() / len(sample)
            if hit_rate >= 0.3:
                return "extract", c

    return None, None


def build_line_series(df):
    """Returns a pandas Series of normalized line labels (or None) aligned to df's index."""
    mode, col = detect_line_column(df)
    if mode is None:
        return pd.Series([None] * len(df), index=df.index), None

    if mode == "direct":
        def norm_direct(v):
            if v is None:
                return None
            s = str(v).strip()
            if s == "" or s.lower() in ("nan", "none"):
                return None
            lab = extract_line_label(s)
            if lab:
                return lab
            if re.match(r"^\d{1,3}$", s):
                return f"Line {int(s)}"
            return s  # keep custom textual line name as-is (e.g. "Line A")
        return df[col].apply(norm_direct), col
    else:
        return df[col].apply(extract_line_label), col


def line_sort_key(label):
    m = re.search(r"\d+", str(label))
    if m:
        return (0, int(m.group(0)), str(label))
    return (1, 0, str(label))


# ----------------------------------------------------------------------------
# MES MAJOR LOSS FILE PARSER  (LossRegisterAnalysis_YYYYMMDD.csv/.xlsx/.txt)
# ----------------------------------------------------------------------------
def parse_major_loss_file(uploaded_file, manual_unit_hint="min"):
    warnings, errors = [], []
    try:
        payload, kind, meta = load_uploaded_table(uploaded_file)
    except Exception as e:
        return {"ok": False, "df": None, "total_minutes": np.nan, "lines": [],
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
        return {"ok": False, "df": None, "total_minutes": np.nan, "lines": [],
                "warnings": warnings, "errors": ["The Major Loss file appears to be empty."]}

    # ---- locate the duration column (Column K = minutes, Column L = seconds by spec) ----
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
        return {"ok": False, "df": None, "total_minutes": np.nan, "lines": [],
                "warnings": warnings,
                "errors": ["Could not find a 'Loss Time (Min/Sec)' column (expected around column K or L)."]}

    remarks_col = find_column(raw, must_contain_any=["remark", "categor", "reason"])
    if remarks_col is None:
        remarks_col = col_by_letter(raw, "M")
        if remarks_col is not None:
            warnings.append(f"Remarks column not found by header — using column M ('{remarks_col}') by position.")

    line_series, line_col = build_line_series(raw)
    if line_col is not None:
        warnings.append(f"SMT Line identifiers detected in column '{line_col}' of the Major Loss file.")

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
    clean["Line"] = clean["Line"].fillna("Unspecified")

    if clean.empty:
        return {"ok": False, "df": None, "total_minutes": np.nan, "lines": [],
                "warnings": warnings, "errors": ["No valid loss-duration rows were found after parsing."]}

    below_threshold = (clean["Duration_Min"] < MAJOR_LOSS_THRESHOLD_MIN).sum()
    if below_threshold > 0:
        warnings.append(
            f"{below_threshold} row(s) in the Major Loss file are below the {MAJOR_LOSS_THRESHOLD_MIN:.0f}-minute "
            "threshold — they were still included in the Major Loss total as provided by the file."
        )

    total_minutes = float(clean["Duration_Min"].sum())
    detected_lines = sorted([l for l in clean["Line"].unique() if l != "Unspecified"], key=line_sort_key)
    return {"ok": True, "df": clean, "total_minutes": total_minutes, "lines": detected_lines,
            "warnings": warnings, "errors": errors}


# ----------------------------------------------------------------------------
# MACHINE SW STOPPAGE LOG PARSER (.xlsx multi-sheet, or .csv/.txt single table)
# ----------------------------------------------------------------------------
def parse_sw_log_file(uploaded_file):
    warnings, errors = [], []
    try:
        payload, kind, meta = load_uploaded_table(uploaded_file)
    except Exception as e:
        return {"ok": False, "stoppage_df": None, "status_df": None, "module_df": None, "lines": [],
                "warnings": warnings, "errors": [f"Could not read the Machine SW Log file: {e}"]}

    sheets = payload if kind == "sheets" else {"Data": payload}

    # ---- find the best event-level sheet (Span/Status columns present) ----
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
            line_series, line_col = build_line_series(raw1)
            if line_col is not None:
                warnings.append(f"SMT Line identifiers detected in column '{line_col}' of the Machine SW Log.")
            stoppage_df = pd.DataFrame({
                "Machine": raw1[machine_col].astype(str) if machine_col is not None else "N/A",
                "Module": raw1[module_col].astype(str) if module_col is not None else "N/A",
                "Status": raw1[status_col].astype(str).str.strip(),
                "Duration_Sec": raw1[span_col].apply(parse_hms_to_seconds),
                "Line": line_series,
            })
            stoppage_df = stoppage_df.dropna(subset=["Duration_Sec"])
            stoppage_df["Line"] = stoppage_df["Line"].fillna("Unspecified")
            if stoppage_df.empty:
                stoppage_df = None
                warnings.append(f"No valid duration rows found in sheet '{event_sheet_name}'.")
    else:
        warnings.append("Could not identify an event-level Stoppage sheet/table with Span & Status columns.")

    status_df, module_df = None, None

    if stoppage_df is not None and not stoppage_df.empty:
        status_df = (
            stoppage_df.groupby("Status")["Duration_Sec"]
            .agg(Qty="size", Min_Sec="min", Max_Sec="max", Total_Sec="sum")
            .reset_index()
        )
        module_df = (
            stoppage_df.groupby("Module")["Duration_Sec"]
            .agg(Total_Sec="sum")
            .reset_index()
        )
    elif kind == "sheets":
        # fallback: use pre-aggregated Stoppage(Status)/Stoppage(Module) sheets if present
        # (line segmentation will not be available in this fallback path)
        status_sheet = None
        for n in sheets:
            if "status" in str(n).lower():
                status_sheet = n
                break
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
                    warnings.append(
                        f"Using pre-aggregated sheet '{status_sheet}' — per-line breakdown is unavailable "
                        "in this fallback mode because SMT Line could not be tied to individual events."
                    )
            except Exception as e:
                warnings.append(f"Could not parse fallback sheet '{status_sheet}': {e}")

        module_sheet = None
        for n in sheets:
            if "module" in str(n).lower():
                module_sheet = n
                break
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
        errors.append(
            "Could not build a stoppage-by-status summary — the Pareto chart and Priority "
            "Action Plan need either an event-level Stoppage table or a Stoppage(Status) sheet."
        )

    detected_lines = []
    if stoppage_df is not None and not stoppage_df.empty:
        detected_lines = sorted([l for l in stoppage_df["Line"].unique() if l != "Unspecified"], key=line_sort_key)

    return {"ok": ok, "stoppage_df": stoppage_df, "status_df": status_df, "module_df": module_df,
            "lines": detected_lines, "warnings": warnings, "errors": errors}


def aggregate_status(stoppage_df):
    if stoppage_df is None or stoppage_df.empty:
        return None
    return (
        stoppage_df.groupby("Status")["Duration_Sec"]
        .agg(Qty="size", Min_Sec="min", Max_Sec="max", Total_Sec="sum")
        .reset_index()
    )


def aggregate_module(stoppage_df):
    if stoppage_df is None or stoppage_df.empty:
        return None
    return stoppage_df.groupby("Module")["Duration_Sec"].agg(Total_Sec="sum").reset_index()


# ----------------------------------------------------------------------------
# ACTION PLAN ENGINE
# ----------------------------------------------------------------------------
def build_action_plan(status_df, top_n=5):
    if status_df is None or status_df.empty:
        return []

    ranked = status_df.sort_values("Total_Sec", ascending=False).reset_index(drop=True)
    total_all = ranked["Total_Sec"].sum()
    plan = []
    for i, row in ranked.head(top_n).iterrows():
        key = str(row["Status"]).strip().lower()
        entry = None
        for pk, pv in STOPPAGE_PLAYBOOK.items():
            if pk in key:
                entry = pv
                break
        if entry is None:
            entry = DEFAULT_PLAYBOOK_ENTRY

        share = safe_div(row["Total_Sec"], total_all) * 100 if total_all else np.nan
        priority = 1 if i == 0 else (2 if i in (1, 2) else 3)
        plan.append({
            "rank": i + 1, "priority": priority, "status_raw": row["Status"],
            "label": entry["label"], "root_cause": entry["root_cause"], "action": entry["action"],
            "total_sec": row["Total_Sec"], "qty": row.get("Qty", np.nan), "share_pct": share,
        })
    return plan


# ----------------------------------------------------------------------------
# SIDEBAR — MANUAL INPUTS & FILE UPLOADS
# ----------------------------------------------------------------------------
with st.sidebar:
    st.markdown(f"<span class='sidebar-version'>⚡ {APP_VERSION}</span>", unsafe_allow_html=True)
    st.markdown("## ⚙️ Data Input")
    st.markdown("#### 1. Manual Time Entry (Minutes)")
    plant_name = st.text_input("Facility / Plant Name", value="SMT Plant")
    plant_loading_min = st.number_input(
        "Loading Time (Minutes)", min_value=0.0, value=1440.0, step=30.0,
        help="Total scheduled/available production time for the period being analyzed, in minutes."
    )
    plant_working_min = st.number_input(
        "Actual Working Time (Minutes)", min_value=0.0, value=1080.0, step=30.0,
        help="Actual machine running/production time for the period, in minutes."
    )

    st.markdown("---")
    st.markdown("#### 2. MES Major Loss File (> 5 min)")
    st.caption("`LossRegisterAnalysis_YYYYMMDD` — .xlsx, .csv, or .txt. Durations in column K/L, remarks in column M.")
    major_loss_file = st.file_uploader("Upload Major Loss Register", type=UPLOAD_TYPES, key="major_loss_upl")
    unit_override = st.selectbox(
        "If unit can't be auto-detected, assume:", options=["Minutes", "Seconds"], index=0,
        help="Only used as a fallback when the duration column can't be identified by its header name."
    )

    st.markdown("---")
    st.markdown("#### 3. Machine SW Stoppage Log")
    st.caption("`Stoppage@...` — .xlsx (multi-sheet), .csv, or .txt event log.")
    sw_log_file = st.file_uploader("Upload Machine SW Log", type=UPLOAD_TYPES, key="sw_log_upl")

    st.markdown("---")
    top_n_pareto = st.slider("Pareto: number of stoppage categories to show", 3, 15, 8)


# ----------------------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------------------
st.markdown(
    f"""
    <div class="app-header">
        <div>
            <h1>🏭 SMT Mounter Efficiency &amp; Small-Loss Analyzer</h1>
            <p>{plant_name} &nbsp;•&nbsp; Small Loss (min) = Loading − (Working + Major Loss &gt; 5 min)</p>
        </div>
        <span class="version-badge">⚡ {APP_VERSION}</span>
    </div>
    """,
    unsafe_allow_html=True,
)

# ----------------------------------------------------------------------------
# PROCESS UPLOADS
# ----------------------------------------------------------------------------
major_loss_result = None
sw_log_result = None

if major_loss_file is not None:
    try:
        major_loss_result = parse_major_loss_file(
            major_loss_file, manual_unit_hint=("min" if unit_override == "Minutes" else "sec")
        )
    except Exception as e:
        major_loss_result = {"ok": False, "df": None, "total_minutes": np.nan, "lines": [],
                              "warnings": [], "errors": [f"Unexpected error while parsing Major Loss file: {e}"]}

if sw_log_file is not None:
    try:
        sw_log_result = parse_sw_log_file(sw_log_file)
    except Exception as e:
        sw_log_result = {"ok": False, "stoppage_df": None, "status_df": None, "module_df": None, "lines": [],
                          "warnings": [], "errors": [f"Unexpected error while parsing SW Log file: {e}"]}

for label, result in [("Major Loss file", major_loss_result), ("Machine SW Log file", sw_log_result)]:
    if result is None:
        continue
    for err in result.get("errors", []):
        st.error(f"**{label}:** {err}")
    for warn in result.get("warnings", []):
        st.warning(f"**{label}:** {warn}")

# ----------------------------------------------------------------------------
# DETECT SMT LINES (union across both files)
# ----------------------------------------------------------------------------
lines_from_major = major_loss_result["lines"] if (major_loss_result and major_loss_result.get("ok")) else []
lines_from_sw = sw_log_result["lines"] if (sw_log_result and sw_log_result.get("ok")) else []
detected_lines = sorted(set(lines_from_major) | set(lines_from_sw), key=line_sort_key)

if detected_lines:
    pills = "".join(f"<span class='line-pill'>📍 {l}</span>" for l in detected_lines)
    st.markdown(f"<div style='margin-bottom:14px;'>{pills}</div>", unsafe_allow_html=True)


# ----------------------------------------------------------------------------
# REUSABLE DASHBOARD RENDERER (used for Plant Overview + each Line tab)
# ----------------------------------------------------------------------------
def render_dashboard(section_label, loading_min, working_min, major_loss_min,
                      status_df, module_df, top_n, key_prefix):

    loading_min = max(loading_min or 0.0, 0.0)
    working_min = max(working_min or 0.0, 0.0)
    major_loss_min = max(major_loss_min or 0.0, 0.0)

    small_loss_min = loading_min - (working_min + major_loss_min)

    data_issue = None
    if loading_min == 0:
        data_issue = "Loading Time is 0 — enter a valid Loading Time to see results."
    elif working_min + major_loss_min > loading_min:
        data_issue = (
            "Actual Working Time + Major Loss exceeds Loading Time. Small Loss is negative — "
            "please double-check the time entries and the Major Loss file totals for this line."
        )

    efficiency_pct = safe_div(working_min, loading_min) * 100 if loading_min else np.nan
    major_loss_pct = safe_div(major_loss_min, loading_min) * 100 if loading_min else np.nan
    small_loss_pct = safe_div(small_loss_min, loading_min) * 100 if loading_min else np.nan

    # ---- KPI CARDS ----
    k1, k2, k3, k4 = st.columns(4)
    eff_class = "kpi-good" if (not np.isnan(efficiency_pct) and efficiency_pct >= 85) else (
        "kpi-warn" if (not np.isnan(efficiency_pct) and efficiency_pct >= 70) else "kpi-bad"
    )
    with k1:
        st.markdown(f"""
        <div class="kpi-card {eff_class}">
            <div class="kpi-label">Machine Efficiency</div>
            <div class="kpi-value">{fmt_pct(efficiency_pct)}</div>
            <div class="kpi-sub">Actual Working / Loading Time</div>
        </div>""", unsafe_allow_html=True)
    with k2:
        st.markdown(f"""
        <div class="kpi-card kpi-info">
            <div class="kpi-label">Loading Time</div>
            <div class="kpi-value">{fmt_min(loading_min)}</div>
            <div class="kpi-sub">≈ {loading_min/60:,.2f} h total available time</div>
        </div>""", unsafe_allow_html=True)
    with k3:
        st.markdown(f"""
        <div class="kpi-card kpi-warn">
            <div class="kpi-label">Major Loss (&gt; 5 min)</div>
            <div class="kpi-value">{fmt_min(major_loss_min)}</div>
            <div class="kpi-sub">{fmt_pct(major_loss_pct)} of Loading Time</div>
        </div>""", unsafe_allow_html=True)
    with k4:
        small_loss_class = "kpi-bad" if (not np.isnan(small_loss_pct) and small_loss_pct > 15) else "kpi-warn"
        st.markdown(f"""
        <div class="kpi-card {small_loss_class}">
            <div class="kpi-label">Small Loss (Hidden, &lt; 5 min)</div>
            <div class="kpi-value">{fmt_min(small_loss_min)}</div>
            <div class="kpi-sub">{fmt_pct(small_loss_pct)} of Loading Time</div>
        </div>""", unsafe_allow_html=True)

    if data_issue:
        st.warning(data_issue)

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    # ---- DONUT + PARETO ----
    col_donut, col_pareto = st.columns([1, 1.4])

    with col_donut:
        st.markdown(f"""
        <div class="section-card">
            <div class="section-title">⏱️ Time Distribution — {section_label}</div>
            <div class="section-caption">Share of Loading Time by category (minutes)</div>
        """, unsafe_allow_html=True)

        donut_working = max(working_min, 0)
        donut_major = max(major_loss_min, 0)
        donut_small = max(small_loss_min, 0)

        if loading_min > 0 and (donut_working + donut_major + donut_small) > 0:
            fig_donut = go.Figure(data=[go.Pie(
                labels=["Actual Working Time", "Major Loss (>5m)", "Small Loss (Hidden)"],
                values=[donut_working, donut_major, donut_small],
                hole=0.62,
                marker=dict(colors=["#3ddc84", "#ffb547", "#ff5c5c"], line=dict(color="#0b0f14", width=2)),
                textinfo="percent",
                textfont=dict(color="#eef3f8", size=13),
                sort=False,
            )])
            fig_donut.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#a9b7c6"), showlegend=True,
                legend=dict(orientation="h", yanchor="bottom", y=-0.18, x=0.5, xanchor="center"),
                margin=dict(t=10, b=10, l=10, r=10), height=340,
                annotations=[dict(
                    text=f"{fmt_pct(efficiency_pct)}<br><span style='font-size:11px;color:#71828f'>Efficiency</span>",
                    x=0.5, y=0.5, showarrow=False, font=dict(size=22, color="#eef3f8")
                )],
            )
            st.plotly_chart(fig_donut, use_container_width=True, key=f"{key_prefix}_donut")
        else:
            st.info("Enter a valid Loading Time (and Working Time) to see the distribution chart.")
        st.markdown("</div>", unsafe_allow_html=True)

    with col_pareto:
        st.markdown(f"""
        <div class="section-card">
            <div class="section-title">📊 Machine SW Stoppage Pareto — {section_label}</div>
            <div class="section-caption">Total stoppage duration by status, sorted descending</div>
        """, unsafe_allow_html=True)

        if status_df is not None and not status_df.empty:
            ranked = status_df.sort_values("Total_Sec", ascending=False).head(top_n).copy()
            ranked["Total_Min"] = ranked["Total_Sec"] / 60.0
            ranked["Cumulative_Pct"] = ranked["Total_Sec"].cumsum() / status_df["Total_Sec"].sum() * 100

            fig_pareto = go.Figure()
            fig_pareto.add_trace(go.Bar(
                x=ranked["Status"], y=ranked["Total_Min"], name="Total Duration (min)",
                marker_color="#3ea6ff",
                text=[f"{v:,.0f}m" for v in ranked["Total_Min"]],
                textposition="outside", textfont=dict(color="#eef3f8"),
            ))
            fig_pareto.add_trace(go.Scatter(
                x=ranked["Status"], y=ranked["Cumulative_Pct"], name="Cumulative %",
                yaxis="y2", mode="lines+markers",
                line=dict(color="#ffb547", width=2), marker=dict(size=6),
            ))
            fig_pareto.update_layout(
                paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                font=dict(color="#a9b7c6"), margin=dict(t=10, b=10, l=10, r=10), height=340,
                xaxis=dict(tickangle=-25, gridcolor="#1c2734"),
                yaxis=dict(title="Minutes", gridcolor="#1c2734"),
                yaxis2=dict(title="Cumulative %", overlaying="y", side="right", range=[0, 105], showgrid=False),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, xanchor="left"),
            )
            st.plotly_chart(fig_pareto, use_container_width=True, key=f"{key_prefix}_pareto")
        else:
            st.info("Upload the Machine SW Log to see the Pareto chart for this section.")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("<hr class='divider'>", unsafe_allow_html=True)

    # ---- MODULE BREAKDOWN ----
    if module_df is not None and not module_df.empty:
        st.markdown(f"""
        <div class="section-card">
            <div class="section-title">🧩 Module-Level Stoppage Breakdown — {section_label}</div>
            <div class="section-caption">Total duration by machine module</div>
        """, unsafe_allow_html=True)

        mod_ranked = module_df.sort_values("Total_Sec", ascending=False).copy()
        mod_ranked["Total_Min"] = mod_ranked["Total_Sec"] / 60.0

        fig_mod = go.Figure(go.Bar(
            x=mod_ranked["Total_Min"], y=mod_ranked["Module"], orientation="h",
            marker_color="#00e0c6",
            text=[f"{v:,.0f}m" for v in mod_ranked["Total_Min"]],
            textposition="outside", textfont=dict(color="#eef3f8"),
        ))
        fig_mod.update_layout(
            paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#a9b7c6"), margin=dict(t=10, b=10, l=10, r=10),
            height=max(280, 32 * len(mod_ranked)),
            xaxis=dict(title="Minutes", gridcolor="#1c2734"),
            yaxis=dict(autorange="reversed"),
        )
        st.plotly_chart(fig_mod, use_container_width=True, key=f"{key_prefix}_module")
        st.markdown("</div>", unsafe_allow_html=True)

    # ---- ACTION PLAN ----
    st.markdown(f"""
    <div class="section-card">
        <div class="section-title">🚨 Automated Priority Action Plan — {section_label}</div>
        <div class="section-caption">
            Ranked by total stoppage duration — mapped to the standard Task Force response playbook.
        </div>
    """, unsafe_allow_html=True)

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
                <div class="action-meta">
                    Total: {item['total_sec']/60:,.1f} min ({seconds_to_hms(item['total_sec'])}) &nbsp;•&nbsp;
                    {qty_txt} &nbsp;•&nbsp; {fmt_pct(item['share_pct'])} of total SW stoppage time
                </div>
                <div class="action-body">{item['action']}</div>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("Upload the Machine SW Log to generate the automated action plan for this section.")

    st.markdown("</div>", unsafe_allow_html=True)

    # ---- DETAIL TABLES ----
    with st.expander(f"🔍 View Raw / Parsed Data — {section_label}"):
        if status_df is not None and not status_df.empty:
            df_show = status_df.copy()
            df_show["Total_Min"] = df_show["Total_Sec"] / 60.0
            df_show["Min_HMS"] = df_show["Min_Sec"].apply(seconds_to_hms)
            df_show["Max_HMS"] = df_show["Max_Sec"].apply(seconds_to_hms)
            df_show["Total_HMS"] = df_show["Total_Sec"].apply(seconds_to_hms)
            st.dataframe(df_show[["Status", "Qty", "Min_HMS", "Max_HMS", "Total_HMS", "Total_Min"]],
                         use_container_width=True)
        else:
            st.caption("No SW stoppage summary available for this section yet.")

    return small_loss_min, efficiency_pct


# ----------------------------------------------------------------------------
# MAIN LAYOUT: PLANT OVERVIEW TAB + PER-LINE TABS
# ----------------------------------------------------------------------------
overall_major_loss_min = major_loss_result["total_minutes"] if (major_loss_result and major_loss_result.get("ok")) else 0.0
overall_status_df = sw_log_result["status_df"] if (sw_log_result and sw_log_result.get("ok")) else None
overall_module_df = sw_log_result["module_df"] if (sw_log_result and sw_log_result.get("ok")) else None
major_loss_source = (
    f"From uploaded file ({len(major_loss_result['df'])} loss events)"
    if (major_loss_result and major_loss_result.get("ok")) else "No file uploaded — treated as 0 min"
)

if not detected_lines:
    st.markdown(f"<p class='small-note'>Major Loss source: {major_loss_source} · No SMT Line identifiers were detected in the uploaded files, so results are shown at plant level only.</p>", unsafe_allow_html=True)
    render_dashboard(
        section_label="Plant Overview",
        loading_min=plant_loading_min, working_min=plant_working_min, major_loss_min=overall_major_loss_min,
        status_df=overall_status_df, module_df=overall_module_df, top_n=top_n_pareto, key_prefix="plant",
    )
else:
    st.markdown(f"<p class='small-note'>Major Loss source: {major_loss_source} · {len(detected_lines)} SMT Line(s) detected — see per-line tabs below.</p>", unsafe_allow_html=True)

    tab_labels = ["🏭 Plant Overview"] + [f"📍 {l}" for l in detected_lines]
    tabs = st.tabs(tab_labels)

    with tabs[0]:
        st.caption("Aggregate figures across all detected lines and any unspecified stoppages.")
        render_dashboard(
            section_label="Plant Overview",
            loading_min=plant_loading_min, working_min=plant_working_min, major_loss_min=overall_major_loss_min,
            status_df=overall_status_df, module_df=overall_module_df, top_n=top_n_pareto, key_prefix="plant",
        )

    major_df = major_loss_result["df"] if (major_loss_result and major_loss_result.get("ok")) else None
    stoppage_df_all = sw_log_result["stoppage_df"] if (sw_log_result and sw_log_result.get("ok")) else None

    n_lines = len(detected_lines)
    for i, line in enumerate(detected_lines, start=1):
        with tabs[i]:
            st.caption(
                f"Enter this line's Loading & Working Time — defaults are the Plant total split evenly "
                f"across the {n_lines} detected line(s); adjust to match this line's actual schedule."
            )
            c1, c2 = st.columns(2)
            default_loading = plant_loading_min / n_lines if n_lines else plant_loading_min
            default_working = plant_working_min / n_lines if n_lines else plant_working_min
            with c1:
                line_loading_min = st.number_input(
                    f"Loading Time (Minutes) — {line}", min_value=0.0,
                    value=float(round(default_loading, 1)), step=10.0, key=f"loading_{line}",
                )
            with c2:
                line_working_min = st.number_input(
                    f"Actual Working Time (Minutes) — {line}", min_value=0.0,
                    value=float(round(default_working, 1)), step=10.0, key=f"working_{line}",
                )

            if major_df is not None and (major_df["Line"] == line).any():
                line_major_loss_min = float(major_df.loc[major_df["Line"] == line, "Duration_Min"].sum())
            elif major_df is not None and line not in set(major_df["Line"].unique()):
                line_major_loss_min = 0.0
                st.caption("⚠️ This line was not found in the Major Loss file — Major Loss is shown as 0 min for it.")
            else:
                line_major_loss_min = 0.0

            if stoppage_df_all is not None and (stoppage_df_all["Line"] == line).any():
                line_stoppage_df = stoppage_df_all[stoppage_df_all["Line"] == line]
                line_status_df = aggregate_status(line_stoppage_df)
                line_module_df = aggregate_module(line_stoppage_df)
            else:
                line_status_df, line_module_df = None, None
                if stoppage_df_all is not None:
                    st.caption("⚠️ This line was not found in the Machine SW Log — Pareto/Action Plan unavailable for it.")

            render_dashboard(
                section_label=line,
                loading_min=line_loading_min, working_min=line_working_min, major_loss_min=line_major_loss_min,
                status_df=line_status_df, module_df=line_module_df, top_n=top_n_pareto,
                key_prefix=re.sub(r"[^a-zA-Z0-9]", "_", line),
            )

st.markdown(
    f"<p class='small-note'>SMT Mounter Efficiency & Small-Loss Analyzer · {APP_VERSION} · "
    "All calculations are performed locally in this session — no data is stored.</p>",
    unsafe_allow_html=True,
)

