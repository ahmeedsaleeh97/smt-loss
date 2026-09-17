"""
SMT Mounter Machine Efficiency & Micro-Loss (Small Loss) Analyzer
===================================================================
Production-ready single-file Streamlit application.

Purpose
-------
SMT (Surface Mount Technology) mounter machines are usually monitored by an
MES (Manufacturing Execution System) that only logs "Major Losses" (stoppages
longer than 5 minutes). In reality, machines also suffer from a large number
of very short stoppages (feeder waits, small part-supply hiccups, micro
jams, sensor faults that self-clear, etc.) that individually last only a
few seconds to a few minutes. These "Small Losses" (a.k.a. Micro-Losses)
never show up as a discrete row in the MES major-loss report, but their
cumulative effect can be a very large hidden efficiency killer.

This app reconstructs the Small Loss bucket using a simple time-balance
identity:

    Small Loss = Loading Time - (Actual Working Time + Major Loss (> 5 min))

It then cross-references that hidden bucket against the machine's own
SW (software) stoppage logs -- which DO capture every stoppage, regardless
of duration -- to work out *which* stoppage category is most likely
responsible for the hidden loss, and produces an automatic, prioritized
action plan for the shop-floor task force.

Run with:
    streamlit run app.py
"""

import io
import re
import traceback
from datetime import timedelta

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ----------------------------------------------------------------------------
# PAGE CONFIG
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="SMT Mounter Efficiency & Small Loss Analyzer",
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
}
.app-header h1{
    margin:0;
    font-size: 1.65rem;
}
.app-header p{
    margin: 4px 0 0 0;
    color: var(--text-1);
    font-size: 0.92rem;
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
    font-size: 1.9rem;
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
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# CONSTANTS
# ----------------------------------------------------------------------------
MAJOR_LOSS_THRESHOLD_MIN = 5.0  # minutes - anything above this is a "Major Loss"

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
def fmt_hours(h):
    if h is None or (isinstance(h, float) and np.isnan(h)):
        return "—"
    return f"{h:,.2f} h"


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
    cols = list(df.columns)
    for c in cols:
        name = str(c).strip().lower()
        ok_all = True
        if must_contain_all:
            ok_all = all(k in name for k in must_contain_all)
        ok_any = True
        if must_contain_any:
            ok_any = any(k in name for k in must_contain_any)
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

    # datetime.time or pandas Timestamp-like (openpyxl sometimes returns time objects)
    if hasattr(value, "hour") and hasattr(value, "minute") and hasattr(value, "second"):
        try:
            return value.hour * 3600 + value.minute * 60 + value.second
        except Exception:
            pass

    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none", "nat"):
        return np.nan

    # HH:MM:SS or MM:SS or H:MM:SS.sss
    match = re.match(r"^(\d+):(\d{1,2}):(\d{1,2}(?:\.\d+)?)$", s)
    if match:
        h, m, sec = match.groups()
        return int(h) * 3600 + int(m) * 60 + float(sec)

    match = re.match(r"^(\d{1,2}):(\d{1,2}(?:\.\d+)?)$", s)
    if match:
        m, sec = match.groups()
        return int(m) * 60 + float(sec)

    # plain number as string (e.g. "125.5")
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


def read_any_table(uploaded_file, sheet_name=0):
    """Read a CSV or Excel file into a DataFrame, handling common quirks."""
    name = uploaded_file.name.lower()
    uploaded_file.seek(0)
    if name.endswith(".csv"):
        # try a couple of encodings/separators defensively
        for enc in ("utf-8-sig", "utf-8", "latin1"):
            try:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, encoding=enc)
            except Exception:
                continue
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file, engine="python")
    else:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="openpyxl")


def read_all_sheets(uploaded_file):
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")


def find_sheet(sheets_dict, keywords):
    """Find the best-matching sheet name for a list of lower-case keyword hints, in order."""
    names = list(sheets_dict.keys())
    lower_map = {n: str(n).strip().lower() for n in names}
    for kw in keywords:
        for n in names:
            if kw in lower_map[n]:
                return n
    return None


# ----------------------------------------------------------------------------
# MES MAJOR LOSS FILE PARSER  (LossRegisterAnalysis_YYYYMMDD.csv/.xlsx)
# ----------------------------------------------------------------------------
def parse_major_loss_file(uploaded_file, manual_unit_hint=None):
    """
    Returns:
        result dict with:
            ok (bool)
            df (DataFrame) - cleaned [Duration_Min, Remarks]
            total_minutes (float)
            warnings (list[str])
            errors (list[str])
    """
    warnings, errors = [], []
    try:
        raw = read_any_table(uploaded_file)
    except Exception as e:
        return {
            "ok": False, "df": None, "total_minutes": np.nan,
            "warnings": warnings, "errors": [f"Could not read the Major Loss file: {e}"],
        }

    if raw is None or raw.empty:
        return {
            "ok": False, "df": None, "total_minutes": np.nan,
            "warnings": warnings, "errors": ["The Major Loss file appears to be empty."],
        }

    # ---- locate the duration column (Column K = minutes, Column L = seconds by spec) ----
    min_col = find_column(raw, must_contain_all=["loss"], must_contain_any=["min"])
    sec_col = find_column(raw, must_contain_all=["loss"], must_contain_any=["sec"])

    duration_col, unit = None, None
    if min_col is not None:
        duration_col, unit = min_col, "min"
    elif sec_col is not None:
        duration_col, unit = sec_col, "sec"
    else:
        # positional fallback: K = index 10 (minutes), L = index 11 (seconds)
        k_col = col_by_letter(raw, "K")
        l_col = col_by_letter(raw, "L")
        if k_col is not None and raw[k_col].apply(lambda v: parse_hms_to_seconds(v)).notna().sum() > 0:
            duration_col, unit = k_col, (manual_unit_hint or "min")
            warnings.append(
                f"Loss duration column not found by header name — falling back to column K "
                f"('{k_col}') by position, assumed unit = {unit}."
            )
        elif l_col is not None and raw[l_col].apply(lambda v: parse_hms_to_seconds(v)).notna().sum() > 0:
            duration_col, unit = l_col, (manual_unit_hint or "sec")
            warnings.append(
                f"Loss duration column not found by header name — falling back to column L "
                f"('{l_col}') by position, assumed unit = {unit}."
            )

    if duration_col is None:
        return {
            "ok": False, "df": None, "total_minutes": np.nan,
            "warnings": warnings,
            "errors": [
                "Could not find a 'Loss Time (Min/Sec)' column (expected around column K or L). "
                "Please check the file layout."
            ],
        }

    # ---- locate the remarks column (Column M by spec) ----
    remarks_col = find_column(raw, must_contain_any=["remark", "categor", "reason"])
    if remarks_col is None:
        remarks_col = col_by_letter(raw, "M")
        if remarks_col is not None:
            warnings.append(f"Remarks column not found by header name — using column M ('{remarks_col}') by position.")

    seconds_series = raw[duration_col].apply(parse_hms_to_seconds)
    n_bad = seconds_series.isna().sum()
    if n_bad > 0:
        warnings.append(f"{n_bad} row(s) in the loss duration column could not be parsed and were ignored.")

    minutes_series = seconds_series / 60.0 if unit == "sec" else seconds_series

    clean = pd.DataFrame({
        "Duration_Min": minutes_series,
    })
    clean["Remarks"] = raw[remarks_col].astype(str).str.strip() if remarks_col is not None else "Unclassified"
    clean = clean.dropna(subset=["Duration_Min"])
    clean = clean[clean["Duration_Min"] >= 0]

    if clean.empty:
        return {
            "ok": False, "df": None, "total_minutes": np.nan,
            "warnings": warnings, "errors": ["No valid loss-duration rows were found after parsing."],
        }

    # Sanity check vs 5-minute definition (informational only, not a hard filter,
    # since the source file is expected to already be the >5 min register).
    below_threshold = (clean["Duration_Min"] < MAJOR_LOSS_THRESHOLD_MIN).sum()
    if below_threshold > 0:
        warnings.append(
            f"{below_threshold} row(s) in the Major Loss file are below the {MAJOR_LOSS_THRESHOLD_MIN:.0f}-minute "
            "threshold — they were still included in the Major Loss total as provided by the file."
        )

    total_minutes = float(clean["Duration_Min"].sum())
    return {"ok": True, "df": clean, "total_minutes": total_minutes, "warnings": warnings, "errors": errors}


# ----------------------------------------------------------------------------
# MACHINE SW STOPPAGE LOG PARSER (multi-sheet Stoppage@....xlsx)
# ----------------------------------------------------------------------------
def parse_sw_log_file(uploaded_file):
    warnings, errors = [], []
    try:
        sheets = read_all_sheets(uploaded_file)
    except Exception as e:
        return {
            "ok": False, "stoppage_df": None, "status_df": None, "module_df": None,
            "warnings": warnings, "errors": [f"Could not read the Machine SW Log workbook: {e}"],
        }

    if not sheets:
        return {
            "ok": False, "stoppage_df": None, "status_df": None, "module_df": None,
            "warnings": warnings, "errors": ["The Machine SW Log workbook has no readable sheets."],
        }

    # ---------------- Sheet 1: Stoppage (event-level) ----------------
    stoppage_df = None
    sheet1_name = find_sheet(sheets, ["stoppage(event", "stoppage_event"]) or (
        "Stoppage" if "Stoppage" in sheets else find_sheet(sheets, ["stoppage"])
    )
    # avoid grabbing the (Status)/(Module) sheets for sheet 1
    candidate_names = [n for n in sheets.keys() if "status" not in str(n).lower() and "module" not in str(n).lower()]
    if sheet1_name is None or "status" in str(sheet1_name).lower() or "module" in str(sheet1_name).lower():
        sheet1_name = candidate_names[0] if candidate_names else list(sheets.keys())[0]

    try:
        raw1 = sheets[sheet1_name]
        span_col = find_column(raw1, must_contain_any=["span"])
        machine_col = find_column(raw1, must_contain_any=["machine"])
        module_col = find_column(raw1, must_contain_any=["module"])
        status_col = find_column(raw1, must_contain_any=["status"])

        if span_col is None:
            span_col = col_by_letter(raw1, "C")
        if machine_col is None:
            machine_col = col_by_letter(raw1, "F")
        if module_col is None:
            module_col = col_by_letter(raw1, "G")
        if status_col is None:
            status_col = col_by_letter(raw1, "H")

        missing = [n for n, c in [("Span", span_col), ("Machine", machine_col),
                                   ("Module", module_col), ("Status", status_col)] if c is None]
        if missing:
            warnings.append(f"Sheet '{sheet1_name}': could not locate column(s) {', '.join(missing)}.")

        if span_col is not None and status_col is not None:
            stoppage_df = pd.DataFrame({
                "Machine": raw1[machine_col] if machine_col is not None else "N/A",
                "Module": raw1[module_col] if module_col is not None else "N/A",
                "Status": raw1[status_col].astype(str).str.strip(),
                "Duration_Sec": raw1[span_col].apply(parse_hms_to_seconds),
            })
            stoppage_df = stoppage_df.dropna(subset=["Duration_Sec"])
    except Exception as e:
        warnings.append(f"Could not fully parse the event-level Stoppage sheet ('{sheet1_name}'): {e}")

    # ---------------- Sheet 2: Stoppage(Status) (aggregated by status) ----------------
    status_df = None
    sheet2_name = find_sheet(sheets, ["stoppage(status", "stoppage_status", "status"])
    if sheet2_name is not None:
        try:
            raw2 = sheets[sheet2_name]
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
            else:
                warnings.append(f"Sheet '{sheet2_name}': could not locate Status/Total columns.")
        except Exception as e:
            warnings.append(f"Could not fully parse the '{sheet2_name}' sheet: {e}")
    else:
        warnings.append("Could not find a 'Stoppage(Status)' sheet in the workbook.")

    # ---------------- Sheet 3: Stoppage(Module) (aggregated by module) ----------------
    module_df = None
    sheet3_name = find_sheet(sheets, ["stoppage(module", "stoppage_module", "module"])
    if sheet3_name is not None:
        try:
            raw3 = sheets[sheet3_name]
            mod_col = find_column(raw3, must_contain_any=["module"]) or col_by_letter(raw3, "A")
            tot_col = find_column(raw3, must_contain_any=["total"])
            if tot_col is None:
                # take the last column that looks like a duration
                tot_col = raw3.columns[-1]
            if mod_col is not None and tot_col is not None:
                module_df = pd.DataFrame({
                    "Module": raw3[mod_col].astype(str).str.strip(),
                    "Total_Sec": raw3[tot_col].apply(parse_hms_to_seconds),
                })
                module_df = module_df.dropna(subset=["Total_Sec"])
                module_df = module_df[module_df["Module"].str.lower() != "nan"]
        except Exception as e:
            warnings.append(f"Could not fully parse the '{sheet3_name}' sheet: {e}")
    else:
        warnings.append("Could not find a 'Stoppage(Module)' sheet in the workbook (optional).")

    ok = status_df is not None and not status_df.empty
    if not ok:
        errors.append(
            "The Stoppage(Status) sheet could not be parsed — the Pareto chart and Priority "
            "Action Plan need this sheet to work."
        )

    return {
        "ok": ok,
        "stoppage_df": stoppage_df,
        "status_df": status_df,
        "module_df": module_df,
        "warnings": warnings,
        "errors": errors,
    }


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
            "rank": i + 1,
            "priority": priority,
            "status_raw": row["Status"],
            "label": entry["label"],
            "root_cause": entry["root_cause"],
            "action": entry["action"],
            "total_sec": row["Total_Sec"],
            "qty": row.get("Qty", np.nan),
            "share_pct": share,
        })
    return plan


# ----------------------------------------------------------------------------
# SIDEBAR — MANUAL INPUTS & FILE UPLOADS
# ----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## ⚙️ Data Input")
    st.markdown("#### 1. Manual Time Entry")
    line_name = st.text_input("Line / Machine ID", value="SMT Line 1 - Mounter")
    loading_time_h = st.number_input(
        "Loading Time (Hours)", min_value=0.0, value=24.0, step=0.5,
        help="Total scheduled/available production time for the period being analyzed."
    )
    working_time_h = st.number_input(
        "Actual Working Time (Hours)", min_value=0.0, value=18.0, step=0.5,
        help="Actual machine running/production time for the period."
    )

    st.markdown("---")
    st.markdown("#### 2. MES Major Loss File (> 5 min)")
    st.caption("`LossRegisterAnalysis_YYYYMMDD.csv` or `.xlsx` — durations in column K/L, remarks in column M.")
    major_loss_file = st.file_uploader(
        "Upload Major Loss Register", type=["csv", "xlsx", "xls"], key="major_loss_upl"
    )
    unit_override = st.selectbox(
        "If unit can't be auto-detected, assume:",
        options=["Minutes", "Seconds"], index=0,
        help="Only used as a fallback when the duration column can't be identified by its header name."
    )

    st.markdown("---")
    st.markdown("#### 3. Machine SW Stoppage Log")
    st.caption("`Stoppage@...xlsx` — multi-sheet workbook: Stoppage, Stoppage(Status), Stoppage(Module).")
    sw_log_file = st.file_uploader(
        "Upload Machine SW Log Workbook", type=["xlsx", "xls"], key="sw_log_upl"
    )

    st.markdown("---")
    top_n_pareto = st.slider("Pareto: number of stoppage categories to show", 3, 15, 8)


# ----------------------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------------------
st.markdown(
    f"""
    <div class="app-header">
        <h1>🏭 SMT Mounter Efficiency &amp; Small-Loss Analyzer</h1>
        <p>{line_name} &nbsp;•&nbsp; Small Loss = Loading Time − (Actual Working Time + Major Loss &gt; 5 min)</p>
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
        major_loss_result = {
            "ok": False, "df": None, "total_minutes": np.nan,
            "warnings": [], "errors": [f"Unexpected error while parsing Major Loss file: {e}"],
        }

if sw_log_file is not None:
    try:
        sw_log_result = parse_sw_log_file(sw_log_file)
    except Exception as e:
        sw_log_result = {
            "ok": False, "stoppage_df": None, "status_df": None, "module_df": None,
            "warnings": [], "errors": [f"Unexpected error while parsing SW Log workbook: {e}"],
        }

# Surface warnings / errors from file parsing
for label, result in [("Major Loss file", major_loss_result), ("Machine SW Log file", sw_log_result)]:
    if result is None:
        continue
    for err in result.get("errors", []):
        st.error(f"**{label}:** {err}")
    for warn in result.get("warnings", []):
        st.warning(f"**{label}:** {warn}")

# ----------------------------------------------------------------------------
# CORE CALCULATIONS
# ----------------------------------------------------------------------------
major_loss_h = 0.0
major_loss_source = "No file uploaded — treated as 0 h"
if major_loss_result is not None and major_loss_result.get("ok"):
    major_loss_h = major_loss_result["total_minutes"] / 60.0
    major_loss_source = f"From uploaded file ({len(major_loss_result['df'])} loss events)"

# guard against nonsensical manual entries
loading_time_h = max(loading_time_h, 0.0)
working_time_h = max(working_time_h, 0.0)

small_loss_h = loading_time_h - (working_time_h + major_loss_h)
small_loss_h_display = small_loss_h  # keep sign for warnings, clip only for chart

data_issue = None
if loading_time_h == 0:
    data_issue = "Loading Time is 0 — enter a valid Loading Time in the sidebar to see results."
elif working_time_h + major_loss_h > loading_time_h:
    data_issue = (
        "Actual Working Time + Major Loss exceeds Loading Time. Small Loss is negative — "
        "please double-check your manual entries and the Major Loss file totals."
    )

efficiency_pct = safe_div(working_time_h, loading_time_h) * 100 if loading_time_h else np.nan
major_loss_pct = safe_div(major_loss_h, loading_time_h) * 100 if loading_time_h else np.nan
small_loss_pct = safe_div(small_loss_h, loading_time_h) * 100 if loading_time_h else np.nan

# ----------------------------------------------------------------------------
# KPI CARDS
# ----------------------------------------------------------------------------
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
    </div>
    """, unsafe_allow_html=True)

with k2:
    st.markdown(f"""
    <div class="kpi-card kpi-info">
        <div class="kpi-label">Loading Time</div>
        <div class="kpi-value">{fmt_hours(loading_time_h)}</div>
        <div class="kpi-sub">Total available time</div>
    </div>
    """, unsafe_allow_html=True)

with k3:
    st.markdown(f"""
    <div class="kpi-card kpi-warn">
        <div class="kpi-label">Major Loss (&gt; 5 min)</div>
        <div class="kpi-value">{fmt_hours(major_loss_h)}</div>
        <div class="kpi-sub">{fmt_pct(major_loss_pct)} of Loading Time</div>
    </div>
    """, unsafe_allow_html=True)

with k4:
    small_loss_class = "kpi-bad" if (not np.isnan(small_loss_pct) and small_loss_pct > 15) else "kpi-warn"
    st.markdown(f"""
    <div class="kpi-card {small_loss_class}">
        <div class="kpi-label">Small Loss (Hidden, &lt; 5 min)</div>
        <div class="kpi-value">{fmt_hours(small_loss_h)}</div>
        <div class="kpi-sub">{fmt_pct(small_loss_pct)} of Loading Time</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown(f"<p class='small-note'>Major Loss source: {major_loss_source}</p>", unsafe_allow_html=True)

if data_issue:
    st.warning(data_issue)

st.markdown("<hr class='divider'>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# TIME DISTRIBUTION DONUT + PARETO
# ----------------------------------------------------------------------------
col_donut, col_pareto = st.columns([1, 1.4])

with col_donut:
    st.markdown("""
    <div class="section-card">
        <div class="section-title">⏱️ Time Distribution</div>
        <div class="section-caption">Share of Loading Time by category</div>
    """, unsafe_allow_html=True)

    donut_working = max(working_time_h, 0)
    donut_major = max(major_loss_h, 0)
    donut_small = max(small_loss_h, 0)

    if loading_time_h > 0 and (donut_working + donut_major + donut_small) > 0:
        fig_donut = go.Figure(data=[go.Pie(
            labels=["Actual Working Time", "Major Loss (>5m)", "Small Loss (Hidden)"],
            values=[donut_working, donut_major, donut_small],
            hole=0.62,
            marker=dict(colors=["#3ddc84", "#ffb547", "#ff5c5c"],
                        line=dict(color="#0b0f14", width=2)),
            textinfo="percent",
            textfont=dict(color="#eef3f8", size=13),
            sort=False,
        )])
        fig_donut.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#a9b7c6"),
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.18, x=0.5, xanchor="center"),
            margin=dict(t=10, b=10, l=10, r=10),
            height=340,
            annotations=[dict(
                text=f"{fmt_pct(efficiency_pct)}<br><span style='font-size:11px;color:#71828f'>Efficiency</span>",
                x=0.5, y=0.5, showarrow=False, font=dict(size=22, color="#eef3f8")
            )],
        )
        st.plotly_chart(fig_donut, use_container_width=True)
    else:
        st.info("Enter a valid Loading Time (and Working Time) in the sidebar to see the distribution chart.")

    st.markdown("</div>", unsafe_allow_html=True)

with col_pareto:
    st.markdown("""
    <div class="section-card">
        <div class="section-title">📊 Machine SW Stoppage Pareto</div>
        <div class="section-caption">Total stoppage duration by status, sorted descending — from Stoppage(Status) sheet</div>
    """, unsafe_allow_html=True)

    status_df = sw_log_result["status_df"] if (sw_log_result and sw_log_result.get("status_df") is not None) else None

    if status_df is not None and not status_df.empty:
        ranked = status_df.sort_values("Total_Sec", ascending=False).head(top_n_pareto).copy()
        ranked["Total_Hr"] = ranked["Total_Sec"] / 3600.0
        ranked["Cumulative_Pct"] = ranked["Total_Sec"].cumsum() / status_df["Total_Sec"].sum() * 100

        fig_pareto = go.Figure()
        fig_pareto.add_trace(go.Bar(
            x=ranked["Status"], y=ranked["Total_Hr"],
            name="Total Duration (h)",
            marker_color="#3ea6ff",
            text=[f"{v:,.1f}h" for v in ranked["Total_Hr"]],
            textposition="outside",
            textfont=dict(color="#eef3f8"),
        ))
        fig_pareto.add_trace(go.Scatter(
            x=ranked["Status"], y=ranked["Cumulative_Pct"],
            name="Cumulative %",
            yaxis="y2",
            mode="lines+markers",
            line=dict(color="#ffb547", width=2),
            marker=dict(size=6),
        ))
        fig_pareto.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#a9b7c6"),
            margin=dict(t=10, b=10, l=10, r=10),
            height=340,
            xaxis=dict(tickangle=-25, gridcolor="#1c2734"),
            yaxis=dict(title="Hours", gridcolor="#1c2734"),
            yaxis2=dict(title="Cumulative %", overlaying="y", side="right", range=[0, 105], showgrid=False),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, xanchor="left"),
        )
        st.plotly_chart(fig_pareto, use_container_width=True)
    else:
        st.info("Upload the Machine SW Log workbook (with a 'Stoppage(Status)' sheet) to see the Pareto chart.")

    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<hr class='divider'>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# MODULE BREAKDOWN (optional, if available)
# ----------------------------------------------------------------------------
module_df = sw_log_result["module_df"] if (sw_log_result and sw_log_result.get("module_df") is not None) else None
if module_df is not None and not module_df.empty:
    st.markdown("""
    <div class="section-card">
        <div class="section-title">🧩 Module-Level Stoppage Breakdown</div>
        <div class="section-caption">From Stoppage(Module) sheet</div>
    """, unsafe_allow_html=True)

    mod_ranked = module_df.sort_values("Total_Sec", ascending=False).copy()
    mod_ranked["Total_Hr"] = mod_ranked["Total_Sec"] / 3600.0

    fig_mod = go.Figure(go.Bar(
        x=mod_ranked["Total_Hr"], y=mod_ranked["Module"], orientation="h",
        marker_color="#00e0c6",
        text=[f"{v:,.1f}h" for v in mod_ranked["Total_Hr"]],
        textposition="outside",
        textfont=dict(color="#eef3f8"),
    ))
    fig_mod.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#a9b7c6"),
        margin=dict(t=10, b=10, l=10, r=10),
        height=max(280, 32 * len(mod_ranked)),
        xaxis=dict(title="Hours", gridcolor="#1c2734"),
        yaxis=dict(autorange="reversed"),
    )
    st.plotly_chart(fig_mod, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# AUTOMATED PRIORITY ACTION PLAN
# ----------------------------------------------------------------------------
st.markdown("""
<div class="section-card">
    <div class="section-title">🚨 Automated Priority Action Plan</div>
    <div class="section-caption">
        Ranked by total stoppage duration from the Machine SW log — mapped to the standard
        Task Force response playbook. Because the SW log captures every stoppage regardless
        of length, this is the primary tool for tracing the hidden Small Loss bucket back to
        a root cause.
    </div>
""", unsafe_allow_html=True)

if status_df is not None and not status_df.empty:
    plan = build_action_plan(status_df, top_n=top_n_pareto)
    for item in plan:
        badge_class = "badge-danger" if item["priority"] == 1 else ("badge-warn" if item["priority"] == 2 else "badge-info")
        priority_label = "TOP PRIORITY" if item["priority"] == 1 else (f"PRIORITY {item['priority']}")
        qty_txt = f"{int(item['qty'])} events" if pd.notna(item["qty"]) else "qty n/a"
        st.markdown(f"""
        <div class="action-card action-priority-{item['priority']}">
            <span class="badge {badge_class}">{priority_label}</span>
            <span class="action-title">#{item['rank']} — {item['status_raw']} → {item['root_cause']}</span>
            <div class="action-meta">
                Total: {seconds_to_hms(item['total_sec'])} ({item['total_sec']/3600:,.2f} h) &nbsp;•&nbsp;
                {qty_txt} &nbsp;•&nbsp; {fmt_pct(item['share_pct'])} of total SW stoppage time
            </div>
            <div class="action-body">{item['action']}</div>
        </div>
        """, unsafe_allow_html=True)
else:
    st.info(
        "Upload the Machine SW Log workbook to generate the automated action plan. "
        "The engine needs the 'Stoppage(Status)' sheet to rank root causes."
    )

st.markdown("</div>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# DETAIL TABLES (expandable, for audit / traceability)
# ----------------------------------------------------------------------------
with st.expander("🔍 View Raw / Parsed Data Tables"):
    tabs = st.tabs(["Major Loss Events", "SW Stoppage (Status) Summary", "SW Stoppage (Event Log)"])

    with tabs[0]:
        if major_loss_result and major_loss_result.get("ok"):
            df_show = major_loss_result["df"].copy()
            df_show["Duration_Hr"] = df_show["Duration_Min"] / 60.0
            st.dataframe(df_show, use_container_width=True)
        else:
            st.caption("No Major Loss data parsed yet.")

    with tabs[1]:
        if status_df is not None and not status_df.empty:
            df_show = status_df.copy()
            df_show["Total_Hr"] = df_show["Total_Sec"] / 3600.0
            df_show["Min_HMS"] = df_show["Min_Sec"].apply(seconds_to_hms)
            df_show["Max_HMS"] = df_show["Max_Sec"].apply(seconds_to_hms)
            df_show["Total_HMS"] = df_show["Total_Sec"].apply(seconds_to_hms)
            st.dataframe(
                df_show[["Status", "Qty", "Min_HMS", "Max_HMS", "Total_HMS", "Total_Hr"]],
                use_container_width=True,
            )
        else:
            st.caption("No Stoppage(Status) data parsed yet.")

    with tabs[2]:
        stoppage_df = sw_log_result["stoppage_df"] if (sw_log_result and sw_log_result.get("stoppage_df") is not None) else None
        if stoppage_df is not None and not stoppage_df.empty:
            df_show = stoppage_df.copy()
            df_show["Duration_HMS"] = df_show["Duration_Sec"].apply(seconds_to_hms)
            st.dataframe(df_show, use_container_width=True, height=320)
        else:
            st.caption("No event-level Stoppage data parsed yet.")

st.markdown(
    "<p class='small-note'>Built for SMT production efficiency analysis · "
    "All calculations are performed locally in this session — no data is stored.</p>",
    unsafe_allow_html=True,
)
import io
import re
import traceback
from datetime import timedelta

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ----------------------------------------------------------------------------
# PAGE CONFIG
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="SMT Mounter Efficiency & Small Loss Analyzer",
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
}
.app-header h1{
    margin:0;
    font-size: 1.65rem;
}
.app-header p{
    margin: 4px 0 0 0;
    color: var(--text-1);
    font-size: 0.92rem;
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
    font-size: 1.9rem;
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
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# CONSTANTS
# ----------------------------------------------------------------------------
MAJOR_LOSS_THRESHOLD_MIN = 5.0  # minutes - anything above this is a "Major Loss"

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
def fmt_hours(h):
    if h is None or (isinstance(h, float) and np.isnan(h)):
        return "—"
    return f"{h:,.2f} h"


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
    cols = list(df.columns)
    for c in cols:
        name = str(c).strip().lower()
        ok_all = True
        if must_contain_all:
            ok_all = all(k in name for k in must_contain_all)
        ok_any = True
        if must_contain_any:
            ok_any = any(k in name for k in must_contain_any)
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

    # datetime.time or pandas Timestamp-like (openpyxl sometimes returns time objects)
    if hasattr(value, "hour") and hasattr(value, "minute") and hasattr(value, "second"):
        try:
            return value.hour * 3600 + value.minute * 60 + value.second
        except Exception:
            pass

    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none", "nat"):
        return np.nan

    # HH:MM:SS or MM:SS or H:MM:SS.sss
    match = re.match(r"^(\d+):(\d{1,2}):(\d{1,2}(?:\.\d+)?)$", s)
    if match:
        h, m, sec = match.groups()
        return int(h) * 3600 + int(m) * 60 + float(sec)

    match = re.match(r"^(\d{1,2}):(\d{1,2}(?:\.\d+)?)$", s)
    if match:
        m, sec = match.groups()
        return int(m) * 60 + float(sec)

    # plain number as string (e.g. "125.5")
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


def read_any_table(uploaded_file, sheet_name=0):
    """Read a CSV or Excel file into a DataFrame, handling common quirks."""
    name = uploaded_file.name.lower()
    uploaded_file.seek(0)
    if name.endswith(".csv"):
        # try a couple of encodings/separators defensively
        for enc in ("utf-8-sig", "utf-8", "latin1"):
            try:
                uploaded_file.seek(0)
                return pd.read_csv(uploaded_file, encoding=enc)
            except Exception:
                continue
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file, engine="python")
    else:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, sheet_name=sheet_name, engine="openpyxl")


def read_all_sheets(uploaded_file):
    uploaded_file.seek(0)
    return pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")


def find_sheet(sheets_dict, keywords):
    """Find the best-matching sheet name for a list of lower-case keyword hints, in order."""
    names = list(sheets_dict.keys())
    lower_map = {n: str(n).strip().lower() for n in names}
    for kw in keywords:
        for n in names:
            if kw in lower_map[n]:
                return n
    return None


# ----------------------------------------------------------------------------
# MES MAJOR LOSS FILE PARSER  (LossRegisterAnalysis_YYYYMMDD.csv/.xlsx)
# ----------------------------------------------------------------------------
def parse_major_loss_file(uploaded_file, manual_unit_hint=None):
    """
    Returns:
        result dict with:
            ok (bool)
            df (DataFrame) - cleaned [Duration_Min, Remarks]
            total_minutes (float)
            warnings (list[str])
            errors (list[str])
    """
    warnings, errors = [], []
    try:
        raw = read_any_table(uploaded_file)
    except Exception as e:
        return {
            "ok": False, "df": None, "total_minutes": np.nan,
            "warnings": warnings, "errors": [f"Could not read the Major Loss file: {e}"],
        }

    if raw is None or raw.empty:
        return {
            "ok": False, "df": None, "total_minutes": np.nan,
            "warnings": warnings, "errors": ["The Major Loss file appears to be empty."],
        }

    # ---- locate the duration column (Column K = minutes, Column L = seconds by spec) ----
    min_col = find_column(raw, must_contain_all=["loss"], must_contain_any=["min"])
    sec_col = find_column(raw, must_contain_all=["loss"], must_contain_any=["sec"])

    duration_col, unit = None, None
    if min_col is not None:
        duration_col, unit = min_col, "min"
    elif sec_col is not None:
        duration_col, unit = sec_col, "sec"
    else:
        # positional fallback: K = index 10 (minutes), L = index 11 (seconds)
        k_col = col_by_letter(raw, "K")
        l_col = col_by_letter(raw, "L")
        if k_col is not None and raw[k_col].apply(lambda v: parse_hms_to_seconds(v)).notna().sum() > 0:
            duration_col, unit = k_col, (manual_unit_hint or "min")
            warnings.append(
                f"Loss duration column not found by header name — falling back to column K "
                f"('{k_col}') by position, assumed unit = {unit}."
            )
        elif l_col is not None and raw[l_col].apply(lambda v: parse_hms_to_seconds(v)).notna().sum() > 0:
            duration_col, unit = l_col, (manual_unit_hint or "sec")
            warnings.append(
                f"Loss duration column not found by header name — falling back to column L "
                f"('{l_col}') by position, assumed unit = {unit}."
            )

    if duration_col is None:
        return {
            "ok": False, "df": None, "total_minutes": np.nan,
            "warnings": warnings,
            "errors": [
                "Could not find a 'Loss Time (Min/Sec)' column (expected around column K or L). "
                "Please check the file layout."
            ],
        }

    # ---- locate the remarks column (Column M by spec) ----
    remarks_col = find_column(raw, must_contain_any=["remark", "categor", "reason"])
    if remarks_col is None:
        remarks_col = col_by_letter(raw, "M")
        if remarks_col is not None:
            warnings.append(f"Remarks column not found by header name — using column M ('{remarks_col}') by position.")

    seconds_series = raw[duration_col].apply(parse_hms_to_seconds)
    n_bad = seconds_series.isna().sum()
    if n_bad > 0:
        warnings.append(f"{n_bad} row(s) in the loss duration column could not be parsed and were ignored.")

    minutes_series = seconds_series / 60.0 if unit == "sec" else seconds_series

    clean = pd.DataFrame({
        "Duration_Min": minutes_series,
    })
    clean["Remarks"] = raw[remarks_col].astype(str).str.strip() if remarks_col is not None else "Unclassified"
    clean = clean.dropna(subset=["Duration_Min"])
    clean = clean[clean["Duration_Min"] >= 0]

    if clean.empty:
        return {
            "ok": False, "df": None, "total_minutes": np.nan,
            "warnings": warnings, "errors": ["No valid loss-duration rows were found after parsing."],
        }

    # Sanity check vs 5-minute definition (informational only, not a hard filter,
    # since the source file is expected to already be the >5 min register).
    below_threshold = (clean["Duration_Min"] < MAJOR_LOSS_THRESHOLD_MIN).sum()
    if below_threshold > 0:
        warnings.append(
            f"{below_threshold} row(s) in the Major Loss file are below the {MAJOR_LOSS_THRESHOLD_MIN:.0f}-minute "
            "threshold — they were still included in the Major Loss total as provided by the file."
        )

    total_minutes = float(clean["Duration_Min"].sum())
    return {"ok": True, "df": clean, "total_minutes": total_minutes, "warnings": warnings, "errors": errors}


# ----------------------------------------------------------------------------
# MACHINE SW STOPPAGE LOG PARSER (multi-sheet Stoppage@....xlsx)
# ----------------------------------------------------------------------------
def parse_sw_log_file(uploaded_file):
    warnings, errors = [], []
    try:
        sheets = read_all_sheets(uploaded_file)
    except Exception as e:
        return {
            "ok": False, "stoppage_df": None, "status_df": None, "module_df": None,
            "warnings": warnings, "errors": [f"Could not read the Machine SW Log workbook: {e}"],
        }

    if not sheets:
        return {
            "ok": False, "stoppage_df": None, "status_df": None, "module_df": None,
            "warnings": warnings, "errors": ["The Machine SW Log workbook has no readable sheets."],
        }

    # ---------------- Sheet 1: Stoppage (event-level) ----------------
    stoppage_df = None
    sheet1_name = find_sheet(sheets, ["stoppage(event", "stoppage_event"]) or (
        "Stoppage" if "Stoppage" in sheets else find_sheet(sheets, ["stoppage"])
    )
    # avoid grabbing the (Status)/(Module) sheets for sheet 1
    candidate_names = [n for n in sheets.keys() if "status" not in str(n).lower() and "module" not in str(n).lower()]
    if sheet1_name is None or "status" in str(sheet1_name).lower() or "module" in str(sheet1_name).lower():
        sheet1_name = candidate_names[0] if candidate_names else list(sheets.keys())[0]

    try:
        raw1 = sheets[sheet1_name]
        span_col = find_column(raw1, must_contain_any=["span"])
        machine_col = find_column(raw1, must_contain_any=["machine"])
        module_col = find_column(raw1, must_contain_any=["module"])
        status_col = find_column(raw1, must_contain_any=["status"])

        if span_col is None:
            span_col = col_by_letter(raw1, "C")
        if machine_col is None:
            machine_col = col_by_letter(raw1, "F")
        if module_col is None:
            module_col = col_by_letter(raw1, "G")
        if status_col is None:
            status_col = col_by_letter(raw1, "H")

        missing = [n for n, c in [("Span", span_col), ("Machine", machine_col),
                                   ("Module", module_col), ("Status", status_col)] if c is None]
        if missing:
            warnings.append(f"Sheet '{sheet1_name}': could not locate column(s) {', '.join(missing)}.")

        if span_col is not None and status_col is not None:
            stoppage_df = pd.DataFrame({
                "Machine": raw1[machine_col] if machine_col is not None else "N/A",
                "Module": raw1[module_col] if module_col is not None else "N/A",
                "Status": raw1[status_col].astype(str).str.strip(),
                "Duration_Sec": raw1[span_col].apply(parse_hms_to_seconds),
            })
            stoppage_df = stoppage_df.dropna(subset=["Duration_Sec"])
    except Exception as e:
        warnings.append(f"Could not fully parse the event-level Stoppage sheet ('{sheet1_name}'): {e}")

    # ---------------- Sheet 2: Stoppage(Status) (aggregated by status) ----------------
    status_df = None
    sheet2_name = find_sheet(sheets, ["stoppage(status", "stoppage_status", "status"])
    if sheet2_name is not None:
        try:
            raw2 = sheets[sheet2_name]
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
            else:
                warnings.append(f"Sheet '{sheet2_name}': could not locate Status/Total columns.")
        except Exception as e:
            warnings.append(f"Could not fully parse the '{sheet2_name}' sheet: {e}")
    else:
        warnings.append("Could not find a 'Stoppage(Status)' sheet in the workbook.")

    # ---------------- Sheet 3: Stoppage(Module) (aggregated by module) ----------------
    module_df = None
    sheet3_name = find_sheet(sheets, ["stoppage(module", "stoppage_module", "module"])
    if sheet3_name is not None:
        try:
            raw3 = sheets[sheet3_name]
            mod_col = find_column(raw3, must_contain_any=["module"]) or col_by_letter(raw3, "A")
            tot_col = find_column(raw3, must_contain_any=["total"])
            if tot_col is None:
                # take the last column that looks like a duration
                tot_col = raw3.columns[-1]
            if mod_col is not None and tot_col is not None:
                module_df = pd.DataFrame({
                    "Module": raw3[mod_col].astype(str).str.strip(),
                    "Total_Sec": raw3[tot_col].apply(parse_hms_to_seconds),
                })
                module_df = module_df.dropna(subset=["Total_Sec"])
                module_df = module_df[module_df["Module"].str.lower() != "nan"]
        except Exception as e:
            warnings.append(f"Could not fully parse the '{sheet3_name}' sheet: {e}")
    else:
        warnings.append("Could not find a 'Stoppage(Module)' sheet in the workbook (optional).")

    ok = status_df is not None and not status_df.empty
    if not ok:
        errors.append(
            "The Stoppage(Status) sheet could not be parsed — the Pareto chart and Priority "
            "Action Plan need this sheet to work."
        )

    return {
        "ok": ok,
        "stoppage_df": stoppage_df,
        "status_df": status_df,
        "module_df": module_df,
        "warnings": warnings,
        "errors": errors,
    }


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
            "rank": i + 1,
            "priority": priority,
            "status_raw": row["Status"],
            "label": entry["label"],
            "root_cause": entry["root_cause"],
            "action": entry["action"],
            "total_sec": row["Total_Sec"],
            "qty": row.get("Qty", np.nan),
            "share_pct": share,
        })
    return plan


# ----------------------------------------------------------------------------
# SIDEBAR — MANUAL INPUTS & FILE UPLOADS
# ----------------------------------------------------------------------------
with st.sidebar:
    st.markdown("## ⚙️ Data Input")
    st.markdown("#### 1. Manual Time Entry")
    line_name = st.text_input("Line / Machine ID", value="SMT Line 1 - Mounter")
    loading_time_h = st.number_input(
        "Loading Time (Hours)", min_value=0.0, value=24.0, step=0.5,
        help="Total scheduled/available production time for the period being analyzed."
    )
    working_time_h = st.number_input(
        "Actual Working Time (Hours)", min_value=0.0, value=18.0, step=0.5,
        help="Actual machine running/production time for the period."
    )

    st.markdown("---")
    st.markdown("#### 2. MES Major Loss File (> 5 min)")
    st.caption("`LossRegisterAnalysis_YYYYMMDD.csv` or `.xlsx` — durations in column K/L, remarks in column M.")
    major_loss_file = st.file_uploader(
        "Upload Major Loss Register", type=["csv", "xlsx", "xls"], key="major_loss_upl"
    )
    unit_override = st.selectbox(
        "If unit can't be auto-detected, assume:",
        options=["Minutes", "Seconds"], index=0,
        help="Only used as a fallback when the duration column can't be identified by its header name."
    )

    st.markdown("---")
    st.markdown("#### 3. Machine SW Stoppage Log")
    st.caption("`Stoppage@...xlsx` — multi-sheet workbook: Stoppage, Stoppage(Status), Stoppage(Module).")
    sw_log_file = st.file_uploader(
        "Upload Machine SW Log Workbook", type=["xlsx", "xls"], key="sw_log_upl"
    )

    st.markdown("---")
    top_n_pareto = st.slider("Pareto: number of stoppage categories to show", 3, 15, 8)


# ----------------------------------------------------------------------------
# HEADER
# ----------------------------------------------------------------------------
st.markdown(
    f"""
    <div class="app-header">
        <h1>🏭 SMT Mounter Efficiency &amp; Small-Loss Analyzer</h1>
        <p>{line_name} &nbsp;•&nbsp; Small Loss = Loading Time − (Actual Working Time + Major Loss &gt; 5 min)</p>
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
        major_loss_result = {
            "ok": False, "df": None, "total_minutes": np.nan,
            "warnings": [], "errors": [f"Unexpected error while parsing Major Loss file: {e}"],
        }

if sw_log_file is not None:
    try:
        sw_log_result = parse_sw_log_file(sw_log_file)
    except Exception as e:
        sw_log_result = {
            "ok": False, "stoppage_df": None, "status_df": None, "module_df": None,
            "warnings": [], "errors": [f"Unexpected error while parsing SW Log workbook: {e}"],
        }

# Surface warnings / errors from file parsing
for label, result in [("Major Loss file", major_loss_result), ("Machine SW Log file", sw_log_result)]:
    if result is None:
        continue
    for err in result.get("errors", []):
        st.error(f"**{label}:** {err}")
    for warn in result.get("warnings", []):
        st.warning(f"**{label}:** {warn}")

# ----------------------------------------------------------------------------
# CORE CALCULATIONS
# ----------------------------------------------------------------------------
major_loss_h = 0.0
major_loss_source = "No file uploaded — treated as 0 h"
if major_loss_result is not None and major_loss_result.get("ok"):
    major_loss_h = major_loss_result["total_minutes"] / 60.0
    major_loss_source = f"From uploaded file ({len(major_loss_result['df'])} loss events)"

# guard against nonsensical manual entries
loading_time_h = max(loading_time_h, 0.0)
working_time_h = max(working_time_h, 0.0)

small_loss_h = loading_time_h - (working_time_h + major_loss_h)
small_loss_h_display = small_loss_h  # keep sign for warnings, clip only for chart

data_issue = None
if loading_time_h == 0:
    data_issue = "Loading Time is 0 — enter a valid Loading Time in the sidebar to see results."
elif working_time_h + major_loss_h > loading_time_h:
    data_issue = (
        "Actual Working Time + Major Loss exceeds Loading Time. Small Loss is negative — "
        "please double-check your manual entries and the Major Loss file totals."
    )

efficiency_pct = safe_div(working_time_h, loading_time_h) * 100 if loading_time_h else np.nan
major_loss_pct = safe_div(major_loss_h, loading_time_h) * 100 if loading_time_h else np.nan
small_loss_pct = safe_div(small_loss_h, loading_time_h) * 100 if loading_time_h else np.nan

# ----------------------------------------------------------------------------
# KPI CARDS
# ----------------------------------------------------------------------------
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
    </div>
    """, unsafe_allow_html=True)

with k2:
    st.markdown(f"""
    <div class="kpi-card kpi-info">
        <div class="kpi-label">Loading Time</div>
        <div class="kpi-value">{fmt_hours(loading_time_h)}</div>
        <div class="kpi-sub">Total available time</div>
    </div>
    """, unsafe_allow_html=True)

with k3:
    st.markdown(f"""
    <div class="kpi-card kpi-warn">
        <div class="kpi-label">Major Loss (&gt; 5 min)</div>
        <div class="kpi-value">{fmt_hours(major_loss_h)}</div>
        <div class="kpi-sub">{fmt_pct(major_loss_pct)} of Loading Time</div>
    </div>
    """, unsafe_allow_html=True)

with k4:
    small_loss_class = "kpi-bad" if (not np.isnan(small_loss_pct) and small_loss_pct > 15) else "kpi-warn"
    st.markdown(f"""
    <div class="kpi-card {small_loss_class}">
        <div class="kpi-label">Small Loss (Hidden, &lt; 5 min)</div>
        <div class="kpi-value">{fmt_hours(small_loss_h)}</div>
        <div class="kpi-sub">{fmt_pct(small_loss_pct)} of Loading Time</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown(f"<p class='small-note'>Major Loss source: {major_loss_source}</p>", unsafe_allow_html=True)

if data_issue:
    st.warning(data_issue)

st.markdown("<hr class='divider'>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# TIME DISTRIBUTION DONUT + PARETO
# ----------------------------------------------------------------------------
col_donut, col_pareto = st.columns([1, 1.4])

with col_donut:
    st.markdown("""
    <div class="section-card">
        <div class="section-title">⏱️ Time Distribution</div>
        <div class="section-caption">Share of Loading Time by category</div>
    """, unsafe_allow_html=True)

    donut_working = max(working_time_h, 0)
    donut_major = max(major_loss_h, 0)
    donut_small = max(small_loss_h, 0)

    if loading_time_h > 0 and (donut_working + donut_major + donut_small) > 0:
        fig_donut = go.Figure(data=[go.Pie(
            labels=["Actual Working Time", "Major Loss (>5m)", "Small Loss (Hidden)"],
            values=[donut_working, donut_major, donut_small],
            hole=0.62,
            marker=dict(colors=["#3ddc84", "#ffb547", "#ff5c5c"],
                        line=dict(color="#0b0f14", width=2)),
            textinfo="percent",
            textfont=dict(color="#eef3f8", size=13),
            sort=False,
        )])
        fig_donut.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#a9b7c6"),
            showlegend=True,
            legend=dict(orientation="h", yanchor="bottom", y=-0.18, x=0.5, xanchor="center"),
            margin=dict(t=10, b=10, l=10, r=10),
            height=340,
            annotations=[dict(
                text=f"{fmt_pct(efficiency_pct)}<br><span style='font-size:11px;color:#71828f'>Efficiency</span>",
                x=0.5, y=0.5, showarrow=False, font=dict(size=22, color="#eef3f8")
            )],
        )
        st.plotly_chart(fig_donut, use_container_width=True)
    else:
        st.info("Enter a valid Loading Time (and Working Time) in the sidebar to see the distribution chart.")

    st.markdown("</div>", unsafe_allow_html=True)

with col_pareto:
    st.markdown("""
    <div class="section-card">
        <div class="section-title">📊 Machine SW Stoppage Pareto</div>
        <div class="section-caption">Total stoppage duration by status, sorted descending — from Stoppage(Status) sheet</div>
    """, unsafe_allow_html=True)

    status_df = sw_log_result["status_df"] if (sw_log_result and sw_log_result.get("status_df") is not None) else None

    if status_df is not None and not status_df.empty:
        ranked = status_df.sort_values("Total_Sec", ascending=False).head(top_n_pareto).copy()
        ranked["Total_Hr"] = ranked["Total_Sec"] / 3600.0
        ranked["Cumulative_Pct"] = ranked["Total_Sec"].cumsum() / status_df["Total_Sec"].sum() * 100

        fig_pareto = go.Figure()
        fig_pareto.add_trace(go.Bar(
            x=ranked["Status"], y=ranked["Total_Hr"],
            name="Total Duration (h)",
            marker_color="#3ea6ff",
            text=[f"{v:,.1f}h" for v in ranked["Total_Hr"]],
            textposition="outside",
            textfont=dict(color="#eef3f8"),
        ))
        fig_pareto.add_trace(go.Scatter(
            x=ranked["Status"], y=ranked["Cumulative_Pct"],
            name="Cumulative %",
            yaxis="y2",
            mode="lines+markers",
            line=dict(color="#ffb547", width=2),
            marker=dict(size=6),
        ))
        fig_pareto.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            font=dict(color="#a9b7c6"),
            margin=dict(t=10, b=10, l=10, r=10),
            height=340,
            xaxis=dict(tickangle=-25, gridcolor="#1c2734"),
            yaxis=dict(title="Hours", gridcolor="#1c2734"),
            yaxis2=dict(title="Cumulative %", overlaying="y", side="right", range=[0, 105], showgrid=False),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, xanchor="left"),
        )
        st.plotly_chart(fig_pareto, use_container_width=True)
    else:
        st.info("Upload the Machine SW Log workbook (with a 'Stoppage(Status)' sheet) to see the Pareto chart.")

    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<hr class='divider'>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# MODULE BREAKDOWN (optional, if available)
# ----------------------------------------------------------------------------
module_df = sw_log_result["module_df"] if (sw_log_result and sw_log_result.get("module_df") is not None) else None
if module_df is not None and not module_df.empty:
    st.markdown("""
    <div class="section-card">
        <div class="section-title">🧩 Module-Level Stoppage Breakdown</div>
        <div class="section-caption">From Stoppage(Module) sheet</div>
    """, unsafe_allow_html=True)

    mod_ranked = module_df.sort_values("Total_Sec", ascending=False).copy()
    mod_ranked["Total_Hr"] = mod_ranked["Total_Sec"] / 3600.0

    fig_mod = go.Figure(go.Bar(
        x=mod_ranked["Total_Hr"], y=mod_ranked["Module"], orientation="h",
        marker_color="#00e0c6",
        text=[f"{v:,.1f}h" for v in mod_ranked["Total_Hr"]],
        textposition="outside",
        textfont=dict(color="#eef3f8"),
    ))
    fig_mod.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(color="#a9b7c6"),
        margin=dict(t=10, b=10, l=10, r=10),
        height=max(280, 32 * len(mod_ranked)),
        xaxis=dict(title="Hours", gridcolor="#1c2734"),
        yaxis=dict(autorange="reversed"),
    )
    st.plotly_chart(fig_mod, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# AUTOMATED PRIORITY ACTION PLAN
# ----------------------------------------------------------------------------
st.markdown("""
<div class="section-card">
    <div class="section-title">🚨 Automated Priority Action Plan</div>
    <div class="section-caption">
        Ranked by total stoppage duration from the Machine SW log — mapped to the standard
        Task Force response playbook. Because the SW log captures every stoppage regardless
        of length, this is the primary tool for tracing the hidden Small Loss bucket back to
        a root cause.
    </div>
""", unsafe_allow_html=True)

if status_df is not None and not status_df.empty:
    plan = build_action_plan(status_df, top_n=top_n_pareto)
    for item in plan:
        badge_class = "badge-danger" if item["priority"] == 1 else ("badge-warn" if item["priority"] == 2 else "badge-info")
        priority_label = "TOP PRIORITY" if item["priority"] == 1 else (f"PRIORITY {item['priority']}")
        qty_txt = f"{int(item['qty'])} events" if pd.notna(item["qty"]) else "qty n/a"
        st.markdown(f"""
        <div class="action-card action-priority-{item['priority']}">
            <span class="badge {badge_class}">{priority_label}</span>
            <span class="action-title">#{item['rank']} — {item['status_raw']} → {item['root_cause']}</span>
            <div class="action-meta">
                Total: {seconds_to_hms(item['total_sec'])} ({item['total_sec']/3600:,.2f} h) &nbsp;•&nbsp;
                {qty_txt} &nbsp;•&nbsp; {fmt_pct(item['share_pct'])} of total SW stoppage time
            </div>
            <div class="action-body">{item['action']}</div>
        </div>
        """, unsafe_allow_html=True)
else:
    st.info(
        "Upload the Machine SW Log workbook to generate the automated action plan. "
        "The engine needs the 'Stoppage(Status)' sheet to rank root causes."
    )

st.markdown("</div>", unsafe_allow_html=True)

# ----------------------------------------------------------------------------
# DETAIL TABLES (expandable, for audit / traceability)
# ----------------------------------------------------------------------------
with st.expander("🔍 View Raw / Parsed Data Tables"):
    tabs = st.tabs(["Major Loss Events", "SW Stoppage (Status) Summary", "SW Stoppage (Event Log)"])

    with tabs[0]:
        if major_loss_result and major_loss_result.get("ok"):
            df_show = major_loss_result["df"].copy()
            df_show["Duration_Hr"] = df_show["Duration_Min"] / 60.0
            st.dataframe(df_show, use_container_width=True)
        else:
            st.caption("No Major Loss data parsed yet.")

    with tabs[1]:
        if status_df is not None and not status_df.empty:
            df_show = status_df.copy()
            df_show["Total_Hr"] = df_show["Total_Sec"] / 3600.0
            df_show["Min_HMS"] = df_show["Min_Sec"].apply(seconds_to_hms)
            df_show["Max_HMS"] = df_show["Max_Sec"].apply(seconds_to_hms)
            df_show["Total_HMS"] = df_show["Total_Sec"].apply(seconds_to_hms)
            st.dataframe(
                df_show[["Status", "Qty", "Min_HMS", "Max_HMS", "Total_HMS", "Total_Hr"]],
                use_container_width=True,
            )
        else:
            st.caption("No Stoppage(Status) data parsed yet.")

    with tabs[2]:
        stoppage_df = sw_log_result["stoppage_df"] if (sw_log_result and sw_log_result.get("stoppage_df") is not None) else None
        if stoppage_df is not None and not stoppage_df.empty:
            df_show = stoppage_df.copy()
            df_show["Duration_HMS"] = df_show["Duration_Sec"].apply(seconds_to_hms)
            st.dataframe(df_show, use_container_width=True, height=320)
        else:
            st.caption("No event-level Stoppage data parsed yet.")

st.markdown(
    "<p class='small-note'>Built for SMT production efficiency analysis · "
    "All calculations are performed locally in this session — no data is stored.</p>",
    unsafe_allow_html=True,
)

