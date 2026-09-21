"""
SMT Mounter Micro-Loss Analyzer & Task Force Action Tracker — Ver 2.1
==============================================================================
Production-ready single-file Streamlit application.

WHAT'S NEW IN VER 2.1 (on top of every Ver 2.0 feature)
----------------------------------------------------------
1. Auto-launch: on startup the app tries to open your default browser to
   http://localhost:<port> automatically (see "AUTO BROWSER LAUNCH" below).
2. Every log input now has TWO modes: Upload File (.xlsx/.xls/.csv/.txt) OR
   Paste Text (raw tab/comma/semicolon-separated text copied straight out of
   Excel) — both go through the same parsing/validation logic.
3. Column parsing is aligned to the standard Samsung-style MES Loss Register
   (`Occur. Time`, `Shift`, `Line`, `Loss Time(min)`, `Remarks`) and Machine
   SW Stoppage log (`Start`, `End`, `Span`, `Time(min)`, `Machine`, `Module`,
   `Status`, `Detail`), with header-name matching first and positional
   fallbacks second.
4. A Shift selector (sidebar) filters the Major Loss data by shift when a
   Shift column is present, and shift is now saved alongside every daily
   database record for date+shift historical comparisons.
5. A dedicated "📁 Data Input & Log Analysis" tab hosts all upload/paste
   controls, a parsed-data preview, and a small duration-conversion helper.
6. New charts: a Line-by-Line Efficiency (OEE-proxy) bar chart on the Overall
   Plant dashboard, and a Loading→Major Loss→Small Loss→Working waterfall
   chart on each line's dashboard.

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

HOW TO RUN
-----------
