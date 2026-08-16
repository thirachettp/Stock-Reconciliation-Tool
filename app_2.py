import streamlit as st
import pandas as pd
import numpy as np
import json
import io
import re
import uuid
from datetime import datetime, date
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Stock Reconciliation Tool",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ─────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap');

/* ── Theme tokens — Light (default) ── */
:root {
    --bg-app: #f8fafc;
    --bg-card: #ffffff;
    --bg-card-alt: #f1f5f9;
    --border: #e2e8f0;
    --border-strong: #cbd5e1;
    --text-primary: #0f172a;
    --text-secondary: #475569;
    --text-muted: #94a3b8;
    --accent: #3b82f6;
    --accent-text: #ffffff;
    --accent-soft-bg: #eff6ff;
    --code-bg: #f1f5f9;
    --code-border: #bfdbfe;
    --code-text: #2563eb;
    --input-bg: #ffffff;
    --input-border: #cbd5e1;
    --kpi-match: #16a34a;
    --kpi-mismatch: #d97706;
    --kpi-missing: #dc2626;
    --status-match-bg: #dcfce7;    --status-match-text: #15803d;
    --status-mismatch-bg: #fef3c7; --status-mismatch-text: #b45309;
    --status-missing-bg: #fee2e2;  --status-missing-text: #b91c1c;
    --step-done-bg: #16a34a;
    --step-pending-bg: #e2e8f0;
    --step-pending-text: #94a3b8;
}
/* ── Theme tokens — Dark (auto, via OS/browser preference) ── */
@media (prefers-color-scheme: dark) {
    :root {
        --bg-app: #0f1117;
        --bg-card: #1a1f2e;
        --bg-card-alt: #161b28;
        --border: #2d3748;
        --border-strong: #374151;
        --text-primary: #e2e8f0;
        --text-secondary: #94a3b8;
        --text-muted: #64748b;
        --accent: #3b82f6;
        --accent-text: #ffffff;
        --accent-soft-bg: #1e3a5f;
        --code-bg: #0f172a;
        --code-border: #1e3a5f;
        --code-text: #22d3ee;
        --input-bg: #252d3d;
        --input-border: #374151;
        --kpi-match: #22c55e;
        --kpi-mismatch: #f59e0b;
        --kpi-missing: #ef4444;
        --status-match-bg: #14532d;    --status-match-text: #4ade80;
        --status-mismatch-bg: #451a03; --status-mismatch-text: #fbbf24;
        --status-missing-bg: #450a0a;  --status-missing-text: #f87171;
        --step-done-bg: #22c55e;
        --step-pending-bg: #252d3d;
        --step-pending-text: #64748b;
    }
}

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background: var(--bg-app); color: var(--text-primary); }

/* ── Tabs ── */
.stTabs [data-baseweb="tab-list"] {
    background: var(--bg-card); border-radius: 12px;
    padding: 4px; gap: 4px; border: 1px solid var(--border);
}
.stTabs [data-baseweb="tab"] {
    background: transparent; color: var(--text-secondary);
    border-radius: 8px; font-weight: 500; font-size: 14px; padding: 8px 20px;
}
.stTabs [aria-selected="true"] { background: var(--accent) !important; color: var(--accent-text) !important; }
.stTabs [data-baseweb="tab"] p { color: inherit !important; }

/* ── Cards ── */
.card {
    background: var(--bg-card); border: 1px solid var(--border);
    border-radius: 14px; padding: 20px 24px; margin-bottom: 16px;
}
.card-sm {
    background: var(--bg-card); border: 1px solid var(--border);
    border-radius: 10px; padding: 14px 18px; margin-bottom: 12px;
}
.card-label {
    font-size: 11px; font-weight: 700; color: var(--text-muted);
    text-transform: uppercase; letter-spacing: 0.09em; margin-bottom: 14px;
}

/* ── System selector pills ── */
.sys-pill-row { display:flex; gap:8px; margin-bottom:16px; }
.sys-pill {
    padding: 6px 18px; border-radius: 20px; font-size: 13px; font-weight: 600;
    cursor: pointer; border: 1.5px solid transparent; transition: all 0.15s;
}
.sys-FC  { background:var(--accent-soft-bg); color:#3b82f6; border-color:#2563eb; }
.sys-JDA { background:var(--status-match-bg); color:var(--status-match-text); border-color:#16a34a; }
.sys-MHT { background:var(--status-missing-bg); color:#db2777; border-color:#db2777; }

/* ── KPI ── */
.kpi-grid { display:flex; gap:14px; margin-bottom:20px; flex-wrap:wrap; }
.kpi-box {
    flex:1; min-width:130px; background:var(--bg-card); border:1px solid var(--border);
    border-radius:12px; padding:18px 20px; text-align:center;
}
.kpi-value { font-size:30px; font-weight:700; font-family:'JetBrains Mono',monospace; }
.kpi-label { font-size:11px; color:var(--text-muted); margin-top:4px; font-weight:600; text-transform:uppercase; letter-spacing:.06em; }
.kpi-total .kpi-value  { color:var(--text-primary); }
.kpi-match .kpi-value  { color:var(--kpi-match); }
.kpi-mis .kpi-value    { color:var(--kpi-mismatch); }
.kpi-miss .kpi-value   { color:var(--kpi-missing); }

/* ── Formula preview ── */
.fpv {
    font-family:'JetBrains Mono',monospace; font-size:13px; color:var(--code-text);
    background:var(--code-bg); border:1px solid var(--code-border); border-radius:6px;
    padding:7px 12px; min-height:36px; display:flex; align-items:center;
}

/* ── Buttons ── */
.stButton > button {
    border-radius:8px; font-weight:500; font-size:14px;
    border:1px solid var(--border); background:var(--bg-card); color:var(--text-primary); transition:all .15s;
}
.stButton > button:hover { background:var(--bg-card-alt); border-color:var(--accent); color:var(--accent); }

/* ── Section titles ── */
.sec-title { font-size:18px; font-weight:700; color:var(--text-primary); margin-bottom:4px; }
.sec-sub   { font-size:13px; color:var(--text-muted); margin-bottom:20px; }

/* ── Divider ── */
hr.s { border:none; border-top:1px solid var(--border); margin:20px 0; }

/* ── Inputs / Selects ── */
.stSelectbox > div > div, .stMultiSelect > div > div {
    background:var(--input-bg) !important; border:1px solid var(--input-border) !important; border-radius:8px !important;
}
.stTextInput > div > div > input {
    background:var(--input-bg) !important; border:1px solid var(--input-border) !important;
    border-radius:8px !important; color:var(--text-primary) !important;
}
/* Make selectbox text follow theme */
.stSelectbox [data-baseweb="select"] div { color:var(--text-primary) !important; }
.stSelectbox svg { fill:var(--text-secondary) !important; }
.stMultiSelect span { color:var(--text-primary) !important; }

/* ── File uploader ── */
[data-testid="stFileUploader"] {
    background:var(--input-bg); border:2px dashed var(--input-border); border-radius:12px; padding:8px;
}
[data-testid="stFileUploader"]:hover { border-color:var(--accent); }
[data-testid="stFileUploader"] label { color:var(--text-secondary) !important; }

/* ── Expander ── */
details > summary { background:var(--bg-card); border-radius:8px; padding:8px 12px; color:var(--text-primary); }

/* ── Dataframe ── */
[data-testid="stDataFrame"] { border-radius:10px; overflow:hidden; }

/* ── Column header labels ── */
.col-hdr {
    font-size:10px; font-weight:700; color:var(--text-muted);
    text-transform:uppercase; letter-spacing:.08em; padding:4px 0;
}
.col-grp-hdr {
    font-size:11px; font-weight:700; color:var(--accent);
    text-transform:uppercase; letter-spacing:.08em; padding:4px 0 4px 0;
    text-align:center; border-bottom:2px solid var(--border); margin-bottom:6px;
}

/* upload area label hidden */
[data-testid="stFileUploaderDropzone"] p { font-size:13px; color:var(--text-secondary); }

/* ── Step indicator ── */
.step-track { display:flex; align-items:center; margin: 4px 0 24px 0; }
.step-item { display:flex; align-items:center; gap:10px; }
.step-circle {
    width:28px; height:28px; border-radius:50%; display:flex; align-items:center; justify-content:center;
    font-size:13px; font-weight:700; flex-shrink:0;
}
.step-circle.done    { background:var(--step-done-bg); color:#ffffff; }
.step-circle.current { background:var(--accent); color:var(--accent-text); }
.step-circle.pending { background:var(--step-pending-bg); color:var(--step-pending-text); }
.step-label { font-size:13px; font-weight:600; color:var(--text-secondary); white-space:nowrap; }
.step-label.active { color:var(--text-primary); }
.step-line { flex:1; height:2px; background:var(--border); margin:0 12px; min-width:24px; }
.step-line.done { background:var(--step-done-bg); }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────
FUNCTIONS = ["FIRST", "SUM", "MIN", "MAX", "COUNT", "วดป", "ปดว", "ดวป"]

# Three dedicated date-order functions for fields where the source date text
# is genuinely ambiguous (e.g. '03/04/2026'). Unlike FIRST/MIN/MAX (which use
# best-effort auto-detection), picking one of these tells the app exactly how
# to read the value — no guessing:
#   วดป = วัน-เดือน-ปี = Day-Month-Year   (D/M/Y)
#   ปดว = ปี-เดือน-วัน = Year-Month-Day   (Y/M/D)
#   ดวป = เดือน-วัน-ปี = Month-Day-Year   (M/D/Y)
# Each behaves like FIRST (takes the first source column's value) but forces
# that specific day/month/year order regardless of ambiguity.
DATE_ORDER_FUNCTIONS = {"วดป": "dmy", "ปดว": "ymd", "ดวป": "mdy"}

# Pandas caps how many cells its Styler will render (default 262,144) to avoid
# slow browser rendering. Raise the cap so normal-sized reconciliation tables
# don't hit it — a hard ceiling (STYLE_CELL_HARD_LIMIT below) still protects
# against truly huge tables that would make the page unusably slow.
STYLE_CELL_HARD_LIMIT = 2_000_000
try:
    pd.set_option("styler.render.max_elements", STYLE_CELL_HARD_LIMIT)
except Exception:
    pass  # older pandas versions don't have this option; safe to ignore
SYSTEMS   = ["FC", "JDA", "MHT"]

DEFAULT_MAPPINGS = {
    "JDA": {
        "SKU":      {"function": "FIRST", "columns": ["SKU"]},
        "Location": {"function": "FIRST", "columns": ["SLOT"]},
        "Qty":      {"function": "SUM",   "columns": ["SLOT_SOH"]},
        "BBDate":   {"function": "MIN",   "columns": ["EXP_DATE"]},
    },
    "FC": {
        "SKU":      {"function": "FIRST", "columns": ["sku_id"]},
        "Location": {"function": "FIRST", "columns": ["location_id"]},
        "Qty":      {"function": "SUM",   "columns": ["binbalance_qtybal"]},
        "BBDate":   {"function": "MIN",   "columns": ["exp_date"]},
    },
    "MHT": {
        "SKU":      {"function": "FIRST", "columns": ["SKU"]},
        "Location": {"function": "FIRST", "columns": ["Location"]},
        "Qty":      {"function": "SUM",   "columns": ["On Hand"]},
        "BBDate":   {"function": "MIN",   "columns": []},
    },
}

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
def init_state():
    defs = {
        "sys_a": "FC", "sys_b": "JDA",
        "df_a_raw": None, "df_b_raw": None,
        "file_configs_a": {}, "file_configs_b": {},
        "formulas_a": None, "formulas_b": None,
        "formulas_unified": None,
        "std_a": None, "std_b": None,
        "compare_result": None,
        "key_cols": ["SKU", "Location"],
        "compare_fields": [],
        "auto_combine_sig_a": None, "auto_combine_sig_b": None,
    }
    for k, v in defs.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# Auto-init formulas from system defaults (on first load or system change)
def load_default_formulas(system):
    mapping = DEFAULT_MAPPINGS.get(system, {})
    return [
        {"target": t, "function": cfg["function"], "columns": cfg["columns"]}
        for t, cfg in mapping.items()
    ]

def load_default_formulas_unified(sys_a, sys_b):
    """
    Build the unified Formula Builder rows: one row per target field name,
    shared by both File A and File B, each carrying its own function/columns
    per side (since both sides already use the same target-field names by
    design — SKU, Location, Qty, BBDate).
    """
    map_a = DEFAULT_MAPPINGS.get(sys_a, {})
    map_b = DEFAULT_MAPPINGS.get(sys_b, {})
    targets = list(dict.fromkeys(list(map_a.keys()) + list(map_b.keys())))
    rows = []
    for t in targets:
        a_cfg = map_a.get(t, {"function": "FIRST", "columns": []})
        b_cfg = map_b.get(t, {"function": "FIRST", "columns": []})
        rows.append({
            "id": uuid.uuid4().hex[:8],
            "target": t,
            "a": {"function": a_cfg["function"], "columns": list(a_cfg["columns"])},
            "b": {"function": b_cfg["function"], "columns": list(b_cfg["columns"])},
        })
    return rows

def ensure_formulas_unified():
    if st.session_state.get("formulas_unified") is None:
        st.session_state.formulas_unified = load_default_formulas_unified(
            st.session_state.sys_a, st.session_state.sys_b
        )

# ─────────────────────────────────────────────
# CORE HELPERS
# ─────────────────────────────────────────────
def is_excel_file(filename):
    return filename.lower().endswith((".xlsx", ".xls"))

@st.cache_data(show_spinner=False)
def get_excel_sheets_cached(file_bytes):
    """Return list of sheet names in an Excel file (bytes). Cached — only re-parsed when bytes change."""
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    return xls.sheet_names

@st.cache_data(show_spinner=False)
def read_sheet_cached(file_bytes, filename, sheet_name, header_row):
    """
    Read one CSV or one Excel sheet from raw bytes, with a chosen header row.
    header_row is 1-indexed: 1 = first row is the header (no rows skipped).
    Cached on (bytes, filename, sheet_name, header_row) so re-combining unchanged
    files/sheets is instant.
    """
    skip = max(int(header_row) - 1, 0)
    name = filename.lower()
    if name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(file_bytes), skiprows=skip)
    else:
        return pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, skiprows=skip)

def sync_file_configs(uploaded_files, configs_key):
    """
    Keep st.session_state[configs_key] in sync with the files currently sitting
    in the uploader widget: add config entries for newly added files, drop
    entries for files the user removed. Each config tracks which sheets to
    include and the header-row for each sheet, per file.
    """
    configs = st.session_state[configs_key]
    current_keys = set()
    for f in uploaded_files or []:
        fkey = f"{f.name}__{f.size}"
        current_keys.add(fkey)
        if fkey not in configs:
            fbytes = f.getvalue()
            excel = is_excel_file(f.name)
            try:
                sheets = get_excel_sheets_cached(fbytes) if excel else [None]
            except Exception as e:
                st.error(f"ไม่สามารถอ่านไฟล์ {f.name}: {e}")
                continue
            configs[fkey] = {
                "filename": f.name,
                "bytes": fbytes,
                "is_excel": excel,
                "sheets": sheets,
                "selected_sheets": list(sheets),   # default: include every sheet
                "header_rows": {s: 1 for s in sheets},  # default: header on row 1
            }
    # Drop configs for files the user removed from the uploader
    for k in list(configs.keys()):
        if k not in current_keys:
            del configs[k]

def combine_file_configs(configs_key):
    """Read every selected (file, sheet) per its own header row and concat into one table."""
    configs = st.session_state[configs_key]
    frames = []
    errors = []
    for cfg in configs.values():
        for s in cfg["selected_sheets"]:
            hr = cfg["header_rows"].get(s, 1)
            try:
                d = read_sheet_cached(cfg["bytes"], cfg["filename"], s, hr)
                frames.append(d)
            except Exception as e:
                label = f"{cfg['filename']}" + (f" [{s}]" if s else "")
                errors.append(f"{label}: {e}")
    for e in errors:
        st.error(f"อ่านไม่สำเร็จ — {e}")
    if not frames:
        return None
    return pd.concat(frames, ignore_index=True, sort=False)

_ISO_DATE_RE = re.compile(r'^\d{4}-\d{2}-\d{2}')

def _to_datetime_mixed(series, date_format="auto"):
    """
    pd.to_datetime infers ONE format from the first values in a column and
    silently returns NaT for any element that doesn't match it — so a column
    mixing '2026-01-15' and '2026-01-15 14:32:00' loses the second one.
    format='mixed' parses each element on its own. Falls back for older
    pandas (<2.0) that don't support format='mixed'.

    date_format controls how AMBIGUOUS date strings ('03/04/2026') are read:
    'auto'/'dmy' = day-first (our default convention), 'mdy' = month-first,
    'ymd' = year-first (e.g. '26/07/09' as Y/M/D). Unambiguous values
    ('17/7/2026') are read correctly regardless of this setting.

    ISO-format values (YYYY-MM-DD...) are ALWAYS parsed separately from the
    ambiguous ones and never get a dayfirst/yearfirst hint — pandas has a
    quirk where dayfirst=True + format='mixed' can otherwise corrupt
    unambiguous ISO dates ('2026-07-09' incorrectly becoming 2026-09-07)
    when mixed with D/M-style strings in the same column.
    """
    if pd.api.types.is_datetime64_any_dtype(series):
        return pd.to_datetime(series, errors="coerce")

    yearfirst = date_format == "ymd"
    dayfirst = date_format in ("auto", "dmy")  # ignored by pandas when yearfirst determines order first
    s = series if isinstance(series, pd.Series) else pd.Series(series)
    str_s = s.astype(str)
    is_iso = str_s.str.match(_ISO_DATE_RE).fillna(False)

    result = pd.Series(pd.NaT, index=s.index, dtype="datetime64[us]")
    try:
        if is_iso.any():
            result.loc[is_iso] = pd.to_datetime(s[is_iso], errors="coerce", format="mixed")
        if (~is_iso).any():
            result.loc[~is_iso] = pd.to_datetime(s[~is_iso], errors="coerce", format="mixed", dayfirst=dayfirst, yearfirst=yearfirst)
    except TypeError:
        # Older pandas (<2.0) without format='mixed' support
        if is_iso.any():
            result.loc[is_iso] = pd.to_datetime(s[is_iso], errors="coerce")
        if (~is_iso).any():
            result.loc[~is_iso] = pd.to_datetime(s[~is_iso], errors="coerce", dayfirst=dayfirst, yearfirst=yearfirst)
    return result

def _is_date_like(series):
    """True if series looks like dates rather than plain numbers."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return True
    sample = series.dropna().astype(str).head(20)
    if sample.empty:
        return False
    # If it parses as plain int/float, treat as numeric not date
    try:
        pd.to_numeric(sample)
        return False
    except (ValueError, TypeError):
        pass
    try:
        parsed = _to_datetime_mixed(sample)
        return parsed.notna().mean() >= 0.5
    except Exception:
        return False

def _coerce_for_minmax(sub, date_format="auto"):
    """
    Coerce DataFrame columns for min/max.
    Date-like columns → datetime (avoids -9223372036854775808 from NaT→int64).
    Numeric columns   → float.
    """
    if any(_is_date_like(sub[c]) for c in sub.columns):
        return sub.apply(
            lambda s: _to_datetime_mixed(s, date_format=date_format)
        )
    return sub.apply(pd.to_numeric, errors="coerce")

def _format_date_result(result, date_format="auto"):
    """
    If a formula's result is date-like, present it as a consistent YYYY-MM-DD
    string regardless of how the source data was formatted — e.g. one file's
    dates coming through as real datetime64 ('2026-07-17 00:00:00') and the
    other's as 'D/M/Y' text ('17/7/2026') would otherwise display completely
    differently in the two Standard Tables even though they're the same date.
    """
    if pd.api.types.is_datetime64_any_dtype(result):
        return result.dt.strftime("%Y-%m-%d").where(result.notna(), other=None)
    if _is_date_like(result):
        parsed = _to_datetime_mixed(result, date_format=date_format)
        return parsed.dt.strftime("%Y-%m-%d").where(parsed.notna(), other=None)
    return result

def _warn_unparseable_numeric(before, after, label):
    """
    pd.to_numeric(errors='coerce') silently turns anything it can't parse
    (e.g. 'N/A', '-', '150 pcs') into a blank value. That blank then gets
    treated as 0 in Diff calculations — so without a warning, dirty source
    data quietly looks identical to 'this quantity is genuinely zero'.
    """
    was_present = before.notna() if hasattr(before, "notna") else pd.notna(before)
    became_nan = after.isna() if hasattr(after, "notna") else pd.isna(after)
    n_bad = int((was_present & became_nan).sum())
    if n_bad > 0:
        st.warning(f"⚠️ {label}: มี {n_bad} แถวที่ค่าไม่ใช่ตัวเลข (parse ไม่ได้) ถูกนับเป็น 0 โดยอัตโนมัติ")
    return n_bad

def apply_formula(df, formula, date_format="auto"):
    fn   = formula["function"]
    cols = formula["columns"]
    target = formula.get("target", "")
    valid = [c for c in cols if c in df.columns]
    if not valid:
        return pd.Series([None] * len(df), index=df.index)
    sub = df[valid]
    if fn in DATE_ORDER_FUNCTIONS:
        # Explicit date-order function — take the first column's value like
        # FIRST, but force-parse it as a date in the chosen order regardless
        # of whether it looks ambiguous.
        result = df[valid[0]].reset_index(drop=True)
        forced_format = DATE_ORDER_FUNCTIONS[fn]
        if pd.api.types.is_datetime64_any_dtype(result):
            return result.dt.strftime("%Y-%m-%d").where(result.notna(), other=None)
        parsed = _to_datetime_mixed(result, date_format=forced_format)
        return parsed.dt.strftime("%Y-%m-%d").where(parsed.notna(), other=None)
    elif fn == "FIRST":
        result = df[valid[0]].reset_index(drop=True)
        return _format_date_result(result, date_format=date_format)
    elif fn == "SUM":
        coerced = sub.apply(pd.to_numeric, errors="coerce")
        for c in valid:
            _warn_unparseable_numeric(sub[c], coerced[c], f"Target '{target}' ← คอลัมน์ '{c}'")
        return coerced.sum(axis=1).reset_index(drop=True)
    elif fn == "MIN":
        coerced = _coerce_for_minmax(sub, date_format=date_format)
        result = coerced.min(axis=1).reset_index(drop=True)
        return _format_date_result(result, date_format=date_format)
    elif fn == "MAX":
        coerced = _coerce_for_minmax(sub, date_format=date_format)
        result = coerced.max(axis=1).reset_index(drop=True)
        return _format_date_result(result, date_format=date_format)
    elif fn == "COUNT":
        return sub.count(axis=1).reset_index(drop=True)
    return pd.Series([None] * len(df), index=df.index)

def build_standard(df, formulas, date_format="auto"):
    """date_format: per-side date parsing preset ('auto'/'dmy'/'mdy'), applied
    to every date-like target field built from this dataframe. Per-field
    overrides (formula['date_format']) take precedence when present."""
    result = {}
    for f in formulas:
        t = f.get("target", "").strip()
        if t:
            fmt = f.get("date_format") or date_format
            result[t] = apply_formula(df.reset_index(drop=True), f, date_format=fmt)
    return pd.DataFrame(result)

def formula_preview_str(fn, cols):
    if not cols:
        return f"{fn}(...)"
    return f"{fn}({', '.join(cols)})"

def compare_tables(std_a, std_b, key_cols, compare_fields):
    agg_a = std_a.copy()
    agg_b = std_b.copy()

    # Decide, per compare field, whether it's a date field or a plain numeric
    # field — using both sides' values together so either side alone having
    # blank/short data doesn't mis-detect it.
    date_fields = set()
    for f in compare_fields:
        parts = []
        if f in agg_a.columns:
            parts.append(agg_a[f])
        if f in agg_b.columns:
            parts.append(agg_b[f])
        if parts and _is_date_like(pd.concat(parts, ignore_index=True)):
            date_fields.add(f)

    for f in compare_fields:
        if f in date_fields:
            # Normalize to date-only (midnight) so any time-of-day component
            # is ignored — comparison happens at Year/Month/Day granularity only.
            if f in agg_a.columns:
                agg_a[f] = _to_datetime_mixed(agg_a[f]).dt.normalize()
            if f in agg_b.columns:
                agg_b[f] = _to_datetime_mixed(agg_b[f]).dt.normalize()
        else:
            if f in agg_a.columns:
                before = agg_a[f]
                agg_a[f] = pd.to_numeric(before, errors="coerce")
                _warn_unparseable_numeric(before, agg_a[f], f"Compare Field '{f}' (ฝั่ง A)")
            if f in agg_b.columns:
                before = agg_b[f]
                agg_b[f] = pd.to_numeric(before, errors="coerce")
                _warn_unparseable_numeric(before, agg_b[f], f"Compare Field '{f}' (ฝั่ง B)")

    # Key Columns can end up with mismatched dtypes between the two sides
    # (e.g. one file's date column read as real datetime64, the other's as
    # plain text) — pd.merge refuses to join on mismatched dtypes. Harmonize
    # each shared key column before grouping/merging so this can't happen.
    # Values are otherwise compared exactly as received — no trimming,
    # case-folding, or leading-zero stripping.
    for k in key_cols:
        if k not in agg_a.columns or k not in agg_b.columns:
            continue
        a_col, b_col = agg_a[k], agg_b[k]
        a_is_dt = pd.api.types.is_datetime64_any_dtype(a_col)
        b_is_dt = pd.api.types.is_datetime64_any_dtype(b_col)
        same_kind = a_col.dtype.kind == b_col.dtype.kind or (a_col.dtype.kind in "if" and b_col.dtype.kind in "if")
        if a_is_dt or b_is_dt or not same_kind:
            sample = pd.concat([
                a_col.dropna().astype(str), b_col.dropna().astype(str)
            ], ignore_index=True)
            if a_is_dt or b_is_dt or _is_date_like(sample):
                # Date-like key → normalize both sides to a plain Y-M-D string
                pa = _to_datetime_mixed(a_col)
                pb = _to_datetime_mixed(b_col)
                agg_a[k] = pa.dt.strftime("%Y-%m-%d").where(pa.notna(), None)
                agg_b[k] = pb.dt.strftime("%Y-%m-%d").where(pb.notna(), None)
            else:
                # Otherwise just make sure both sides use the same (string) dtype
                agg_a[k] = a_col.astype(str).where(a_col.notna(), None)
                agg_b[k] = b_col.astype(str).where(b_col.notna(), None)

    vk_a = [k for k in key_cols if k in agg_a.columns]
    vk_b = [k for k in key_cols if k in agg_b.columns]

    if vk_a:
        # Date fields can't be summed — use the earliest date in the group instead
        ag_a = {f: ("min" if f in date_fields else "sum") for f in compare_fields if f in agg_a.columns}
        fi_a = {c: "first" for c in agg_a.columns if c not in vk_a and c not in compare_fields}
        agg_a = agg_a.groupby(vk_a, as_index=False, dropna=False).agg({**ag_a, **fi_a})

    if vk_b:
        ag_b = {f: ("min" if f in date_fields else "sum") for f in compare_fields if f in agg_b.columns}
        fi_b = {c: "first" for c in agg_b.columns if c not in vk_b and c not in compare_fields}
        agg_b = agg_b.groupby(vk_b, as_index=False, dropna=False).agg({**ag_b, **fi_b})

    common_keys = [k for k in key_cols if k in agg_a.columns and k in agg_b.columns]
    if not common_keys:
        st.error("ไม่พบ Key Columns ร่วมกันใน Standard Table ทั้งสองฝั่ง")
        return None

    merged = pd.merge(agg_a, agg_b, on=common_keys, how="outer", suffixes=("_A","_B"), indicator=True)
    idx = merged.index
    n = len(merged)

    result_cols = {k: merged[k] for k in common_keys}
    # One boolean "did this field mismatch" column per compare field —
    # collected so Status/Mismatched-Fields can be derived across all
    # fields at once instead of per-row.
    mismatch_masks = {}

    for f in compare_fields:
        ca, cb = f"{f}_A", f"{f}_B"
        # If a compare field only exists on one side, pd.merge doesn't add
        # a suffix — fall back to the bare column (same values feed both
        # "sides"), matching the previous per-row lookup logic exactly.
        va = merged[ca] if ca in merged.columns else merged.get(f, pd.Series(np.nan, index=idx))
        vb = merged[cb] if cb in merged.columns else merged.get(f, pd.Series(np.nan, index=idx))

        if f in date_fields:
            both_notna = va.notna() & vb.notna()
            diff = pd.Series(np.where(both_notna, (va - vb).dt.days, np.nan), index=idx)
            result_cols[f"{f}_A"] = va.dt.date
            result_cols[f"{f}_B"] = vb.dt.date
            mismatch_masks[f] = diff.isna() | (diff != 0)
        else:
            va_num = pd.to_numeric(va, errors="coerce")
            vb_num = pd.to_numeric(vb, errors="coerce")
            diff = va_num.fillna(0.0) - vb_num.fillna(0.0)
            result_cols[f"{f}_A"] = va
            result_cols[f"{f}_B"] = vb
            mismatch_masks[f] = diff != 0

        result_cols[f"{f}_Diff"] = diff

    mismatch_df = pd.DataFrame(mismatch_masks, index=idx) if mismatch_masks else pd.DataFrame(index=idx)
    any_mismatch = mismatch_df.any(axis=1) if not mismatch_df.empty else pd.Series(False, index=idx)

    status = pd.Series("Match", index=idx, dtype=object)
    status[any_mismatch] = "Mismatch"
    status[merged["_merge"] == "left_only"] = "Missing B"
    status[merged["_merge"] == "right_only"] = "Missing A"
    result_cols["Status"] = status

    # Only meaningful when both sides have a row to compare — for
    # Missing A/B, every field differs "because" the row itself is
    # missing on one side, so listing them adds no information.
    mismatched_fields_col = pd.Series("", index=idx, dtype=object)
    is_mismatch = status == "Mismatch"
    if not mismatch_df.empty and is_mismatch.any():
        sub = mismatch_df.loc[is_mismatch]
        field_names = np.array(sub.columns)
        joined = [", ".join(field_names[row]) for row in sub.to_numpy()]
        mismatched_fields_col.loc[is_mismatch] = joined
    result_cols["Mismatched Fields"] = mismatched_fields_col

    return pd.DataFrame(result_cols, index=idx)

def export_excel(result_df, compare_fields):
    output = io.BytesIO()
    sc = "Status"
    match_df     = result_df[result_df[sc] == "Match"]      if sc in result_df.columns else pd.DataFrame()
    mismatch_df  = result_df[result_df[sc] == "Mismatch"]   if sc in result_df.columns else pd.DataFrame()
    missing_a_df = result_df[result_df[sc] == "Missing A"]  if sc in result_df.columns else pd.DataFrame()
    missing_b_df = result_df[result_df[sc] == "Missing B"]  if sc in result_df.columns else pd.DataFrame()

    total, nm, nmi, nmA, nmB = len(result_df), len(match_df), len(mismatch_df), len(missing_a_df), len(missing_b_df)
    summary_df = pd.DataFrame({
        "Metric": ["Total Rows","Match","Mismatch","Missing A","Missing B"],
        "Count":  [total, nm, nmi, nmA, nmB],
        "Pct": [
            "100%",
            f"{nm/total*100:.1f}%"  if total else "0%",
            f"{nmi/total*100:.1f}%" if total else "0%",
            f"{nmA/total*100:.1f}%" if total else "0%",
            f"{nmB/total*100:.1f}%" if total else "0%",
        ]
    })

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        summary_df.to_excel(writer,    sheet_name="Summary",    index=False)
        result_df.to_excel(writer,     sheet_name="All",        index=False)
        match_df.to_excel(writer,      sheet_name="Match",      index=False)
        mismatch_df.to_excel(writer,   sheet_name="Mismatch",   index=False)
        missing_a_df.to_excel(writer,  sheet_name="Missing A",  index=False)
        missing_b_df.to_excel(writer,  sheet_name="Missing B",  index=False)

        wb = writer.book
        fmt_pos  = wb.add_format({"bg_color":"#fef9c3"})
        fmt_neg  = wb.add_format({"bg_color":"#fee2e2"})
        fmt_zero = wb.add_format({"bg_color":"#dcfce7"})

        for sname, df_s in [("All",result_df),("Match",match_df),("Mismatch",mismatch_df),("Missing A",missing_a_df),("Missing B",missing_b_df)]:
            ws = writer.sheets[sname]
            for i,col in enumerate(df_s.columns):
                ml = max(len(str(col)), df_s[col].astype(str).str.len().max() if len(df_s) else 0)
                ws.set_column(i, i, min(ml+4, 40))
            if len(df_s):
                for f in compare_fields:
                    dc = f"{f}_Diff"
                    if dc in df_s.columns:
                        ci = list(df_s.columns).index(dc)
                        cl = get_column_letter(ci+1)
                        lr = len(df_s)+1
                        ws.conditional_format(f"{cl}2:{cl}{lr}", {"type":"cell","criteria":">","value":0,"format":fmt_pos})
                        ws.conditional_format(f"{cl}2:{cl}{lr}", {"type":"cell","criteria":"<","value":0,"format":fmt_neg})
                        ws.conditional_format(f"{cl}2:{cl}{lr}", {"type":"cell","criteria":"==","value":0,"format":fmt_zero})

    output.seek(0)
    return output

# ─────────────────────────────────────────────
# FORMULA BUILDER WIDGET — unified (A and B share one Target field per row)
# ─────────────────────────────────────────────
COL_WEIGHTS = [1.7, 0.85, 1.9, 1.5, 0.85, 1.9, 1.5, 0.45]

def formula_builder_unified(formulas_key, cols_a, cols_b, sys_a, sys_b, df_a=None, df_b=None, name_a="File A", name_b="File B"):
    rows = st.session_state[formulas_key]

    # ── Data preview — helps pick the right columns / functions, one per side ──
    pv1, pv2 = st.columns(2)
    with pv1:
        if df_a is not None and len(df_a.columns) > 0:
            with st.expander(f"👁️ Preview {name_a} — {len(df_a):,} rows × {len(df_a.columns)} columns", expanded=False):
                st.dataframe(df_a.head(8), use_container_width=True, height=220)
        else:
            st.caption("อัปโหลด File A ก่อน เพื่อดู Preview ข้อมูล")
    with pv2:
        if df_b is not None and len(df_b.columns) > 0:
            with st.expander(f"👁️ Preview {name_b} — {len(df_b):,} rows × {len(df_b.columns)} columns", expanded=False):
                st.dataframe(df_b.head(8), use_container_width=True, height=220)
        else:
            st.caption("อัปโหลด File B ก่อน เพื่อดู Preview ข้อมูล")

    st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)

    # ── Group headers (Target | FILE A | FILE B) — one flexbox block, weights match data rows below ──
    w = COL_WEIGHTS
    label_a = f"{name_a} — {sys_a}"
    label_b = f"{name_b} — {sys_b}"
    st.markdown(
        f'<div style="display:flex;gap:10px;align-items:flex-end;">'
        f'<div style="flex:{w[0]};"></div>'
        f'<div style="flex:{w[1]+w[2]+w[3]};text-align:center;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;'
        f'font-size:11px;font-weight:700;color:var(--accent);text-transform:uppercase;'
        f'letter-spacing:.08em;border-bottom:2px solid var(--border);padding-bottom:6px;" title="{label_a}">{label_a}</div>'
        f'<div style="flex:{w[4]+w[5]+w[6]};text-align:center;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;'
        f'font-size:11px;font-weight:700;color:var(--accent);text-transform:uppercase;'
        f'letter-spacing:.08em;border-bottom:2px solid var(--border);padding-bottom:6px;" title="{label_b}">{label_b}</div>'
        f'<div style="flex:{w[7]};"></div>'
        f'</div>', unsafe_allow_html=True
    )
    st.markdown(
        f'<div style="display:flex;gap:10px;margin-top:6px;">'
        f'<div style="flex:{w[0]};" class="col-hdr">Target Field</div>'
        f'<div style="flex:{w[1]};" class="col-hdr">Fn</div>'
        f'<div style="flex:{w[2]};" class="col-hdr">Source Columns</div>'
        f'<div style="flex:{w[3]};" class="col-hdr">Preview</div>'
        f'<div style="flex:{w[4]};" class="col-hdr">Fn</div>'
        f'<div style="flex:{w[5]};" class="col-hdr">Source Columns</div>'
        f'<div style="flex:{w[6]};" class="col-hdr">Preview</div>'
        f'<div style="flex:{w[7]};"></div>'
        f'</div>', unsafe_allow_html=True
    )

    to_delete_id = None
    for row in rows:
        rid = row.setdefault("id", uuid.uuid4().hex[:8])
        c = st.columns(COL_WEIGHTS)

        with c[0]:
            target = st.text_input(
                "tgt", value=row.get("target", ""),
                key=f"{formulas_key}_tgt_{rid}", label_visibility="collapsed",
                placeholder="เช่น SKU, Qty"
            )
            row["target"] = target

        for side_idx, (side_key, avail_cols) in enumerate([("a", cols_a), ("b", cols_b)]):
            base = 1 + side_idx * 3  # column offset: A starts at 1, B starts at 4
            side_cfg = row.setdefault(side_key, {"function": "FIRST", "columns": []})

            with c[base]:
                fn_idx = FUNCTIONS.index(side_cfg.get("function", "FIRST")) if side_cfg.get("function", "FIRST") in FUNCTIONS else 0
                fn = st.selectbox("fn", FUNCTIONS, index=fn_idx,
                                  key=f"{formulas_key}_fn_{side_key}_{rid}", label_visibility="collapsed")
                side_cfg["function"] = fn

            with c[base + 1]:
                if avail_cols:
                    safe_default = [x for x in side_cfg.get("columns", []) if x in avail_cols]
                    selected = st.multiselect(
                        "cols", avail_cols, default=safe_default,
                        key=f"{formulas_key}_cols_{side_key}_{rid}", label_visibility="collapsed"
                    )
                else:
                    raw = st.text_input(
                        "cols", value=",".join(side_cfg.get("columns", [])),
                        key=f"{formulas_key}_cols_{side_key}_{rid}", label_visibility="collapsed",
                        placeholder="COL1,COL2"
                    )
                    selected = [x.strip() for x in raw.split(",") if x.strip()]
                side_cfg["columns"] = selected

            with c[base + 2]:
                prev = formula_preview_str(fn, selected)
                st.markdown(f'<div class="fpv">{prev}</div>', unsafe_allow_html=True)

            row[side_key] = side_cfg

        with c[7]:
            st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
            if st.button("🗑", key=f"{formulas_key}_del_{rid}", help="ลบ Target Field นี้ (ทั้ง A และ B)"):
                to_delete_id = rid

    if to_delete_id is not None:
        st.session_state[formulas_key] = [r for r in rows if r["id"] != to_delete_id]
        st.rerun()

    st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)
    if st.button("＋ Add Target Field", key=f"add_{formulas_key}"):
        st.session_state[formulas_key].append({
            "id": uuid.uuid4().hex[:8],
            "target": "",
            "a": {"function": "FIRST", "columns": []},
            "b": {"function": "FIRST", "columns": []},
        })
        st.rerun()

    st.session_state[formulas_key] = rows


# ─────────────────────────────────────────────
# ENSURE DEFAULT FORMULAS LOADED ON STARTUP
# ─────────────────────────────────────────────
ensure_formulas_unified()

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div style="padding:28px 0 20px 0;">
  <div style="font-size:26px;font-weight:800;color:var(--text-primary);letter-spacing:-.5px;">📦 Stock Reconciliation Tool</div>
  <div style="font-size:13px;color:var(--text-muted);margin-top:5px;">Compare stock data between two systems — FC · JDA · MHT</div>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# STEP PROGRESS INDICATOR
# ─────────────────────────────────────────────
def render_step_indicator():
    step1_done = st.session_state.df_a_raw is not None and st.session_state.df_b_raw is not None
    step2_done = st.session_state.std_a is not None and st.session_state.std_b is not None
    step3_done = st.session_state.compare_result is not None

    if step3_done:
        current = 3
    elif step2_done:
        current = 3
    elif step1_done:
        current = 2
    else:
        current = 1

    def cls(step_num, done):
        if done: return "done"
        if step_num == current: return "current"
        return "pending"

    steps = [
        (1, "Upload Files", step1_done),
        (2, "Formula Builder", step2_done),
        (3, "Compare", step3_done),
    ]
    html = '<div class="step-track">'
    for idx, (num, label, done) in enumerate(steps):
        c = cls(num, done)
        icon = "✓" if done else str(num)
        active = " active" if (done or num == current) else ""
        html += f'<div class="step-item"><div class="step-circle {c}">{icon}</div><div class="step-label{active}">{label}</div></div>'
        if idx < len(steps) - 1:
            line_done = "done" if steps[idx][2] else ""
            html += f'<div class="step-line {line_done}"></div>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)

render_step_indicator()

# ─────────────────────────────────────────────
# SWAP A ↔ B
# ─────────────────────────────────────────────
def swap_a_and_b():
    s = st.session_state
    s.sys_a, s.sys_b = s.sys_b, s.sys_a
    s.df_a_raw, s.df_b_raw = s.df_b_raw, s.df_a_raw
    s.file_configs_a, s.file_configs_b = s.file_configs_b, s.file_configs_a
    s.auto_combine_sig_a, s.auto_combine_sig_b = s.auto_combine_sig_b, s.auto_combine_sig_a
    # Mapping/build/compare state depended on the old A/B arrangement — reset it
    s.formulas_unified = load_default_formulas_unified(s.sys_a, s.sys_b)
    s.formulas_a = None
    s.formulas_b = None
    s.std_a = None
    s.std_b = None
    s.compare_result = None

# ─────────────────────────────────────────────
# RESET / CLEAR SESSION
# ─────────────────────────────────────────────
def reset_all_session():
    """
    Wipe every piece of session state — uploaded files, formulas/mapping,
    standard tables, compare results, and any per-widget keys (formula
    builder rows, file-uploader selections, etc.) — then reinitialize back
    to a fresh app. Lets the user start a new comparison without having to
    refresh the browser tab.
    """
    for k in list(st.session_state.keys()):
        del st.session_state[k]
    init_state()
    ensure_formulas_unified()

# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tab_upload, tab_formula, tab_compare = st.tabs([
    "📁  Upload Files",
    "🔧  Formula Builder",
    "⚡  Compare & Dashboard",
])

# ─────────────────────────────────────────────
# UPLOAD SIDE RENDERER (shared by File A / File B)
# ─────────────────────────────────────────────
def _short_tag(name, fallback):
    """First ~5 alphanumeric characters of a filename, used as a compact
    column suffix (e.g. 'Qty_RBS02' instead of 'Qty_A')."""
    base = name.split(" +")[0]           # strip any "+N ไฟล์" suffix
    base = base.rsplit(".", 1)[0]        # strip file extension
    base = re.sub(r'[^0-9A-Za-zก-๙]', '', base)  # keep letters/digits (incl. Thai)
    tag = base[:5]
    return tag if tag else fallback

def side_display_name(configs_key, fallback):
    """Show the real uploaded filename(s) instead of a generic 'File A' / 'File B' label."""
    configs = st.session_state.get(configs_key, {})
    names = [cfg["filename"] for cfg in configs.values()]
    if not names:
        return fallback
    if len(names) == 1:
        n = names[0]
        return n if len(n) <= 42 else n[:39] + "..."
    first = names[0]
    first = first if len(first) <= 28 else first[:25] + "..."
    return f"{first} +{len(names)-1} ไฟล์"

def _upload_case_signature(configs):
    """
    Returns a signature tuple if the upload is the 'simple' case (exactly one
    file, one sheet selected, header on row 1 — the common case) so it can be
    auto-combined without requiring a button click. Returns None otherwise
    (multiple files/sheets, or a custom header row) — those cases still need
    the explicit Combine button since there's a real decision being made.
    """
    if len(configs) != 1:
        return None
    fkey, cfg = next(iter(configs.items()))
    if len(cfg["selected_sheets"]) != 1:
        return None
    s = cfg["selected_sheets"][0]
    if cfg["header_rows"].get(s, 1) != 1:
        return None
    return (fkey, s, 1)

def render_upload_side(side_code, sys_key, configs_key, df_key, uploader_key, combine_key, auto_sig_key):
    std_key = "std_" + ("a" if sys_key == "sys_a" else "b")
    display_name = side_display_name(configs_key, f"File {side_code}")

    st.markdown(f'<div class="card-label">ระบบ {display_name}</div>', unsafe_allow_html=True)
    sys_choice = st.radio(
        f"sys_{sys_key}", SYSTEMS,
        index=SYSTEMS.index(st.session_state[sys_key]),
        horizontal=True, key=f"{sys_key}_radio", label_visibility="collapsed"
    )
    if sys_choice != st.session_state[sys_key]:
        st.session_state[sys_key] = sys_choice
        sys_a_now = sys_choice if sys_key == "sys_a" else st.session_state.sys_a
        sys_b_now = sys_choice if sys_key == "sys_b" else st.session_state.sys_b
        st.session_state.formulas_unified = load_default_formulas_unified(sys_a_now, sys_b_now)
        st.session_state.formulas_a = None
        st.session_state.formulas_b = None
        st.session_state[std_key] = None
        st.session_state.compare_result = None
        st.rerun()

    st.markdown('<div class="card-label" style="margin-top:16px;">อัปโหลดไฟล์ (เลือกได้หลายไฟล์)</div>', unsafe_allow_html=True)
    files = st.file_uploader(
        f"File {side_code}", type=["csv", "xlsx", "xls"],
        key=uploader_key, label_visibility="collapsed",
        accept_multiple_files=True
    )
    sync_file_configs(files, configs_key)

    configs = st.session_state[configs_key]
    if configs:
        for fkey, cfg in configs.items():
            n_sheets = len(cfg["sheets"])
            title = f"📄 {cfg['filename']}" + (f" ({n_sheets} sheets)" if cfg["is_excel"] and n_sheets > 1 else "")
            with st.expander(title, expanded=False):
                if cfg["is_excel"] and n_sheets > 1:
                    sel = st.multiselect(
                        "Sheets ที่จะรวม", cfg["sheets"],
                        default=cfg["selected_sheets"],
                        key=f"{configs_key}_{fkey}_sheets"
                    )
                    cfg["selected_sheets"] = sel
                else:
                    cfg["selected_sheets"] = cfg["sheets"]

                for s in cfg["selected_sheets"]:
                    label = f"แถวที่เป็น Header — Sheet '{s}'" if s else "แถวที่เป็น Header"
                    hr = st.number_input(
                        label, min_value=1, max_value=1000,
                        value=int(cfg["header_rows"].get(s, 1)), step=1,
                        key=f"{configs_key}_{fkey}_hdr_{s}",
                        help="ถ้า Header ไม่ได้อยู่แถวแรกของไฟล์ ให้ใส่เลขแถวที่ Header อยู่จริง"
                    )
                    cfg["header_rows"][s] = hr

        st.markdown("<div style='height:4px'></div>", unsafe_allow_html=True)

        sig = _upload_case_signature(configs)
        if sig is not None:
            # Simple case (1 file, 1 sheet, header row 1) → combine automatically, no button needed
            if st.session_state[auto_sig_key] != sig:
                combined = combine_file_configs(configs_key)
                if combined is not None:
                    st.session_state[df_key] = combined
                    st.session_state[std_key] = None
                    st.session_state.compare_result = None
                st.session_state[auto_sig_key] = sig
            st.caption("✅ ไฟล์เดียว รวมข้อมูลให้อัตโนมัติแล้ว")
        else:
            # Multiple files/sheets, or a custom header row → needs an explicit decision
            st.session_state[auto_sig_key] = None
            if st.button(f"🔗 รวมไฟล์ → Table {side_code}", key=combine_key, use_container_width=True):
                with st.spinner("กำลังอ่านและรวมไฟล์..."):
                    combined = combine_file_configs(configs_key)
                if combined is not None:
                    st.session_state[df_key] = combined
                    st.session_state[std_key] = None
                    st.session_state.compare_result = None
                    st.success(f"✅ รวมแล้ว {len(combined):,} rows × {len(combined.columns)} columns")
                else:
                    st.warning("ไม่มีข้อมูลให้รวม — กรุณาเลือกอย่างน้อย 1 Sheet")
    else:
        st.caption("ยังไม่ได้อัปโหลดไฟล์")

    if st.session_state[df_key] is not None:
        df = st.session_state[df_key]
        st.markdown("<hr class='s'>", unsafe_allow_html=True)
        st.success(f"✅ {display_name} พร้อมใช้งาน — {len(df):,} rows × {len(df.columns)} columns")
        with st.expander("Preview 10 rows"):
            st.dataframe(df.head(10), use_container_width=True)

# ═══════════════════════════════════════════════
# TAB 1 — UPLOAD
# ═══════════════════════════════════════════════
with tab_upload:
    hdr_l, hdr_r = st.columns([4.3, 1.7])
    with hdr_l:
        st.markdown('<div class="sec-title">Upload Source Files</div>', unsafe_allow_html=True)
        st.markdown('<div class="sec-sub">อัปโหลดไฟล์เดียวจะรวมให้อัตโนมัติ — ถ้ามีหลายไฟล์/หลาย Sheet/Header ไม่ได้อยู่แถวแรก ค่อยกด "รวมไฟล์" เอง</div>', unsafe_allow_html=True)
    with hdr_r:
        st.markdown("<div style='height:8px'></div>", unsafe_allow_html=True)
        br1, br2 = st.columns(2)
        with br1:
            if st.button("🔄 สลับ A ↔ B", use_container_width=True, help="สลับข้อมูล/ระบบ/ไฟล์ระหว่าง A และ B ทั้งหมด"):
                swap_a_and_b()
                st.rerun()
        with br2:
            with st.popover("🗑 Reset ทั้งหมด", use_container_width=True, help="ล้างไฟล์ที่อัปโหลด, Mapping, และผลลัพธ์ทั้งหมด เริ่มใหม่ตั้งแต่ต้น"):
                st.write("ล้างไฟล์ที่อัปโหลด, Mapping ทั้งหมด และผลลัพธ์ Compare — เริ่มต้นแอปใหม่ทั้งหมด")
                st.caption("การกระทำนี้ไม่สามารถย้อนกลับได้")
                if st.button("⚠️ ยืนยัน Reset ทั้งหมด", type="primary", use_container_width=True):
                    reset_all_session()
                    st.rerun()

    col_a, col_b = st.columns(2)

    with col_a:
        with st.container():
            st.markdown('<div class="card">', unsafe_allow_html=True)
            render_upload_side("A", "sys_a", "file_configs_a", "df_a_raw", "uploader_a", "combine_a", "auto_combine_sig_a")
            st.markdown('</div>', unsafe_allow_html=True)

    with col_b:
        with st.container():
            st.markdown('<div class="card">', unsafe_allow_html=True)
            render_upload_side("B", "sys_b", "file_configs_b", "df_b_raw", "uploader_b", "combine_b", "auto_combine_sig_b")
            st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("<hr class='s'>", unsafe_allow_html=True)
    ra = st.session_state.df_a_raw is not None
    rb = st.session_state.df_b_raw is not None
    name_a = side_display_name("file_configs_a", "File A")
    name_b = side_display_name("file_configs_b", "File B")
    if ra and rb:
        st.success("✅ ทั้งสองไฟล์พร้อมแล้ว — ไปที่ **Formula Builder** เพื่อตรวจสอบ Mapping")
    elif ra:
        st.info(f"ℹ️ รอไฟล์ฝั่ง B... ({name_a} พร้อมแล้ว)")
    elif rb:
        st.info(f"ℹ️ รอไฟล์ฝั่ง A... ({name_b} พร้อมแล้ว)")
    else:
        st.info("ℹ️ กรุณาอัปโหลดทั้งสองไฟล์")

# ═══════════════════════════════════════════════
# TAB 2 — FORMULA BUILDER
# ═══════════════════════════════════════════════
with tab_formula:
    st.markdown('<div class="sec-title">Formula Builder</div>', unsafe_allow_html=True)
    st.markdown('<div class="sec-sub">กำหนด Target Field เดียว ใช้ร่วมกันทั้ง File A และ File B — ไม่ต้องตั้งชื่อซ้ำสองรอบ</div>', unsafe_allow_html=True)

    cols_a = list(st.session_state.df_a_raw.columns) if st.session_state.df_a_raw is not None else []
    cols_b = list(st.session_state.df_b_raw.columns) if st.session_state.df_b_raw is not None else []

    st.markdown('<div class="card">', unsafe_allow_html=True)
    formula_builder_unified(
        "formulas_unified", cols_a, cols_b,
        st.session_state.sys_a, st.session_state.sys_b,
        df_a=st.session_state.df_a_raw, df_b=st.session_state.df_b_raw,
        name_a=side_display_name("file_configs_a", "File A"),
        name_b=side_display_name("file_configs_b", "File B"),
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Build ──
    st.markdown("<hr class='s'>", unsafe_allow_html=True)

    all_targets = [r.get("target","").strip() for r in st.session_state.formulas_unified if r.get("target","").strip()]
    dup_targets = sorted(set(t for t in all_targets if all_targets.count(t) > 1))
    if dup_targets:
        st.warning(f"⚠️ มี Target Field ชื่อซ้ำ: {', '.join(dup_targets)} — กรุณาตั้งชื่อให้ไม่ซ้ำกัน")

    if st.button("🔨 Build Standard Tables", use_container_width=True, type="primary"):
        if st.session_state.df_a_raw is None or st.session_state.df_b_raw is None:
            st.error("กรุณาอัปโหลดทั้งสองไฟล์ก่อน")
        elif not st.session_state.formulas_unified:
            st.error("กรุณาเพิ่ม Target Field อย่างน้อย 1 ข้อก่อน")
        elif dup_targets:
            st.error("กรุณาแก้ไขชื่อ Target Field ที่ซ้ำกันก่อน")
        else:
            with st.spinner("กำลัง Build..."):
                st.session_state.formulas_a = [
                    {"target": r["target"], "function": r["a"]["function"], "columns": r["a"]["columns"]}
                    for r in st.session_state.formulas_unified if r.get("target","").strip()
                ]
                st.session_state.formulas_b = [
                    {"target": r["target"], "function": r["b"]["function"], "columns": r["b"]["columns"]}
                    for r in st.session_state.formulas_unified if r.get("target","").strip()
                ]
                st.session_state.std_a = build_standard(st.session_state.df_a_raw, st.session_state.formulas_a)
                st.session_state.std_b = build_standard(st.session_state.df_b_raw, st.session_state.formulas_b)
            st.session_state.compare_result = None
            st.success("✅ Build Standard Tables สำเร็จ")

    if st.session_state.std_a is not None and st.session_state.std_b is not None:
        tag_a = _short_tag(side_display_name("file_configs_a", "A"), "A")
        tag_b = _short_tag(side_display_name("file_configs_b", "B"), "B")
        p1, p2 = st.columns(2)
        with p1:
            st.markdown(f'<div class="card-label" style="margin-top:16px">Standard Table {tag_a} — {len(st.session_state.std_a):,} rows</div>', unsafe_allow_html=True)
            st.dataframe(st.session_state.std_a, use_container_width=True, height=400)
        with p2:
            st.markdown(f'<div class="card-label" style="margin-top:16px">Standard Table {tag_b} — {len(st.session_state.std_b):,} rows</div>', unsafe_allow_html=True)
            st.dataframe(st.session_state.std_b, use_container_width=True, height=400)

# ═══════════════════════════════════════════════
# TAB 3 — COMPARE
# ═══════════════════════════════════════════════
with tab_compare:
    st.markdown('<div class="sec-title">Compare Engine</div>', unsafe_allow_html=True)
    st.markdown('<div class="sec-sub">เลือก Key Columns และ Compare Fields — สามารถเลือก/ยกเลิก Field ที่ต้องการ Compare ได้</div>', unsafe_allow_html=True)

    if st.session_state.std_a is None or st.session_state.std_b is None:
        st.warning("⚠️ กรุณา Build Standard Tables ใน Formula Builder ก่อน")
    else:
        std_a = st.session_state.std_a
        std_b = st.session_state.std_b
        cols_a = list(std_a.columns)
        cols_b = list(std_b.columns)
        common_cols = [c for c in cols_a if c in cols_b]

        cfg_col, _ = st.columns([1, 1.8])
        with cfg_col:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.markdown('<div class="card-label">⚙️ Key Columns (Join Keys)</div>', unsafe_allow_html=True)

            key_default = [k for k in ["SKU","Location"] if k in common_cols] or (common_cols[:1] if common_cols else [])
            key_cols = st.multiselect(
                "Key Columns", options=common_cols, default=key_default,
                key="key_cols_ms", label_visibility="collapsed"
            )

            # Compare fields — all common non-key cols, user can deselect
            auto_compare_all = [c for c in common_cols if c not in key_cols]

            if auto_compare_all:
                st.markdown("<hr class='s'>", unsafe_allow_html=True)
                st.markdown('<div class="card-label">📋 Compare Fields</div>', unsafe_allow_html=True)
                selected_compare = st.multiselect(
                    "Compare Fields",
                    options=auto_compare_all,
                    default=auto_compare_all,
                    key="compare_fields_ms",
                    label_visibility="collapsed"
                )
            else:
                selected_compare = []
                st.info("ไม่พบ Field ที่ชื่อตรงกันระหว่าง A และ B (นอกจาก Key)")

            st.session_state.key_cols = key_cols
            st.session_state.compare_fields = selected_compare
            auto_compare = selected_compare
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("<hr class='s'>", unsafe_allow_html=True)

        if st.button("⚡ Run Compare", use_container_width=True, type="primary"):
            if not key_cols:
                st.error("กรุณาเลือก Key Columns อย่างน้อย 1")
            elif not auto_compare:
                st.error("ไม่มี Field ร่วมกันที่จะ Compare")
            else:
                with st.spinner("กำลัง Compare..."):
                    result = compare_tables(std_a, std_b, key_cols, auto_compare)
                if result is not None:
                    tag_a = _short_tag(side_display_name("file_configs_a", "A"), "A")
                    tag_b = _short_tag(side_display_name("file_configs_b", "B"), "B")
                    rename_map = {}
                    for col in result.columns:
                        if col.endswith("_A") and not col.endswith("_Diff"):
                            rename_map[col] = f"{col[:-2]}_{tag_a}"
                        elif col.endswith("_B") and not col.endswith("_Diff"):
                            rename_map[col] = f"{col[:-2]}_{tag_b}"
                    result = result.rename(columns=rename_map)
                    st.session_state.compare_result = result
                    st.session_state.dash_filter = "All"  # reset filter every fresh Compare run
                    st.success(f"✅ Compare เสร็จสิ้น — {len(result):,} rows")

        # ── Dashboard — shown right here, same page, as soon as a result exists ──
        if st.session_state.compare_result is not None:
            result = st.session_state.compare_result
            cf = st.session_state.get("compare_fields", [])

            st.markdown("<hr class='s'>", unsafe_allow_html=True)
            st.markdown('<div class="sec-title" style="font-size:16px;">📊 Dashboard</div>', unsafe_allow_html=True)

            total = len(result)
            nm  = len(result[result.Status=="Match"])      if "Status" in result.columns else 0
            nmi = len(result[result.Status=="Mismatch"])   if "Status" in result.columns else 0
            nmA = len(result[result.Status=="Missing A"])  if "Status" in result.columns else 0
            nmB = len(result[result.Status=="Missing B"])  if "Status" in result.columns else 0
            pct = nm/total*100 if total else 0

            # KPI
            st.markdown(f"""
            <div class="kpi-grid">
              <div class="kpi-box kpi-total"><div class="kpi-value">{total:,}</div><div class="kpi-label">Total Rows</div></div>
              <div class="kpi-box kpi-match"><div class="kpi-value">{nm:,}</div><div class="kpi-label">Match ({pct:.1f}%)</div></div>
              <div class="kpi-box kpi-mis"  ><div class="kpi-value">{nmi:,}</div><div class="kpi-label">Mismatch</div></div>
              <div class="kpi-box kpi-miss" ><div class="kpi-value">{nmA:,}</div><div class="kpi-label">Missing A</div></div>
              <div class="kpi-box kpi-miss" ><div class="kpi-value">{nmB:,}</div><div class="kpi-label">Missing B</div></div>
            </div>
            """, unsafe_allow_html=True)

            # Match-rate bar
            st.markdown(f"""
            <div style="margin-bottom:20px;">
              <div style="display:flex;justify-content:space-between;font-size:12px;color:var(--text-muted);margin-bottom:5px;">
                <span>Match Rate</span><span>{pct:.1f}%</span>
              </div>
              <div style="height:8px;background:var(--bg-card-alt);border-radius:4px;overflow:hidden;border:1px solid var(--border);">
                <div style="height:100%;width:{pct}%;background:linear-gradient(90deg,#22c55e,#16a34a);border-radius:4px;"></div>
              </div>
            </div>
            """, unsafe_allow_html=True)

            # Total Diff per Compare Field — row counts alone don't show
            # severity (e.g. "3 rows mismatched" could mean 3 units off or
            # 3,000 units off), so surface the summed |Diff| per field too.
            diff_summaries = []
            for f in cf:
                dc = f"{f}_Diff"
                if dc not in result.columns:
                    continue
                numeric_col = pd.to_numeric(result[dc], errors="coerce")
                total_abs = numeric_col.abs().sum(skipna=True)
                n_nonzero = int((numeric_col.fillna(0) != 0).sum())
                a_col = result.get(f"{f}_A")
                is_date = a_col is not None and a_col.dropna().map(lambda v: isinstance(v, date)).any()
                diff_summaries.append((f, total_abs, n_nonzero, "วัน" if is_date else ""))

            if diff_summaries:
                st.markdown('<div class="card-label" style="margin-top:4px;">📐 Total Diff ต่อ Field</div>', unsafe_allow_html=True)
                boxes = "".join(
                    f'<div class="kpi-box"><div class="kpi-value" style="font-size:22px;">{tot:,.0f}{(" " + unit) if unit else ""}</div>'
                    f'<div class="kpi-label">{fld} · {n:,} แถวต่าง</div></div>'
                    for fld, tot, n, unit in diff_summaries
                )
                st.markdown(f'<div class="kpi-grid">{boxes}</div>', unsafe_allow_html=True)

            st.markdown("<hr class='s'>", unsafe_allow_html=True)

            # Filter — right here on the Compare page
            st.markdown('<div class="card-label">🔍 Filter ผลลัพธ์</div>', unsafe_allow_html=True)
            fsel = st.radio("filter", ["All","Match","Mismatch","Missing A","Missing B"],
                            horizontal=True, label_visibility="collapsed", key="dash_filter")
            filtered = result if fsel == "All" else result[result.Status == fsel]

            def highlight_status(row):
                s = row.get("Status","")
                base = [""] * len(row)
                idx = list(row.index).index("Status") if "Status" in row.index else -1
                if s == "Match"    and idx>=0: base[idx] = "background-color:var(--status-match-bg);color:var(--status-match-text)"
                elif s == "Mismatch" and idx>=0: base[idx] = "background-color:var(--status-mismatch-bg);color:var(--status-mismatch-text)"
                elif "Missing" in str(s) and idx>=0: base[idx] = "background-color:var(--status-missing-bg);color:var(--status-missing-text)"
                return base

            def color_diff(val):
                try:
                    v = float(val)
                    if v > 0: return "background-color:var(--status-mismatch-bg);color:var(--status-mismatch-text)"
                    if v < 0: return "background-color:var(--status-missing-bg);color:var(--status-missing-text)"
                    return "background-color:var(--status-match-bg);color:var(--status-match-text)"
                except:
                    return ""

            n_cells = filtered.shape[0] * filtered.shape[1]
            if n_cells <= STYLE_CELL_HARD_LIMIT:
                sty = filtered.style.apply(highlight_status, axis=1)
                for col in filtered.columns:
                    if col.endswith("_Diff"):
                        try:
                            sty = sty.map(color_diff, subset=[col])
                        except AttributeError:
                            sty = sty.applymap(color_diff, subset=[col])
                st.dataframe(sty, use_container_width=True, height=500)
            else:
                st.info(
                    f"ℹ️ ตารางมี {n_cells:,} cells (เกิน {STYLE_CELL_HARD_LIMIT:,}) — "
                    "แสดงแบบไม่มีสีไฮไลต์เพื่อความเร็ว ลองใช้ตัวกรอง Status ด้านบนเพื่อดูข้อมูลแบบมีสี"
                )
                st.dataframe(filtered, use_container_width=True, height=500)
            st.caption(f"แสดง {len(filtered):,} จาก {total:,} rows")

            excel_data = export_excel(result, cf)
            export_filename = f"CompareStock_{datetime.now():%d%m%y_%H%M}.xlsx"
            st.download_button(
                f"📥 Export {export_filename}", data=excel_data,
                file_name=export_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

# ── Footer ──
st.markdown("""
<div style="margin-top:48px;padding-top:20px;border-top:1px solid var(--border);
            text-align:center;color:var(--text-muted);font-size:11px;">
  Stock Reconciliation Tool &nbsp;·&nbsp; FC / JDA / MHT &nbsp;·&nbsp; Streamlit + Pandas
</div>
""", unsafe_allow_html=True)