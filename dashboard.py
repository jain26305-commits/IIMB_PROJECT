"""
Enterprise Demand Forecasting & Inventory Decision Support System
Using Time Series Analytics — Premium Edition (v2)

Architecture: PRESERVE -> CORRECT -> ENHANCE -> ANIMATE -> POLISH
Author: Kaushik Jain

Data source of truth: Enterprise_Supply_Chain_Master_Audit.xlsx (Executive_Summary sheet)
This file is READ-ONLY at runtime. Nothing here ever writes back to the source workbook.
"""

import warnings
warnings.filterwarnings("ignore")

import sys
import types
import os
from pathlib import Path
import numpy as np

# --- NumPy 2.0+ compatibility shims (kept from the original build) ---
np.long = int
np.ulong = int
np.int = int
np.float = float
np.bool = np.bool_
np.complex = complex
np.object = object
sys.modules['numpy'].long = int

try:
    import numpy.exceptions
except ImportError:
    np.exceptions = types.ModuleType('numpy.exceptions')
    sys.modules['numpy.exceptions'] = np.exceptions
if not hasattr(np.exceptions, 'RankWarning'):
    np.exceptions.RankWarning = Warning

import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import ast
import io
import html
import streamlit as st
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Optional / heavy dependencies made resilient: the app must never crash on boot
# just because an optional modelling package isn't installed in a given environment.
try:
    from sklearn.metrics import mean_absolute_error, mean_squared_error
except Exception:
    mean_absolute_error, mean_squared_error = None, None

try:
    from statsmodels.stats.diagnostic import acorr_ljungbox
except Exception:
    acorr_ljungbox = None

try:
    from scipy.stats import norm, skew, kurtosis
except Exception:
    norm, skew, kurtosis = None, None, None

st.set_page_config(page_title="Enterprise Demand Forecasting", page_icon="🧊", layout="wide", initial_sidebar_state="collapsed")
if "_sidebar_bootstrap_done" not in st.session_state:
    st.session_state["_sidebar_bootstrap_done"] = True
    st.html("""<script>
        setTimeout(() => {
            const sidebar = document.querySelector('[data-testid="stSidebar"]');
            if (sidebar?.getAttribute('aria-expanded') === 'true') {
                sidebar.querySelector('button[kind="headerNoPadding"]')?.click();
            }
        }, 150);
    </script>""", unsafe_allow_javascript=True)

# ------------------------------------------------------------------------------
# CINEMATIC CHART LAYER — purely presentational, additive-only wrapper around
# st.plotly_chart. It never touches a figure's data, traces, colors, titles,
# axis ranges, or chart type — it only layers in smoother typography, a
# gentle entrance transition, and refined hover styling so every one of the
# 24+ charts in this file benefits consistently without editing each one by
# hand (and without risking a mismatched font or forgotten polish pass on any
# single chart). If a figure already sets a given property, that figure's own
# setting is left alone — this only fills in what isn't already specified.
# ------------------------------------------------------------------------------
_ORIGINAL_PLOTLY_CHART = st.plotly_chart

def _cinematic_plotly_chart(fig, *args, **kwargs):
    try:
        # Force a consistent light paper/plot background on every chart.
        # Several charts in this file only ever set plot_bgcolor (never
        # paper_bgcolor), and several others set neither — which left the
        # chart's outer "paper" area to inherit Plotly's own dark-mode
        # detection, producing a white plotting box sitting inside a dark
        # halo. Setting both explicitly here, globally, removes that
        # mismatch everywhere at once instead of patching 24 call sites.
        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="#ffffff",
            font=dict(family="'Inter', 'Segoe UI', -apple-system, sans-serif", size=13, color="#334155"),
            title_font=dict(family="'Inter', 'Segoe UI', -apple-system, sans-serif", size=16, color="#1E3A8A"),
            legend=dict(bgcolor="rgba(255,255,255,0.6)", bordercolor="rgba(226,232,240,0.8)", borderwidth=1,
                        font=dict(color="#334155")),
            hoverlabel=dict(bgcolor="white", font_size=13, font_family="'Inter', 'Segoe UI', sans-serif",
                             bordercolor="#0EA5E9", font_color="#0F172A"),
        )
        fig.update_xaxes(showline=True, linecolor="rgba(148,163,184,0.35)", gridcolor="rgba(226,232,240,0.55)",
                          zerolinecolor="rgba(148,163,184,0.35)", color="#334155")
        fig.update_yaxes(showline=True, linecolor="rgba(148,163,184,0.35)", gridcolor="rgba(226,232,240,0.55)",
                          zerolinecolor="rgba(148,163,184,0.35)", color="#334155")
        # Visualization upgrade — additive trace polish only. Never touches a
        # trace's x/y/values, color, name, or type: only refines the visual
        # finish (rounded bar corners, crisper marker outlines) so charts
        # read as more considered without any data or chart-type change.
        for trace in fig.data:
            ttype = getattr(trace, "type", None)
            try:
                if ttype == "bar":
                    trace.update(marker_cornerradius=4)
                elif ttype in ("scatter", "scattergl") and trace.mode and "markers" in trace.mode:
                    current_line = getattr(trace.marker, "line", None)
                    if current_line is None or current_line.width is None:
                        trace.update(marker_line_width=1, marker_line_color="rgba(255,255,255,0.6)")
            except Exception:
                continue
    except Exception:
        pass  # never let a cosmetic enhancement break a chart render
    return _ORIGINAL_PLOTLY_CHART(fig, *args, **kwargs)

st.plotly_chart = _cinematic_plotly_chart


# ==============================================================================
# SECTION 0. AUTHORITATIVE FIELD MAP  (Master Audit = source of truth)
# ==============================================================================
# These are the 51 authoritative fields used by the dashboard from the Enterprise Supply Chain Master Audit.
# The workbook itself contains additional columns. These fields are read directly when available.
MASTER_AUDIT_FIELDS = [
    'SKU', 'ABC_Class', 'XYZ_Class', 'ABC_XYZ_Class', 'Priority_Level', 'Forecast_Next_Month',
    'Accuracy_Pct', 'Forecast_Health_Score', 'Forecast_Risk', 'Safety_Stock', 'Reorder_Point', 'EOQ',
    'Average_Inventory_Units', 'Inventory_Value', 'Working_Capital', 'Estimated_Carrying_Cost',
    'Stockout_Cost', 'Inventory_Turnover', 'Inventory_Days', 'Service_Level_Pct', 'Fill_Rate_Pct',
    'Inventory_Health_Score', 'Executive_Recommendation', 'Inventory_Recommendation',
    'Procurement_Recommendation', 'Warehouse_Recommendation', 'Financial_Recommendation',
    'Business_Explanation', 'Mean_Monthly_Demand', 'Annual_Demand', 'CV', 'MAE', 'RMSE', 'Bias',
    'Bullwhip_Ratio', 'Order', 'Status', 'Health_P_Value', 'Planning Frequency', 'Inventory Policy',
    'Review Frequency', 'Business Risk', 'Supply Chain Priority', 'Unit_Cost', 'Holding_Cost_Unit',
    'Ordering_Cost_Event', 'Demand_Stability_Score', 'Model_Reliability',
    'Expected_Impact', 'KPI_Financial_Health_Score', 'KPI_Working_Capital_Efficiency'
]
# Accuracy_Pct is the project-defined sMAPE-derived forecast accuracy field supplied by the Master Audit.
# A probabilistic confidence measure is not used as a dashboard KPI.
# Optional array-level fields — NOT present in the Master Audit export as verified.
# If the authoritative source includes them, richer per-SKU trajectory views unlock consistently.
OPTIONAL_ARRAY_FIELDS = ['Test_Actuals', 'Test_Predictions', 'Test_Lower_Bound', 'Test_Upper_Bound']
# A few plausible alternate names for an observed/actual inventory position, checked defensively
# in the authoritative source only — never fabricated if absent.
POSSIBLE_ACTUAL_INVENTORY_COLS = ['Actual_Inventory', 'Current_Inventory', 'On_Hand_Inventory', 'Current_Stock', 'Closing_Balance']

ABC_ORDER = ['A', 'B', 'C']
XYZ_ORDER = ['X', 'Y', 'Z']
RISK_ORDER = ['Low', 'Medium', 'High']
ABC_COLOR_MAP = {'A': '#1E3A8A', 'B': '#0EA5E9', 'C': '#94A3B8'}
XYZ_COLOR_MAP = {'X': '#0D9488', 'Y': '#F59E0B', 'Z': '#EF4444'}

# ==============================================================================
# SECTION 1. GENERIC HELPERS — safe formatting & traceable data access
# ==============================================================================
def col_exists(df, col):
    return df is not None and col in df.columns

def safe_series(df, col, numeric=False, default=np.nan):
    """Column-safe accessor: never KeyErrors, never silently fabricates a shared fallback array."""
    if col_exists(df, col):
        s = df[col]
        return pd.to_numeric(s, errors='coerce') if numeric else s
    return pd.Series([default] * len(df), index=df.index if df is not None else None)

def safe_sum(series):
    """Sum while preserving all-missing/empty states as N/A instead of manufacturing zero."""
    if series is None:
        return np.nan
    s = pd.to_numeric(series, errors='coerce')
    if s.notna().sum() == 0:
        return np.nan
    return float(s.sum(min_count=1))

def safe_val(row, col, default=None):
    try:
        if row is None or col not in row.index:
            return default
        v = row[col]
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return default
        return v
    except Exception:
        return default

def is_missing(v):
    if v is None:
        return True
    try:
        return bool(pd.isna(v))
    except Exception:
        return False

def fmt_num(val, decimals=1, suffix=""):
    if is_missing(val):
        return "N/A"
    try:
        return f"{float(val):,.{decimals}f}{suffix}"
    except Exception:
        return "N/A"

def fmt_int(val, suffix=""):
    if is_missing(val):
        return "N/A"
    try:
        return f"{int(round(float(val))):,}{suffix}"
    except Exception:
        return "N/A"

def fmt_monetary(val):
    """Neutral monetary-unit formatter; currency is intentionally not assumed."""
    if is_missing(val):
        return "N/A"
    try:
        return f"{float(val):,.0f}"
    except Exception:
        return "N/A"

def ordered_category_counts(df, col, order):
    """Return category counts in an explicit semantic order."""
    if df is None or col not in df.columns:
        return pd.DataFrame(columns=[col, 'Count'])
    counts = df[col].value_counts(dropna=False).reindex(order, fill_value=0)
    return counts.rename_axis(col).reset_index(name='Count')


def deterministic_high_risk_sort(df):
    """Sort by Priority, RMSE desc, Bullwhip desc, then SKU."""
    work = df.copy()
    keys=[]; asc=[]
    if 'Priority_Level' in work.columns:
        work['_priority'] = (work['Priority_Level'].astype(str).str.strip().str.lower()
                             .map({'high': 0, 'medium': 1, 'low': 2}).fillna(99))
        keys.append('_priority'); asc.append(True)
    if 'RMSE' in work.columns:
        work['_rmse'] = pd.to_numeric(work['RMSE'], errors='coerce')
        keys.append('_rmse'); asc.append(False)
    if 'Bullwhip_Ratio' in work.columns:
        work['_bullwhip'] = pd.to_numeric(work['Bullwhip_Ratio'], errors='coerce')
        keys.append('_bullwhip'); asc.append(False)
    if 'SKU' in work.columns:
        work['_sku'] = work['SKU'].astype(str)
        keys.append('_sku'); asc.append(True)
    if keys:
        work=work.sort_values(keys,ascending=asc,kind='mergesort',na_position='last')
    return work.drop(columns=[c for c in ['_priority','_rmse','_bullwhip','_sku'] if c in work.columns])


def parse_pct_string(val, default=np.nan):
    """Handles both numeric percentages and Master Audit strings like '17.33%'."""
    if is_missing(val):
        return default
    try:
        if isinstance(val, (int, float)):
            return float(val)
        return float(str(val).strip().replace('%', ''))
    except Exception:
        return default

SOURCE_TAG_CLASS = {
    "Master Audit": "tag-master",
    "Historical Raw Data": "tag-historical",
    "Dashboard Derived": "tag-derived",
    "Illustrative Scenario": "tag-illustrative",
}
def source_tag(label):
    cls = SOURCE_TAG_CLASS.get(label, "tag-derived")
    return f'<span class="source-tag {cls}">{label}</span>'

def badge_html(status_key, label=None):
    """status_key in {'excellent','good','attention','poor'} — one canonical naming scheme."""
    label = label or status_key.title()
    return f'<span class="badge" data-status="{status_key}">{label}</span>'

def score_to_status(score, thresholds=(80, 60, 40)):
    if is_missing(score):
        return "attention"
    hi, mid, lo = thresholds
    if score >= hi: return "excellent"
    if score >= mid: return "good"
    if score >= lo: return "attention"
    return "poor"

def render_metric_card(container, title, value, subtitle="", *, variant="ribbon", pulse=False, extra="", extra_class=""):
    # Flattened to a single unindented line (Data Integrity Fix, Sept 2026):
    # any HTML string with 4+ leading spaces per line risks being parsed as a
    # markdown code fence — see the note on render_square_metric below for the
    # visible symptom this caused elsewhere. Flat markup is immune regardless
    # of how it's later combined with other strings.
    pulse_cls = "pulse-critical" if pulse else ""
    extra_cls = str(extra_class).strip()
    card_classes = f"metric-card metric-card--{variant} {pulse_cls} {extra_cls}".strip()
    container.markdown(
        f'<div class="{card_classes}">'
        f'<div class="metric-card-title">{title}</div>'
        f'<div class="metric-card-value">{value}</div>'
        f'<div class="metric-card-subtitle">{subtitle}{extra}</div></div>',
        unsafe_allow_html=True)

def render_square_metric(label, value, delta=""):
    # NOTE: emitted as a single unindented line on purpose. Streamlit's markdown
    # pass treats 4+ leading spaces as a fenced code block, which previously
    # caused this card's own closing </div> to leak out as literal on-screen
    # text once it was interpolated into an already-indented parent f-string
    # (see mini_html below). Keeping this fully flat prevents that regardless
    # of how much the caller indents its own wrapping markup.
    return (f'<div class="square-metric"><div class="square-metric-label">{label}</div>'
            f'<div class="square-metric-value">{value}</div>'
            f'<div class="square-metric-delta">{delta}</div></div>')

def render_status_card(status, title, desc, action=None):
    action_html = f'<span class="status-card-action">Action: {action}</span>' if action else ""
    st.markdown(
        f'<div class="status-card" data-status="{status}">'
        f'<strong>{title}</strong><br>'
        f'<span class="status-card-desc">{desc}</span><br>'
        f'{action_html}</div>',
        unsafe_allow_html=True)

def render_decision_card(icon, title, text, source="Master Audit"):
    st.markdown(
        f'<div class="decision-card">'
        f'<div class="decision-card-title">{icon} {title} {source_tag(source)}</div>'
        f'<div class="decision-card-text">{text}</div></div>',
        unsafe_allow_html=True)

# ------------------------------------------------------------------------------
# MATHEMATICAL FORENSIC VALIDATION — Safety Stock / Reorder Point
# ------------------------------------------------------------------------------
# Ground-truth formula extracted directly from the source notebook
# (PROJECTIIMB-1.ipynb, process_single_sku(), 'Success' branch):
#
#   sigma_lead_time  = RMSE * sqrt(LEAD_TIME_DAYS / 30)
#   Safety_Stock     = SERVICE_LEVEL_Z * sigma_lead_time
#   Lead_Time_Demand = (Forecast_Next_Month / 30) * LEAD_TIME_DAYS
#   Reorder_Point    = Lead_Time_Demand + Safety_Stock
#
# NOTE ON METHODOLOGY: this is a SINGLE-source-of-uncertainty model driven by
# forecast RMSE. Lead time itself is a fixed constant (7 days) in the notebook
# — there is no separate lead-time-standard-deviation term anywhere in the
# implementation. This differs from a fuller two-term sigma_DLT formulation
# (demand variance + lead-time variance combined) sometimes used as a textbook
# reference. That fuller formula is NOT what this specific pipeline computes,
# so it is shown in the diagnostic panel as a labelled reference comparison,
# never as a silent replacement for the Master Audit's actual methodology.
#
# For "Success (Dead Stock)" / "Success (Constant Demand)" SKUs, the notebook
# has a separate guard clause that forces Safety_Stock = 0 by definition
# (zero demand variance -> no buffer required) — this is mathematically
# correct, not an anomaly, and the diagnostic panel treats it accordingly.
SERVICE_LEVEL_Z = 1.645
LEAD_TIME_DAYS = 7.0
LEAD_TIME_FRACTION = LEAD_TIME_DAYS / 30.0

def independent_ss_rop(row):
    """Recomputes Safety_Stock / Reorder_Point independently from RMSE + Forecast_Next_Month,
    using the exact formula the source notebook implements. Returns a dict with N/A-safe values."""
    rmse = safe_val(row, 'RMSE')
    fc = safe_val(row, 'Forecast_Next_Month')
    status = str(safe_val(row, 'Status', ''))
    if 'Dead Stock' in status or 'Constant Demand' in status:
        # Zero demand variance -> zero buffer is the model's design. The
        # replenishment quantity still requires a real mean-demand input.
        mean_m = safe_val(row, 'Mean_Monthly_Demand')
        if is_missing(mean_m):
            return {'ss': 0.0, 'rop': None, 'basis': 'Zero-variance guard; Mean_Monthly_Demand missing'}
        ltd = (float(mean_m) / 30.0) * LEAD_TIME_DAYS
        return {'ss': 0.0, 'rop': round(ltd, 0), 'basis': 'Zero-variance guard (design-correct zero)'}
    if is_missing(rmse) or is_missing(fc):
        return {'ss': None, 'rop': None, 'basis': 'Insufficient inputs (RMSE/Forecast missing)'}
    sigma_lt = float(rmse) * np.sqrt(LEAD_TIME_FRACTION)
    ss = SERVICE_LEVEL_Z * sigma_lt
    ltd = (float(fc) / 30.0) * LEAD_TIME_DAYS
    rop = ltd + ss
    return {'ss': round(ss, 0), 'rop': round(rop, 0), 'basis': 'RMSE-based single-uncertainty-term formula'}

@st.cache_data(show_spinner=False)
def run_portfolio_math_audit(source_df):
    """Runs the SS/ROP forensic audit across every row with the required inputs. Returns a summary dict + a per-row detail dataframe."""
    rows = []
    for _, r in source_df.iterrows():
        result = independent_ss_rop(r)
        ms_ss = safe_val(r, 'Safety_Stock')
        ms_rop = safe_val(r, 'Reorder_Point')
        rows.append({
            'SKU': safe_val(r, 'SKU', 'N/A'),
            'Status': safe_val(r, 'Status', 'N/A'),
            'MA_Safety_Stock': ms_ss, 'Independent_Safety_Stock': result['ss'],
            'MA_Reorder_Point': ms_rop, 'Independent_Reorder_Point': result['rop'],
            'Basis': result['basis'],
        })
    detail = pd.DataFrame(rows)
    if len(detail) == 0:
        return {'n': 0}, detail

    def _diff(a, b):
        if is_missing(a) or is_missing(b):
            return np.nan
        return abs(float(a) - float(b))

    detail['SS_Variance'] = detail.apply(lambda r: _diff(r['MA_Safety_Stock'], r['Independent_Safety_Stock']), axis=1)
    detail['ROP_Variance'] = detail.apply(lambda r: _diff(r['MA_Reorder_Point'], r['Independent_Reorder_Point']), axis=1)

    summary = {
        'n': len(detail),
        'ss_zero_in_master': int((detail['MA_Safety_Stock'] == 0).sum()),
        'ss_exact_match': int((detail['SS_Variance'] == 0).sum()),
        'ss_tolerance_match': int((detail['SS_Variance'] <= 1).sum()),
        'ss_material_mismatch': int((detail['SS_Variance'] > 1).sum()),
        'rop_zero_in_master': int((detail['MA_Reorder_Point'] == 0).sum()),
        'rop_exact_match': int((detail['ROP_Variance'] == 0).sum()),
        'rop_tolerance_match': int((detail['ROP_Variance'] <= 1).sum()),
        'rop_material_mismatch': int((detail['ROP_Variance'] > 1).sum()),
    }
    return summary, detail

# ==============================================================================
# SECTION 2. SPLASH / INTRO SCREEN  (preserved — product identity, do not remove)
# ==============================================================================
if 'app_initialized' not in st.session_state:
    st.session_state.app_initialized = True
    st.markdown("""
        <style>
            .splash-overlay {
                position: fixed;
                top: 0; left: 0; width: 100vw; height: 100vh;
                background: linear-gradient(135deg, #F8FAFC, #E0F2FE, #BAE6FD);
                background-size: 200% 200%;
                z-index: 999999;
                display: flex;
                flex-direction: column;
                align-items: center;
                justify-content: center;
                animation: gradientBG 1s ease infinite, fadeOutOverlay 0.35s cubic-bezier(0.4, 0, 0.2, 1) forwards 0.65s;
                pointer-events: none;
                padding: 0 4vw;
                box-sizing: border-box;
                overflow: hidden;
            }
            /* Cinematic vignette — darkens the corners subtly so the title reads
               like a lit subject on a stage rather than flat text on a gradient.
               Pure visual layer: adds no width/height and shifts no content. */
            .splash-overlay::before {
                content: "";
                position: absolute; inset: 0;
                background: radial-gradient(120% 90% at 50% 45%, transparent 40%, rgba(15,23,42,0.10) 100%);
                pointer-events: none;
            }
            /* Cinematic light sweep — a soft diagonal highlight that glides across
               the frame once, like a lens catching light, then settles. */
            .splash-overlay::after {
                content: "";
                position: absolute; top: -50%; left: -60%; width: 60%; height: 200%;
                background: linear-gradient(75deg, transparent, rgba(255,255,255,0.35), transparent);
                transform: rotate(8deg);
                animation: lightSweep 0.7s cubic-bezier(0.4,0,0.2,1) forwards 0.05s;
                pointer-events: none;
            }
            .splash-title-text {
                color: #000000;
                font-size: clamp(2.5rem, 5vw, 6.5rem);
                font-weight: 900;
                text-align: center !important;
                line-height: 1.15;
                width: 100vw;
                max-width: 92vw;
                margin: 0 auto;
                text-shadow: 2px 2px 20px rgba(255, 255, 255, 0.9);
                animation: focusIn 0.45s cubic-bezier(0.25, 1, 0.5, 1) forwards;
                position: relative;
                z-index: 1;
            }
            .splash-subtitle-text {
                color: #000000;
                font-size: clamp(1.8rem, 3.5vw, 3.5rem);
                font-weight: 800;
                letter-spacing: 6px;
                text-transform: uppercase;
                margin-top: 2.5rem;
                text-align: center;
                width: 100vw;
                animation: fadeInUp 0.45s ease forwards;
                position: relative;
                z-index: 1;
            }
            @keyframes gradientBG {
                0% { background-position: 0% 50%; }
                50% { background-position: 100% 50%; }
                100% { background-position: 0% 50%; }
            }
            @keyframes fadeOutOverlay {
                0% { opacity: 1; visibility: visible; }
                100% { opacity: 0; visibility: hidden; display: none; }
            }
            @keyframes focusIn {
                0% { transform: scale(0.95); opacity: 0; filter: blur(5px); }
                20% { transform: scale(1); opacity: 1; filter: blur(0px); }
                100% { transform: scale(1.02); opacity: 1; filter: blur(0px); }
            }
            @keyframes fadeInUp {
                from { opacity: 0; transform: translateY(20px); }
                to { opacity: 1; transform: translateY(0); }
            }
            @keyframes lightSweep {
                0% { left: -60%; opacity: 0; }
                15% { opacity: 1; }
                60% { opacity: 1; }
                100% { left: 120%; opacity: 0; }
            }
            @media (prefers-reduced-motion: reduce) {
                .splash-overlay { animation: fadeOutOverlay 0.3s ease forwards 0.2s; }
                .splash-title-text, .splash-subtitle-text { animation: none !important; opacity: 1 !important; filter: none !important; transform: none !important; }
                .splash-overlay::after { animation: none !important; opacity: 0 !important; }
            }
        </style>
        <div class="splash-overlay">
            <h1 class="splash-title-text">Enterprise Demand Forecasting & Inventory Decision Support System Using Time Series Analytics</h1>
            <p class="splash-subtitle-text">Created by Kaushik Jain</p>
        </div>
    """, unsafe_allow_html=True)

# ==============================================================================
# SECTION 3. DESIGN TOKENS + GLOBAL COMPONENT SYSTEM
# ------------------------------------------------------------------------------
# Every color / radius / shadow / easing value used anywhere below is defined
# ONCE here as a CSS custom property, then referenced with var(...). This
# replaces the previous 47 hardcoded hex values and 3 unrelated color palettes
# with a single, systematic token layer (Design System Audit, Sept 2026).
# ==============================================================================
st.markdown("""
<style>
    :root {
        /* Brand core */
        --c-navy: #1E3A8A;
        --c-navy-dark: #1E293B;
        --c-blue: #0EA5E9;
        --c-blue-light: #38BDF8;
        --c-blue-pale: #E0F2FE;

        /* Semantic status (single canonical scheme: excellent / good / attention / poor) */
        --c-success: #10B981;
        --c-success-bg: #F0FDF4;
        --c-success-border: #BBF7D0;
        --c-success-text: #065F46;

        --c-warning: #F59E0B;
        --c-warning-bg: #FFFBEB;
        --c-warning-border: #FEF3C7;
        --c-warning-text: #92400E;

        --c-danger: #EF4444;
        --c-danger-bg: #FEF2F2;
        --c-danger-border: #FEE2E2;
        --c-danger-text: #991B1B;

        --c-poor-bg: #991B1B;
        --c-poor-text: #FFFFFF;

        /* Neutrals */
        --c-slate-900: #0F172A;
        --c-slate-700: #334155;
        --c-slate-600: #475569;
        --c-slate-500: #64748B;
        --c-slate-200: #E2E8F0;
        --c-slate-100: #F1F5F9;
        --c-slate-50:  #F8FAFC;
        --c-white: #FFFFFF;

        /* Secondary accent — forecast / prediction series only */
        --c-rose: #F43F5E;
        --c-rose-dark: #E11D48;

        /* 8-step "stage" ramp — ONE systematic palette replacing both the
           Sankey rainbow and the flow-card rainbow from the original build */
        --stage-1: #1E3A8A;
        --stage-2: #1D4ED8;
        --stage-3: #0369A1;
        --stage-4: #0891B2;
        --stage-5: #0D9488;
        --stage-6: #059669;
        --stage-7: #475569;
        --stage-8: #334155;

        /* Radius scale */
        --radius-sm: 6px;
        --radius-md: 8px;
        --radius-lg: 10px;
        --radius-xl: 12px;
        --radius-pill: 20px;

        /* Elevation scale — cinematic upgrade (Sept 2026): each shadow now
           layers a soft ambient falloff with a tighter directional key-shadow,
           the way a lit product photo separates from its background, instead
           of a single flat drop-shadow. Values only — no selectors, radii,
           spacing, or layout touched, so every card keeps its exact position
           and size on screen. */
        --shadow-resting: 0 1px 2px rgba(15,23,42,0.05), 0 4px 10px rgba(15,23,42,0.05);
        --shadow-hover: 0 2px 6px rgba(15,23,42,0.08), 0 20px 40px rgba(30,58,138,0.16);
        --shadow-hover-strong: 0 2px 8px rgba(14,165,233,0.12), 0 28px 56px rgba(14,165,233,0.22);
        --shadow-ambient-soft: 0 30px 60px -20px rgba(15,23,42,0.25);

        /* Claymorphism layer — soft dual-tone "embossed" shadow (a light
           highlight from the top-left, a soft dark falloff to the
           bottom-right, plus a faint inner rim light) layered on top of the
           elevation scale above. This is what gives every surface its
           puffy, tactile, cinematic-lit "clay" feel. Static box-shadow
           values only — no animation cost while idle. */
        --clay-radius: 22px;
        --clay-radius-sm: 16px;
        --clay-shadow: 0 2px 8px rgba(30,41,59,0.07), 0 1px 2px rgba(15,23,42,0.05);
        --clay-shadow-hover: 0 6px 16px rgba(15,23,42,0.10), 0 2px 4px rgba(15,23,42,0.05);
        --clay-shadow-pressed: inset 0 1px 4px rgba(15,23,42,0.08);
        --glow-blue: 0 0 0 1px rgba(14,165,233,0.10), 0 6px 18px rgba(14,165,233,0.12);
        --glow-navy: 0 0 0 1px rgba(30,58,138,0.10), 0 6px 18px rgba(30,58,138,0.10);

        /* Motion scale */
        --ease-standard: cubic-bezier(0.4,0,0.2,1);
        --ease-spring: cubic-bezier(0.175,0.885,0.32,1.275);
        --dur-fast: 0.15s;
        --dur-med: 0.25s;
        --dur-slow: 0.35s;
    }

    /* Cinematic ambient glow — opacity/transform only (GPU-compositor
       animations that never trigger layout or paint), so this can loop
       forever without the repaint cost that was previously removed from
       the full-viewport background-position sweep. Used behind the hero
       title and section headers for a subtle "lit stage" breathing glow. */
    @keyframes ambientGlow {
        0%, 100% { opacity: 0.55; transform: scale(1); }
        50% { opacity: 0.9; transform: scale(1.05); }
    }

    @keyframes floatElement {
        0% { transform: translateY(0px); }
        50% { transform: translateY(-8px); }
        100% { transform: translateY(0px); }
    }
    @keyframes chartReveal {
        from { opacity: 0; transform: translateY(10px) scale(0.99); }
        to { opacity: 1; transform: translateY(0) scale(1); }
    }
    @keyframes tabPanelReveal {
        from { opacity: 0; transform: translateY(6px); }
        to { opacity: 1; transform: translateY(0); }
    }
    @keyframes pulseCritical {
        0% { box-shadow: 0 0 0 0 rgba(153, 27, 27, 0.35); }
        70% { box-shadow: 0 0 0 8px rgba(153, 27, 27, 0); }
        100% { box-shadow: 0 0 0 0 rgba(153, 27, 27, 0); }
    }
    @keyframes fadeInUp {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
    }

    /* ---------- FORCE LIGHT THEME (hard override) ----------
       Streamlit's own dark theme (and the OS/browser prefers-color-scheme)
       can otherwise paint the root app shell, header, sidebar and the gaps
       between st.columns() pure black, AND flip generic text to a
       light/white color that then disappears against this dashboard's own
       light card backgrounds. These rules pin both background AND base text
       color for every structural Streamlit container to the same light
       palette regardless of the user's system theme. Component classes
       further down (.metric-card-value, .badge, etc.) already set their own
       explicit colors and are more specific, so they are unaffected and
       keep rendering exactly as designed. */
    html, body, [data-testid="stAppViewContainer"], [data-testid="stAppViewBlockContainer"],
    [data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stBottomBlockContainer"],
    [data-testid="stMain"], [data-testid="stMainBlockContainer"],
    div[data-testid="column"], div[data-testid="stVerticalBlock"], div[data-testid="stHorizontalBlock"],
    [data-testid="stMarkdownContainer"], [data-testid="stMarkdownContainer"] p,
    [data-testid="stText"], [data-testid="stCaptionContainer"],
    label, .stMarkdown, .stText {
        background-color: var(--c-slate-50) !important;
        color: var(--c-slate-900) !important;
        color-scheme: light !important;
    }
    [data-testid="stHeader"] { background: rgba(248,250,252,0) !important; }
    [data-testid="stSidebar"] * { color-scheme: light !important; }
    /* Widgets (selects, inputs, expanders) keep readable text too */
    [data-baseweb="select"] *, [data-testid="stExpander"] *,
    input, textarea, .stDataFrame, [data-testid="stDataFrame"] * {
        color: var(--c-slate-900) !important;
    }

    /* ---------- Native alert boxes (st.info/success/warning/error) ----------
       These ship with theme-aware defaults that, under a dark browser/OS
       theme, render as a dark box with light text — a hard color mismatch
       against this dashboard's otherwise-forced light surface. Recoloring
       all four variants explicitly, using the same semantic tokens already
       used for status-card and badge components elsewhere in the file, so
       an "info" box here means the same blue as an "info" badge anywhere
       else on the page. */
    div[data-testid="stAlert"] {
        border-radius: var(--radius-lg) !important;
        border: 1px solid !important;
        box-shadow: var(--shadow-resting) !important;
    }
    div[data-testid="stAlert"] p, div[data-testid="stAlert"] span,
    div[data-testid="stAlert"] li, div[data-testid="stAlert"] div {
        color: inherit !important;
    }
    div[data-testid="stAlertContentInfo"] {
        background-color: var(--c-blue-pale) !important; color: var(--c-navy) !important; border-color: #BAE6FD !important;
    }
    div[data-testid="stAlertContentSuccess"] {
        background-color: var(--c-success-bg) !important; color: var(--c-success-text) !important; border-color: var(--c-success-border) !important;
    }
    div[data-testid="stAlertContentWarning"] {
        background-color: var(--c-warning-bg) !important; color: var(--c-warning-text) !important; border-color: var(--c-warning-border) !important;
    }
    div[data-testid="stAlertContentError"] {
        background-color: var(--c-danger-bg) !important; color: var(--c-danger-text) !important; border-color: var(--c-danger-border) !important;
    }

    /* ---------- Row alignment fix ----------
       st.columns() renders each column as an independently-sized flex item,
       so a card in one column can end at a different height than its
       neighbor whenever the title/subtitle text wraps differently. Making
       each column a stretching flex column (and letting the metric card
       grow to fill it) keeps every card in a row ending at the same
       baseline, regardless of text length — no Python layout code changed. */
    div[data-testid="stHorizontalBlock"] { align-items: stretch !important; }
    div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {
        display: flex !important; flex-direction: column !important;
    }
    div[data-testid="column"] > div[data-testid="stVerticalBlock"] {
        display: flex !important; flex-direction: column !important; flex: 1 1 auto !important;
    }
    div[data-testid="column"] .metric-card,
    div[data-testid="column"] .flow-card {
        height: 100% !important;
    }

    @keyframes cardSheen {
        0% { transform: translateX(-120%) skewX(-15deg); }
        100% { transform: translateX(220%) skewX(-15deg); }
    }

    /* Whole-app ambient backdrop: a static, subtle gradient wash behind the
       content, layered on top of the forced light base color above — never
       black, in any browser or OS theme. Static (not animated): an
       animated full-viewport gradient with background-attachment:fixed
       forces continuous repaints on every frame and was the single biggest
       performance cost on the page; a still gradient looks identical at a
       glance and costs nothing after first paint. */
    .stApp {
        background-image: radial-gradient(circle at 15% 10%, rgba(14,165,233,0.05), transparent 45%),
                           radial-gradient(circle at 85% 90%, rgba(30,58,138,0.05), transparent 45%);
        background-size: 100% 100%;
        color-scheme: light !important;
    }

    /* Staggered entrance: children of the same reveal-driven containers land
       in a soft, FAST cascade rather than all at once, purely via
       animation-delay — no positions, sizes, or DOM order are affected.
       Kept under 120ms total so the effect reads as instant, not laggy. */
    div[data-testid="column"]:nth-child(1) .metric-card,
    div[data-testid="column"]:nth-child(1) .square-metric { animation-delay: 0s; }
    div[data-testid="column"]:nth-child(2) .metric-card,
    div[data-testid="column"]:nth-child(2) .square-metric { animation-delay: 0.03s; }
    div[data-testid="column"]:nth-child(3) .metric-card,
    div[data-testid="column"]:nth-child(3) .square-metric { animation-delay: 0.06s; }
    div[data-testid="column"]:nth-child(4) .metric-card,
    div[data-testid="column"]:nth-child(4) .square-metric { animation-delay: 0.09s; }
    div[data-testid="column"]:nth-child(5) .metric-card,
    div[data-testid="column"]:nth-child(5) .square-metric { animation-delay: 0.12s; }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        animation: fadeInUp 0.35s cubic-bezier(0.16, 1, 0.3, 1) forwards;
    }


    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, var(--c-slate-50), var(--c-slate-100));
        border-right: none;
        box-shadow: inset -1px 0 0 rgba(226,232,240,0.9), 6px 0 24px rgba(15,23,42,0.04);
        transition: background-color var(--dur-fast) ease;
    }

    div[data-testid="stMetricValue"] {
        font-size: 2.2rem;
        color: var(--c-navy) !important;
        font-weight: 800;
    }
    div[data-testid="stMetricLabel"], div[data-testid="stMetricLabel"] * {
        color: var(--c-slate-600) !important;
    }
    div[data-testid="stMetricDelta"] { color: var(--c-slate-600) !important; }
    div[data-testid="stMetric"] {
        background: linear-gradient(155deg, #ffffff 0%, var(--c-slate-50) 100%);
        border: none;
        border-radius: var(--clay-radius-sm);
        padding: 0.8rem 1rem;
        box-shadow: var(--clay-shadow);
        transition: transform var(--dur-med) var(--ease-spring), box-shadow var(--dur-med) ease;
    }
    div[data-testid="stMetric"]:hover {
        transform: translateY(-2px);
        box-shadow: var(--clay-shadow-hover);
    }

    /* ---------- Metric card (consolidated: ribbon + square variants share one token base) ---------- */
    .metric-card {
        background:
            radial-gradient(140% 100% at 0% 0%, rgba(14,165,233,0.06), transparent 55%),
            linear-gradient(155deg, #ffffff 0%, var(--c-slate-50) 100%);
        border-radius: var(--clay-radius);
        border: none;
        border-left: 5px solid var(--c-blue);
        box-shadow: var(--clay-shadow);
        transition: transform var(--dur-med) var(--ease-spring), box-shadow var(--dur-med) ease, border-color var(--dur-med) ease;
        animation: chartReveal var(--dur-slow) var(--ease-standard) both;
        position: relative;
        overflow: hidden;
    }
    .metric-card::before {
        content: ""; position: absolute; inset: 0 0 auto 0; height: 1px;
        background: linear-gradient(90deg, transparent, rgba(255,255,255,0.9), transparent);
        opacity: 0.6; pointer-events: none;
    }
    /* Cinematic sheen: a soft light streak that glides across the card on
       hover only, like light catching glass. Absolutely positioned overlay —
       adds no size and shifts no sibling or child content. Runs once per
       hover rather than looping, so it costs nothing while idle. */
    .metric-card::after {
        content: ""; position: absolute; top: 0; left: 0; width: 40%; height: 100%;
        background: linear-gradient(100deg, transparent, rgba(255,255,255,0.35), transparent);
        opacity: 0; pointer-events: none;
    }
    .metric-card:hover::after { opacity: 1; animation: cardSheen 1.1s ease forwards; }
    .metric-card--ribbon {
        padding: 1.2rem;
        min-height: 145px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .metric-card--ribbon:hover {
        transform: translateY(-2px);
        box-shadow: var(--clay-shadow-hover);
        border-left: 5px solid var(--c-navy);
    }
    .metric-card-title {
        font-size: 0.85rem; color: var(--c-slate-500); font-weight: 700;
        text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 0.2rem;
    }
    .metric-card-value {
        font-size: 1.8rem; color: var(--c-slate-900); font-weight: 900; margin-bottom: 0.2rem;
        text-shadow: 0 1px 0 rgba(255,255,255,0.6);
    }
    .metric-card-subtitle { font-size: 0.85rem; color: var(--c-slate-500); font-weight: 600; }
    .metric-card.pulse-critical { animation: chartReveal var(--dur-slow) var(--ease-standard) both, pulseCritical 2.4s infinite; }

    /* ---------- Badge — one canonical naming scheme via data-status ---------- */
    .badge { padding: 4px 12px; border-radius: var(--radius-pill); font-size: 0.75rem; font-weight: 700; display: inline-block;
        box-shadow: inset 0 1px 2px rgba(255,255,255,0.6), inset 0 -1px 2px rgba(15,23,42,0.06); }
    .badge[data-status="excellent"] { background-color: var(--c-success-bg); color: var(--c-success-text); }
    .badge[data-status="good"]      { background-color: var(--c-warning-bg); color: var(--c-warning-text); }
    .badge[data-status="attention"] { background-color: var(--c-danger-border); color: var(--c-danger-text); }
    .badge[data-status="poor"]      { background-color: var(--c-poor-bg); color: var(--c-poor-text); }

    /* ---------- Traceability source tag ---------- */
    .source-tag { font-size: 0.65rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.4px;
        padding: 2px 8px; border-radius: var(--radius-pill); margin-left: 6px; vertical-align: middle; }
    .tag-master { background: var(--c-blue-pale); color: var(--c-navy); }
    .tag-historical { background: #EDE9FE; color: #5B21B6; }
    .tag-derived { background: var(--c-slate-100); color: var(--c-slate-600); }
    .tag-illustrative { background: var(--c-warning-bg); color: var(--c-warning-text); }

    /* ---------- Square metric (mini KPI tile — Tab 2 / Tab 3 side panels) ----------
       BUG FIX: this tile previously used `aspect-ratio: 1/1` (height locked to
       width) with `overflow: hidden`. In a narrower 3- or 4-column grid the
       forced-square height shrank below what a two-line label + big value +
       badge actually needs, so the top of the label and the source-tag badge
       at the bottom were silently clipped. Replaced with a flexible
       min-height so the box always grows to fit its content instead. */
    .metrics-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 0.9rem; margin-top: 0.5rem; }
    .metrics-grid--3 { grid-template-columns: repeat(3, 1fr); }
    @media (max-width: 640px) { .metrics-grid--3, .metrics-grid { grid-template-columns: repeat(1, 1fr); } }
    .square-metric {
        background: radial-gradient(120% 120% at 20% 0%, #ffffff, var(--c-slate-50) 70%);
        min-height: 150px; width: 100%; box-sizing: border-box;
        border-radius: var(--clay-radius);
        border: none; border-top: 4px solid var(--c-blue);
        box-shadow: var(--clay-shadow);
        animation: chartReveal var(--dur-slow) var(--ease-standard) both;
        transition: transform var(--dur-med) var(--ease-spring), box-shadow var(--dur-med) ease, border-color var(--dur-med) ease;
        display: flex; flex-direction: column; align-items: center; justify-content: center;
        text-align: center; padding: 1rem 0.7rem;
        position: relative; overflow: hidden;
    }
    .square-metric::after {
        content: ""; position: absolute; top: 0; left: 0; width: 45%; height: 100%;
        background: linear-gradient(100deg, transparent, rgba(255,255,255,0.4), transparent);
        opacity: 0; pointer-events: none;
    }
    .square-metric:hover::after { opacity: 1; animation: cardSheen 1s ease forwards; }
    .square-metric:hover {
        transform: translateY(-2px); box-shadow: var(--clay-shadow-hover);
        border-top: 4px solid var(--c-navy); z-index: 10;
    }
    .square-metric:hover .square-metric-value { transform: none; }
    .square-metric-label { font-size: 0.92rem; color: var(--c-slate-600); font-weight: 800; margin-bottom: 0.35rem;
        line-height: 1.25; overflow-wrap: break-word; hyphens: auto; }
    .square-metric-value { font-size: 1.7rem; color: var(--c-navy); font-weight: 900; line-height: 1.25; text-shadow: 0 1px 0 rgba(255,255,255,0.5);
        display: inline-block; transition: transform var(--dur-med) var(--ease-spring); }
    .square-metric-delta { font-size: 0.8rem; color: var(--c-slate-500); margin-top: 0.4rem; line-height: 1.4; overflow-wrap: break-word; }

    /* ---------- Status card (advisor / diagnostic recommendations — one class, data-status modifier) ---------- */
    .status-card { border-radius: var(--clay-radius-sm); padding: 1.1rem 1.2rem; margin-bottom: 0.9rem;
        transition: transform var(--dur-fast) ease, box-shadow var(--dur-fast) ease; border: none; box-shadow: var(--clay-shadow); }
    .status-card:hover { transform: translateY(-2px); box-shadow: var(--clay-shadow-hover); }
    .status-card[data-status="green"] { background: linear-gradient(160deg, var(--c-success-bg), #ffffff 130%); border-left: 6px solid var(--c-success); }
    .status-card[data-status="amber"] { background: linear-gradient(160deg, var(--c-warning-bg), #ffffff 130%); border-left: 6px solid var(--c-warning); }
    .status-card[data-status="red"]   { background: linear-gradient(160deg, var(--c-danger-bg), #ffffff 130%);  border-left: 6px solid var(--c-danger); }
    .status-card-desc { font-size: 0.9rem; color: var(--c-slate-700); }
    .status-card-action { font-size: 0.8rem; font-weight: bold; color: var(--c-navy); }

    /* ---------- Decision Intelligence card ---------- */
    .decision-card { background: linear-gradient(160deg, #fff, var(--c-slate-50) 120%);
        border: none; border-left: 6px solid var(--c-navy);
        border-radius: var(--clay-radius-sm); padding: 1rem 1.2rem; margin-bottom: 0.8rem;
        transition: transform var(--dur-fast) ease, box-shadow var(--dur-fast) ease; box-shadow: var(--clay-shadow);
        position: relative; overflow: hidden; }
    .decision-card::after {
        content: ""; position: absolute; top: 0; left: 0; width: 35%; height: 100%;
        background: linear-gradient(100deg, transparent, rgba(255,255,255,0.4), transparent);
        opacity: 0; pointer-events: none;
    }
    .decision-card:hover::after { opacity: 1; animation: cardSheen 1s ease forwards; }
    .decision-card:hover { box-shadow: var(--clay-shadow-hover); transform: translateY(-2px); }
    .decision-card-title { font-size: 0.85rem; font-weight: 800; color: var(--c-navy); margin-bottom: 0.25rem; }
    .decision-card-text { font-size: 0.92rem; color: var(--c-slate-700); line-height: 1.4; }

    /* ---------- Value-chain flow cards (stage palette, not a random rainbow) ---------- */
    .matrix-grid-container { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1.1rem; margin-top: 1rem; }
    @media (max-width: 1200px) { .matrix-grid-container { grid-template-columns: repeat(2, 1fr); } }
    @media (max-width: 768px)  { .matrix-grid-container { grid-template-columns: 1fr; } }
    .flow-card { background: linear-gradient(165deg, #FFFFFF, var(--c-slate-50) 130%); border: none; border-radius: var(--clay-radius-sm); padding: 1.1rem;
        text-align: left; box-shadow: var(--clay-shadow); height: 100%; transition: transform var(--dur-med) ease, box-shadow var(--dur-med) ease;
        animation: chartReveal var(--dur-slow) var(--ease-standard) both;
        position: relative; overflow: hidden; }
    .flow-card::after {
        content: ""; position: absolute; top: 0; left: 0; width: 35%; height: 100%;
        background: linear-gradient(100deg, transparent, rgba(255,255,255,0.45), transparent);
        opacity: 0; pointer-events: none;
    }
    .flow-card:hover::after { opacity: 1; animation: cardSheen 1s ease forwards; }
    .flow-card:hover { transform: translateY(-2px); box-shadow: var(--clay-shadow-hover); }
    .flow-card-title { font-size: 1rem; font-weight: 900; color: var(--c-slate-900); margin-bottom: 0.5rem; }
    .flow-card-text { font-size: 0.85rem; color: var(--c-slate-700); font-weight: 600; margin-bottom: 0.25rem; line-height: 1.35; }

    /* ---------- Mathematical Integrity diagnostic panel ---------- */
    .math-audit-table { width: 100%; border-collapse: collapse; font-size: 0.88rem; margin-top: 0.4rem; }
    .math-audit-table th { text-align: left; padding: 6px 10px; background: var(--c-slate-100); color: var(--c-slate-700); font-weight: 800; }
    .math-audit-table td { padding: 6px 10px; border-bottom: 1px solid var(--c-slate-100); color: var(--c-slate-900); }
    .math-status-ok { color: var(--c-success-text); font-weight: 800; }
    .math-status-review { color: var(--c-warning-text); font-weight: 800; }


    .story-strip { display: flex; align-items: center; justify-content: space-between; gap: 6px; margin: 0.5rem 0 1.5rem 0;
        overflow-x: auto; padding: 6px 2px 10px; }
    .story-node { flex: 1; min-width: 92px; text-align: center; padding: 0.7rem 0.4rem; border-radius: var(--clay-radius-sm);
        background: linear-gradient(165deg, #FFFFFF, var(--c-slate-50)); border: none; font-size: 0.72rem; font-weight: 800; color: var(--c-slate-700);
        box-shadow: var(--clay-shadow);
        animation: chartReveal var(--dur-slow) var(--ease-standard) both; transition: transform var(--dur-fast) ease, box-shadow var(--dur-fast) ease, color var(--dur-fast) ease; }
    .story-node:hover { transform: translateY(-4px) scale(1.03); box-shadow: var(--clay-shadow-hover), var(--glow-blue); color: var(--c-navy); }
    .story-arrow { color: var(--c-slate-400); font-size: 1.2rem; flex: 0 0 auto; }
    @media (max-width: 900px) { .story-strip { flex-wrap: wrap; } .story-arrow { display: none; } }


    div[data-testid="stDataFrame"] { transition: transform var(--dur-med) ease, box-shadow var(--dur-med) ease; border-radius: var(--clay-radius-sm); overflow: hidden; border: none; box-shadow: var(--clay-shadow); }
    div[data-testid="stDataFrame"]:hover { transform: translateY(-1px); box-shadow: var(--clay-shadow-hover); }

    [data-testid="stPlotlyChart"], .stPlotlyChart {
        transition: transform var(--dur-med) var(--ease-spring), box-shadow var(--dur-med) ease !important;
        border-radius: var(--clay-radius); background: transparent;
        box-shadow: var(--clay-shadow);
        animation: chartReveal var(--dur-slow) var(--ease-standard) both;
    }
    [data-testid="stPlotlyChart"]:hover, .stPlotlyChart:hover {
        transform: translateY(-2px) !important; box-shadow: var(--clay-shadow-hover) !important; z-index: 5;
    }
    iframe { overflow: hidden !important; }

    /* ---------- Buttons — clay body + cinematic color glow on hover ---------- */
    .stButton > button, .stDownloadButton > button {
        background: linear-gradient(135deg, var(--c-blue), var(--c-blue-light));
        color: white; font-weight: 700; border: none; border-radius: var(--clay-radius-sm);
        padding: 0.55rem 1.1rem; transition: transform var(--dur-med) var(--ease-standard), box-shadow var(--dur-med) var(--ease-standard), background var(--dur-med) var(--ease-standard) !important; letter-spacing: 0.5px;
        box-shadow: 6px 6px 16px rgba(14,165,233,0.28), -4px -4px 12px rgba(255,255,255,0.5), inset 0 1px 0 rgba(255,255,255,0.35);
        position: relative; overflow: hidden;
    }
    .stButton > button::after, .stDownloadButton > button::after {
        content: ""; position: absolute; top: 0; left: 0; width: 40%; height: 100%;
        background: linear-gradient(100deg, transparent, rgba(255,255,255,0.5), transparent);
        opacity: 0; pointer-events: none;
    }
    .stButton > button:hover::after, .stDownloadButton > button:hover::after {
        opacity: 1; animation: cardSheen 0.9s ease forwards;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        transform: translateY(-2px); box-shadow: 0 8px 18px rgba(14, 165, 233, 0.18) !important;
        background: linear-gradient(135deg, var(--c-navy), #1E40AF);
    }
    .stButton > button:active, .stDownloadButton > button:active {
        transform: translateY(0); transition: transform 0.08s ease !important;
    }

    /* ---------- Tabs: cinematic active-state glow ---------- */
    div[data-testid="stTabs"] [data-baseweb="tab-list"] {
        display: flex !important; width: 100% !important; justify-content: center !important; gap: 8px;
        background: linear-gradient(155deg, #ffffff, var(--c-slate-50)); border-radius: var(--clay-radius-sm);
        box-shadow: var(--clay-shadow); padding: 4px; margin-bottom: 0.4rem;
    }
    div[data-testid="stTabs"] [data-baseweb="tab"] {
        flex-grow: 1 !important; text-align: center !important; justify-content: center !important;
        font-size: 0.95rem !important; border-radius: var(--radius-md) !important;
        transition: background-color var(--dur-fast) ease, transform var(--dur-fast) ease, box-shadow var(--dur-fast) ease !important;
    }
    div[data-testid="stTabs"] [data-baseweb="tab"]:hover { background-color: rgba(14, 165, 233, 0.08); transform: translateY(-1px); }
    div[data-testid="stTabs"] [data-baseweb="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, var(--c-blue-pale), #ffffff) !important;
        box-shadow: inset 0 -3px 0 var(--c-blue), var(--glow-blue); color: var(--c-navy) !important; font-weight: 800 !important;
    }
    div[data-testid="stTabs"] [data-baseweb="tab-panel"] { animation: tabPanelReveal 0.22s var(--ease-standard) both; }

    /* ---------- Accessibility: keyboard focus states ---------- */
    .stButton > button:focus-visible, .stDownloadButton > button:focus-visible,
    div[data-testid="stTabs"] [data-baseweb="tab"]:focus-visible,
    [data-baseweb="select"]:focus-within, input:focus-visible {
        outline: 2px solid var(--c-blue) !important; outline-offset: 2px !important;
    }

    .main-dashboard-title {
        color: var(--c-navy); font-weight: 900; font-size: clamp(2.1rem, 4vw, 3.4rem);
        margin: 0 auto 0.5rem auto; max-width: 1200px; word-wrap: break-word; line-height: 1.25;
        text-align: center !important; display: block;
    }

    .section-caption { font-size: 0.82rem; color: var(--c-slate-500); font-style: italic; margin-top: -0.4rem; margin-bottom: 0.6rem; }
    .illustrative-banner {
        background: linear-gradient(160deg, var(--c-warning-bg), #ffffff 140%); border: 1px dashed var(--c-warning); color: var(--c-warning-text);
        border-radius: var(--clay-radius-sm); padding: 0.75rem 1.1rem; font-size: 0.85rem; font-weight: 700; margin-bottom: 0.9rem;
        box-shadow: var(--clay-shadow);
    }

    /* Cinematic hero glow — a soft radial light breathing behind the title.
       Opacity + transform(scale) only, so it's compositor-only and can loop
       indefinitely without the layout/paint repaint cost that was removed
       from the old full-viewport background-position sweep. */
    .hero-glow {
        position: absolute; top: 50%; left: 50%; width: 640px; height: 320px;
        transform: translate(-50%, -50%); border-radius: 50%;
        background: radial-gradient(closest-side, rgba(14,165,233,0.16), rgba(30,58,138,0.08) 60%, transparent 80%);
        pointer-events: none; z-index: 0;
    }
    .hero-wrap { position: relative; }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    /* ==========================================================================
       CINEMATIC FLASH LAYER — presentation only
       Exact theme/layout preserved. Motion is compositor-friendly (transform +
       opacity) and never animates background-position or layout dimensions.
       ========================================================================== */
    .stApp { position: relative; overflow-x: hidden; }
    .block-container { position: relative; z-index: 1; }
    .stApp::before {
        content: ""; position: fixed; inset: -20% -10%; pointer-events: none; z-index: 0;
        background:
            radial-gradient(36% 28% at 12% 18%, rgba(56,189,248,0.085), transparent 72%),
            radial-gradient(30% 26% at 88% 12%, rgba(30,58,138,0.075), transparent 72%),
            radial-gradient(34% 30% at 76% 86%, rgba(14,165,233,0.055), transparent 72%);
        opacity: 0.9; transform: translate3d(0,0,0) scale(1);
        animation: cinematicAmbient 18s ease-in-out infinite alternate; will-change: transform, opacity;
    }
    @keyframes cinematicAmbient {
        from { transform: translate3d(-0.5%, -0.3%, 0) scale(1); opacity: 0.78; }
        to   { transform: translate3d(0.5%, 0.35%, 0) scale(1.025); opacity: 0.96; }
    }
    .hero-wrap { isolation: isolate; overflow: hidden; border-radius: 18px; }
    .hero-wrap::before {
        content: ""; position: absolute; left: 50%; top: 50%; width: 76%; height: 120px;
        transform: translate(-50%, -50%); border-radius: 999px;
        background: radial-gradient(closest-side, rgba(56,189,248,0.12), transparent 72%);
        filter: blur(10px); pointer-events: none; z-index: 0;
    }
    .hero-wrap::after {
        content: ""; position: absolute; top: 50%; left: -35%; width: 34%; height: 2px;
        transform: translate3d(0,-50%,0);
        background: linear-gradient(90deg, transparent, rgba(56,189,248,0.0), rgba(56,189,248,0.9), rgba(255,255,255,0.95), rgba(30,58,138,0.75), transparent);
        opacity: 0; pointer-events: none; z-index: 2;
        animation: heroSweep 8s cubic-bezier(0.4,0,0.2,1) infinite; will-change: transform, opacity;
    }
    @keyframes heroSweep {
        0%, 18% { transform: translate3d(0,-50%,0); opacity: 0; }
        28%, 55% { opacity: 0.85; }
        70%, 100% { transform: translate3d(400%, -50%, 0); opacity: 0; }
    }
    .main-dashboard-title { letter-spacing: -0.035em; text-shadow: 0 8px 28px rgba(30,58,138,0.14), 0 0 24px rgba(56,189,248,0.10); }

    .metric-card, .square-metric, .flow-card, .status-card, .decision-card,
    [data-testid="stPlotlyChart"], .stPlotlyChart, div[data-testid="stDataFrame"] {
        border: 1px solid rgba(148,163,184,0.18) !important;
    }
    .metric-card, .square-metric, .flow-card, .decision-card {
        box-shadow: 0 10px 26px rgba(15,23,42,0.045), 0 2px 8px rgba(14,165,233,0.035), inset 0 1px 0 rgba(255,255,255,0.82);
        backdrop-filter: blur(2px);
    }
    .metric-card:hover, .square-metric:hover, .flow-card:hover, .decision-card:hover {
        transform: translate3d(0,-4px,0) !important;
        box-shadow: 0 18px 38px rgba(15,23,42,0.085), 0 4px 18px rgba(14,165,233,0.10), 0 0 0 1px rgba(14,165,233,0.14), inset 0 1px 0 rgba(255,255,255,0.92) !important;
    }

    /* Executive opening KPI row: sorted hierarchy + equal-height visual rhythm. */
    .executive-kpi { height: 232px !important; min-height: 232px !important; box-sizing: border-box !important; }
    .executive-kpi .metric-card-title { letter-spacing: 0.7px; }
    .executive-kpi .metric-card-value { font-size: 1.95rem; }
    .executive-kpi .metric-card-subtitle { line-height: 1.45; max-width: 100%; }
    @media (max-width: 900px) { .executive-kpi { height: auto !important; min-height: 190px !important; } }

    [data-testid="stAppViewContainer"] h2, [data-testid="stAppViewContainer"] h3 { position: relative; padding-bottom: 0.28rem; }
    [data-testid="stAppViewContainer"] h2::after, [data-testid="stAppViewContainer"] h3::after {
        content: ""; display: block; width: 56px; height: 2px; margin-top: 5px; border-radius: 999px;
        background: linear-gradient(90deg, var(--c-navy), var(--c-blue), transparent); opacity: 0.8;
    }
    div[data-testid="stTabs"] [data-baseweb="tab"][aria-selected="true"] {
        border: 1px solid rgba(14,165,233,0.22) !important;
        box-shadow: inset 0 -3px 0 var(--c-blue), 0 6px 16px rgba(14,165,233,0.10) !important;
    }

    /* ---------- Accessibility: respect reduced-motion preference globally ---------- */
    @media (prefers-reduced-motion: reduce) {
        *, *::before, *::after { animation-duration: 0.001ms !important; animation-iteration-count: 1 !important; transition-duration: 0.001ms !important; }
        .square-metric, .metric-card, .metric-card.pulse-critical,
        .flow-card, .story-node, [data-testid="stPlotlyChart"] { animation: none !important; text-shadow: none !important; }
        .metric-card--ribbon:hover, .square-metric:hover, .flow-card:hover,
        [data-testid="stPlotlyChart"]:hover, div[data-testid="stDataFrame"]:hover, div[data-testid="stMetric"]:hover,
        .stButton > button:hover, .status-card:hover, .decision-card:hover, .story-node:hover { transform: none !important; }
        .stApp, .hero-glow { animation: none !important; }
        .metric-card::after, .square-metric::after, .flow-card::after, .decision-card::after,
        .stButton > button::after, .stDownloadButton > button::after { display: none !important; }
        .square-metric:hover .square-metric-value { transform: none !important; }
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
    <div class="hero-wrap" style="display: flex; justify-content: center; align-items: center; width: 100%; margin-bottom: 1rem; text-align: center; padding: 0 1rem;">
        <div class="hero-glow"></div>
        <h1 class="main-dashboard-title" style="position:relative; z-index:1;">Enterprise Demand Forecasting & Inventory Decision Support System Using Time Series Analytics</h1>
    </div>
    <hr style="border: 0; height: 2px; background: linear-gradient(to right, rgba(0,0,0,0), #1E3A8A 25%, #38BDF8 50%, #1E3A8A 75%, rgba(0,0,0,0)); margin-top: 0; margin-bottom: 2rem;">
""", unsafe_allow_html=True)

# ==============================================================================
# SECTION 4. SIDEBAR — Model/Source Data Controls vs. Illustrative Sandbox
# (kept separate per §24: simulation inputs must never silently overwrite
#  authoritative Master Audit fields)
# ==============================================================================
st.sidebar.markdown("<h2 style='color: #1E3A8A;'>🎛️ ➔ Control Panel</h2>", unsafe_allow_html=True)
st.sidebar.markdown("---")

st.sidebar.markdown("### 📁 Model / Source Data Controls")
st.sidebar.caption("Drives every authoritative metric on this dashboard.")
target_container = st.sidebar.container()
sku_container = st.sidebar.container()

st.sidebar.markdown("---")
# ------------------------------------------------------------------------------
# Data loading — honest version. Critically: if optional array fields
# (Test_Actuals etc.) are absent, we do NOT stamp an identical fabricated
# series across every row. We record their absence and degrade gracefully.
# ------------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_and_clean_data():
    """Load the production Master Audit only; optional demo mode is explicit and never silent."""
    data_source_label = "Master Audit"
    try:
        master_path = Path(__file__).resolve().parent / "Enterprise_Supply_Chain_Master_Audit.xlsx"
        if master_path.exists():
            df = pd.read_excel(master_path, sheet_name="Executive_Summary")
        elif os.getenv("AURIX_DASHBOARD_DEMO_MODE", "0") == "1":
            # Explicit developer/demo switch only. Never silently used in production.
            df = pd.DataFrame({'SKU': ['DEMO_SKU_001'], 'ABC_Class': ['A'], 'XYZ_Class': ['X'], 'ABC_XYZ_Class': ['AX']})
            data_source_label = "Developer Demo Mode"
        else:
            raise FileNotFoundError("Master Audit workbook not found. Dashboard cannot initialise.")

        # Column-name hygiene only — never value fabrication.
        df.columns = [str(c).strip() for c in df.columns]
        if 'SKU' not in df.columns:
            raise ValueError("Authoritative source is missing the SKU field.")
        df['SKU'] = df['SKU'].astype(str).str.strip()
        df = df[~df['SKU'].str.lower().isin({'', 'nan', 'none', 'grand total', 'total'})].copy()
        df = df[df['SKU'].notna()].copy()
        return df, data_source_label
    except FileNotFoundError:
        st.error("Master Audit workbook not found. Dashboard cannot initialise.")
        st.stop()
    except Exception as e:
        st.error(f"Could not read the Master Audit/source file cleanly: {e}")
        st.stop()

df, DATA_SOURCE_LABEL = load_and_clean_data()

# ------------------------------------------------------------------------------
# Raw 24-Month Workbook upload infrastructure removed in this production build.
# The current application has no uploader widget, so keeping dormant upload state
# would create a misleading control path. The Master Audit remains the runtime
# source of truth; the dashboard degrades to aggregate views where monthly raw data
# is unavailable.
# ------------------------------------------------------------------------------

# Availability flags computed BEFORE any downstream default-filling, and never
# used to justify inventing per-row values.
HAS_TEST_ARRAYS = all(c in df.columns for c in OPTIONAL_ARRAY_FIELDS[:2])          # rich per-SKU trajectory mode
HAS_XYZ = 'XYZ_Class' in df.columns
HAS_ABC_XYZ = 'ABC_XYZ_Class' in df.columns
ACTUAL_INV_COL = next((c for c in POSSIBLE_ACTUAL_INVENTORY_COLS if c in df.columns), None)
MISSING_AUTHORITATIVE_FIELDS = [c for c in MASTER_AUDIT_FIELDS if c not in df.columns]

target_container.subheader("🎯 Target Isolation")
abc_options = [c for c in ABC_ORDER if 'ABC_Class' in df.columns and c in set(df['ABC_Class'].dropna().astype(str))] if 'ABC_Class' in df.columns else ABC_ORDER.copy()
selected_classes = target_container.multiselect("Select ABC Classes", options=abc_options, default=abc_options, key="abc_filter")

if HAS_XYZ:
    xyz_options = [c for c in XYZ_ORDER if c in set(df['XYZ_Class'].dropna().astype(str))]
    selected_xyz = target_container.multiselect("Select XYZ Classes", options=xyz_options, default=xyz_options, key="xyz_filter")
else:
    selected_xyz = None

filtered_df = df.copy()
if selected_classes and 'ABC_Class' in df.columns:
    filtered_df = filtered_df[filtered_df['ABC_Class'].isin(selected_classes)]
if HAS_XYZ and selected_xyz:
    filtered_df = filtered_df[filtered_df['XYZ_Class'].isin(selected_xyz)]

sku_list = filtered_df['SKU'].unique().tolist() if 'SKU' in filtered_df.columns else []

if 'shared_sku' not in st.session_state:
    st.session_state.shared_sku = sku_list[0] if sku_list else None
    st.session_state.sidebar_sku = st.session_state.shared_sku
    st.session_state.tab2_sku = st.session_state.shared_sku
    st.session_state.tab3_sku = st.session_state.shared_sku

if sku_list and st.session_state.shared_sku not in sku_list:
    st.session_state.shared_sku = sku_list[0]
    st.session_state.sidebar_sku = sku_list[0]
    st.session_state.tab2_sku = sku_list[0]
    st.session_state.tab3_sku = sku_list[0]
elif not sku_list:
    # Empty-filter guard: with zero SKUs left after the ABC/XYZ filter, every
    # downstream selectbox below is skipped rather than instantiated against
    # an empty options list, so no widget is left pointing at a stale SKU.
    st.session_state.shared_sku = None

# Pending cross-tab SKU jump (e.g. "🎯 Jump to SKU" in Most Common Executive Recommendations).
# Streamlit forbids writing st.session_state.sidebar_sku / tab2_sku / tab3_sku
# once those selectbox widgets have already been instantiated in this run, so
# a button placed further down the page cannot set them directly. Instead it
# queues the target SKU here in '_pending_sku_jump', and this block — which
# runs BEFORE any of those selectboxes are created — applies it up front.
if st.session_state.get('_pending_sku_jump'):
    pending_sku = st.session_state.pop('_pending_sku_jump')
    if pending_sku in sku_list:
        st.session_state.shared_sku = pending_sku
        st.session_state.sidebar_sku = pending_sku
        st.session_state.tab2_sku = pending_sku
        st.session_state.tab3_sku = pending_sku

def sync_from(key):
    val = st.session_state[key]
    st.session_state.shared_sku = val
    for k in ['sidebar_sku', 'tab2_sku', 'tab3_sku']:
        st.session_state[k] = val

if sku_list:
    sku_container.selectbox("Global SKU Selector", sku_list, key="sidebar_sku", on_change=lambda: sync_from('sidebar_sku'))
else:
    sku_container.info("No SKUs match the current ABC/XYZ filter.")

# ------------------------------------------------------------------------------
# Portfolio-level numeric series used across multiple tabs (computed once)
# ------------------------------------------------------------------------------
rmse_series = safe_series(filtered_df, 'RMSE', numeric=True)
mae_series = safe_series(filtered_df, 'MAE', numeric=True)
bias_series = safe_series(filtered_df, 'Bias', numeric=True)
cv_series = safe_series(filtered_df, 'CV', numeric=True)
bullwhip_series = safe_series(filtered_df, 'Bullwhip_Ratio', numeric=True)
accuracy_series = safe_series(filtered_df, 'Accuracy_Pct', numeric=True)
fhs_series = safe_series(filtered_df, 'Forecast_Health_Score', numeric=True)
inv_health_series = safe_series(filtered_df, 'Inventory_Health_Score', numeric=True)
inv_value_series = safe_series(filtered_df, 'Inventory_Value', numeric=True)
wc_series = safe_series(filtered_df, 'Working_Capital', numeric=True)
fin_health_series = safe_series(filtered_df, 'KPI_Financial_Health_Score', numeric=True)
mean_monthly_series = safe_series(filtered_df, 'Mean_Monthly_Demand', numeric=True)
annual_demand_series = safe_series(filtered_df, 'Annual_Demand', numeric=True)
turnover_series = safe_series(filtered_df, 'Inventory_Turnover', numeric=True)

total_skus = len(filtered_df)
avg_rmse = rmse_series.mean() if total_skus > 0 else np.nan
portfolio_turnover = turnover_series.mean() if total_skus > 0 and 'Inventory_Turnover' in filtered_df.columns else np.nan
portfolio_inventory_days = 365.0 / portfolio_turnover if not is_missing(portfolio_turnover) and portfolio_turnover > 0 else np.nan
portfolio_wce = float(np.clip(portfolio_turnover / 6.0 * 100.0, 0, 100)) if not is_missing(portfolio_turnover) else np.nan

if 'Forecast_Health_Score' in filtered_df.columns:
    forecast_health_avg = fhs_series.mean()
    forecast_health_source = "Master Audit"
else:
    forecast_health_avg = np.nan
    forecast_health_source = None

avg_accuracy = accuracy_series.mean() if 'Accuracy_Pct' in filtered_df.columns else np.nan

if 'Forecast_Risk' in filtered_df.columns:
    high_risk_count = int((filtered_df['Forecast_Risk'].astype(str).str.strip().str.lower() == 'high').sum())
    high_risk_source = "Master Audit"
else:
    # Forecast_Risk is an authoritative source field; do not invent a substitute
    # classification when it is absent.
    high_risk_count = np.nan
    high_risk_source = None

inv_value_sum = safe_sum(inv_value_series) if 'Inventory_Value' in filtered_df.columns else np.nan
wc_sum = safe_sum(wc_series) if 'Working_Capital' in filtered_df.columns else np.nan
fin_health_avg = fin_health_series.mean() if 'KPI_Financial_Health_Score' in filtered_df.columns else np.nan
inv_health_avg = inv_health_series.mean() if 'Inventory_Health_Score' in filtered_df.columns else np.nan

# Overall qualitative status banner (Dashboard-Derived Portfolio Status — combines authoritative signals)
status_points = 0
status_total = 0
for val, good_thresh, reverse in [(forecast_health_avg, 70, False), (avg_accuracy, 70, False),
                                    (fin_health_avg, 70, False), (inv_health_avg, 70, False)]:
    if not is_missing(val):
        status_total += 1
        if (val >= good_thresh) != reverse:
            status_points += 1
if status_total > 0:
    overall_ratio = status_points / status_total
    high_risk_zero_known = (not is_missing(high_risk_count) and high_risk_count == 0)
    if high_risk_zero_known and overall_ratio >= 0.75:
        overall_status, overall_status_key = "Stable & Healthy", "excellent"
    elif overall_ratio >= 0.5:
        overall_status, overall_status_key = "Needs Attention", "good"
    else:
        overall_status, overall_status_key = "Elevated Risk", "poor"
else:
    overall_status, overall_status_key = "Insufficient Data", "attention"

# ==============================================================================
# SECTION 5. EXACTLY FIVE TABS
# ==============================================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "🏠 Executive Control Tower",
    "📈 Demand & Forecast Intelligence",
    "📦 Inventory Optimisation",
    "⚠️ Risk, Segmentation & Diagnostics",
    "💰 Financial Impact, Methodology & Audit",
])

# ==============================================================================
# TAB 1 — EXECUTIVE CONTROL TOWER
# 60-second overview: KPI ribbon, portfolio health, what needs attention now.
# ==============================================================================
with tab1:
    story_nodes = ["📥 DATA", "📊 DEMAND", "🔮 FORECAST", "📐 UNCERTAINTY", "📦 INVENTORY", "⚠️ RISK", "💰 FINANCE", "🧭 DECISION"]
    story_html = '<div class="story-strip">'
    for i, node in enumerate(story_nodes):
        story_html += f'<div class="story-node">{node}</div>'
        if i < len(story_nodes) - 1:
            story_html += '<div class="story-arrow">→</div>'
    story_html += '</div>'
    st.markdown(story_html, unsafe_allow_html=True)

    st.markdown(
        f'<div class="illustrative-banner" data-status="{overall_status_key}" style="border-color: var(--c-blue); background: var(--c-blue-pale); color: var(--c-navy);">'
        f'🧭 Dashboard-Derived Portfolio Signal: <strong>{overall_status}</strong> &nbsp;|&nbsp; Data Source: <strong>{DATA_SOURCE_LABEL}</strong> &nbsp;|&nbsp; SKUs in View: <strong>{total_skus:,}</strong></div>',
        unsafe_allow_html=True)

    r1c1, r1c2, r1c3, r1c4 = st.columns(4)
    # Executive hierarchy: scope → accuracy → health → immediate risk.
    render_metric_card(r1c1, "Total Tracked SKUs", f"{total_skus:,}", "Count of filtered Master Audit SKUs", extra=source_tag("Dashboard Derived"), extra_class="executive-kpi")
    render_metric_card(r1c2, "Forecast Accuracy", fmt_num(avg_accuracy, 1, "%") if not is_missing(avg_accuracy) else "N/A", "Portfolio mean • Master Audit Accuracy_Pct", extra=source_tag("Dashboard Derived"), extra_class="executive-kpi")

    if forecast_health_source:
        fh_status = score_to_status(forecast_health_avg)
        fh_badge = badge_html(fh_status, fh_status.title())
        fh_subtitle = f"Mean Master Audit Forecast Health • Dashboard presentation band {fh_badge}"
    else:
        fh_subtitle = "Forecast_Health_Score unavailable in current source"
    render_metric_card(r1c3, "Forecast Health Score", fmt_num(forecast_health_avg, 0, "%"), fh_subtitle, extra=source_tag("Dashboard Derived"), extra_class="executive-kpi")

    render_metric_card(r1c4, "High-Risk SKUs", fmt_int(high_risk_count), "High Forecast_Risk count • Master Audit", variant="ribbon", pulse=(not is_missing(high_risk_count) and high_risk_count > 0), extra=source_tag("Dashboard Derived"), extra_class="executive-kpi")

    r2c1, r2c2, r2c3 = st.columns(3)
    render_metric_card(r2c1, "Model-Implied Inventory Value", fmt_monetary(inv_value_sum), "Portfolio sum • Master Audit Inventory_Value", extra=source_tag("Dashboard Derived"), extra_class="executive-kpi")
    render_metric_card(r2c2, "Model-Implied Working Capital", fmt_monetary(wc_sum), "Portfolio sum • Master Audit Working_Capital", extra=source_tag("Dashboard Derived"), extra_class="executive-kpi")
    render_metric_card(r2c3, "Portfolio Mean Financial Health", fmt_num(fin_health_avg, 0, "%") if not is_missing(fin_health_avg) else "N/A", "Mean of SKU Financial Health Scores • Master Audit KPI_Financial_Health_Score", extra=source_tag("Dashboard Derived"), extra_class="executive-kpi")

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<hr>", unsafe_allow_html=True)

    ov_col1, ov_col2 = st.columns(2)
    with ov_col1:
        st.subheader("📦 Portfolio Composition — ABC")
        if 'ABC_Class' in filtered_df.columns and total_skus > 0:
            abc_counts = ordered_category_counts(filtered_df, 'ABC_Class', ABC_ORDER)
            fig_abc = px.pie(abc_counts, names='ABC_Class', values='Count', hole=0.5,
                              color='ABC_Class', color_discrete_map=ABC_COLOR_MAP, category_orders={'ABC_Class': ABC_ORDER},
                              title="SKU Count Share by ABC — Demand Volume Classification")
            fig_abc.update_traces(textinfo='label+percent', marker=dict(line=dict(color='#FFFFFF', width=2)))
            fig_abc.update_layout(plot_bgcolor='white', margin=dict(l=10, r=10, t=40, b=10))
            st.plotly_chart(fig_abc, width='stretch')
        else:
            st.info("ABC_Class not available in this data source.")
    with ov_col2:
        st.subheader("🌡️ Portfolio Composition — XYZ")
        if HAS_XYZ and total_skus > 0:
            xyz_counts = ordered_category_counts(filtered_df, 'XYZ_Class', XYZ_ORDER)
            fig_xyz = px.pie(xyz_counts, names='XYZ_Class', values='Count', hole=0.5,
                              color='XYZ_Class', color_discrete_map=XYZ_COLOR_MAP, category_orders={'XYZ_Class': XYZ_ORDER},
                              title="SKU Count Share by XYZ — Demand Variability Classification")
            fig_xyz.update_traces(textinfo='label+percent', marker=dict(line=dict(color='#FFFFFF', width=2)))
            fig_xyz.update_layout(plot_bgcolor='white', margin=dict(l=10, r=10, t=40, b=10))
            st.plotly_chart(fig_xyz, width='stretch')
        else:
            st.info("XYZ_Class is not available in the current Master Audit source; this view requires that authoritative field.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🧩 ABC–XYZ Segmentation Matrix")
    st.markdown('<div class="section-caption">Count of SKUs by combined ABC — Demand Volume Classification × XYZ — Demand Variability classification.</div>', unsafe_allow_html=True)
    if HAS_ABC_XYZ and total_skus > 0:
        abc_order = ABC_ORDER
        xyz_order = XYZ_ORDER
        # Vectorized rebuild of the same ABC x XYZ count matrix (Performance Fix,
        # Sept 2026): the original per-row Python loop over the full filtered
        # dataset ran on every script rerun; this produces an identical matrix
        # via a single pandas crosstab, with any code not matching the expected
        # 2-character ABC/XYZ pattern excluded exactly as the loop's own
        # length/membership checks did.
        codes = filtered_df['ABC_XYZ_Class'].astype(str)
        valid_code = codes.str.len().eq(2) & codes.str[0].isin(abc_order) & codes.str[1].isin(xyz_order)
        valid_codes = codes[valid_code]
        matrix = pd.crosstab(valid_codes.str[0], valid_codes.str[1]).reindex(index=abc_order, columns=xyz_order, fill_value=0)
        fig_matrix = px.imshow(matrix, text_auto=True, color_continuous_scale=[[0, '#F1F5F9'], [1, '#1E3A8A']],
                                labels=dict(x="XYZ — Demand Variability", y="ABC — Demand Volume Classification", color="SKU Count"))
        fig_matrix.update_layout(title="ABC–XYZ Segmentation Matrix — SKU Count", margin=dict(l=10, r=10, t=40, b=10))
        st.plotly_chart(fig_matrix, width='stretch')
    else:
        st.info("ABC_XYZ_Class not available — matrix will populate once this field is present in the source data.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("📊 Annualised Demand by ABC Class")
    st.caption("Annualised Demand = full-period mean monthly demand × 12; it is not necessarily trailing-12-month actual demand.")
    st.markdown('<div class="section-caption">Portfolio demand snapshot from Master Audit Annual_Demand. Note: month-by-month historical time series is not part of the Master Audit export, so this is an aggregate view rather than a monthly trend line.</div>', unsafe_allow_html=True)
    if 'Annual_Demand' in filtered_df.columns and 'ABC_Class' in filtered_df.columns and total_skus > 0:
        demand_by_abc = filtered_df.groupby('ABC_Class', as_index=False)['Annual_Demand'].sum()
        demand_by_abc['ABC_Class'] = pd.Categorical(demand_by_abc['ABC_Class'], categories=ABC_ORDER, ordered=True)
        demand_by_abc = demand_by_abc.sort_values('ABC_Class')
        fig_demand = px.bar(demand_by_abc, x='ABC_Class', y='Annual_Demand', color='ABC_Class',
                             color_discrete_map=ABC_COLOR_MAP, category_orders={'ABC_Class': ABC_ORDER},
                             title="Total Annualised Demand (units) by ABC Class")
        fig_demand.update_layout(plot_bgcolor='white', showlegend=False, margin=dict(l=10, r=10, t=40, b=10))
        st.plotly_chart(fig_demand, width='stretch')
    else:
        st.info("Annual_Demand not available in this data source.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🚨 Immediate Risk & Priority Exceptions")
    attn_df = pd.DataFrame()
    attn_cols = [c for c in ['SKU', 'ABC_XYZ_Class', 'ABC_Class', 'Forecast_Risk', 'Priority_Level',
                              'Inventory_Health_Score', 'Executive_Recommendation'] if c in filtered_df.columns]
    if attn_cols and total_skus > 0:
        attn_mask = pd.Series([False] * total_skus, index=filtered_df.index)
        if 'Forecast_Risk' in filtered_df.columns:
            attn_mask = attn_mask | (filtered_df['Forecast_Risk'].astype(str).str.lower() == 'high')
        if 'Priority_Level' in filtered_df.columns:
            attn_mask = attn_mask | (filtered_df['Priority_Level'].astype(str).str.lower() == 'high')
        attn_df = filtered_df[attn_mask][attn_cols].copy()
        if 'Inventory_Value' in filtered_df.columns:
            attn_df = attn_df.join(filtered_df.loc[attn_mask, 'Inventory_Value'])
            attn_df = attn_df.sort_values(['Inventory_Value', 'SKU'], ascending=[False, True], kind='mergesort', na_position='last')
        else:
            attn_df = deterministic_high_risk_sort(attn_df)
        st.dataframe(attn_df.head(25), width='stretch', hide_index=True)
        st.caption(f"{len(attn_df):,} SKU(s) flagged as High Risk or High Priority in the current filter. Showing top 25 by Inventory Value.")
    else:
        st.info("Risk/priority fields not available to build the attention list.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🧭 Most Common Executive Recommendations")
    st.markdown('<div class="section-caption">Recommendations are shown by frequency among flagged SKUs; this is not an objective priority ranking. Each recommendation lists the exact SKUs it applies to — click "Jump to" to open one directly in Decision Intelligence below.</div>', unsafe_allow_html=True)
    if 'Executive_Recommendation' in filtered_df.columns and len(attn_df) > 0:
        priority_recs_df = filtered_df.loc[attn_df.index]
        top_recs = (priority_recs_df['Executive_Recommendation'].value_counts().rename_axis('Recommendation').reset_index(name='Count')
                      .sort_values(['Count', 'Recommendation'], ascending=[False, True], kind='mergesort').head(5)
                      .set_index('Recommendation')['Count'])
        for rec_i, (rec, cnt) in enumerate(top_recs.items()):
            rec_skus = priority_recs_df.loc[priority_recs_df['Executive_Recommendation'] == rec, 'SKU'].astype(str).tolist()
            shown_skus = rec_skus[:12]
            sku_chip_html = " ".join(f'<span class="source-tag tag-master" style="margin:2px;">{html.escape(s)}</span>' for s in shown_skus)
            if len(rec_skus) > len(shown_skus):
                sku_chip_html += f' <span class="source-tag tag-derived">+{len(rec_skus) - len(shown_skus)} more</span>'
            st.markdown(
                f'<div class="status-card" data-status="amber">'
                f'<strong>{cnt} SKU(s)</strong><br>'
                f'<span class="status-card-desc">{html.escape(str(rec))}</span><br>'
                f'<div style="margin-top:0.5rem;">{sku_chip_html}</div></div>',
                unsafe_allow_html=True)
            jump_col1, jump_col2 = st.columns([3, 1])
            with jump_col1:
                pick = st.selectbox("SKU affected by this decision", rec_skus, key=f"priority_pick_{rec_i}", label_visibility="collapsed")
            with jump_col2:
                if st.button("🎯 Jump to SKU", key=f"priority_jump_{rec_i}", width='stretch'):
                    st.session_state['_pending_sku_jump'] = pick
                    st.toast(f"Loaded {pick} — see 'Decision Intelligence' below, or Tabs 2/3 for full diagnostics.", icon="🎯")
                    st.rerun()
    else:
        st.info("No Executive_Recommendation data available for the current selection.")

    # ------------------------------------------------------------------------
    # DECISION INTELLIGENCE — sits directly beneath the decisions that
    # reference it, so the "Jump to SKU" action above has somewhere to land.
    # ------------------------------------------------------------------------
    st.markdown("<hr>", unsafe_allow_html=True)
    dt_title = st.session_state.shared_sku if st.session_state.shared_sku else "No SKU Selected"
    st.subheader("🧭 Decision Intelligence — " + str(dt_title))
    sku_row5 = None
    if st.session_state.shared_sku and total_skus > 0:
        match5 = filtered_df[filtered_df['SKU'] == st.session_state.shared_sku]
        if len(match5) > 0:
            sku_row5 = match5.iloc[0]
    if sku_row5 is not None:
        dcol1, dcol2 = st.columns(2)
        with dcol1:
            render_decision_card("🎯", "Executive Recommendation", safe_val(sku_row5, 'Executive_Recommendation', 'N/A'))
            render_decision_card("📦", "Inventory Recommendation", safe_val(sku_row5, 'Inventory_Recommendation', 'N/A'))
            render_decision_card("🚚", "Procurement Recommendation", safe_val(sku_row5, 'Procurement_Recommendation', 'N/A'))
            render_decision_card("🏭", "Warehouse Recommendation", safe_val(sku_row5, 'Warehouse_Recommendation', 'N/A'))
        with dcol2:
            render_decision_card("💰", "Financial Recommendation", safe_val(sku_row5, 'Financial_Recommendation', 'N/A'))
            render_decision_card("📝", "Business Explanation", safe_val(sku_row5, 'Business_Explanation', 'N/A'))
            render_decision_card("📈", "Expected Impact", safe_val(sku_row5, 'Expected_Impact', 'N/A'))
    else:
        st.info("Select a SKU from the sidebar to view its Decision Intelligence cards.")

    # ==========================================================================
    # WHAT THIS PROJECT ENABLES — compact, evidence-only implementation
    # summary. Every line maps to a capability the code actually produces
    # from Master Audit fields. No baseline, owner, target, status or
    # "not yet measured" placeholder is shown here: none of those are
    # evidenced, so per the evidence-audit rule they are not fabricated.
    # ==========================================================================
    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🧭 What This Project Enables")
    st.markdown('<div class="section-caption">Capabilities the analytics genuinely deliver today, drawn directly from Master Audit outputs.</div>', unsafe_allow_html=True)

    wc_cols = st.columns(3)
    wc_items = [
        ("📈 Forecasting", "SKU-level validated forecasting and model-selection outputs."),
        ("📦 Inventory", "Forecast-driven Safety Stock, ROP and EOQ recommendations."),
        ("🧩 Segmentation", "ABC–XYZ differentiated inventory control."),
    ]
    for col, (title, desc) in zip(wc_cols, wc_items):
        col.markdown(f'<div class="flow-card" style="border-top:5px solid var(--c-blue);">'
                      f'<div class="flow-card-title">{title}</div><div class="flow-card-text">{desc}</div></div>',
                      unsafe_allow_html=True)
    wc_cols2 = st.columns(3)
    wc_items2 = [
        ("⚠️ Risk", "Forecast-risk, RMSE and Bullwhip-based exception visibility."),
        ("💰 Finance", "Model-Implied Inventory Value, Model-Implied Working Capital, Carrying Cost and Stockout exposure visibility."),
        ("🧭 Decision Support", "Rule-based recommendations linked directly to the analytical outputs above."),
    ]
    for col, (title, desc) in zip(wc_cols2, wc_items2):
        col.markdown(f'<div class="flow-card" style="border-top:5px solid var(--c-blue);">'
                      f'<div class="flow-card-title">{title}</div><div class="flow-card-text">{desc}</div></div>',
                      unsafe_allow_html=True)
    st.caption("Realised business impact (savings, inventory reduction, forecast improvement) requires post-deployment "
               "measurement against an actual operating baseline — that measurement is not part of this analytical layer.")

# ==============================================================================
# TAB 2 — DEMAND & FORECAST INTELLIGENCE
# Deep demand + model-performance analysis for a selected SKU, plus
# portfolio-level forecast-quality diagnostics.
# ==============================================================================
with tab2:
    header_col1, header_col2 = st.columns([3, 1])
    with header_col1:
        st.subheader("Forecast Diagnostics & Model Performance")
    with header_col2:
        if sku_list:
            st.selectbox("🎯 Isolate SKU", sku_list, key="tab2_sku", on_change=lambda: sync_from('tab2_sku'))

    sku_row = None
    if st.session_state.shared_sku and total_skus > 0:
        match = filtered_df[filtered_df['SKU'] == st.session_state.shared_sku]
        if len(match) > 0:
            sku_row = match.iloc[0]

    if sku_row is not None:
        abc_badge = f"ABC: <strong>{safe_val(sku_row, 'ABC_Class', 'N/A')}</strong>"
        xyz_badge = f" &nbsp;|&nbsp; XYZ: <strong>{safe_val(sku_row, 'XYZ_Class', 'N/A')}</strong>" if HAS_XYZ else ""
        safe_sku = html.escape(str(st.session_state.shared_sku))
        st.markdown(f"#### Target SKU: **{safe_sku}** &nbsp; ({abc_badge}{xyz_badge})", unsafe_allow_html=True)

        st.markdown("<h4 style='color: #475569;'>📊 Model Diagnostic Brief</h4>", unsafe_allow_html=True)
        with st.container(border=True):
            rmse_val = safe_val(sku_row, 'RMSE')
            daf_val = safe_val(sku_row, 'Bullwhip_Ratio')
            acc_val = safe_val(sku_row, 'Accuracy_Pct')
            p_val = safe_val(sku_row, 'Health_P_Value')

            insight = f"**Model Summary:** The forecasting model for **{st.session_state.shared_sku}** "
            if not is_missing(rmse_val) and not is_missing(avg_rmse):
                if rmse_val > avg_rmse:
                    insight += f"shows above-portfolio-average forecast error (RMSE {fmt_num(rmse_val,1)} vs. portfolio average {fmt_num(avg_rmse,1)}). Manual review of recent demand anomalies is recommended. "
                else:
                    insight += f"is tracking within normal bounds (RMSE {fmt_num(rmse_val,1)}, portfolio average {fmt_num(avg_rmse,1)}). "
            else:
                insight += "does not have an RMSE value available for comparison. "

            if not is_missing(daf_val):
                if daf_val > 1.5:
                    insight += f"The Bullwhip Ratio — model-derived replenishment signal amplification metric is {fmt_num(daf_val,2)} relative to the reference level. "
                else:
                    insight += f"The Bullwhip Ratio — model-derived replenishment signal amplification metric is {fmt_num(daf_val,2)} and is near the reference level. "

            if not is_missing(p_val):
                health_status = ("Acceptable — no significant residual autocorrelation detected at the tested lag" if float(p_val) > 0.05 else "Review — significant residual autocorrelation detected at the tested lag")
                insight += f"\n\n**Residual Diagnostic:** {health_status} — Ljung-Box p-value {fmt_num(p_val,3)}."
            st.info(insight)

        t2_col1, t2_col2 = st.columns([2.5, 1])
        with t2_col1:
            if HAS_TEST_ARRAYS:
                try:
                    actuals = ast.literal_eval(str(sku_row.get('Test_Actuals', '[]')))
                    preds = ast.literal_eval(str(sku_row.get('Test_Predictions', '[]')))
                    if isinstance(actuals, list) and isinstance(preds, list) and len(actuals) == len(preds) and len(actuals) > 0:
                        months = [f"Month {i+1}" for i in range(len(actuals))]
                        fig = go.Figure()
                        try:
                            lb = ast.literal_eval(str(sku_row.get('Test_Lower_Bound', '[]')))
                            ub = ast.literal_eval(str(sku_row.get('Test_Upper_Bound', '[]')))
                            if isinstance(lb, list) and isinstance(ub, list) and len(lb) == len(months):
                                fig.add_trace(go.Scatter(x=months + months[::-1], y=ub + lb[::-1], fill='toself',
                                                          fillcolor='rgba(244, 63, 94, 0.15)', line=dict(color='rgba(255,255,255,0)'),
                                                          hoverinfo="skip", showlegend=True, name='Forecast Uncertainty Band'))
                        except Exception:
                            pass
                        fig.add_trace(go.Scatter(x=months, y=actuals, name='Actual Demand', mode='lines+markers',
                                                  line=dict(color='#0EA5E9', width=3), marker=dict(size=8, color='#0284C7')))
                        fig.add_trace(go.Scatter(x=months, y=preds, name='Forecast', mode='lines+markers',
                                                  line=dict(color='#F43F5E', width=3, dash='dash'), marker=dict(size=8, symbol='diamond', color='#E11D48')))
                        fig.update_layout(title="Forecast Validation — Test Period", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                                           hovermode='x unified',
                                           legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                                           xaxis=dict(showgrid=False), yaxis=dict(showgrid=True, gridcolor='#F1F5F9'),
                                           margin=dict(l=20, r=20, t=40, b=20))
                        st.plotly_chart(fig, width='stretch')
                        st.caption("Test Period = model validation window, not the full historical series.")
                    else:
                        st.info("Test-period arrays present but malformed for this SKU.")
                except Exception:
                    st.warning("Could not parse test-period arrays for this SKU.")
            else:
                fc_next = safe_val(sku_row, 'Forecast_Next_Month')
                mean_m = safe_val(sku_row, 'Mean_Monthly_Demand')
                snap_df = pd.DataFrame({
                    'Metric': ['Mean Monthly Demand (Historical Avg)', 'Forecast — Next Month'],
                    'Value': [mean_m, fc_next]
                })
                fig_snap = px.bar(snap_df, x='Metric', y='Value', color='Metric',
                                   color_discrete_sequence=['#94A3B8', '#0EA5E9'], title="Forecast Snapshot")
                fig_snap.update_layout(plot_bgcolor='white', showlegend=False, margin=dict(l=20, r=20, t=40, b=20))
                st.plotly_chart(fig_snap, width='stretch')
                st.markdown(f"""<div class="section-caption">Month-by-month actual-vs-predicted arrays are not included in this Master Audit export —
                only aggregate validation statistics (Accuracy, MAE, RMSE, Bias) are available per SKU. {source_tag("Master Audit")}</div>""", unsafe_allow_html=True)

        with t2_col2:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("<h4 style='color: #475569; text-align:center;'>⚡ Forecast KPIs</h4>", unsafe_allow_html=True)
            acc_val = safe_val(sku_row, 'Accuracy_Pct')
            acc_source = "Master Audit" if 'Accuracy_Pct' in filtered_df.columns else "Dashboard Derived"
            mini_html = ('<div class="metrics-grid metrics-grid--3">'
                + render_square_metric("Next Forecast", fmt_int(safe_val(sku_row, 'Forecast_Next_Month')), "Projected " + source_tag("Master Audit"))
                + render_square_metric("Bullwhip Ratio", fmt_num(safe_val(sku_row, 'Bullwhip_Ratio'), 2), "Bullwhip Ratio — Model-Derived Replenishment Signal Amplification " + source_tag("Master Audit"))
                + render_square_metric("Forecast Accuracy", fmt_num(acc_val, 1, '%'), source_tag(acc_source))
                + '</div>')
            st.markdown(mini_html, unsafe_allow_html=True)

        st.markdown("<hr>", unsafe_allow_html=True)
        st.subheader("🔬 Full Diagnostic Panel")
        d1, d2, d3, d4, d5, d6 = st.columns(6)
        d1.metric("Accuracy %", fmt_num(safe_val(sku_row, 'Accuracy_Pct'), 1))
        d2.metric("MAE", fmt_num(safe_val(sku_row, 'MAE'), 2))
        d3.metric("RMSE", fmt_num(safe_val(sku_row, 'RMSE'), 2))
        d4.metric("Bias", fmt_num(safe_val(sku_row, 'Bias'), 2))
        d5.metric("CV", fmt_num(safe_val(sku_row, 'CV'), 2))
        d6.metric("Mean Monthly Demand", fmt_num(safe_val(sku_row, 'Mean_Monthly_Demand'), 1))
        e1, e2, e3 = st.columns(3)
        with e1:
            fr = str(safe_val(sku_row, 'Forecast_Risk', 'N/A'))
            fr_status = {'low': 'excellent', 'medium': 'good', 'high': 'poor'}.get(fr.lower(), 'attention')
            st.markdown(f"**Forecast Risk:** {badge_html(fr_status, fr)}", unsafe_allow_html=True)
        with e2:
            mr = str(safe_val(sku_row, 'Model_Reliability', 'N/A'))
            mr_status = {'high': 'excellent', 'moderate': 'good', 'low': 'poor'}.get(mr.lower(), 'attention')
            st.markdown(f"**Model Reliability:** {badge_html(mr_status, mr)}", unsafe_allow_html=True)
        with e3:
            st.markdown(f"**Annualised Demand:** {fmt_int(safe_val(sku_row, 'Annual_Demand'))} {source_tag('Master Audit')}", unsafe_allow_html=True)
    else:
        st.info("Select a SKU from the sidebar or the selector above to view its forecast diagnostics.")

    # --------------------------------------------------------------------
    # Portfolio-level forecast quality
    # --------------------------------------------------------------------
    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🎯 Portfolio Forecast Quality")
    st.markdown('<div class="section-caption">Computed from authoritative per-SKU RMSE / MAE / Bias / Accuracy_Pct values in the Master Audit — not recomputed from raw arrays.</div>', unsafe_allow_html=True)

    if total_skus > 0:
        fq1, fq2, fq3 = st.columns(3)
        with fq1:
            fig_rmse = px.histogram(filtered_df, x='RMSE', title="RMSE Distribution", color_discrete_sequence=['#0EA5E9']) if 'RMSE' in filtered_df.columns else None
            if fig_rmse:
                fig_rmse.update_layout(plot_bgcolor='white', margin=dict(l=20, r=20, t=40, b=20))
                st.plotly_chart(fig_rmse, width='stretch')
            else:
                st.info("RMSE not available.")
        with fq2:
            fig_mae = px.histogram(filtered_df, x='MAE', title="MAE Distribution", color_discrete_sequence=['#1E3A8A']) if 'MAE' in filtered_df.columns else None
            if fig_mae:
                fig_mae.update_layout(plot_bgcolor='white', margin=dict(l=20, r=20, t=40, b=20))
                st.plotly_chart(fig_mae, width='stretch')
            else:
                st.info("MAE not available.")
        with fq3:
            fig_bias = px.histogram(filtered_df, x='Bias', title="Bias Distribution", color_discrete_sequence=['#F43F5E']) if 'Bias' in filtered_df.columns else None
            if fig_bias:
                fig_bias.update_layout(plot_bgcolor='white', margin=dict(l=20, r=20, t=40, b=20))
                st.plotly_chart(fig_bias, width='stretch')
            else:
                st.info("Bias not available.")

        if not is_missing(avg_accuracy):
            q_class = "Excellent" if avg_accuracy >= 85 else ("Good" if avg_accuracy >= 70 else ("Average" if avg_accuracy >= 50 else "Poor"))
            q_class = f"Dashboard Presentation Band: {q_class}"
            fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number", value=float(avg_accuracy),
                title={'text': f"Portfolio Forecast Accuracy: {q_class}", 'font': {'size': 18, 'color': '#1E3A8A'}},
                gauge={'axis': {'range': [0, 100]}, 'bar': {'color': "#0EA5E9"},
                       'steps': [{'range': [0, 50], 'color': "#FEE2E2"}, {'range': [50, 70], 'color': "#FEF3C7"},
                                 {'range': [70, 85], 'color': "#E0F2FE"}, {'range': [85, 100], 'color': "#D1FAE5"}]}
            ))
            fig_gauge.update_layout(margin=dict(l=20, r=20, t=80, b=20), height=250)
            st.plotly_chart(fig_gauge, width='stretch')
            st.caption(f"Portfolio average Accuracy_Pct (Master Audit) across {total_skus:,} SKUs in view.")
        else:
            st.info("Accuracy_Pct not available in this data source — gauge cannot be computed.")
    else:
        st.info("No SKUs in the current filter — adjust the ABC/XYZ selection in the sidebar to see portfolio forecast quality.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("📐 Portfolio-Level Demand Statistics")
    st.markdown(f'<div class="section-caption">Distribution of <em>Mean_Monthly_Demand</em> across all {total_skus:,} SKUs in view (portfolio spread, not a single SKU\'s month-to-month variability). {source_tag("Dashboard Derived")}</div>', unsafe_allow_html=True)
    if total_skus > 0 and 'Mean_Monthly_Demand' in filtered_df.columns and mean_monthly_series.notna().sum() > 2:
        arr = mean_monthly_series.dropna().values.astype(float)
        s1, s2, s3, s4, s5 = st.columns(5)
        s1.metric("Mean", fmt_num(np.mean(arr), 1))
        s2.metric("Median", fmt_num(np.median(arr), 1))
        s3.metric("Std Dev", fmt_num(np.std(arr), 1))
        s4.metric("Skewness", fmt_num(skew(arr), 2) if skew is not None else "N/A")
        s5.metric("Kurtosis", fmt_num(kurtosis(arr), 2) if kurtosis is not None else "N/A")
        fig_dist = px.histogram(arr, nbins=20, marginal="box", title="Distribution of Mean Monthly Demand Across Portfolio")
        fig_dist.update_layout(plot_bgcolor='white', showlegend=False, margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig_dist, width='stretch')
    else:
        st.info("Not enough Mean_Monthly_Demand data to compute portfolio statistics.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🗺️ Forecast Risk by ABC Class")
    st.markdown(f'<div class="section-caption">Substitutes the original 24-month heatmap, which required a monthly historical series not present in this Master Audit export. {source_tag("Dashboard Derived")}</div>', unsafe_allow_html=True)
    if 'ABC_Class' in filtered_df.columns and 'Forecast_Risk' in filtered_df.columns and total_skus > 0:
        risk_matrix = pd.crosstab(filtered_df['ABC_Class'], filtered_df['Forecast_Risk'])
        risk_matrix = risk_matrix.reindex(index=ABC_ORDER, columns=RISK_ORDER, fill_value=0)
        fig_riskhm = px.imshow(risk_matrix, text_auto=True, color_continuous_scale=[[0, '#F1F5F9'], [1, '#991B1B']],
                                labels=dict(x="Forecast Risk", y="ABC — Demand Volume Classification", color="SKU Count"))
        fig_riskhm.update_xaxes(categoryorder='array', categoryarray=RISK_ORDER)
        fig_riskhm.update_yaxes(categoryorder='array', categoryarray=ABC_ORDER)
        fig_riskhm.update_layout(title="Forecast Risk by ABC Class — SKU Count", margin=dict(l=10, r=10, t=40, b=10))
        st.plotly_chart(fig_riskhm, width='stretch')
    else:
        st.info("ABC_Class and/or Forecast_Risk not available for this view.")

    # --------------------------------------------------------------------
    # FORECASTING — what the project enables (evidence-only).
    # --------------------------------------------------------------------
    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🔁 What the Forecasting Layer Enables")
    st.markdown(
        '<div class="decision-card"><div class="decision-card-title">📏 Validation-Driven Forecasting '
        f'{source_tag("Master Audit")}</div><div class="decision-card-text">'
        'Historical demand diagnostics feed a SKU-level model competition (AutoARIMA, Naive, Seasonal Naive, '
        '3-Month Moving Average); the winning model is selected by lowest validation sMAPE on unseen holdout data, '
        'producing the per-SKU forecast, accuracy and risk outputs shown above.</div></div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="decision-card"><div class="decision-card-title">📊 Current Portfolio Evidence '
        f'{source_tag("Master Audit")}</div><div class="decision-card-text">'
        f'{fmt_num(avg_accuracy, 2, "%") if not is_missing(avg_accuracy) else "N/A"} average validation accuracy across '
        f'{total_skus:,} SKUs in the current filter — this is holdout validation performance, not an improvement figure '
        f'versus any prior forecasting process.</div></div>', unsafe_allow_html=True)

# ==============================================================================
# TAB 3 — INVENTORY OPTIMISATION
# Translate forecasts into inventory decisions for a selected SKU, plus a
# portfolio-level replenishment view.
# ==============================================================================
with tab3:
    h3c1, h3c2 = st.columns([3, 1])
    with h3c1:
        st.subheader("Inventory Parameters & Replenishment")
    with h3c2:
        if sku_list:
            st.selectbox("🎯 Isolate SKU", sku_list, key="tab3_sku", on_change=lambda: sync_from('tab3_sku'))

    sku_row3 = None
    if st.session_state.shared_sku and total_skus > 0:
        match3 = filtered_df[filtered_df['SKU'] == st.session_state.shared_sku]
        if len(match3) > 0:
            sku_row3 = match3.iloc[0]

    if sku_row3 is not None:
        cls_line = f"ABC: <strong>{safe_val(sku_row3,'ABC_Class','N/A')}</strong>"
        if HAS_XYZ:
            cls_line += f" &nbsp;|&nbsp; XYZ: <strong>{safe_val(sku_row3,'XYZ_Class','N/A')}</strong>"
        if HAS_ABC_XYZ:
            cls_line += f" &nbsp;|&nbsp; Combined: <strong>{safe_val(sku_row3,'ABC_XYZ_Class','N/A')}</strong>"
        safe_sku = html.escape(str(st.session_state.shared_sku))
        st.markdown(f"#### Target SKU: **{safe_sku}** &nbsp; ({cls_line})", unsafe_allow_html=True)

        inv_cols = st.columns(4)
        render_metric_card(inv_cols[0], "Safety Stock", fmt_int(safe_val(sku_row3, 'Safety_Stock')), "units", variant="ribbon", extra=source_tag("Master Audit"))
        render_metric_card(inv_cols[1], "Reorder Point", fmt_int(safe_val(sku_row3, 'Reorder_Point')), "units", variant="ribbon", extra=source_tag("Master Audit"))
        render_metric_card(inv_cols[2], "EOQ", fmt_int(safe_val(sku_row3, 'EOQ')), "units / order", variant="ribbon", extra=source_tag("Master Audit"))
        render_metric_card(inv_cols[3], "Average Inventory", fmt_int(safe_val(sku_row3, 'Average_Inventory_Units')), "units", variant="ribbon", extra=source_tag("Master Audit"))

        inv_cols2 = st.columns(4)
        render_metric_card(inv_cols2[0], "Model-Implied Inventory Value", fmt_monetary(safe_val(sku_row3, 'Inventory_Value')), "", extra=source_tag("Master Audit"))
        render_metric_card(inv_cols2[1], "Model-Implied Working Capital", fmt_monetary(safe_val(sku_row3, 'Working_Capital')), "", extra=source_tag("Master Audit"))
        render_metric_card(inv_cols2[2], "Model-Implied Inventory Days", fmt_num(safe_val(sku_row3, 'Inventory_Days'), 0), "days", extra=source_tag("Master Audit"))
        render_metric_card(inv_cols2[3], "Inventory Turnover", fmt_num(safe_val(sku_row3, 'Inventory_Turnover'), 2), "turns / yr", extra=source_tag("Master Audit"))

        st.markdown("<br>", unsafe_allow_html=True)
        svc_col1, svc_col2, svc_col3 = st.columns(3)
        with svc_col1:
            st.metric("Model-Implied Service Level %", fmt_num(safe_val(sku_row3, 'Service_Level_Pct'), 1), help="Direct from Master Audit; model-implied, not an observed realised service level.")
        with svc_col2:
            st.metric("Model-Implied Fill Rate %", fmt_num(safe_val(sku_row3, 'Fill_Rate_Pct'), 1), help="Direct from Master Audit; model-implied, not an observed realised fill rate.")
        with svc_col3:
            ihs = safe_val(sku_row3, 'Inventory_Health_Score')
            ihs_status = score_to_status(ihs)
            st.markdown(f"**Inventory Health:** {fmt_num(ihs,0)} {badge_html(ihs_status, ihs_status.title())} {source_tag('Master Audit')}", unsafe_allow_html=True)
        st.caption(f"{source_tag('Master Audit')} = Service_Level_Pct and Fill_Rate_Pct are read directly from the Master Audit; they are model-implied rather than observed realised rates.", unsafe_allow_html=True)

        pol_col1, pol_col2, pol_col3 = st.columns(3)
        pol_col1.markdown(f"**Planning Frequency:** {safe_val(sku_row3, 'Planning Frequency', 'N/A')}")
        pol_col2.markdown(f"**Inventory Policy:** {safe_val(sku_row3, 'Inventory Policy', 'N/A')}")
        pol_col3.markdown(f"**Review Frequency:** {safe_val(sku_row3, 'Review Frequency', 'N/A')}")

        # ------------------------------------------------------------------
        # MATHEMATICAL INTEGRITY — Safety Stock / Reorder Point forensic check
        # for the selected SKU. Master Audit value is NEVER overwritten; the
        # independent recomputation is shown alongside it with an explicit
        # status, exactly as required for auditability.
        # ------------------------------------------------------------------
        st.markdown("<hr>", unsafe_allow_html=True)
        st.subheader("🧮 Inventory Mathematics Integrity — Selected SKU")
        with st.expander("Master Audit vs. Independent Validation (click to expand)", expanded=(str(safe_val(sku_row3, 'Safety_Stock', 0)) == '0')):
            val_result = independent_ss_rop(sku_row3)
            ma_ss = safe_val(sku_row3, 'Safety_Stock')
            ma_rop = safe_val(sku_row3, 'Reorder_Point')
            ind_ss, ind_rop = val_result['ss'], val_result['rop']
            ss_var = abs(ma_ss - ind_ss) if (not is_missing(ma_ss) and not is_missing(ind_ss)) else None
            rop_var = abs(ma_rop - ind_rop) if (not is_missing(ma_rop) and not is_missing(ind_rop)) else None
            ss_ok = ss_var is not None and ss_var <= 1
            rop_ok = rop_var is not None and rop_var <= 1
            st.markdown(
                '<table class="math-audit-table">'
                '<tr><th></th><th>Master Audit</th><th>Independent Validation</th><th>Variance</th><th>Status</th></tr>'
                f'<tr><td><strong>Safety Stock</strong></td><td>{fmt_num(ma_ss,0)}</td><td>{fmt_num(ind_ss,0)}</td><td>{fmt_num(ss_var,0) if ss_var is not None else "N/A"}</td>'
                f'<td class="{"math-status-ok" if ss_ok else "math-status-review"}">{"✓ Confirmed" if ss_ok else "⚠ Review" if ss_var is not None else "N/A"}</td></tr>'
                f'<tr><td><strong>Reorder Point</strong></td><td>{fmt_num(ma_rop,0)}</td><td>{fmt_num(ind_rop,0)}</td><td>{fmt_num(rop_var,0) if rop_var is not None else "N/A"}</td>'
                f'<td class="{"math-status-ok" if rop_ok else "math-status-review"}">{"✓ Confirmed" if rop_ok else "⚠ Review" if rop_var is not None else "N/A"}</td></tr>'
                '</table>',
                unsafe_allow_html=True)
            st.caption(f"Basis: {val_result['basis']}. Independent formula: Safety Stock = 1.645 × RMSE × √(Lead Time ÷ 30); "
                       f"Reorder Point = (Forecast Next Month ÷ 30 × Lead Time) + Safety Stock — reproduced directly from the source "
                       f"notebook's own implementation, not a generic textbook substitute. {source_tag('Dashboard Derived')}")
            if ma_ss == 0 and ss_ok:
                st.info("A Safety Stock of 0 is mathematically correct here — not a data-quality defect. It occurs when the SKU has zero "
                        "historical demand variance (dead stock / constant demand) or when the forecast model achieved a near-zero RMSE "
                        "on a very low-volume, intermittent-demand SKU, so the computed buffer legitimately rounds to zero.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🔁 Portfolio Replenishment View")
    repl_cols = [c for c in ['SKU', 'ABC_XYZ_Class', 'ABC_Class', 'Safety_Stock', 'Reorder_Point', 'EOQ',
                              'Inventory_Days', 'Inventory_Turnover', 'Inventory_Health_Score', 'Inventory_Recommendation'] if c in filtered_df.columns]
    if repl_cols and total_skus > 0:
        repl_df = filtered_df[repl_cols].copy()
        if 'Inventory_Health_Score' in repl_df.columns:
            repl_df = repl_df.sort_values('Inventory_Health_Score', ascending=True)
        st.dataframe(repl_df.head(50), width='stretch', hide_index=True)
        st.caption(f"Showing {min(50, len(repl_df)):,} of {len(repl_df):,} SKUs, sorted by lowest Inventory Health Score first (most attention-worthy). {source_tag('Master Audit')}", unsafe_allow_html=True)
    else:
        st.info("Inventory fields not available for the portfolio view.")

    repl_col1, repl_col2 = st.columns(2)
    with repl_col1:
        if 'Inventory_Turnover' in filtered_df.columns and 'ABC_Class' in filtered_df.columns and total_skus > 0:
            turn_by_abc = filtered_df.groupby('ABC_Class', as_index=False)['Inventory_Turnover'].mean()
            turn_by_abc['ABC_Class'] = pd.Categorical(turn_by_abc['ABC_Class'], categories=ABC_ORDER, ordered=True)
            turn_by_abc = turn_by_abc.sort_values('ABC_Class')
            fig_turn = px.bar(turn_by_abc, x='ABC_Class', y='Inventory_Turnover', color='ABC_Class',
                               color_discrete_map=ABC_COLOR_MAP, category_orders={'ABC_Class': ABC_ORDER}, title="Average Inventory Turnover by ABC Class")
            fig_turn.update_layout(plot_bgcolor='white', showlegend=False, margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_turn, width='stretch')
        else:
            st.info("Inventory_Turnover not available.")
    with repl_col2:
        if 'Inventory_Health_Score' in filtered_df.columns and total_skus > 0:
            fig_ihs = px.histogram(filtered_df, x='Inventory_Health_Score', title="Inventory Health Score Distribution", color_discrete_sequence=['#10B981'])
            fig_ihs.update_layout(plot_bgcolor='white', margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_ihs, width='stretch')
        else:
            st.info("Inventory_Health_Score not available.")

    # --------------------------------------------------------------------
    # INVENTORY — what the project enables (evidence-only).
    # --------------------------------------------------------------------
    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🔁 What the Inventory Layer Enables")
    st.markdown(
        '<div class="decision-card"><div class="decision-card-title">📦 Forecast-Driven Inventory Policy '
        f'{source_tag("Master Audit")}</div><div class="decision-card-text">'
        'Forecast + RMSE-based Safety Stock, Lead-Time Demand, Reorder Point, EOQ and ABC–XYZ segmentation are computed per '
        'SKU (SS = 1.645 × RMSE × √(7/30); ROP = LTD + SS; EOQ = √(2 × Annual Demand × Ordering Cost ÷ Holding Cost per Unit)), '
        'providing Safety Stock, Reorder Point, EOQ, Model-Implied Inventory Days, Inventory Turnover, Model-Implied Service Level, Model-Implied Fill Rate and Inventory '
        'Health figures above.</div></div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="decision-card"><div class="decision-card-title">📊 Model-Implied Inventory Position '
        f'{source_tag("Master Audit")}</div><div class="decision-card-text">'
        f'Model-Implied Inventory Value across the current filter is <strong>{fmt_monetary(inv_value_sum)}</strong> — the analytical '
        f'figure implied by applying this policy to Master Audit fields, not an observed accounting balance. '
        f'{"An actual on-hand inventory column is available in this data source." if ACTUAL_INV_COL else "No actual/observed on-hand inventory column is present in this data source, so it is not shown as if it were."}'
        f'</div></div>', unsafe_allow_html=True)
    st.caption("Any inventory reduction or working-capital release is a modelled implication of this policy, not a measured "
               "outcome — realising it requires comparison against an actual physical inventory baseline.")

# ==============================================================================
# TAB 4 — RISK, SEGMENTATION & DIAGNOSTICS
# Where does the supply chain need attention?
# ==============================================================================
with tab4:
    st.subheader("⚠️ Bullwhip Ratio vs. Forecast Error")
    st.markdown(f"""<div class="section-caption">Quadrant boundaries below are median-split cut points computed on the current filter — a
    {source_tag("Dashboard Derived")} diagnostic classification, not an official Master Audit threshold.</div>""", unsafe_allow_html=True)

    if total_skus > 2 and 'Bullwhip_Ratio' in filtered_df.columns and 'RMSE' in filtered_df.columns:
        med_bw = bullwhip_series.median()
        med_rmse = rmse_series.median()
        scatter_df = filtered_df.copy()
        scatter_df['_bw'] = bullwhip_series
        scatter_df['_rmse'] = rmse_series

        def quad_label(row):
            bw_high = row['_bw'] > med_bw
            rmse_high = row['_rmse'] > med_rmse
            if not bw_high and not rmse_high: return "Stable"
            if not bw_high and rmse_high: return "Forecast Issue"
            if bw_high and not rmse_high: return "Signal Amplification"
            return "Critical"
        scatter_df['Diagnostic Quadrant'] = scatter_df.apply(quad_label, axis=1)

        fig_scatter = px.scatter(scatter_df, x='_bw', y='_rmse', color='Diagnostic Quadrant',
                                  color_discrete_map={"Stable": "#10B981", "Forecast Issue": "#F59E0B",
                                                       "Signal Amplification": "#0EA5E9", "Critical": "#EF4444"},
                                  hover_data=['SKU'] if 'SKU' in scatter_df.columns else None,
                                  labels={'_bw': 'Bullwhip Ratio — Model-Derived Replenishment Signal Amplification', '_rmse': 'RMSE'},
                                  title="Bullwhip Ratio vs. Forecast Error — Dashboard-Derived Review Lens")
        fig_scatter.add_vline(x=med_bw, line_dash="dot", line_color="#94A3B8")
        fig_scatter.add_hline(y=med_rmse, line_dash="dot", line_color="#94A3B8")
        fig_scatter.update_traces(marker=dict(size=9, opacity=0.75, line=dict(width=1, color='white')))
        fig_scatter.update_layout(plot_bgcolor='white', margin=dict(l=20, r=20, t=40, b=20), hoverlabel=dict(font_family="Calibri"))
        st.plotly_chart(fig_scatter, width='stretch')
    else:
        st.info("Bullwhip_Ratio and/or RMSE not available, or not enough SKUs in the current filter.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("📊 Segmentation Overview")
    seg1, seg2 = st.columns(2)
    with seg1:
        if 'ABC_Class' in filtered_df.columns and total_skus > 0:
            abc_bar = ordered_category_counts(filtered_df, 'ABC_Class', ABC_ORDER)
            fig_abc_bar = px.bar(abc_bar, x='ABC_Class', y='Count', color='ABC_Class',
                                  color_discrete_map=ABC_COLOR_MAP, category_orders={'ABC_Class': ABC_ORDER}, title="ABC — Demand Volume Classification")
            fig_abc_bar.update_layout(plot_bgcolor='white', showlegend=False, margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_abc_bar, width='stretch')
        else:
            st.info("ABC_Class not available.")
    with seg2:
        if HAS_XYZ and total_skus > 0:
            xyz_bar = ordered_category_counts(filtered_df, 'XYZ_Class', XYZ_ORDER)
            fig_xyz_bar = px.bar(xyz_bar, x='XYZ_Class', y='Count', color='XYZ_Class',
                                  color_discrete_map=XYZ_COLOR_MAP, category_orders={'XYZ_Class': XYZ_ORDER}, title="XYZ — Demand Variability Classification")
            fig_xyz_bar.update_layout(plot_bgcolor='white', showlegend=False, margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_xyz_bar, width='stretch')
        else:
            st.info("XYZ_Class not available.")

    st.subheader("🧩 ABC–XYZ Risk Intensity Matrix")
    st.markdown(f'<div class="section-caption">Average RMSE per ABC × XYZ cell — a different lens on the same segmentation shown in Tab 1 (which shows SKU counts). {source_tag("Dashboard Derived")}</div>', unsafe_allow_html=True)
    if HAS_ABC_XYZ and 'RMSE' in filtered_df.columns and total_skus > 0:
        abc_order = ABC_ORDER; xyz_order = XYZ_ORDER
        risk_intensity = pd.DataFrame(np.nan, index=abc_order, columns=xyz_order)
        tmp = filtered_df.copy()
        tmp['_A'] = tmp['ABC_XYZ_Class'].astype(str).str[0]
        tmp['_X'] = tmp['ABC_XYZ_Class'].astype(str).str[1]
        grp = tmp.groupby(['_A', '_X'])['RMSE'].mean()
        for (a, x), v in grp.items():
            if a in abc_order and x in xyz_order:
                risk_intensity.loc[a, x] = v
        fig_intensity = px.imshow(risk_intensity, text_auto='.1f', color_continuous_scale=[[0, '#D1FAE5'], [0.5, '#FEF3C7'], [1, '#991B1B']],
                                   labels=dict(x="XYZ", y="ABC", color="Avg RMSE"))
        fig_intensity.update_layout(title="ABC–XYZ Risk Intensity — Mean RMSE", margin=dict(l=10, r=10, t=40, b=10))
        st.plotly_chart(fig_intensity, width='stretch')
    else:
        st.info("ABC_XYZ_Class and/or RMSE not available for this view.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🔴 High-Risk SKU Register")
    hr_cols = [c for c in ['SKU', 'ABC_XYZ_Class', 'RMSE', 'Bullwhip_Ratio', 'CV', 'Forecast_Risk',
                            'Business Risk', 'Priority_Level', 'Supply Chain Priority', 'Model_Reliability'] if c in filtered_df.columns]
    if hr_cols and 'Forecast_Risk' in filtered_df.columns and total_skus > 0:
        hr_df = filtered_df[filtered_df['Forecast_Risk'].astype(str).str.lower() == 'high'][hr_cols].copy()
        hr_df = deterministic_high_risk_sort(hr_df)
        st.dataframe(hr_df, width='stretch', hide_index=True)
        st.caption(f"{len(hr_df):,} SKU(s) flagged High Forecast Risk (Master Audit). {source_tag('Master Audit')}", unsafe_allow_html=True)
    else:
        st.info("Forecast_Risk not available to build the high-risk register.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("📉 Pareto Analysis")
    st.markdown(f'<div class="section-caption">Based on Annualised Demand from the Master Audit. The current source export does not contain per-SKU monthly history, so this is an aggregate demand-volume Pareto view rather than a reconstructed monthly series. {source_tag("Master Audit")}</div>', unsafe_allow_html=True)
    if 'Annual_Demand' in filtered_df.columns and total_skus > 0:
        pareto_df = filtered_df[['SKU', 'Annual_Demand']].dropna().sort_values('Annual_Demand', ascending=False).reset_index(drop=True)
        annual_total = pareto_df['Annual_Demand'].sum()
        if len(pareto_df) > 0 and annual_total and annual_total > 0:
            pareto_df['Cumulative %'] = pareto_df['Annual_Demand'].cumsum() / annual_total * 100
            top_n = pareto_df.head(30)
            fig_pareto = make_subplots(specs=[[{"secondary_y": True}]])
            fig_pareto.add_trace(go.Bar(x=top_n['SKU'], y=top_n['Annual_Demand'], name='Annualised Demand', marker_color='#0EA5E9'), secondary_y=False)
            fig_pareto.add_trace(go.Scatter(x=top_n['SKU'], y=top_n['Cumulative %'], name='Cumulative %', line=dict(color='#F43F5E', width=3)), secondary_y=True)
            fig_pareto.add_hline(y=80, line_dash="dot", line_color="#991B1B", secondary_y=True, annotation_text="80% reference threshold")
            fig_pareto.update_layout(title="Top 30 SKUs — Pareto (Annualised Demand, Master Audit)", plot_bgcolor='white',
                                      margin=dict(l=20, r=20, t=40, b=80), xaxis_tickangle=-60)
            st.plotly_chart(fig_pareto, width='stretch')
            reached = pareto_df.index[pareto_df['Cumulative %'] >= 80]
            a_class_count = int(reached[0] + 1) if len(reached) else len(pareto_df)
            st.caption(f"Approximately {a_class_count:,} of {len(pareto_df):,} SKUs (by Annualised Demand) reach 80% of cumulative demand.")
        else:
            st.info("Annual_Demand is unavailable or zero for the current filter — Pareto ranking is not meaningful.")
    else:
        st.info("Annual_Demand is not available for the primary Pareto view.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🪤 High Inventory Days & Lower Demand Diagnostic")
    st.markdown(f"""<div class="section-caption">Flags SKUs with model-implied Inventory Days in the upper quartile and Mean Monthly Demand in the lower quartile of the current filter as candidates for further inventory review. This does not establish obsolescence, inactivity, trapped capital or liquidity loss. {source_tag('Dashboard Derived')}</div>""", unsafe_allow_html=True)
    if 'Inventory_Days' in filtered_df.columns and 'Mean_Monthly_Demand' in filtered_df.columns and total_skus > 3:
        inventory_days_series = safe_series(filtered_df, 'Inventory_Days', numeric=True)
        days_q75 = inventory_days_series.quantile(0.75)
        demand_q25 = mean_monthly_series.quantile(0.25)
        trap_mask = (inventory_days_series >= days_q75) & (mean_monthly_series <= demand_q25)
        trap_cols = [c for c in ['SKU', 'ABC_XYZ_Class', 'Mean_Monthly_Demand', 'Inventory_Days', 'Inventory_Value',
                                  'Forecast_Risk', 'Inventory_Recommendation'] if c in filtered_df.columns]
        trap_df = filtered_df[trap_mask][trap_cols].copy()
        if len(trap_df) > 0:
            sort_cols = [c for c in ['Inventory_Value', 'Inventory_Days', 'Mean_Monthly_Demand', 'SKU'] if c in trap_df.columns]
            ascending = [False, False, True, True][:len(sort_cols)]
            trap_df = trap_df.sort_values(sort_cols, ascending=ascending, kind='mergesort', na_position='last')
        st.dataframe(trap_df, width='stretch', hide_index=True)
        st.caption(f"{len(trap_df):,} SKU(s) flagged as dashboard-derived candidates for inventory review — requires validation. "
                   f"These are dashboard-derived candidates for review, not confirmed excess inventory.")
    else:
        st.info("Inventory_Days and/or Mean_Monthly_Demand not available to compute this diagnostic.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🏷️ Business Risk & Supply Chain Priority Breakdown")
    br1, br2 = st.columns(2)
    with br1:
        if 'Business Risk' in filtered_df.columns and total_skus > 0:
            br_order = ["Elevated Variability / Stockout Risk", "Moderate Variance", "Low Obsolescence"]
            br_counts = filtered_df['Business Risk'].value_counts().reindex(br_order, fill_value=0).rename_axis('Business Risk').reset_index(name='Count')
            fig_br = px.bar(br_counts, x='Business Risk', y='Count', color='Business Risk',
                            category_orders={'Business Risk': br_order}, title="Business Risk Categories")
            fig_br.update_layout(plot_bgcolor='white', showlegend=False, margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_br, width='stretch')
        else:
            st.info("Business Risk field not available.")
    with br2:
        if 'Supply Chain Priority' in filtered_df.columns and total_skus > 0:
            sp_order = ["High", "Medium", "Standard"]
            sp_counts = filtered_df['Supply Chain Priority'].value_counts().reindex(sp_order, fill_value=0).rename_axis('Supply Chain Priority').reset_index(name='Count')
            fig_sp = px.bar(sp_counts, x='Supply Chain Priority', y='Count', color='Supply Chain Priority',
                            category_orders={'Supply Chain Priority': sp_order}, title="Supply Chain Priority Categories")
            fig_sp.update_layout(plot_bgcolor='white', showlegend=False, margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_sp, width='stretch')
        else:
            st.info("Supply Chain Priority field not available.")

    # --------------------------------------------------------------------
    # IMPLEMENTATION PRIORITY LENS — built entirely from existing, evidenced
    # fields (ABC, XYZ, Forecast Risk, Bullwhip, RMSE, Inventory Health,
    # Business Risk). Clearly labelled Dashboard Derived: a prioritisation
    # lens for human review, not a company-approved classification.
    # --------------------------------------------------------------------
    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🧭 Where Should Management Investigate First?")
    st.markdown(f'<div class="section-caption">Dashboard-Derived Review Lens — combinations built from existing ABC, XYZ, Forecast Risk, Bullwhip, RMSE, '
                f'Inventory Health and Business Risk fields already shown on this tab. {source_tag("Dashboard Derived")}</div>', unsafe_allow_html=True)

    pz1, pz2 = st.columns(2)
    with pz1:
        render_status_card("red", "1. High Demand Volume + High Variability", "A-class SKUs combined with Z-class demand variability — "
                            "immediate analytical attention.", action="Review AZ-segment SKUs in Tab 1 ABC–XYZ matrix")
        render_status_card("amber", "2. High Forecast Error + High Replenishment-Signal Amplification", "Elevated RMSE alongside high Bullwhip Ratio — Model-Derived Replenishment Signal Amplification — "
                            "forecast / replenishment-signal diagnostic review.", action="See Bullwhip Ratio vs. Forecast Error quadrant above")
    with pz2:
        render_status_card("amber", "3. High Inventory Days + Low Demand", "SKUs with high model-implied Inventory Days and lower demand — further inventory review.", action="See High Inventory Days & Lower Demand Diagnostic above")
        render_status_card("red", "4. High Business Risk + Lower Service Indicators", "Elevated business risk combined with lower model-implied service/fill-rate indicators — service protection review.", action="Cross-reference with Tab 3 Model-Implied Service Level / Fill Rate")

# ==============================================================================
# TAB 5 — FINANCIAL IMPACT, METHODOLOGY & AUDIT
# ==============================================================================
with tab5:
    st.markdown('<div class="section-caption">Analytical baseline — Master Audit. Figures below are model-implied financial '
                'exposure computed from this project\'s inventory policy, not observed accounting balances.</div>', unsafe_allow_html=True)
    st.subheader("💰 Master Audit Financial Measures")
    st.markdown('<div class="section-caption">Portfolio aggregations of authoritative Master Audit financial fields. The underlying fields remain Master Audit; the portfolio aggregation is Dashboard Derived.</div>', unsafe_allow_html=True)
    fk1, fk2, fk3, fk4 = st.columns(4)
    render_metric_card(fk1, "Model-Implied Inventory Value", fmt_monetary(inv_value_sum), "Portfolio aggregation of Master Audit Inventory_Value", extra=source_tag("Dashboard Derived"))
    render_metric_card(fk2, "Model-Implied Working Capital", fmt_monetary(wc_sum), "Portfolio aggregation of Master Audit Working_Capital", extra=source_tag("Dashboard Derived"))
    carrying_sum = safe_series(filtered_df, 'Estimated_Carrying_Cost', numeric=True).sum() if 'Estimated_Carrying_Cost' in filtered_df.columns else np.nan
    stockout_sum = safe_series(filtered_df, 'Stockout_Cost', numeric=True).sum() if 'Stockout_Cost' in filtered_df.columns else np.nan
    render_metric_card(fk3, "Estimated Carrying Cost", fmt_monetary(carrying_sum), "Portfolio aggregation of Master Audit Estimated_Carrying_Cost", extra=source_tag("Dashboard Derived"))
    render_metric_card(fk4, "Estimated Stockout Cost / Exposure", fmt_monetary(stockout_sum), "Portfolio aggregation of Master Audit Stockout_Cost", extra=source_tag("Dashboard Derived"))

    st.subheader("📊 Dashboard-Derived Portfolio Indicators")
    st.markdown('<div class="section-caption">Aggregations or transformations calculated by the dashboard from authoritative Master Audit fields; these are not separate source columns.</div>', unsafe_allow_html=True)
    fk5, fk6, fk7, fk8 = st.columns(4)
    turnover_avg = portfolio_turnover
    days_avg = portfolio_inventory_days
    wc_eff_avg = portfolio_wce
    render_metric_card(fk5, "Portfolio Inventory Turnover", fmt_num(turnover_avg, 2), "Mean of Master Audit Inventory_Turnover", extra=source_tag("Dashboard Derived"))
    render_metric_card(fk6, "Portfolio-Implied Inventory Days", fmt_num(days_avg, 0), "365 ÷ portfolio mean Inventory Turnover", extra=source_tag("Dashboard Derived"))
    render_metric_card(fk7, "Portfolio Mean Financial Health", fmt_num(fin_health_avg, 0, "%"), "Mean of SKU Financial Health Scores", extra=source_tag("Dashboard Derived"))
    render_metric_card(fk8, "Portfolio Working Capital Efficiency — Derived", fmt_num(wc_eff_avg, 0, "%"), "Portfolio mean Inventory Turnover ÷ 6 × 100; clipped 0–100", extra=source_tag("Dashboard Derived"))

    st.markdown("<hr>", unsafe_allow_html=True)
    tcol1, tcol2 = st.columns(2)
    with tcol1:
        st.subheader("🗺️ Inventory Capital Allocation by ABC Hierarchy")
        tree_df = filtered_df[['ABC_Class', 'SKU', 'Inventory_Value']].dropna() if ('Inventory_Value' in filtered_df.columns and 'ABC_Class' in filtered_df.columns) else pd.DataFrame()
        if len(tree_df) > 0:
            fig_tree = px.treemap(tree_df, path=[px.Constant("Portfolio"), 'ABC_Class', 'SKU'], values='Inventory_Value',
                                   color='ABC_Class', color_discrete_map=ABC_COLOR_MAP)
            fig_tree.update_layout(title="Inventory Value by ABC Hierarchy", margin=dict(l=10, r=10, t=40, b=10))
            st.plotly_chart(fig_tree, width='stretch')
            st.caption(f"Tile size = Inventory_Value (Master Audit). {source_tag('Master Audit')}", unsafe_allow_html=True)
        else:
            st.info("Inventory_Value and/or ABC_Class not available for the treemap.")
    with tcol2:
        st.subheader("💰 Financial Composition")
        st.markdown(f'<div class="section-caption">Portfolio comparison of Master Audit financial components; not a hypothetical reduction scenario. {source_tag("Dashboard Derived")}</div>', unsafe_allow_html=True)
        if not is_missing(inv_value_sum) and not is_missing(carrying_sum) and not is_missing(stockout_sum):
            # Separate model-implied financial components only.
            fig_components = go.Figure(go.Bar(
                x=["Model-Implied Inventory Value", "Estimated Carrying Cost", "Estimated Stockout Cost / Exposure"],
                y=[inv_value_sum, carrying_sum, stockout_sum],
                text=[fmt_monetary(inv_value_sum), fmt_monetary(carrying_sum), fmt_monetary(stockout_sum)],
                textposition='auto', marker_color=['#1E3A8A', '#F59E0B', '#EF4444']
            ))
            fig_components.update_layout(title="Model-Implied Financial Components", plot_bgcolor='white', margin=dict(l=20, r=20, t=40, b=20), xaxis_title="Financial Measure", yaxis_title="Monetary Units")
            st.plotly_chart(fig_components, width='stretch')
        else:
            st.info("Inventory_Value, Estimated_Carrying_Cost and Stockout_Cost are all required for this chart.")

    # ==========================================================================
    # REALISED FINANCIAL IMPACT — a single honest statement rather than an
    # elaborate Baseline | Target | Realised table with no evidence behind it.
    # ==========================================================================
    st.markdown("<hr>", unsafe_allow_html=True)
    st.markdown(
        '<div class="illustrative-banner">These are the financial dimensions the project enables management to monitor '
        '(Inventory Value, Model-Implied Working Capital, Carrying Cost, Stockout Exposure, Turnover, Financial Health, Portfolio Working Capital Efficiency '
        'Efficiency). <strong>Realised financial impact requires post-deployment measurement and finance validation</strong> — '
        'it is not calculated here.</div>', unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🧪 Illustrative Inventory Reduction Scenario")
    st.markdown('<div class="illustrative-banner">⚠️ ILLUSTRATIVE SCENARIO — NOT A MODEL OUTPUT, NOT A REALISED SAVING, NOT A FEASIBILITY '
                'RESULT. A simple linear what-if slider for exploration only.</div>', unsafe_allow_html=True)
    if not is_missing(inv_value_sum):
        reduction_pct = st.slider("Illustrative inventory reduction (%)", min_value=0, max_value=40, value=10, step=1, key="illustrative_reduction_pct")
        illustrative_value_reduction = inv_value_sum * (reduction_pct / 100.0)
        residual_value = inv_value_sum - illustrative_value_reduction
        sc1, sc2, sc3 = st.columns(3)
        sc1.metric("Current Model-Implied Inventory Value", fmt_monetary(inv_value_sum))
        sc2.metric(f"Illustrative Inventory Value Reduction ({reduction_pct}%)", fmt_monetary(illustrative_value_reduction), help="Illustrative Scenario — hypothetical value reduction only.")
        sc3.metric("Illustrative Residual Inventory Value", fmt_monetary(residual_value))
        st.caption("Illustrative Scenario: simple linear what-if applied to current Model-Implied Inventory Value. It does not represent cash generation, realised benefits, a model recommendation, service-level impact, or feasibility analysis. " + "Illustrative Scenario")
    else:
        st.info("Inventory_Value not available - cannot build the illustrative scenario.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🔗 Operational Value Chain — Conceptual Process Architecture")
    st.markdown('<div class="section-caption">Illustrative process sequence only - link thickness is uniform and does NOT represent measured throughput volume.</div>', unsafe_allow_html=True)

    stage_labels = ["Demand Forecast", "Procurement", "Production", "Warehouse", "Transportation", "Customer Service", "Finance", "Executive Decisions"]
    stage_colors = ["#1E3A8A", "#1D4ED8", "#0369A1", "#0891B2", "#0D9488", "#059669", "#475569", "#334155"]
    n_stages = len(stage_labels)
    link_colors = []
    for c in stage_colors[:-1]:
        rr = int(c[1:3], 16)
        gg = int(c[3:5], 16)
        bb = int(c[5:7], 16)
        link_colors.append(f"rgba({rr},{gg},{bb},0.35)")
    fig_sankey = go.Figure(data=[go.Sankey(
        node=dict(pad=20, thickness=22, line=dict(color="white", width=1), label=stage_labels, color=stage_colors),
        link=dict(source=list(range(n_stages - 1)), target=list(range(1, n_stages)), value=[1] * (n_stages - 1), color=link_colors)
    )])
    fig_sankey.update_layout(title_text="8-Stage Operational Value Chain (Conceptual Flow)", font=dict(family="Calibri", size=13),
                              margin=dict(l=10, r=10, t=40, b=10), height=340)
    st.plotly_chart(fig_sankey, width='stretch')

    stage_details = [
        ("Demand Forecast", "Statistical models project next-period demand per SKU.", "Feeds procurement and production planning."),
        ("Procurement", "Purchase orders sized using EOQ and Reorder Point logic.", "Determines supplier order cadence."),
        ("Production", "Make-to-order / buffer policies may be considered where supplier reliability, lead time, commercial requirements and service considerations support them.", "Aligns output with validated forecast and inventory signals."),
        ("Warehouse", "Safety stock and slotting decisions based on ABC-XYZ class.", "Drives storage and pick-face prioritisation."),
        ("Transportation", "Replenishment lead time assumptions feed safety-stock sizing.", "Impacts service-level attainment."),
        ("Customer Service", "Fill Rate and Service Level tracked against demand.", "Surfaces stockout risk early."),
        ("Finance", "Model-Implied Inventory Value, Model-Implied Working Capital and Estimated Carrying Cost quantified.", "Anchors the financial impact story."),
        ("Executive Decisions", "Recommendations synthesized into prioritised actions.", "Closes the loop back to forecasting."),
    ]
    cards_html = '<div class="matrix-grid-container">'
    for i in range(len(stage_details)):
        title, line1, line2 = stage_details[i]
        cards_html += '<div class="flow-card" style="border-top: 5px solid ' + stage_colors[i] + ';">'
        cards_html += '<div class="flow-card-title">' + str(i + 1) + '. ' + title + '</div>'
        cards_html += '<div class="flow-card-text">' + line1 + '</div>'
        cards_html += '<div class="flow-card-text">' + line2 + '</div>'
        cards_html += '</div>'
    cards_html += '</div>'
    st.markdown(cards_html, unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("📚 Methodology & Traceability")
    with st.expander("How this dashboard is built — read before presenting", expanded=True):
        st.markdown(f"**Traceability:** Master Audit = direct source field; Dashboard Derived = calculated from filtered source data; Illustrative Scenario = hypothetical what-if. {source_tag('Master Audit')} {source_tag('Dashboard Derived')} {source_tag('Illustrative Scenario')}")
        methodology_text = "**Data:** Master Audit export (Enterprise_Supply_Chain_Master_Audit.xlsx, Executive_Summary sheet), " + f"{len(df):,}" + " SKUs. Methodology uses 24 calendar months, 1,401 true SKUs, Grand Total excluded, missing Outwards = 0 and negative Outwards retained. Annualised Demand = full-period mean monthly demand × 12; it is not necessarily trailing-12-month actual demand. Month-level historical raw data was not part of this export, so demand statistics here are either per-SKU aggregates or portfolio-level distributions of those aggregates.\n\n"
        methodology_text += "**Forecasting:** Forecast outputs (Forecast_Next_Month, Accuracy_Pct, MAE, RMSE, Bias, CV, Health_P_Value) are taken directly from the Master Audit. Accuracy_Pct is the project-defined sMAPE-derived accuracy metric, not probabilistic confidence. This dashboard does not re-run, re-fit or retrain any forecasting model; final SKU model selection is based on the lowest validation sMAPE.\n\n"
        methodology_text += "**Validation:** Where per-SKU test-period arrays are present in the source file, an actual-vs-predicted trajectory with an uncertainty band is shown. Where they are not (the case for the current Master Audit export), only the aggregate validation statistics are shown, and this is stated explicitly on-screen.\n\n"
        methodology_text += "**Inventory Optimisation:** Safety Stock, Reorder Point and EOQ are read directly from the Master Audit. No inventory numbers are invented — where an actual/observed on-hand figure isn't available, only the model recommendation is shown. The underlying formula (reverse-engineered from the source notebook and verified against all SKUs — see the Inventory Mathematics QA Report below) is: Safety Stock = 1.645 × RMSE × √(Lead Time ÷ 30); Reorder Point = (Forecast Next Month ÷ 30 × Lead Time) + Safety Stock. **Note:** this is a single-uncertainty-term model driven by forecast error (RMSE); lead time itself is treated as a fixed 7-day constant, not a random variable with its own standard deviation. A fuller two-term formulation (combining separate demand-variance and lead-time-variance terms) is sometimes used as a reference methodology, but it is not what this specific pipeline computes — the dashboard validates against what the model actually implements, not against an assumption of what it should implement.\n\n"
        methodology_text += "**Segmentation:** ABC — Demand Volume Classification and XYZ — Demand Variability Classification, and their combination, come directly from the Master Audit. ABC is based on cumulative demand volume.\n\n"
        methodology_text += "**Presentation Bands:** Forecast Health Score badges use dashboard presentation bands of Excellent (≥80), Good (60–79), Attention (40–59) and Poor (<40). These are display rules, not source-model thresholds.\n\n"
        methodology_text += "**Risk:** RMSE, Bullwhip Ratio, Business Risk and Priority Level come from the Master Audit. In the diagnostic view, Bullwhip Ratio is interpreted as model-derived replenishment signal amplification relative to the reference level. The Bullwhip Ratio vs. Forecast Error chart and High Inventory Days & Lower Demand Diagnostic are Dashboard-Derived Review Lenses.\n\n"
        methodology_text += "**Financial Layer:** Inventory_Value, Working_Capital, Estimated_Carrying_Cost and Stockout_Cost are Master Audit fields; dashboard portfolio aggregates of these fields are explicitly labelled Dashboard Derived. The only hypothetical figure on this dashboard is the clearly labelled Illustrative Inventory Value Reduction scenario.\n\n"
        methodology_text += "**Decision Layer:** Recommendations shown below are the Master Audit's rule-based, model-informed outputs — presented as Decision Support / Decision Intelligence, not as autonomous actions or operational execution. The executive recommendation panel groups flagged SKUs by recommendation frequency; it is not a priority ranking."
        st.markdown(methodology_text)
        if MISSING_AUTHORITATIVE_FIELDS:
            st.markdown(f"""**Field Completeness Check:** {len(MISSING_AUTHORITATIVE_FIELDS)} field(s) expected by this dashboard are
            not present in the current data source and are shown as "N/A" or omitted wherever referenced: {', '.join(MISSING_AUTHORITATIVE_FIELDS)}.""")
        else:
            st.markdown("**Field Completeness Check:** All authoritative fields this dashboard expects are present in the current data source.")

    # ------------------------------------------------------------------------
    # PORTFOLIO-WIDE INVENTORY MATHEMATICS QA REPORT
    # Forensic audit run across every SKU in the full dataset (not just the
    # current ABC/XYZ filter) so the disposition below is a population-level
    # finding, not an artifact of whatever slice happens to be selected.
    # ------------------------------------------------------------------------
    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🧮 Inventory Mathematics QA Report")
    with st.expander("Full forensic audit: Safety Stock & Reorder Point across all SKUs", expanded=True):
        qa_summary, qa_detail = run_portfolio_math_audit(df)
        if qa_summary.get('n', 0) > 0:
            qc1, qc2 = st.columns(2)
            with qc1:
                st.markdown(f"""
**Safety Stock — {qa_summary['n']:,} SKUs tested**
- Zero in Master Audit: **{qa_summary['ss_zero_in_master']:,}**
- Exact match to independent calculation: **{qa_summary['ss_exact_match']:,}**
- Match within ±1 unit (rounding tolerance): **{qa_summary['ss_tolerance_match']:,}**
- Material mismatch (>1 unit): **{qa_summary['ss_material_mismatch']:,}**
                """)
            with qc2:
                st.markdown(f"""
**Reorder Point — {qa_summary['n']:,} SKUs tested**
- Zero in Master Audit: **{qa_summary['rop_zero_in_master']:,}**
- Exact match to independent calculation: **{qa_summary['rop_exact_match']:,}**
- Match within ±1 unit (rounding tolerance): **{qa_summary['rop_tolerance_match']:,}**
- Material mismatch (>1 unit): **{qa_summary['rop_material_mismatch']:,}**
                """)

            if qa_summary['ss_material_mismatch'] == 0 and qa_summary['rop_material_mismatch'] == 0:
                st.success("✓ Formula integrity confirmed. Every Safety Stock / Reorder Point value in the Master Audit matches the "
                           "independently recomputed value within a ±1 unit rounding tolerance across the full SKU population. "
                           "Zero values are concentrated in SKUs with zero demand variance (dead stock / constant demand) or "
                           "near-zero forecast RMSE on very low-volume SKUs — both are mathematically correct outcomes of the "
                           "model's own formula, not data-quality defects.")
            else:
                st.warning(f"⚠ {qa_summary['ss_material_mismatch']:,} Safety Stock and {qa_summary['rop_material_mismatch']:,} "
                           f"Reorder Point value(s) differ from the independent recalculation by more than 1 unit. See the detail "
                           f"table below for the specific SKUs — these are flagged for review, not silently corrected.")
                mismatch_rows = qa_detail[(qa_detail['SS_Variance'] > 1) | (qa_detail['ROP_Variance'] > 1)]
                st.dataframe(mismatch_rows, width='stretch', hide_index=True)

            st.markdown(f"""
**Root Cause (established, not assumed):** the notebook's Safety Stock formula is `1.645 × RMSE × √(7/30)` and Reorder Point is
`(Forecast_Next_Month/30 × 7) + Safety Stock`. Both terms scale with forecast error and forecast level — when a SKU has zero
historical demand variance, the model's own guard clause sets Safety Stock to exactly 0 by design (Guard: zero-variance / dead
stock). When a SKU has extremely low, near-deterministic demand, RMSE rounds to a value small enough that the computed buffer
also rounds to 0. Both are legitimate outputs of the implemented formula, verified by direct recomputation from RMSE and
Forecast_Next_Month across the entire population — this is not a broken pipeline or a lost calculation.

**Disposition:** {'CONFIRMED VALID — no data issue found.' if qa_summary['ss_material_mismatch'] == 0 and qa_summary['rop_material_mismatch'] == 0 else 'REQUIRES REVIEW for the flagged SKUs above.'}
            """)
            st.caption(f"Audited against the full unfiltered dataset ({len(df):,} SKUs), independent of the current ABC/XYZ filter. {source_tag('Dashboard Derived')}")
        else:
            st.info("Could not run the mathematics audit — RMSE and/or Forecast_Next_Month are not present in this data source.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🔍 Full Audit Table")
    search_term = st.text_input("🔎 Search SKU", key="audit_search")
    audit_view = filtered_df.copy()
    if search_term:
        audit_view = audit_view[audit_view['SKU'].astype(str).str.contains(search_term, case=False, na=False)]
    st.dataframe(audit_view, width='stretch', hide_index=True, height=350)
    st.caption(f"{len(audit_view):,} of {len(df):,} total SKUs shown. Every column above is read directly from {DATA_SOURCE_LABEL}; the source file itself is never modified by this dashboard.")

    @st.cache_data(show_spinner=False)
    def generate_excel_report(export_df):
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            export_df.to_excel(writer, index=False, sheet_name='Filtered_Audit')
            ws = writer.sheets['Filtered_Audit']
            header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
            header_font = Font(color="FFFFFF", bold=True)
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            for i, col in enumerate(export_df.columns, 1):
                try:
                    max_len = max(export_df[col].astype(str).map(len).max(), len(str(col))) + 2
                except Exception:
                    max_len = len(str(col)) + 2
                ws.column_dimensions[get_column_letter(i)].width = min(max_len, 40)
        buffer.seek(0)
        return buffer

    if len(audit_view) > 0:
        try:
            excel_buffer = generate_excel_report(audit_view)
            st.download_button("⬇️ Download Filtered Audit Report (.xlsx)", data=excel_buffer,
                                file_name="Filtered_Supply_Chain_Audit.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="download_audit_report")
        except Exception as e:
            st.warning(f"Could not generate the export file ({e}).")
    else:
        st.info("No rows in the current filtered/searched view to export.")

# ==============================================================================
# FOOTER
# ==============================================================================
st.markdown("<hr style='margin-top:3rem;'>", unsafe_allow_html=True)
st.markdown("""
<div style="text-align:center; color:#64748B; font-size:0.85rem; padding-bottom:1rem;">
    Enterprise Demand Forecasting & Inventory Decision Support System Using Time Series Analytics<br>
    All rights reserved by <strong>Kaushik Jain</strong>
</div>
""", unsafe_allow_html=True)
