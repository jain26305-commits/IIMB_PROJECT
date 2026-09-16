"""
Enterprise Demand Forecasting & Inventory Decision Support System
Using Time Series Analytics — Premium Edition (v2.1 Audited)

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
import numpy as np

# --- NumPy 2.0+ compatibility shims ---
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

# Optional / heavy dependencies made resilient
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

st.set_page_config(
    page_title="Enterprise Demand Forecasting & Inventory Decision Support",
    page_icon="🧊",
    layout="wide",
    initial_sidebar_state="collapsed"
)

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
# PRESENTATIONAL PLOTLY WRAPPER
# ------------------------------------------------------------------------------
_ORIGINAL_PLOTLY_CHART = st.plotly_chart

def _cinematic_plotly_chart(fig, *args, **kwargs):
    try:
        fig.update_layout(
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="#ffffff",
            font=dict(family="'Inter', 'Segoe UI', -apple-system, sans-serif", size=13, color="#334155"),
            title_font=dict(family="'Inter', 'Segoe UI', -apple-system, sans-serif", size=16, color="#1E3A8A"),
            legend=dict(
                bgcolor="rgba(255,255,255,0.6)",
                bordercolor="rgba(226,232,240,0.8)",
                borderwidth=1,
                font=dict(color="#334155")
            ),
            hoverlabel=dict(
                bgcolor="white",
                font_size=13,
                font_family="'Inter', 'Segoe UI', sans-serif",
                bordercolor="#0EA5E9",
                font_color="#0F172A"
            ),
        )
        fig.update_xaxes(
            showline=True, linecolor="rgba(148,163,184,0.35)", gridcolor="rgba(226,232,240,0.55)",
            zerolinecolor="rgba(148,163,184,0.35)", color="#334155"
        )
        fig.update_yaxes(
            showline=True, linecolor="rgba(148,163,184,0.35)", gridcolor="rgba(226,232,240,0.55)",
            zerolinecolor="rgba(148,163,184,0.35)", color="#334155"
        )
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
        pass
    return _ORIGINAL_PLOTLY_CHART(fig, *args, **kwargs)

st.plotly_chart = _cinematic_plotly_chart

# ==============================================================================
# SECTION 0. AUTHORITATIVE FIELD MAP (Master Audit = source of truth)
# ==============================================================================
# These are the 51 authoritative fields used by the dashboard from the Enterprise Supply Chain Master Audit.
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

OPTIONAL_ARRAY_FIELDS = ['Test_Actuals', 'Test_Predictions', 'Test_Lower_Bound', 'Test_Upper_Bound']
POSSIBLE_ACTUAL_INVENTORY_COLS = ['Actual_Inventory', 'Current_Inventory', 'On_Hand_Inventory', 'Current_Stock', 'Closing_Balance']

# Fixed colour schemes
ABC_COLORS = {'A': '#1E3A8A', 'B': '#0EA5E9', 'C': '#94A3B8'}
XYZ_COLORS = {'X': '#0D9488', 'Y': '#F59E0B', 'Z': '#EF4444'}

# ==============================================================================
# SECTION 1. GENERIC HELPERS & METRIC COMPUTATION
# ==============================================================================
def col_exists(df, col):
    return df is not None and col in df.columns

def safe_series(df, col, numeric=False, default=np.nan):
    if col_exists(df, col):
        s = df[col]
        return pd.to_numeric(s, errors='coerce') if numeric else s
    return pd.Series([default] * len(df), index=df.index if df is not None else None)

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

def fmt_currency(val):
    if is_missing(val):
        return "N/A"
    try:
        return f"₹{float(val):,.2f}"
    except Exception:
        return "N/A"

def parse_pct_string(val, default=np.nan):
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

def render_metric_card(container, title, value, subtitle="", variant="ribbon", pulse=False, extra=""):
    pulse_cls = "pulse-critical" if pulse else ""
    container.markdown(
        f'<div class="metric-card metric-card--{variant} {pulse_cls}">'
        f'<div class="metric-card-title">{title}</div>'
        f'<div class="metric-card-value">{value}</div>'
        f'<div class="metric-card-subtitle">{subtitle}{extra}</div></div>',
        unsafe_allow_html=True
    )

def render_square_metric(label, value, delta=""):
    return (
        f'<div class="square-metric"><div class="square-metric-label">{label}</div>'
        f'<div class="square-metric-value">{value}</div>'
        f'<div class="square-metric-delta">{delta}</div></div>'
    )

def render_status_card(status, title, desc, action=None):
    action_html = f'<span class="status-card-action">Action: {action}</span>' if action else ""
    st.markdown(
        f'<div class="status-card" data-status="{status}">'
        f'<strong>{title}</strong><br>'
        f'<span class="status-card-desc">{desc}</span><br>'
        f'{action_html}</div>',
        unsafe_allow_html=True
    )

def render_decision_card(icon, title, text, source="Master Audit"):
    st.markdown(
        f'<div class="decision-card">'
        f'<div class="decision-card-title">{icon} {title} {source_tag(source)}</div>'
        f'<div class="decision-card-text">{text}</div></div>',
        unsafe_allow_html=True
    )

# ------------------------------------------------------------------------------
# MATHEMATICAL FORENSIC AUDIT (Safety Stock & Reorder Point)
# ------------------------------------------------------------------------------
SERVICE_LEVEL_Z = 1.645
LEAD_TIME_DAYS = 7.0
LEAD_TIME_FRACTION = LEAD_TIME_DAYS / 30.0

def independent_ss_rop(row):
    rmse = safe_val(row, 'RMSE')
    fc = safe_val(row, 'Forecast_Next_Month')
    status = str(safe_val(row, 'Status', ''))
    if 'Dead Stock' in status or 'Constant Demand' in status:
        mmd = safe_val(row, 'Mean_Monthly_Demand')
        ltd = ((float(mmd) / 30.0) * LEAD_TIME_DAYS) if not is_missing(mmd) else 0.0
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
# SECTION 2. INTRO OVERLAY
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
            .splash-overlay::before {
                content: "";
                position: absolute; inset: 0;
                background: radial-gradient(120% 90% at 50% 45%, transparent 40%, rgba(15,23,42,0.10) 100%);
                pointer-events: none;
            }
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
# SECTION 3. DESIGN TOKENS & CSS
# ==============================================================================
st.markdown("""
<style>
    :root {
        --c-navy: #1E3A8A;
        --c-navy-dark: #1E293B;
        --c-blue: #0EA5E9;
        --c-blue-light: #38BDF8;
        --c-blue-pale: #E0F2FE;

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

        --c-slate-900: #0F172A;
        --c-slate-700: #334155;
        --c-slate-600: #475569;
        --c-slate-500: #64748B;
        --c-slate-200: #E2E8F0;
        --c-slate-100: #F1F5F9;
        --c-slate-50:  #F8FAFC;
        --c-white: #FFFFFF;

        --c-rose: #F43F5E;
        --c-rose-dark: #E11D48;

        --radius-sm: 6px;
        --radius-md: 8px;
        --radius-lg: 10px;
        --radius-xl: 12px;
        --radius-pill: 20px;

        --shadow-resting: 0 1px 2px rgba(15,23,42,0.05), 0 4px 10px rgba(15,23,42,0.05);
        --shadow-hover: 0 2px 6px rgba(15,23,42,0.08), 0 20px 40px rgba(30,58,138,0.16);
        --clay-radius: 22px;
        --clay-radius-sm: 16px;
        --clay-shadow: 10px 10px 24px rgba(30,41,59,0.10), -8px -8px 18px rgba(255,255,255,0.9), inset 0 1px 0 rgba(255,255,255,0.7);
        --clay-shadow-hover: 16px 16px 36px rgba(30,41,59,0.14), -10px -10px 24px rgba(255,255,255,0.95), inset 0 1px 0 rgba(255,255,255,0.8);
        --glow-blue: 0 0 0 1px rgba(14,165,233,0.18), 0 14px 44px rgba(14,165,233,0.28);
        --glow-navy: 0 0 0 1px rgba(30,58,138,0.16), 0 14px 44px rgba(30,58,138,0.22);

        --ease-standard: cubic-bezier(0.4,0,0.2,1);
        --ease-spring: cubic-bezier(0.175,0.885,0.32,1.275);
        --dur-fast: 0.15s;
        --dur-med: 0.25s;
        --dur-slow: 0.35s;
    }

    @keyframes ambientGlow {
        0%, 100% { opacity: 0.55; transform: scale(1); }
        50% { opacity: 0.9; transform: scale(1.05); }
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
    @keyframes cardSheen {
        0% { transform: translateX(-120%) skewX(-15deg); }
        100% { transform: translateX(220%) skewX(-15deg); }
    }

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
    [data-baseweb="select"] *, [data-testid="stExpander"] *,
    input, textarea, .stDataFrame, [data-testid="stDataFrame"] * {
        color: var(--c-slate-900) !important;
    }

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

    .stApp {
        background-image: radial-gradient(circle at 15% 10%, rgba(14,165,233,0.05), transparent 45%),
                           radial-gradient(circle at 85% 90%, rgba(30,58,138,0.05), transparent 45%);
        background-size: 100% 100%;
        color-scheme: light !important;
    }

    .block-container {
        padding-top: 1.5rem;
        padding-bottom: 2rem;
        animation: fadeInUp 0.35s cubic-bezier(0.16, 1, 0.3, 1) forwards;
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, var(--c-slate-50), var(--c-slate-100));
        border-right: none;
        box-shadow: inset -1px 0 0 rgba(226,232,240,0.9), 6px 0 24px rgba(15,23,42,0.04);
    }

    .metric-card {
        background: radial-gradient(140% 100% at 0% 0%, rgba(14,165,233,0.06), transparent 55%),
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
    .metric-card--ribbon {
        padding: 1.2rem;
        min-height: 145px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
    }
    .metric-card--ribbon:hover {
        transform: translateY(-7px) scale(1.015);
        box-shadow: var(--clay-shadow-hover), var(--glow-blue);
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

    .badge { padding: 4px 12px; border-radius: var(--radius-pill); font-size: 0.75rem; font-weight: 700; display: inline-block;
        box-shadow: inset 0 1px 2px rgba(255,255,255,0.6), inset 0 -1px 2px rgba(15,23,42,0.06); }
    .badge[data-status="excellent"] { background-color: var(--c-success-bg); color: var(--c-success-text); }
    .badge[data-status="good"]      { background-color: var(--c-warning-bg); color: var(--c-warning-text); }
    .badge[data-status="attention"] { background-color: var(--c-danger-border); color: var(--c-danger-text); }
    .badge[data-status="poor"]      { background-color: var(--c-poor-bg); color: var(--c-poor-text); }

    .source-tag { font-size: 0.65rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.4px;
        padding: 2px 8px; border-radius: var(--radius-pill); margin-left: 6px; vertical-align: middle; }
    .tag-master { background: var(--c-blue-pale); color: var(--c-navy); }
    .tag-historical { background: #EDE9FE; color: #5B21B6; }
    .tag-derived { background: var(--c-slate-100); color: var(--c-slate-600); }
    .tag-illustrative { background: var(--c-warning-bg); color: var(--c-warning-text); }

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
    .square-metric:hover {
        transform: scale(1.05) translateY(-4px); box-shadow: var(--clay-shadow-hover), var(--glow-blue);
        border-top: 4px solid var(--c-navy); z-index: 10;
    }
    .square-metric-label { font-size: 0.92rem; color: var(--c-slate-600); font-weight: 800; margin-bottom: 0.35rem; line-height: 1.25; }
    .square-metric-value { font-size: 1.7rem; color: var(--c-navy); font-weight: 900; line-height: 1.25; }
    .square-metric-delta { font-size: 0.8rem; color: var(--c-slate-500); margin-top: 0.4rem; line-height: 1.4; }

    .status-card { border-radius: var(--clay-radius-sm); padding: 1.1rem 1.2rem; margin-bottom: 0.9rem;
        transition: transform var(--dur-fast) ease, box-shadow var(--dur-fast) ease; border: none; box-shadow: var(--clay-shadow); }
    .status-card:hover { transform: translateX(4px) translateY(-3px); box-shadow: var(--clay-shadow-hover); }
    .status-card[data-status="green"] { background: linear-gradient(160deg, var(--c-success-bg), #ffffff 130%); border-left: 6px solid var(--c-success); }
    .status-card[data-status="amber"] { background: linear-gradient(160deg, var(--c-warning-bg), #ffffff 130%); border-left: 6px solid var(--c-warning); }
    .status-card[data-status="red"]   { background: linear-gradient(160deg, var(--c-danger-bg), #ffffff 130%);  border-left: 6px solid var(--c-danger); }
    .status-card-desc { font-size: 0.9rem; color: var(--c-slate-700); }
    .status-card-action { font-size: 0.8rem; font-weight: bold; color: var(--c-navy); }

    .decision-card { background: linear-gradient(160deg, #fff, var(--c-slate-50) 120%);
        border: none; border-left: 6px solid var(--c-navy);
        border-radius: var(--clay-radius-sm); padding: 1rem 1.2rem; margin-bottom: 0.8rem;
        transition: transform var(--dur-fast) ease, box-shadow var(--dur-fast) ease; box-shadow: var(--clay-shadow); }
    .decision-card:hover { box-shadow: var(--clay-shadow-hover), var(--glow-navy); transform: translateY(-4px); }
    .decision-card-title { font-size: 0.85rem; font-weight: 800; color: var(--c-navy); margin-bottom: 0.25rem; }
    .decision-card-text { font-size: 0.92rem; color: var(--c-slate-700); line-height: 1.4; }

    .matrix-grid-container { display: grid; grid-template-columns: repeat(4, 1fr); gap: 1.1rem; margin-top: 1rem; }
    @media (max-width: 1200px) { .matrix-grid-container { grid-template-columns: repeat(2, 1fr); } }
    @media (max-width: 768px)  { .matrix-grid-container { grid-template-columns: 1fr; } }
    .flow-card { background: linear-gradient(165deg, #FFFFFF, var(--c-slate-50) 130%); border: none; border-radius: var(--clay-radius-sm); padding: 1.1rem;
        text-align: left; box-shadow: var(--clay-shadow); height: 100%; transition: transform var(--dur-med) ease, box-shadow var(--dur-med) ease;
        animation: chartReveal var(--dur-slow) var(--ease-standard) both; }
    .flow-card:hover { transform: translateY(-5px) scale(1.02); box-shadow: var(--clay-shadow-hover), var(--glow-blue); }
    .flow-card-title { font-size: 1rem; font-weight: 900; color: var(--c-slate-900); margin-bottom: 0.5rem; }
    .flow-card-text { font-size: 0.85rem; color: var(--c-slate-700); font-weight: 600; margin-bottom: 0.25rem; line-height: 1.35; }

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
    .hero-glow {
        position: absolute; top: 50%; left: 50%; width: 640px; height: 320px;
        transform: translate(-50%, -50%); border-radius: 50%;
        background: radial-gradient(closest-side, rgba(14,165,233,0.16), rgba(30,58,138,0.08) 60%, transparent 80%);
        animation: ambientGlow 6s ease-in-out infinite; pointer-events: none; z-index: 0;
    }
    .hero-wrap { position: relative; }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
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
# SECTION 4. SIDEBAR CONTROLS & DATA LOADING
# ==============================================================================
st.sidebar.markdown("<h2 style='color: #1E3A8A;'>🎛️ ➔ Control Panel</h2>", unsafe_allow_html=True)
st.sidebar.markdown("---")

st.sidebar.markdown("### 📁 Model / Source Data Controls")
st.sidebar.caption("Interactive Analytical Dashboard driven by Enterprise Supply Chain Master Audit.")
target_container = st.sidebar.container()
sku_container = st.sidebar.container()

dev_mode_toggle = st.sidebar.checkbox("Developer Mode: Allow synthetic demo fallback", value=False, key="dev_mode_toggle")
st.sidebar.markdown("---")

@st.cache_data(show_spinner=False)
def load_and_clean_data(file_bytes, file_name, allow_demo=False):
    data_source_label = "Master Audit"
    df = None
    try:
        if file_bytes is not None:
            if file_name.endswith('.csv'):
                df = pd.read_csv(io.BytesIO(file_bytes))
            else:
                df = pd.read_excel(io.BytesIO(file_bytes))
            data_source_label = f"Sandbox Upload ({file_name})"
        else:
            if os.path.exists("Enterprise_Supply_Chain_Master_Audit.xlsx"):
                df = pd.read_excel("Enterprise_Supply_Chain_Master_Audit.xlsx", sheet_name="Executive_Summary")
            elif allow_demo:
                df = pd.DataFrame({
                    'SKU': ['DEMO_SKU_001', 'DEMO_SKU_002', 'DEMO_SKU_003'],
                    'ABC_Class': ['A', 'B', 'C'], 'XYZ_Class': ['X', 'Y', 'Z'],
                    'ABC_XYZ_Class': ['AX', 'BY', 'CZ'],
                    'RMSE': [12.5, 22.1, 8.4], 'MAE': [9.1, 17.4, 6.2], 'Bias': [1.2, -3.4, 0.5],
                    'CV': [0.4, 1.1, 0.2], 'Bullwhip_Ratio': [1.2, 2.4, 1.1],
                    'Accuracy_Pct': [88.0, 61.0, 94.0], 'Forecast_Next_Month': [125, 65, 215],
                    'Mean_Monthly_Demand': [120, 60, 210], 'Annual_Demand': [1440, 720, 2520],
                    'Safety_Stock': [45, 80, 25], 'Reorder_Point': [90, 150, 60], 'EOQ': [60, 40, 90],
                    'Health_P_Value': [0.12, 0.03, 0.45], 'Inventory_Turnover': [5.2, 3.1, 8.4]
                })
                data_source_label = "Developer Demo Data (explicit toggle)"
            else:
                return None, "Missing Master Audit"

        if df is not None:
            df.columns = [str(c).strip() for c in df.columns]
            if 'SKU' in df.columns:
                df['SKU'] = df['SKU'].astype(str).str.strip()
                df = df[~df['SKU'].str.lower().isin(['grand total', 'total', 'nan', 'none', ''])]
            else:
                df['SKU'] = [f"SKU_{i+1}" for i in range(len(df))]
        return df, data_source_label
    except Exception as e:
        if allow_demo:
            st.error(f"⚠️ Could not read file cleanly ({e}). Falling back to minimal demo.")
            df = pd.DataFrame({
                'SKU': ['FALLBACK_001', 'FALLBACK_002'], 'ABC_Class': ['A', 'B'], 'XYZ_Class': ['X', 'Z'],
                'RMSE': [12.5, 22.1], 'CV': [0.4, 1.1], 'Bullwhip_Ratio': [1.2, 2.4], 'Inventory_Turnover': [4.0, 2.5]
            })
            return df, "Fallback (Read Error)"
        return None, "Read Error"

uploaded_file = None
file_bytes = None
file_name = "default"
df, DATA_SOURCE_LABEL = load_and_clean_data(file_bytes, file_name, allow_demo=dev_mode_toggle)

if df is None:
    st.error("Master Audit workbook not found. Dashboard cannot initialise.")
    st.stop()

@st.cache_data(show_spinner=False)
def load_raw_workbook(raw_bytes):
    if raw_bytes is None:
        return None
    try:
        xls = pd.ExcelFile(io.BytesIO(raw_bytes))
        expected_cols = {'Particulars', 'Inwards', 'Outwards'}
        frames = []
        for i, sheet in enumerate(xls.sheet_names):
            sdf = pd.read_excel(xls, sheet_name=sheet)
            sdf.columns = [str(c).strip() for c in sdf.columns]
            if not expected_cols.issubset(set(sdf.columns)):
                continue
            sdf = sdf.rename(columns={'Particulars': 'SKU'})
            sdf['SKU'] = sdf['SKU'].astype(str).str.strip()
            sdf = sdf[~sdf['SKU'].str.lower().isin(['grand total', 'total', 'nan', 'none', ''])]
            sdf['Month'] = sheet
            sdf['Month_Index'] = i
            for c in ['Opening Balance', 'Inwards', 'Outwards', 'Closing Balance']:
                if c in sdf.columns:
                    sdf[c] = pd.to_numeric(sdf[c], errors='coerce')
            if 'Outwards' in sdf.columns:
                sdf['Outwards'] = sdf['Outwards'].fillna(0.0)
            keep = [c for c in ['SKU', 'Month', 'Month_Index', 'Opening Balance', 'Inwards', 'Outwards', 'Closing Balance'] if c in sdf.columns]
            frames.append(sdf[keep])
        if not frames:
            return None
        combined = pd.concat(frames, ignore_index=True)
        return combined
    except Exception:
        return None

raw_workbook_file = None
raw_history_df = load_raw_workbook(None)
HAS_RAW_HISTORY = raw_history_df is not None and len(raw_history_df) > 0

HAS_TEST_ARRAYS = all(c in df.columns for c in OPTIONAL_ARRAY_FIELDS[:2])
HAS_XYZ = 'XYZ_Class' in df.columns
HAS_ABC_XYZ = 'ABC_XYZ_Class' in df.columns
ACTUAL_INV_COL = next((c for c in POSSIBLE_ACTUAL_INVENTORY_COLS if c in df.columns), None)
MISSING_AUTHORITATIVE_FIELDS = [c for c in MASTER_AUDIT_FIELDS if c not in df.columns]

target_container.subheader("🎯 Target Isolation")
abc_options = ['A', 'B', 'C'] if 'ABC_Class' in df.columns else []
selected_classes = target_container.multiselect("Select ABC — Demand Volume Classes", options=abc_options, default=abc_options, key="abc_filter")

if HAS_XYZ:
    xyz_options = ['X', 'Y', 'Z']
    selected_xyz = target_container.multiselect("Select XYZ — Demand Variability Classes", options=xyz_options, default=xyz_options, key="xyz_filter")
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
    st.session_state.shared_sku = None

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

if 'Forecast_Health_Score' in filtered_df.columns:
    forecast_health_avg = fhs_series.mean()
    forecast_health_source = "Master Audit"
else:
    forecast_health_avg = np.nan
    forecast_health_source = None

if total_skus > 0 and not is_missing(avg_rmse):
    share_below_rmse = float((rmse_series <= avg_rmse).sum()) / total_skus * 100
else:
    share_below_rmse = np.nan

avg_accuracy = accuracy_series.mean() if 'Accuracy_Pct' in filtered_df.columns else np.nan

if 'Forecast_Risk' in filtered_df.columns:
    high_risk_count = int((filtered_df['Forecast_Risk'].astype(str).str.strip().str.lower() == 'high').sum())
    high_risk_source = "Master Audit"
else:
    high_risk_mask = ((bullwhip_series > 2) | (cv_series > 1)) & (rmse_series > avg_rmse)
    high_risk_count = int(high_risk_mask.fillna(False).sum())
    high_risk_source = "Dashboard Derived"

inv_value_sum = inv_value_series.sum() if 'Inventory_Value' in filtered_df.columns else np.nan
wc_sum = wc_series.sum() if 'Working_Capital' in filtered_df.columns else np.nan
fin_health_avg = fin_health_series.mean() if 'KPI_Financial_Health_Score' in filtered_df.columns else np.nan
inv_health_avg = inv_health_series.mean() if 'Inventory_Health_Score' in filtered_df.columns else np.nan

# Overall qualitative status banner
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
    if high_risk_count == 0 and overall_ratio >= 0.75:
        overall_status, overall_status_key = "Stable & Healthy", "excellent"
    elif overall_ratio >= 0.5:
        overall_status, overall_status_key = "Needs Attention", "good"
    else:
        overall_status, overall_status_key = "Elevated Risk", "poor"
else:
    overall_status, overall_status_key = "Insufficient Data", "attention"

# ==============================================================================
# SECTION 5. DASHBOARD TABS
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
        f'🧭 Dashboard-Derived Portfolio Status: <strong>{overall_status}</strong> &nbsp;|&nbsp; Data Source: <strong>{DATA_SOURCE_LABEL}</strong> &nbsp;|&nbsp; SKUs in View: <strong>{total_skus:,}</strong></div>',
        unsafe_allow_html=True)

    r1c1, r1c2, r1c3, r1c4 = st.columns(4)
    render_metric_card(r1c1, "Total Tracked SKUs", f"{total_skus:,}", "System Population", extra=source_tag("Master Audit"))

    if forecast_health_source:
        fh_status = score_to_status(forecast_health_avg)
        render_metric_card(r1c2, "Forecast Health Score", fmt_num(forecast_health_avg, 0, "%"), badge_html(fh_status, fh_status.title()), extra=source_tag("Master Audit"))
    else:
        sb_status = score_to_status(share_below_rmse)
        render_metric_card(r1c2, "Share Below Portfolio RMSE", fmt_num(share_below_rmse, 0, "%"), badge_html(sb_status, sb_status.title()), extra=source_tag("Dashboard Derived"))

    render_metric_card(r1c3, "Forecast Accuracy", fmt_num(avg_accuracy, 1, "%") if not is_missing(avg_accuracy) else "N/A", "Portfolio Average", extra=source_tag("Master Audit") if 'Accuracy_Pct' in filtered_df.columns else source_tag("Dashboard Derived"))
    render_metric_card(r1c4, "High-Risk SKUs", f"{high_risk_count:,}", "Requires Review", variant="ribbon", pulse=(high_risk_count > 0), extra=source_tag(high_risk_source))

    r2c1, r2c2, r2c3 = st.columns(3)
    render_metric_card(r2c1, "Model-Implied Inventory Value", fmt_currency(inv_value_sum), "Sum Across Filtered SKUs", extra=source_tag("Master Audit"))
    render_metric_card(r2c2, "Model-Implied Working Capital", fmt_currency(wc_sum), "Sum Across Filtered SKUs", extra=source_tag("Master Audit"))
    render_metric_card(r2c3, "Financial Health Score", fmt_num(fin_health_avg, 0, "%") if not is_missing(fin_health_avg) else "N/A", "Portfolio Average", extra=source_tag("Master Audit"))

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<hr>", unsafe_allow_html=True)

    ov_col1, ov_col2 = st.columns(2)
    with ov_col1:
        st.subheader("📦 Portfolio Composition — ABC — Demand Volume")
        st.markdown('<div class="section-caption">Cumulative demand volume classification (A: top 80%, B: next 15%, C: tail 5%).</div>', unsafe_allow_html=True)
        if 'ABC_Class' in filtered_df.columns and total_skus > 0:
            abc_counts = filtered_df['ABC_Class'].value_counts().reindex(['A', 'B', 'C'], fill_value=0).reset_index()
            abc_counts.columns = ['ABC_Class', 'Count']
            fig_abc = px.pie(
                abc_counts, names='ABC_Class', values='Count', hole=0.5,
                category_orders={'ABC_Class': ['A', 'B', 'C']},
                color='ABC_Class', color_discrete_map=ABC_COLORS,
                title="SKU Count Share by ABC — Demand Volume"
            )
            fig_abc.update_traces(textinfo='label+percent', marker=dict(line=dict(color='#FFFFFF', width=2)))
            fig_abc.update_layout(plot_bgcolor='white', margin=dict(l=10, r=10, t=40, b=10))
            st.plotly_chart(fig_abc, width='stretch')
        else:
            st.info("ABC_Class not available in this data source.")

    with ov_col2:
        st.subheader("🌡️ Portfolio Composition — XYZ — Demand Variability")
        st.markdown('<div class="section-caption">Demand coefficient of variation classification (X: low variability, Y: moderate, Z: high).</div>', unsafe_allow_html=True)
        if HAS_XYZ and total_skus > 0:
            xyz_counts = filtered_df['XYZ_Class'].value_counts().reindex(['X', 'Y', 'Z'], fill_value=0).reset_index()
            xyz_counts.columns = ['XYZ_Class', 'Count']
            fig_xyz = px.pie(
                xyz_counts, names='XYZ_Class', values='Count', hole=0.5,
                category_orders={'XYZ_Class': ['X', 'Y', 'Z']},
                color='XYZ_Class', color_discrete_map=XYZ_COLORS,
                title="SKU Count Share by XYZ — Demand Variability"
            )
            fig_xyz.update_traces(textinfo='label+percent', marker=dict(line=dict(color='#FFFFFF', width=2)))
            fig_xyz.update_layout(plot_bgcolor='white', margin=dict(l=10, r=10, t=40, b=10))
            st.plotly_chart(fig_xyz, width='stretch')
        else:
            st.info("XYZ_Class not available in this data source.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🧩 ABC — Demand Volume × XYZ — Demand Variability Matrix")
    st.markdown('<div class="section-caption">Count of SKUs by combined ABC (Demand Volume) × XYZ (Demand Variability) classification based on cumulative demand volume.</div>', unsafe_allow_html=True)
    if HAS_ABC_XYZ and total_skus > 0:
        abc_order = ['A', 'B', 'C']
        xyz_order = ['X', 'Y', 'Z']
        codes = filtered_df['ABC_XYZ_Class'].astype(str)
        valid_code = codes.str.len().eq(2) & codes.str[0].isin(abc_order) & codes.str[1].isin(xyz_order)
        valid_codes = codes[valid_code]
        matrix = pd.crosstab(valid_codes.str[0], valid_codes.str[1]).reindex(index=abc_order, columns=xyz_order, fill_value=0)
        fig_matrix = px.imshow(
            matrix, text_auto=True, color_continuous_scale=[[0, '#F1F5F9'], [1, '#1E3A8A']],
            labels=dict(x="XYZ — Demand Variability", y="ABC — Demand Volume", color="SKU Count")
        )
        fig_matrix.update_layout(margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig_matrix, width='stretch')
    else:
        st.info("ABC_XYZ_Class not available in source data.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("📊 Annualised Demand by ABC — Demand Volume")
    st.markdown(f'<div class="section-caption">Annualised Demand = full-period mean monthly demand × 12 (annualised model baseline, not necessarily trailing-12-month actual demand). {source_tag("Master Audit")}</div>', unsafe_allow_html=True)
    if 'Annual_Demand' in filtered_df.columns and 'ABC_Class' in filtered_df.columns and total_skus > 0:
        demand_by_abc = filtered_df.groupby('ABC_Class', as_index=False)['Annual_Demand'].sum()
        demand_by_abc['ABC_Class'] = pd.Categorical(demand_by_abc['ABC_Class'], categories=['A', 'B', 'C'], ordered=True)
        demand_by_abc = demand_by_abc.sort_values('ABC_Class').dropna(subset=['ABC_Class'])
        fig_demand = px.bar(
            demand_by_abc, x='ABC_Class', y='Annual_Demand', color='ABC_Class',
            category_orders={'ABC_Class': ['A', 'B', 'C']},
            color_discrete_map=ABC_COLORS,
            title="Total Annualised Demand (units) by ABC — Demand Volume"
        )
        fig_demand.update_layout(plot_bgcolor='white', showlegend=False, margin=dict(l=10, r=10, t=40, b=10))
        st.plotly_chart(fig_demand, width='stretch')
    else:
        st.info("Annual_Demand not available in this data source.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🚨 Priority Management Review")
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
            attn_df = attn_df.sort_values('Inventory_Value', ascending=False)
        st.dataframe(attn_df.head(25), width='stretch', hide_index=True)
        st.caption(f"{len(attn_df):,} SKU(s) flagged as High Risk or High Priority in current filter. Showing top 25 by Model-Implied Inventory Value.")
    else:
        st.info("Risk/priority fields not available to build the attention list.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🧭 Decision Support — Priority Decisions")
    st.markdown('<div class="section-caption">Model-informed recommendations synthesized from validation signals. Conditional on operational feasibility.</div>', unsafe_allow_html=True)
    if 'Executive_Recommendation' in filtered_df.columns and len(attn_df) > 0:
        priority_recs_df = filtered_df.loc[attn_df.index]
        top_recs = priority_recs_df['Executive_Recommendation'].value_counts().head(5)
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
                unsafe_allow_html=True
            )
            jump_col1, jump_col2 = st.columns([3, 1])
            with jump_col1:
                pick = st.selectbox("SKU affected by this decision", rec_skus, key=f"priority_pick_{rec_i}", label_visibility="collapsed")
            with jump_col2:
                if st.button("🎯 Jump to SKU", key=f"priority_jump_{rec_i}", width='stretch'):
                    st.session_state['_pending_sku_jump'] = pick
                    st.rerun()

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

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🧭 What This Project Enables")
    st.markdown('<div class="section-caption">Capabilities the analytics deliver, drawn directly from Master Audit outputs.</div>', unsafe_allow_html=True)

    wc_cols = st.columns(3)
    wc_items = [
        ("📈 Forecasting", "SKU-level validated forecasting and model-competition outputs."),
        ("📦 Inventory", "Forecast-driven Safety Stock, ROP, and EOQ recommendations."),
        ("🧩 Segmentation", "ABC–XYZ differentiated inventory control based on demand volume and variability."),
    ]
    for col, (title, desc) in zip(wc_cols, wc_items):
        col.markdown(f'<div class="flow-card" style="border-top:5px solid var(--c-blue);"><div class="flow-card-title">{title}</div><div class="flow-card-text">{desc}</div></div>', unsafe_allow_html=True)

    wc_cols2 = st.columns(3)
    wc_items2 = [
        ("⚠️ Risk", "Forecast-risk, RMSE, and Bullwhip-based signal amplification visibility."),
        ("💰 Finance", "Inventory Value, Working Capital, Carrying Cost, and Stockout exposure visibility."),
        ("🧭 Decision Support", "Rule-based decision recommendations linked directly to analytical outputs."),
    ]
    for col, (title, desc) in zip(wc_cols2, wc_items2):
        col.markdown(f'<div class="flow-card" style="border-top:5px solid var(--c-blue);"><div class="flow-card-title">{title}</div><div class="flow-card-text">{desc}</div></div>', unsafe_allow_html=True)

# ==============================================================================
# TAB 2 — DEMAND & FORECAST INTELLIGENCE
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
        abc_badge = f"ABC — Demand Volume: <strong>{safe_val(sku_row, 'ABC_Class', 'N/A')}</strong>"
        xyz_badge = f" &nbsp;|&nbsp; XYZ — Demand Variability: <strong>{safe_val(sku_row, 'XYZ_Class', 'N/A')}</strong>" if HAS_XYZ else ""
        safe_sku = html.escape(str(st.session_state.shared_sku))
        st.markdown(f"#### Target SKU: **{safe_sku}** &nbsp; ({abc_badge}{xyz_badge})", unsafe_allow_html=True)

        st.markdown("<h4 style='color: #475569;'>📊 Model Diagnostic Brief</h4>", unsafe_allow_html=True)
        with st.container(border=True):
            rmse_val = safe_val(sku_row, 'RMSE')
            bw_val = safe_val(sku_row, 'Bullwhip_Ratio')
            acc_val = safe_val(sku_row, 'Accuracy_Pct')
            p_val = safe_val(sku_row, 'Health_P_Value')

            insight = f"**Model Summary:** The forecasting model for **{st.session_state.shared_sku}** "
            if not is_missing(rmse_val) and not is_missing(avg_rmse):
                if rmse_val > avg_rmse:
                    insight += f"shows above-portfolio-average forecast error (RMSE {fmt_num(rmse_val,1)} vs. portfolio average {fmt_num(avg_rmse,1)}). Analytical review is recommended. "
                else:
                    insight += f"tracks within normal bounds (RMSE {fmt_num(rmse_val,1)}, portfolio average {fmt_num(avg_rmse,1)}). "
            else:
                insight += "does not have an RMSE value available for comparison. "

            if not is_missing(bw_val):
                if bw_val > 1.5:
                    insight += f"Bullwhip Ratio of {fmt_num(bw_val,2)} indicates Model-Derived Replenishment-Signal Amplification relative to reference level (1.0). Where elevated, review model replenishment dynamics and order batching assumptions. "
                else:
                    insight += f"Bullwhip Ratio of {fmt_num(bw_val,2)} indicates baseline replenishment-signal amplification close to reference level. "
            else:
                insight += "Bullwhip Ratio is N/A. "

            if not is_missing(p_val):
                health_status = (
                    "Acceptable — no significant residual autocorrelation detected at the tested lag"
                    if float(p_val) > 0.05
                    else "Review — significant residual autocorrelation detected at the tested lag"
                )
                insight += f"\n\n**Residual Diagnostic:** {health_status} (Ljung-Box p-value: {fmt_num(p_val,3)})."
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
                                                          hoverinfo="skip", showlegend=True, name='Model Validation Band'))
                        except Exception:
                            pass
                        fig.add_trace(go.Scatter(x=months, y=actuals, name='Actual Demand', mode='lines+markers',
                                                  line=dict(color='#0EA5E9', width=3), marker=dict(size=8, color='#0284C7')))
                        fig.add_trace(go.Scatter(x=months, y=preds, name='Forecast', mode='lines+markers',
                                                  line=dict(color='#F43F5E', width=3, dash='dash'), marker=dict(size=8, symbol='diamond', color='#E11D48')))
                        fig.update_layout(title="Forecast Validation — Test Period", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                                           hovermode='x unified', legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                                           xaxis=dict(showgrid=False), yaxis=dict(showgrid=True, gridcolor='#F1F5F9'), margin=dict(l=20, r=20, t=40, b=20))
                        st.plotly_chart(fig, width='stretch')
                        st.caption("Test Period = 4-month chronological validation window.")
                    else:
                        st.info("Test-period arrays present but malformed for this SKU.")
                except Exception:
                    st.warning("Could not parse test-period arrays for this SKU.")
            else:
                fc_next = safe_val(sku_row, 'Forecast_Next_Month')
                mean_m = safe_val(sku_row, 'Mean_Monthly_Demand')
                snap_df = pd.DataFrame({
                    'Metric': ['Mean Monthly Demand (Full Period)', 'Forecast — Next Month'],
                    'Value': [mean_m if not is_missing(mean_m) else np.nan, fc_next if not is_missing(fc_next) else np.nan]
                }).dropna()
                if len(snap_df) > 0:
                    fig_snap = px.bar(snap_df, x='Metric', y='Value', color='Metric',
                                       color_discrete_sequence=['#94A3B8', '#0EA5E9'], title="Forecast Snapshot")
                    fig_snap.update_layout(plot_bgcolor='white', showlegend=False, margin=dict(l=20, r=20, t=40, b=20))
                    st.plotly_chart(fig_snap, width='stretch')
                st.markdown(f"""<div class="section-caption">Month-by-month validation arrays are not included in this Master Audit export —
                aggregate validation statistics (Accuracy_Pct, MAE, RMSE, Bias) are shown. {source_tag("Master Audit")}</div>""", unsafe_allow_html=True)

        with t2_col2:
            st.markdown("<br>", unsafe_allow_html=True)
            st.markdown("<h4 style='color: #475569; text-align:center;'>⚡ Forecast KPIs</h4>", unsafe_allow_html=True)
            acc_val_disp = safe_val(sku_row, 'Accuracy_Pct')
            fc_val_disp = safe_val(sku_row, 'Forecast_Next_Month')
            bw_val_disp = safe_val(sku_row, 'Bullwhip_Ratio')

            mini_html = (
                '<div class="metrics-grid metrics-grid--3">'
                + render_square_metric("Next Forecast", fmt_int(fc_val_disp), f"Projected {source_tag('Master Audit')}")
                + render_square_metric("Signal Amp.", fmt_num(bw_val_disp, 2), f"Bullwhip Ratio {source_tag('Master Audit')}")
                + render_square_metric("Forecast Accuracy", fmt_num(acc_val_disp, 1, '%'), f"sMAPE Derived {source_tag('Master Audit')}")
                + '</div>'
            )
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
            st.markdown(f"**Annualised Demand:** {fmt_int(safe_val(sku_row, 'Annual_Demand'))} units", unsafe_allow_html=True)
    else:
        st.info("Select a SKU from the sidebar to view its forecast diagnostics.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🎯 Portfolio Forecast Quality")
    st.markdown('<div class="section-caption">Computed from authoritative per-SKU RMSE, MAE, Bias, and Accuracy_Pct fields in the Master Audit.</div>', unsafe_allow_html=True)

    if total_skus > 0:
        fq1, fq2, fq3 = st.columns(3)
        with fq1:
            fig_rmse = px.histogram(filtered_df, x='RMSE', title="RMSE Distribution", color_discrete_sequence=['#0EA5E9']) if 'RMSE' in filtered_df.columns else None
            if fig_rmse:
                fig_rmse.update_layout(plot_bgcolor='white', margin=dict(l=20, r=20, t=40, b=20))
                st.plotly_chart(fig_rmse, width='stretch')
        with fq2:
            fig_mae = px.histogram(filtered_df, x='MAE', title="MAE Distribution", color_discrete_sequence=['#1E3A8A']) if 'MAE' in filtered_df.columns else None
            if fig_mae:
                fig_mae.update_layout(plot_bgcolor='white', margin=dict(l=20, r=20, t=40, b=20))
                st.plotly_chart(fig_mae, width='stretch')
        with fq3:
            fig_bias = px.histogram(filtered_df, x='Bias', title="Bias Distribution", color_discrete_sequence=['#F43F5E']) if 'Bias' in filtered_df.columns else None
            if fig_bias:
                fig_bias.update_layout(plot_bgcolor='white', margin=dict(l=20, r=20, t=40, b=20))
                st.plotly_chart(fig_bias, width='stretch')

        if not is_missing(avg_accuracy):
            q_class = "Excellent" if avg_accuracy >= 85 else ("Good" if avg_accuracy >= 70 else ("Average" if avg_accuracy >= 50 else "Poor"))
            fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number", value=float(avg_accuracy),
                title={'text': f"Portfolio Forecast Accuracy: {fmt_num(avg_accuracy, 1)}%<br><span style='font-size:12px;color:#64748B;'>Dashboard Presentation Band: {q_class}</span>", 'font': {'size': 16, 'color': '#1E3A8A'}},
                gauge={'axis': {'range': [0, 100]}, 'bar': {'color': "#0EA5E9"},
                       'steps': [{'range': [0, 50], 'color': "#FEE2E2"}, {'range': [50, 70], 'color': "#FEF3C7"},
                                 {'range': [70, 85], 'color': "#E0F2FE"}, {'range': [85, 100], 'color': "#D1FAE5"}]}
            ))
            fig_gauge.update_layout(margin=dict(l=20, r=20, t=80, b=20), height=250)
            st.plotly_chart(fig_gauge, width='stretch')
            st.caption(f"Portfolio average Accuracy_Pct across {total_skus:,} SKUs in view. Qualitative bands are a Dashboard Presentation Band, not an engine classification.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🗺️ Forecast Risk by ABC — Demand Volume")
    if 'ABC_Class' in filtered_df.columns and 'Forecast_Risk' in filtered_df.columns and total_skus > 0:
        risk_matrix = pd.crosstab(filtered_df['ABC_Class'], filtered_df['Forecast_Risk']).reindex(index=['A', 'B', 'C'], fill_value=0)
        fig_riskhm = px.imshow(
            risk_matrix, text_auto=True, color_continuous_scale=[[0, '#F1F5F9'], [1, '#991B1B']],
            labels=dict(x="Forecast Risk", y="ABC — Demand Volume", color="SKU Count")
        )
        fig_riskhm.update_layout(margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig_riskhm, width='stretch')

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🔁 What the Forecasting Layer Enables")
    st.markdown(
        f'<div class="decision-card"><div class="decision-card-title">📏 Validation-Driven Model Competition {source_tag("Master Audit")}</div>'
        f'<div class="decision-card-text">Four forecasting pathways (3-Month Moving Average, Naive, Seasonal Naive, and demand-aware AutoARIMA) '
        f'compete over a 20-month training window. The winning model is determined by lowest validation sMAPE on a 4-month chronological holdout, '
        f'yielding an average validation accuracy of {fmt_num(avg_accuracy, 1, "%") if not is_missing(avg_accuracy) else "N/A"} across {total_skus:,} SKUs.</div></div>',
        unsafe_allow_html=True
    )

# ==============================================================================
# TAB 3 — INVENTORY OPTIMISATION
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
        cls_line = f"ABC — Demand Volume: <strong>{safe_val(sku_row3,'ABC_Class','N/A')}</strong>"
        if HAS_XYZ:
            cls_line += f" &nbsp;|&nbsp; XYZ — Demand Variability: <strong>{safe_val(sku_row3,'XYZ_Class','N/A')}</strong>"
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
        render_metric_card(inv_cols2[0], "Model-Implied Inventory Value", fmt_currency(safe_val(sku_row3, 'Inventory_Value')), "", extra=source_tag("Master Audit"))
        render_metric_card(inv_cols2[1], "Model-Implied Working Capital", fmt_currency(safe_val(sku_row3, 'Working_Capital')), "", extra=source_tag("Master Audit"))
        render_metric_card(inv_cols2[2], "Model-Implied Inventory Days", fmt_num(safe_val(sku_row3, 'Inventory_Days'), 0), "days", extra=source_tag("Master Audit"))
        render_metric_card(inv_cols2[3], "Inventory Turnover", fmt_num(safe_val(sku_row3, 'Inventory_Turnover'), 2), "turns / yr", extra=source_tag("Master Audit"))

        st.markdown("<br>", unsafe_allow_html=True)
        svc_col1, svc_col2, svc_col3 = st.columns(3)
        with svc_col1:
            st.metric("Model-Implied Service Level", fmt_num(safe_val(sku_row3, 'Service_Level_Pct'), 1, "%"))
        with svc_col2:
            st.metric("Model-Implied Fill Rate", fmt_num(safe_val(sku_row3, 'Fill_Rate_Pct'), 1, "%"))
        with svc_col3:
            ihs = safe_val(sku_row3, 'Inventory_Health_Score')
            ihs_status = score_to_status(ihs)
            st.markdown(f"**Inventory Health Score:** {fmt_num(ihs,0)} {badge_html(ihs_status, ihs_status.title())}", unsafe_allow_html=True)

        pol_col1, pol_col2, pol_col3 = st.columns(3)
        pol_col1.markdown(f"**Planning Frequency:** {safe_val(sku_row3, 'Planning Frequency', 'N/A')}")
        pol_col2.markdown(f"**Inventory Policy:** {safe_val(sku_row3, 'Inventory Policy', 'N/A')}")
        pol_col3.markdown(f"**Review Frequency:** {safe_val(sku_row3, 'Review Frequency', 'N/A')}")

        st.markdown("<hr>", unsafe_allow_html=True)
        st.subheader("🧮 Inventory Mathematics Integrity — Selected SKU")
        with st.expander("Master Audit vs. Independent Validation", expanded=False):
            val_result = independent_ss_rop(sku_row3)
            ma_ss = safe_val(sku_row3, 'Safety_Stock')
            ma_rop = safe_val(sku_row3, 'Reorder_Point')
            ind_ss, ind_rop = val_result['ss'], val_result['rop']
            ss_var = abs(ma_ss - ind_ss) if (not is_missing(ma_ss) and not is_missing(ind_ss)) else None
            rop_var = abs(ma_rop - ind_rop) if (not is_missing(ma_rop) and not is_missing(ind_rop)) else None
            ss_ok = ss_var is not None and ss_var <= 1
            rop_ok = rop_var is not None and rop_var <= 1
            st.markdown(
                f'<table class="math-audit-table">'
                f'<tr><th></th><th>Master Audit</th><th>Independent Validation</th><th>Variance</th><th>Status</th></tr>'
                f'<tr><td><strong>Safety Stock</strong></td><td>{fmt_num(ma_ss,0)}</td><td>{fmt_num(ind_ss,0)}</td><td>{fmt_num(ss_var,0) if ss_var is not None else "N/A"}</td>'
                f'<td class="{"math-status-ok" if ss_ok else "math-status-review"}">{"✓ Confirmed" if ss_ok else "⚠ Review"}</td></tr>'
                f'<tr><td><strong>Reorder Point</strong></td><td>{fmt_num(ma_rop,0)}</td><td>{fmt_num(ind_rop,0)}</td><td>{fmt_num(rop_var,0) if rop_var is not None else "N/A"}</td>'
                f'<td class="{"math-status-ok" if rop_ok else "math-status-review"}">{"✓ Confirmed" if rop_ok else "⚠ Review"}</td></tr>'
                f'</table>',
                unsafe_allow_html=True
            )
            st.caption(f"Basis: {val_result['basis']}. Formula: Safety Stock = 1.645 × RMSE × √(7/30); Reorder Point = (Forecast_Next_Month/30 × 7) + Safety Stock. {source_tag('Dashboard Derived')}")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🔁 Portfolio Replenishment View")
    repl_cols = [c for c in ['SKU', 'ABC_XYZ_Class', 'ABC_Class', 'Safety_Stock', 'Reorder_Point', 'EOQ',
                             'Inventory_Days', 'Inventory_Turnover', 'Inventory_Health_Score', 'Inventory_Recommendation'] if c in filtered_df.columns]
    if repl_cols and total_skus > 0:
        repl_df = filtered_df[repl_cols].copy()
        if 'Inventory_Health_Score' in repl_df.columns:
            repl_df = repl_df.sort_values('Inventory_Health_Score', ascending=True)
        st.dataframe(repl_df.head(50), width='stretch', hide_index=True)
        st.caption(f"Showing {min(50, len(repl_df)):,} of {len(repl_df):,} SKUs, sorted by lowest Inventory Health Score first. {source_tag('Master Audit')}")

    repl_col1, repl_col2 = st.columns(2)
    with repl_col1:
        if 'Inventory_Turnover' in filtered_df.columns and 'ABC_Class' in filtered_df.columns and total_skus > 0:
            turn_by_abc = filtered_df.groupby('ABC_Class', as_index=False)['Inventory_Turnover'].mean()
            turn_by_abc['ABC_Class'] = pd.Categorical(turn_by_abc['ABC_Class'], categories=['A', 'B', 'C'], ordered=True)
            turn_by_abc = turn_by_abc.sort_values('ABC_Class').dropna(subset=['ABC_Class'])
            fig_turn = px.bar(
                turn_by_abc, x='ABC_Class', y='Inventory_Turnover', color='ABC_Class',
                category_orders={'ABC_Class': ['A', 'B', 'C']},
                color_discrete_map=ABC_COLORS,
                title="Average Inventory Turnover by ABC — Demand Volume"
            )
            fig_turn.update_layout(plot_bgcolor='white', showlegend=False, margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_turn, width='stretch')
    with repl_col2:
        if 'Inventory_Health_Score' in filtered_df.columns and total_skus > 0:
            fig_ihs = px.histogram(filtered_df, x='Inventory_Health_Score', title="Inventory Health Score Distribution", color_discrete_sequence=['#10B981'])
            fig_ihs.update_layout(plot_bgcolor='white', margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_ihs, width='stretch')

# ==============================================================================
# TAB 4 — RISK, SEGMENTATION & DIAGNOSTICS
# ==============================================================================
with tab4:
    st.subheader("⚠️ Replenishment-Signal Amplification vs. Forecast Error")
    st.markdown(f'<div class="section-caption">Quadrant boundaries represent median-split reference cuts on the active filter. {source_tag("Dashboard Derived")}</div>', unsafe_allow_html=True)

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
        fig_scatter = px.scatter(
            scatter_df, x='_bw', y='_rmse', color='Diagnostic Quadrant',
            color_discrete_map={"Stable": "#10B981", "Forecast Issue": "#F59E0B", "Signal Amplification": "#0EA5E9", "Critical": "#EF4444"},
            hover_data=['SKU'] if 'SKU' in scatter_df.columns else None,
            labels={'_bw': 'Bullwhip Ratio (Model-Derived Replenishment-Signal Amplification)', '_rmse': 'RMSE'},
            title="Replenishment-Signal Amplification vs. Forecast Error — Dashboard-Derived Review Lens"
        )
        fig_scatter.add_vline(x=med_bw, line_dash="dot", line_color="#94A3B8")
        fig_scatter.add_hline(y=med_rmse, line_dash="dot", line_color="#94A3B8")
        fig_scatter.update_traces(marker=dict(size=9, opacity=0.75, line=dict(width=1, color='white')))
        fig_scatter.update_layout(plot_bgcolor='white', margin=dict(l=20, r=20, t=40, b=20))
        st.plotly_chart(fig_scatter, width='stretch')

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("📊 Segmentation Distributions")
    seg1, seg2 = st.columns(2)
    with seg1:
        if 'ABC_Class' in filtered_df.columns and total_skus > 0:
            abc_bar = filtered_df['ABC_Class'].value_counts().reindex(['A', 'B', 'C'], fill_value=0).reset_index()
            abc_bar.columns = ['ABC_Class', 'Count']
            fig_abc_bar = px.bar(
                abc_bar, x='ABC_Class', y='Count', color='ABC_Class',
                category_orders={'ABC_Class': ['A', 'B', 'C']},
                color_discrete_map=ABC_COLORS,
                title="ABC — Demand Volume Distribution"
            )
            fig_abc_bar.update_layout(plot_bgcolor='white', showlegend=False, margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_abc_bar, width='stretch')
    with seg2:
        if HAS_XYZ and total_skus > 0:
            xyz_bar = filtered_df['XYZ_Class'].value_counts().reindex(['X', 'Y', 'Z'], fill_value=0).reset_index()
            xyz_bar.columns = ['XYZ_Class', 'Count']
            fig_xyz_bar = px.bar(
                xyz_bar, x='XYZ_Class', y='Count', color='XYZ_Class',
                category_orders={'XYZ_Class': ['X', 'Y', 'Z']},
                color_discrete_map=XYZ_COLORS,
                title="XYZ — Demand Variability Distribution"
            )
            fig_xyz_bar.update_layout(plot_bgcolor='white', showlegend=False, margin=dict(l=20, r=20, t=40, b=20))
            st.plotly_chart(fig_xyz_bar, width='stretch')

    st.subheader("🧩 ABC — Demand Volume × XYZ — Demand Variability Risk Intensity")
    st.markdown(f'<div class="section-caption">Average RMSE per ABC × XYZ segment cell. {source_tag("Dashboard Derived")}</div>', unsafe_allow_html=True)
    if HAS_ABC_XYZ and 'RMSE' in filtered_df.columns and total_skus > 0:
        abc_order = ['A', 'B', 'C']
        xyz_order = ['X', 'Y', 'Z']
        risk_intensity = pd.DataFrame(np.nan, index=abc_order, columns=xyz_order)
        tmp = filtered_df.copy()
        tmp['_A'] = tmp['ABC_XYZ_Class'].astype(str).str[0]
        tmp['_X'] = tmp['ABC_XYZ_Class'].astype(str).str[1]
        grp = tmp.groupby(['_A', '_X'])['RMSE'].mean()
        for (a, x), v in grp.items():
            if a in abc_order and x in xyz_order:
                risk_intensity.loc[a, x] = v
        fig_intensity = px.imshow(
            risk_intensity, text_auto='.1f', color_continuous_scale=[[0, '#D1FAE5'], [0.5, '#FEF3C7'], [1, '#991B1B']],
            labels=dict(x="XYZ — Demand Variability", y="ABC — Demand Volume", color="Avg RMSE")
        )
        fig_intensity.update_layout(margin=dict(l=10, r=10, t=20, b=10))
        st.plotly_chart(fig_intensity, width='stretch')

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🔴 High-Risk SKU Register")
    hr_cols = [c for c in ['SKU', 'ABC_XYZ_Class', 'RMSE', 'Bullwhip_Ratio', 'CV', 'Forecast_Risk',
                           'Business Risk', 'Priority_Level', 'Supply Chain Priority', 'Model_Reliability'] if c in filtered_df.columns]
    if hr_cols and 'Forecast_Risk' in filtered_df.columns and total_skus > 0:
        hr_df = filtered_df[filtered_df['Forecast_Risk'].astype(str).str.lower() == 'high'].copy()
        sort_cols = []
        sort_asc = []
        if 'Priority_Level' in hr_df.columns:
            prio_map = {'high': 1, 'medium': 2, 'low': 3}
            hr_df['_prio'] = hr_df['Priority_Level'].astype(str).str.lower().map(prio_map).fillna(99)
            sort_cols.append('_prio')
            sort_asc.append(True)
        elif 'Supply Chain Priority' in hr_df.columns:
            prio_map = {'high': 1, 'medium': 2, 'low': 3}
            hr_df['_prio'] = hr_df['Supply Chain Priority'].astype(str).str.lower().map(prio_map).fillna(99)
            sort_cols.append('_prio')
            sort_asc.append(True)
        if 'RMSE' in hr_df.columns:
            sort_cols.append('RMSE')
            sort_asc.append(False)
        if 'Bullwhip_Ratio' in hr_df.columns:
            sort_cols.append('Bullwhip_Ratio')
            sort_asc.append(False)
        if 'SKU' in hr_df.columns:
            sort_cols.append('SKU')
            sort_asc.append(True)

        if sort_cols:
            hr_df = hr_df.sort_values(by=sort_cols, ascending=sort_asc)
        disp_cols = [c for c in hr_cols if c in hr_df.columns and not c.startswith('_')]
        st.dataframe(hr_df[disp_cols], width='stretch', hide_index=True)
        st.caption(f"{len(hr_df):,} SKU(s) flagged High Forecast Risk. Deterministically sorted by Priority, RMSE desc, Bullwhip_Ratio desc, SKU. {source_tag('Master Audit')}")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("📉 Pareto Analysis")
    if HAS_RAW_HISTORY:
        st.markdown(f'<div class="section-caption">Based on 24-month historical Outwards. {source_tag("Historical Raw Data")}</div>', unsafe_allow_html=True)
        hist_totals = raw_history_df[raw_history_df['SKU'].isin(filtered_df['SKU'])].groupby('SKU', as_index=False)['Outwards'].sum()
        hist_totals = hist_totals.rename(columns={'Outwards': 'Total_24M_Outwards'}).sort_values('Total_24M_Outwards', ascending=False).reset_index(drop=True)
        outwards_total = hist_totals['Total_24M_Outwards'].sum()
        if outwards_total and outwards_total > 0:
            hist_totals['Cumulative %'] = hist_totals['Total_24M_Outwards'].cumsum() / outwards_total * 100
            top_n = hist_totals.head(30)
            fig_pareto = make_subplots(specs=[[{"secondary_y": True}]])
            fig_pareto.add_trace(go.Bar(x=top_n['SKU'], y=top_n['Total_24M_Outwards'], name='24M Total Outwards', marker_color='#0EA5E9'), secondary_y=False)
            fig_pareto.add_trace(go.Scatter(x=top_n['SKU'], y=top_n['Cumulative %'], name='Cumulative %', line=dict(color='#F43F5E', width=3)), secondary_y=True)
            fig_pareto.add_hline(y=80, line_dash="dot", line_color="#991B1B", secondary_y=True, annotation_text="80% threshold")
            fig_pareto.update_layout(title="Top 30 SKUs — Pareto (24-Month Historical Outwards)", plot_bgcolor='white', margin=dict(l=20, r=20, t=40, b=80), xaxis_tickangle=-60)
            st.plotly_chart(fig_pareto, width='stretch')
            idx_80 = (hist_totals['Cumulative %'] >= 80.0).idxmax() if (hist_totals['Cumulative %'] >= 80.0).any() else len(hist_totals) - 1
            a_count = int(idx_80) + 1 if (hist_totals['Cumulative %'] >= 80.0).any() else len(hist_totals)
            st.caption(f"Approximately {a_count:,} of {len(hist_totals):,} SKUs reach 80% of cumulative demand.")
    elif 'Annual_Demand' in filtered_df.columns and total_skus > 0:
        st.markdown(f'<div class="section-caption">Based on Annualised Demand from Master Audit. {source_tag("Master Audit")}</div>', unsafe_allow_html=True)
        pareto_df = filtered_df[['SKU', 'Annual_Demand']].dropna().sort_values('Annual_Demand', ascending=False).reset_index(drop=True)
        annual_total = pareto_df['Annual_Demand'].sum()
        if len(pareto_df) > 0 and annual_total and annual_total > 0:
            pareto_df['Cumulative %'] = pareto_df['Annual_Demand'].cumsum() / annual_total * 100
            top_n = pareto_df.head(30)
            fig_pareto = make_subplots(specs=[[{"secondary_y": True}]])
            fig_pareto.add_trace(go.Bar(x=top_n['SKU'], y=top_n['Annual_Demand'], name='Annualised Demand', marker_color='#0EA5E9'), secondary_y=False)
            fig_pareto.add_trace(go.Scatter(x=top_n['SKU'], y=top_n['Cumulative %'], name='Cumulative %', line=dict(color='#F43F5E', width=3)), secondary_y=True)
            fig_pareto.add_hline(y=80, line_dash="dot", line_color="#991B1B", secondary_y=True, annotation_text="80% threshold")
            fig_pareto.update_layout(title="Top 30 SKUs — Pareto (Annualised Demand, Master Audit)", plot_bgcolor='white', margin=dict(l=20, r=20, t=40, b=80), xaxis_tickangle=-60)
            st.plotly_chart(fig_pareto, width='stretch')
            idx_80 = (pareto_df['Cumulative %'] >= 80.0).idxmax() if (pareto_df['Cumulative %'] >= 80.0).any() else len(pareto_df) - 1
            a_count = int(idx_80) + 1 if (pareto_df['Cumulative %'] >= 80.0).any() else len(pareto_df)
            st.caption(f"Approximately {a_count:,} of {len(pareto_df):,} SKUs reach 80% of cumulative demand.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("⚠️ High Inventory Days & Lower Demand Diagnostic")
    st.markdown(f'<div class="section-caption">Flags SKUs with model-implied Inventory Days in the upper quartile and Mean Monthly Demand in the lower quartile of the current filter as candidates for further inventory review. {source_tag("Dashboard Derived")}</div>', unsafe_allow_html=True)
    if 'Inventory_Days' in filtered_df.columns and 'Mean_Monthly_Demand' in filtered_df.columns and total_skus > 3:
        inventory_days_series = safe_series(filtered_df, 'Inventory_Days', numeric=True)
        days_q75 = inventory_days_series.quantile(0.75)
        demand_q25 = mean_monthly_series.quantile(0.25)
        diag_mask = (inventory_days_series >= days_q75) & (mean_monthly_series <= demand_q25)
        diag_cols = [c for c in ['SKU', 'ABC_XYZ_Class', 'Mean_Monthly_Demand', 'Inventory_Days', 'Inventory_Value',
                                 'Forecast_Risk', 'Inventory_Recommendation'] if c in filtered_df.columns]
        diag_df = filtered_df[diag_mask][diag_cols].sort_values('Inventory_Value', ascending=False) if 'Inventory_Value' in diag_cols else filtered_df[diag_mask][diag_cols]
        st.dataframe(diag_df, width='stretch', hide_index=True)
        st.caption(f"{len(diag_df):,} SKU(s) flagged for inventory review. These are dashboard-derived candidates for review, not confirmed obsolescence, inactivity, or trapped capital.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🧭 Dashboard-Derived Review Lenses")
    st.markdown(f'<div class="section-caption">Prioritisation review lenses synthesized from existing Master Audit fields. {source_tag("Dashboard Derived")}</div>', unsafe_allow_html=True)

    pz1, pz2 = st.columns(2)
    with pz1:
        render_status_card("red", "1. High Demand Volume + High Variability", "A-class demand volume items with Z-class variability — prioritize for safety stock and buffer review.", action="Review AZ-segment SKUs in Tab 1 ABC–XYZ matrix")
        render_status_card("amber", "2. High Forecast Error + High Replenishment-Signal Amplification", "Elevated RMSE alongside high Bullwhip Ratio — evaluate model parameters and batching rules.", action="See Signal Amplification vs. Forecast Error scatter above")
    with pz2:
        render_status_card("amber", "3. High Inventory Days & Lower Demand Diagnostic", "Model-implied high days on hand with low demand — candidates for stocking review.", action="See High Inventory Days & Lower Demand Diagnostic above")
        render_status_card("red", "4. High Business Risk + Lower Service Indicators", "Elevated business risk combined with lower service level or fill rate.", action="Cross-reference with Tab 3 Service Level and Fill Rate")

# ==============================================================================
# TAB 5 — FINANCIAL IMPACT, METHODOLOGY & AUDIT
# ==============================================================================
with tab5:
    st.markdown(f'<div class="section-caption">Model-implied financial figures based on project policy assumptions, not observed accounting ledger balances. {source_tag("Master Audit")}</div>', unsafe_allow_html=True)
    st.subheader("💰 Financial KPIs")
    fk1, fk2, fk3, fk4 = st.columns(4)
    render_metric_card(fk1, "Model-Implied Inventory Value", fmt_currency(inv_value_sum), "Portfolio Sum", extra=source_tag("Master Audit"))
    render_metric_card(fk2, "Model-Implied Working Capital", fmt_currency(wc_sum), "Portfolio Sum", extra=source_tag("Master Audit"))
    carrying_sum = safe_series(filtered_df, 'Estimated_Carrying_Cost', numeric=True).sum() if 'Estimated_Carrying_Cost' in filtered_df.columns else np.nan
    stockout_sum = safe_series(filtered_df, 'Stockout_Cost', numeric=True).sum() if 'Stockout_Cost' in filtered_df.columns else np.nan
    render_metric_card(fk3, "Estimated Carrying Cost", fmt_currency(carrying_sum), "Portfolio Sum", extra=source_tag("Master Audit"))
    render_metric_card(fk4, "Estimated Stockout Cost", fmt_currency(stockout_sum), "Portfolio Sum", extra=source_tag("Master Audit"))

    fk5, fk6, fk7, fk8 = st.columns(4)
    turnover_avg = safe_series(filtered_df, 'Inventory_Turnover', numeric=True).mean() if 'Inventory_Turnover' in filtered_df.columns else np.nan
    portfolio_inv_days = (365.0 / turnover_avg) if (not is_missing(turnover_avg) and turnover_avg > 0) else np.nan
    wc_eff_avg = safe_series(filtered_df, 'KPI_Working_Capital_Efficiency', numeric=True).mean() if 'KPI_Working_Capital_Efficiency' in filtered_df.columns else np.nan

    render_metric_card(fk5, "Avg Inventory Turnover", fmt_num(turnover_avg, 2), "turns / yr", extra=source_tag("Master Audit"))
    render_metric_card(fk6, "Portfolio Inventory Days", fmt_num(portfolio_inv_days, 1), "365 / mean(Turnover)", extra=source_tag("Dashboard Derived"))
    render_metric_card(fk7, "Financial Health Score", fmt_num(fin_health_avg, 0, "%"), "Portfolio Average", extra=source_tag("Master Audit"))
    render_metric_card(fk8, "Working Capital Efficiency", fmt_num(wc_eff_avg, 0, "%"), "Portfolio Average", extra=source_tag("Master Audit"))

    st.markdown("<hr>", unsafe_allow_html=True)
    tcol1, tcol2 = st.columns(2)
    with tcol1:
        st.subheader("🗺️ Inventory Allocation by ABC — Demand Volume")
        tree_df = filtered_df[['ABC_Class', 'SKU', 'Inventory_Value']].dropna() if ('Inventory_Value' in filtered_df.columns and 'ABC_Class' in filtered_df.columns) else pd.DataFrame()
        if len(tree_df) > 0:
            fig_tree = px.treemap(
                tree_df, path=[px.Constant("Portfolio"), 'ABC_Class', 'SKU'], values='Inventory_Value',
                color='ABC_Class', color_discrete_map={'A': '#1E3A8A', 'B': '#0EA5E9', 'C': '#94A3B8', '(?)': '#E2E8F0'}
            )
            fig_tree.update_layout(margin=dict(l=10, r=10, t=30, b=10))
            st.plotly_chart(fig_tree, width='stretch')
            st.caption(f"Tile size = Model-Implied Inventory Value (Master Audit). {source_tag('Master Audit')}")
    with tcol2:
        st.subheader("📊 Model-Implied Financial Exposure Components")
        st.markdown(f'<div class="section-caption">Separate model-implied financial dimensions from Master Audit. {source_tag("Master Audit")}</div>', unsafe_allow_html=True)
        if not is_missing(inv_value_sum) and not is_missing(carrying_sum) and not is_missing(stockout_sum):
            fig_fin_comp = go.Figure(go.Bar(
                x=["Inventory Value", "Carrying Cost", "Stockout Exposure"],
                y=[inv_value_sum, carrying_sum, stockout_sum],
                marker_color=['#1E3A8A', '#F59E0B', '#EF4444'],
                text=[fmt_currency(inv_value_sum), fmt_currency(carrying_sum), fmt_currency(stockout_sum)],
                textposition='outside'
            ))
            fig_fin_comp.update_layout(
                title="Model-Implied Financial Components",
                plot_bgcolor='white',
                yaxis_title="Amount (₹)",
                margin=dict(l=20, r=20, t=40, b=20)
            )
            st.plotly_chart(fig_fin_comp, width='stretch')

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🧪 Illustrative Inventory Value Reduction Scenario")
    st.markdown('<div class="illustrative-banner">⚠️ ILLUSTRATIVE SCENARIO — NOT A MODEL OUTPUT, NOT A REALISED SAVING, NOT A FEASIBILITY RESULT. A simple linear what-if exploration.</div>', unsafe_allow_html=True)
    if not is_missing(inv_value_sum):
        reduction_pct = st.slider("Illustrative inventory reduction (%)", min_value=0, max_value=40, value=10, step=1, key="illustrative_reduction_pct")
        hypo_release = inv_value_sum * (reduction_pct / 100.0)
        hypo_residual = inv_value_sum - hypo_release
        sc1, sc2, sc3 = st.columns(3)
        sc1.metric("Current Model-Implied Value", fmt_currency(inv_value_sum))
        sc2.metric(f"Illustrative Inventory Value Reduction ({reduction_pct}%)", fmt_currency(hypo_release))
        sc3.metric("Illustrative Residual Value", fmt_currency(hypo_residual))
        st.caption("This hypothetical calculation does not reflect an engine recommendation, realised cash savings, or operational service feasibility analysis.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🔗 Operational Value Chain — Conceptual Flow Architecture")
    st.markdown('<div class="section-caption">Illustrative process sequence only — link thickness does not represent physical throughput.</div>', unsafe_allow_html=True)

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
        ("Demand Forecast", "Statistical models evaluate candidate methods per SKU.", "Feeds procurement and replenishment planning."),
        ("Procurement", "Order sizing guided by EOQ and Reorder Point logic.", "Provides cadence where commercial constraints permit."),
        ("Production", "Make-to-order or buffer policies may be considered where inventory position, lead time, and demand profiles support them.", "Aligns output with validated forecast and inventory signals."),
        ("Warehouse", "Safety stock and slotting guided by ABC–XYZ classification; JIT or VMI may be considered conditionally where supplier reliability and commercial terms allow.", "Supports pick-face prioritisation."),
        ("Transportation", "Lead time parameters inform safety-stock buffering.", "Affects service-level attainment."),
        ("Customer Service", "Fill Rate and Service Level evaluated against demand.", "Surfaces potential stockout exposure."),
        ("Finance", "Model-implied Inventory Value, Working Capital, and Carrying Cost quantified.", "Provides comparative cost baseline."),
        ("Executive Decisions", "Decision intelligence recommendations synthesized for operational review.", "Closes governance feedback loop."),
    ]
    cards_html = '<div class="matrix-grid-container">'
    for i in range(len(stage_details)):
        title, line1, line2 = stage_details[i]
        cards_html += f'<div class="flow-card" style="border-top: 5px solid {stage_colors[i]};">'
        cards_html += f'<div class="flow-card-title">{i+1}. {title}</div>'
        cards_html += f'<div class="flow-card-text">{line1}</div>'
        cards_html += f'<div class="flow-card-text">{line2}</div></div>'
    cards_html += '</div>'
    st.markdown(cards_html, unsafe_allow_html=True)

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("📚 Methodology & Traceability")
    with st.expander("Methodological Foundations — Enterprise Supply Chain Master Audit", expanded=True):
        methodology_text = (
            "**Data Foundation:** Enterprise Supply Chain Master Audit (`Enterprise_Supply_Chain_Master_Audit.xlsx`, Executive_Summary sheet). "
            "Covers 24 calendar months and 1,401 true SKUs (Grand Total excluded, Missing Outwards = 0, Negative Outwards retained). "
            "All demand statistics are per-SKU aggregates or portfolio distributions.\n\n"
            "**Forecasting Engine:** Model competition runs across four forecasting pathways (3-Month Moving Average, Naive, Seasonal Naive, "
            "and demand-aware AutoARIMA) evaluated on a 20-month training / 4-month chronological validation split. Lowest validation sMAPE "
            "determines the winning model per SKU. Accuracy_Pct is project-defined sMAPE-derived accuracy.\n\n"
            "**Inventory Optimisation:** Safety Stock = 1.645 × RMSE × √(7/30); Lead Time Demand = (Forecast_Next_Month / 30) × 7; "
            "Reorder Point = Lead Time Demand + Safety Stock; EOQ = √(2 × Annual Demand × Ordering Cost / Holding Cost per Unit); "
            "Inventory Turnover = Annual Demand / Average Inventory; Inventory Days = 365 / Inventory Turnover.\n\n"
            "**Portfolio Inventory Days:** Computed at portfolio level as `365 / mean(Inventory_Turnover)` when positive.\n\n"
            "**Segmentation:** ABC classifies demand volume based on cumulative demand contribution; XYZ classifies demand variability based on CV.\n\n"
            "**Diagnostics & Policies:** Bullwhip_Ratio represents model-derived replenishment-signal amplification diagnostic. "
            "JIT, VMI, and MTO are conditional options for management review, not automatic operational actions."
        )
        st.markdown(methodology_text)
        if MISSING_AUTHORITATIVE_FIELDS:
            st.markdown(f"**Field Completeness Check:** {len(MISSING_AUTHORITATIVE_FIELDS)} field(s) absent in source: {', '.join(MISSING_AUTHORITATIVE_FIELDS)}.")
        else:
            st.markdown("**Field Completeness Check:** All 51 authoritative fields used by the dashboard are verified present.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🧮 Inventory Mathematics QA Report")
    with st.expander("Forensic Validation: Safety Stock & Reorder Point across all SKUs", expanded=True):
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
                st.success("✓ Formula integrity confirmed. Every Safety Stock and Reorder Point value in the Master Audit matches the independent calculation within ±1 unit tolerance.")
            else:
                st.warning(f"⚠ Mismatches detected for {qa_summary['ss_material_mismatch']:,} SS and {qa_summary['rop_material_mismatch']:,} ROP values.")

    st.markdown("<hr>", unsafe_allow_html=True)
    st.subheader("🔍 Full Audit Table")
    search_term = st.text_input("🔎 Search SKU", key="audit_search")
    audit_view = filtered_df.copy()
    if search_term:
        audit_view = audit_view[audit_view['SKU'].astype(str).str.contains(search_term, case=False, na=False)]
    st.dataframe(audit_view, width='stretch', hide_index=True, height=350)
    st.caption(f"{len(audit_view):,} of {len(df):,} SKUs displayed directly from {DATA_SOURCE_LABEL}.")

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
            st.download_button(
                "⬇️ Download Filtered Audit Report (.xlsx)", data=excel_buffer,
                file_name="Filtered_Supply_Chain_Audit.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_audit_report"
            )
        except Exception as e:
            st.warning(f"Could not generate export file ({e}).")

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
