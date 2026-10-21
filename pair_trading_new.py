"""
Pair Trading Dashboard - Streamlit App (Mobile Master v9 - Clean Mobile)
=========================================================
- Permanent lock via cell B75 of SIGNAL sheet
- Mobile: ONLY green/red status shown (no controls)
- Desktop: full controls available
- Fixed settings on mobile: Z=1.5 / 0.5 / 60 days
"""
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timezone, timedelta
import time
import os
import re
import warnings
warnings.filterwarnings("ignore")
try:
    import yfinance as yf
    HAS_YF = True
except ImportError:
    HAS_YF = False
try:
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
    HAS_PLOTLY = True
except ImportError:
    HAS_PLOTLY = False
try:
    from scipy import stats
    HAS_SCIPY = True
except ImportError:
    HAS_SCIPY = False
try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

st.set_page_config(
    page_title="Pair Trade Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    [data-testid="stSidebar"] {display: none !important;}
    [data-testid="collapsedControl"] {display: none !important;}
    section[data-testid="stSidebar"] {display: none !important;}
    #MainMenu {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    .viewerBadge_container__1QS1n {display: none !important;}
    [data-testid="stToolbar"] {display: none !important;}
    [data-testid="stDecoration"] {display: none !important;}

    .stApp { background-color: #0E1117; color: #E0E0E0; }
    
    .main-header {
        font-size: 2.0rem; font-weight: 800;
        background: linear-gradient(90deg, #00E5FF, #76FF03, #FFEA00);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        margin-bottom: 0.1rem;
    }
    .sub-header { font-size: 0.9rem; color: #90A4AE !important; margin-bottom: 1.0rem; }
    
    [data-testid="stMetric"] {
        background: #1A1F2E; border-radius: 10px; padding: 10px 14px;
        border-left: 4px solid #00E5FF; box-shadow: 0 2px 8px rgba(0,0,0,0.4);
    }
    [data-testid="stMetric"] label { color: #90CAF9 !important; font-size: 0.75rem !important; }
    [data-testid="stMetric"] [data-testid="stMetricValue"] { color: #FFFFFF !important; font-weight: 700; font-size: 1.2rem !important; }
    
    .ratio-box {
        background: #1A1F2E; border: 1px solid #37474F; border-radius: 10px;
        padding: 12px 14px; margin: 8px 0 16px 0; overflow-x: auto;
    }
    .ratio-box table { width: 100%; border-collapse: collapse; color: #E0E0E0; }
    .ratio-box th { background: #263238; color: #80D8FF; padding: 6px 8px; text-align: center; font-size: 0.8rem; }
    .ratio-box td { padding: 6px 8px; text-align: center; border-bottom: 1px solid #37474F; font-size: 0.85rem; }
    
    .js-plotly-plot, .plotly, .plot-container { touch-action: pan-y !important; }
    .stPlotlyChart { overflow: hidden !important; }

    /* Hide entire controls section on mobile */
    @media (max-width: 768px) {
        .main-header { font-size: 1.5rem; }
        .sub-header { font-size: 0.8rem; }
        [data-testid="column"] { width: 100% !important; flex: 100% !important; min-width: 100% !important; }
        .desktop-controls { display: none !important; }
    }
</style>
""", unsafe_allow_html=True)

DEFAULT_PAIRS = [
    {"name": "ICICIBANK-HDFCBANK", "s1": "ICICIBANK.NS", "s2": "HDFCBANK.NS", "lot1": 700, "lot2": 550, "sector": "Banking", "excel_key": "ICICIBANK"},
    {"name": "AXISBANK-ICICIBANK", "s1": "AXISBANK.NS", "s2": "ICICIBANK.NS", "lot1": 625, "lot2": 700, "sector": "Banking", "excel_key": "AXISBANK"},
    {"name": "CIPLA-SUNPHARMA", "s1": "CIPLA.NS", "s2": "SUNPHARMA.NS", "lot1": 650, "lot2": 350, "sector": "Pharma", "excel_key": "CIPLA"},
    {"name": "BPCL-HINDPETRO", "s1": "BPCL.NS", "s2": "HINDPETRO.NS", "lot1": 1975, "lot2": 2025, "sector": "Oil & Gas", "excel_key": "BPCL"},
    {"name": "ACC-AMBUJACEM", "s1": "ACC.NS", "s2": "AMBUJACEM.NS", "lot1": 300, "lot2": 900, "sector": "Cement", "excel_key": "ACC"},
    {"name": "TCS-HCLTECH", "s1": "TCS.NS", "s2": "HCLTECH.NS", "lot1": 175, "lot2": 350, "sector": "IT", "excel_key": "TCS"},
    {"name": "TATASTEEL-JSWSTEEL", "s1": "TATASTEEL.NS", "s2": "JSWSTEEL.NS", "lot1": 5500, "lot2": 675, "sector": "Metals", "excel_key": "TATASTEEL"},
    {"name": "NESTLE-HINDUNILVR", "s1": "NESTLEIND.NS", "s2": "HINDUNILVR.NS", "lot1": 200, "lot2": 300, "sector": "FMCG", "excel_key": "NESTLEIND"},
    {"name": "NIFTY-BANKNIFTY", "s1": "^NSEI", "s2": "^NSEBANK", "lot1": 25, "lot2": 15, "sector": "Index", "excel_key": "INDEXNIFTY_50"},
    {"name": "BRITANNIA-HINDUNILVR", "s1": "BRITANNIA.NS", "s2": "HINDUNILVR.NS", "lot1": 200, "lot2": 300, "sector": "FMCG", "excel_key": "BRITANNIA"},
    {"name": "BAJFINANCE-BAJAJFINSV", "s1": "BAJFINANCE.NS", "s2": "BAJAJFINSV.NS", "lot1": 100, "lot2": 500, "sector": "Finance", "excel_key": "BAJFINANCE"},
    {"name": "HDFCBANK-BANKNIFTY", "s1": "HDFCBANK.NS", "s2": "^NSEBANK", "lot1": 550, "lot2": 15, "sector": "Banking", "excel_key": "HDFCBANK"},
    {"name": "RECLTD-PFC", "s1": "RECLTD.NS", "s2": "PFC.NS", "lot1": 2000, "lot2": 1300, "sector": "Finance", "excel_key": "RECLTD"},
    {"name": "IGL-GAIL", "s1": "IGL.NS", "s2": "GAIL.NS", "lot1": 1375, "lot2": 4575, "sector": "Oil & Gas", "excel_key": "IGL"},
    {"name": "ALKEM-AUROPHARMA", "s1": "ALKEM.NS", "s2": "AUROPHARMA.NS", "lot1": 100, "lot2": 550, "sector": "Pharma", "excel_key": "ALKEM"},
    {"name": "BERGEPAINT-ASIANPAINT", "s1": "BERGEPAINT.NS", "s2": "ASIANPAINT.NS", "lot1": 1320, "lot2": 200, "sector": "Paints", "excel_key": "BERGEPAINT"},
    {"name": "BANKBARODA-CANBK", "s1": "BANKBARODA.NS", "s2": "CANBK.NS", "lot1": 2925, "lot2": 6750, "sector": "Banking", "excel_key": "BANKBARODA"},
    {"name": "COALINDIA-GAIL", "s1": "COALINDIA.NS", "s2": "GAIL.NS", "lot1": 2925, "lot2": 6750, "sector": "Energy", "excel_key": "COALINDIA"},
    {"name": "SBIN-PNB", "s1": "SBIN.NS", "s2": "PNB.NS", "lot1": 1500, "lot2": 8000, "sector": "Banking", "excel_key": "SBIN"},
    {"name": "KOTAKBANK-HDFCBANK", "s1": "KOTAKBANK.NS", "s2": "HDFCBANK.NS", "lot1": 400, "lot2": 550, "sector": "Banking", "excel_key": "KOTAKBANK"},
    {"name": "INFY-TCS", "s1": "INFY.NS", "s2": "TCS.NS", "lot1": 400, "lot2": 175, "sector": "IT", "excel_key": "INFY"},
    {"name": "WIPRO-HCLTECH", "s1": "WIPRO.NS", "s2": "HCLTECH.NS", "lot1": 1500, "lot2": 350, "sector": "IT", "excel_key": "WIPRO"},
    {"name": "RELIANCE-ONGC", "s1": "RELIANCE.NS", "s2": "ONGC.NS", "lot1": 250, "lot2": 1925, "sector": "Oil & Gas", "excel_key": "RELIANCE"},
    {"name": "MARUTI-M&M", "s1": "MARUTI.NS", "s2": "M&M.NS", "lot1": 100, "lot2": 350, "sector": "Auto", "excel_key": "MARUTI"},
    {"name": "TITAN-TATASTEEL", "s1": "TITAN.NS", "s2": "TATASTEEL.NS", "lot1": 175, "lot2": 5500, "sector": "Consumer", "excel_key": "TITAN"},
    {"name": "ULTRACEMCO-SHREECEM", "s1": "ULTRACEMCO.NS", "s2": "SHREECEM.NS", "lot1": 100, "lot2": 25, "sector": "Cement", "excel_key": "ULTRACEMCO"},
    {"name": "DRREDDY-CIPLA", "s1": "DRREDDY.NS", "s2": "CIPLA.NS", "lot1": 125, "lot2": 650, "sector": "Pharma", "excel_key": "DRREDDY"},
    {"name": "HDFCLIFE-SBILIFE", "s1": "HDFCLIFE.NS", "s2": "SBILIFE.NS", "lot1": 1100, "lot2": 750, "sector": "Insurance", "excel_key": "HDFCLIFE"},
    {"name": "POWERGRID-NTPC", "s1": "POWERGRID.NS", "s2": "NTPC.NS", "lot1": 2700, "lot2": 1500, "sector": "Power", "excel_key": "POWERGRID"},
    {"name": "ADANIENT-ADANIPORTS", "s1": "ADANIENT.NS", "s2": "ADANIPORTS.NS", "lot1": 300, "lot2": 1250, "sector": "Infrastructure", "excel_key": "ADANIENT"},
]

def extract_sheet_id(csv_url: str) -> str:
    if not csv_url:
        return ""
    m = re.search(r"/d/([a-zA-Z0-9-_]+)", csv_url)
    if m:
        return m.group(1)
    m = re.search(r"key=([a-zA-Z0-9-_]+)", csv_url)
    if m:
        return m.group(1)
    return ""

@st.cache_data(ttl=30, show_spinner=False)
def read_locked_url_from_b75(csv_url: str) -> str:
    if not csv_url.strip():
        return ""
    sheet_id = extract_sheet_id(csv_url)
    if not sheet_id:
        return ""
    candidates = [
        f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet=SIGNAL&range=B75",
        f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv&gid=0&range=B75",
        f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&range=SIGNAL!B75",
    ]
    for url in candidates:
        try:
            df = pd.read_csv(url, header=None)
            if not df.empty:
                val = str(df.iloc[0, 0]).strip()
                if val.startswith("http"):
                    return val
        except Exception:
            continue
    return ""

# ==================== LOAD LOCKED URL ====================
query_params = st.query_params
saved_from_query = query_params.get("csv_url", "")

if "csv_url" not in st.session_state:
    st.session_state["csv_url"] = saved_from_query

possible_source = st.session_state.get("csv_url") or saved_from_query
locked_from_b75 = ""
if possible_source:
    locked_from_b75 = read_locked_url_from_b75(possible_source)

if locked_from_b75:
    st.session_state["csv_url"] = locked_from_b75
    st.query_params["csv_url"] = locked_from_b75
elif saved_from_query:
    st.session_state["csv_url"] = saved_from_query

# ==================== HEADER ====================
st.markdown('<p class="main-header">📈 Live Pair Trading Suite</p>', unsafe_allow_html=True)

utc_now = datetime.now(timezone.utc)
ist_now = utc_now + timedelta(hours=5, minutes=30)
live_clock_ist = ist_now.strftime("%Y-%m-%d %H:%M:%S IST")
st.markdown(f'<p class="sub-header">Dual confirmation • Updated {live_clock_ist}</p>', unsafe_allow_html=True)

# ==================== STATUS BAR (always visible) ====================
active_url = st.session_state.get("csv_url", "")
if active_url:
    st.success("✅ Google Sheet CSV  →  LOCKED & ACTIVE")
else:
    st.error("❌ Google Sheet CSV  →  NOT LOCKED")

# ==================== DESKTOP CONTROLS ONLY ====================
st.markdown('<div class="desktop-controls">', unsafe_allow_html=True)

with st.expander("⚙️ Desktop Controls (hidden on mobile)", expanded=False):
    st.markdown("#### Google Sheets CSV Link")
    url_input = st.text_input("Published CSV URL:", value=st.session_state.get("csv_url", ""))
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("💾 Save & Lock", use_container_width=True, type="primary"):
            clean = url_input.strip()
            st.session_state["csv_url"] = clean
            st.query_params["csv_url"] = clean
            st.cache_data.clear()
            st.rerun()
    with c2:
        if st.button("🗑️ Clear", use_container_width=True):
            st.session_state["csv_url"] = ""
            if "csv_url" in st.query_params:
                del st.query_params["csv_url"]
            st.cache_data.clear()
            st.rerun()
    
    st.info(
        "**One-time setup for all devices:**\n"
        "1. Paste CSV link → click Save & Lock\n"
        "2. Copy the same URL\n"
        "3. Open Google Sheet → SIGNAL sheet → paste into cell **B75**\n"
        "4. Done. All devices will pick it up automatically."
    )
    
    st.markdown("#### Parameters")
    col1, col2, col3 = st.columns(3)
    with col1:
        z_entry = st.slider("Z Entry", 1.0, 3.0, 1.5, 0.1)
    with col2:
        z_exit = st.slider("Z Exit", 0.1, 1.5, 0.5, 0.1)
    with col3:
        lookback = st.selectbox("Lookback", [30, 60, 90], index=1)
    
    sector_filter = st.multiselect(
        "Sector Filter",
        options=sorted(set(p["sector"] for p in DEFAULT_PAIRS)),
        default=sorted(set(p["sector"] for p in DEFAULT_PAIRS)),
    )
    auto_refresh = st.checkbox("Auto-refresh 60s", value=False)
    show_charts = st.checkbox("Show charts", value=True)
    if st.button("🔄 Refresh Now", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

# Force fixed settings (especially important for mobile)
z_entry = 1.5
z_exit = 0.5
lookback = 60
sector_filter = sorted(set(p["sector"] for p in DEFAULT_PAIRS))
auto_refresh = False
show_charts = True

# ==================== DATA FUNCTIONS ====================
@st.cache_data(ttl=1, show_spinner=False)
def load_live_signals(csv_url: str = ""):
    excel_map = {}
    if csv_url.strip():
        try:
            sep = "&" if "?" in csv_url else "?"
            fetch_url = f"{csv_url}{sep}t={time.time()}"
            df = pd.read_csv(fetch_url)
            for idx in range(0, len(df) - 1, 3):
                try:
                    row1 = df.iloc[idx]
                    row2 = df.iloc[idx + 1]
                    s1_raw = str(row1.iloc[1]) if len(row1) > 1 else ""
                    s2_raw = str(row2.iloc[1]) if len(row2) > 1 else ""
                    if not s1_raw or s1_raw == "nan":
                        continue
                    s1 = s1_raw.replace("NSE:", "").replace("INDEXNSE:", "").replace("MCX:", "").strip()
                    s2 = s2_raw.replace("NSE:", "").replace("INDEXNSE:", "").replace("MCX:", "").strip()
                    lot1 = int(float(row1.iloc[0])) if pd.notnull(row1.iloc[0]) else 0
                    lot2 = int(float(row2.iloc[0])) if pd.notnull(row2.iloc[0]) else 0
                    z_val = float(row1.iloc[11]) if len(row1) > 11 and pd.notnull(row1.iloc[11]) else None
                    ratio_val = float(row1.iloc[14]) if len(row1) > 14 and pd.notnull(row1.iloc[14]) else None
                    excel_buy = str(row1.iloc[29]) if len(row1) > 29 and pd.notnull(row1.iloc[29]) else ""
                    excel_action = "NEUTRAL"
                    if "Buy" in excel_buy and "Do Not" not in excel_buy:
                        excel_action = "BUY SIGNAL"
                    key = s1.upper()
                    excel_map[key] = {
                        "lot1": lot1, "lot2": lot2, "s1": s1, "s2": s2,
                        "excel_z": z_val, "excel_ratio": ratio_val,
                        "excel_action": excel_action, "excel_buy_flag": excel_buy,
                    }
                except Exception:
                    continue
            if excel_map:
                return excel_map, "Live Google Sheet CSV connected & active"
        except Exception as e:
            return {}, f"CSV Error: {e}"
    return {}, "No live sheet – Independent only"

def merge_excel_lots(pairs, excel_map):
    updated = []
    for p in pairs:
        key = p.get("excel_key", p["name"].split("-")[0]).upper()
        key2 = p["s1"].replace(".NS", "").replace("^", "").upper()
        info = excel_map.get(key) or excel_map.get(key2)
        new_p = dict(p)
        if info:
            if info["lot1"] > 0: new_p["lot1"] = info["lot1"]
            if info["lot2"] > 0: new_p["lot2"] = info["lot2"]
            new_p["excel"] = info
        else:
            new_p["excel"] = {}
        updated.append(new_p)
    return updated

@st.cache_data(ttl=55, show_spinner=False)
def fetch_pair_history(ticker1: str, ticker2: str, period: str = "1y") -> pd.DataFrame:
    if not HAS_YF:
        return pd.DataFrame()
    try:
        h1 = yf.Ticker(ticker1).history(period=period, auto_adjust=True)
        h2 = yf.Ticker(ticker2).history(period=period, auto_adjust=True)
        if h1.empty or h2.empty:
            return pd.DataFrame()
        df = pd.DataFrame({"price1": h1["Close"], "price2": h2["Close"]}).dropna()
        df.index = pd.to_datetime(df.index).tz_localize(None)
        return df
    except Exception:
        return pd.DataFrame()

def compute_pair_metrics(df: pd.DataFrame, lookback: int = 60) -> dict:
    if df.empty or len(df) < 30:
        return {}
    ratio = df["price1"] / df["price2"]
    window = min(lookback, len(ratio) - 1)
    ratio_mean = ratio.rolling(window).mean()
    ratio_std = ratio.rolling(window).std()
    zscore = (ratio - ratio_mean) / ratio_std
    corr_30 = df["price1"].tail(30).corr(df["price2"].tail(30)) if len(df) >= 30 else np.nan
    corr_60 = df["price1"].tail(60).corr(df["price2"].tail(60)) if len(df) >= 60 else np.nan
    cur_ratio = float(ratio.iloc[-1])
    cur_z = float(zscore.iloc[-1]) if not np.isnan(zscore.iloc[-1]) else 0.0
    cur_p1 = float(df["price1"].iloc[-1])
    cur_p2 = float(df["price2"].iloc[-1])
    recent60 = df.tail(60)
    support1 = float(recent60["price1"].min())
    resist1 = float(recent60["price1"].max())
    support2 = float(recent60["price2"].min())
    resist2 = float(recent60["price2"].max())
    r30 = ratio.tail(30)
    r60 = ratio.tail(60)
    half_life = np.nan
    if HAS_SCIPY and len(ratio) > 30:
        try:
            lag = ratio.shift(1).dropna()
            delta = ratio.diff().dropna()
            lag = lag.loc[delta.index]
            slope, _, _, _, _ = stats.linregress(lag, delta)
            if slope < 0:
                half_life = -np.log(2) / slope
        except Exception:
            pass
    ratio_sma20 = ratio.rolling(20).mean()
    divergence = (ratio - ratio_sma20) / ratio_sma20
    last5_div = divergence.tail(5).tolist() if len(divergence) >= 5 else []
    last5_dates = [d.strftime("%d-%b") for d in divergence.tail(5).index] if len(divergence) >= 5 else []
    return {
        "ratio": cur_ratio, "zscore": cur_z, "corr_30": corr_30, "corr_60": corr_60,
        "price1": cur_p1, "price2": cur_p2,
        "support1": support1, "resist1": resist1, "support2": support2, "resist2": resist2,
        "ratio_max_30": float(r30.max()), "ratio_min_30": float(r30.min()),
        "ratio_max_60": float(r60.max()), "ratio_min_60": float(r60.min()),
        "ratio_avg_30": float(r30.mean()), "ratio_avg_60": float(r60.mean()),
        "ratio_mean": float(r60.mean()), "half_life": half_life,
        "ratio_series": ratio, "zscore_series": zscore,
        "last5_div": last5_div, "last5_dates": last5_dates, "df": df,
    }

def generate_signal(metrics: dict, z_entry: float = 1.5, z_exit: float = 0.5) -> dict:
    if not metrics:
        return {"action": "NO DATA", "leg1": "—", "leg2": "—", "confidence": 0, "reason": "Insufficient data", "deep_reason": ""}
    z = metrics["zscore"]
    corr = metrics.get("corr_30") or metrics.get("corr_60") or 0
    abs_z = abs(z)
    corr_ok = abs(corr) >= 0.5 if not np.isnan(corr) else False
    if abs_z >= z_entry and corr_ok:
        if z > 0:
            action, leg1, leg2 = "ENTER PAIR", "SELL", "BUY"
            reason = f"Z-score +{z:.2f} (ratio stretched high)"
            deep = f"Ratio is {abs_z:.1f} SD above mean. Correlation {corr:.2f}."
        else:
            action, leg1, leg2 = "ENTER PAIR", "BUY", "SELL"
            reason = f"Z-score {z:.2f} (ratio stretched low)"
            deep = f"Ratio is {abs_z:.1f} SD below mean. Correlation {corr:.2f}."
        confidence = min(95, int(40 + abs_z * 15 + (abs(corr) * 20)))
    elif abs_z <= z_exit:
        action, leg1, leg2 = "EXIT / HOLD", "CLOSE", "CLOSE"
        reason = f"Z-score near mean ({z:.2f})"
        deep = "Ratio close to average."
        confidence = 30
    else:
        action, leg1, leg2 = "WATCH", "—", "—"
        reason = f"Z-score {z:.2f} inside neutral band"
        deep = "Waiting for stretch."
        confidence = 20
    return {"action": action, "leg1": leg1, "leg2": leg2, "confidence": confidence, "reason": reason, "deep_reason": deep, "z": z, "corr": corr}

def recommend_cash_neutral_lots(metrics: dict, lot1: int, lot2: int) -> dict:
    if not metrics:
        return {}
    p1, p2 = metrics["price1"], metrics["price2"]
    v1, v2 = p1 * lot1, p2 * lot2
    best = None
    for n1 in range(1, 6):
        for n2 in range(1, 6):
            diff = abs(n1 * v1 - n2 * v2)
            pct = diff / max(n1 * v1, n2 * v2) * 100
            if best is None or pct < best["imbalance_pct"]:
                best = {"lots1": n1, "lots2": n2, "shares1": n1*lot1, "shares2": n2*lot2,
                        "value1": round(n1*v1,0), "value2": round(n2*v2,0),
                        "imbalance": round(n1*v1 - n2*v2,0), "imbalance_pct": round(pct,2)}
    return best

def profit_potential(metrics: dict, signal: dict, lots_info: dict) -> list:
    if not metrics or not lots_info:
        return []
    p1, p2 = metrics["price1"], metrics["price2"]
    cur_r = metrics["ratio"]
    targets = [("Mean 60D", metrics["ratio_mean"]), ("30D Avg", metrics["ratio_avg_30"]),
               ("Halfway", (cur_r + metrics["ratio_mean"])/2),
               ("30D Min", metrics["ratio_min_30"]), ("30D Max", metrics["ratio_max_30"])]
    sh1, sh2 = lots_info.get("shares1",0), lots_info.get("shares2",0)
    if sh1 == 0: return []
    rows = []
    direction = signal.get("leg1","")
    for label, tgt_r in targets:
        if tgt_r <= 0 or cur_r <= 0: continue
        p1_new = tgt_r * p2
        p2_new = p1 / tgt_r
        if direction == "SELL":
            pnl = (sh1*(p1 - p1_new) + sh2*(p2_new - p2))/2
        elif direction == "BUY":
            pnl = (sh1*(p1_new - p1) + sh2*(p2 - p2_new))/2
        else:
            pnl = 0
        rows.append({"Target": label, "Target Ratio": round(tgt_r,4),
                     "Move %": round(abs(tgt_r-cur_r)/cur_r*100,2), "Est. P&L ₹": round(pnl,0)})
    return rows

def recommend_prices(metrics: dict, signal: dict) -> dict:
    if not metrics or signal["action"] not in ("ENTER PAIR",):
        return {}
    z = metrics["zscore"]
    p1, p2 = metrics["price1"], metrics["price2"]
    ratio = metrics["ratio"]
    mean_r = metrics["ratio_mean"]
    stop_ratio = metrics["ratio_max_60"]*1.02 if z > 0 else metrics["ratio_min_60"]*0.98
    if signal["leg1"] == "SELL":
        entry1, target1 = p1, p1*(mean_r/ratio)*0.98
        stop1 = p1*(stop_ratio/ratio)*1.01
        entry2, target2 = p2, p2*(ratio/mean_r)*1.02
        stop2 = p2*(ratio/stop_ratio)*0.99
    else:
        entry1, target1 = p1, p1*(mean_r/ratio)*1.02
        stop1 = p1*(stop_ratio/ratio)*0.99
        entry2, target2 = p2, p2*(ratio/mean_r)*0.98
        stop2 = p2*(ratio/stop_ratio)*1.01
    return {"entry1": round(entry1,2), "target1": round(target1,2), "stop1": round(stop1,2),
            "entry2": round(entry2,2), "target2": round(target2,2), "stop2": round(stop2,2),
            "support1": round(metrics["support1"],2), "resist1": round(metrics["resist1"],2),
            "support2": round(metrics["support2"],2), "resist2": round(metrics["resist2"],2)}

def historical_pair_stats(df: pd.DataFrame, z_entry: float = 1.5, lookback: int = 60) -> dict:
    if df.empty or len(df) < lookback + 20:
        return {"trades": 0, "wins": 0, "win_rate": None, "avg_pnl_pct": None, "details": []}
    ratio = df["price1"] / df["price2"]
    rmean = ratio.rolling(lookback).mean()
    rstd = ratio.rolling(lookback).std()
    z = (ratio - rmean) / rstd
    trades = []
    i = lookback
    while i < len(z) - 5:
        zi = z.iloc[i]
        if np.isnan(zi):
            i += 1
            continue
        if abs(zi) >= z_entry:
            entry_ratio = ratio.iloc[i]
            entry_z = zi
            exit_i = None
            for j in range(i+1, min(i+31, len(z))):
                if abs(z.iloc[j]) <= 0.5:
                    exit_i = j
                    break
            if exit_i is None:
                exit_i = min(i+30, len(z)-1)
            exit_ratio = ratio.iloc[exit_i]
            pnl_pct = (entry_ratio - exit_ratio)/entry_ratio*100 if entry_z > 0 else (exit_ratio - entry_ratio)/entry_ratio*100
            trades.append({"entry_date": z.index[i].strftime("%Y-%m-%d"),
                           "exit_date": z.index[exit_i].strftime("%Y-%m-%d"),
                           "entry_z": round(entry_z,2), "pnl_pct": round(pnl_pct,2), "win": pnl_pct > 0})
            i = exit_i + 1
        else:
            i += 1
    if not trades:
        return {"trades": 0, "wins": 0, "win_rate": None, "avg_pnl_pct": None, "details": []}
    wins = sum(1 for t in trades if t["win"])
    return {"trades": len(trades), "wins": wins,
            "win_rate": round(wins/len(trades)*100,1),
            "avg_pnl_pct": round(np.mean([t["pnl_pct"] for t in trades]),2),
            "details": trades[-12:]}

# ==================== LOAD & SCAN ====================
excel_map, signal_status = load_live_signals(st.session_state.get("csv_url", ""))
ALL_PAIRS = merge_excel_lots(DEFAULT_PAIRS, excel_map)
st.caption(f"📂 {signal_status}")

pairs_to_scan = [p for p in ALL_PAIRS if p["sector"] in sector_filter]
progress = st.progress(0, text="Scanning pairs...")
results = []
for i, pair in enumerate(pairs_to_scan):
    progress.progress((i+1)/len(pairs_to_scan), text=f"Scanning {pair['name']}...")
    hist = fetch_pair_history(pair["s1"], pair["s2"])
    metrics = compute_pair_metrics(hist, lookback=lookback)
    signal = generate_signal(metrics, z_entry=z_entry, z_exit=z_exit)
    prices = recommend_prices(metrics, signal)
    lots_info = recommend_cash_neutral_lots(metrics, pair["lot1"], pair["lot2"]) if metrics else {}
    hist_stats = historical_pair_stats(hist, z_entry=z_entry, lookback=lookback) if not hist.empty else {}
    results.append({"pair": pair, "metrics": metrics, "signal": signal, "prices": prices,
                    "lots_info": lots_info, "hist_stats": hist_stats, "hist": hist})
progress.empty()

# ==================== LIVE SIGNAL BOARD ====================
st.subheader("🎯 Live Signal Board")
rows = []
for r in results:
    m, s, p = r["metrics"], r["signal"], r["pair"]
    if not m: continue
    excel_act = p.get("excel", {}).get("excel_action", "—")
    dual = "—"
    if "ENTER" in s["action"] and excel_act in ("BUY SIGNAL", "PAIR ENTRY"):
        dual = "✅ DUAL CONFIRM"
    elif "ENTER" in s["action"]:
        dual = "⚡ Independent"
    elif excel_act in ("BUY SIGNAL", "PAIR ENTRY"):
        dual = "📋 Sheet Only"
    rows.append({
        "Pair": p["name"], "Sector": p["sector"],
        "Price1": round(m["price1"],2), "Price2": round(m["price2"],2),
        "Ratio": round(m["ratio"],4), "Z-Score": round(m["zscore"],2),
        "Corr30": round(m["corr_30"],2) if not np.isnan(m["corr_30"]) else "—",
        "Action": s["action"], "Leg1": s["leg1"], "Leg2": s["leg2"],
        "Conf%": s["confidence"], "Sheet": excel_act, "Dual": dual,
    })

if rows:
    df_signals = pd.DataFrame(rows)
    df_signals.index = range(1, len(df_signals)+1)
    df_signals.index.name = "#"
    def color_rows(row):
        act = str(row["Action"]).upper()
        dual = str(row["Dual"])
        if "ENTER" in act:
            if "DUAL" in dual:
                return ["background-color: #00C853; color: #000; font-weight: bold"] * len(row)
            return ["background-color: #2E7D32; color: #FFF; font-weight: bold"] * len(row)
        elif "EXIT" in act:
            return ["background-color: #C62828; color: #FFF"] * len(row)
        return [""] * len(row)
    st.dataframe(df_signals.style.apply(color_rows, axis=1), use_container_width=True, height=520)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Pairs", len(rows))
c2.metric("🟢 Enter", sum(1 for r in rows if "ENTER" in r["Action"]))
c3.metric("✅ Dual", sum(1 for r in rows if "DUAL" in r["Dual"]))
c4.metric("🟡 Watch", sum(1 for r in rows if "WATCH" in r["Action"]))
c5.metric("Avg |Z|", round(np.mean([abs(r["Z-Score"]) for r in rows]),2) if rows else 0)

# INDEPENDENT
st.markdown("---")
st.subheader("🧠 High-Conviction Entries")
notes = []
for r in results:
    if not r["metrics"]: continue
    s = r["signal"]
    if "ENTER" in s["action"]:
        tag = " ✅ DUAL" if r["pair"].get("excel",{}).get("excel_action") in ("BUY SIGNAL","PAIR ENTRY") else ""
        notes.append(f"**{r['pair']['name']}** — {s['action']} (conf {s['confidence']}%){tag}  \n{s['reason']}")
if notes:
    for n in notes:
        st.markdown(f"- {n}")
else:
    st.info("No high-conviction entries right now.")

# DEEP DIVE
st.markdown("---")
st.subheader("🔍 Pair Deep-Dive")
selected = st.selectbox("Select Pair", [r["pair"]["name"] for r in results if r["metrics"]])
selected_res = next(r for r in results if r["pair"]["name"] == selected)
m, s, pr, p = selected_res["metrics"], selected_res["signal"], selected_res["prices"], selected_res["pair"]
lots = selected_res["lots_info"]

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Ratio", f"{m['ratio']:.4f}")
k2.metric("Z-Score", f"{m['zscore']:.2f}")
k3.metric("30D Corr", f"{m['corr_30']:.2f}" if not np.isnan(m["corr_30"]) else "—")
k4.metric("60D Corr", f"{m['corr_60']:.2f}" if not np.isnan(m["corr_60"]) else "—")
k5.metric("Half-Life", f"{m['half_life']:.1f}d" if not np.isnan(m.get("half_life", np.nan)) else "—")

action_bg = "#1B5E20" if "ENTER" in s["action"] else ("#F9A825" if "WATCH" in s["action"] else "#37474F")
st.markdown(f"""
<div style="background:{action_bg};color:white;padding:14px;border-radius:10px;margin:10px 0;">
    <b>{s['action']}</b> — Conf {s['confidence']}%<br>
    {p['s1']} → <b>{s['leg1']}</b> &nbsp;|&nbsp; {p['s2']} → <b>{s['leg2']}</b><br>
    <small>{s['reason']}</small>
</div>
""", unsafe_allow_html=True)

if lots:
    st.write(f"**Lots:** {lots['lots1']}×{p['s1']} ({lots['shares1']}) vs {lots['lots2']}×{p['s2']} ({lots['shares2']}) | Imbalance {lots['imbalance_pct']}%")

pnl_rows = profit_potential(m, s, lots)
if pnl_rows:
    st.dataframe(pd.DataFrame(pnl_rows), use_container_width=True, hide_index=True)

if pr:
    st.write(f"**{p['s1']}**: Entry {pr['entry1']} | Target {pr['target1']} | Stop {pr['stop1']}")
    st.write(f"**{p['s2']}**: Entry {pr['entry2']} | Target {pr['target2']} | Stop {pr['stop2']}")

if show_charts and HAS_PLOTLY and not selected_res["hist"].empty:
    hist = selected_res["hist"]
    norm1 = hist["price1"] / hist["price1"].iloc[0] * 100
    norm2 = hist["price2"] / hist["price2"].iloc[0] * 100
    fig = make_subplots(rows=3, cols=1, shared_xaxes=True, vertical_spacing=0.08,
                        subplot_titles=("Normalized Prices", "Ratio", "Z-Score"),
                        row_heights=[0.35,0.35,0.30])
    fig.add_trace(go.Scatter(x=hist.index, y=norm1, name=p["s1"], line=dict(color="#00E5FF")), row=1, col=1)
    fig.add_trace(go.Scatter(x=hist.index, y=norm2, name=p["s2"], line=dict(color="#FF5252")), row=1, col=1)
    fig.add_trace(go.Scatter(x=m["ratio_series"].index, y=m["ratio_series"], name="Ratio", line=dict(color="#69F0AE")), row=2, col=1)
    fig.add_hline(y=m["ratio_mean"], line_dash="dash", line_color="#FFEA00", row=2, col=1)
    fig.add_trace(go.Scatter(x=m["zscore_series"].index, y=m["zscore_series"], name="Z", line=dict(color="#E040FB")), row=3, col=1)
    fig.add_hline(y=z_entry, line_dash="dash", line_color="#FF5252", row=3, col=1)
    fig.add_hline(y=-z_entry, line_dash="dash", line_color="#00E676", row=3, col=1)
    fig.update_layout(height=650, template="plotly_dark", paper_bgcolor="#0E1117", plot_bgcolor="#1A1F2E",
                      showlegend=True, margin=dict(l=20,r=10,t=30,b=20), dragmode=False)
    st.plotly_chart(fig, use_container_width=True, config={"scrollZoom":False,"displayModeBar":False,"responsive":True})

if auto_refresh:
    time.sleep(60)
    st.rerun()
