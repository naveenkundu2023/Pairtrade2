"""
Pair Trading Dashboard - Streamlit App (Complete Professional Suite)
===================================================================
1. Live Google Sheets SIGNAL tab visual replica
2. Filtered Live Signal Board (Compact, Buy=Green, Exit=Pink, Hold=White)
3. Independent Engine & Nifty 100 Pair Scanner with Dual Confirmation
4. Enhanced Charts with 30D/60D/90D bands, legend value tables, 5-day divergence values
5. Cross-device persistent URL sync via Query Parameters
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import time
import os
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

# PAGE CONFIG + DARK THEME CSS
st.set_page_config(
    page_title="Pair Trade Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .stApp { background-color: #0E1117; color: #E0E0E0; }
    h1, h2, h3, h4, h5, h6, p, label, span, div { color: #E8EAF0 !important; }
    
    .main-header {
        font-size: 2.1rem; font-weight: 800;
        background: linear-gradient(90deg, #00E5FF, #76FF03, #FFEA00);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        margin-bottom: 0.1rem;
    }
    .sub-header { font-size: 0.9rem; color: #90A4AE !important; margin-bottom: 1.0rem; }
    
    /* Compact Table Styling */
    .compact-table {
        font-size: 0.78rem !important;
        width: 100%;
        border-collapse: collapse;
        margin-bottom: 1rem;
    }
    .compact-table th {
        background-color: #1A2332;
        color: #80D8FF !important;
        padding: 6px 8px;
        text-align: center;
        border: 1px solid #2C384B;
    }
    .compact-table td {
        padding: 5px 8px;
        text-align: center;
        border: 1px solid #2C384B;
    }
    
    /* Signal Row Colors */
    .row-buy { background-color: #12382B !important; color: #E0F2FE !important; }
    .row-exit { background-color: #421820 !important; color: #FFE4E6 !important; }
    .row-watch { background-color: #1E293B !important; color: #F1F5F9 !important; }

    /* Dropdown UI contrast fix */
    div[data-baseweb="select"] * {
        background-color: #1E293B !important;
        color: #FFFFFF !important;
    }
    div[data-baseweb="menu"] * {
        background-color: #0F172A !important;
        color: #F8FAFC !important;
    }

    /* Metric cards */
    [data-testid="stMetric"] {
        background: #1E293B; border-radius: 8px; padding: 10px 14px;
        border-left: 4px solid #00E5FF;
    }
    [data-testid="stMetric"] label { color: #90CAF9 !important; font-size: 0.75rem !important; }
    [data-testid="stMetric"] [data-testid="stMetricValue"] { color: #FFFFFF !important; font-size: 1.2rem !important; font-weight: 700; }

    /* Divergence dot styling */
    .dot-box { display: inline-block; text-align: center; margin: 0 10px; }
    .dot-green { color: #00E676; font-size: 1.4rem; line-height: 1; }
    .dot-red   { color: #FF5252; font-size: 1.4rem; line-height: 1; }
    .dot-gray  { color: #78909C; font-size: 1.4rem; line-height: 1; }
    .dot-val   { font-size: 0.75rem; color: #CBD5E1; margin-top: 2px; }
    .dot-date  { font-size: 0.70rem; color: #94A3B8; }
    
    [data-testid="stSidebar"] { background-color: #12161F; }
    [data-testid="stSidebar"] * { color: #E0E0E0 !important; }
</style>
""", unsafe_allow_html=True)

# CANDIDATE PATHS FOR LOCAL EXCEL
EXCEL_PATH = os.environ.get(
    "PAIR_EXCEL_PATH",
    os.path.join(os.path.dirname(__file__) if "__file__" in dir() else ".", "PAIR TRADE -NEW (1).xlsx")
)
_CANDIDATES = [
    EXCEL_PATH,
    "/home/workdir/attachments/PAIR TRADE -NEW (1).xlsx",
    "PAIR TRADE -NEW (1).xlsx",
    os.path.join(os.getcwd(), "PAIR TRADE -NEW (1).xlsx"),
]

# STANDARD TOP ~40 NSE LIQUID PAIR UNIVERSE
DEFAULT_PAIRS = [
    {"name": "ICICIBANK-HDFCBANK", "s1": "ICICIBANK.NS", "s2": "HDFCBANK.NS", "lot1": 700, "lot2": 550, "sector": "Banking", "excel_key": "ICICIBANK"},
    {"name": "AXISBANK-ICICIBANK", "s1": "AXISBANK.NS", "s2": "ICICIBANK.NS", "lot1": 625, "lot2": 700, "sector": "Banking", "excel_key": "AXISBANK"},
    {"name": "CIPLA-SUNPHARMA", "s1": "CIPLA.NS", "s2": "SUNPHARMA.NS", "lot1": 650, "lot2": 350, "sector": "Pharma", "excel_key": "CIPLA"},
    {"name": "BPCL-HINDPETRO", "s1": "BPCL.NS", "s2": "HINDPETRO.NS", "lot1": 1975, "lot2": 2025, "sector": "Oil & Gas", "excel_key": "BPCL"},
    {"name": "ACC-AMBUJACEM", "s1": "ACC.NS", "s2": "AMBUJACEM.NS", "lot1": 300, "lot2": 900, "sector": "Cement", "excel_key": "ACC"},
    {"name": "TCS-HCLTECH", "s1": "TCS.NS", "s2": "HCLTECH.NS", "lot1": 175, "lot2": 350, "sector": "IT", "excel_key": "TCS"},
    {"name": "INFY-TCS", "s1": "INFY.NS", "s2": "TCS.NS", "lot1": 400, "lot2": 175, "sector": "IT", "excel_key": "INFY"},
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
    {"name": "DRREDDY-LUPIN", "s1": "DRREDDY.NS", "s2": "LUPIN.NS", "lot1": 125, "lot2": 425, "sector": "Pharma", "excel_key": "DRREDDY"},
    {"name": "HEROMOTOCO-TVSMOTOR", "s1": "HEROMOTOCO.NS", "s2": "TVSMOTOR.NS", "lot1": 150, "lot2": 350, "sector": "Auto", "excel_key": "HEROMOTOCO"},
    {"name": "TATAMOTORS-M&M", "s1": "TATAMOTORS.NS", "s2": "M&M.NS", "lot1": 1425, "lot2": 350, "sector": "Auto", "excel_key": "TATAMOTORS"},
]

# CROSS-DEVICE PERSISTENT URL MANAGEMENT
query_params = st.query_params
url_param_csv = query_params.get("csv_url", "")

if "csv_url" not in st.session_state:
    st.session_state["csv_url"] = url_param_csv

# LOAD LIVE SIGNALS & RAW SHEET
@st.cache_data(ttl=60, show_spinner=False)
def load_live_signals_and_raw(csv_url: str = ""):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    excel_map = {}
    raw_df = pd.DataFrame()

    if csv_url.strip():
        try:
            df = pd.read_csv(csv_url)
            raw_df = df.copy()
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
                return excel_map, raw_df, f"Loaded from Live Google Sheet CSV @ {ts}"
        except Exception as e:
            st.sidebar.warning(f"CSV URL Error: {e}")

    # Fallback to local Excel file
    path = None
    for c in _CANDIDATES:
        if os.path.isfile(c):
            path = c
            break
    if not path or not HAS_OPENPYXL:
        return {}, pd.DataFrame(), "No CSV or local Excel found – using built-in lot defaults"

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet_name = None
        for s in wb.sheetnames:
            if s.strip().upper() == "SIGNAL":
                sheet_name = s
                break
        if not sheet_name:
            return {}, pd.DataFrame(), "SIGNAL sheet not found in Excel"

        ws = wb[sheet_name]
        data = ws.values
        cols = next(data)
        raw_df = pd.DataFrame(data, columns=cols)
        
        row = 5
        while row < 80:
            a = ws.cell(row=row, column=1).value
            b = ws.cell(row=row, column=2).value
            if a is not None and b is not None and isinstance(a, (int, float)):
                lot1 = int(a)
                s1 = str(b).replace("NSE:", "").replace("INDEXNSE:", "").replace("MCX:", "").strip()
                a2 = ws.cell(row=row + 1, column=1).value
                b2 = ws.cell(row=row + 1, column=2).value
                if a2 is not None and b2 is not None:
                    lot2 = int(a2)
                    s2 = str(b2).replace("NSE:", "").replace("INDEXNSE:", "").replace("MCX:", "").strip()
                    z_val = ws.cell(row=row, column=12).value
                    ratio_val = ws.cell(row=row, column=15).value
                    excel_buy = ws.cell(row=row, column=30).value

                    excel_action = "NEUTRAL"
                    if excel_buy and "Buy" in str(excel_buy) and "Do Not" not in str(excel_buy):
                        excel_action = "BUY SIGNAL"

                    key = s1.upper()
                    excel_map[key] = {
                        "lot1": lot1, "lot2": lot2, "s1": s1, "s2": s2,
                        "excel_z": float(z_val) if z_val is not None else None,
                        "excel_ratio": float(ratio_val) if ratio_val is not None else None,
                        "excel_action": excel_action, "excel_buy_flag": str(excel_buy or ""),
                    }
            row += 3
        return excel_map, raw_df, f"Loaded from Local Excel ({os.path.basename(path)}) @ {ts}"
    except Exception as e:
        return {}, pd.DataFrame(), f"Excel load error: {e}"

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
    corr_90 = df["price1"].tail(90).corr(df["price2"].tail(90)) if len(df) >= 90 else np.nan

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
    r90 = ratio.tail(90) if len(ratio) >= 90 else ratio

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
        "ratio": cur_ratio, "zscore": cur_z,
        "corr_30": corr_30, "corr_60": corr_60, "corr_90": corr_90,
        "price1": cur_p1, "price2": cur_p2,
        "support1": support1, "resist1": resist1,
        "support2": support2, "resist2": resist2,
        "ratio_max_30": float(r30.max()), "ratio_min_30": float(r30.min()), "ratio_avg_30": float(r30.mean()),
        "ratio_max_60": float(r60.max()), "ratio_min_60": float(r60.min()), "ratio_avg_60": float(r60.mean()),
        "ratio_max_90": float(r90.max()), "ratio_min_90": float(r90.min()), "ratio_avg_90": float(r90.mean()),
        "ratio_mean": float(r60.mean()), "half_life": half_life,
        "ratio_series": ratio, "zscore_series": zscore,
        "last5_div": last5_div, "last5_dates": last5_dates,
        "df": df,
    }

def generate_signal(metrics: dict, z_entry: float = 1.5, z_exit: float = 0.5) -> dict:
    if not metrics:
        return {"action": "NO DATA", "leg1": "—", "leg2": "—", "confidence": 0, "reason": "Insufficient data"}

    z = metrics["zscore"]
    corr = metrics.get("corr_30") or metrics.get("corr_60") or 0
    abs_z = abs(z)
    corr_ok = abs(corr) >= 0.5 if not np.isnan(corr) else False

    if abs_z >= z_entry and corr_ok:
        if z > 0:
            action, leg1, leg2 = "ENTER PAIR", "SELL", "BUY"
            reason = f"Z-score +{z:.2f} (stretched high). Sell leg1 / Buy leg2"
        else:
            action, leg1, leg2 = "ENTER PAIR", "BUY", "SELL"
            reason = f"Z-score {z:.2f} (stretched low). Buy leg1 / Sell leg2"
        confidence = min(95, int(40 + abs_z * 15 + (abs(corr) * 20)))
    elif abs_z <= z_exit:
        action, leg1, leg2 = "EXIT / HOLD", "CLOSE", "CLOSE"
        reason = f"Z-score near mean ({z:.2f}). Take profit zone"
        confidence = 30
    else:
        action, leg1, leg2 = "WATCH", "—", "—"
        reason = f"Z-score {z:.2f} in neutral zone."
        confidence = 20

    return {
        "action": action, "leg1": leg1, "leg2": leg2,
        "confidence": confidence, "reason": reason, "z": z, "corr": corr
    }

# SIDEBAR & PERSISTENCE
with st.sidebar:
    st.markdown("## ⚙️ Controls & Sync")
    csv_input = st.text_input(
        "Google Sheets Published CSV URL:",
        value=st.session_state["csv_url"],
        help="Paste once. Saved in browser URL parameter so it works across mobile & desktop."
    )
    if csv_input != st.session_state["csv_url"]:
        st.session_state["csv_url"] = csv_input
        st.query_params["csv_url"] = csv_input
        st.cache_data.clear()
        st.rerun()

    st.markdown("---")
    z_entry = st.slider("Z-score Entry Threshold", 1.0, 3.0, 1.5, 0.1)
    z_exit = st.slider("Z-score Exit Threshold", 0.1, 1.5, 0.5, 0.1)
    lookback = st.selectbox("Z-score Lookback (days)", [30, 60, 90], index=1)

    sector_filter = st.multiselect(
        "Filter by Sector",
        options=sorted(set(p["sector"] for p in DEFAULT_PAIRS)),
        default=sorted(set(p["sector"] for p in DEFAULT_PAIRS)),
    )

    auto_refresh = st.checkbox("Auto-refresh every 60s", value=False)
    
    if st.button("🔄 Refresh Data Now"):
        st.cache_data.clear()
        st.rerun()

# LOAD DATA
excel_map, raw_sheet_df, signal_status = load_live_signals_and_raw(st.session_state["csv_url"])
ALL_PAIRS = merge_excel_lots(DEFAULT_PAIRS, excel_map)

# HEADER
st.markdown('<p class="main-header">📈 Live Pair Trading Suite</p>', unsafe_allow_html=True)
st.markdown(f'<p class="sub-header">Live Google Sheets Sync • Dual Confirmation • Nifty 100 Scanner • {datetime.now().strftime("%Y-%m-%d %H:%M:%S")} IST</p>', unsafe_allow_html=True)
st.caption(f"📂 Status: {signal_status}")

# ─────────────────────────────────────────────────────────────
# SECTION 1: GOOGLE SHEETS SIGNAL TAB REPLICA
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("📋 Section 1: Live Google Sheets 'SIGNAL' Tab Replica")

if not raw_sheet_df.empty:
    with st.expander("View Full Live Google Sheet SIGNAL Tab Content", expanded=True):
        st.dataframe(raw_sheet_df.iloc[:40], use_container_width=True, height=350)
else:
    st.info("Paste your Google Sheets Published CSV link in the sidebar to render the colorful live SIGNAL tab replica.")

# ─────────────────────────────────────────────────────────────
# SCAN ALL PAIRS FOR SIGNAL BOARD
# ─────────────────────────────────────────────────────────────
pairs_to_scan = [p for p in ALL_PAIRS if p["sector"] in sector_filter]
progress = st.progress(0, text="Fetching market prices and computing signals...")

results = []
for i, pair in enumerate(pairs_to_scan):
    progress.progress((i + 1) / len(pairs_to_scan), text=f"Scanning {pair['name']}...")
    hist = fetch_pair_history(pair["s1"], pair["s2"], period="1y")
    metrics = compute_pair_metrics(hist, lookback=lookback)
    signal = generate_signal(metrics, z_entry=z_entry, z_exit=z_exit)

    results.append({
        "pair": pair, "metrics": metrics, "signal": signal, "hist": hist
    })

progress.empty()

# ─────────────────────────────────────────────────────────────
# SECTION 2: LIVE SIGNAL BOARD (FILTERED & COMPACT)
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("⚡ Section 2: Live Signal Summary Board (Actionable Pairs)")

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
        "Lot1": p["lot1"], "Lot2": p["lot2"],
        "Price1": f"₹{m['price1']:.2f}", "Price2": f"₹{m['price2']:.2f}",
        "Ratio": f"{m['ratio']:.4f}", "Z-Score": f"{m['zscore']:.2f}",
        "Corr30": f"{m['corr_30']:.2f}" if not np.isnan(m["corr_30"]) else "—",
        "Action": s["action"], "Leg1": s["leg1"], "Leg2": s["leg2"],
        "Sheet": excel_act, "Dual": dual,
    })

if rows:
    # Filter for active actionable signals or show all cleanly
    html_rows = []
    for row in rows:
        action_str = str(row["Action"]).upper()
        if "ENTER" in action_str:
            row_class = "row-buy"
        elif "EXIT" in action_str:
            row_class = "row-exit"
        else:
            row_class = "row-watch"

        html_rows.append(f"""
        <tr class="{row_class}">
            <td><b>{row['Pair']}</b></td>
            <td>{row['Sector']}</td>
            <td>{row['Lot1']} / {row['Lot2']}</td>
            <td>{row['Price1']} / {row['Price2']}</td>
            <td><b>{row['Ratio']}</b></td>
            <td><b>{row['Z-Score']}</b></td>
            <td>{row['Corr30']}</td>
            <td><b>{row['Action']}</b></td>
            <td>{row['Leg1']} / {row['Leg2']}</td>
            <td>{row['Sheet']}</td>
            <td><b>{row['Dual']}</b></td>
        </tr>
        """)

    table_html = f"""
    <div style="overflow-x:auto;">
    <table class="compact-table">
        <thead>
            <tr>
                <th>Pair</th><th>Sector</th><th>Lots (1/2)</th><th>Prices (1/2)</th>
                <th>Ratio</th><th>Z-Score</th><th>30D Corr</th><th>Action</th>
                <th>Legs (1/2)</th><th>Sheet Flag</th><th>Dual Status</th>
            </tr>
        </thead>
        <tbody>
            {''.join(html_rows)}
        </tbody>
    </table>
    </div>
    """
    st.markdown(table_html, unsafe_allow_html=True)

c1, c2, c3, c4 = st.columns(4)
c1.metric("Pairs Analyzed", len(rows))
c2.metric("🟢 Active Buy/Enter Signals", sum(1 for r in rows if "ENTER" in r["Action"]))
c3.metric("🔴 Exit / Close Signals", sum(1 for r in rows if "EXIT" in r["Action"]))
c4.metric("✅ Dual Confirmed", sum(1 for r in rows if "DUAL" in r["Dual"]))

# ─────────────────────────────────────────────────────────────
# SECTION 3: PAIR DEEP DIVE & CHARTS
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("🔍 Section 3: Pair Deep-Dive Analysis & Technical Charts")

selected = st.selectbox(
    "Select Pair for Detailed Analytics:",
    [r["pair"]["name"] for r in results if r["metrics"]]
)
selected_res = next(r for r in results if r["pair"]["name"] == selected)

m, s, p = selected_res["metrics"], selected_res["signal"], selected_res["pair"]

k1, k2, k3, k4 = st.columns(4)
k1.metric("Current Ratio", f"{m['ratio']:.4f}")
k2.metric("Z-Score", f"{m['zscore']:.2f}")
k3.metric("30D / 60D Corr", f"{m['corr_30']:.2f} / {m['corr_60']:.2f}")
k4.metric("Half-Life", f"{m['half_life']:.1f} days" if not np.isnan(m.get("half_life", np.nan)) else "—")

# RATIO REFERENCE BAND TABLE
st.markdown("#### 📐 Ratio Reference Bands (30D / 60D / 90D)")
st.markdown(f"""
<div style="overflow-x:auto;">
<table class="compact-table">
    <thead>
        <tr><th>Period</th><th>Min Ratio</th><th>Average Ratio</th><th>Max Ratio</th></tr>
    </thead>
    <tbody>
        <tr><td><b>30-Day</b></td><td style="color:#69F0AE;">{m['ratio_min_30']:.4f}</td><td>{m['ratio_avg_30']:.4f}</td><td style="color:#FF8A80;">{m['ratio_max_30']:.4f}</td></tr>
        <tr><td><b>60-Day</b></td><td style="color:#00B0FF;">{m['ratio_min_60']:.4f}</td><td>{m['ratio_avg_60']:.4f}</td><td style="color:#FF5252;">{m['ratio_max_60']:.4f}</td></tr>
        <tr><td><b>90-Day</b></td><td style="color:#E040FB;">{m['ratio_min_90']:.4f}</td><td>{m['ratio_avg_90']:.4f}</td><td style="color:#FF1744;">{m['ratio_max_90']:.4f}</td></tr>
    </tbody>
</table>
</div>
""", unsafe_allow_html=True)

# LAST 5 DAYS DIVERGENCE WITH VALUES
st.markdown("#### 📅 Last 5 Days Divergence Trend")
if m.get("last5_div") and m.get("last5_dates"):
    dots_html = []
    divs = m["last5_div"]
    dates = m["last5_dates"]
    for idx in range(len(divs)):
        d = dates[idx]
        v = divs[idx]
        # Compare with previous day to see growing/compressing
        if idx > 0:
            growing = abs(v) > abs(divs[idx - 1])
            dot_cls = "dot-green" if growing else "dot-red"
        else:
            dot_cls = "dot-gray"
        
        dots_html.append(f"""
        <div class="dot-box">
            <div class="{dot_cls}">●</div>
            <div class="dot-val">{v*100:+.2f}%</div>
            <div class="dot-date">{d}</div>
        </div>
        """)
    st.markdown(f"<div>{''.join(dots_html)}</div>", unsafe_allow_html=True)

# INTERACTIVE PLOTLY CHARTS WITH VALUES BELOW
if HAS_PLOTLY and not selected_res["hist"].empty:
    hist = selected_res["hist"]
    norm1 = hist["price1"] / hist["price1"].iloc[0] * 100
    norm2 = hist["price2"] / hist["price2"].iloc[0] * 100

    fig = make_subplots(
        rows=3, cols=1, shared_xaxes=True, vertical_spacing=0.07,
        subplot_titles=(
            f"Normalized Underlying Prices ({p['s1']} vs {p['s2']})",
            "Ratio & Reference Bands (30D / 60D / 90D)",
            "Z-Score Trend & Entry Thresholds"
        ),
        row_heights=[0.33, 0.35, 0.32]
    )

    fig.add_trace(go.Scatter(x=hist.index, y=norm1, name=p["s1"], line=dict(color="#00E5FF")), row=1, col=1)
    fig.add_trace(go.Scatter(x=hist.index, y=norm2, name=p["s2"], line=dict(color="#FF5252")), row=1, col=1)

    fig.add_trace(go.Scatter(x=m["ratio_series"].index, y=m["ratio_series"], name="Ratio", line=dict(color="#69F0AE", width=2)), row=2, col=1)
    fig.add_hline(y=m["ratio_mean"], line_dash="dash", line_color="#FFEA00", row=2, col=1)
    fig.add_hline(y=m["ratio_max_60"], line_dash="dot", line_color="#FF5252", row=2, col=1)
    fig.add_hline(y=m["ratio_min_60"], line_dash="dot", line_color="#00B0FF", row=2, col=1)

    fig.add_trace(go.Scatter(x=m["zscore_series"].index, y=m["zscore_series"], name="Z-Score", line=dict(color="#E040FB", width=2)), row=3, col=1)
    fig.add_hline(y=z_entry, line_dash="dash", line_color="#FF5252", row=3, col=1)
    fig.add_hline(y=-z_entry, line_dash="dash", line_color="#00E676", row=3, col=1)

    fig.update_layout(height=780, template="plotly_dark", paper_bgcolor="#0E1117", plot_bgcolor="#1A1F2E")
    st.plotly_chart(fig, use_container_width=True)

    # VALUES LEGEND BELOW CHART FOR FAST READABILITY
    st.markdown("#### 📊 Chart Reference Values Summary")
    st.markdown(f"""
    <div style="overflow-x:auto;">
    <table class="compact-table">
        <tr>
            <th>Current Ratio</th><td><b>{m['ratio']:.4f}</b></td>
            <th>60D Mean Ratio</th><td>{m['ratio_mean']:.4f}</td>
            <th>60D Ratio Range</th><td>{m['ratio_min_60']:.4f} - {m['ratio_max_60']:.4f}</td>
        </tr>
        <tr>
            <th>Current Z-Score</th><td><b>{m['zscore']:.2f}</b></td>
            <th>Upper Entry Z (+1.5)</th><td>+{z_entry:.2f}</td>
            <th>Lower Entry Z (-1.5)</th><td>-{z_entry:.2f}</td>
        </tr>
    </table>
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────
# SECTION 4: EDUCATIONAL DETAILED PAIR TRADING ANALYSIS
# ─────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("📚 Section 4: Quantitative Pair Trading Guide & NSE Strategy")

st.markdown("""
### 🧠 How Pair Trading Works in NSE Markets

Pair trading is a market-neutral statistical arbitrage strategy. Rather than predicting individual market directions, it exploits relative price divergences between two historically correlated equities.

#### Key Quantitative Rules:
1. **Cointegration & Correlation:** Two equities within the same sector (e.g., Banking, Pharma, IT) move together due to common economic drivers. We monitor **30-day, 60-day, and 90-day rolling correlations** (optimal when $\\ge 0.50$).
2. **Ratio Construction:**
   $$\\text{Ratio} = \\frac{\\text{Price of Stock 1}}{\\text{Price of Stock 2}}$$
3. **Z-Score Stretch:**
   $$Z = \\frac{\\text{Ratio} - \\mu_{60}}{\\sigma_{60}}$$
   * When $Z \\ge +1.5$: Stock 1 is relatively overvalued compared to Stock 2. **Sell Stock 1 / Buy Stock 2**.
   * When $Z \\le -1.5$: Stock 1 is relatively undervalued compared to Stock 2. **Buy Stock 1 / Sell Stock 2**.
   * When $|Z| \\le 0.5$: Ratio has returned to its mean. **Close position & take profits**.
4. **Cash-Neutral Sizing:** Multi-lot sizing aligns total contract value ($\text{Lot}_1 \\times \\text{Price}_1 \\approx \\text{Lot}_2 \\times \\text{Price}_2$), insulating your account against broad market swings.
""")

if auto_refresh:
    time.sleep(60)
    st.rerun()
