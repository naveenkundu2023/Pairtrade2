"""
Pair Trading Dashboard - Streamlit App (Fully Dynamic & Clean v6)
=================================================================
- Real-time un-cached live clock evaluation on every page load
- Cache-busting live Google Sheets CSV fetch (forces fresh data pull on every refresh)
- Clean neutral rows (no navy blue background on watch/hold rows)
- High-contrast signal row coloring (Green for Buy, Red for Exit)
- Mobile-optimized touch-scroll Plotly charts
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

# PAGE CONFIG + DARK THEME
st.set_page_config(
    page_title="Pair Trade Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .stApp { background-color: #0E1117; color: #E0E0E0; }
    
    .main-header {
        font-size: 2.2rem; font-weight: 800;
        background: linear-gradient(90deg, #00E5FF, #76FF03, #FFEA00);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        margin-bottom: 0.1rem;
    }
    .sub-header { font-size: 0.95rem; color: #90A4AE !important; margin-bottom: 1.0rem; }
    
    /* Metric Cards */
    [data-testid="stMetric"] {
        background: #1A1F2E; border-radius: 10px; padding: 12px 16px;
        border-left: 4px solid #00E5FF; box-shadow: 0 2px 8px rgba(0,0,0,0.4);
    }
    [data-testid="stMetric"] label { color: #90CAF9 !important; font-size: 0.8rem !important; }
    [data-testid="stMetric"] [data-testid="stMetricValue"] { color: #FFFFFF !important; font-weight: 700; }
    
    /* Ratio Box */
    .ratio-box {
        background: #1A1F2E; border: 1px solid #37474F; border-radius: 10px;
        padding: 14px 18px; margin: 8px 0 16px 0;
    }
    .ratio-box table { width: 100%; border-collapse: collapse; color: #E0E0E0; }
    .ratio-box th { background: #263238; color: #80D8FF; padding: 8px 12px; text-align: center; font-size: 0.85rem; }
    .ratio-box td { padding: 7px 12px; text-align: center; border-bottom: 1px solid #37474F; font-size: 0.9rem; }
    
    /* Divergence dots */
    .dot-green { color: #00E676; font-size: 1.3rem; }
    .dot-red   { color: #FF5252; font-size: 1.3rem; }
    .dot-gray  { color: #78909C; font-size: 1.3rem; }
    
    [data-testid="stSidebar"] { background-color: #12161F; }
    [data-testid="stSidebar"] label { color: #E0E0E0 !important; font-weight: 600; }
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

# TOP NSE PAIR UNIVERSE
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
]

# PERSISTENT CROSS-DEVICE LINK SAVING
query_params = st.query_params
saved_csv_url = query_params.get("csv_url", "")

if "csv_url" not in st.session_state:
    st.session_state["csv_url"] = saved_csv_url

# SIDEBAR CONTROL
with st.sidebar:
    st.markdown("## ⚙️ Controls & Sync")
    
    url_input = st.text_input(
        "Google Sheets Published CSV URL:",
        value=st.session_state["csv_url"],
        help="Paste published CSV link from Google Sheets SIGNAL tab."
    )
    
    if st.button("💾 Save & Lock Link", use_container_width=True):
        st.session_state["csv_url"] = url_input
        st.query_params["csv_url"] = url_input
        st.cache_data.clear()
        st.success("Link saved permanently!")
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
    show_charts = st.checkbox("Show detailed charts", value=True)

    if st.button("🔄 Refresh Data Now", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

@st.cache_data(ttl=1, show_spinner=False)
def load_live_signals(csv_url: str = ""):
    excel_map = {}

    if csv_url.strip():
        try:
            # Append cache-busting timestamp to force fresh fetch from Google Sheets on every refresh
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
                return excel_map, "Live Google Sheet CSV connected & freshly fetched"
        except Exception as e:
            st.sidebar.warning(f"CSV URL Error: {e}")

    # Fallback to local Excel file
    path = None
    for c in _CANDIDATES:
        if os.path.isfile(c):
            path = c
            break
    if not path or not HAS_OPENPYXL:
        return {}, "No CSV or local Excel found – using built-in lot defaults"

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        sheet_name = None
        for s in wb.sheetnames:
            if s.strip().upper() == "SIGNAL":
                sheet_name = s
                break
        if not sheet_name:
            return {}, "SIGNAL sheet not found in Excel"

        ws = wb[sheet_name]
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
        return excel_map, f"Local Excel ({os.path.basename(path)}) loaded"
    except Exception as e:
        return {}, f"Excel load error: {e}"

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
        "ratio": cur_ratio, "zscore": cur_z,
        "corr_30": corr_30, "corr_60": corr_60,
        "price1": cur_p1, "price2": cur_p2,
        "support1": support1, "resist1": resist1,
        "support2": support2, "resist2": resist2,
        "ratio_max_30": float(r30.max()), "ratio_min_30": float(r30.min()),
        "ratio_max_60": float(r60.max()), "ratio_min_60": float(r60.min()),
        "ratio_avg_30": float(r30.mean()), "ratio_avg_60": float(r60.mean()),
        "ratio_mean": float(r60.mean()), "half_life": half_life,
        "ratio_series": ratio, "zscore_series": zscore,
        "last5_div": last5_div, "last5_dates": last5_dates,
        "df": df,
    }

def generate_signal(metrics: dict, z_entry: float = 1.5, z_exit: float = 0.5) -> dict:
    if not metrics:
        return {"action": "NO DATA", "leg1": "—", "leg2": "—", "confidence": 0, "reason": "Insufficient data", "deep_reason": ""}

    z = metrics["zscore"]
    corr = metrics.get("corr_30") or metrics.get("corr_60") or 0
    abs_z = abs(z)
    corr_ok = abs(corr) >= 0.5 if not np.isnan(corr) else False

    deep = ""
    if abs_z >= z_entry and corr_ok:
        if z > 0:
            action, leg1, leg2 = "ENTER PAIR", "SELL", "BUY"
            reason = f"Z-score +{z:.2f} (ratio stretched high). Expect mean-reversion → Sell leg1 / Buy leg2"
            deep = f"The price ratio is {abs_z:.1f} standard deviations above its 60-day mean. 30-day correlation is {corr:.2f}, indicating the pair still moves together. Selling the richer leg and buying the cheaper leg captures the reversion."
        else:
            action, leg1, leg2 = "ENTER PAIR", "BUY", "SELL"
            reason = f"Z-score {z:.2f} (ratio stretched low). Expect mean-reversion → Buy leg1 / Sell leg2"
            deep = f"The price ratio is {abs_z:.1f} standard deviations below its rolling mean. Mean-reversion probability is elevated. Correlation ({corr:.2f}) supports pair quality. Buying the cheaper leg and shorting the richer leg is the classic setup."
        confidence = min(95, int(40 + abs_z * 15 + (abs(corr) * 20)))
    elif abs_z <= z_exit:
        action, leg1, leg2 = "EXIT / HOLD", "CLOSE", "CLOSE"
        reason = f"Z-score near mean ({z:.2f}). No edge / take-profit zone"
        deep = "Ratio is close to its average; edge has largely been realized or never appeared."
        confidence = 30
    else:
        action, leg1, leg2 = "WATCH", "—", "—"
        reason = f"Z-score {z:.2f} inside neutral band. Waiting for stretch"
        deep = "Wait for |Z| to exceed the entry threshold before committing capital."
        confidence = 20

    return {
        "action": action, "leg1": leg1, "leg2": leg2,
        "confidence": confidence, "reason": reason, "deep_reason": deep,
        "z": z, "corr": corr
    }

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
                best = {
                    "lots1": n1, "lots2": n2,
                    "shares1": n1 * lot1, "shares2": n2 * lot2,
                    "value1": round(n1 * v1, 0), "value2": round(n2 * v2, 0),
                    "imbalance": round(n1 * v1 - n2 * v2, 0),
                    "imbalance_pct": round(pct, 2),
                }
    best["single_value1"] = round(v1, 0)
    best["single_value2"] = round(v2, 0)
    best["single_imbalance"] = round(v1 - v2, 0)
    best["single_imbalance_pct"] = round(abs(v1 - v2) / max(v1, v2) * 100, 2)
    return best

def profit_potential(metrics: dict, signal: dict, lots_info: dict) -> list:
    if not metrics or not lots_info:
        return []
    p1, p2 = metrics["price1"], metrics["price2"]
    cur_r = metrics["ratio"]
    targets = [
        ("Mean (60D avg)", metrics["ratio_mean"]),
        ("30D Avg", metrics["ratio_avg_30"]),
        ("Halfway to mean", (cur_r + metrics["ratio_mean"]) / 2),
        ("30D Min", metrics["ratio_min_30"]),
        ("30D Max", metrics["ratio_max_30"]),
    ]
    sh1, sh2 = lots_info.get("shares1", 0), lots_info.get("shares2", 0)
    if sh1 == 0: return []

    rows = []
    direction = signal.get("leg1", "")
    for label, tgt_r in targets:
        if tgt_r <= 0 or cur_r <= 0: continue
        p1_new_a = tgt_r * p2
        p2_new_b = p1 / tgt_r

        if direction == "SELL":
            pnl = (sh1 * (p1 - p1_new_a) + sh2 * (p2_new_b - p2)) / 2
        elif direction == "BUY":
            pnl = (sh1 * (p1_new_a - p1) + sh2 * (p2 - p2_new_b)) / 2
        else: pnl = 0

        rows.append({
            "Target": label, "Target Ratio": round(tgt_r, 4),
            "Ratio Move %": round(abs(tgt_r - cur_r) / cur_r * 100, 2),
            "Est. P&L (₹)": round(pnl, 0)
        })
    return rows

def recommend_prices(metrics: dict, signal: dict) -> dict:
    if not metrics or signal["action"] not in ("ENTER PAIR",):
        return {}
    z = metrics["zscore"]
    p1, p2 = metrics["price1"], metrics["price2"]
    ratio = metrics["ratio"]
    mean_r = metrics["ratio_mean"]
    target_ratio = mean_r
    stop_ratio = metrics["ratio_max_60"] * 1.02 if z > 0 else metrics["ratio_min_60"] * 0.98

    if signal["leg1"] == "SELL":
        entry1, target1 = p1, p1 * (target_ratio / ratio) * 0.98
        stop1 = p1 * (stop_ratio / ratio) * 1.01
        entry2, target2 = p2, p2 * (ratio / target_ratio) * 1.02
        stop2 = p2 * (ratio / stop_ratio) * 0.99
    else:
        entry1, target1 = p1, p1 * (target_ratio / ratio) * 1.02
        stop1 = p1 * (stop_ratio / ratio) * 0.99
        entry2, target2 = p2, p2 * (ratio / target_ratio) * 0.98
        stop2 = p2 * (ratio / stop_ratio) * 1.01

    return {
        "entry1": round(entry1, 2), "target1": round(target1, 2), "stop1": round(stop1, 2),
        "entry2": round(entry2, 2), "target2": round(target2, 2), "stop2": round(stop2, 2),
        "support1": round(metrics["support1"], 2), "resist1": round(metrics["resist1"], 2),
        "support2": round(metrics["support2"], 2), "resist2": round(metrics["resist2"], 2),
    }

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
            for j in range(i + 1, min(i + 31, len(z))):
                if abs(z.iloc[j]) <= 0.5:
                    exit_i = j
                    break
            if exit_i is None:
                exit_i = min(i + 30, len(z) - 1)
            exit_ratio = ratio.iloc[exit_i]
            pnl_pct = (entry_ratio - exit_ratio) / entry_ratio * 100 if entry_z > 0 else (exit_ratio - entry_ratio) / entry_ratio * 100
            trades.append({
                "entry_date": z.index[i].strftime("%Y-%m-%d"),
                "exit_date": z.index[exit_i].strftime("%Y-%m-%d"),
                "entry_z": round(entry_z, 2),
                "pnl_pct": round(pnl_pct, 2),
                "win": pnl_pct > 0,
            })
            i = exit_i + 1
        else:
            i += 1

    if not trades:
        return {"trades": 0, "wins": 0, "win_rate": None, "avg_pnl_pct": None, "details": []}
    wins = sum(1 for t in trades if t["win"])
    return {
        "trades": len(trades), "wins": wins,
        "win_rate": round(wins / len(trades) * 100, 1),
        "avg_pnl_pct": round(np.mean([t["pnl_pct"] for t in trades]), 2),
        "details": trades[-12:],
    }

# LOAD LIVE SIGNALS
excel_map, signal_status = load_live_signals(st.session_state["csv_url"])
ALL_PAIRS = merge_excel_lots(DEFAULT_PAIRS, excel_map)

# UN-CACHED LIVE CLOCK EVALUATION (ALWAYS SHOWS CURRENT REAL-TIME SECONDS)
live_clock_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S IST")

# HEADER
st.markdown('<p class="main-header">📈 Live Pair Trading Suite</p>', unsafe_allow_html=True)
st.markdown(f'<p class="sub-header">Dual confirmation (Google Sheets / Excel + Independent) • Updated {live_clock_now}</p>', unsafe_allow_html=True)
st.caption(f"📂 {signal_status}")

# SCAN PAIRS
pairs_to_scan = [p for p in ALL_PAIRS if p["sector"] in sector_filter]
progress = st.progress(0, text="Fetching live market data...")

results = []
for i, pair in enumerate(pairs_to_scan):
    progress.progress((i + 1) / len(pairs_to_scan), text=f"Scanning {pair['name']}...")
    hist = fetch_pair_history(pair["s1"], pair["s2"], period="1y")
    metrics = compute_pair_metrics(hist, lookback=lookback)
    signal = generate_signal(metrics, z_entry=z_entry, z_exit=z_exit)
    prices = recommend_prices(metrics, signal)
    lots_info = recommend_cash_neutral_lots(metrics, pair["lot1"], pair["lot2"]) if metrics else {}
    hist_stats = historical_pair_stats(hist, z_entry=z_entry, lookback=lookback) if not hist.empty else {}

    results.append({
        "pair": pair, "metrics": metrics, "signal": signal, "prices": prices,
        "lots_info": lots_info, "hist_stats": hist_stats, "hist": hist
    })

progress.empty()

# 🎯 LIVE SIGNAL BOARD
st.subheader("🎯 Live Signal Board — All Pairs")
rows = []
for r in results:
    m, s, p = r["metrics"], r["signal"], r["pair"]
    if not m: continue
    excel_info = p.get("excel", {})
    excel_act = excel_info.get("excel_action", "—")

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
        "Price1": round(m["price1"], 2), "Price2": round(m["price2"], 2),
        "Ratio": round(m["ratio"], 4), "Z-Score": round(m["zscore"], 2),
        "Corr30": round(m["corr_30"], 2) if not np.isnan(m["corr_30"]) else "—",
        "Action": s["action"], "Leg1": s["leg1"], "Leg2": s["leg2"],
        "Conf%": s["confidence"], "Sheet Signal": excel_act, "Dual Status": dual,
        "HalfLife": round(m["half_life"], 1) if not np.isnan(m.get("half_life", np.nan)) else "—",
    })

if rows:
    df_signals = pd.DataFrame(rows)

    def color_rows(row):
        act = str(row["Action"]).upper()
        dual = str(row["Dual Status"])
        if "ENTER" in act:
            if "DUAL" in dual:
                # Vibrant bright green for dual conviction
                return ["background-color: #00C853; color: #000000; font-weight: bold"] * len(row)
            # Solid green for single signal
            return ["background-color: #2E7D32; color: #FFFFFF; font-weight: bold"] * len(row)
        elif "EXIT" in act:
            if "DUAL" in dual:
                # Vibrant crimson red for dual exit
                return ["background-color: #D50000; color: #FFFFFF; font-weight: bold"] * len(row)
            # Medium red for single exit
            return ["background-color: #C62828; color: #FFFFFF"] * len(row)
        else:
            # NO COLOR FOR NEUTRAL / WATCH ROWS (Transparent / Default Table Background)
            return [""] * len(row)

    styled_df = df_signals.style.apply(color_rows, axis=1)
    st.dataframe(styled_df, use_container_width=True, height=520)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Pairs Scanned", len(rows))
c2.metric("🟢 Enter Signals", sum(1 for r in rows if "ENTER" in r["Action"]))
c3.metric("✅ Dual Confirm", sum(1 for r in rows if "DUAL" in r["Dual Status"]))
c4.metric("🟡 Watchlist", sum(1 for r in rows if "WATCH" in r["Action"]))
c5.metric("Avg |Z|", round(np.mean([abs(r["Z-Score"]) for r in rows]), 2) if rows else 0)

# 🧠 INDEPENDENT ANALYTICS - HIGH CONVICTION
st.markdown("---")
st.subheader("🧠 Independent Analytics — High-Conviction Entries")

notes = []
for r in results:
    if not r["metrics"]: continue
    s = r["signal"]
    m = r["metrics"]
    p = r["pair"]
    if "ENTER" in s["action"]:
        excel_tag = ""
        if p.get("excel", {}).get("excel_action") in ("BUY SIGNAL", "PAIR ENTRY"):
            excel_tag = " ✅ **DUAL CONFIRM with Excel**"
        notes.append(
            f"**{p['name']}** — {s['action']} (conf {s['confidence']}%). "
            f"Z={m['zscore']:.2f}, Corr30={m['corr_30']:.2f}. {s['reason']}{excel_tag}\n\n"
            f"&nbsp;&nbsp;&nbsp;*Deep rational:* {s.get('deep_reason', '')}"
        )

if notes:
    for n in notes:
        st.markdown(f"- {n}")
else:
    st.info("No high-conviction entry signals at the moment. All pair Z-scores are inside normal bands.")

# 🔍 PAIR DEEP-DIVE
st.markdown("---")
st.subheader("🔍 Pair Deep-Dive Analysis")

selected = st.selectbox("Select Pair for Detailed Analysis", [r["pair"]["name"] for r in results if r["metrics"]])
selected_res = next(r for r in results if r["pair"]["name"] == selected)

m, s, pr, p = selected_res["metrics"], selected_res["signal"], selected_res["prices"], selected_res["pair"]
lots = selected_res["lots_info"]
hist_stats = selected_res["hist_stats"]

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Current Ratio", f"{m['ratio']:.4f}")
k2.metric("Z-Score", f"{m['zscore']:.2f}", delta="Stretched" if abs(m["zscore"]) > z_entry else "Normal")
k3.metric("30D Corr", f"{m['corr_30']:.2f}" if not np.isnan(m["corr_30"]) else "—")
k4.metric("60D Corr", f"{m['corr_60']:.2f}" if not np.isnan(m["corr_60"]) else "—")
k5.metric("Half-Life", f"{m['half_life']:.1f} d" if not np.isnan(m.get("half_life", np.nan)) else "—")

# SIGNAL BOX BANNER
action_bg = "#1B5E20" if "ENTER" in s["action"] else ("#F9A825" if "WATCH" in s["action"] else "#37474F")
st.markdown(
    f"""
    <div style="background:{action_bg};color:white;padding:18px;border-radius:12px;margin:12px 0;">
        <h3 style="margin:0;color:white;">{s['action']} — Confidence {s['confidence']}%</h3>
        <p style="margin:8px 0 0 0;font-size:1.1rem;">
            <b>{p['s1']}</b> → <b>{s['leg1']}</b> &nbsp;&nbsp;|&nbsp;&nbsp; <b>{p['s2']}</b> → <b>{s['leg2']}</b>
        </p>
        <p style="margin:8px 0 0 0;opacity:0.95;">{s['reason']}</p>
    </div>
    """,
    unsafe_allow_html=True
)

if s.get("deep_reason"):
    st.markdown("#### 🧠 Independent Deep Rational")
    st.info(s["deep_reason"])

st.markdown("#### 📐 Ratio Reference Levels")
st.markdown(f"""
<div class="ratio-box">
<table>
    <tr><th>Metric</th><th>Current</th><th>30-Day</th><th>60-Day</th></tr>
    <tr><td><b>Ratio</b></td><td style="color:#00E5FF;font-weight:700;">{m['ratio']:.4f}</td><td>—</td><td>—</td></tr>
    <tr><td>Average</td><td>—</td><td>{m['ratio_avg_30']:.4f}</td><td>{m['ratio_avg_60']:.4f}</td></tr>
    <tr><td>Maximum</td><td>—</td><td style="color:#FF8A80;">{m['ratio_max_30']:.4f}</td><td style="color:#FF8A80;">{m['ratio_max_60']:.4f}</td></tr>
    <tr><td>Minimum</td><td>—</td><td style="color:#69F0AE;">{m['ratio_min_30']:.4f}</td><td style="color:#69F0AE;">{m['ratio_min_60']:.4f}</td></tr>
</table>
</div>
""", unsafe_allow_html=True)

if lots:
    st.markdown("#### 📦 Recommended Cash-Neutral Lot Sizes")
    c1, c2 = st.columns(2)
    with c1:
        st.write(f"**Multi-Lot Recommendation:** Use **{lots['lots1']} x lot1** ({lots['shares1']} shares) of {p['s1']} vs **{lots['lots2']} x lot2** ({lots['shares2']} shares) of {p['s2']}")
    with c2:
        st.write(f"Position Imbalance: ₹{lots['imbalance']:,.0f} ({lots['imbalance_pct']}%)")

pnl_rows = profit_potential(m, s, lots)
if pnl_rows:
    st.markdown("#### 💰 Profit Potential Grid")
    st.dataframe(pd.DataFrame(pnl_rows), use_container_width=True, hide_index=True)

if pr:
    st.markdown("#### 🎯 Suggested Entry / Target / Stop Levels")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"**{p['s1']}** (Lot {p['lot1']})")
        st.write(f"Entry: {pr['entry1']} | Target: {pr['target1']} | Stop: {pr['stop1']} | Support: {pr['support1']} | Resistance: {pr['resist1']}")
    with c2:
        st.markdown(f"**{p['s2']}** (Lot {p['lot2']})")
        st.write(f"Entry: {pr['entry2']} | Target: {pr['target2']} | Stop: {pr['stop2']} | Support: {pr['support2']} | Resistance: {pr['resist2']}")

st.markdown("#### 📅 Last 5 Trading Days — Divergence Trend")
if m.get("last5_div") and m.get("last5_dates"):
    dots_html = []
    divs = m["last5_div"]
    dates = m["last5_dates"]
    for idx in range(len(divs)):
        d = dates[idx]
        v = divs[idx]
        dot_cls = "dot-green" if v > 0.005 else ("dot-red" if v < -0.005 else "dot-gray")
        dots_html.append(f'<span class="{dot_cls}">●</span> {d} ({v*100:+.2f}%)')
    st.markdown(" &nbsp;&nbsp;&nbsp;&nbsp; ".join(dots_html), unsafe_allow_html=True)

if show_charts and HAS_PLOTLY and not selected_res["hist"].empty:
    hist = selected_res["hist"]
    norm1 = hist["price1"] / hist["price1"].iloc[0] * 100
    norm2 = hist["price2"] / hist["price2"].iloc[0] * 100

    fig = make_subplots(
        rows=3, cols=1, shared_xaxes=True, vertical_spacing=0.07,
        subplot_titles=(
            f"Normalized Prices — {p['s1']} vs {p['s2']} (Base=100)",
            f"Ratio (Price1/Price2) | Mean={m['ratio_mean']:.4f} | 60D Max={m['ratio_max_60']:.4f} | 60D Min={m['ratio_min_60']:.4f}",
            f"Z-Score Trend | Current Z={m['zscore']:.2f} | Entry Band=±{z_entry:.1f}"
        ),
        row_heights=[0.35, 0.35, 0.30]
    )

    fig.add_trace(go.Scatter(x=hist.index, y=norm1, name=p["s1"], line=dict(color="#00E5FF")), row=1, col=1)
    fig.add_trace(go.Scatter(x=hist.index, y=norm2, name=p["s2"], line=dict(color="#FF5252")), row=1, col=1)

    fig.add_trace(go.Scatter(x=m["ratio_series"].index, y=m["ratio_series"], name="Ratio", line=dict(color="#69F0AE")), row=2, col=1)
    fig.add_hline(y=m["ratio_mean"], line_dash="dash", line_color="#FFEA00", row=2, col=1)
    fig.add_hline(y=m["ratio_max_60"], line_dash="dot", line_color="#FF8A80", row=2, col=1)
    fig.add_hline(y=m["ratio_min_60"], line_dash="dot", line_color="#80D8FF", row=2, col=1)

    fig.add_trace(go.Scatter(x=m["zscore_series"].index, y=m["zscore_series"], name="Z-Score", line=dict(color="#E040FB")), row=3, col=1)
    fig.add_hline(y=z_entry, line_dash="dash", line_color="#FF5252", row=3, col=1)
    fig.add_hline(y=-z_entry, line_dash="dash", line_color="#00E676", row=3, col=1)

    fig.update_layout(
        height=750,
        template="plotly_dark",
        paper_bgcolor="#0E1117",
        plot_bgcolor="#1A1F2E",
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02)
    )
    st.plotly_chart(fig, use_container_width=True, config={'scrollZoom': False, 'displayModeBar': False})

# 📊 HISTORICAL PAIR PERFORMANCE
st.markdown("---")
st.subheader("📊 Historical Pair Performance (Simulated — Last ~12 Months)")
st.caption("Simple back-test: enter when |Z| ≥ entry threshold, exit when |Z| ≤ 0.5 or after 30 trading days.")

hist_rows = []
for r in results:
    hs = r.get("hist_stats") or {}
    if hs.get("trades", 0) > 0:
        hist_rows.append({
            "Pair": r["pair"]["name"], "Opportunities": hs["trades"],
            "Wins": hs["wins"], "Win Rate %": hs["win_rate"], "Avg P&L %": hs["avg_pnl_pct"]
        })

if hist_rows:
    st.dataframe(pd.DataFrame(hist_rows), use_container_width=True, hide_index=True)
    if hist_stats.get("details"):
        st.markdown(f"**Recent simulated trades for {selected}**")
        st.dataframe(pd.DataFrame(hist_stats["details"]), use_container_width=True, hide_index=True)

if auto_refresh:
    time.sleep(60)
    st.rerun()
