"""
Pair Trading Dashboard - Streamlit App (Mobile Master v9 - JSONBin Persist)
=========================================================
- Permanent storage via JSONBin (works on all devices)
- Mobile: only green/red status + fixed settings (Z=1.5 / 0.5 / 60)
- Desktop: full controls
"""
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timezone, timedelta
import time
import requests
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

st.set_page_config(
    page_title="Pair Trade Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ==================== JSONBIN CONFIG ====================
JSONBIN_BIN_ID = "6ad88032ac6210605a450ba5"
JSONBIN_MASTER_KEY = "$2a$10$PV6.NiiOEcgZCXk1UdV4GexJxG1LNrSOt3.ewXwxF1nE8Wtu.z3He"
JSONBIN_URL = f"https://api.jsonbin.io/v3/b/{JSONBIN_BIN_ID}"

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
    .sub-header { font-size: 0.9rem; color: #90A4AE !important; margin-bottom: 0.8rem; }
    
    [data-testid="stMetric"] {
        background: #1A1F2E; border-radius: 10px; padding: 10px 14px;
        border-left: 4px solid #00E5FF; box-shadow: 0 2px 8px rgba(0,0,0,0.4);
    }
    [data-testid="stMetric"] label { color: #90CAF9 !important; font-size: 0.75rem !important; }
    [data-testid="stMetric"] [data-testid="stMetricValue"] { color: #FFFFFF !important; font-weight: 700; font-size: 1.2rem !important; }
    
    .js-plotly-plot, .plotly, .plot-container { touch-action: pan-y !important; }
    .stPlotlyChart { overflow: hidden !important; }

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

# ==================== JSONBIN FUNCTIONS ====================
@st.cache_data(ttl=15, show_spinner=False)
def load_csv_url_from_jsonbin():
    """Read the locked CSV URL from JSONBin."""
    try:
        headers = {"X-Master-Key": JSONBIN_MASTER_KEY}
        resp = requests.get(f"{JSONBIN_URL}/latest", headers=headers, timeout=8)
        if resp.status_code == 200:
            data = resp.json()
            record = data.get("record", {})
            url = record.get("csv_url", "").strip()
            return url
    except Exception:
        pass
    return ""

def save_csv_url_to_jsonbin(url: str) -> bool:
    """Save / update the CSV URL in JSONBin."""
    try:
        headers = {
            "Content-Type": "application/json",
            "X-Master-Key": JSONBIN_MASTER_KEY
        }
        payload = {"csv_url": url.strip()}
        resp = requests.put(JSONBIN_URL, json=payload, headers=headers, timeout=8)
        return resp.status_code in (200, 201)
    except Exception:
        return False

# ==================== LOAD LOCKED URL ====================
locked_url = load_csv_url_from_jsonbin()

# ==================== HEADER ====================
st.markdown('<p class="main-header">📈 Live Pair Trading Suite</p>', unsafe_allow_html=True)

utc_now = datetime.now(timezone.utc)
ist_now = utc_now + timedelta(hours=5, minutes=30)
st.markdown(f'<p class="sub-header">Dual confirmation • {ist_now.strftime("%Y-%m-%d %H:%M:%S IST")}</p>', unsafe_allow_html=True)

# ==================== STATUS (visible everywhere) ====================
if locked_url:
    st.success("✅ Google Sheet CSV  →  LOCKED & ACTIVE")
    st.caption(locked_url[:130] + ("..." if len(locked_url) > 130 else ""))
else:
    st.error("❌ Google Sheet CSV  →  NOT LOCKED")
    st.caption("Go to Desktop → paste CSV link → click Save & Lock")

# ==================== DESKTOP CONTROLS ====================
st.markdown('<div class="desktop-controls">', unsafe_allow_html=True)

with st.expander("⚙️ Desktop Controls", expanded=True):
    st.markdown("#### Google Sheets CSV Link")
    url_input = st.text_input("Published CSV URL", value=locked_url)
    
    c1, c2 = st.columns(2)
    with c1:
        if st.button("💾 Save & Lock", type="primary", use_container_width=True):
            clean = url_input.strip()
            if clean:
                success = save_csv_url_to_jsonbin(clean)
                if success:
                    st.cache_data.clear()
                    st.success("Saved permanently! All devices will now see this link.")
                    st.rerun()
                else:
                    st.error("Failed to save to JSONBin. Please try again.")
            else:
                st.warning("Please enter a valid URL")
    with c2:
        if st.button("🗑️ Clear Lock", use_container_width=True):
            save_csv_url_to_jsonbin("")
            st.cache_data.clear()
            st.rerun()
    
    st.info("Once you click **Save & Lock**, the link is stored permanently. Every mobile and desktop will automatically load it.")
    
    st.markdown("#### Parameters (Desktop only)")
    col1, col2, col3 = st.columns(3)
    with col1:
        z_entry = st.slider("Z Entry", 1.0, 3.0, 1.5, 0.1)
    with col2:
        z_exit = st.slider("Z Exit", 0.1, 1.5, 0.5, 0.1)
    with col3:
        lookback = st.selectbox("Lookback", [30, 60, 90], index=1)
    
    sector_filter = st.multiselect(
        "Sectors",
        sorted(set(p["sector"] for p in DEFAULT_PAIRS)),
        default=sorted(set(p["sector"] for p in DEFAULT_PAIRS))
    )
    auto_refresh = st.checkbox("Auto-refresh 60s", value=False)
    show_charts = st.checkbox("Show charts", value=True)
    if st.button("🔄 Refresh Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

st.markdown('</div>', unsafe_allow_html=True)

# Fixed settings (especially for mobile)
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
    if not csv_url.strip():
        return {}, "No CSV link locked"
    try:
        sep = "&" if "?" in csv_url else "?"
        df = pd.read_csv(f"{csv_url}{sep}t={time.time()}")
        for idx in range(0, len(df)-1, 3):
            try:
                row1 = df.iloc[idx]
                row2 = df.iloc[idx+1]
                s1_raw = str(row1.iloc[1]) if len(row1) > 1 else ""
                if not s1_raw or s1_raw == "nan":
                    continue
                s1 = s1_raw.replace("NSE:", "").replace("INDEXNSE:", "").replace("MCX:", "").strip()
                s2 = str(row2.iloc[1]).replace("NSE:", "").replace("INDEXNSE:", "").replace("MCX:", "").strip()
                lot1 = int(float(row1.iloc[0])) if pd.notnull(row1.iloc[0]) else 0
                lot2 = int(float(row2.iloc[0])) if pd.notnull(row2.iloc[0]) else 0
                excel_buy = str(row1.iloc[29]) if len(row1) > 29 and pd.notnull(row1.iloc[29]) else ""
                excel_action = "BUY SIGNAL" if "Buy" in excel_buy and "Do Not" not in excel_buy else "NEUTRAL"
                excel_map[s1.upper()] = {
                    "lot1": lot1, "lot2": lot2, "s1": s1, "s2": s2,
                    "excel_action": excel_action
                }
            except Exception:
                continue
        if excel_map:
            return excel_map, "Live Google Sheet connected"
    except Exception as e:
        return {}, f"CSV Error: {e}"
    return {}, "Sheet loaded but no pairs found"

def merge_excel_lots(pairs, excel_map):
    updated = []
    for p in pairs:
        key = p.get("excel_key", p["name"].split("-")[0]).upper()
        key2 = p["s1"].replace(".NS","").replace("^","").upper()
        info = excel_map.get(key) or excel_map.get(key2)
        new_p = dict(p)
        if info:
            if info.get("lot1", 0) > 0: new_p["lot1"] = info["lot1"]
            if info.get("lot2", 0) > 0: new_p["lot2"] = info["lot2"]
            new_p["excel"] = info
        else:
            new_p["excel"] = {}
        updated.append(new_p)
    return updated

@st.cache_data(ttl=55, show_spinner=False)
def fetch_pair_history(t1, t2, period="1y"):
    if not HAS_YF:
        return pd.DataFrame()
    try:
        h1 = yf.Ticker(t1).history(period=period, auto_adjust=True)
        h2 = yf.Ticker(t2).history(period=period, auto_adjust=True)
        if h1.empty or h2.empty:
            return pd.DataFrame()
        df = pd.DataFrame({"price1": h1["Close"], "price2": h2["Close"]}).dropna()
        df.index = pd.to_datetime(df.index).tz_localize(None)
        return df
    except Exception:
        return pd.DataFrame()

def compute_pair_metrics(df, lookback=60):
    if df.empty or len(df) < 30:
        return {}
    ratio = df["price1"] / df["price2"]
    window = min(lookback, len(ratio)-1)
    zscore = (ratio - ratio.rolling(window).mean()) / ratio.rolling(window).std()
    corr_30 = df["price1"].tail(30).corr(df["price2"].tail(30)) if len(df) >= 30 else np.nan
    corr_60 = df["price1"].tail(60).corr(df["price2"].tail(60)) if len(df) >= 60 else np.nan
    r30, r60 = ratio.tail(30), ratio.tail(60)
    half_life = np.nan
    if HAS_SCIPY and len(ratio) > 30:
        try:
            lag = ratio.shift(1).dropna()
            delta = ratio.diff().dropna()
            lag = lag.loc[delta.index]
            slope = stats.linregress(lag, delta)[0]
            if slope < 0:
                half_life = -np.log(2) / slope
        except Exception:
            pass
    return {
        "ratio": float(ratio.iloc[-1]),
        "zscore": float(zscore.iloc[-1]) if not np.isnan(zscore.iloc[-1]) else 0.0,
        "corr_30": corr_30, "corr_60": corr_60,
        "price1": float(df["price1"].iloc[-1]), "price2": float(df["price2"].iloc[-1]),
        "support1": float(df["price1"].tail(60).min()), "resist1": float(df["price1"].tail(60).max()),
        "support2": float(df["price2"].tail(60).min()), "resist2": float(df["price2"].tail(60).max()),
        "ratio_max_30": float(r30.max()), "ratio_min_30": float(r30.min()),
        "ratio_max_60": float(r60.max()), "ratio_min_60": float(r60.min()),
        "ratio_avg_30": float(r30.mean()), "ratio_avg_60": float(r60.mean()),
        "ratio_mean": float(r60.mean()), "half_life": half_life,
        "ratio_series": ratio, "zscore_series": zscore, "df": df
    }

def generate_signal(metrics, z_entry=1.5, z_exit=0.5):
    if not metrics:
        return {"action": "NO DATA", "leg1": "—", "leg2": "—", "confidence": 0, "reason": "No data"}
    z = metrics["zscore"]
    corr = metrics.get("corr_30") or metrics.get("corr_60") or 0
    abs_z = abs(z)
    corr_ok = abs(corr) >= 0.5 if not np.isnan(corr) else False
    if abs_z >= z_entry and corr_ok:
        if z > 0:
            return {"action": "ENTER PAIR", "leg1": "SELL", "leg2": "BUY",
                    "confidence": min(95, int(40 + abs_z*15 + abs(corr)*20)),
                    "reason": f"Z +{z:.2f} stretched high"}
        else:
            return {"action": "ENTER PAIR", "leg1": "BUY", "leg2": "SELL",
                    "confidence": min(95, int(40 + abs_z*15 + abs(corr)*20)),
                    "reason": f"Z {z:.2f} stretched low"}
    elif abs_z <= z_exit:
        return {"action": "EXIT / HOLD", "leg1": "CLOSE", "leg2": "CLOSE", "confidence": 30,
                "reason": f"Z near mean ({z:.2f})"}
    return {"action": "WATCH", "leg1": "—", "leg2": "—", "confidence": 20,
            "reason": f"Z {z:.2f} neutral"}

def recommend_cash_neutral_lots(metrics, lot1, lot2):
    if not metrics:
        return {}
    p1, p2 = metrics["price1"], metrics["price2"]
    v1, v2 = p1 * lot1, p2 * lot2
    best = None
    for n1 in range(1, 6):
        for n2 in range(1, 6):
            pct = abs(n1*v1 - n2*v2) / max(n1*v1, n2*v2) * 100
            if best is None or pct < best["imbalance_pct"]:
                best = {"lots1": n1, "lots2": n2, "shares1": n1*lot1, "shares2": n2*lot2,
                        "imbalance_pct": round(pct, 2)}
    return best

def profit_potential(metrics, signal, lots):
    if not metrics or not lots:
        return []
    p1, p2, cur_r = metrics["price1"], metrics["price2"], metrics["ratio"]
    targets = [("Mean 60D", metrics["ratio_mean"]), ("30D Avg", metrics["ratio_avg_30"]),
               ("Halfway", (cur_r + metrics["ratio_mean"])/2),
               ("30D Min", metrics["ratio_min_30"]), ("30D Max", metrics["ratio_max_30"])]
    sh1, sh2 = lots.get("shares1", 0), lots.get("shares2", 0)
    if sh1 == 0:
        return []
    rows = []
    direction = signal.get("leg1", "")
    for label, tgt in targets:
        if tgt <= 0 or cur_r <= 0:
            continue
        p1n = tgt * p2
        p2n = p1 / tgt
        if direction == "SELL":
            pnl = (sh1*(p1 - p1n) + sh2*(p2n - p2)) / 2
        elif direction == "BUY":
            pnl = (sh1*(p1n - p1) + sh2*(p2 - p2n)) / 2
        else:
            pnl = 0
        rows.append({"Target": label, "Ratio": round(tgt, 4),
                     "Move%": round(abs(tgt - cur_r)/cur_r*100, 2), "P&L ₹": round(pnl, 0)})
    return rows

def recommend_prices(metrics, signal):
    if not metrics or signal["action"] != "ENTER PAIR":
        return {}
    z = metrics["zscore"]
    p1, p2, ratio, mean_r = metrics["price1"], metrics["price2"], metrics["ratio"], metrics["ratio_mean"]
    stop_r = metrics["ratio_max_60"]*1.02 if z > 0 else metrics["ratio_min_60"]*0.98
    if signal["leg1"] == "SELL":
        return {
            "entry1": round(p1, 2), "target1": round(p1*(mean_r/ratio)*0.98, 2), "stop1": round(p1*(stop_r/ratio)*1.01, 2),
            "entry2": round(p2, 2), "target2": round(p2*(ratio/mean_r)*1.02, 2), "stop2": round(p2*(ratio/stop_r)*0.99, 2)
        }
    return {
        "entry1": round(p1, 2), "target1": round(p1*(mean_r/ratio)*1.02, 2), "stop1": round(p1*(stop_r/ratio)*0.99, 2),
        "entry2": round(p2, 2), "target2": round(p2*(ratio/mean_r)*0.98, 2), "stop2": round(p2*(ratio/stop_r)*1.01, 2)
    }

# ==================== RUN ====================
excel_map, signal_status = load_live_signals(locked_url)
ALL_PAIRS = merge_excel_lots(DEFAULT_PAIRS, excel_map)
st.caption(f"📂 {signal_status}")

pairs_to_scan = [p for p in ALL_PAIRS if p["sector"] in sector_filter]
progress = st.progress(0, text="Scanning pairs...")
results = []
for i, pair in enumerate(pairs_to_scan):
    progress.progress((i+1)/len(pairs_to_scan), text=pair["name"])
    hist = fetch_pair_history(pair["s1"], pair["s2"])
    metrics = compute_pair_metrics(hist, lookback)
    signal = generate_signal(metrics, z_entry, z_exit)
    prices = recommend_prices(metrics, signal)
    lots = recommend_cash_neutral_lots(metrics, pair["lot1"], pair["lot2"]) if metrics else {}
    results.append({"pair": pair, "metrics": metrics, "signal": signal, "prices": prices, "lots_info": lots, "hist": hist})
progress.empty()

# ==================== LIVE SIGNAL BOARD ====================
st.subheader("🎯 Live Signal Board")
rows = []
for r in results:
    m, s, p = r["metrics"], r["signal"], r["pair"]
    if not m:
        continue
    excel_act = p.get("excel", {}).get("excel_action", "—")
    dual = "✅ DUAL CONFIRM" if "ENTER" in s["action"] and excel_act == "BUY SIGNAL" else \
           "⚡ Independent" if "ENTER" in s["action"] else \
           "📋 Sheet Only" if excel_act == "BUY SIGNAL" else "—"
    rows.append({
        "Pair": p["name"], "Sector": p["sector"],
        "Price1": round(m["price1"], 2), "Price2": round(m["price2"], 2),
        "Ratio": round(m["ratio"], 4), "Z": round(m["zscore"], 2),
        "Corr": round(m["corr_30"], 2) if not np.isnan(m["corr_30"]) else "—",
        "Action": s["action"], "Leg1": s["leg1"], "Leg2": s["leg2"],
        "Conf": s["confidence"], "Sheet": excel_act, "Dual": dual
    })

if rows:
    df = pd.DataFrame(rows)
    df.index = range(1, len(df)+1)
    df.index.name = "#"
    def color(row):
        if "ENTER" in str(row["Action"]):
            if "DUAL" in str(row["Dual"]):
                return ["background-color:#00C853;color:#000;font-weight:bold"] * len(row)
            return ["background-color:#2E7D32;color:#FFF;font-weight:bold"] * len(row)
        if "EXIT" in str(row["Action"]):
            return ["background-color:#C62828;color:#FFF"] * len(row)
        return [""] * len(row)
    st.dataframe(df.style.apply(color, axis=1), use_container_width=True, height=500)

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Pairs", len(rows))
c2.metric("🟢 Enter", sum(1 for r in rows if "ENTER" in r["Action"]))
c3.metric("✅ Dual", sum(1 for r in rows if "DUAL" in r["Dual"]))
c4.metric("🟡 Watch", sum(1 for r in rows if "WATCH" in r["Action"]))
c5.metric("Avg |Z|", round(np.mean([abs(r["Z"]) for r in rows]), 2) if rows else 0)

# HIGH CONVICTION
st.markdown("---")
st.subheader("🧠 High-Conviction Entries")
notes = [f"**{r['pair']['name']}** — {r['signal']['action']} (conf {r['signal']['confidence']}%)  \n{r['signal']['reason']}"
         for r in results if r["metrics"] and "ENTER" in r["signal"]["action"]]
if notes:
    for n in notes:
        st.markdown(f"- {n}")
else:
    st.info("No high-conviction entries right now.")

# DEEP DIVE
st.markdown("---")
st.subheader("🔍 Deep-Dive")
sel = st.selectbox("Select Pair", [r["pair"]["name"] for r in results if r["metrics"]])
res = next(r for r in results if r["pair"]["name"] == sel)
m, s, pr, p, lots = res["metrics"], res["signal"], res["prices"], res["pair"], res["lots_info"]

k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Ratio", f"{m['ratio']:.4f}")
k2.metric("Z-Score", f"{m['zscore']:.2f}")
k3.metric("30D Corr", f"{m['corr_30']:.2f}" if not np.isnan(m["corr_30"]) else "—")
k4.metric("60D Corr", f"{m['corr_60']:.2f}" if not np.isnan(m["corr_60"]) else "—")
k5.metric("Half-Life", f"{m['half_life']:.1f}d" if not np.isnan(m.get("half_life", np.nan)) else "—")

bg = "#1B5E20" if "ENTER" in s["action"] else ("#F9A825" if "WATCH" in s["action"] else "#37474F")
st.markdown(f"""
<div style="background:{bg};color:white;padding:14px;border-radius:10px;margin:8px 0;">
<b>{s['action']}</b> — {s['confidence']}% &nbsp;|&nbsp; {p['s1']} → <b>{s['leg1']}</b> &nbsp; {p['s2']} → <b>{s['leg2']}</b><br>
<small>{s['reason']}</small>
</div>
""", unsafe_allow_html=True)

if lots:
    st.write(f"**Lots:** {lots['lots1']}× {p['s1']} vs {lots['lots2']}× {p['s2']} | Imbalance {lots['imbalance_pct']}%")

pnl = profit_potential(m, s, lots)
if pnl:
    st.dataframe(pd.DataFrame(pnl), use_container_width=True, hide_index=True)

if pr:
    st.write(f"**{p['s1']}**: Entry {pr['entry1']} | Target {pr['target1']} | Stop {pr['stop1']}")
    st.write(f"**{p['s2']}**: Entry {pr['entry2']} | Target {pr['target2']} | Stop {pr['stop2']}")

if show_charts and HAS_PLOTLY and not res["hist"].empty:
    hist = res["hist"]
    norm1 = hist["price1"] / hist["price1"].iloc[0] * 100
    norm2 = hist["price2"] / hist["price2"].iloc[0] * 100
    fig = make_subplots(rows=3, cols=1, shared_xaxes=True, vertical_spacing=0.07,
                        subplot_titles=("Normalized", "Ratio", "Z-Score"), row_heights=[0.34, 0.33, 0.33])
    fig.add_trace(go.Scatter(x=hist.index, y=norm1, name=p["s1"], line=dict(color="#00E5FF")), row=1, col=1)
    fig.add_trace(go.Scatter(x=hist.index, y=norm2, name=p["s2"], line=dict(color="#FF5252")), row=1, col=1)
    fig.add_trace(go.Scatter(x=m["ratio_series"].index, y=m["ratio_series"], name="Ratio", line=dict(color="#69F0AE")), row=2, col=1)
    fig.add_hline(y=m["ratio_mean"], line_dash="dash", line_color="#FFEA00", row=2, col=1)
    fig.add_trace(go.Scatter(x=m["zscore_series"].index, y=m["zscore_series"], name="Z", line=dict(color="#E040FB")), row=3, col=1)
    fig.add_hline(y=z_entry, line_dash="dash", line_color="#FF5252", row=3, col=1)
    fig.add_hline(y=-z_entry, line_dash="dash", line_color="#00E676", row=3, col=1)
    fig.update_layout(height=620, template="plotly_dark", paper_bgcolor="#0E1117", plot_bgcolor="#1A1F2E",
                      showlegend=True, margin=dict(l=20, r=10, t=30, b=20), dragmode=False)
    st.plotly_chart(fig, use_container_width=True, config={"scrollZoom": False, "displayModeBar": False, "responsive": True})

if auto_refresh:
    time.sleep(60)
    st.rerun()
