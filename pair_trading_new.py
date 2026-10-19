"""
Pair Trading Dashboard - Streamlit App (Enhanced v2 + Live CSV Sync)

Automated pair-trading signal dashboard with dual confirmation
(Google Sheets Live CSV / Excel SIGNAL sheet + Independent analytics).

How to run:
1. pip install streamlit pandas numpy yfinance plotly openpyxl scipy
2. Place "PAIR TRADE -NEW (1).xlsx" in the same folder or paste Google Sheets CSV link in sidebar
3. streamlit run pair_trading_new.py
"""

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
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
st.set_page_config(
page_title="Pair Trade Dashboard",
page_icon="📈",
layout="wide",
initial_sidebar_state="expanded",
)

st.markdown("""

""", unsafe_allow_html=True)


EXCEL_PATH = os.environ.get(
"PAIR_EXCEL_PATH",
os.path.join(os.path.dirname(file) if "file" in dir() else ".", "PAIR TRADE -NEW (1).xlsx")
)
_CANDIDATES = [
EXCEL_PATH,
"/home/workdir/attachments/PAIR TRADE -NEW (1).xlsx",
"PAIR TRADE -NEW (1).xlsx",
os.path.join(os.getcwd(), "PAIR TRADE -NEW (1).xlsx"),
]

DEFAULT_PAIRS = [
{"name": "ICICIBANK-HDFCBANK", "s1": "ICICIBANK.NS", "s2": "HDFCBANK.NS",
"lot1": 700, "lot2": 550, "sector": "Banking", "excel_key": "ICICIBANK"},
{"name": "AXISBANK-ICICIBANK", "s1": "AXISBANK.NS", "s2": "ICICIBANK.NS",
"lot1": 625, "lot2": 700, "sector": "Banking", "excel_key": "AXISBANK"},
{"name": "CIPLA-SUNPHARMA", "s1": "CIPLA.NS", "s2": "SUNPHARMA.NS",
"lot1": 650, "lot2": 350, "sector": "Pharma", "excel_key": "CIPLA"},
{"name": "BPCL-HINDPETRO", "s1": "BPCL.NS", "s2": "HINDPETRO.NS",
"lot1": 1975, "lot2": 2025, "sector": "Oil & Gas", "excel_key": "BPCL"},
{"name": "ACC-AMBUJACEM", "s1": "ACC.NS", "s2": "AMBUJACEM.NS",
"lot1": 300, "lot2": 900, "sector": "Cement", "excel_key": "ACC"},
{"name": "TCS-HCLTECH", "s1": "TCS.NS", "s2": "HCLTECH.NS",
"lot1": 175, "lot2": 350, "sector": "IT", "excel_key": "TCS"},
{"name": "TATASTEEL-JSWSTEEL", "s1": "TATASTEEL.NS", "s2": "JSWSTEEL.NS",
"lot1": 5500, "lot2": 675, "sector": "Metals", "excel_key": "TATASTEEL"},
{"name": "NESTLE-HINDUNILVR", "s1": "NESTLEIND.NS", "s2": "HINDUNILVR.NS",
"lot1": 200, "lot2": 300, "sector": "FMCG", "excel_key": "NESTLEIND"},
{"name": "NIFTY-BANKNIFTY", "s1": "^NSEI", "s2": "^NSEBANK",
"lot1": 25, "lot2": 15, "sector": "Index", "excel_key": "INDEXNIFTY_50"},
{"name": "BRITANNIA-HINDUNILVR", "s1": "BRITANNIA.NS", "s2": "HINDUNILVR.NS",
"lot1": 200, "lot2": 300, "sector": "FMCG", "excel_key": "BRITANNIA"},
{"name": "BAJFINANCE-BAJAJFINSV", "s1": "BAJFINANCE.NS", "s2": "BAJAJFINSV.NS",
"lot1": 100, "lot2": 500, "sector": "Finance", "excel_key": "BAJFINANCE"},
{"name": "HDFCBANK-BANKNIFTY", "s1": "HDFCBANK.NS", "s2": "^NSEBANK",
"lot1": 550, "lot2": 15, "sector": "Banking", "excel_key": "HDFCBANK"},
{"name": "RECLTD-PFC", "s1": "RECLTD.NS", "s2": "PFC.NS",
"lot1": 2000, "lot2": 1300, "sector": "Finance", "excel_key": "RECLTD"},
{"name": "IGL-GAIL", "s1": "IGL.NS", "s2": "GAIL.NS",
"lot1": 1375, "lot2": 4575, "sector": "Oil & Gas", "excel_key": "IGL"},
{"name": "ALKEM-AUROPHARMA", "s1": "ALKEM.NS", "s2": "AUROPHARMA.NS",
"lot1": 100, "lot2": 550, "sector": "Pharma", "excel_key": "ALKEM"},
{"name": "BERGEPAINT-ASIANPAINT", "s1": "BERGEPAINT.NS", "s2": "ASIANPAINT.NS",
"lot1": 1320, "lot2": 200, "sector": "Paints", "excel_key": "BERGEPAINT"},
{"name": "BANKBARODA-CANBK", "s1": "BANKBARODA.NS", "s2": "CANBK.NS",
"lot1": 2925, "lot2": 6750, "sector": "Banking", "excel_key": "BANKBARODA"},
{"name": "COALINDIA-GAIL", "s1": "COALINDIA.NS", "s2": "GAIL.NS",
"lot1": 2925, "lot2": 6750, "sector": "Energy", "excel_key": "COALINDIA"},
]


@st.cache_data(ttl=60, show_spinner=False)
def load_live_signals(csv_url: str = ""):
"""Parse Google Sheets CSV or local Excel SIGNAL sheet."""
ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
excel_map = {}

# Try Google Sheets CSV first
if csv_url.strip():
    try:
        df = pd.read_csv(csv_url)
        # Find rows with valid tickers and lots
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
            return excel_map, f"Loaded from Live Google Sheet CSV @ {ts}"
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
    return excel_map, f"Loaded from Local Excel ({os.path.basename(path)}) @ {ts}"
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

sma20_1 = float(df["price1"].tail(20).mean())
sma20_2 = float(df["price2"].tail(20).mean())

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
last5_z = zscore.tail(5).tolist() if len(zscore) >= 5 else []

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
    "last5_div": last5_div, "last5_dates": last5_dates, "last5_z": last5_z,
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
        reason = f"Z-score +{z:.2f} (ratio high). Mean-reversion → Sell {leg1} / Buy {leg2}"
    else:
        action, leg1, leg2 = "ENTER PAIR", "BUY", "SELL"
        reason = f"Z-score {z:.2f} (ratio low). Mean-reversion → Buy {leg1} / Sell {leg2}"
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


def recommend_cash_neutral_lots(metrics: dict, lot1: int, lot2: int) -> dict:
if not metrics:
return {}
p1, p2 = metrics["price1"], metrics["price2"]
v1, v2 = p1 * lot1, p2 * lot2
best = None
for n1 in range(1, 5):
for n2 in range(1, 5):
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
("30D Min", metrics["ratio_min_30"]),
("30D Max", metrics["ratio_max_30"]),
]
sh1, sh2 = lots_info.get("shares1", 0), lots_info.get("shares2", 0)
if sh1 == 0: return []

rows = []
direction = signal.get("leg1", "")
for label, tgt_r in targets:
    if tgt_r <= 0 or cur_r <= 0: continue
    p1_new = tgt_r * p2
    p2_new = p1 / tgt_r

    if direction == "SELL":
        pnl = (sh1 * (p1 - p1_new) + sh2 * (p2_new - p2)) / 2
    elif direction == "BUY":
        pnl = (sh1 * (p1_new - p1) + sh2 * (p2 - p2_new)) / 2
    else: pnl = 0

    rows.append({
        "Target": label, "Target Ratio": round(tgt_r, 4),
        "Ratio Move %": round(abs(tgt_r - cur_r) / cur_r * 100, 2),
        "Est. P&L (₹)": round(pnl, 0)
    })
return rows

with st.sidebar:
st.markdown("## ⚙️ Controls")
csv_url = st.text_input("Google Sheets Published CSV URL:", value="", help="Paste published CSV link from Google Sheets SIGNAL tab")
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

if st.button("🔄 Refresh Data"):
    st.cache_data.clear()
    st.rerun()


excel_map, signal_status = load_live_signals(csv_url)
ALL_PAIRS = merge_excel_lots(DEFAULT_PAIRS, excel_map)

st.markdown('📈 Live Pair Trading Dashboard', unsafe_allow_html=True)
st.markdown(f'Dual Confirmation Analytics • {datetime.now().strftime("%Y-%m-%d %H:%M:%S")} IST', unsafe_allow_html=True)
st.caption(f"📂 Status: {signal_status}")

pairs_to_scan = [p for p in ALL_PAIRS if p["sector"] in sector_filter]
progress = st.progress(0, text="Fetching market data...")

results = []
for i, pair in enumerate(pairs_to_scan):
progress.progress((i + 1) / len(pairs_to_scan), text=f"Scanning {pair['name']}...")
hist = fetch_pair_history(pair["s1"], pair["s2"], period="1y")
metrics = compute_pair_metrics(hist, lookback=lookback)
signal = generate_signal(metrics, z_entry=z_entry, z_exit=z_exit)
lots_info = recommend_cash_neutral_lots(metrics, pair["lot1"], pair["lot2"]) if metrics else {}

results.append({
    "pair": pair, "metrics": metrics, "signal": signal,
    "lots_info": lots_info, "hist": hist
})


progress.empty()


st.subheader("🎯 Live Signal Board — All Pairs")
rows = []
for r in results:
m, s, p = r["metrics"], r["signal"], r["pair"]
if not m: continue
excel_act = p.get("excel", {}).get("excel_action", "—")

dual = "—"
if "ENTER" in s["action"] and excel_act in ("BUY SIGNAL", "PAIR ENTRY"):
    dual = "✅ DUAL CONFIRM"
elif "ENTER" in s["action"]:
    dual = "⚡ Independent only"
elif excel_act in ("BUY SIGNAL", "PAIR ENTRY"):
    dual = "📋 Sheet only"

rows.append({
    "Pair": p["name"], "Sector": p["sector"],
    "Lot1": p["lot1"], "Lot2": p["lot2"],
    "Price1": round(m["price1"], 2), "Price2": round(m["price2"], 2),
    "Ratio": round(m["ratio"], 4), "Z-Score": round(m["zscore"], 2),
    "Corr30": round(m["corr_30"], 2) if not np.isnan(m["corr_30"]) else "—",
    "Action": s["action"], "Leg1": s["leg1"], "Leg2": s["leg2"],
    "Sheet": excel_act, "Dual": dual
})


if rows:
df_signals = pd.DataFrame(rows)
st.dataframe(df_signals, use_container_width=True, height=500)

Metrics bar

c1, c2, c3, c4 = st.columns(4)
c1.metric("Pairs Scanned", len(rows))
c2.metric("🟢 Active Signals", sum(1 for r in rows if "ENTER" in r["Action"]))
c3.metric("✅ Dual Confirmed", sum(1 for r in rows if "DUAL" in r["Dual"]))
c4.metric("Avg |Z-Score|", round(np.mean([abs(r["Z-Score"]) for r in rows]), 2) if rows else 0)


st.markdown("---")
st.subheader("🔍 Pair Deep-Dive Analysis")
selected = st.selectbox("Select Pair", [r["pair"]["name"] for r in results if r["metrics"]])
selected_res = next(r for r in results if r["pair"]["name"] == selected)

m, s, p = selected_res["metrics"], selected_res["signal"], selected_res["pair"]
lots = selected_res["lots_info"]

KPI Cards

k1, k2, k3, k4 = st.columns(4)
k1.metric("Current Ratio", f"{m['ratio']:.4f}")
k2.metric("Z-Score", f"{m['zscore']:.2f}")
k3.metric("30D Correlation", f"{m['corr_30']:.2f}" if not np.isnan(m["corr_30"]) else "—")
k4.metric("Half-Life", f"{m['half_life']:.1f} days" if not np.isnan(m.get("half_life", np.nan)) else "—")

Ratio Reference Table

st.markdown("#### 📐 Ratio Reference Levels")
st.markdown(f"""

Cash Neutrality Multi-Lot

if lots:
st.markdown("#### 📦 Recommended Cash-Neutral Lot Sizes")
st.write(f"Recommended Multiples: {lots['lots1']} x {p['s1']} ({lots['shares1']} shares) vs {lots['lots2']} x {p['s2']} ({lots['shares2']} shares)")
st.caption(f"Position Value Imbalance: ₹{lots['imbalance']:,.0f} ({lots['imbalance_pct']}%)")

Interactive Charts

if show_charts and HAS_PLOTLY and not selected_res["hist"].empty:
fig = make_subplots(rows=2, cols=1, shared_xaxes=True, subplot_titles=("Ratio & Moving Reference Bands", "Z-Score Trend"))

# Ratio Trace
fig.add_trace(go.Scatter(x=m["ratio_series"].index, y=m["ratio_series"], name="Ratio", line=dict(color="#00E5FF")), row=1, col=1)
fig.add_hline(y=m["ratio_mean"], line_dash="dash", line_color="#FFEA00", row=1, col=1)

# Z-Score Trace
fig.add_trace(go.Scatter(x=m["zscore_series"].index, y=m["zscore_series"], name="Z-Score", line=dict(color="#E040FB")), row=2, col=1)
fig.add_hline(y=z_entry, line_dash="dash", line_color="#FF5252", row=2, col=1)
fig.add_hline(y=-z_entry, line_dash="dash", line_color="#00E676", row=2, col=1)

fig.update_layout(height=500, template="plotly_dark", paper_bgcolor="#0E1117", plot_bgcolor="#1A1F2E")
st.plotly_chart(fig, use_container_width=True)


if auto_refresh:
time.sleep(60)
st.rerun()
