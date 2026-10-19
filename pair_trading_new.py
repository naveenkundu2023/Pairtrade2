import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
import yfinance as yf
from datetime import datetime, timedelta

# Set page layout to wide mode
st.set_page_config(
    page_title="Pair Trading Live Dashboard",
    page_icon="📈",
    layout="wide"
)

# -----------------------------------------------------------------------------
# HELPERS & DATA LOADING
# -----------------------------------------------------------------------------
@st.cache_data(ttl=60)  # Auto refresh cache every 60 seconds
def load_excel_or_google_sheet(file_path_or_url):
    """
    Loads the Excel file or Google Sheet CSV export link.
    """
    try:
        if file_path_or_url.endswith('.xlsx') or file_path_or_url.endswith('.xls'):
            df_signal = pd.read_excel(file_path_or_url, sheet_name='SIGNAL')
        else:
            df_signal = pd.read_csv(file_path_or_url)
        return df_signal
    except Exception as e:
        st.error(f"Error loading file: {e}")
        return pd.DataFrame()

# Sample Pair Definitions for Independent Verification & Charting
DEFAULT_PAIRS = [
    {"name": "ICICIBANK - HDFCBANK", "s1": "ICICIBANK.NS", "s2": "HDFCBANK.NS", "lot1": 700, "lot2": 550},
    {"name": "AXISBANK - ICICIBANK", "s1": "AXISBANK.NS", "s2": "ICICIBANK.NS", "lot1": 625, "lot2": 700},
    {"name": "CIPLA - SUNPHARMA", "s1": "CIPLA.NS", "s2": "SUNPHARMA.NS", "lot1": 650, "lot2": 350},
    {"name": "BPCL - HINDPETRO", "s1": "BPCL.NS", "s2": "HINDPETRO.NS", "lot1": 1975, "lot2": 2025},
    {"name": "ACC - AMBUJACEM", "s1": "ACC.NS", "s2": "AMBUJACEM.NS", "lot1": 300, "lot2": 900},
    {"name": "TCS - HCLTECH", "s1": "TCS.NS", "s2": "HCLTECH.NS", "lot1": 175, "lot2": 350},
    {"name": "RECLTD - PFC", "s1": "RECLTD.NS", "s2": "PFC.NS", "lot1": 2000, "lot2": 1300},
    {"name": "ALKEM - AUROPHARMA", "s1": "ALKEM.NS", "s2": "AUROPHARMA.NS", "lot1": 100, "lot2": 550},
    {"name": "BANKBARODA - CANBK", "s1": "BANKBARODA.NS", "s2": "CANBK.NS", "lot1": 2925, "lot2": 6750},
]

@st.cache_data(ttl=300)
def fetch_historical_pair_data(s1_ticker, s2_ticker, days=90):
    end = datetime.now()
    start = end - timedelta(days=days + 30)
    data = yf.download([s1_ticker, s2_ticker], start=start, end=end)['Close']
    data = data.dropna()
    if s1_ticker in data and s2_ticker in data:
        df = pd.DataFrame()
        df['Price1'] = data[s1_ticker]
        df['Price2'] = data[s2_ticker]
        df['Ratio'] = df['Price1'] / df['Price2']
        
        # Calculate Rolling Z-Score (20-day SMA)
        df['Mean'] = df['Ratio'].rolling(window=20).mean()
        df['Std'] = df['Ratio'].rolling(window=20).std()
        df['ZScore'] = (df['Ratio'] - df['Mean']) / df['Std']
        return df.tail(days)
    return pd.DataFrame()

# -----------------------------------------------------------------------------
# SIDEBAR / CONTROLS
# -----------------------------------------------------------------------------
st.sidebar.title("⚙️ Dashboard Controls")

file_input = st.sidebar.text_input(
    "Google Sheets Export URL / Local File Path:",
    value="PAIR TRADE -NEW (1).xlsx"
)

z_entry_threshold = st.sidebar.slider("Z-Score Entry Threshold", 1.0, 3.0, 2.0, 0.1)
z_exit_threshold = st.sidebar.slider("Z-Score Exit Threshold", 0.0, 1.0, 0.2, 0.1)

st.title("📊 Live Pair Trading Dashboard")

# -----------------------------------------------------------------------------
# SECTION 1 & 2 & 3: LIVE SIGNAL SHEET REFLECTION (WITH HORIZONTAL SLIDER)
# -----------------------------------------------------------------------------
st.header("1. Live Signal Sheet Reflection")

df_signals = load_excel_or_google_sheet(file_input)

if not df_signals.empty:
    st.write("Below is the live reflection from your Excel/Google Sheet file:")
    # Horizontal slider enabled automatically via Streamlit dataframe display inside a scrollable container
    st.dataframe(
        df_signals,
        use_container_width=True,
        height=350
    )
else:
    st.warning("No data found or sheet 'SIGNAL' could not be loaded. Please verify the file path.")

st.markdown("---")

# -----------------------------------------------------------------------------
# SECTION 4 & 5: INDEPENDENT PAIR TRADE LOGIC & DUAL CONFIRMATION
# -----------------------------------------------------------------------------
st.header("2. Independent Logic Verification & Dual Confirmation")

verified_results = []

for pair in DEFAULT_PAIRS:
    hist = fetch_historical_pair_data(pair['s1'], pair['s2'], days=90)
    if not hist.empty:
        curr_ratio = hist['Ratio'].iloc[-1]
        curr_z = hist['ZScore'].iloc[-1]
        p1 = hist['Price1'].iloc[-1]
        p2 = hist['Price2'].iloc[-1]
        
        # Independent Signal Logic
        ind_signal = "NEUTRAL"
        if curr_z >= z_entry_threshold:
            ind_signal = "SELL Leg1 / BUY Leg2"
        elif curr_z <= -z_entry_threshold:
            ind_signal = "BUY Leg1 / SELL Leg2"
            
        # Exit Signal Logic
        exit_signal = "HOLD / NO POSITION"
        if abs(curr_z) <= z_exit_threshold:
            exit_signal = "EXIT / TAKE PROFIT"

        # Check for Dual Confirmation with Excel Sheet
        excel_matched = False
        sheet_signal = "N/A"
        if not df_signals.empty and 'PAIR' in df_signals.columns:
            matching_rows = df_signals[df_signals['PAIR'].astype(str).str.contains(pair['s1'].split('.')[0], case=False, na=False)]
            if not matching_rows.empty:
                sheet_signal = matching_rows.iloc[0].get('SIGNAL', 'N/A')
                if ind_signal != "NEUTRAL" and str(sheet_signal).upper() in ["BUY", "SELL", "ACTIVE"]:
                    excel_matched = True

        verified_results.append({
            "Pair": pair['name'],
            "Price 1": round(p1, 2),
            "Price 2": round(p2, 2),
            "Current Ratio": round(curr_ratio, 4),
            "Current Z-Score": round(curr_z, 2),
            "Independent Signal": ind_signal,
            "Excel Signal": sheet_signal,
            "Dual Confirmation": "✅ SOLID ENTRY" if excel_matched else "❌ SINGLE / NONE",
            "Exit Condition": exit_signal
        })

df_verified = pd.DataFrame(verified_results)

# Display Dual Confirmations Highlighted
st.subheader("Dual Confirmation Entries (Solid Signals)")
dual_confirmed = df_verified[df_verified['Dual Confirmation'] == "✅ SOLID ENTRY"]
if not dual_confirmed.empty:
    st.success("High-probability pair trades found with dual confirmation:")
    st.dataframe(dual_confirmed, use_container_width=True)
else:
    st.info("No dual confirmed entry signals at this time.")

st.subheader("All Verified Pairs")
st.dataframe(df_verified, use_container_width=True)

st.markdown("---")

# -----------------------------------------------------------------------------
# SECTION 6: CASH NEUTRALITY & SUGGESTIVE LOT SIZES (MAX RATIO <= 4)
# -----------------------------------------------------------------------------
st.header("3. Cash Neutrality & Suggestive Lot Sizes")

neutrality_data = []

for pair in DEFAULT_PAIRS:
    hist = fetch_historical_pair_data(pair['s1'], pair['s2'], days=5)
    if not hist.empty:
        p1 = hist['Price1'].iloc[-1]
        p2 = hist['Price2'].iloc[-1]
        
        lot1 = pair['lot1']
        lot2 = pair['lot2']
        
        val1_std = lot1 * p1
        val2_std = lot2 * p2
        
        # Exact Cash Neutral Lot 2 calculation
        exact_lot2 = (lot1 * p1) / p2
        
        # Multiplier bounded to maximum 4x standard lot size
        bounded_lot2 = min(round(exact_lot2), lot2 * 4)
        val2_rec = bounded_lot2 * p2
        
        neutrality_data.append({
            "Pair": pair['name'],
            "Leg 1 Price": round(p1, 2),
            "Leg 2 Price": round(p2, 2),
            "Standard Lot (1:1)": f"{lot1} : {lot2}",
            "Std Value Leg 1 (₹)": round(val1_std, 2),
            "Std Value Leg 2 (₹)": round(val2_std, 2),
            "Recommended Lot 2 (Max 4x)": bounded_lot2,
            "Recommended Value Leg 2 (₹)": round(val2_rec, 2),
            "Cash Neutrality Diff (₹)": round(abs(val1_std - val2_rec), 2)
        })

df_neutrality = pd.DataFrame(neutrality_data)
st.dataframe(df_neutrality, use_container_width=True)

st.markdown("---")

# -----------------------------------------------------------------------------
# SECTION 7: LAST 5 TRADING DAYS — DIVERGENCE TREND (RED / GREEN DOTS)
# -----------------------------------------------------------------------------
st.header("4. Last 5 Trading Days — Divergence Trend")

divergence_rows = []

for pair in DEFAULT_PAIRS:
    hist = fetch_historical_pair_data(pair['s1'], pair['s2'], days=10)
    if len(hist) >= 5:
        last_5 = hist.tail(5)
        dots = []
        for z in last_5['ZScore']:
            if abs(z) >= z_entry_threshold:
                dots.append("🔴")  # Diverged (Divergence / Entry threshold hit)
            else:
                dots.append("🟢")  # Converged / Normal
        
        divergence_rows.append({
            "Pair": pair['name'],
            "Day -4": dots[0],
            "Day -3": dots[1],
            "Day -2": dots[2],
            "Day -1": dots[3],
            "Today": dots[4],
            "Current Z": round(last_5['ZScore'].iloc[-1], 2)
        })

df_divergence = pd.DataFrame(divergence_rows)
st.caption("🔴 = Divergence / Threshold Exceeded | 🟢 = Normal / Converged Range")
st.dataframe(df_divergence, use_container_width=True)

st.markdown("---")

# -----------------------------------------------------------------------------
# SECTION 8 & 9: CHARTS & HISTORICAL PERFORMANCE (DROP-DOWN SELECTION)
# -----------------------------------------------------------------------------
st.header("5. Pair Charts & Historical Performance Analysis")

selected_pair_name = st.selectbox(
    "Select Pair for Detailed Technical & Historical Analysis:",
    options=[p['name'] for p in DEFAULT_PAIRS]
)

selected_pair_info = next(p for p in DEFAULT_PAIRS if p['name'] == selected_pair_name)
df_hist_90 = fetch_historical_pair_data(selected_pair_info['s1'], selected_pair_info['s2'], days=90)

if not df_hist_90.empty:
    col1, col2 = st.columns(2)
    
    # --- CHART 1: 90-DAY RATIO CHART WITH MIN, MAX, AVG ---
    with col1:
        st.subheader("90-Day Ratio Chart (Ratio = Price1 / Price2)")
        max_ratio = df_hist_90['Ratio'].max()
        min_ratio = df_hist_90['Ratio'].min()
        avg_ratio = df_hist_90['Ratio'].mean()
        
        fig_ratio = px.line(df_hist_90, x=df_hist_90.index, y='Ratio', title=f"Ratio: {selected_pair_name}")
        fig_ratio.add_hline(y=max_ratio, line_dash="dash", line_color="red", annotation_text=f"Max: {max_ratio:.4f}")
        fig_ratio.add_hline(y=avg_ratio, line_dash="solid", line_color="blue", annotation_text=f"Avg: {avg_ratio:.4f}")
        fig_ratio.add_hline(y=min_ratio, line_dash="dash", line_color="green", annotation_text=f"Min: {min_ratio:.4f}")
        fig_ratio.update_layout(showlegend=True)
        st.plotly_chart(fig_ratio, use_container_width=True)
        
    # --- CHART 2: Z-SCORE CHART WITH REFERENCE LINES ---
    with col2:
        st.subheader("Z-Score Chart against Max/Min/Avg")
        max_z = df_hist_90['ZScore'].max()
        min_z = df_hist_90['ZScore'].min()
        avg_z = df_hist_90['ZScore'].mean()
        
        fig_z = px.line(df_hist_90, x=df_hist_90.index, y='ZScore', title=f"Z-Score: {selected_pair_name}")
        fig_z.add_hline(y=max_z, line_dash="dash", line_color="red", annotation_text=f"Max Z: {max_z:.2f}")
        fig_z.add_hline(y=avg_z, line_dash="solid", line_color="blue", annotation_text=f"Avg Z: {avg_z:.2f}")
        fig_z.add_hline(y=min_z, line_dash="dash", line_color="green", annotation_text=f"Min Z: {min_z:.2f}")
        fig_z.add_hline(y=z_entry_threshold, line_dash="dot", line_color="orange", annotation_text="Upper Entry")
        fig_z.add_hline(y=-z_entry_threshold, line_dash="dot", line_color="orange", annotation_text="Lower Entry")
        fig_z.update_layout(showlegend=True)
        st.plotly_chart(fig_z, use_container_width=True)

    # --- HISTORICAL TRADE PERFORMANCE SIMULATION ---
    st.subheader(f"Historical Trade Performance (Last 90 Days) - {selected_pair_name}")
    
    # Backtest simulation logic based on Z-Score entries/exits
    in_trade = False
    trade_type = None
    entry_ratio = 0
    trades = []
    
    for date, row in df_hist_90.iterrows():
        z = row['ZScore']
        r = row['Ratio']
        
        if not in_trade:
            if z >= z_entry_threshold:
                in_trade = True
                trade_type = "SHORT RATIO (Sell Leg1 / Buy Leg2)"
                entry_ratio = r
                entry_date = date
            elif z <= -z_entry_threshold:
                in_trade = True
                trade_type = "LONG RATIO (Buy Leg1 / Sell Leg2)"
                entry_ratio = r
                entry_date = date
        else:
            if abs(z) <= z_exit_threshold:
                exit_ratio = r
                if trade_type == "SHORT RATIO (Sell Leg1 / Buy Leg2)":
                    pnl_pct = ((entry_ratio - exit_ratio) / entry_ratio) * 100
                else:
                    pnl_pct = ((exit_ratio - entry_ratio) / entry_ratio) * 100
                
                trades.append({
                    "Entry Date": entry_date.strftime('%Y-%m-%d'),
                    "Exit Date": date.strftime('%Y-%m-%d'),
                    "Trade Type": trade_type,
                    "Entry Ratio": round(entry_ratio, 4),
                    "Exit Ratio": round(exit_ratio, 4),
                    "Return (%)": round(pnl_pct, 2),
                    "Outcome": "WIN ✅" if pnl_pct > 0 else "LOSS ❌"
                })
                in_trade = False

    df_trades = pd.DataFrame(trades)
    
    if not df_trades.empty:
        total_trades = len(df_trades)
        win_rate = round((len(df_trades[df_trades['Return (%)'] > 0]) / total_trades) * 100, 2)
        avg_return = round(df_trades['Return (%)'].mean(), 2)
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Executed Trades", total_trades)
        m2.metric("Win Rate (%)", f"{win_rate}%")
        m3.metric("Average Return per Trade", f"{avg_return}%")
        
        st.dataframe(df_trades, use_container_width=True)
    else:
        st.info("No completed historical trade setups found in the last 90 days for this pair at the chosen threshold.")