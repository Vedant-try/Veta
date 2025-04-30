import streamlit as st
import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from io import BytesIO
import time

st.markdown("""
**CAGR Formula**  
$$ CAGR = \\left( \\frac{Ending\\ Value}{Beginning\\ Value} \\right)^{\\frac{1}{Years}} - 1 $$
""")

if 'tickers' not in st.session_state:
    st.session_state.tickers = ['']

if 'cagr_data' not in st.session_state:
    st.session_state.cagr_data = pd.DataFrame()

# Date inputs
col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Start date", datetime.today() - timedelta(days=365*5))
with col2:
    end_date = st.date_input("End date", datetime.today())

# Ticker input section
st.header("Ticker Input")
input_method = st.radio("Input method:", ["Manual Entry", "CSV Upload"])

if input_method == "Manual Entry":
    st.subheader("Add/Remove Tickers")
    for idx in range(len(st.session_state.tickers)):
        cols = st.columns([6, 1])
        with cols[0]:
            st.session_state.tickers[idx] = st.text_input(
                f"Ticker {idx+1}",
                value=st.session_state.tickers[idx],
                key=f"ticker_{idx}",
                help="Use Yahoo Finance format (e.g., 'TCS.NS', 'AAPL')"
            )
        with cols[1]:
            if st.button("❌", key=f"remove_{idx}"):
                st.session_state.tickers.pop(idx)
                st.experimental_rerun()
    st.button("➕ Add Ticker", on_click=lambda: st.session_state.tickers.append(''))
else:
    uploaded_file = st.file_uploader("Upload CSV", type=["csv"])
    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file)
            st.session_state.tickers = df['Ticker'].dropna().tolist()[:1000]
        except:
            st.error("Invalid CSV format - must contain 'Ticker' column")

# CAGR Calculation
def calculate_cagr(tickers, start, end):
    results = []
    for t in tickers:
        t = t.strip()
        try:
            # Try to fetch data directly
            data = yf.download(t, start=start, end=end, progress=False)
            if data.empty or len(data['Close'].dropna()) < 2:
                raise ValueError("No data found for this ticker and date range.")
            years = (pd.to_datetime(end) - pd.to_datetime(start)).days / 365.25
            start_price = data['Close'].iloc[0]
            end_price = data['Close'].iloc[-1]
            if start_price <= 0:
                raise ValueError("Invalid starting price.")
            cagr = (end_price/start_price)**(1/years) - 1
            results.append({
                'Ticker': t,
                'Start Date': start,
                'End Date': end,
                'CAGR (%)': cagr * 100,
                'Status': 'Success'
            })
        except Exception as e:
            results.append({
                'Ticker': t,
                'Start Date': start,
                'End Date': end,
                'CAGR (%)': np.nan,
                'Status': f"Error: {str(e)}"
            })
    return pd.DataFrame(results)

if st.button("Calculate CAGR"):
    valid_tickers = [t for t in st.session_state.tickers if t.strip()]
    if not valid_tickers:
        st.error("Please enter at least one valid ticker")
    else:
        with st.spinner(f"Fetching data for {len(valid_tickers)} tickers..."):
            st.session_state.cagr_data = calculate_cagr(valid_tickers, start_date, end_date)

# Display results
if not st.session_state.cagr_data.empty:
    st.subheader("Results")
    st.dataframe(st.session_state.cagr_data.style.format({'CAGR (%)': '{:.2f}%'}), use_container_width=True)
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        st.session_state.cagr_data.to_excel(writer, index=False)
    st.download_button(
        "📥 Download Excel Report",
        output.getvalue(),
        "cagr_report.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

st.markdown("""
**Tips:**  
- Use Yahoo Finance tickers (e.g. `TCS.NS`, `AAPL`, `GOOG`, `INFY.NS`)
- If you get "No data found", double-check the ticker and date range.
""")
