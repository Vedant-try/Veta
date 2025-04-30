import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
from io import BytesIO
import numpy as np

st.title("CAGR Calculator (Simple & Robust)")

st.markdown("""
This tool fetches the price for each ticker on the start and end dates (or nearest trading days) and computes CAGR.
""")

# --- Sidebar: Ticker input ---
st.sidebar.header("Inputs")
uploaded_file = st.sidebar.file_uploader("Upload a CSV file with 'Ticker' column", type=["csv"])

tickers = []

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    if 'Ticker' in df.columns:
        tickers = df['Ticker'].dropna().astype(str).tolist()
        st.sidebar.success(f"Loaded {len(tickers)} tickers from file.")
    else:
        st.sidebar.error("CSV must have a column named 'Ticker'.")
else:
    n = st.sidebar.number_input("Number of tickers", 1, 100, 2)
    for i in range(n):
        t = st.sidebar.text_input(f"Ticker {i+1}", "")
        if t: tickers.append(t.strip())

# --- Sidebar: Date input ---
today = datetime.today().date()
default_start = today - timedelta(days=365*5)
start_date = st.sidebar.date_input("Start Date", default_start, min_value=today - timedelta(days=365*30), max_value=today)
end_date = st.sidebar.date_input("End Date", today, min_value=start_date, max_value=today)

# --- CAGR function ---
def get_nearest_price(data, target_date):
    """Return the price on the target_date or closest previous date."""
    if data.empty:
        return np.nan, None
    # Ensure index is datetime.date
    data = data.copy()
    data.index = pd.to_datetime(data.index).date
    available_dates = data.index[data.index <= target_date]
    if len(available_dates) == 0:
        # All data is after the target date; use earliest available
        nearest_date = min(data.index)
    else:
        nearest_date = max(available_dates)
    price = data.loc[nearest_date, 'Adj Close']
    return price, nearest_date

def calculate_cagr(start_price, end_price, years):
    if start_price > 0 and end_price > 0 and years > 0:
        return (end_price/start_price)**(1/years) - 1
    else:
        return np.nan

# --- Main calculation ---
if st.sidebar.button("Calculate CAGR"):
    if not tickers:
        st.error("Please provide at least one ticker.")
    else:
        results = []
        for ticker in tickers:
            try:
                data = yf.download(ticker, start=start_date - timedelta(days=7), end=end_date + timedelta(days=7), progress=False)
                if data.empty or 'Adj Close' not in data.columns:
                    results.append({
                        'Ticker': ticker,
                        'Beginning Value': 'N/A',
                        'Ending Value': 'N/A',
                        'CAGR (%)': 'No data'
                    })
                    continue
                start_price, actual_start = get_nearest_price(data, start_date)
                end_price, actual_end = get_nearest_price(data, end_date)
                if pd.isna(start_price) or pd.isna(end_price):
                    results.append({
                        'Ticker': ticker,
                        'Beginning Value': 'N/A',
                        'Ending Value': 'N/A',
                        'CAGR (%)': 'No data'
                    })
                    continue
                years = (actual_end - actual_start).days / 365.25
                cagr = calculate_cagr(start_price, end_price, years)
                results.append({
                    'Ticker': ticker,
                    'Beginning Value': f"{start_price:.2f} ({actual_start})",
                    'Ending Value': f"{end_price:.2f} ({actual_end})",
                    'CAGR (%)': f"{cagr*100:.2f}%" if not np.isnan(cagr) else "N/A"
                })
            except Exception as e:
                results.append({
                    'Ticker': ticker,
                    'Beginning Value': 'N/A',
                    'Ending Value': 'N/A',
                    'CAGR (%)': f"Error: {str(e)}"
                })
        result_df = pd.DataFrame(results)
        st.subheader("CAGR Results")
        st.dataframe(result_df, use_container_width=True)

        # Download as Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result_df.to_excel(writer, index=False)
        st.download_button(
            "Download Results as Excel",
            output.getvalue(),
            "cagr_results.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

st.markdown("""
---
**How it works:**  
- For each ticker, the app fetches prices as close as possible to your chosen start and end dates.
- If the exact date is a holiday/weekend, it uses the last available price before that date.
- If no data is available at all, "No data" is shown.
""")
