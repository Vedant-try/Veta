import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime

# ------------------ Page Config ------------------
st.set_page_config(page_title="CAGR Calculator", layout="wide")

# ------------------ Title ------------------
st.title("📊 Multi-Stock CAGR Calculator")
st.markdown("""
Upload a CSV with a `Ticker` column **or** manually enter a ticker below. 
Choose the date range, and get price performance + CAGR summary.

**Example Ticker:** RELIANCE.NS, INFY.NS, AAPL, MSFT, etc.
""")

# ------------------ Download Sample CSV ------------------
def get_sample_csv():
    sample = pd.DataFrame({"Ticker": ["RELIANCE.NS", "INFY.NS", "TCS.NS"]})
    return sample.to_csv(index=False).encode('utf-8')

st.download_button(
    label="📥 Download Sample CSV",
    data=get_sample_csv(),
    file_name="sample_ticker_list.csv",
    mime="text/csv"
)

# ------------------ File Upload ------------------
uploaded_file = st.file_uploader("Upload CSV with a 'Ticker' column", type=["csv"])

# ------------------ Manual Ticker Input ------------------
man_ticker = st.text_input("Or manually enter a single Ticker (e.g., RELIANCE.NS)", value="")

# ------------------ Date Inputs ------------------
today = datetime.today().date()
default_start = today.replace(year=today.year - 5)
col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Start Date", value=default_start)
with col2:
    end_date = st.date_input("End Date", value=today)

# ------------------ CAGR Function ------------------
def calculate_cagr(start_price, end_price, years):
    if start_price > 0 and end_price > 0 and years > 0:
        return ((end_price / start_price) ** (1 / years)) - 1
    return None

# ------------------ Prepare Ticker List ------------------
tickers = []
if uploaded_file:
    df = pd.read_csv(uploaded_file)
    if 'Ticker' in df.columns:
        tickers = df['Ticker'].dropna().unique().tolist()
    else:
        st.error("CSV must contain a 'Ticker' column.")
        st.stop()
elif man_ticker:
    tickers = [man_ticker.strip()]
else:
    st.info("Please upload a CSV or manually enter a ticker above to begin.")

# ------------------ Process Data ------------------
if tickers:
    result_data = []
    failed_tickers = []

    with st.spinner("Fetching data from Yahoo Finance..."):
        for ticker in tickers:
            try:
                data = yf.Ticker(ticker).history(start=start_date, end=end_date)
                if data.empty or 'Close' not in data.columns:
                    failed_tickers.append(ticker)
                    continue

                data = data.dropna(subset=['Close'])
                start_price = data['Close'].iloc[0]
                end_price = data['Close'].iloc[-1]
                num_years = (end_date - start_date).days / 365.25
                cagr = calculate_cagr(start_price, end_price, num_years)

                result_data.append({
                    "Ticker": ticker,
                    "Start Price": round(start_price, 2),
                    "End Price": round(end_price, 2),
                    "CAGR (%)": round(cagr * 100, 2) if cagr is not None else None
                })

            except Exception as e:
                failed_tickers.append(ticker)

    # ------------------ Display Results ------------------
    if result_data:
        result_df = pd.DataFrame(result_data)
        st.subheader("📈 CAGR Summary")
        st.dataframe(result_df)

        # Download option
        csv_data = result_df.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Results as CSV",
            data=csv_data,
            file_name="cagr_results.csv",
            mime="text/csv"
        )

    if failed_tickers:
        st.warning(f"Could not fetch data for: {', '.join(failed_tickers)}")

# ------------------ Footer ------------------
st.markdown("""
---
💡 **CAGR Formula:**

\[ \text{CAGR} = \left( \frac{\text{Final Price}}{\text{Initial Price}} \right)^{\frac{1}{\text{Years}}} - 1 \]
""")
