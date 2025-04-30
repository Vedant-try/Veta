import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime
import io

# ------------------ Streamlit Page Setup ------------------
st.set_page_config("📊 Indian Stock CAGR Calculator", layout="centered")

st.title("📈 Indian Stock CAGR Calculator")
st.write("This tool calculates the CAGR of Indian listed stocks (NSE) using Yahoo Finance.")

st.latex(r"""
\text{CAGR} = \left( \frac{\text{Ending Price}}{\text{Beginning Price}} \right)^{\frac{1}{\text{Years}}} - 1
""")

# ------------------ Ticker Input Section ------------------
st.subheader("1. Input Stock Tickers")

if "ticker_count" not in st.session_state:
    st.session_state.ticker_count = 1

col1, col2 = st.columns(2)
with col1:
    if st.button("➕ Add Ticker"):
        if st.session_state.ticker_count < 1000:
            st.session_state.ticker_count += 1
with col2:
    if st.button("➖ Remove Ticker"):
        if st.session_state.ticker_count > 1:
            st.session_state.ticker_count -= 1

tickers = []
for i in range(st.session_state.ticker_count):
    ticker = st.text_input(f"Ticker {i + 1}", value="", key=f"ticker_{i}")
    if ticker.strip():
        tickers.append(ticker.strip().upper() + ".NS")  # Adding ".NS" for NSE tickers

# ------------------ Date Inputs ------------------
st.subheader("2. Select Date Range")
col1, col2 = st.columns(2)
start_date = col1.date_input("From Date", value=datetime(2015, 1, 1))
end_date = col2.date_input("To Date", value=datetime.today())

# ------------------ CAGR Calculation Function ------------------
def calculate_cagr(start_price, end_price, years):
    if start_price <= 0 or years <= 0:
        return None
    return ((end_price / start_price) ** (1 / years)) - 1

# ------------------ Generate Button ------------------
st.subheader("3. Generate Results")
if st.button("🚀 Generate CAGR Results"):

    if not tickers:
        st.warning("Please enter at least one stock ticker.")
    else:
        st.info("Fetching data... please wait ⏳")
        result_data = []
        errors = []

        for ticker in tickers:
            try:
                data = yf.download(ticker, start=start_date, end=end_date)
                if data.empty or 'Adj Close' not in data:
                    errors.append(ticker.replace(".NS", ""))
                    continue

                start_price = data['Adj Close'].iloc[0]
                end_price = data['Adj Close'].iloc[-1]
                years = (end_date - start_date).days / 365.25

                cagr = calculate_cagr(start_price, end_price, years)
                result_data.append({
                    "Ticker": ticker.replace(".NS", ""),
                    "Start Price (₹)": round(start_price, 2),
                    "End Price (₹)": round(end_price, 2),
                    "Years": round(years, 2),
                    "CAGR (%)": round(cagr * 100, 2) if cagr is not None else "N/A"
                })

            except Exception as e:
                errors.append(ticker.replace(".NS", ""))

        # ------------------ Display Results ------------------
        if result_data:
            result_df = pd.DataFrame(result_data)
            st.success("CAGR Calculation Completed ✅")
            st.dataframe(result_df)

            # ------------------ Excel Download ------------------
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='CAGR Results')
                workbook = writer.book
                worksheet = writer.sheets['CAGR Results']
                worksheet.set_column('A:E', 18)

                # Add formula at top
                worksheet.write('G1', 'Formula:')
                worksheet.write('G2', 'CAGR = (End / Start)^(1/Years) - 1')

            st.download_button(
                label="📥 Download Excel Report",
                data=output.getvalue(),
                file_name="CAGR_Results_India.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        if errors:
            st.error("No data found for these tickers: " + ", ".join(errors))
