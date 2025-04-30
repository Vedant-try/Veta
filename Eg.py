import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime
import io

# -------------------- Setup --------------------
st.set_page_config("📈 Global Stock CAGR Calculator", layout="centered")
st.title("📈 Global Stock CAGR Calculator")

st.latex(r"""
\text{CAGR} = \left( \frac{\text{Ending Price}}{\text{Beginning Price}} \right)^{\frac{1}{\text{Years}}} - 1
""")

# -------------------- Ticker Inputs --------------------
st.subheader("1. Add Tickers Manually or Upload a CSV")

# --- Sample CSV Download ---
with st.expander("📄 Sample CSV Format"):
    st.markdown("Upload a CSV file with a column named `Ticker` (no header row also works).")
    sample_csv = pd.DataFrame({"Ticker": ["AAPL", "MSFT", "GOOGL"]})
    csv_bytes = sample_csv.to_csv(index=False).encode()
    st.download_button("📥 Download Sample CSV", data=csv_bytes, file_name="sample_tickers.csv", mime="text/csv")

# --- Upload CSV or Use Manual Input ---
use_csv = st.radio("Select Input Method", ["Manual Entry", "Upload CSV"])

tickers = []

if use_csv == "Manual Entry":
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

    for i in range(st.session_state.ticker_count):
        ticker = st.text_input(f"Ticker {i + 1}", key=f"ticker_{i}")
        if ticker.strip():
            tickers.append(ticker.strip().upper())

else:
    uploaded_file = st.file_uploader("Upload your CSV", type=["csv"])
    if uploaded_file:
        try:
            df_uploaded = pd.read_csv(uploaded_file, header=None)
            tickers = df_uploaded.iloc[:, 0].astype(str).str.upper().tolist()
            st.success(f"✅ Loaded {len(tickers)} tickers from CSV")
        except Exception as e:
            st.error("❌ Error reading CSV. Please ensure it has one column with tickers.")

# -------------------- Date Inputs --------------------
st.subheader("2. Select Date Range")
col1, col2 = st.columns(2)
start_date = col1.date_input("From Date", value=datetime(2015, 1, 1))
end_date = col2.date_input("To Date", value=datetime.today())

# -------------------- CAGR Function --------------------
def calculate_cagr(start_price, end_price, years):
    if start_price <= 0 or years <= 0:
        return None
    return ((end_price / start_price) ** (1 / years)) - 1

# -------------------- Results Section --------------------
st.subheader("3. Generate Results")
if st.button("🚀 Generate CAGR Results"):
    if not tickers:
        st.warning("⚠️ Please enter or upload at least one valid ticker.")
    else:
        st.info("⏳ Fetching data from Yahoo Finance...")
        results = []
        errors = []

        for ticker in tickers:
            try:
                # Add 7-day buffer before/after to handle missing trading days
                data = yf.download(ticker, start=start_date - pd.Timedelta(days=7), end=end_date + pd.Timedelta(days=7))

                if data.empty or 'Adj Close' not in data:
                    errors.append(ticker)
                    continue

                # Drop NaNs and find first and last valid prices within range
                data = data[['Adj Close']].dropna()
                data = data[(data.index >= pd.to_datetime(start_date)) & (data.index <= pd.to_datetime(end_date))]

                if data.empty:
                    errors.append(ticker)
                    continue

                start_price = data['Adj Close'].iloc[0]
                end_price = data['Adj Close'].iloc[-1]
                actual_start = data.index[0]
                actual_end = data.index[-1]
                years = (actual_end - actual_start).days / 365.25

                cagr = calculate_cagr(start_price, end_price, years)

                results.append({
                    "Ticker": ticker,
                    "Start Price": round(start_price, 2),
                    "End Price": round(end_price, 2),
                    "Years": round(years, 2),
                    "CAGR (%)": round(cagr * 100, 2) if cagr is not None else "N/A"
                })

            except Exception as e:
                errors.append(ticker)

        if results:
            result_df = pd.DataFrame(results)
            st.success("✅ CAGR Calculation Completed")
            st.dataframe(result_df)

            # Excel export
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='CAGR Results')
                worksheet = writer.sheets['CAGR Results']
                worksheet.set_column('A:E', 20)
                worksheet.write('G1', 'Formula:')
                worksheet.write('G2', 'CAGR = (End / Start)^(1/Years) - 1')

            st.download_button(
                label="📥 Download Excel",
                data=output.getvalue(),
                file_name="CAGR_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        if errors:
            st.error(f"❌ No data found or no trading days for: {', '.join(errors)}")

            # Excel export
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                result_df.to_excel(writer, index=False, sheet_name='CAGR Results')
                worksheet = writer.sheets['CAGR Results']
                worksheet.set_column('A:E', 20)
                worksheet.write('G1', 'Formula:')
                worksheet.write('G2', 'CAGR = (End / Start)^(1/Years) - 1')

            st.download_button(
                label="📥 Download Excel",
                data=output.getvalue(),
                file_name="CAGR_Results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        if errors:
            st.error(f"❌ No data found for: {', '.join(errors)}")
