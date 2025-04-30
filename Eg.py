import streamlit as st
import pandas as pd
import yfinance as yf
import io
import datetime

st.set_page_config(page_title="CAGR Calculator", layout="wide")
st.title("📈 CAGR Calculator from Yahoo Finance Data")

# --- CAGR Formula Display ---
with st.expander("📌 What is CAGR?"):
    st.markdown(r"""
    **Compound Annual Growth Rate (CAGR)** is calculated using the formula:

    \[
    \text{CAGR} = \left( \frac{\text{End Price}}{\text{Start Price}} \right)^{\frac{1}{\text{Years}}} - 1
    \]

    It measures the smoothed annual growth rate of an investment over a period of time.
    """)

# --- Date Inputs ---
col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Start Date", value=datetime.date(2018, 1, 1))
with col2:
    end_date = st.date_input("End Date", value=datetime.date.today())

# --- Manual Ticker Entry ---
st.subheader("📝 Enter Tickers Manually or Upload CSV")
input_method = st.radio("Choose input method", ["Manual Entry", "Upload CSV"])

tickers = []

if input_method == "Manual Entry":
    if 'ticker_count' not in st.session_state:
        st.session_state.ticker_count = 1

    col1, col2 = st.columns([3, 1])
    with col1:
        for i in range(st.session_state.ticker_count):
            tickers.append(st.text_input(f"Ticker {i+1}", key=f"ticker_{i}"))

    with col2:
        if st.button("➕ Add Ticker") and st.session_state.ticker_count < 1000:
            st.session_state.ticker_count += 1
        if st.button("➖ Remove Ticker") and st.session_state.ticker_count > 1:
            st.session_state.ticker_count -= 1

else:
    sample_df = pd.DataFrame({"Ticker": ["AAPL", "MSFT", "HDFCBANK.NS"]})
    csv_buffer = io.BytesIO()
    sample_df.to_csv(csv_buffer, index=False)
    st.download_button("📥 Download Sample CSV", data=csv_buffer.getvalue(), file_name="sample_tickers.csv")

    uploaded_file = st.file_uploader("Upload a CSV with a column named 'Ticker'", type=["csv"])
    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file)
            if "Ticker" in df.columns:
                tickers = df["Ticker"].dropna().astype(str).tolist()
            else:
                st.error("❌ CSV must contain a 'Ticker' column.")
        except Exception as e:
            st.error(f"❌ Error reading CSV: {e}")

# --- Function to fetch and compute CAGR ---
def get_clean_cagr_data(ticker, start_date, end_date):
    try:
        data = yf.download(ticker, start=start_date - pd.Timedelta(days=7), end=end_date + pd.Timedelta(days=7))
        if data.empty or 'Adj Close' not in data:
            return None
        data = data[['Adj Close']].dropna()
        data = data[(data.index >= pd.to_datetime(start_date)) & (data.index <= pd.to_datetime(end_date))]
        if data.empty or len(data) < 2:
            return None
        actual_start = data.index[0]
        actual_end = data.index[-1]
        start_price = data['Adj Close'].iloc[0]
        end_price = data['Adj Close'].iloc[-1]
        years = (actual_end - actual_start).days / 365.25
        if years <= 0:
            return None
        cagr = (end_price / start_price) ** (1 / years) - 1
        return {
            "Ticker": ticker,
            "Start Price": round(start_price, 2),
            "End Price": round(end_price, 2),
            "Years": round(years, 2),
            "CAGR (%)": round(cagr * 100, 2)
        }
    except:
        return None

# --- Generate Button ---
if st.button("🚀 Generate Results"):
    if not tickers:
        st.warning("⚠️ Please provide at least one valid ticker.")
    else:
        results = []
        errors = []

        with st.spinner("Fetching data..."):
            for ticker in tickers:
                if ticker.strip() == "":
                    continue
                result = get_clean_cagr_data(ticker.strip(), pd.to_datetime(start_date), pd.to_datetime(end_date))
                if result:
                    results.append(result)
                else:
                    errors.append(ticker)

        if results:
            result_df = pd.DataFrame(results)
            st.success("✅ Data fetched successfully!")
            st.dataframe(result_df)

            # --- Excel Export ---
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
            st.warning(f"⚠️ No data found for these tickers: {', '.join(errors)}")
