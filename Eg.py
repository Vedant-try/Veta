import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime
import io

# ---------- UI SETUP ----------
st.set_page_config(page_title="Stock CAGR Calculator", layout="centered")

st.markdown(
    """
    <h2 style='text-align: center;'>📈 Stock CAGR Calculator</h2>
    <p style='text-align: center;'>Calculate the Compound Annual Growth Rate (CAGR) for up to 1000 stocks</p>
    <hr>
    <p><b>CAGR Formula:</b></p>
    <p style='background-color: #f0f2f6; padding: 10px; border-radius: 5px;'>
    \\[
    CAGR = \\left( \\frac{\\text{Ending Price}}{\\text{Beginning Price}} \\right)^{\\frac{1}{\\text{Years}}} - 1
    \\]
    </p>
    """,
    unsafe_allow_html=True
)

# ---------- USER INPUT ----------
upload_option = st.radio("Choose input method for tickers:", ("Manual Entry", "Upload CSV"))

tickers = []

if upload_option == "Manual Entry":
    num = st.number_input("Number of ticker input fields (Max 1000)", min_value=1, max_value=1000, value=1)
    for i in range(num):
        ticker = st.text_input(f"Enter Ticker {i+1}", key=f"ticker_{i}")
        if ticker:
            tickers.append(ticker.upper())
else:
    sample = pd.DataFrame({'Tickers': ['AAPL', 'MSFT', 'GOOG']})
    csv = sample.to_csv(index=False).encode('utf-8')
    st.download_button("Download Sample CSV", csv, "sample_tickers.csv", "text/csv")
    
    uploaded_file = st.file_uploader("Upload a CSV with one column 'Tickers'", type=['csv'])
    if uploaded_file:
        try:
            df_uploaded = pd.read_csv(uploaded_file)
            tickers = df_uploaded['Tickers'].dropna().astype(str).str.upper().tolist()
        except Exception as e:
            st.error(f"Error reading uploaded CSV: {e}")

# ---------- DATE INPUT ----------
col1, col2 = st.columns(2)
start_date = col1.date_input("From Date", value=datetime(2010, 1, 1))
end_date = col2.date_input("To Date", value=datetime.today())

# ---------- CAGR CALCULATION ----------
def calculate_cagr(start_price, end_price, years):
    if start_price <= 0 or years <= 0:
        return None
    return ((end_price / start_price) ** (1 / years)) - 1

# ---------- RESULTS GENERATION ----------
if st.button("Generate Results") and tickers:
    st.info("Fetching data... Please wait.")
    
    result_data = []
    errors = []

    for ticker in tickers:
        try:
            data = yf.download(ticker, start=start_date, end=end_date)
            if data.empty:
                errors.append(ticker)
                continue

            start_price = data['Adj Close'].iloc[0]
            end_price = data['Adj Close'].iloc[-1]
            years = (end_date - start_date).days / 365.25

            cagr = calculate_cagr(start_price, end_price, years)
            if cagr is not None:
                result_data.append({
                    "Ticker": ticker,
                    "Start Price": round(start_price, 2),
                    "End Price": round(end_price, 2),
                    "Years": round(years, 2),
                    "CAGR (%)": round(cagr * 100, 2)
                })
        except Exception as e:
            errors.append(ticker)

    # ---------- DISPLAY RESULTS ----------
    if result_data:
        result_df = pd.DataFrame(result_data)
        st.success("CAGR Calculation Completed!")
        st.dataframe(result_df)

        # Excel download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result_df.to_excel(writer, index=False, sheet_name='CAGR Results')
            workbook = writer.book
            worksheet = writer.sheets['CAGR Results']
            format1 = workbook.add_format({'num_format': '0.00', 'align': 'center'})
            worksheet.set_column('A:E', 15, format1)
            worksheet.write('G1', 'Formula: CAGR = (Ending / Beginning)^(1/Years) - 1')

        st.download_button(
            label="📥 Download Excel Report",
            data=output.getvalue(),
            file_name="CAGR_Results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No valid data found for the tickers provided.")

    if errors:
        st.error(f"No data found or error for: {', '.join(errors)}")

elif st.button("Generate Results") and not tickers:
    st.warning("Please provide at least one ticker to calculate CAGR.")
