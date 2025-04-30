import streamlit as st
import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO

# CAGR formula display
st.markdown("""
**CAGR Formula**  
The Compound Annual Growth Rate is calculated using:
$$ CAGR = \\left( \\frac{Ending\\ Value}{Beginning\\ Value} \\right)^{\\frac{1}{Years}} - 1 $$
""")

# Initialize session state for dynamic fields
if 'tickers' not in st.session_state:
    st.session_state.tickers = ['']
if 'cagr_data' not in st.session_state:
    st.session_state.cagr_data = pd.DataFrame()

# Dynamic ticker input functions
def add_ticker():
    st.session_state.tickers.append('')

def remove_ticker(idx):
    del st.session_state.tickers[idx]

# File upload/download handlers
def download_template():
    template = pd.DataFrame(columns=['Ticker'])
    towrite = BytesIO()
    template.to_excel(towrite, index=False)
    towrite.seek(0)
    return towrite.read()

# Date inputs
col1, col2 = st.columns(2)
with col1:
    start_date = st.date_input("Start date", datetime(2020, 1, 1))
with col2:
    end_date = st.date_input("End date", datetime.today())

# Ticker input section
st.header("Ticker Input Methods")
input_method = st.radio("Choose input method:", 
                       ["Manual Entry", "CSV Upload"])

if input_method == "Manual Entry":
    st.subheader("Add/Remove Tickers")
    for idx, ticker in enumerate(st.session_state.tickers):
        cols = st.columns([6, 1])
        with cols[0]:
            st.text_input(f"Ticker {idx+1}", key=f"ticker_{idx}",
                         value=ticker)
        with cols[1]:
            st.button("❌", key=f"remove_{idx}", 
                     on_click=remove_ticker, args=(idx,))
    st.button("➕ Add Ticker", on_click=add_ticker)
else:
    uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])
    if uploaded_file:
        df = pd.read_csv(uploaded_file)
        st.session_state.tickers = df['Ticker'].tolist()[:1000]
    
    st.download_button("Download CSV Template", 
                      download_template(),
                      "ticker_template.csv",
                      "application/vnd.ms-excel")

# Calculation function
def calculate_cagr(tickers, start, end):
    results = []
    for t in tickers:
        try:
            # Add delay to avoid 429 errors
            time.sleep(0.5)
            
            data = yf.download(t, start=start, end=end)
            if len(data) < 2:
                raise ValueError("Insufficient data")
                
            years = (end - start).days / 365.25
            start_price = data['Close'].iloc[0]
            end_price = data['Close'].iloc[-1]
            cagr = (end_price/start_price)**(1/years) - 1
            
            results.append({
                'Ticker': t,
                'CAGR (%)': cagr * 100,
                'Status': 'Success'
            })
        except Exception as e:
            results.append({
                'Ticker': t,
                'CAGR (%)': np.nan,
                'Status': f"Error: {str(e)[:50]}"
            })
    return pd.DataFrame(results)

# Main calculation
if st.button("Calculate CAGR"):
    valid_tickers = [t for t in st.session_state.tickers if t.strip()]
    
    if not valid_tickers:
        st.error("Please enter at least one valid ticker")
    else:
        with st.spinner("Calculating..."):
            try:
                st.session_state.cagr_data = calculate_cagr(
                    valid_tickers, start_date, end_date
                )
            except Exception as e:
                st.error(f"Critical error: {str(e)}")

# Display results
if not st.session_state.cagr_data.empty:
    st.subheader("Results")
    st.dataframe(
        st.session_state.cagr_data.style.format({'CAGR (%)': '{:.2f}%'}),
        use_container_width=True
    )
    
    # Excel download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        st.session_state.cagr_data.to_excel(writer, index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        format_percent = workbook.add_format({'num_format': '0.00%'})
        worksheet.set_column('B:B', 12, format_percent)
    
    st.download_button(
        "Download Excel Report",
        output.getvalue(),
        "cagr_results.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Error handling documentation
st.markdown("""
**Common Errors:**
- Invalid ticker symbols
- Insufficient historical data
- Yahoo Finance API limits (wait 60 seconds and try again)
""")
