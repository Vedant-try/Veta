import streamlit as st
import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from io import BytesIO
import time

# CAGR formula display
st.markdown("""
**CAGR Formula**  
The Compound Annual Growth Rate is calculated using:
$$ CAGR = \\left( \\frac{Ending\\ Value}{Beginning\\ Value} \\right)^{\\frac{1}{Years}} - 1 $$
""")

# Initialize session state
if 'tickers' not in st.session_state:
    st.session_state.tickers = ['']
if 'cagr_data' not in st.session_state:
    st.session_state.cagr_data = pd.DataFrame()

# Date inputs with validation
def get_dates():
    col1, col2 = st.columns(2)
    with col1:
        start_date = st.date_input("Start date", 
                                 datetime.today() - timedelta(days=365*5))
    with col2:
        end_date = st.date_input("End date", datetime.today())
    
    if start_date > end_date:
        st.error("End date must be after start date")
        return None, None
    return start_date.strftime('%Y-%m-%d'), end_date.strftime('%Y-%m-%d')

# Enhanced ticker validation
def validate_ticker(ticker):
    try:
        info = yf.Ticker(ticker).info
        if info['regularMarketPrice'] is None:
            raise ValueError
        return True
    except:
        return False

# Improved calculation function
def calculate_cagr(tickers, start, end):
    results = []
    for t in tickers:
        try:
            if not validate_ticker(t):
                raise ValueError(f"Invalid ticker: {t}")
            
            time.sleep(0.7)  # Increased delay for rate limiting
            data = yf.download(t, start=start, end=end, progress=False)
            
            if len(data) < 2:
                raise ValueError("Insufficient historical data")
                
            years = (datetime.strptime(end, '%Y-%m-%d') - 
                    datetime.strptime(start, '%Y-%m-%d')).days / 365.25
            start_price = data['Close'].iloc[0]
            end_price = data['Close'].iloc[-1]
            
            if start_price <= 0:
                raise ValueError("Invalid starting price")
                
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

# Ticker input section
st.header("Ticker Input")
input_method = st.radio("Input method:", ["Manual Entry", "CSV Upload"])

if input_method == "Manual Entry":
    st.subheader("Add/Remove Tickers")
    for idx in range(len(st.session_state.tickers)):
        cols = st.columns([6, 1])
        with cols[0]:
            st.text_input(f"Ticker {idx+1}", 
                        key=f"ticker_{idx}",
                        value=st.session_state.tickers[idx],
                        help="Use Yahoo Finance format (e.g., 'TCS.NS' for NSE)")
        with cols[1]:
            st.button("❌", key=f"remove_{idx}", 
                    on_click=lambda idx=idx: st.session_state.tickers.pop(idx))
    st.button("➕ Add Ticker", on_click=lambda: st.session_state.tickers.append(''))
else:
    uploaded_file = st.file_uploader("Upload CSV", type=["csv"])
    if uploaded_file:
        try:
            df = pd.read_csv(uploaded_file)
            st.session_state.tickers = df['Ticker'].dropna().tolist()[:1000]
        except:
            st.error("Invalid CSV format - must contain 'Ticker' column")

# Main processing
start, end = get_dates()
if st.button("Calculate CAGR") and start and end:
    valid_tickers = [t.strip() for t in st.session_state.tickers if t.strip()]
    
    if not valid_tickers:
        st.error("Please enter at least one valid ticker")
    else:
        with st.spinner(f"Fetching data for {len(valid_tickers)} tickers..."):
            try:
                st.session_state.cagr_data = calculate_cagr(valid_tickers, start, end)
            except Exception as e:
                st.error(f"Critical error: {str(e)}")

# Display results
if not st.session_state.cagr_data.empty:
    st.subheader("Results")
    styled_df = st.session_state.cagr_data.style.format({
        'CAGR (%)': '{:.2f}%',
        'Start Date': lambda x: str(x)[:10],
        'End Date': lambda x: str(x)[:10]
    })
    st.dataframe(styled_df, use_container_width=True)
    
    # Excel download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        st.session_state.cagr_data.to_excel(writer, index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        format_percent = workbook.add_format({'num_format': '0.00%'})
        worksheet.set_column('D:D', 12, format_percent)
    
    st.download_button(
        "📥 Download Excel Report",
        output.getvalue(),
        "cagr_report.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Help section
with st.expander("Troubleshooting Guide"):
    st.markdown("""
    **Common Issues Solutions:**
    1. **Unrecognized Tickers**  
       - Use Yahoo Finance format (e.g., `INFY.NS` for NSE, `TATAMOTORS.BO` for BSE)
       - Check symbol on [Yahoo Finance](https://finance.yahoo.com/)
    
    2. **Date Range Errors**  
       - Ensure end date is after start date
       - Avoid weekends/market holidays
    
    3. **Data Not Found**  
       - Try different date ranges
       - Check if company was listed during selected period
    """)
