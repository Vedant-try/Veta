import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
from io import BytesIO
import xlsxwriter
import matplotlib.pyplot as plt
import numpy as np

# Initialize Streamlit App
st.title("CAGR Calculator")
st.sidebar.header("Input Fields")

# Note for tickers
st.sidebar.markdown("**Note:** Use tickers from Yahoo Finance (e.g., RELIANCE.NS, AAPL)")

# Upload CSV file or manual entry
uploaded_file = st.sidebar.file_uploader("Upload a CSV file with stock tickers", type=["csv"])

stock_symbols = []

if uploaded_file is not None:
    tickers_df = pd.read_csv(uploaded_file)
    if 'Ticker' in tickers_df.columns:
        stock_symbols = tickers_df['Ticker'].dropna().unique().tolist()
        st.sidebar.success(f"Loaded {len(stock_symbols)} tickers from CSV.")
    else:
        st.sidebar.error("CSV must have a column named 'Ticker'.")
else:
    st.sidebar.warning("No CSV uploaded. Enter manually below.")
    num_companies = st.sidebar.number_input("Number of companies:", min_value=1, max_value=1000, value=2)
    for i in range(num_companies):
        symbol = st.sidebar.text_input(f"Enter Stock Symbol {i + 1} (e.g., RELIANCE.NS):", "")
        if symbol:
            stock_symbols.append(symbol)

# Small check
if not stock_symbols:
    st.sidebar.warning("Please upload a CSV or manually enter at least one stock symbol.")

# Download Sample CSV
def get_sample_csv():
    sample = pd.DataFrame({
        "Ticker": ["RELIANCE.NS", "TCS.NS", "INFY.NS"]
    })
    return sample.to_csv(index=False)

st.sidebar.download_button(
    label="📥 Download Sample CSV",
    data=get_sample_csv(),
    file_name="Sample_Tickers.csv",
    mime="text/csv",
)

# Date inputs for selecting start and end dates
start_date = st.sidebar.date_input(
    "Select Start Date:",
    datetime.now().date() - timedelta(days=365*5),
    min_value=datetime.now().date() - timedelta(days=20 * 365),
    max_value=datetime.now().date() - timedelta(days=30),
)

end_date = st.sidebar.date_input(
    "Select End Date:",
    datetime.now().date(),
    min_value=start_date + timedelta(days=30),
    max_value=datetime.now().date(),
)

# Initialize session state to store data
if "stock_data_dict" not in st.session_state:
    st.session_state.stock_data_dict = {}
if "cagr_summary" not in st.session_state:
    st.session_state.cagr_summary = []

# Definition and formula for CAGR
st.markdown("### Understanding CAGR (Compound Annual Growth Rate)")
st.markdown(
    r"""
    **CAGR** measures the mean annual growth rate of an investment over a specified time period. 
    It represents one of the most accurate ways to calculate returns for anything that rises or falls in value over time.
    
    The formula for CAGR is:
    """
)
st.latex(r"CAGR = \left(\frac{Ending\ Value}{Beginning\ Value}\right)^{\frac{1}{Number\ of\ Years}} - 1")

# Fetch Data Button
if st.sidebar.button("Calculate CAGR"):
    try:
        stock_data_dict = {}
        cagr_summary = []

        for stock_symbol in stock_symbols:
            try:
                # Download stock data
                stock_data = yf.download(stock_symbol, start=start_date, end=end_date)
                
                if not stock_data.empty and len(stock_data) >= 2:
                    # Calculate years between start and end date
                    years = (end_date - start_date).days / 365.25
                    
                    # Get beginning and ending values
                    beginning_value = stock_data['Adj Close'].iloc[0]
                    ending_value = stock_data['Adj Close'].iloc[-1]
                    
                    # Calculate CAGR
                    cagr = (ending_value / beginning_value) ** (1 / years) - 1
                    
                    # Store data for display
                    stock_data_dict[stock_symbol] = {
                        "data": stock_data,
                        "beginning_value": beginning_value,
                        "ending_value": ending_value,
                        "years": years,
                        "cagr": cagr
                    }
                    
                    cagr_summary.append({
                        "Stock Symbol": stock_symbol, 
                        "Beginning Value": beginning_value,
                        "Ending Value": ending_value,
                        "Years": years,
                        "CAGR (%)": cagr * 100
                    })
                else:
                    st.warning(f"Insufficient data for {stock_symbol}. Please check ticker or date range.")
            except Exception as e:
                st.warning(f"Error processing {stock_symbol}: {e}")
                
        st.session_state.stock_data_dict = stock_data_dict
        st.session_state.cagr_summary = cagr_summary

        st.success("CAGR calculated successfully!")
    except Exception as e:
        st.error(f"An error occurred: {e}")

# Display data in Streamlit
if st.session_state.stock_data_dict:
    for stock_symbol, data in st.session_state.stock_data_dict.items():
        st.subheader(f"Data for {stock_symbol}")
        
        # Display price chart
        st.write("**Price History**")
        fig, ax = plt.subplots(figsize=(10, 6))
        ax.plot(data["data"].index, data["data"]['Adj Close'])
        ax.set_title(f"{stock_symbol} Price History")
        ax.set_xlabel("Date")
        ax.set_ylabel("Adjusted Close Price")
        st.pyplot(fig)
        
        # Display CAGR calculation
        st.write(f"**CAGR Calculation for {stock_symbol}:**")
        st.write(f"Beginning Value: {data['beginning_value']:.2f}")
        st.write(f"Ending Value: {data['ending_value']:.2f}")
        st.write(f"Time Period: {data['years']:.2f} years")
        st.write(f"CAGR: {data['cagr']*100:.2f}%")

    # Display CAGR summary
    st.subheader("CAGR Summary")
    cagr_df = pd.DataFrame(st.session_state.cagr_summary)
    
    # Format the dataframe for display
    formatted_df = cagr_df.copy()
    formatted_df["Beginning Value"] = formatted_df["Beginning Value"].map("${:.2f}".format)
    formatted_df["Ending Value"] = formatted_df["Ending Value"].map("${:.2f}".format)
    formatted_df["Years"] = formatted_df["Years"].map("{:.2f}".format)
    formatted_df["CAGR (%)"] = formatted_df["CAGR (%)"].map("{:.2f}%".format)
    
    st.table(formatted_df)

# Function to generate Excel file with improved styling
def generate_excel(stock_data_dict, cagr_summary):
    try:
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True, 'nan_inf_to_errors': True})

        # Define formatting styles
        header_format = workbook.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1, 'align': 'center'})
        cell_format = workbook.add_format({'border': 1, 'align': 'center'})
        percent_format = workbook.add_format({'num_format': '0.00%', 'border': 1, 'align': 'center'})
        currency_format = workbook.add_format({'num_format': '$#,##0.00', 'border': 1, 'align': 'center'})
        bold_format = workbook.add_format({'bold': True})
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1, 'align': 'center'})

        # Create summary sheet
        summary_sheet = workbook.add_worksheet("CAGR Summary")
        headers = ["Stock Symbol", "Beginning Value", "Ending Value", "Years", "CAGR"]
        
        for col, header in enumerate(headers):
            summary_sheet.write(0, col, header, header_format)
        
        for row, data in enumerate(cagr_summary, start=1):
            summary_sheet.write(row, 0, data["Stock Symbol"], cell_format)
            summary_sheet.write(row, 1, data["Beginning Value"], currency_format)
            summary_sheet.write(row, 2, data["Ending Value"], currency_format)
            summary_sheet.write(row, 3, data["Years"], cell_format)
            summary_sheet.write(row, 4, data["CAGR (%)"] / 100, percent_format)
        
        # Auto-fit column widths
        for col_num, _ in enumerate(headers):
            summary_sheet.set_column(col_num, col_num, 15)

        # Create individual sheets for each stock
        for stock_symbol, data in stock_data_dict.items():
            worksheet = workbook.add_worksheet(stock_symbol[:31])  # Sheet names max 31 chars
            stock_data = data['data']
            
            # Write CAGR value at the top
            worksheet.write(0, 0, f"CAGR for {stock_symbol}:", bold_format)
            worksheet.write(0, 1, data['cagr'], percent_format)
            
            worksheet.write(1, 0, "Beginning Value:", bold_format)
            worksheet.write(1, 1, data['beginning_value'], currency_format)
            
            worksheet.write(2, 0, "Ending Value:", bold_format)
            worksheet.write(2, 1, data['ending_value'], currency_format)
            
            worksheet.write(3, 0, "Years:", bold_format)
            worksheet.write(3, 1, data['years'], cell_format)

            # Write price history data
            worksheet.write(5, 0, "Price History", bold_format)
            headers = ["Date", "Open", "High", "Low", "Close", "Adj Close", "Volume"]
            for col, header in enumerate(headers):
                worksheet.write(6, col, header, header_format)

            for row, (date, values) in enumerate(stock_data.iterrows(), start=7):
                worksheet.write_datetime(row, 0, date.to_pydatetime(), date_format)
                for col, value in enumerate(values, start=1):
                    if col < 6:  # Price columns
                        worksheet.write(row, col, value, currency_format)
                    else:  # Volume column
                        worksheet.write(row, col, value, cell_format)

            # Add a price chart to the Excel sheet
            chart = workbook.add_chart({'type': 'line'})
            chart.add_series({
                'name': 'Adj Close',
                'categories': [stock_symbol[:31], 7, 0, 7 + len(stock_data) - 1, 0],
                'values': [stock_symbol[:31], 7, 5, 7 + len(stock_data) - 1, 5],
            })
            chart.set_title({'name': f'{stock_symbol} Price History'})
            chart.set_x_axis({'name': 'Date', 'date_axis': True})
            chart.set_y_axis({'name': 'Price'})
            worksheet.insert_chart(7 + len(stock_data) + 2, 0, chart, {'x_scale': 1.5, 'y_scale': 1.5})

        workbook.close()
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"An error occurred while generating the Excel file: {e}")
        return None

# Download Excel Button
if st.button("Download CAGR Report as Excel"):
    if st.session_state.stock_data_dict:
        excel_file = generate_excel(st.session_state.stock_data_dict, st.session_state.cagr_summary)
        if excel_file:
            st.download_button(
                label="Download Excel File",
                data=excel_file,
                file_name="cagr_calculator_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("No data available to generate Excel.")

# Note on adjusted closing price
st.markdown("---")
st.markdown(
    "**Note:** The adjusted closing price is used for CAGR calculations as it accounts for corporate actions like stock splits and dividends, providing a more accurate measure of investment performance over time."
)
