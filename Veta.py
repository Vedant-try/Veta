import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
from io import BytesIO
import xlsxwriter

# Initialize Streamlit App
st.title("Beta Calculator")
st.sidebar.header("Input Fields")

# Input fields
stock_symbols = st.sidebar.text_area("Enter Stock Symbols (comma-separated, e.g., RELIANCE.NS,TCS.NS):", "RELIANCE.NS,TCS.NS")
index_symbol = st.sidebar.text_input("Enter the Index Symbol (e.g., ^NSEI):", "^NSEI")
num_days = st.sidebar.number_input("Enter the number of days (max 1000):", min_value=1, max_value=1000, value=30)

if "stock_data_dict" not in st.session_state:
    st.session_state.stock_data_dict = {}
if "beta_summary" not in st.session_state:
    st.session_state.beta_summary = []

# Fetch Data Button
if st.sidebar.button("Fetch Data"):
    try:
        stock_symbols_list = [symbol.strip() for symbol in stock_symbols.split(",")]
        end_date = datetime.now().date()
        start_date = end_date - timedelta(days=num_days)

        stock_data_dict = {}
        beta_summary = []

        for stock_symbol in stock_symbols_list:
            stock_data = yf.download(stock_symbol, start=start_date, end=end_date)
            index_data = yf.download(index_symbol, start=start_date, end=end_date)

            if not stock_data.empty and not index_data.empty:
                stock_data['Daily Change (%)'] = stock_data['Close'].pct_change() * 100
                index_data['Daily Change (%)'] = index_data['Close'].pct_change() * 100

                intersection = pd.merge(
                    stock_data[['Close', 'Daily Change (%)']].reset_index(),
                    index_data[['Close', 'Daily Change (%)']].reset_index(),
                    on='Date',
                    suffixes=("_Stock", "_Index")
                )

                non_intersection = pd.concat([stock_data, index_data], axis=0).drop_duplicates(keep=False)

                covariance = intersection['Daily Change (%)_Stock'].cov(intersection['Daily Change (%)_Index'])
                variance = intersection['Daily Change (%)_Index'].var()
                beta = round(covariance / variance, 2) if variance != 0 else None

                stock_data_dict[stock_symbol] = {
                    "intersection": intersection,
                    "non_intersection": non_intersection,
                    "beta": beta
                }

                beta_summary.append({"Stock Symbol": stock_symbol, "Beta": beta})

        st.session_state.stock_data_dict = stock_data_dict
        st.session_state.beta_summary = beta_summary

        st.success("Data fetched successfully!")
    except Exception as e:
        st.error(f"An error occurred: {e}")

# Display Data Button
if st.button("Display Data"):
    if st.session_state.stock_data_dict:
        for stock_symbol, data in st.session_state.stock_data_dict.items():
            st.subheader(f"Data for {stock_symbol}")
            st.write("Intersection Data")
            st.dataframe(data['intersection'])
            st.write("Non-Intersection Data")
            st.dataframe(data['non_intersection'])
            st.write(f"Beta: {data['beta']}")
    else:
        st.error("No data available. Please fetch data first.")

# Function to generate Excel file with improved styling
def generate_excel(stock_data_dict, beta_summary):
    try:
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})

        # Define formatting styles
        bold_format = workbook.add_format({'bold': True})
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})  # Short date format
        table_header_format = workbook.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1})
        cell_border_format = workbook.add_format({'border': 1})

        # Create individual sheets for each stock
        for stock_symbol, data in stock_data_dict.items():
            worksheet = workbook.add_worksheet(stock_symbol[:31])  # Sheet names max 31 chars
            intersection = data['intersection']
            non_intersection = data['non_intersection']
            beta = data['beta']

            # Replace NaN/Inf with None
            intersection = intersection.fillna("").replace([float("inf"), float("-inf")], "")
            non_intersection = non_intersection.fillna("").replace([float("inf"), float("-inf")], "")

            # Write beta value at the top
            worksheet.write(0, 0, f"Beta for {stock_symbol}: {beta if beta is not None else 'Insufficient Data'}", bold_format)

            # Write intersection data
            worksheet.write(2, 0, "Intersection Data", bold_format)
            headers = ["Date", "Stock Price", "Stock Daily Change (%)", "Index Price", "Index Daily Change (%)"]
            for col, header in enumerate(headers):
                worksheet.write(3, col, header, table_header_format)

            for row, record in enumerate(intersection.itertuples(index=False), start=4):
                for col, value in enumerate(record):
                    if col == 0:  # Apply date format for the first column
                        worksheet.write_datetime(row, col, pd.to_datetime(value), date_format)
                    else:
                        worksheet.write(row, col, value, cell_border_format)

            # Write non-intersection data
            start_row = len(intersection) + 6
            worksheet.write(start_row, 0, "Non-Intersection Data", bold_format)
            headers = ["Date", "Stock Price", "Index Price"]
            for col, header in enumerate(headers):
                worksheet.write(start_row + 1, col, header, table_header_format)

            for row, record in enumerate(non_intersection.itertuples(index=False), start=start_row + 2):
                for col, value in enumerate(record):
                    if col == 0:  # Apply date format for the first column
                        worksheet.write_datetime(row, col, pd.to_datetime(value), date_format)
                    else:
                        worksheet.write(row, col, value, cell_border_format)

            # Autofit columns
            for col_num, _ in enumerate(headers):
                worksheet.set_column(col_num, col_num, 18)  # Adjust column width to fit data

        # Create a summary sheet
        summary_sheet = workbook.add_worksheet("Summary")
        summary_sheet.write(0, 0, "Stock Symbol", table_header_format)
        summary_sheet.write(0, 1, "Beta", table_header_format)
        for row, beta_data in enumerate(beta_summary, start=1):
            summary_sheet.write(row, 0, beta_data["Stock Symbol"], cell_border_format)
            summary_sheet.write(row, 1, beta_data["Beta"], cell_border_format)

        # Write average beta
        summary_sheet.write(len(beta_summary) + 2, 0, "Average Beta", bold_format)
        average_beta = sum(d['Beta'] for d in beta_summary if d['Beta'] is not None) / len(beta_summary)
        summary_sheet.write(len(beta_summary) + 2, 1, round(average_beta, 2), cell_border_format)

        # Autofit columns
        summary_sheet.set_column(0, 0, 18)  # Column for Stock Symbol
        summary_sheet.set_column(1, 1, 12)  # Column for Beta

        workbook.close()
        output.seek(0)
        return output
    except Exception as e:
        st.error(f"An error occurred while generating the Excel file: {e}")
        return None

# Download Workings Button
if st.button("Download Workings as Excel"):
    if st.session_state.stock_data_dict:
        excel_file = generate_excel(st.session_state.stock_data_dict, st.session_state.beta_summary)
        if excel_file:
            st.download_button(
                label="Download Excel File",
                data=excel_file,
                file_name="beta_calculator_workings.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.error("No data available to generate Excel.")
