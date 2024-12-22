import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
from io import BytesIO
import xlsxwriter

# Initialize Streamlit App
st.title("Beta Calculator")
st.sidebar.header("Input Fields")

# Input fields for stock symbols
num_companies = st.sidebar.number_input("Number of companies:", min_value=1, max_value=10, value=2)
stock_symbols = []
for i in range(num_companies):
    symbol = st.sidebar.text_input(f"Enter Stock Symbol {i + 1} (e.g., RELIANCE.NS):", "")
    if symbol:
        stock_symbols.append(symbol)

# Input field for index symbol
index_symbol = st.sidebar.text_input("Enter the Index Symbol (e.g., ^NSEI):", "^NSEI")

# Date input for selecting the start date and number of days to go back
selected_date = st.sidebar.date_input("Select a particular date:", datetime.now().date())
num_days = st.sidebar.number_input("Enter the number of days to go back:", min_value=1, max_value=1000, value=30)

# Initialize session state to store data
if "stock_data_dict" not in st.session_state:
    st.session_state.stock_data_dict = {}
if "beta_summary" not in st.session_state:
    st.session_state.beta_summary = []

# Calculate start date based on selected date
start_date = pd.to_datetime(selected_date) - timedelta(days=num_days)

# Fetch Data Button
if st.sidebar.button("Fetch Data"):
    try:
        stock_data_dict = {}
        beta_summary = []

        for stock_symbol in stock_symbols:
            stock_data = yf.download(stock_symbol, start=start_date, end=selected_date)
            index_data = yf.download(index_symbol, start=start_date, end=selected_date)

            if not stock_data.empty and not index_data.empty:
                stock_data['Daily Change (%)'] = stock_data['Close'].pct_change() * 100
                index_data['Daily Change (%)'] = index_data['Close'].pct_change() * 100

                # Handling intersection of stock and index data
                intersection = pd.merge(
                    stock_data[['Close', 'Daily Change (%)']].reset_index(),
                    index_data[['Close', 'Daily Change (%)']].reset_index(),
                    on='Date',
                    suffixes=("_Stock", "_Index")
                )

                # Handling non-intersection data
                stock_non_intersection = stock_data[~stock_data.index.isin(intersection['Date'])]
                index_non_intersection = index_data[~index_data.index.isin(intersection['Date'])]
                non_intersection = pd.concat(
                    [stock_non_intersection[['Close']].rename(columns={'Close': 'Stock Price'}),
                     index_non_intersection[['Close']].rename(columns={'Close': 'Index Price'})],
                    axis=1
                ).reset_index()

                # Calculating Beta
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

# Display data in Streamlit
if st.session_state.stock_data_dict:
    for stock_symbol, data in st.session_state.stock_data_dict.items():
        st.subheader(f"Data for {stock_symbol}")

        # Display intersection data
        st.write("**Intersection Data**")
        st.dataframe(data["intersection"])

        # Display non-intersection data
        st.write("**Non-Intersection Data**")
        st.dataframe(data["non_intersection"])

        # Display beta
        beta = data["beta"]
        st.write(f"**Beta for {stock_symbol}:** {beta if beta is not None else 'Insufficient Data'}")

    # Display beta summary
    st.subheader("Beta Summary")
    st.table(pd.DataFrame(st.session_state.beta_summary))

# Function to generate Excel file with improved styling
def generate_excel(stock_data_dict, beta_summary):
    try:
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True, 'nan_inf_to_errors': True})

        # Define formatting styles
        header_format = workbook.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1, 'align': 'center'})
        cell_format = workbook.add_format({'border': 1, 'align': 'center'})
        percent_format = workbook.add_format({'num_format': '0.00%', 'border': 1, 'align': 'center'})
        bold_format = workbook.add_format({'bold': True})
        date_format = workbook.add_format({'num_format': 'yyyy-mm-dd', 'border': 1, 'align': 'center'})

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
                worksheet.write(3, col, header, header_format)

            for row, record in enumerate(intersection.itertuples(index=False), start=4):
                for col, value in enumerate(record):
                    if col == 0:  # Apply date format for the first column
                        worksheet.write_datetime(row, col, pd.to_datetime(value), date_format)
                    elif col in [2, 4]:  # Apply percentage format for change columns
                        worksheet.write_number(row, col, value / 100 if value != "" else 0, percent_format)
                    else:
                        worksheet.write(row, col, value, cell_format)

            # Write non-intersection data
            start_row = len(intersection) + 6
            worksheet.write(start_row, 0, "Non-Intersection Data", bold_format)
            headers = ["Date", "Stock Price", "Index Price"]
            for col, header in enumerate(headers):
                worksheet.write(start_row + 1, col, header, header_format)

            for row, record in enumerate(non_intersection.itertuples(index=False), start=start_row + 2):
                for col, value in enumerate(record):
                    if col == 0:  # Apply date format for the first column
                        worksheet.write_datetime(row, col, pd.to_datetime(value), date_format)
                    else:
                        worksheet.write(row, col, value, cell_format)

            # Autofit columns
            for col_num in range(len(headers)):
                worksheet.set_column(col_num, col_num, 18)  # Adjust column width to fit data

        # Create a summary sheet
        summary_sheet = workbook.add_worksheet("Summary")
        summary_sheet.write(0, 0, "Stock Symbol", header_format)
        summary_sheet.write(0, 1, "Beta", header_format)
        for row, beta_data in enumerate(beta_summary, start=1):
            summary_sheet.write(row, 0, beta_data["Stock Symbol"], cell_format)
            summary_sheet.write(row, 1, beta_data["Beta"], cell_format)

        # Write average beta
        summary_sheet.write(len(beta_summary) + 2, 0, "Average Beta", bold_format)
        average_beta = sum(d['Beta'] for d in beta_summary if d['Beta'] is not None) / len(beta_summary)
        summary_sheet.write(len(beta_summary) + 2, 1, round(average_beta, 2), cell_format)

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