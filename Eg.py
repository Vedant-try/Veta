import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime, timedelta
from io import BytesIO
import xlsxwriter
import matplotlib.pyplot as plt
from matplotlib.dates import DateFormatter, MonthLocator, YearLocator

# Initialize Streamlit App
st.title("Beta Calculator")
st.sidebar.header("Input Fields")

# Note for tickers
st.sidebar.markdown("**Note:** Take tickers from Yahoo Finance (e.g., RELIANCE.NS, ^NSEI).")

# Input fields for stock symbols
num_companies = st.sidebar.number_input("Number of companies:", min_value=1, max_value=10, value=2)
stock_symbols = []
for i in range(num_companies):
    symbol = st.sidebar.text_input(f"Enter Stock Symbol {i + 1} (e.g., RELIANCE.NS):", "")
    if symbol:
        stock_symbols.append(symbol)

# Input field for index symbol
index_symbol = st.sidebar.text_input("Enter the Index Symbol (e.g., ^NSEI):", "^NSEI")

# Date inputs for selecting start and end dates
start_date = st.sidebar.date_input(
    "Select Start Date (dd/mm/yyyy):",
    datetime.now().date() - timedelta(days=30),
    min_value=datetime.now().date() - timedelta(days=5 * 365),
    max_value=datetime.now().date(),
).strftime('%d/%m/%Y')

end_date = st.sidebar.date_input(
    "Select End Date (dd/mm/yyyy):",
    datetime.now().date(),
    min_value=datetime.strptime(start_date, '%d/%m/%Y'),
    max_value=datetime.now().date(),
).strftime('%d/%m/%Y')

start_date = datetime.strptime(start_date, '%d/%m/%Y').date()
end_date = datetime.strptime(end_date, '%d/%m/%Y').date()

# Initialize session state to store data
if "stock_data_dict" not in st.session_state:
    st.session_state.stock_data_dict = {}
if "beta_summary" not in st.session_state:
    st.session_state.beta_summary = []

# Definition and formula for beta
st.markdown("### Understanding Stock Beta")
st.markdown(
    r"""
    **Beta (β)** is a measure of the volatility or systematic risk of a stock compared to the overall market. 
    A beta value helps investors understand how sensitive a stock is to market movements:
    - β > 1: The stock is more volatile than the market.
    - β < 1: The stock is less volatile than the market.
    - β = 1: The stock moves in sync with the market.

    The formula for beta is:
    """
)
st.latex(r"\beta = \frac{\text{Cov}(R_{\text{stock}}, R_{\text{index}})}{\text{Var}(R_{\text{index}})}")

# Fetch Data Button
if st.sidebar.button("Fetch Data"):
    try:
        stock_data_dict = {}
        beta_summary = []

        for stock_symbol in stock_symbols:
            stock_data = yf.download(stock_symbol, start=start_date, end=end_date)
            index_data = yf.download(index_symbol, start=start_date, end=end_date)

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

        # Display line graph for percentage changes
        st.write("**Volatility Comparison: Stock vs. Index**")
        plt.figure(figsize=(10, 6))
        plt.plot(data["intersection"]['Date'], data["intersection"]['Daily Change (%)_Stock'], label="Stock")
        plt.plot(data["intersection"]['Date'], data["intersection"]['Daily Change (%)_Index'], label="Index")
        plt.xlabel("Date")
        plt.ylabel("Daily Change (%)")
        plt.title(f"Daily Change: {stock_symbol} vs. {index_symbol}")

        # Improve date formatting
        ax = plt.gca()
        if (end_date - start_date).days > 365:
            ax.xaxis.set_major_locator(YearLocator())
            ax.xaxis.set_major_formatter(DateFormatter("%Y"))
        else:
            ax.xaxis.set_major_locator(MonthLocator(interval=6))
            ax.xaxis.set_major_formatter(DateFormatter("%b %y"))

        plt.legend()
        st.pyplot(plt)

    # Display beta summary
    st.subheader("Beta Summary")
    beta_df = pd.DataFrame(st.session_state.beta_summary)
    average_beta = beta_df['Beta'].mean()
    
    # Replacing append() with pd.concat()
    average_beta_df = pd.DataFrame([{"Stock Symbol": "Average", "Beta": round(average_beta, 2)}])
    beta_df = pd.concat([beta_df, average_beta_df], ignore_index=True)
    
    st.table(beta_df)

# Function to generate Excel file with improved styling and graphs
def generate_excel(stock_data_dict, beta_summary):
    try:
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True, 'nan_inf_to_errors': True})

        # Define formatting styles
        header_format = workbook.add_format({'bold': True, 'bg_color': '#DDEBF7', 'border': 1, 'align': 'center'})
        cell_format = workbook.add_format({'border': 1, 'align': 'center'})
        percent_format = workbook.add_format({'num_format': '0.00%', 'border': 1, 'align': 'center'})
        bold_format = workbook.add_format({'bold': True})
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy', 'border': 1, 'align': 'center'})

        # Create individual sheets for each stock
        for stock_symbol, data in stock_data_dict.items():
            worksheet = workbook.add_worksheet(stock_symbol[:31])  # Sheet names max 31 chars
            intersection = data['intersection']
            beta = data['beta']

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

            # Auto-fit column widths
            for col_num, _ in enumerate(headers):
                worksheet.set_column(col_num, col_num, 15)

            # Add a graph to the Excel sheet
            chart = workbook.add_chart({'type': 'line'})
            chart.add_series({
                'name': f'{stock_symbol} Daily Change (%)',
                'categories': [stock_symbol[:31], 4, 0, len(intersection) + 3, 0],
                'values': [stock_symbol[:31], 4, 2, len(intersection) + 3, 2],
            })
            chart.add_series({
                'name': 'Index Daily Change (%)',
                'categories': [stock_symbol[:31], 4, 0, len(intersection) + 3, 0],
                'values': [stock_symbol[:31], 4, 4, len(intersection) + 3, 4],
            })
            chart.set_title({'name': f'Daily Change: {stock_symbol} vs. {index_symbol}'})
            chart.set_x_axis({'name': 'Date', 'date_axis': True})
            chart.set_y_axis({'name': 'Daily Change (%)'})
            worksheet.insert_chart(len(intersection) + 5, 0, chart)

        # Create a summary sheet
        summary_sheet = workbook.add_worksheet("Summary")
        summary_sheet.write(0, 0, "Stock Symbol", header_format)
        summary_sheet.write(0, 1, "Beta", header_format)
        for row, beta_data in enumerate(beta_summary, start=1):
            summary_sheet.write(row, 0, beta_data["Stock Symbol"], cell_format)
            summary_sheet.write(row, 1, beta_data["Beta"], cell_format)

        # Add average beta at the bottom
        summary_sheet.write(len(beta_summary) + 1, 0, "Average", bold_format)
        summary_sheet.write(len(beta_summary) + 1, 1, round(average_beta, 2), bold_format)

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

# Note on adjusted closing price
st.markdown("---")
st.markdown(
    "**Note:** The adjusted closing price is used instead of the regular closing price because it accounts for corporate actions like stock splits and dividends, making it more accurate for calculating beta."
)
