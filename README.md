
# **Beta Calculator App**

The **Beta Calculator App** is a powerful, interactive tool built with **Streamlit** for calculating the Beta of a stock relative to a benchmark index. Designed for financial analysts, investors, and CFA candidates, it provides an intuitive way to assess systematic risk with exportable results.

---

## **Features**

- 📈 **Stock Beta Calculation**: Analyze a stock's volatility relative to a benchmark index.
- 📅 **Custom Date Range**: Select specific time periods or rolling windows for analysis.
- 🧹 **Data Cleaning**: Automatically handles non-overlapping stock and index data points.
- 📤 **Excel Export**: Generate well-formatted Excel reports for further analysis.
- ⚡ **Fast and Interactive**: Built on Streamlit for seamless user experience.

---

## **How to Use**

1. Input the **stock ticker symbol** (e.g., `AAPL`) and **benchmark index symbol** (e.g., `^GSPC` for S&P 500).
2. Select a date range or specify a rolling period for analysis.
3. Click **Calculate** to compute the Beta.
4. Export the results to an Excel file for professional use.

---

## **Installation**

Follow these steps to run the app locally:

1. Clone this repository:
   ```bash
   git clone https://github.com/Vedant-try/Veta.git
   cd Veta
   ```

2. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Launch the Streamlit app:
   ```bash
   streamlit run Eg.py
   ```

---

## **Requirements**

- Python 3.9 or later  
- Dependencies (listed in `requirements.txt`):
  - `streamlit`
  - `pandas`
  - `numpy`
  - `yfinance`
  - `xlsxwriter`

---

## **Project Structure**

```
Veta/
├── Eg.py                # Main Streamlit app file
├── requirements.txt     # Python dependencies
├── README.md            # Project documentation
```

---

## **Contributing**

Contributions are welcome! Follow these steps to contribute:
1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Commit your changes (`git commit -m "Add feature"`).
4. Push to the branch (`git push origin feature-branch`).
5. Open a Pull Request.

---

## **License**

This project is licensed under the [MIT License](LICENSE).

---

## **Author**

Developed by **Vedant Shah**, a CFA student passionate about finance, statistics, and workflow automation.
