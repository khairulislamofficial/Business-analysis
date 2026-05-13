# 📊 Sales & Customer Analysis Dashboard

A premium, interactive business intelligence dashboard built with **Streamlit**, **Pandas**, and **Plotly**. This tool allows business owners to upload their sales data, automatically clean it, and gain deep insights into their performance, customer behavior, and delivery metrics.

![Dashboard Mockup](assets/dashboard_mockup.png)

## 🌟 Key Features

-   **📁 Intelligent Data Loading**: Support for Excel files with automatic column mapping and cleaning.
-   **📈 Real-time KPIs**: Instant tracking of Total Revenue, Order Volume, Average Order Value, and Return Rates.
-   **📅 Trend Analysis**: Visualize sales performance over time (Monthly, Weekly, and Day of Week).
-   **🏆 Product Intelligence**: Identify top-selling products and slow-moving items to optimize inventory.
-   **💳 Payment & Delivery Insights**: Analyze popular payment methods and monitor delivery success vs. returns.
-   **🎯 Actionable Recommendations**: Automated business advice based on data patterns.

## 🛠️ Tech Stack

-   **Core**: Python 3.x
-   **Frontend**: [Streamlit](https://streamlit.io/)
-   **Data Processing**: [Pandas](https://pandas.pydata.org/)
-   **Visualization**: [Plotly](https://plotly.com/python/), [Seaborn](https://seaborn.pydata.org/), [Matplotlib](https://matplotlib.org/)
-   **Data Handling**: [OpenPyXL](https://openpyxl.readthedocs.io/), [XlsxWriter](https://xlsxwriter.readthedocs.io/)

## 🚀 Getting Started

### 1. Prerequisites
Ensure you have Python installed. Then, install the required dependencies:

```bash
pip install -r requirements.txt
```

### 2. Generate Sample Data (Optional)
If you don't have your own data yet, you can generate a realistic sample dataset:

```bash
python generate_sample_data.py
```
This will create a `data/sample_sales.xlsx` file.

### 3. Run the Application
Start the Streamlit dashboard:

```bash
streamlit run app.py
```

## 📋 Expected Data Format
The application expects an Excel file with the following columns (it will attempt to auto-map similar names):
- Order ID
- Order Date
- Customer Name
- Customer Phone
- Product Name
- Quantity
- Unit Price (BDT)
- Total Revenue (BDT)
- Payment Method
- Delivery Status

## 📁 Project Structure

```text
Business-analysis/
├── assets/               # Dashboard screenshots and mockups
├── data/                 # Data storage (Excel files)
├── app.py                # Main Streamlit application
├── generate_sample_data.py # Script to create mock data
├── requirements.txt      # Python dependencies
└── README.md             # Project documentation
```

---
Built with ❤️ for Data-Driven Businesses.
