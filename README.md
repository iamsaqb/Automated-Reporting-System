# Automated Reporting System using Python

## Project Overview
The Automated Reporting System is a Python-based project that processes retail sales data and automatically generates a formatted Excel report. It uses Pandas for data processing and OpenPyXL to create, format, and visualize reports.

This project demonstrates how repetitive reporting tasks can be automated to save time and improve reporting consistency.

## Features
- Loads retail sales data from a CSV file.
- Validates required columns and checks the input file.
- Calculates key sales metrics, including total revenue, transactions, units sold, and average transaction value.
- Generates product category, monthly sales, and gender-based analyses.
- Exports results into a multi-sheet Excel workbook.
- Formats worksheets with styled headers, adjusted column widths, and frozen header rows.
- Adds a bar chart for revenue by product category and a line chart for monthly revenue trends.
- Records report-generation activity in a log file.

## Technologies Used
- Python
- Pandas
- OpenPyXL
- pathlib
- CSV and Excel

## Project Structure
```text
Automated_Reporting_System/
├── input_files/
│   └── retail_sales_dataset.csv
├── output_reports/
│   └── Automated_Sales_Report.xlsx
├── logs/
│   └── report.log
├── automated_report.py
├── automated_report.ipynb
├── requirements.txt
└── README.md
```

## Installation and Setup

1. Clone or download this repository.
2. Open the project folder in VS Code.
3. Install the required packages:

```bash
python -m pip install -r requirements.txt
```

4. Place `retail_sales_dataset.csv` inside the `input_files` folder.

## How to Run

Open the terminal in the project folder and run:

```bash
python automated_report.py
```

After successful execution, the Excel report will be saved in the `output_reports` folder. The application log is saved in the `logs` folder.

## Generated Report

The Excel workbook contains the following sheets:

| Worksheet | Description |
|---|---|
| Sales Summary | Key sales performance metrics |
| Category Analysis | Revenue by product category |
| Monthly Analysis | Monthly revenue trends |
| Gender Analysis | Revenue, transaction count, and average transaction value by gender |

The report also includes Excel charts for category revenue and monthly revenue trends.

## Skills Demonstrated
- Data processing with Pandas
- Data validation and error handling
- Excel report generation and formatting
- Data aggregation and analysis
- Chart creation using OpenPyXL
- File and folder management with pathlib
- Dependency management and logging

## Future Improvements
- Schedule automatic report generation.
- Add email delivery of completed reports.
- Include additional KPIs and visualizations.
- Support multiple input datasets.

## Author
Saquib Akhter

## Note
The sample retail sales dataset is required to run the project. If sharing the project publicly, ensure that you have permission to redistribute the dataset.
