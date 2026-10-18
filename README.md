# Automated Reporting System Using Python

## Project Overview

The **Automated Reporting System** is a Python-based project that processes retail sales data and automatically generates a formatted Excel report. It uses Pandas for data processing and OpenPyXL to create, format, and visualize reports.

This project demonstrates how repetitive reporting tasks can be automated to save time, reduce manual effort, and improve reporting consistency.

## Features

- Loads retail sales data from a CSV file.
- Validates required columns and checks the input file.
- Calculates key sales metrics, including total revenue, transaction count, units sold, and average transaction value.
- Performs product category, monthly sales, and gender-based analysis.
- Generates a multi-sheet Excel workbook.
- Formats worksheets with styled headers, adjusted column widths, and frozen header rows.
- Creates a bar chart for revenue by product category and a line chart for monthly revenue trends.
- Records report-generation activity in a log file.

## Technologies Used

- **Python** — Core programming language
- **Pandas** — Data processing and analysis
- **OpenPyXL** — Excel report creation, formatting, and charts
- **pathlib** — File and folder management
- **CSV and Excel** — Data input and report output

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
├── .gitignore
└── README.md
```

## Installation and Setup

### 1. Clone the repository

```bash
git clone https://github.com/iamsaqb/Automated-Reporting-System.git
```

### 2. Navigate to the project folder

```bash
cd Automated-Reporting-System
```

### 3. Install the required dependencies

```bash
python -m pip install -r requirements.txt
```

### 4. Prepare the input data

Ensure that `retail_sales_dataset.csv` is available inside the `input_files` folder.

## How to Run

Open a terminal in the project directory and execute:

```bash
python automated_report.py
```

After successful execution, the generated Excel report will be saved in the `output_reports` folder. Report-generation activity is recorded in the `logs` folder.

## Generated Excel Report

The generated workbook contains the following worksheets:

| Worksheet | Description |
|---|---|
| Sales Summary | Key sales performance metrics |
| Category Analysis | Revenue analysis by product category |
| Monthly Analysis | Monthly revenue trends |
| Gender Analysis | Revenue, transaction count, and average transaction value by gender |

The workbook also includes charts to help visualize category revenue and monthly sales trends.

## Skills Demonstrated

- Data processing and analysis using Pandas
- Data validation and error handling
- Excel report generation and formatting
- Data aggregation and business metric calculation
- Chart creation using OpenPyXL
- File and folder management using pathlib
- Python dependency management
- Logging and reporting automation
- Git and GitHub version control

## Future Improvements

- Schedule automatic report generation.
- Add email delivery of completed reports.
- Introduce additional KPIs and visualizations.
- Support multiple input datasets.
- Add configurable reporting periods.

## Author

**Saquib Akhter**

- GitHub: [@iamsaqb](https://github.com/iamsaqb)
- Project Repository: [Automated Reporting System](https://github.com/iamsaqb/Automated-Reporting-System)

## Note

The sample retail sales dataset is required to run this project. If redistributing the dataset, ensure that you have permission to share it.
