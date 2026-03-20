
# Excel Sales Report Automation with Python

This project demonstrates how to automate the process of cleaning messy sales data and generating a structured Excel report using Python.

The script reads a raw CSV dataset containing common data issues (duplicates, missing values), cleans the data, and automatically generates an Excel report with summarized sales information.

## Project Objective

Many companies store raw sales data in CSV or Excel files that contain errors such as:

- Missing values
- Duplicate rows
- Inconsistent data

This project shows how Python can automate the cleaning process and produce a structured report ready for analysis.

## Technologies Used

- Python
- Pandas
- CSV data processing
- Excel automation

## Project Structure
  excel-report-automation
  │
  ├── data
  │ └── sales_raw.csv
  │
  ├── output
  │ └── sales_report.xlsx
  │
  ├── main.py
  └── README.md

## Dataset Description

The dataset contains sales records with the following fields:

| Column | Description |
|------|------|
| date | date of the sale |
| product | product name |
| price | product price |
| quantity | number of units sold |

The raw dataset intentionally includes:

- missing values
- duplicated records
- inconsistent entries

to simulate real-world business data.

## Data Cleaning Steps

The script performs the following operations:

1. Remove duplicate records
2. Fill missing prices using the most common value (mode) within the same product
3. If the mode is unavailable, the median is used
4. Apply the same logic to the quantity column
5. Convert data types for correct processing
6. Create a new column called **total** (price × quantity)

## Report Generation

The script generates an Excel file with two sheets:

### Cleaned Data

Contains the corrected dataset ready for analysis.

### Summary

Aggregated statistics by product:

- Total sales
- Total quantity sold
- Average price

Example:

| product | total_sales | total_quantity | avg_price |
|------|------|------|------|
| laptop | 4000 | 4 | 1000 |
| mouse | 80 | 8 | 10 |

## How to Run the Project

Install dependencies:
pip install pandas openpyxl

Run the script:
python main.py

The report will be generated automatically in:
output/sales_report.xlsx


## Example Use Cases

This type of automation is useful for:

- Sales reporting
- Data cleaning pipelines
- Excel report automation
- Business data preparation

## Author
Jonathan Bejarano
GitHub:
https://github.com/jonatanbejarano
