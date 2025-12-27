# Company Data Cleaning Pipeline (Python)

## Overview
This project provides a **robust and business-safe data cleaning pipeline** built using Python, Pandas, and NumPy.  
It is designed for **real-world company datasets** where data quality issues such as missing values, duplicates, and outliers are common.

The script cleans the data, generates business-friendly summaries, and exports ready-to-use Excel files.

---

## Key Features
- Safely removes duplicate records
- Handles missing values intelligently:
  - Numeric columns → filled with median values
  - Categorical columns → filled with `"Unknown"`
- Automatically processes date columns
- Caps outliers using the IQR (Interquartile Range) method
- Generates both numeric and full descriptive summaries
- Exports clean data and summaries to Excel
- Prints quick validation checks for confidence

---

## Technologies Used
- Python 3
- Pandas
- NumPy
- Excel (.xlsx) files

---

## Input
- **Excel file** containing raw company data  
  Update this line in the script:
  ```python
  df = pd.read_excel("[FILE PATH]")

---

## Excel Data Cleaning & Summary Automation

# Problem :
-- Excel files contain missing values and duplicates

-- Manual cleaning causes mistakes

-- Large datasets make analysis slow

-- Blind deletion of data can harm business decisions

---

Solution :
-- Automatically inspects missing values

-- Applies safe, column-based data handling

-- Preserves business-critical data

-- Generates clean summary reports

-- Creates audit-ready output files

---

##Why this is safe for companies

-- No blind deletion of data

-- Logical null handling

-- Transparent reporting

-- Suitable for large datasets
