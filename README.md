# Product Catalog Data Entry & QC

A small project that transfers product records from a CSV file into a
structured Excel catalog, **keeping barcodes as text (leading zeros preserved)**
and running basic quality-control (QC) checks.

## What it does
- Reads the source CSV with all columns as text (no lost leading zeros)
- Writes a formatted Excel catalog (Product ID, Barcode, Name, Brand, Category, Quantity, Country)
- Adds a QC sheet with live Excel formulas:
  - total records
  - duplicate barcodes
  - missing values
  - barcode length check (13 digits)

## Project structure
```
data/source_products.csv     raw input (15 demo records)
data/product_catalog.xlsx    final output (Catalog + QC sheets)
clean_catalog.py             the script
qc_report.md                 QC summary
```

## Run
```
pip install pandas openpyxl
python clean_catalog.py
```

## Tools
Python, Pandas, openpyxl, Excel

## Data
All records are demo data created for this project. No third-party data used.
