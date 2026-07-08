# Data Cleaning & Reporting Automation

Automates the cleaning of messy sales data and produces a polished, chart-rich Excel report — built with Python (pandas + openpyxl + matplotlib).

## What it does

1. **Loads** a raw CSV that has real-world problems: missing values, duplicate rows, inconsistent casing/whitespace, mixed date formats, and invalid (negative) numbers.
2. **Cleans** it automatically:
   - Drops blank rows and exact duplicates
   - Standardizes text (trims whitespace, fixes casing on Region/Product)
   - Parses mixed date formats into one consistent format
   - Corrects invalid (negative) quantities
   - Fills missing numeric values with the column median (documented in the log)
   - Drops rows missing essential identifiers
   - Recomputes `Revenue = Units x UnitPrice` so totals are trustworthy
3. **Analyzes** the cleaned data: totals, revenue by region/product/sales rep, and monthly trend.
4. **Reports**: generates `Sales_Data_Report.xlsx` with a Summary tab (embedded charts + cleaning log), breakdown tabs, a native Excel bar chart, and both the cleaned and original raw data for auditability.

## Files

- `generate_sample_data.py` – creates a realistic messy dataset (only needed to demo the project; skip if you have your own data)
- `clean_and_report.py` – the automation itself
- `data/raw_sales_data.csv` – sample messy input
- `output/cleaned_sales_data.csv` – cleaned output
- `output/Sales_Data_Report.xlsx` – final automated report
- `output/*.png` – standalone chart images

## How to run it on your own data

```bash
python clean_and_report.py path/to/your_data.csv path/to/output_folder
```

Your CSV should ideally include columns similar to: `OrderID, Date, Region, Product, SalesRep, Units, UnitPrice`. If your column names differ, adjust the column references near the top of `clean_data()` in `clean_and_report.py`.

## Extending this project

- Swap in `pd.read_excel()` / `pd.read_sql()` to pull from other sources
- Add more cleaning rules (e.g., outlier detection, category mapping) in `clean_data()`
- Add more breakdowns/charts in `summarize()` and `build_excel_report()`
- Schedule it with `cron` or Windows Task Scheduler to re-run automatically each time new data lands, turning this into a fully hands-off reporting pipeline

## Expected outcome mapped to task requirements

| Requirement | How it's addressed |
|---|---|
| Use Python for automation | `clean_and_report.py` is a single reusable script |
| Handle missing values, duplicates, inconsistent data | Median-fill, dedup, text/date standardization, invalid-value correction (all logged) |
| Generate automated reports and visual summaries | Multi-sheet Excel report with embedded + native charts, generated in one run |
| Understand data preprocessing / reporting efficiency | README + inline cleaning log make every transformation traceable |
