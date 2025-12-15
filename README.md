# Jumbo & Company – Attach Percentage Analysis

This project analyzes device insurance attach percentage data for Jumbo & Company retail stores.

## Objective
- Analyze attach % across months, branches, and stores
- Categorize store performance
- Predict January attach % using historical trends

## Data
- Monthly attach % data from Aug–Dec
- January data predicted using 3-month moving average

## Methodology
- Data transformation (wide to long format)
- Exploratory data analysis using Python
- Store categorization: High / Medium / Low performers
- January prediction using moving average

## Tools Used
- Python
- Pandas, NumPy
- Matplotlib, Seaborn

## How to Run
```bash
pip install pandas numpy matplotlib seaborn xlrd openpyxl
python zopper_attach_analysis.py
