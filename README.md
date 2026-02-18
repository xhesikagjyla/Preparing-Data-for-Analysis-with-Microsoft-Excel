# Preparing-Data-for-Analysis-with-Microsoft-Excel

# Project Summary

This project transforms raw transactional sales data into an executive-ready performance report using Microsoft Excel.

The goal was to analyze Quarter 1 sales performance for two consecutive years (2022 vs 2023), calculate financial KPIs, and present structured insights suitable for senior management decision-making.

This project demonstrates applied business analytics, financial modeling logic, and structured data transformation.

# Business Scenario

A management team required a Q1 executive summary including:

	-Total sales for 2022 and 2023
	-Monthly breakdown (January–March)
	-Year-over-year growth rates
	-Order-level tax logic for high-value transactions
  
The challenge was to restructure raw data into meaningful performance metrics.

# Tools & Techniques

	-Microsoft Excel
	-Conditional logic (IF)
	-Aggregation functions (SUMIFS)
	-Date feature engineering (MONTH, YEAR)
	-Financial KPI calculations
	-Data transformation & formatting

# Dataset Overview

The dataset contained:

	-Product ID
	-Product Name
	-Wholesale Price
	-Retail Price
	-Order Quantity
	-Order Date
	-Sales Value
  
The raw structure required cleaning and restructuring before analysis.

# Analytical Workflow

# 1. Data Cleaning & Standardization

Product names were standardized using:
```excel
=PROPER(G2)
```

This ensured consistent formatting across the dataset and improved presentation quality.

After transformation, formulas were replaced with static values to prevent dependency errors.

# 2. Feature Engineering

To enable time-based aggregation:
```excel
=MONTH(J2)
=YEAR(J2)
```

This created structured columns for:

	-Month (1–3 for Q1)
	-Year (2022 / 2023)
  
These helper features allowed scalable conditional aggregation without manual filtering.

# 3. Revenue Modeling

Each order’s revenue was calculated using:
```excel
=M2 * N2
```
This created the base financial metric used for tax logic, annual aggregation and monthly performance comparison.

# 4. Business Rule Implementation 

Orders above €2000 were subject to 5% tax:
```excel
=IF(P2>2000, P2*5%, 0)
```

# 5. Aggregation & KPI Construction

🔹 Total Q1 Sales by Year
```excel
=SUMIFS(R2:R246, L2:L246, 2022)
=SUMIFS(R2:R246, L2:L246, 2023)
```
Using conditional aggregation enabled structured yearly comparisons.

🔹 Monthly Breakdown (January–March)
Example (January 2022):
```excel
=SUMIFS($R$2:$R$103, $K$2:$K$103, 1)
```
This shows:

	-Advanced use of SUMIFS
	-Absolute referencing for scalability
	-Structured monthly segmentation
  
# 6. Year-over-Year Growth Calculation

Core KPI formula:
```excel
=(C6 - B6) / B6
```
This calculates percentage growth from 2022 to 2023.

Mathematically:
```excel
Growth=(New-Old)/Old
```

# Key Results

| Metric            | 2022      | 2023      | YoY Growth |
|-------------------|-----------|-----------|------------|
| Total Q1 Sales   | $330,500  | $453,830  | **+37.32%** |
| January          | $101,595  | $143,555  | **+41.30%** |
| February         | $113,445  | $145,535  | **+28.29%** |
| March            | $115,460  | $164,740  | **+42.68%** |


# Insights

	-Strong overall Q1 growth (+37%)
	-March shows highest performance increase
	-Consistent improvement across all months
	-Tax rule correctly applied to high-value orders

# Executive Reporting Enhancements

To improve readability and presentation quality:

	-Headings formatted with Merge & Center
	-Wrap text for structured headers
	-Freeze Panes for usability
	-Hidden irrelevant columns
	-Sorted chronologically for structured aggregation



