# Retail_data_processor

# Retail Operations Reporting System

This project contains an automated Python script designed to consolidate, analyze, and generate performance reports for a retail store network.

## What does this script do?

The script ingests multiple datasets (Excel files) related to sales, store mergers, operational performance, and customer services, and processes them into a unified, comprehensive report.

Specifically, the system handles:
- Sales Aggregation: Calculates total sales, breaking them down into categories like Electronics, Apparel, Home Goods, and Subscriptions.
- Store Mergers Management: Historically re-proportions data when a store (Target) absorbs another, ensuring accurate "Like-for-Like" (YTD vs YTD-1) comparisons.
- Operational KPI Calculation: Computes key metrics such as the Returns/Sales ratio and the adoption rate of Premium Services.
- Workforce and Headcount Analysis: Calculates the number of employees, the ratio between sales staff and administrators, and overall productivity per headcount (Sales/Headcount).
- Customer Analysis: Generates Top 20 city rankings based on customer volume and calculates portfolio concentration indices.

## Input Data Structure

The script expects to find Excel files in specific paths (configurable at the top of the file), divided into the following categories:
1. Base Performance: Current and previous year's sales data.
2. Mergers: The registry file tracking store closures and consolidations.
3. Customer Locations: Order files used to geolocate customers.
4. Premium Services / Upgrades / Promo Index: Detailed files to calculate category-specific metrics.

## Generated Output

Running the script produces a single Excel file (sales_report.xlsx) containing multiple worksheets:
- Sales Report: The consolidated dataset with all stores, absolute metrics, mix percentages, and YoY variations.
- Quintiles (Quintili commerciale): A statistical view (P20, P40, Median, P60, P80) of the entire network's performance.
- Customer Cities (Comuni clienti): Top 20 city tables for each store based on unique customers.
- Summary Views: Condensed sheets highlighting the most critical KPIs and their statistical distribution.

