# Retail Operations Reporting System

This project contains an automated Python script designed to consolidate, analyze, and generate performance reports for a retail store network.

## What does this script do?

The script ingests multiple datasets (Excel files) related to sales, store mergers, operational performance, and customer services, and processes them into a unified, comprehensive report.

Specifically, the system handles:
- Sales Aggregation: Calculates total sales, breaking them down into categories like Electronics, Apparel, Home Goods, Subscriptions, One-Time, and B2B.
- Store Mergers Management: Historically re-proportions data when a store (Target) absorbs another, ensuring accurate Like-for-Like (YTD vs YTD-1) comparisons using both additive and average-based logic.
- Operational KPI Calculation: Computes key metrics such as the Returns/Sales ratio, Promo Index, Upgrade and Renewal rates, and the adoption rate of Premium Services.
- Workforce and Headcount Analysis: Calculates the number of employees across different roles (admins, leads, ops), the ratio between sales staff and administrators, and overall productivity per headcount.
- Customer Analysis: Generates Top 20 city rankings based on customer volume and calculates portfolio concentration indices based on top customers and order managers.

## Input Data Structure

The script expects to find Excel files in specific paths (configurable at the top of the file), divided into the following categories:
1. Base Performance: Current and previous year's sales, new customers, and headcount data.
2. Mergers: The registry file tracking store closures and consolidations.
3. Customer Locations: Order files used to geolocate customers and calculate concentration metrics.
4. Store Upgrades and Expiring Contracts: Files tracking physical store upgrades and subscription renewals.
5. Premium Services and Promo Index: Detailed files to calculate category-specific operational metrics.

## Generated Output

Running the script produces a single Excel file (sales_report.xlsx) containing five specific worksheets:
- Sales Report: The consolidated dataset with all stores, absolute metrics, mix percentages, YoY variations, and productivity ratios.
- Sales Quintiles: A statistical view (P20, P40, Median, P60, P80) of the entire network's performance.
- Customer Cities: Top 20 city tables for each store based on unique customers.
- Summary View: A condensed sheet highlighting the most critical KPIs for quick executive reading.
- Summary Quintiles: A statistical distribution of the critical KPIs from the summary view.

