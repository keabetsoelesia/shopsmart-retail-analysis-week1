# shopsmart-retail-analysis-week1
Excel retail sales analysis for ShopSmart Retail (AnalystLab Africa internship, Week 1)
##Project Overview 
Stores in the most demanded areas are low stocked. A large amount of orders are pending or cancelled. 

###Files included
- ’ShopSmart_Analysis.xlsx’: cleaned data, formulas, lookup, pivot tables , charts and dashboard
- Executive_Dashboard: the cleaned dashboard exposed as .pdf
- Business_Report.pdf: one-page of findings, recommendations, risks and opportunity

##Data Cleaning Steps 
- Converted datasets into Excel Tables 
- Fixed date formats , removed duplicates, tag placeholder or incomplete phone numbers, renamed sheets.
- Checked for missing values and duplicates. 

##Key decisions and assumptions
- Only orders with status “completed” are counted as sales. For accurate calculations of revenue. 
- Return rate= returned orders / total orders , using a ‘Return_Flag’ helper column

##Workbook guide: 
- tblOrders includes new columns: Customer_Name, Store_Name, Product_Name, Category, Brand, 
Selling_Price, Supplier_Name, Sales_Amount, Order_Year_Cleaned, State, Gender, Return_Flag
- tblInventory includes: Low stock 
- Pivot Sheets: Top 10 products , Bottom 10 products, Sales by month, sales by store, sales by category, sales by brand, sales by state, sales by gender, sales by membership level , return rate by category, return rate by product , return rate by brand. 
- Dashboard: KPI cards, five charts and slicers 

##Tools Used: 
Microsoft Excel: XLOOKUP, SUMIFS, PivotTables, PivotCharts, Slicers). 
