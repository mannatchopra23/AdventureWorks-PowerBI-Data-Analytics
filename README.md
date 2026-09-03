# AdventureWorks Power BI Data Analytics

## Week 1 — Data Loading, Power Query & Data Modelling

### Project Goal
Build a clean Power BI data model with proper relationships before writing DAX measures.

### Data Model
The project uses the following tables:
- **Sales** — Fact table containing sales transactions.
- **Product** — Dimension table containing product information.
- **Customer** — Dimension table containing customer information.
- **Reseller** — Dimension table containing reseller information.
- **SalesTerritory** — Dimension table containing territory information.
- **Calendar** — Date dimension used for date analysis and time-based reporting.
- **SalesOrder** — Additional sales order information.

### Relationships
The main relationships in the model are:
- Customer (Customer Key) → Sales (Customer Key)
- Product (Product Key) → Sales (Product Key)
- Reseller (Reseller Key) → Sales (Reseller Key)
- SalesTerritory (Territory Key) → Sales (Sales Territory Key)
- Calendar (Date Key) → Sales (Order Date Key)
- SalesOrder (Sales Order Line Key) → Sales (Sales Order Line Key)

### Model Structure
The model follows a star-schema approach, with **Sales** acting as the central fact table and the descriptive tables surrounding it.
Relationships between the dimension tables and Sales are primarily one-to-many (`1:*`) with single-direction filtering.

### Validation
A Total Sales measure was created to validate the model:
```DAX
Total Sales = SUM(Sales[Sales Amount])
```

## Week 2 — DAX Measures, Time Intelligence & Advanced Calculations

### Project Goal
Write 25+ DAX measures covering foundational aggregations, time intelligence, VAR/RETURN patterns, error handling, TREATAS cross-table filtering, field parameters, ranking/segmentation, and CALCULATE modifiers — all stored in a dedicated Measures table.

### Data Limitation
No Returns table exists in this model, so **Total Returns** and **Return Rate %** from the original task list were not built.

### Foundation Measures
- **Total Sales** — `SUM(Sales[Sales Amount])` — Total revenue from all sales transactions.
- **Total Cost** — `SUM(Sales[Total Product Cost])` — Total cost of goods sold.
- **Gross Profit** — `[Total Sales] - [Total Cost]` — Revenue minus cost of goods sold.
- **Gross Profit Margin %** *(VAR)* — `VAR TotalSalesAmt = [Total Sales] VAR TotalGrossProfit = [Gross Profit] RETURN DIVIDE(TotalGrossProfit, TotalSalesAmt, 0)` — Gross profit as a % of total sales.
- **Total Customers** — `DISTINCTCOUNT(Sales[Customer Key])` — Count of unique customers who made a purchase.
- **Total Units Sold** — `SUM(Sales[Order Quantity])` — Total quantity of items sold.
- **Total Orders** — `DISTINCTCOUNT(Sales[Sales Order Line Key])` — Count of distinct order lines.
- **Average Order Value** *(VAR)* — `VAR TotalSalesAmt = [Total Sales] VAR TotalOrderCount = [Total Orders] RETURN DIVIDE(TotalSalesAmt, TotalOrderCount, 0)` — Average revenue per order.

### Time Intelligence Measures
- **Sales YTD** *(CALCULATE modifiers)* — `CALCULATE(TOTALYTD([Total Sales], Calendar[Date]), REMOVEFILTERS(Calendar[Year]))` — Cumulative sales from Jan 1 to the latest date in context.
- **Sales PY** — `CALCULATE([Total Sales], SAMEPERIODLASTYEAR(Calendar[Date]))` — Sales for the same period, one year earlier.
- **Sales YoY Change $** *(VAR)* — `VAR CurrentSales = [Sales YTD] VAR PriorSales = [Sales PY] RETURN CurrentSales - PriorSales` — Absolute change in sales vs. same period last year.
- **Sales YoY Change %** *(VAR)* — `VAR CurrentSales = [Sales YTD] VAR PriorSales = [Sales PY] RETURN DIVIDE(CurrentSales - PriorSales, PriorSales, 0)` — % change in sales vs. same period last year.
- **Sales MTD** — `TOTALMTD([Total Sales], Calendar[Date])` — Cumulative sales, month-to-date.
- **Sales QTD** — `TOTALQTD([Total Sales], Calendar[Date])` — Cumulative sales, quarter-to-date.
- **Rolling 3-Month Sales** *(VAR)* — `VAR MaxSaleDate = MAX(Calendar[Date]) VAR RollingDates = DATESINPERIOD(Calendar[Date], MaxSaleDate, -3, MONTH) RETURN CALCULATE([Total Sales], RollingDates)` — Trailing 3-month sales total.
- **Rolling 12-Month Sales** *(VAR)* — `VAR MaxSaleDate = MAX(Calendar[Date]) VAR RollingDates = DATESINPERIOD(Calendar[Date], MaxSaleDate, -12, MONTH) RETURN CALCULATE([Total Sales], RollingDates)` — Trailing 12-month sales total.

Tested in a matrix with Year in rows: PY is blank for the first year, and YoY figures correctly reference the prior year across all measures.

### Error Handling
- All DIVIDE-based measures use its third argument (0) as a safe fallback instead of IFERROR, per best practice.
- **Has Sales** *(VAR, CALCULATE modifiers)* — `VAR CustomerSales = CALCULATE(COUNTROWS(RELATEDTABLE(Sales)), KEEPFILTERS(ALL(Sales[Order Quantity]))) RETURN CustomerSales > 0` — Flags whether the current context has any related sales.

### TREATAS — Cross-Table Filtering
A standalone **Budget** table (Category, Budget Amount) was created with no formal relationship to Product, to demonstrate TREATAS.
- **Budget Amount by Category** *(CALCULATE modifiers, TREATAS)* — `CALCULATE(SUM(Budget[Budget Amount]), ALL(Budget), TREATAS(VALUES(Product[Category]), Budget[Category]))` — Pulls budget amount from the unrelated Budget table by matching Category.
- **Budget vs Actual Variance** *(VAR, TREATAS)* — `VAR ActualSales = [Total Sales] VAR BudgetTarget = [Budget Amount by Category] RETURN ActualSales - BudgetTarget` — Actual sales minus budgeted amount.

### Field Parameters
- **Metric Selector** — switches a chart's metric between Total Sales, Gross Profit, and Total Orders.
- **Dimension Selector** — switches a chart's axis between Category, Region, and Country-Region.

### Ranking & Segmentation
- **Top 10 Products by Sales** — `IF(RANKX(ALL(Product[Product Name]), [Total Sales]) <= 10, [Total Sales])` — Shows sales value only for the top 10 ranked products.
- **Product Rank within Category** — `RANKX(FILTER(ALL(Product), Product[Category] = MAX(Product[Category])), [Total Sales])` — Product's sales rank within its own category.
- **High Value Customer** — `IF([Total Sales] > 50000, "Yes", "No")` — Flags customers with sales above a threshold.
- **Customer Segment** *(VAR)* — `VAR CustSales = [Total Sales] RETURN SWITCH(TRUE(), CustSales >= 100000, "Gold", CustSales >= 50000, "Silver", CustSales > 0, "Bronze", "No Sales")` — Tiered customer segment label based on sales.
- **Days Since Last Purchase** — `DATEDIFF(CALCULATE(MAX(Sales[Order Date Key]), ALLEXCEPT(Sales, Sales[Customer Key])), TODAY(), DAY)` — Days elapsed since a customer's most recent order.
- **Budget Attainment %** — `DIVIDE([Total Sales], [Budget Amount by Category], 0)` — Actual sales as a % of budget.
- **Pareto %** *(VAR)* — cumulative % of total sales contributed, ranked by product (80/20 analysis).

### CALCULATE Deep Dive
Three measures were rewritten to use explicit CALCULATE with modifiers, each documented inline with a comment explaining the filter context before and after:
- **Sales YTD** — uses `REMOVEFILTERS(Calendar[Year])` to clear existing year filters before reapplying a year-to-date filter.
- **Has Sales** — uses `KEEPFILTERS` with `RELATEDTABLE` to check for related sales rows without overriding existing filters.
- **Budget Amount by Category** — uses `ALL(Budget)` to clear prior filters on Budget before applying `TREATAS` to cross-filter from Product.
