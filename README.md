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
- **Customer Segment** *(VAR)* — `VAR CustSales = [Total Sales] RETURN SWITCH(TRUE(), CustSales >= 100000, "Gold", CustSales >= 50000, "Silver", CustSales > 0, "Bronze", "No Sales")` — Tiered customer segment label based on sales. *(Legacy measure — see Week 3 note; the Customer Segment used on report visuals is now a calculated column on the Customer table.)*
- **Days Since Last Purchase** — `DATEDIFF(CALCULATE(MAX(Sales[Order Date Key]), ALLEXCEPT(Sales, Sales[Customer Key])), TODAY(), DAY)` — Days elapsed since a customer's most recent order.
- **Budget Attainment %** — `DIVIDE([Total Sales], [Budget Amount by Category], 0)` — Actual sales as a % of budget.
- **Pareto %** *(VAR)* —
```DAX
  Pareto % =
  VAR RunningTotal =
      CALCULATE(
          [Total Sales],
          FILTER(
              ALL(Product[Product Name]),
              RANKX(ALL(Product[Product Name]), [Total Sales]) <=
                  RANKX(ALL(Product[Product Name]), [Total Sales])
          )
      )
  VAR GrandTotal =
      CALCULATE([Total Sales], ALL(Product))
  RETURN
      DIVIDE(RunningTotal, GrandTotal, 0)
```
  Cumulative % of total sales contributed by products at or above the current product's sales rank, used for 80/20 (Pareto) analysis.

### CALCULATE Deep Dive
Three measures were rewritten to use explicit CALCULATE with modifiers, each documented inline with a comment explaining the filter context before and after:
- **Sales YTD** — uses `REMOVEFILTERS(Calendar[Year])` to clear existing year filters before reapplying a year-to-date filter.
- **Has Sales** — uses `KEEPFILTERS` with `RELATEDTABLE` to check for related sales rows without overriding existing filters.
- **Budget Amount by Category** — uses `ALL(Budget)` to clear prior filters on Budget before applying `TREATAS` to cross-filter from Product.

## Week 3 — Report Design, UX & Advanced Visuals

### Project Goal
Build a 5-page executive-ready report with calculation groups, field parameters, and conditional visibility.

### Executive Summary Page
- 5 KPI cards: Total Sales, Gross Profit %, Sales YoY Change %, Total Orders, Total Customers (substituted for Return Rate %, since no Returns table exists in this model — see Week 2 note).
- Monthly sales line chart with a prior-year comparison line, sorted correctly by Month Number (Month Name alone sorts alphabetically, so Month Number was used as the sort-by column).
- Sales by Territory bar chart, showing all 10 regions.
- A Page Navigator visual gives one-click navigation to Exec Summary, Sales Analysis, and Customer Analysis. A dedicated Mobile Layout was built for this page (View → Mobile Layout), rearranging the KPI cards and charts into a vertical phone-friendly stack.

### Calculation Groups
A **Time Intelligence** Calculation Group was built in Tabular Editor 3, with six Calculation Items: Current Period, Prior Year, YoY Change $, YoY Change %, YTD, MTD. Each item wraps `SELECTEDMEASURE()` with the relevant time-intelligence pattern (e.g. `CALCULATE(SELECTEDMEASURE(), SAMEPERIODLASTYEAR(Calendar[Date]))`).

**Why a Calculation Group instead of individual measures?** Without it, every metric (Total Sales, Gross Profit, Total Orders, etc.) would need its own YTD, PY, YoY, and MTD variant — multiplying the measure count and creating duplicate logic to maintain. A Calculation Group lets a single slicer switch the time-period behaviour for *every* measure on a page at once, without touching the underlying measures. This keeps the Measures table lean and makes time-period logic a single, centrally maintained set of expressions.

The Calculation Group slicer is applied to the Sales Analysis page and the Tooltip page.

### Sales Analysis Page
- A matrix of Category/Subcategory rows against fixed columns: Total Sales, Gross Profit, Sales YoY Change %, Gross Profit Margin %. Per a strict reading of the project plan, the Field Parameters (Metric Selector, Dimension Selector) do not replace the matrix's own metrics — they instead drive a **separate** clustered column chart ("Total Sales by Category"), letting users dynamically switch its metric and axis.
- The Calculation Group slicer is applied to the matrix, allowing users to toggle between Current Period, Prior Year, YTD, MTD, YoY Change $, and YoY Change % for every value in the matrix at once.
- Conditional formatting: Sales YoY Change % uses a red/green gradient; Gross Profit Margin % uses rule-based colour thresholds (red below 20%, yellow 20–40%, green at or above 40%).

### Conditional Visibility
Two patterns were implemented on the Sales Analysis page:

1. **Bar/line chart toggle** — the "Total Sales by Category" chart exists in two versions (bar and line), stacked in the same position via the Selection Pane. Two bookmarks ("Show Bar", "Show Line") capture each visibility state, grouped together under a "Chart Toggle" bookmark group so they don't interfere with the page-navigation bookmarks. A "Switch View" button is wired to the "Show Line" bookmark via an Action → Bookmark interaction.

2. **Measure-driven empty state** — a Card visual (rather than a text box, which does not support visual-level filtering in this Power BI version) is bound to a dedicated measure:
```DAX
   No Data Message = IF(ISBLANK([Total Sales]), "No data available for the selected filters", BLANK())
```
   The Card only displays text when `Total Sales` is blank for the current filter context, and is otherwise empty — no bookmark required. **Root cause note:** this Card initially threw a repeated "Error fetching data for this visual" error; the cause was the Time Intelligence Calculation Group slicer on the same page cross-filtering the text-based measure (its calculation items attempt numeric operations on `SELECTEDMEASURE()`, which breaks on non-numeric output). The fix was ensuring the Calculation Group slicer is set to **Current Period** for this Card's context.

   A related measure, **Show No Data Message** — `IF(ISBLANK([Total Sales]), 1, 0)` — returns a 1/0 flag version of the same condition, used to drive visual-level filtering (rather than the text output itself) where a numeric flag is more convenient than a text-blank check.

### Customer Analysis Page
- Top 20 Customers table (Customer Name, Customer Segment, Days Since Last Purchase, Total Sales), sorted descending by Total Sales, with a Top 20 data limit and `[Not Applicable]` rows excluded via a page-level filter.
- Sales by Customer Segment bar chart (Gold vs Bronze split).
- Sales by Country-Region and City treemap.
- High Value Customer Count card.
- Customer Segment slicer and a "Between"-style date range slicer on `Calendar[Date]`.

**Bugs fixed on this page:**
- *Days Since Last Purchase* originally used `DATEDIFF` against `Sales[Order Date Key]` (a numeric surrogate key, not a real date), which errored out. Fixed by referencing `Calendar[Date]` via `RELATEDTABLE`/`RELATED` instead of `ALLEXCEPT`.
- *Customer Segment* was originally a measure, which meant it couldn't be placed on a chart axis (axes require columns) and, when first converted to a calculated column using `ALLEXCEPT`, incorrectly returned "Gold" for every customer. Fixed by rebuilding it as a calculated column on the Customer table using a plain `CALCULATE([Total Sales])`, relying on relationship context transition rather than `ALLEXCEPT`. The original measure was kept (unchanged, in the `measures` table) as a reference implementation of the VAR/SWITCH pattern for Week 2's Ranking & Segmentation requirement; the calculated column on Customer is what report visuals actually use.
- *High Value Customer Count* returned 1 despite looking correct; the single row crossing the 50,000 threshold turned out to be the `[Not Applicable]` customer row, not a real customer (real customers top out around ₹15–17K). The measure was rewritten to explicitly exclude `[Not Applicable]` and only count Silver/Gold-segment customers:
```DAX
  High Value Customer Count =
  IF(
      ISBLANK(
          CALCULATE(
              DISTINCTCOUNT(Customer[Customer Key]),
              Customer[Customer Segment] IN {"Silver", "Gold"},
              Customer[Customer Name] <> "[Not Applicable]"
          )
      ),
      0,
      CALCULATE(
          DISTINCTCOUNT(Customer[Customer Key]),
          Customer[Customer Segment] IN {"Silver", "Gold"},
          Customer[Customer Name] <> "[Not Applicable]"
      )
  )
```
  It now displays cleanly as 0. **Business insight:** by this dataset's numbers, no real customer currently qualifies as Silver or Gold tier — worth surfacing in the Week 4 business insights summary.

### Product Detail Drill-Through Page
A 5th, hidden page with drill-through enabled on `Product[Product Name]`. Right-clicking a product name anywhere in the report opens this page filtered to that product. It shows:
- Three KPI cards: Total Sales, Gross Profit, Total Orders.
- A monthly sales trend line chart (Month Number/Month Name on the X-axis, same sorting approach as the Exec Summary chart).
- An orders table (Date, Customer Name, Product Name, Total Sales).

### Tooltip Page
A dedicated Tooltip page ("Allow use as tooltip" enabled, canvas sized as a tooltip) shows a Total Sales card filtered to the Calculation Group's Current Period, so the tooltip always reflects the base metric regardless of what time-period toggle is active elsewhere. It is wired as a **Report page tooltip** on the Sales Analysis line chart and matrix, so hovering over either shows the custom tooltip instead of the Power BI default.

### Navigation & Bookmark Organisation
A Page Navigator visual was added to Exec Summary, Sales Analysis, Customer Analysis, and Product Detail. Page 1–3 (legacy test pages) and Product Detail are hidden from the navigator's page list, since Product Detail is only meant to be reached via drill-through. The conditional-visibility bookmarks are kept in their own "Chart Toggle" group, separate from page-navigation, to prevent the two bookmark types from interfering with each other.

### QA Pass
A full QA pass was completed across the report: slicers (Category, Customer Segment, Date range, Calculation Group), drill-through (Product Name → Product Detail), bookmarks (bar/line chart toggle), tooltips (Sales Analysis line chart and matrix), field parameters (Metric Selector, Dimension Selector), Calculation Group toggling (Current Period / MTD / Prior Year / YoY $ / YoY % / YTD), and navigation buttons across all visible pages. All items confirmed working correctly.

## Week 4 — Security, Performance, Documentation & Presentation

### Project Goal
Finalise the report, add enterprise features, and build a portfolio-ready submission.

### Row-Level Security — Static Roles
Three static security roles were created under Manage Roles, each filtering the `SalesTerritory[Group]` column:
- **North America Manager** — `SalesTerritory[Group] = "North America"`
- **Europe Manager** — `SalesTerritory[Group] = "Europe"`
- **Pacific Manager** — `SalesTerritory[Group] = "Pacific"`

Each role was tested using **View As Roles**, confirming that switching to a given role restricts all report visuals (KPI cards, region charts, matrices) to that territory's data only, and no other territory's data is visible.

### Row-Level Security — Dynamic RLS
An **Employees** table (Name, Territory) was created via Enter Data, with no formal relationship to SalesTerritory. A dynamic role, **Dynamic Employee Territory**, was written using a DAX filter expression on `SalesTerritory[Group]`:
```DAX
[Group] = LOOKUPVALUE(Employees[Territory], Employees[Name], USERPRINCIPALNAME())
```
This looks up the current signed-in user's territory from the Employees table and filters SalesTerritory to match. As noted in the project plan, this rule cannot be fully tested inside Power BI Desktop, since `USERPRINCIPALNAME()` only resolves to a real value once the report is published to the Power BI Service under Power BI Pro — it is documented here rather than validated locally.

### Performance Analyser
Performance Analyser was run from the View tab on the Sales Analysis page. On a **cold cache** (first refresh), the 3 slowest visuals were:
- Two Category/Region **Slicers** — approximately 7000ms each
- The **Matrix** (Category/Subcategory breakdown) — 7675ms
- A **Card** visual — 7013ms

Breaking down the Matrix and Slicer durations showed that the DAX query time itself was negligible (15–65ms) — the overwhelming majority of the time (5000–7000ms, logged as "Other") was first-load rendering overhead, not calculation cost. Re-running Performance Analyser a second time (warm cache) confirmed this: Slicer durations dropped to roughly 560–2000ms, while the Matrix remained the comparatively slowest visual even warm, at around 3000ms.

**Optimisation:** The Matrix was carrying four Calculation-Group-driven measures (Total Sales, Gross Profit, Sales YoY Change %, Gross Profit Margin %). Removing one column — **Gross Profit Margin %** — and re-running Performance Analyser dropped the Matrix's duration from **7675ms to roughly 726–927ms**, an approximately 8x improvement. This confirms that each additional Calculation-Group-driven measure in a single visual adds real, avoidable rendering cost, and that trimming a visual's measure count is an effective, low-risk optimisation.

### Advanced Power Query
The **Product** and **Budget** queries were combined using **Merge Queries** (Home → Merge Queries), with **Product as the base query** and **Budget merged in**, joined on the shared `Category` column with a **Left Outer** join (all rows from Product, matching rows from Budget). The resulting `Budget` column was expanded to bring in just `Budget Amount`.

**Why the merge direction matters:** the first attempt merged Product into Budget (Budget as the base query). Since many products share a single category, this duplicated Budget's rows — one row per matching product instead of one row per category — which broke Budget's row uniqueness on `Category`, the exact column the Week 2 TREATAS relationship (`Budget Amount by Category`) depends on to filter correctly. The fix was reversing the merge direction: Product as the base, Budget's `Budget Amount` pulled in as a lookup column. This keeps Product's row count unchanged (397 rows, unduplicated) and leaves Budget's table structure — and the TREATAS measures built on it — untouched.

A custom column, Budget Status, was then added using Add Column → Custom Column with an if-then-else expression in M:

if [Budget Amount] > 10000000 then "High Budget" else "Standard Budget"

Why this was done in Power Query rather than DAX: the merge and the "Budget Status" label are structural, one-time transformations of the data itself, not calculations that need to respond to filter context or user interaction. Power Query runs once at refresh time and produces a static shaped table, which is more efficient than recalculating a row-level classification with DAX on every visual interaction. DAX measures are reserved for aggregations that must react dynamically to slicers and filters; static row-level labels like "High Budget"/"Standard Budget" belong in the data layer, not the calculation layer.
