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
