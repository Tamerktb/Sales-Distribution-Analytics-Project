# Excel Analysis Plan - Al Baraka Distribution Co.

## Step 1: Load & Inspect Raw Data
- Open `data/01_raw_sales_data.csv` in Excel
- Observe issues: mixed date formats, missing values, inconsistent naming, currency symbols

## Step 2: Data Cleaning Transformations

### 2.1 Standardize Dates
Create helper column `OrderDateClean`:
```
=DATEVALUE(SUBSTITUTE(A2,"-","/"))
```
Then format as `YYYY-MM-DD`.

### 2.2 Clean Customer Names
Create lookup using `VLOOKUP` or `XLOOKUP` against `customer_master.csv`:
```
=XLOOKUP(TRIM(C2), customer_master[CustomerName], customer_master[CustomerID], "UNKNOWN")
```

### 2.3 Product Name Standardization
Use `SWITCH` or `IFS` to normalize product names (e.g., "Coke" → "Coca Cola 250ml"):
```
=IFS(ISNUMBER(SEARCH("Coke",D2)), "Coca Cola 250ml", ISNUMBER(SEARCH("Coca Cola",D2)), "Coca Cola 250ml", TRUE, D2)
```

### 2.4 Handle Missing Quantities
```
=IF(G2="", 0, G2)
```

### 2.5 Standardize Currency
Remove "JD" and "JOD" from TotalAmount:
```
=IF(RIGHT(H2,2)="JD", LEFT(H2, LEN(H2)-3), IF(RIGHT(H2,3)="JOD", LEFT(H2, LEN(H2)-4), H2))*1
```

### 2.6 Region/City Standardization
Use `PROPER` to fix inconsistent casing, then `TRIM`:
```
=PROPER(TRIM(K2))
```

## Step 3: Pivot Tables & Analysis

### 3.1 Monthly Sales Trend
- **Rows:** OrderDate (grouped by Month)
- **Values:** Sum of TotalAmount
- **Chart:** Line chart — shows sales trend Jan-Jun 2026

### 3.2 Top Products by Revenue
- **Rows:** ProductName
- **Values:** Sum of TotalAmount
- **Filter:** Exclude returns (Quantity < 0)
- **Chart:** Bar chart — top 10 products

### 3.3 Customer Ranking
- **Rows:** CustomerName
- **Values:** Sum of TotalAmount
- **Chart:** Horizontal bar — top customers

### 3.4 Category Distribution
- **Rows:** Category
- **Values:** Sum of TotalAmount
- **Chart:** Pie/donut chart — revenue by category

### 3.5 Regional Performance
- **Rows:** Region, City
- **Values:** Sum of TotalAmount, Count of Orders
- **Chart:** Stacked bar — performance by region

### 3.6 Payment Status Analysis
- **Rows:** PaymentStatus
- **Values:** Sum of TotalAmount, Count of Orders
- **Insight:** % of unpaid invoices and collection risk

### 3.7 Sales Rep Performance
- **Rows:** SalesRep
- **Values:** Sum of TotalAmount, Count of Orders
- **Chart:** Bar chart — compare rep performance

## Step 4: Advanced Excel Formulas

### Month-over-Month Growth
```
=(CurrentMonth - PreviousMonth) / PreviousMonth
```

### Customer Acquisition Count (new customers per month)
Use `COUNTIFS` with first occurrence logic.

### Running Total (YTD Sales)
```
=SUMIFS(TotalAmount, OrderDate, "<="&OrderDate)
```

### Moving Average (3-month)
```
=AVERAGE(SUMIFS(TotalAmount, Month, {M-2,M-1,M}))
```

### Unpaid Invoice Aging
```
=IF(PaymentStatus="Unpaid", TODAY()-OrderDate, 0)
```

## Step 5: Dashboard Layout (Single Sheet)

```
+---------------------------------------------------+
|  AL BARAKA DISTRIBUTION CO. - SALES DASHBOARD      |
|  Period: Jan-Jun 2026  |  Al Qastal, Jordan       |
+---------------------------------------------------+
| Key KPIs (Cards):                                  |
| [Total Revenue] [Orders] [Avg Order] [Unpaid %]   |
+---------------------------------------------------+
| Monthly Sales Trend  | Category Breakdown          |
| [Line Chart]         | [Pie/Donut Chart]          |
+----------------------+----------------------------+
| Top Products         | Top Customers              |
| [Bar Chart]          | [Horizontal Bar]           |
+----------------------+----------------------------+
| Regional Map/Table   | Sales Rep Performance      |
| [Table + Bar]        | [Bar Chart]                |
+---------------------------------------------------+
```

## Step 6: Insights to Extract

1. **Overall Performance:** Total revenue, order count, average order value
2. **Growth Trend:** Month-over-month and overall growth rate
3. **Top Products:** Identify best-selling SKUs and categories
4. **Customer Concentration:** % revenue from top 3 customers
5. **Regional Gaps:** Underperforming areas (e.g., Zarqa vs Amman)
6. **Payment Risk:** Unpaid invoices by customer and aging
7. **Sales Rep Performance:** Who is driving revenue
8. **Seasonality:** Pre-Ramadan buying patterns (March-April)
9. **Returns Analysis:** Rate and reasons for returns
10. **Actionable Recommendations:** Based on findings
