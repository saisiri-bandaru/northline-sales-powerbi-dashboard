# Northline Sales Performance & Time-Intelligence Dashboard

A Power BI–style analytics project for a fictional CPG company (**Northline Consumer Products**), designed to track trade sales revenue, volume, channel mix, and historical performance with month-to-date and same-period-last-year time intelligence.

**Built by:** Sai Siri Bandaru — Financial Analyst | FP&A

---

## Project Overview

This dashboard provides business visibility into Northline retail / trade sales performance across Grocery, Mass, Club, E-comm, and Convenience channels. It helps FP&A and commercial finance analyze which brand lines generate the highest revenue, how customers settle invoices, and how sales trend versus prior-year baselines.

Because a binary `.pbix` cannot be generated in this environment, the repo ships:

1. Transaction-level Excel data (`data/Northline_Sales_Data.xlsx`)
2. A polished multi-tab Excel dashboard that recreates the **Executive / MTD / SPLY** views
3. Copy-paste **DAX measures** and rebuild steps for Power BI Desktop

---

## Key Performance Indicators (KPIs)

Computed from the generated sample (`2022-01-03` → `2025-06-29`, **4,000** transactions):

| KPI | Value | Description |
| :--- | :--- | :--- |
| **Total Sales** | **$4.0M** ($4,011,324) | Total gross trade sales across all transactions. |
| **Total Quantity** | **404.1K** (404,117) | Total units / cases sold. |
| **Transactions** | **4,000** | Completed sales orders / shipment lines. |
| **Average Price** | **$9.93** | Weighted average selling price per unit. |

---

## Key Insights from the Data

* **Brand / Line Performance:** **Northline Pro** at **$1.7M** (864 transactions), followed by **Northline Core** at **$1.4M** (1,873 transactions), and **Northline Fresh** at **$841.9K** (1,263 transactions).
* **Top Selling Products:** Highest revenue SKUs — **Kitchen Towel Roll Case** ($432.0K), **Floor Cleaner Concentrate** ($421.8K), **Industrial Sanitizer 64oz** ($324.0K).
* **Channel Mix:** Leading channels by sales — **Club** ($1.3M), **Grocery** ($1.2M), **Mass** ($1.1M).
* **Settlement Methods:** Mix across **ACH / Net-30** (34.5%), **ACH / Net-45** (25.4%), **Portal Pay** (15.7%), **Credit Card** (15.2%), **Wire** (9.2%).
* **Top Cities by Volume:** **New York** (39,058), **Los Angeles** (38,489), **Chicago** (38,285), **Atlanta** (30,244).
* **Customer Rating:** **Promoter (9-10)** (47.9%), **Passive (7-8)** (32.2%), **Detractor (0-6)** (19.9%).
* **Seasonality:** Across the full sample, quarterly sales are highest in **Q2** ($1.2M) and **Q4** ($1.0M).
* **Yearly Sales:** **2022**: $968.0K, **2023**: $1.2M, **2024**: $1.2M, **2025**: $659.1K.
* **YoY (SPLY):** **2023** vs prior: **+22.0%** ($1.2M vs $968.0K); **2024** vs prior: **+1.9%** ($1.2M vs $1.2M); **2025** vs prior (Jan–Jun same-period): **+19.8%** ($659.1K vs $550.3K).

---

## Dashboard Views

### 1. Executive Dashboard (`01_Executive`)
* KPI cards: Total Sales, Total Quantity, Transactions, Avg Price.
* Sales by channel and brand/line, top products, settlement mix, city volume, rating mix, order type.
* In Power BI, add slicers for Channel, Brand, Product, Settlement Method, and City.

### 2. Month-to-Date (MTD) Report (`02_MTD`)
* **Yellow cells** = Year / Month filter inputs.
* Daily sales and cumulative MTD chart/table driven by `SUMIFS` against a daily cache.
* MTD KPIs: month sales, days with sales, average daily sales.
* Power BI equivalent: `TOTALMTD([Total_Sales], Calendar[Date])`.

### 3. Same Period Last Year (SPLY) Analysis (`03_SPLY`)
* Year, quarter, and month comparisons with prior-year (SPLY) columns.
* YoY $ and YoY % as Excel formulas (black font).
* Identifies growth trends and gaps versus the same calendar period last year.

### Supporting tabs
* `00_Cover` — package overview and headline KPIs
* `04_Measures` — Excel equivalents + DAX to paste
* `05_Data_Dictionary` — field definitions

---

## Data Model

| File | Sheets | Role |
| --- | --- | --- |
| `data/Northline_Sales_Data.xlsx` | `sales_data`, `Calendar` | Fact + date dimension for Power BI |
| `Northline_Sales_Performance_Dashboard.xlsx` | `00_Cover` … `05_Data_Dictionary` | Excel recreation of the three dashboard views |
| `assets/executive_preview.svg` | — | Static preview of the executive KPI layout |

**Relationship:** `Calendar[Date]` (1) → `sales_data[Date]` (many). Mark `Calendar` as a date table in Power BI.

### Fact columns (CPG-flavored)
Transaction ID, Date, Day/Month/Year/Day Name, Channel, Brand, SKU, Product, Units Sold, Price Per Unit, Sales Amount, Customer Name, City, Region, Order Type (Wholesale / DSD / E-comm), Settlement Method, Customer Rating, Sales Rep.

---

## Core DAX Measures

Paste into Power BI Desktop (table names: `sales_data` and `Calendar`):

```dax
// 1. Total Units Sold
Total_Quantity = SUM(sales_data[Units Sold])

// 2. Total Gross Sales
Total_Sales =
SUMX(
    sales_data,
    sales_data[Units Sold] * sales_data[Price Per Unit]
)

// 3. Total Transaction Count
Transactions = COUNTROWS(sales_data)

// 4. Weighted Average Selling Price
Average_price = DIVIDE([Total_Sales], [Total_Quantity], 0)

// 5. Month-to-Date Cumulative Sales
MTD =
TOTALMTD(
    [Total_Sales],
    Calendar[Date]
)

// 6. Prior Year Sales (Same Period Last Year)
Same Period Last Year =
CALCULATE(
    [Total_Sales],
    SAMEPERIODLASTYEAR(Calendar[Date])
)

// 7. YoY Growth (optional)
YoY_Growth =
DIVIDE([Total_Sales] - [Same Period Last Year], [Same Period Last Year], 0)
```

---

## How to Rebuild in Power BI Desktop

1. Open **Power BI Desktop** → **Get Data** → **Excel workbook**.
2. Select `data/Northline_Sales_Data.xlsx` → load **`sales_data`** and **`Calendar`**.
3. Model view: create a relationship on **Date**; mark **Calendar** as a date table.
4. Create a measures table (or use `sales_data`) and paste the DAX above.
5. Build three report pages:
   - **Executive** — KPI cards + bar/pie visuals by Channel, Brand, Product, Settlement, City, Rating
   - **MTD** — line chart of daily / cumulative sales with a month slicer
   - **SPLY** — clustered bars or matrix of current vs `Same Period Last Year` by Year / Quarter / Month
6. Optional: use the Excel workbook as a reference layout while you arrange visuals.

---

## Stack

* **Power BI Desktop** — intended primary delivery (DAX + date intelligence)
* **Excel** — source data + interactive dashboard recreation (openpyxl charts / formulas)
* **Python** (`build.py`) — reproducible sample generation

Yellow cells with blue font are filter inputs. Black font is formulas or computed values.

---

## Portfolio note

Part of Sai Siri Bandaru's FP&A GitHub platform (Northline Consumer Products fictional universe): variance packs, forecasts, working capital, KPI scorecards, and sales analytics.

---

## Disclaimer

All company names, accounts, SKUs, and figures are **fictional** and created for portfolio demonstration only. No confidential employer or customer data is included.

---

## Reproduce

```bash
python3 -m venv venv && source venv/bin/activate
pip install openpyxl pandas numpy
python build.py
```
