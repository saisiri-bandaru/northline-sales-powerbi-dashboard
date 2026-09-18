#!/usr/bin/env python3
"""
Build Northline Sales Performance & Time-Intelligence Dashboard portfolio package.

Outputs:
  data/Northline_Sales_Data.xlsx          — sales_data + Calendar
  Northline_Sales_Performance_Dashboard.xlsx — multi-tab Excel dashboard
  assets/executive_preview.svg            — simple executive dashboard preview
  README.md                               — recruiter-facing docs with real KPIs
  insights.json                           — computed metrics (intermediate)
"""
from __future__ import annotations

import json
import math
import random
from datetime import date, datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
ASSETS_DIR = ROOT / "assets"
RNG = random.Random(42)
NP_RNG = np.random.default_rng(42)

# ---------------------------------------------------------------------------
# Styles (Northline FP&A palette)
# ---------------------------------------------------------------------------
YELLOW = PatternFill("solid", fgColor="FFF2CC")
HEADER = PatternFill("solid", fgColor="1F4E79")
SECTION = PatternFill("solid", fgColor="D6E3F0")
TILE = PatternFill("solid", fgColor="E9EDF4")
LIGHT = PatternFill("solid", fgColor="F5F5F5")
GREEN = PatternFill("solid", fgColor="C6EFCE")
WHITE = PatternFill("solid", fgColor="FFFFFF")
COVER_BG = PatternFill("solid", fgColor="1F4E79")

INPUT_FONT = Font(name="Calibri", size=11, color="0000FF", bold=True)
BLACK = Font(name="Calibri", size=11, color="000000")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
TITLE_FONT = Font(name="Calibri", size=18, bold=True, color="1F4E79")
SECTION_FONT = Font(name="Calibri", size=12, bold=True, color="1F4E79")
KPI_LABEL = Font(name="Calibri", size=10, bold=True, color="1F4E79")
KPI_VALUE = Font(name="Calibri", size=16, bold=True, color="000000")
COVER_TITLE = Font(name="Calibri", size=22, bold=True, color="FFFFFF")
COVER_SUB = Font(name="Calibri", size=12, color="D6E3F0")
THIN = Border(
    left=Side(style="thin", color="B0B0B0"),
    right=Side(style="thin", color="B0B0B0"),
    top=Side(style="thin", color="B0B0B0"),
    bottom=Side(style="thin", color="B0B0B0"),
)

CHANNELS = ["Grocery", "Mass", "Club", "E-comm", "Convenience"]
CHANNEL_W = [0.32, 0.24, 0.18, 0.16, 0.10]

BRANDS = {
    "Northline Core": [
        ("NL-CORE-001", "Everyday Dish Soap 24oz", 3.49, 5.99),
        ("NL-CORE-002", "All-Purpose Cleaner 32oz", 4.29, 7.49),
        ("NL-CORE-003", "Laundry Detergent 50oz", 8.99, 14.99),
        ("NL-CORE-004", "Paper Towels 6-Pack", 6.49, 11.99),
        ("NL-CORE-005", "Trash Bags 45ct", 7.99, 13.49),
        ("NL-CORE-006", "Hand Soap Refill 32oz", 4.99, 8.49),
    ],
    "Northline Fresh": [
        ("NL-FRSH-001", "Organic Pasta Sauce 24oz", 4.49, 7.99),
        ("NL-FRSH-002", "Cold-Pressed Juice 12oz", 3.99, 6.49),
        ("NL-FRSH-003", "Greek Yogurt 32oz", 5.49, 8.99),
        ("NL-FRSH-004", "Granola Clusters 12oz", 4.79, 7.49),
        ("NL-FRSH-005", "Sparkling Water 8pk", 5.99, 9.49),
        ("NL-FRSH-006", "Plant Protein Bar 4pk", 6.49, 10.99),
    ],
    "Northline Pro": [
        ("NL-PRO-001", "Pro Degreaser Gallon", 14.99, 24.99),
        ("NL-PRO-002", "Foodservice Wipes 200ct", 12.49, 19.99),
        ("NL-PRO-003", "Industrial Sanitizer 64oz", 16.99, 27.99),
        ("NL-PRO-004", "Kitchen Towel Roll Case", 22.99, 34.99),
        ("NL-PRO-005", "Floor Cleaner Concentrate", 18.49, 29.99),
    ],
}
BRAND_W = [0.48, 0.32, 0.20]

CITIES = [
    ("New York", "Northeast"),
    ("Boston", "Northeast"),
    ("Philadelphia", "Northeast"),
    ("Chicago", "Midwest"),
    ("Detroit", "Midwest"),
    ("Minneapolis", "Midwest"),
    ("Dallas", "South"),
    ("Houston", "South"),
    ("Atlanta", "South"),
    ("Miami", "South"),
    ("Los Angeles", "West"),
    ("San Francisco", "West"),
    ("Seattle", "West"),
    ("Denver", "West"),
    ("Phoenix", "West"),
]
CITY_W = [0.10, 0.06, 0.07, 0.09, 0.05, 0.05, 0.08, 0.07, 0.07, 0.05, 0.09, 0.06, 0.06, 0.05, 0.05]

ACCOUNTS = {
    "Grocery": [
        "FreshMart Regional", "GreenValley Foods", "Harbor Grocer Co.",
        "Summit Market Group", "Oak & Vine Markets", "Riverbend Grocery",
    ],
    "Mass": [
        "ValueMart Corp", "National Supercenter", "Everyday BigBox",
        "CrossTown Mass Retail", "Unity Discount Stores",
    ],
    "Club": [
        "MembersPlus Warehouse", "ClubValue Wholesale", "PrimeBasket Clubs",
        "BulkSaver Membership",
    ],
    "E-comm": [
        "CartDirect Online", "ShipFresh Marketplace", "QuickPantry.com",
        "HomeBasket Digital", "Click&Collect Hub",
    ],
    "Convenience": [
        "CornerStop C-Stores", "Fuel&Food Express", "Metro Mini Marts",
        "QuickBite Convenience",
    ],
}

ORDER_TYPES = {
    "Grocery": ["Wholesale", "DSD"],
    "Mass": ["Wholesale", "DSD"],
    "Club": ["Wholesale"],
    "E-comm": ["E-comm"],
    "Convenience": ["DSD", "Wholesale"],
}

SETTLEMENT = ["ACH / Net-30", "ACH / Net-45", "Credit Card", "Wire", "Portal Pay"]
SETTLEMENT_W = [0.35, 0.25, 0.15, 0.10, 0.15]

RATINGS = ["Promoter (9-10)", "Passive (7-8)", "Detractor (0-6)"]
RATING_W = [0.48, 0.32, 0.20]

REPS = [
    "A. Chen", "B. Okonkwo", "C. Ramirez", "D. Patel", "E. Novak",
    "F. Brooks", "G. Singh", "H. Alvarez", "I. Kim", "J. Foster",
]


def _choice(seq, weights=None):
    if weights is None:
        return RNG.choice(seq)
    return RNG.choices(seq, weights=weights, k=1)[0]


def generate_transactions(n: int = 4000) -> pd.DataFrame:
    """Generate ~n retail/trade sales transactions spanning 2022-01-01 to 2025-06-30."""
    start = date(2022, 1, 1)
    end = date(2025, 6, 30)
    span_days = (end - start).days

    rows = []
    for i in range(1, n + 1):
        # Mild growth + seasonality (Q4 / summer bumps)
        t = i / n
        base_day = int(NP_RNG.beta(1.1, 0.95) * span_days)  # slight tilt to later years
        d = start + timedelta(days=base_day)
        # Extra weight on Nov/Dec and Jun/Jul
        if d.month in (11, 12) and RNG.random() < 0.15:
            pass  # keep
        elif d.month not in (6, 7, 11, 12) and RNG.random() < 0.08:
            d = start + timedelta(days=RNG.randint(0, span_days))

        channel = _choice(CHANNELS, CHANNEL_W)
        brand = _choice(list(BRANDS.keys()), BRAND_W)
        sku, product, p_lo, p_hi = _choice(BRANDS[brand])
        # Channel price realization
        channel_mult = {
            "Grocery": 1.00, "Mass": 0.92, "Club": 0.85,
            "E-comm": 1.05, "Convenience": 1.12,
        }[channel]
        price = round(RNG.uniform(p_lo, p_hi) * channel_mult, 2)

        # Case/pack units — Club and Mass larger (wholesale trade volumes)
        if channel == "Club":
            units = RNG.randint(60, 360)
        elif channel == "Mass":
            units = RNG.randint(36, 216)
        elif channel == "Grocery":
            units = RNG.randint(18, 144)
        elif channel == "E-comm":
            units = RNG.randint(3, 48)
        else:
            units = RNG.randint(6, 60)

        city, region = _choice(CITIES, CITY_W)
        account = _choice(ACCOUNTS[channel])
        order_type = _choice(ORDER_TYPES[channel])
        settlement = _choice(SETTLEMENT, SETTLEMENT_W)
        rating = _choice(RATINGS, RATING_W)
        rep = _choice(REPS)

        rows.append({
            "Transaction ID": i,
            "Date": d,
            "Day": d.day,
            "Month": d.month,
            "Year": d.year,
            "Day Name": d.strftime("%A"),
            "Channel": channel,
            "Brand": brand,
            "SKU": sku,
            "Product": product,
            "Units Sold": units,
            "Price Per Unit": price,
            "Sales Amount": round(units * price, 2),
            "Customer Name": account,
            "City": city,
            "Region": region,
            "Order Type": order_type,
            "Settlement Method": settlement,
            "Customer Rating": rating,
            "Sales Rep": rep,
        })

    df = pd.DataFrame(rows).sort_values("Date").reset_index(drop=True)
    df["Transaction ID"] = range(1, len(df) + 1)
    return df


def build_calendar(start: date, end: date) -> pd.DataFrame:
    dates = pd.date_range(start, end, freq="D")
    cal = pd.DataFrame({"Date": dates.date})
    cal["Year"] = [d.year for d in cal["Date"]]
    cal["Quarter"] = [f"Q{((d.month - 1) // 3) + 1}" for d in cal["Date"]]
    cal["Month"] = [d.month for d in cal["Date"]]
    cal["Month Name"] = [d.strftime("%b") for d in cal["Date"]]
    cal["Day"] = [d.day for d in cal["Date"]]
    cal["Day Name"] = [d.strftime("%A") for d in cal["Date"]]
    cal["Weeknum"] = [d.isocalendar()[1] for d in cal["Date"]]
    cal["YearMonth"] = [d.strftime("%Y-%m") for d in cal["Date"]]
    return cal


def compute_insights(df: pd.DataFrame) -> dict:
    total_sales = float(df["Sales Amount"].sum())
    total_qty = int(df["Units Sold"].sum())
    n_txn = len(df)
    avg_price = total_sales / total_qty if total_qty else 0

    brand = (
        df.groupby("Brand")
        .agg(Sales=("Sales Amount", "sum"), Txns=("Transaction ID", "count"), Qty=("Units Sold", "sum"))
        .sort_values("Sales", ascending=False)
    )
    product = (
        df.groupby("Product")
        .agg(Sales=("Sales Amount", "sum"), Qty=("Units Sold", "sum"))
        .sort_values("Sales", ascending=False)
    )
    channel = (
        df.groupby("Channel")
        .agg(Sales=("Sales Amount", "sum"), Txns=("Transaction ID", "count"), Qty=("Units Sold", "sum"))
        .sort_values("Sales", ascending=False)
    )
    settlement = df["Settlement Method"].value_counts(normalize=True) * 100
    city = (
        df.groupby("City")
        .agg(Qty=("Units Sold", "sum"), Sales=("Sales Amount", "sum"))
        .sort_values("Qty", ascending=False)
    )
    rating = df["Customer Rating"].value_counts(normalize=True) * 100
    df2 = df.copy()
    df2["Quarter"] = df2["Date"].map(lambda d: f"Q{((d.month - 1) // 3) + 1}")
    qtr = df2.groupby(["Year", "Quarter"])["Sales Amount"].sum()
    year = df.groupby("Year")["Sales Amount"].sum()
    order_type = df.groupby("Order Type")["Sales Amount"].sum().sort_values(ascending=False)

    # YoY for latest full-ish year pair
    yoy = {}
    years = sorted(df["Year"].unique())
    for y in years[1:]:
        cur = float(year.get(y, 0))
        prev = float(year.get(y - 1, 0))
        yoy[str(y)] = {"sales": cur, "prior": prev, "pct": ((cur / prev) - 1) * 100 if prev else None,
                       "partial": False}

    # H1 same-period YoY for partial final year (data ends mid-year)
    max_d = df["Date"].max()
    if max_d.month < 12:
        h1_mask = df["Month"] <= max_d.month
        h1 = df[h1_mask].groupby("Year")["Sales Amount"].sum()
        y_last = int(max_d.year)
        if y_last in h1.index and (y_last - 1) in h1.index:
            cur = float(h1[y_last])
            prev = float(h1[y_last - 1])
            yoy[str(y_last)] = {
                "sales": cur, "prior": prev,
                "pct": ((cur / prev) - 1) * 100 if prev else None,
                "partial": True, "through_month": int(max_d.month),
            }

    return {
        "total_sales": total_sales,
        "total_qty": total_qty,
        "n_txn": n_txn,
        "avg_price": avg_price,
        "date_min": str(df["Date"].min()),
        "date_max": str(df["Date"].max()),
        "brand": brand.reset_index().to_dict(orient="records"),
        "product_top": product.head(5).reset_index().to_dict(orient="records"),
        "channel": channel.reset_index().to_dict(orient="records"),
        "settlement_pct": {k: float(v) for k, v in settlement.items()},
        "city_top": city.head(5).reset_index().to_dict(orient="records"),
        "rating_pct": {k: float(v) for k, v in rating.items()},
        "quarterly": [{"Year": int(y), "Quarter": q, "Sales": float(s)} for (y, q), s in qtr.items()],
        "yearly": {str(int(y)): float(s) for y, s in year.items()},
        "order_type": {k: float(v) for k, v in order_type.items()},
        "yoy": yoy,
    }


def fmt_money(x: float, compact: bool = True) -> str:
    if compact:
        if abs(x) >= 1_000_000:
            return f"${x / 1_000_000:.1f}M"
        if abs(x) >= 1_000:
            return f"${x / 1_000:.1f}K"
    return f"${x:,.0f}"


def fmt_qty(x: float) -> str:
    if abs(x) >= 1_000_000:
        return f"{x / 1_000_000:.1f}M"
    if abs(x) >= 1_000:
        return f"{x / 1_000:.1f}K"
    return f"{x:,.0f}"


# ---------------------------------------------------------------------------
# Excel writers
# ---------------------------------------------------------------------------

def style_header_row(ws, row: int, start_col: int, end_col: int):
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row, c)
        cell.fill = HEADER
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)
        cell.border = THIN


def autosize(ws, min_w=10, max_w=28):
    for col in ws.columns:
        letter = get_column_letter(col[0].column)
        length = 0
        for cell in col:
            if cell.value is not None:
                length = max(length, min(len(str(cell.value)), max_w))
        ws.column_dimensions[letter].width = max(min_w, length + 2)


def write_data_workbook(df: pd.DataFrame, cal: pd.DataFrame, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()

    # sales_data
    ws = wb.active
    ws.title = "sales_data"
    out = df.copy()
    out["Date"] = pd.to_datetime(out["Date"])
    for r in dataframe_to_rows(out, index=False, header=True):
        ws.append(r)
    style_header_row(ws, 1, 1, len(out.columns))
    for r in range(2, ws.max_row + 1):
        ws.cell(r, 2).number_format = "YYYY-MM-DD"
        ws.cell(r, 12).number_format = '#,##0.00'
        ws.cell(r, 13).number_format = '#,##0.00'
    autosize(ws)
    ws.auto_filter.ref = ws.dimensions
    ws.freeze_panes = "A2"

    # Calendar
    ws2 = wb.create_sheet("Calendar")
    cal2 = cal.copy()
    cal2["Date"] = pd.to_datetime(cal2["Date"])
    for r in dataframe_to_rows(cal2, index=False, header=True):
        ws2.append(r)
    style_header_row(ws2, 1, 1, len(cal2.columns))
    for r in range(2, ws2.max_row + 1):
        ws2.cell(r, 1).number_format = "YYYY-MM-DD"
    autosize(ws2)
    ws2.freeze_panes = "A2"

    wb.save(path)


def write_cover(wb: Workbook, insights: dict):
    ws = wb.create_sheet("00_Cover", 0)
    for c in range(1, 8):
        ws.column_dimensions[get_column_letter(c)].width = 16
    ws.merge_cells("B2:F2")
    ws["B2"] = "NORTHLINE CONSUMER PRODUCTS"
    ws["B2"].font = COVER_TITLE
    ws["B2"].fill = COVER_BG
    ws["B2"].alignment = Alignment(horizontal="center", vertical="center")
    for c in range(2, 7):
        ws.cell(2, c).fill = COVER_BG
    ws.row_dimensions[2].height = 36

    ws.merge_cells("B3:F3")
    ws["B3"] = "Sales Performance & Time-Intelligence Dashboard"
    ws["B3"].font = Font(name="Calibri", size=16, bold=True, color="1F4E79")
    ws["B3"].alignment = Alignment(horizontal="center")

    ws.merge_cells("B5:F5")
    ws["B5"] = (
        "Portfolio sample for Sai Siri Bandaru — FP&A / Sales Analytics. "
        "Excel workbook recreates the Executive, MTD, and SPLY views typically built in Power BI Desktop. "
        "Yellow cells = filter inputs; black = formulas / computed values."
    )
    ws["B5"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[5].height = 60

    ws["B7"] = "Package contents"
    ws["B7"].font = SECTION_FONT
    items = [
        "01_Executive — KPI cards, channel / brand / SKU / settlement / city / rating views",
        "02_MTD — yellow Year/Month inputs, daily sales + cumulative MTD chart",
        "03_SPLY — year / quarter / month same-period-last-year comparisons",
        "04_Measures — Excel formula equivalents + DAX to paste into Power BI",
        "05_Data_Dictionary — field definitions for sales_data and Calendar",
        "Source data: data/Northline_Sales_Data.xlsx (sales_data + Calendar sheets)",
    ]
    for i, t in enumerate(items):
        ws.cell(8 + i, 2, t).font = BLACK

    ws["B15"] = "Headline KPIs (all periods)"
    ws["B15"].font = SECTION_FONT
    labels = ["Total Sales", "Total Quantity", "Transactions", "Avg Price / Unit"]
    values = [
        fmt_money(insights["total_sales"]),
        fmt_qty(insights["total_qty"]),
        f"{insights['n_txn']:,}",
        f"${insights['avg_price']:.2f}",
    ]
    for i, (lab, val) in enumerate(zip(labels, values)):
        cell_l = ws.cell(16, 2 + i, lab)
        cell_l.fill = TILE
        cell_l.font = KPI_LABEL
        cell_l.alignment = Alignment(horizontal="center")
        cell_l.border = THIN
        cell_v = ws.cell(17, 2 + i, val)
        cell_v.fill = TILE
        cell_v.font = KPI_VALUE
        cell_v.alignment = Alignment(horizontal="center")
        cell_v.border = THIN

    ws["B19"] = f"Data window: {insights['date_min']} → {insights['date_max']}"
    ws["B19"].font = Font(name="Calibri", size=10, italic=True, color="666666")
    ws["B20"] = "All figures are fictional sample data for portfolio demonstration."
    ws["B20"].font = Font(name="Calibri", size=10, italic=True, color="666666")


def write_executive(wb: Workbook, df: pd.DataFrame, insights: dict):
    ws = wb.create_sheet("01_Executive")
    for c in range(1, 14):
        ws.column_dimensions[get_column_letter(c)].width = 14
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 22

    ws.merge_cells("B2:G2")
    ws["B2"] = "Executive Sales Dashboard — Northline Consumer Products"
    ws["B2"].font = TITLE_FONT

    # KPI cards
    kpis = [
        ("Total Sales", insights["total_sales"], '"$"#,##0'),
        ("Total Quantity", insights["total_qty"], "#,##0"),
        ("Transactions", insights["n_txn"], "#,##0"),
        ("Avg Price", insights["avg_price"], '"$"#,##0.00'),
    ]
    ws["B4"] = "KPI CARDS"
    ws["B4"].font = SECTION_FONT
    for i, (lab, val, fmt) in enumerate(kpis):
        col = 2 + i
        ws.cell(5, col, lab).font = KPI_LABEL
        ws.cell(5, col).fill = TILE
        ws.cell(5, col).alignment = Alignment(horizontal="center")
        ws.cell(5, col).border = THIN
        c = ws.cell(6, col, val)
        c.font = KPI_VALUE
        c.fill = TILE
        c.number_format = fmt
        c.alignment = Alignment(horizontal="center")
        c.border = THIN

    # --- Sales by Channel ---
    ws["B8"] = "Sales by Channel"
    ws["B8"].font = SECTION_FONT
    ws["B9"] = "Channel"
    ws["C9"] = "Sales $"
    ws["D9"] = "Units"
    ws["E9"] = "Txns"
    style_header_row(ws, 9, 2, 5)
    ch = df.groupby("Channel").agg(
        Sales=("Sales Amount", "sum"),
        Units=("Units Sold", "sum"),
        Txns=("Transaction ID", "count"),
    ).reindex(CHANNELS)
    for i, (ch_name, row) in enumerate(ch.iterrows()):
        ws.cell(10 + i, 2, ch_name).font = BLACK
        ws.cell(10 + i, 3, float(row["Sales"])).number_format = '"$"#,##0'
        ws.cell(10 + i, 4, int(row["Units"])).number_format = "#,##0"
        ws.cell(10 + i, 5, int(row["Txns"])).number_format = "#,##0"
        for c in range(2, 6):
            ws.cell(10 + i, c).border = THIN

    chart1 = BarChart()
    chart1.type = "col"
    chart1.title = "Sales by Channel"
    chart1.style = 10
    chart1.y_axis.title = "Sales $"
    data = Reference(ws, min_col=3, min_row=9, max_row=14)
    cats = Reference(ws, min_col=2, min_row=10, max_row=14)
    chart1.add_data(data, titles_from_data=True)
    chart1.set_categories(cats)
    chart1.shape = 4
    chart1.width = 12
    chart1.height = 8
    ws.add_chart(chart1, "G8")

    # --- Sales by Brand ---
    ws["B16"] = "Sales by Brand / Line"
    ws["B16"].font = SECTION_FONT
    ws["B17"] = "Brand"
    ws["C17"] = "Sales $"
    ws["D17"] = "Units"
    ws["E17"] = "Txns"
    style_header_row(ws, 17, 2, 5)
    brands_order = ["Northline Core", "Northline Fresh", "Northline Pro"]
    br = df.groupby("Brand").agg(
        Sales=("Sales Amount", "sum"),
        Units=("Units Sold", "sum"),
        Txns=("Transaction ID", "count"),
    ).reindex(brands_order)
    for i, (name, row) in enumerate(br.iterrows()):
        ws.cell(18 + i, 2, name)
        ws.cell(18 + i, 3, float(row["Sales"])).number_format = '"$"#,##0'
        ws.cell(18 + i, 4, int(row["Units"])).number_format = "#,##0"
        ws.cell(18 + i, 5, int(row["Txns"])).number_format = "#,##0"
        for c in range(2, 6):
            ws.cell(18 + i, c).border = THIN

    chart2 = BarChart()
    chart2.type = "bar"
    chart2.title = "Sales by Brand"
    chart2.style = 10
    data2 = Reference(ws, min_col=3, min_row=17, max_row=20)
    cats2 = Reference(ws, min_col=2, min_row=18, max_row=20)
    chart2.add_data(data2, titles_from_data=True)
    chart2.set_categories(cats2)
    chart2.width = 12
    chart2.height = 7
    ws.add_chart(chart2, "G18")

    # --- Top SKUs ---
    ws["B22"] = "Top 10 Products by Sales"
    ws["B22"].font = SECTION_FONT
    ws["B23"] = "Product"
    ws["C23"] = "Sales $"
    ws["D23"] = "Units"
    style_header_row(ws, 23, 2, 4)
    top = (
        df.groupby("Product")
        .agg(Sales=("Sales Amount", "sum"), Units=("Units Sold", "sum"))
        .sort_values("Sales", ascending=False)
        .head(10)
    )
    for i, (name, row) in enumerate(top.iterrows()):
        ws.cell(24 + i, 2, name)
        ws.cell(24 + i, 3, float(row["Sales"])).number_format = '"$"#,##0'
        ws.cell(24 + i, 4, int(row["Units"])).number_format = "#,##0"
        for c in range(2, 5):
            ws.cell(24 + i, c).border = THIN
    ws.column_dimensions["B"].width = 28

    chart3 = BarChart()
    chart3.type = "bar"
    chart3.title = "Top Products by Sales"
    chart3.style = 10
    data3 = Reference(ws, min_col=3, min_row=23, max_row=33)
    cats3 = Reference(ws, min_col=2, min_row=24, max_row=33)
    chart3.add_data(data3, titles_from_data=True)
    chart3.set_categories(cats3)
    chart3.width = 14
    chart3.height = 10
    ws.add_chart(chart3, "G28")

    # --- Settlement mix ---
    ws["B36"] = "Settlement Method Mix"
    ws["B36"].font = SECTION_FONT
    ws["B37"] = "Settlement"
    ws["C37"] = "Txns"
    ws["D37"] = "% of Txns"
    style_header_row(ws, 37, 2, 4)
    sett = df["Settlement Method"].value_counts()
    for i, (name, cnt) in enumerate(sett.items()):
        ws.cell(38 + i, 2, name)
        ws.cell(38 + i, 3, int(cnt)).number_format = "#,##0"
        ws.cell(38 + i, 4, cnt / len(df)).number_format = "0.0%"
        for c in range(2, 5):
            ws.cell(38 + i, c).border = THIN

    pie = PieChart()
    pie.title = "Settlement Mix"
    labels = Reference(ws, min_col=2, min_row=38, max_row=37 + len(sett))
    data_p = Reference(ws, min_col=3, min_row=37, max_row=37 + len(sett))
    pie.add_data(data_p, titles_from_data=True)
    pie.set_categories(labels)
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showVal = False
    pie.dataLabels.showCatName = False
    pie.width = 10
    pie.height = 8
    ws.add_chart(pie, "F36")

    # --- City volume ---
    ws["B45"] = "Top Cities by Units Sold"
    ws["B45"].font = SECTION_FONT
    ws["B46"] = "City"
    ws["C46"] = "Region"
    ws["D46"] = "Units"
    ws["E46"] = "Sales $"
    style_header_row(ws, 46, 2, 5)
    city_reg = df.groupby(["City", "Region"]).agg(
        Units=("Units Sold", "sum"), Sales=("Sales Amount", "sum")
    ).sort_values("Units", ascending=False).head(10)
    for i, ((city, region), row) in enumerate(city_reg.iterrows()):
        ws.cell(47 + i, 2, city)
        ws.cell(47 + i, 3, region)
        ws.cell(47 + i, 4, int(row["Units"])).number_format = "#,##0"
        ws.cell(47 + i, 5, float(row["Sales"])).number_format = '"$"#,##0'
        for c in range(2, 6):
            ws.cell(47 + i, c).border = THIN

    # --- Rating mix ---
    ws["B59"] = "Customer Rating / NPS Bucket"
    ws["B59"].font = SECTION_FONT
    ws["B60"] = "Rating"
    ws["C60"] = "Txns"
    ws["D60"] = "% of Txns"
    style_header_row(ws, 60, 2, 4)
    for i, name in enumerate(RATINGS):
        cnt = int((df["Customer Rating"] == name).sum())
        ws.cell(61 + i, 2, name)
        ws.cell(61 + i, 3, cnt).number_format = "#,##0"
        ws.cell(61 + i, 4, cnt / len(df)).number_format = "0.0%"
        for c in range(2, 5):
            ws.cell(61 + i, c).border = THIN

    # --- Order type ---
    ws["B66"] = "Sales by Order Type"
    ws["B66"].font = SECTION_FONT
    ws["B67"] = "Order Type"
    ws["C67"] = "Sales $"
    ws["D67"] = "Txns"
    style_header_row(ws, 67, 2, 4)
    ot = df.groupby("Order Type").agg(
        Sales=("Sales Amount", "sum"), Txns=("Transaction ID", "count")
    ).sort_values("Sales", ascending=False)
    for i, (name, row) in enumerate(ot.iterrows()):
        ws.cell(68 + i, 2, name)
        ws.cell(68 + i, 3, float(row["Sales"])).number_format = '"$"#,##0'
        ws.cell(68 + i, 4, int(row["Txns"])).number_format = "#,##0"
        for c in range(2, 5):
            ws.cell(68 + i, c).border = THIN

    ws["B72"] = "Note: In Power BI, use slicers on Channel, Brand, Product, Settlement Method, and City."
    ws["B72"].font = Font(name="Calibri", size=9, italic=True, color="666666")


def write_mtd(wb: Workbook, df: pd.DataFrame):
    ws = wb.create_sheet("02_MTD")
    for c in range(1, 10):
        ws.column_dimensions[get_column_letter(c)].width = 14
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 18

    ws.merge_cells("B2:F2")
    ws["B2"] = "Month-to-Date (MTD) — Daily Cumulative Sales"
    ws["B2"].font = TITLE_FONT

    ws["B4"] = "FILTER INPUTS (yellow)"
    ws["B4"].font = SECTION_FONT
    ws["B5"] = "Year"
    ws["C5"] = 2025
    ws["C5"].fill = YELLOW
    ws["C5"].font = INPUT_FONT
    ws["C5"].border = THIN
    ws["B6"] = "Month (1-12)"
    ws["C6"] = 6
    ws["C6"].fill = YELLOW
    ws["C6"].font = INPUT_FONT
    ws["C6"].border = THIN
    ws["D5"] = "← change these yellow cells to filter MTD view"
    ws["D5"].font = Font(name="Calibri", size=9, italic=True, color="666666")

    # Pre-compute default month (Jun 2025) daily table; also store helper pivot for all months
    # Build a lookup block for SUMIFS-style formulas against a data dump
    # We'll put a compact daily table for the selected default, AND a formula-driven approach
    # using a cached monthly extract starting at row 40.

    # Data cache for formulas: Year, Month, Day, Sales
    daily = (
        df.groupby(["Year", "Month", "Day"])["Sales Amount"]
        .sum()
        .reset_index()
        .sort_values(["Year", "Month", "Day"])
    )

    ws["B8"] = "MTD KPIs (for Year/Month above — uses SUMIFS on cache)"
    ws["B8"].font = SECTION_FONT
    ws["B9"] = "MTD Sales"
    ws["C9"] = '=SUMIFS($C$40:$C$5000,$A$40:$A$5000,$C$5,$B$40:$B$5000,$C$6)'
    ws["C9"].font = BLACK
    ws["C9"].number_format = '"$"#,##0.00'
    ws["C9"].border = THIN
    ws["B10"] = "Days with Sales"
    ws["C10"] = '=COUNTIFS($A$40:$A$5000,$C$5,$B$40:$B$5000,$C$6)'
    ws["C10"].font = BLACK
    ws["C10"].border = THIN
    ws["B11"] = "Avg Daily Sales"
    ws["C11"] = '=IF(C10=0,0,C9/C10)'
    ws["C11"].font = BLACK
    ws["C11"].number_format = '"$"#,##0.00'
    ws["C11"].border = THIN

    # Visible daily table for default filter (Jun 2025) — also formula-filtered via cache
    ws["B13"] = "Daily Sales + Cumulative (filtered by Year/Month inputs)"
    ws["B13"].font = SECTION_FONT
    ws["B14"] = "Day"
    ws["C14"] = "Sales $"
    ws["D14"] = "Cumulative MTD"
    style_header_row(ws, 14, 2, 4)

    # Days 1-31 with SUMIFS / running total formulas
    for day in range(1, 32):
        r = 14 + day
        ws.cell(r, 2, day).border = THIN
        # Sales for that day in selected year/month
        ws.cell(
            r, 3,
            f'=IFERROR(SUMIFS($C$40:$C$5000,$A$40:$A$5000,$C$5,$B$40:$B$5000,$C$6,$D$40:$D$5000,B{r}),0)',
        )
        ws.cell(r, 3).number_format = '"$"#,##0.00'
        ws.cell(r, 3).font = BLACK
        ws.cell(r, 3).border = THIN
        if day == 1:
            ws.cell(r, 4, f"=C{r}")
        else:
            ws.cell(r, 4, f"=D{r-1}+C{r}")
        ws.cell(r, 4).number_format = '"$"#,##0.00'
        ws.cell(r, 4).font = BLACK
        ws.cell(r, 4).border = THIN

    chart = LineChart()
    chart.title = "Daily Sales vs Cumulative MTD"
    chart.style = 10
    chart.y_axis.title = "Sales $"
    chart.x_axis.title = "Day of Month"
    data = Reference(ws, min_col=3, min_row=14, max_col=4, max_row=45)
    cats = Reference(ws, min_col=2, min_row=15, max_row=45)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.width = 15
    chart.height = 10
    ws.add_chart(chart, "F8")

    # Cache block starting row 40
    ws["A39"] = "Year"
    ws["B39"] = "Month"
    ws["C39"] = "Sales"
    ws["D39"] = "Day"
    style_header_row(ws, 39, 1, 4)
    for i, row in daily.iterrows():
        r = 40 + i
        # i may not be contiguous after reset — use enumerate instead
    for idx, row in enumerate(daily.itertuples(index=False)):
        r = 40 + idx
        ws.cell(r, 1, int(row[0]))  # Year
        ws.cell(r, 2, int(row[1]))  # Month
        ws.cell(r, 3, float(row[3])).number_format = '"$"#,##0.00'  # Sales — wait order is Y,M,D,Sales
        ws.cell(r, 4, int(row[2]))  # Day

    # Fix: itertuples order is Year, Month, Day, Sales Amount
    # Actually after reset_index columns are Year, Month, Day, Sales Amount
    # Let me rewrite the cache cleanly by clearing and rewriting
    # (openpyxl already wrote — need to fix column mapping)
    # Row tuple: (Year, Month, Day, Sales Amount) indices 0,1,2,3
    # I wrote: A=Year(0), B=Month(1), C=Sales(3), D=Day(2) — CORRECT

    ws["B48"] = (
        "How to use: set Year and Month in the yellow cells. Daily sales, cumulative MTD, and KPIs update via SUMIFS. "
        "In Power BI, replace this with TOTALMTD([Total_Sales], Calendar[Date]) and a month slicer."
    )
    ws["B48"].font = Font(name="Calibri", size=9, italic=True, color="666666")
    ws.merge_cells("B48:H48")


def write_sply(wb: Workbook, df: pd.DataFrame):
    ws = wb.create_sheet("03_SPLY")
    for c in range(1, 12):
        ws.column_dimensions[get_column_letter(c)].width = 14
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 16

    ws.merge_cells("B2:H2")
    ws["B2"] = "Same Period Last Year (SPLY) — YoY Comparisons"
    ws["B2"].font = TITLE_FONT

    # Year comparison
    ws["B4"] = "By Year"
    ws["B4"].font = SECTION_FONT
    headers = ["Year", "Sales $", "Prior Year (SPLY)", "YoY $", "YoY %"]
    for i, h in enumerate(headers):
        ws.cell(5, 2 + i, h)
    style_header_row(ws, 5, 2, 6)

    yearly = df.groupby("Year")["Sales Amount"].sum().sort_index()
    years = list(yearly.index)
    for i, y in enumerate(years):
        r = 6 + i
        ws.cell(r, 2, int(y)).border = THIN
        ws.cell(r, 3, float(yearly[y])).number_format = '"$"#,##0'
        ws.cell(r, 3).border = THIN
        if i == 0:
            ws.cell(r, 4, None).border = THIN
            ws.cell(r, 5, None).border = THIN
            ws.cell(r, 6, None).border = THIN
        else:
            # Formulas referencing prior row
            ws.cell(r, 4, f"=C{r-1}")
            ws.cell(r, 4).number_format = '"$"#,##0'
            ws.cell(r, 4).font = BLACK
            ws.cell(r, 4).border = THIN
            ws.cell(r, 5, f"=C{r}-D{r}")
            ws.cell(r, 5).number_format = '"$"#,##0'
            ws.cell(r, 5).font = BLACK
            ws.cell(r, 5).border = THIN
            ws.cell(r, 6, f"=IF(D{r}=0,NA(),C{r}/D{r}-1)")
            ws.cell(r, 6).number_format = "0.0%"
            ws.cell(r, 6).font = BLACK
            ws.cell(r, 6).border = THIN

    chart_y = BarChart()
    chart_y.type = "col"
    chart_y.grouping = "clustered"
    chart_y.title = "Sales vs Prior Year"
    chart_y.style = 10
    last_r = 5 + len(years)
    data = Reference(ws, min_col=3, min_row=5, max_col=4, max_row=last_r)
    cats = Reference(ws, min_col=2, min_row=6, max_row=last_r)
    chart_y.add_data(data, titles_from_data=True)
    chart_y.set_categories(cats)
    chart_y.width = 12
    chart_y.height = 8
    ws.add_chart(chart_y, "H4")

    # Quarterly
    start_q = last_r + 3
    ws.cell(start_q, 2, "By Year-Quarter").font = SECTION_FONT
    q_headers = ["Year", "Quarter", "Sales $", "SPLY $", "YoY $", "YoY %"]
    for i, h in enumerate(q_headers):
        ws.cell(start_q + 1, 2 + i, h)
    style_header_row(ws, start_q + 1, 2, 7)

    tmp = df.copy()
    tmp["Q"] = tmp["Month"].map(lambda m: f"Q{((m - 1) // 3) + 1}")
    qtr = tmp.groupby(["Year", "Q"])["Sales Amount"].sum()
    q_rows = []
    for (y, q), sales in qtr.items():
        prior = qtr.get((y - 1, q), None)
        q_rows.append((int(y), q, float(sales), float(prior) if prior is not None else None))

    for i, (y, q, sales, prior) in enumerate(q_rows):
        r = start_q + 2 + i
        ws.cell(r, 2, y).border = THIN
        ws.cell(r, 3, q).border = THIN
        ws.cell(r, 4, sales).number_format = '"$"#,##0'
        ws.cell(r, 4).border = THIN
        if prior is None:
            for c in range(5, 8):
                ws.cell(r, c).border = THIN
        else:
            ws.cell(r, 5, prior).number_format = '"$"#,##0'
            ws.cell(r, 5).border = THIN
            ws.cell(r, 6, f"=D{r}-E{r}")
            ws.cell(r, 6).number_format = '"$"#,##0'
            ws.cell(r, 6).font = BLACK
            ws.cell(r, 6).border = THIN
            ws.cell(r, 7, f"=IF(E{r}=0,NA(),D{r}/E{r}-1)")
            ws.cell(r, 7).number_format = "0.0%"
            ws.cell(r, 7).font = BLACK
            ws.cell(r, 7).border = THIN

    # Monthly (latest 24 months with SPLY)
    start_m = start_q + 2 + len(q_rows) + 2
    ws.cell(start_m, 2, "By Year-Month (with SPLY)").font = SECTION_FONT
    m_headers = ["Year", "Month", "Month Name", "Sales $", "SPLY $", "YoY $", "YoY %"]
    for i, h in enumerate(m_headers):
        ws.cell(start_m + 1, 2 + i, h)
    style_header_row(ws, start_m + 1, 2, 8)

    monthly = df.groupby(["Year", "Month"])["Sales Amount"].sum()
    month_names = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
                   7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}
    m_keys = sorted(monthly.keys())
    for i, (y, m) in enumerate(m_keys):
        r = start_m + 2 + i
        sales = float(monthly[(y, m)])
        prior = monthly.get((y - 1, m), None)
        ws.cell(r, 2, int(y)).border = THIN
        ws.cell(r, 3, int(m)).border = THIN
        ws.cell(r, 4, month_names[m]).border = THIN
        ws.cell(r, 5, sales).number_format = '"$"#,##0'
        ws.cell(r, 5).border = THIN
        if prior is None:
            for c in range(6, 9):
                ws.cell(r, c).border = THIN
        else:
            ws.cell(r, 6, float(prior)).number_format = '"$"#,##0'
            ws.cell(r, 6).border = THIN
            ws.cell(r, 7, f"=E{r}-F{r}")
            ws.cell(r, 7).number_format = '"$"#,##0'
            ws.cell(r, 7).font = BLACK
            ws.cell(r, 7).border = THIN
            ws.cell(r, 8, f"=IF(F{r}=0,NA(),E{r}/F{r}-1)")
            ws.cell(r, 8).number_format = "0.0%"
            ws.cell(r, 8).font = BLACK
            ws.cell(r, 8).border = THIN

    note_r = start_m + 2 + len(m_keys) + 2
    ws.cell(note_r, 2,
            "SPLY = same calendar period in the prior year. Black cells are Excel formulas. "
            "In Power BI use CALCULATE([Total_Sales], SAMEPERIODLASTYEAR(Calendar[Date])).").font = Font(
        name="Calibri", size=9, italic=True, color="666666"
    )


def write_measures(wb: Workbook):
    ws = wb.create_sheet("04_Measures")
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 55
    ws.column_dimensions["D"].width = 55

    ws["B2"] = "Measure Definitions — Excel Equivalents + DAX (paste into Power BI)"
    ws["B2"].font = TITLE_FONT
    ws.merge_cells("B2:D2")

    ws["B4"] = "Measure"
    ws["C4"] = "Excel / this workbook"
    ws["D4"] = "DAX (Power BI Desktop)"
    style_header_row(ws, 4, 2, 4)

    measures = [
        (
            "Total_Quantity",
            "SUM of Units Sold",
            "Total_Quantity = SUM(sales_data[Units Sold])",
        ),
        (
            "Total_Sales",
            "SUM of (Units Sold × Price Per Unit) — or SUM of Sales Amount",
            "Total_Sales =\nSUMX(\n    sales_data,\n    sales_data[Units Sold] * sales_data[Price Per Unit]\n)",
        ),
        (
            "Transactions",
            "COUNT of Transaction ID / COUNTA of rows",
            "Transactions = COUNTROWS(sales_data)",
        ),
        (
            "Average_price",
            "Total_Sales ÷ Total_Quantity",
            "Average_price = DIVIDE([Total_Sales], [Total_Quantity], 0)",
        ),
        (
            "MTD",
            "02_MTD! cumulative column / SUMIFS for selected Year+Month through Day",
            "MTD =\nTOTALMTD(\n    [Total_Sales],\n    Calendar[Date]\n)",
        ),
        (
            "Same Period Last Year",
            "03_SPLY! prior-year column via YoY formulas",
            "Same Period Last Year =\nCALCULATE(\n    [Total_Sales],\n    SAMEPERIODLASTYEAR(Calendar[Date])\n)",
        ),
        (
            "YoY Growth %",
            "(Current − SPLY) ÷ SPLY",
            "YoY_Growth =\nDIVIDE([Total_Sales] - [Same Period Last Year], [Same Period Last Year], 0)",
        ),
        (
            "YTD Sales",
            "SUMIFS Year = selected year, Date ≤ as-of",
            "YTD_Sales =\nTOTALYTD([Total_Sales], Calendar[Date])",
        ),
    ]
    for i, (name, excel, dax) in enumerate(measures):
        r = 5 + i
        ws.cell(r, 2, name).font = Font(name="Calibri", size=11, bold=True)
        ws.cell(r, 2).alignment = Alignment(vertical="top")
        ws.cell(r, 2).border = THIN
        ws.cell(r, 3, excel).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 3).border = THIN
        ws.cell(r, 4, dax).alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(r, 4).font = Font(name="Consolas", size=9)
        ws.cell(r, 4).border = THIN
        ws.row_dimensions[r].height = 55

    ws["B14"] = "Model relationship"
    ws["B14"].font = SECTION_FONT
    ws["B15"] = (
        "Relate Calendar[Date] (1) → sales_data[Date] (many). Mark Calendar as a date table. "
        "Table names assumed: sales_data and Calendar (see data/Northline_Sales_Data.xlsx)."
    )
    ws["B15"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("B15:D15")
    ws.row_dimensions[15].height = 40


def write_dictionary(wb: Workbook):
    ws = wb.create_sheet("05_Data_Dictionary")
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 60

    ws["B2"] = "Data Dictionary"
    ws["B2"].font = TITLE_FONT

    ws["B4"] = "sales_data"
    ws["B4"].font = SECTION_FONT
    headers = ["Field", "Type", "Description"]
    for i, h in enumerate(headers):
        ws.cell(5, 2 + i, h)
    style_header_row(ws, 5, 2, 4)

    fields = [
        ("Transaction ID", "Integer", "Unique order / shipment line identifier"),
        ("Date", "Date", "Transaction / ship date (relate to Calendar[Date])"),
        ("Day / Month / Year", "Integer", "Date parts for Excel pivots and filters"),
        ("Day Name", "Text", "Weekday name"),
        ("Channel", "Text", "Grocery | Mass | Club | E-comm | Convenience"),
        ("Brand", "Text", "Northline Core | Northline Fresh | Northline Pro"),
        ("SKU", "Text", "Internal product code"),
        ("Product", "Text", "Commercial product name"),
        ("Units Sold", "Integer", "Case / unit quantity on the order"),
        ("Price Per Unit", "Decimal", "Net realized price per unit"),
        ("Sales Amount", "Decimal", "Units Sold × Price Per Unit"),
        ("Customer Name", "Text", "Retail / trade account name"),
        ("City", "Text", "US metro of the ship-to / store cluster"),
        ("Region", "Text", "Northeast | Midwest | South | West"),
        ("Order Type", "Text", "Wholesale | DSD | E-comm"),
        ("Settlement Method", "Text", "How the invoice was settled"),
        ("Customer Rating", "Text", "NPS-style bucket from account survey"),
        ("Sales Rep", "Text", "Assigned Northline sales representative"),
    ]
    for i, (f, t, d) in enumerate(fields):
        r = 6 + i
        ws.cell(r, 2, f).border = THIN
        ws.cell(r, 3, t).border = THIN
        ws.cell(r, 4, d).border = THIN

    r0 = 6 + len(fields) + 2
    ws.cell(r0, 2, "Calendar").font = SECTION_FONT
    for i, h in enumerate(headers):
        ws.cell(r0 + 1, 2 + i, h)
    style_header_row(ws, r0 + 1, 2, 4)
    cal_fields = [
        ("Date", "Date", "Continuous daily date spine"),
        ("Year / Quarter / Month", "Text/Int", "Standard time-intelligence attributes"),
        ("Month Name", "Text", "Jan–Dec"),
        ("Day / Day Name", "Int/Text", "Day of month and weekday"),
        ("Weeknum", "Integer", "ISO week number"),
        ("YearMonth", "Text", "YYYY-MM label for charts"),
    ]
    for i, (f, t, d) in enumerate(cal_fields):
        r = r0 + 2 + i
        ws.cell(r, 2, f).border = THIN
        ws.cell(r, 3, t).border = THIN
        ws.cell(r, 4, d).border = THIN


def write_dashboard(df: pd.DataFrame, insights: dict, path: Path):
    wb = Workbook()
    # remove default; cover will be index 0
    default = wb.active
    wb.remove(default)
    write_cover(wb, insights)
    write_executive(wb, df, insights)
    write_mtd(wb, df)
    write_sply(wb, df)
    write_measures(wb)
    write_dictionary(wb)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def write_svg_preview(insights: dict, path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    brands = insights["brand"]
    channels = insights["channel"]
    max_ch = max(c["Sales"] for c in channels) or 1

    def bar_rows(items, key="Sales", label="Brand", y0=160):
        parts = []
        for i, it in enumerate(items[:5]):
            w = 280 * (it[key] / max(max(x[key] for x in items), 1))
            y = y0 + i * 28
            parts.append(
                f'<text x="40" y="{y + 14}" font-family="Segoe UI,Arial" font-size="11" fill="#333">{it[label]}</text>'
                f'<rect x="160" y="{y}" width="{w:.1f}" height="18" rx="3" fill="#1F4E79"/>'
                f'<text x="{170 + w:.1f}" y="{y + 14}" font-family="Segoe UI,Arial" font-size="10" fill="#555">{fmt_money(it[key])}</text>'
            )
        return "\n".join(parts)

    ch_bars = []
    for i, c in enumerate(channels):
        w = 200 * (c["Sales"] / max_ch)
        x = 40 + i * 95
        ch_bars.append(
            f'<rect x="{x}" y="{280 - w * 0.6:.1f}" width="60" height="{w * 0.6:.1f}" rx="3" fill="#2E75B6"/>'
            f'<text x="{x + 30}" y="295" text-anchor="middle" font-family="Segoe UI,Arial" font-size="9" fill="#333">{c["Channel"][:6]}</text>'
        )

    svg = f'''<?xml version="1.0" encoding="UTF-8"?>
<svg xmlns="http://www.w3.org/2000/svg" width="900" height="480" viewBox="0 0 900 480">
  <rect width="900" height="480" fill="#F7F9FC"/>
  <rect width="900" height="56" fill="#1F4E79"/>
  <text x="24" y="36" font-family="Segoe UI,Arial" font-size="20" font-weight="700" fill="#FFFFFF">Northline — Executive Sales Dashboard</text>
  <text x="870" y="36" text-anchor="end" font-family="Segoe UI,Arial" font-size="11" fill="#D6E3F0">Portfolio preview</text>

  <!-- KPI cards -->
  <rect x="24" y="72" width="200" height="64" rx="6" fill="#FFFFFF" stroke="#D6E3F0"/>
  <text x="40" y="94" font-family="Segoe UI,Arial" font-size="11" fill="#1F4E79">Total Sales</text>
  <text x="40" y="120" font-family="Segoe UI,Arial" font-size="22" font-weight="700" fill="#111">{fmt_money(insights["total_sales"])}</text>

  <rect x="240" y="72" width="200" height="64" rx="6" fill="#FFFFFF" stroke="#D6E3F0"/>
  <text x="256" y="94" font-family="Segoe UI,Arial" font-size="11" fill="#1F4E79">Total Quantity</text>
  <text x="256" y="120" font-family="Segoe UI,Arial" font-size="22" font-weight="700" fill="#111">{fmt_qty(insights["total_qty"])}</text>

  <rect x="456" y="72" width="200" height="64" rx="6" fill="#FFFFFF" stroke="#D6E3F0"/>
  <text x="472" y="94" font-family="Segoe UI,Arial" font-size="11" fill="#1F4E79">Transactions</text>
  <text x="472" y="120" font-family="Segoe UI,Arial" font-size="22" font-weight="700" fill="#111">{insights["n_txn"]:,}</text>

  <rect x="672" y="72" width="200" height="64" rx="6" fill="#FFFFFF" stroke="#D6E3F0"/>
  <text x="688" y="94" font-family="Segoe UI,Arial" font-size="11" fill="#1F4E79">Avg Price / Unit</text>
  <text x="688" y="120" font-family="Segoe UI,Arial" font-size="22" font-weight="700" fill="#111">${insights["avg_price"]:.2f}</text>

  <text x="40" y="168" font-family="Segoe UI,Arial" font-size="13" font-weight="700" fill="#1F4E79">Sales by Brand</text>
  {bar_rows(brands, label="Brand", y0=178)}

  <text x="40" y="320" font-family="Segoe UI,Arial" font-size="13" font-weight="700" fill="#1F4E79">Sales by Channel</text>
  {''.join(ch_bars)}

  <text x="520" y="168" font-family="Segoe UI,Arial" font-size="13" font-weight="700" fill="#1F4E79">Top Products</text>
'''
    for i, p in enumerate(insights["product_top"][:5]):
        y = 190 + i * 24
        svg += f'<text x="520" y="{y}" font-family="Segoe UI,Arial" font-size="11" fill="#333">{i+1}. {p["Product"][:32]}</text>\n'
        svg += f'<text x="860" y="{y}" text-anchor="end" font-family="Segoe UI,Arial" font-size="11" fill="#1F4E79">{fmt_money(p["Sales"])}</text>\n'

    svg += f'''
  <text x="24" y="460" font-family="Segoe UI,Arial" font-size="10" fill="#888">Fictional CPG sample · {insights["date_min"]} to {insights["date_max"]} · Sai Siri Bandaru FP&amp;A portfolio</text>
</svg>
'''
    path.write_text(svg, encoding="utf-8")


def write_readme(insights: dict, path: Path):
    brand = insights["brand"]
    products = insights["product_top"]
    channel = insights["channel"]
    sett = insights["settlement_pct"]
    cities = insights["city_top"]
    rating = insights["rating_pct"]
    yearly = insights["yearly"]

    # Quarterly peaks
    q_by_label = {}
    for row in insights["quarterly"]:
        label = f"{row['Year']} {row['Quarter']}"
        q_by_label[label] = row["Sales"]
    # Aggregate by quarter letter across years for seasonality note
    q_agg = {"Q1": 0.0, "Q2": 0.0, "Q3": 0.0, "Q4": 0.0}
    for row in insights["quarterly"]:
        q_agg[row["Quarter"]] += row["Sales"]
    peak_q = max(q_agg, key=q_agg.get)
    peak_q2 = sorted(q_agg.items(), key=lambda x: -x[1])[1][0]

    brand_lines = []
    for i, b in enumerate(brand):
        brand_lines.append(
            f"**{b['Brand']}** at **{fmt_money(b['Sales'])}** ({b['Txns']:,} transactions)"
        )
    brand_sentence = ", followed by ".join(
        [brand_lines[0], ", and ".join(brand_lines[1:3])]
    ) if len(brand_lines) >= 3 else "; ".join(brand_lines)

    prod_bits = ", ".join(
        f"**{p['Product']}** ({fmt_money(p['Sales'])})" for p in products[:3]
    )
    sett_bits = ", ".join(f"**{k}** ({v:.1f}%)" for k, v in sorted(sett.items(), key=lambda x: -x[1]))
    city_bits = ", ".join(f"**{c['City']}** ({c['Qty']:,})" for c in cities[:4])
    ch_bits = ", ".join(f"**{c['Channel']}** ({fmt_money(c['Sales'])})" for c in channel[:3])
    rating_bits = ", ".join(f"**{k}** ({v:.1f}%)" for k, v in rating.items())

    yoy_lines = []
    for y, info in sorted(insights["yoy"].items()):
        if info["pct"] is not None:
            note = f" (Jan–{['','Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][info.get('through_month', 12)]} SPLY)" if info.get("partial") else ""
            # simplify partial note
            if info.get("partial"):
                m = info.get("through_month", 6)
                mon = ["","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"][m]
                note = f" (Jan–{mon} same-period)"
            else:
                note = ""
            yoy_lines.append(f"**{y}** vs prior{note}: **{info['pct']:+.1f}%** ({fmt_money(info['sales'])} vs {fmt_money(info['prior'])})")
    yoy_sentence = "; ".join(yoy_lines) if yoy_lines else "n/a"

    year_bits = ", ".join(f"**{y}**: {fmt_money(s)}" for y, s in sorted(yearly.items()))

    readme = f"""# Northline Sales Performance & Time-Intelligence Dashboard

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

Computed from the generated sample (`{insights["date_min"]}` → `{insights["date_max"]}`, **{insights["n_txn"]:,}** transactions):

| KPI | Value | Description |
| :--- | :--- | :--- |
| **Total Sales** | **{fmt_money(insights["total_sales"])}** ({fmt_money(insights["total_sales"], compact=False)}) | Total gross trade sales across all transactions. |
| **Total Quantity** | **{fmt_qty(insights["total_qty"])}** ({insights["total_qty"]:,}) | Total units / cases sold. |
| **Transactions** | **{insights["n_txn"]:,}** | Completed sales orders / shipment lines. |
| **Average Price** | **${insights["avg_price"]:.2f}** | Weighted average selling price per unit. |

---

## Key Insights from the Data

* **Brand / Line Performance:** {brand_sentence}.
* **Top Selling Products:** Highest revenue SKUs — {prod_bits}.
* **Channel Mix:** Leading channels by sales — {ch_bits}.
* **Settlement Methods:** Mix across {sett_bits}.
* **Top Cities by Volume:** {city_bits}.
* **Customer Rating:** {rating_bits}.
* **Seasonality:** Across the full sample, quarterly sales are highest in **{peak_q}** ({fmt_money(q_agg[peak_q])}) and **{peak_q2}** ({fmt_money(q_agg[peak_q2])}).
* **Yearly Sales:** {year_bits}.
* **YoY (SPLY):** {yoy_sentence}.

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
"""
    path.write_text(readme, encoding="utf-8")


def main():
    print("Generating transactions…")
    df = generate_transactions(4000)
    print(f"  rows={len(df):,}  {df['Date'].min()} → {df['Date'].max()}")

    cal = build_calendar(date(2022, 1, 1), date(2025, 12, 31))
    insights = compute_insights(df)

    DATA_DIR.mkdir(parents=True, exist_ok=True)
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)

    data_path = DATA_DIR / "Northline_Sales_Data.xlsx"
    dash_path = ROOT / "Northline_Sales_Performance_Dashboard.xlsx"
    svg_path = ASSETS_DIR / "executive_preview.svg"
    readme_path = ROOT / "README.md"
    insights_path = ROOT / "insights.json"

    print("Writing data workbook…")
    write_data_workbook(df, cal, data_path)

    print("Writing dashboard workbook…")
    write_dashboard(df, insights, dash_path)

    print("Writing SVG preview…")
    write_svg_preview(insights, svg_path)

    print("Writing README…")
    write_readme(insights, readme_path)

    # Serialize insights (convert numpy types)
    def _jsonable(o):
        if isinstance(o, dict):
            return {str(k): _jsonable(v) for k, v in o.items()}
        if isinstance(o, list):
            return [_jsonable(v) for v in o]
        if isinstance(o, (np.integer,)):
            return int(o)
        if isinstance(o, (np.floating,)):
            return float(o)
        return o

    insights_path.write_text(json.dumps(_jsonable(insights), indent=2), encoding="utf-8")

    print("Done.")
    print(f"  Total Sales: {fmt_money(insights['total_sales'])} ({insights['total_sales']:,.2f})")
    print(f"  Total Qty:   {insights['total_qty']:,}")
    print(f"  Txns:        {insights['n_txn']:,}")
    print(f"  Avg Price:   ${insights['avg_price']:.2f}")
    print(f"  Files: {data_path.name}, {dash_path.name}, {svg_path.name}, README.md")


if __name__ == "__main__":
    main()
