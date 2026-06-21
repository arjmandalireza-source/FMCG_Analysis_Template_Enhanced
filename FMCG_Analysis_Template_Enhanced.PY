"""
FMCG Analytics Learning & Enterprise Template
==============================================
Generates a fully-formula-driven Excel workbook with 126 KPIs.
All KPIs are calculated via live Excel formulas — no hard-coded values.

Sheets:
  Data_Sales       — 5,000 transaction rows
  Data_Customers   — 200 customer master records
  Data_Products    — 50 SKU master records
  Data_Inventory   — 250 warehouse stock records
  Data_Visits      — 3,000 sales visit records
  Data_AR_Aging    — 400 AR invoice records
  Data_External    — 12 monthly external/market records
  Data_Targets     — monthly rep targets
  Assumptions      — WACC, tax rate, margins, etc.
  KPI_Sales        — SM01–SM31
  KPI_Customer     — CA01–CA20
  KPI_Inventory    — IN01–IN12
  KPI_Distribution — DR01–DR16
  KPI_Finance      — FI01–FI17
  KPI_Advanced     — AE01–AE30
  Dashboard        — summary of core KPIs

Requirements:
  pip install openpyxl numpy pandas faker
"""

import random
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
from faker import Faker
from openpyxl import Workbook
from openpyxl.styles import (
    Alignment, Border, Font, PatternFill, Side
)
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

# ─────────────────────────────────────────────
# 0. SEEDS & CONSTANTS
# ─────────────────────────────────────────────
np.random.seed(42)
random.seed(42)
fake = Faker()
Faker.seed(42)

N_SALES    = 5_000
N_VISITS   = 3_000
N_AR       = 400
START_DATE = datetime(2024, 1, 1)
END_DATE   = datetime(2024, 12, 31)

REGIONS    = ["Tehran", "Isfahan", "Mashhad", "Shiraz", "Tabriz"]
CHANNELS   = ["Modern Trade", "Traditional Trade", "E-Commerce", "HoReCa"]
CATEGORIES = ["Dairy", "Beverages", "Snacks", "Personal Care", "Home Care"]
BRANDS     = ["BrandA", "BrandB", "BrandC", "BrandD"]
CUSTOMERS  = [f"CUST_{i:04d}" for i in range(1, 201)]
PRODUCTS   = [f"SKU_{i:03d}" for i in range(1, 51)]
ROUTES     = [f"ROUTE_{i:02d}" for i in range(1, 21)]
SALESREPS  = [f"REP_{i:02d}" for i in range(1, 16)]
WAREHOUSES = ["WH_Tehran", "WH_Isfahan", "WH_Mashhad", "WH_Shiraz"]

# ─────────────────────────────────────────────
# 1. DATA GENERATORS
# ─────────────────────────────────────────────

def rand_dates(n, start=START_DATE, end=END_DATE):
    delta = (end - start).days
    return [start + timedelta(days=int(d)) for d in np.random.randint(0, delta + 1, n)]


def make_sales():
    dates        = rand_dates(N_SALES)
    customers    = np.random.choice(CUSTOMERS, N_SALES)
    products     = np.random.choice(PRODUCTS, N_SALES)
    routes       = np.random.choice(ROUTES, N_SALES)
    reps         = np.random.choice(SALESREPS, N_SALES)
    regions      = np.random.choice(REGIONS, N_SALES)
    channels     = np.random.choice(CHANNELS, N_SALES)
    categories   = np.random.choice(CATEGORIES, N_SALES)
    brands       = np.random.choice(BRANDS, N_SALES)
    quantity     = np.random.randint(1, 200, N_SALES)
    unit_price   = np.round(np.random.uniform(5, 500, N_SALES), 2)
    discount_rate= np.round(np.random.uniform(0, 0.25, N_SALES), 4)
    return_rate  = np.round(np.random.uniform(0, 0.05, N_SALES), 4)
    cogs_rate    = np.round(np.random.uniform(0.45, 0.75, N_SALES), 4)

    gross_sales  = np.round(quantity * unit_price, 2)
    discount_amt = np.round(gross_sales * discount_rate, 2)
    net_sales    = np.round(gross_sales - discount_amt, 2)
    returns_amt  = np.round(net_sales * return_rate, 2)
    net_sales_after_returns = np.round(net_sales - returns_amt, 2)
    cogs         = np.round(net_sales_after_returns * cogs_rate, 2)
    gross_profit = np.round(net_sales_after_returns - cogs, 2)

    df = pd.DataFrame({
        "Transaction_ID"  : [f"TXN_{i:05d}" for i in range(1, N_SALES + 1)],
        "Date"            : dates,
        "Month"           : [d.month for d in dates],
        "Quarter"         : [f"Q{((d.month-1)//3)+1}" for d in dates],
        "Year"            : [d.year for d in dates],
        "Customer_ID"     : customers,
        "Product_ID"      : products,
        "Route_ID"        : routes,
        "SalesRep_ID"     : reps,
        "Region"          : regions,
        "Channel"         : channels,
        "Category"        : categories,
        "Brand"           : brands,
        "Quantity"        : quantity,
        "Unit_Price"      : unit_price,
        "Discount_Rate"   : discount_rate,
        "Gross_Sales"     : gross_sales,
        "Discount_Amount" : discount_amt,
        "Return_Rate"     : return_rate,
        "Returns_Amount"  : returns_amt,
        "Net_Sales"       : net_sales_after_returns,
        "COGS_Rate"       : cogs_rate,
        "COGS"            : cogs,
        "Gross_Profit"    : gross_profit,
    })
    return df.sort_values("Date").reset_index(drop=True)


def make_customers():
    segments = ["Platinum", "Gold", "Silver", "Bronze"]
    rows = []
    for cid in CUSTOMERS:
        region  = random.choice(REGIONS)
        channel = random.choice(CHANNELS)
        since   = START_DATE - timedelta(days=random.randint(30, 1825))
        rows.append({
            "Customer_ID"    : cid,
            "Customer_Name"  : fake.company(),
            "Region"         : region,
            "Channel"        : channel,
            "Route_ID"       : random.choice(ROUTES),
            "Segment"        : random.choice(segments),
            "Credit_Limit"   : round(random.uniform(5000, 100000), 2),
            "Payment_Terms"  : random.choice([30, 45, 60, 90]),
            "Customer_Since" : since.strftime("%Y-%m-%d"),
            "Last_Sale_Date" : (START_DATE + timedelta(days=random.randint(0, 364))).strftime("%Y-%m-%d"),
        })
    return pd.DataFrame(rows)


def make_products():
    rows = []
    for pid in PRODUCTS:
        cat   = random.choice(CATEGORIES)
        brand = random.choice(BRANDS)
        cost  = round(random.uniform(2, 200), 2)
        price = round(cost * random.uniform(1.3, 2.5), 2)
        rows.append({
            "Product_ID"      : pid,
            "Product_Name"    : f"{brand} {cat} {pid}",
            "Category"        : cat,
            "Brand"           : brand,
            "Unit_Cost"       : cost,
            "List_Price"      : price,
            "Weight_kg"       : round(random.uniform(0.1, 10), 2),
            "Shelf_Life_Days" : random.choice([30, 60, 90, 180, 365]),
            "Is_Active"       : random.choice([1, 1, 1, 0]),
            "Launch_Date"     : (START_DATE - timedelta(days=random.randint(0, 730))).strftime("%Y-%m-%d"),
        })
    return pd.DataFrame(rows)


def make_inventory():
    rows = []
    idx = 1
    for pid in PRODUCTS:
        for wh in random.sample(WAREHOUSES, k=random.randint(1, 4)):
            stock    = random.randint(0, 2000)
            cost     = round(random.uniform(2, 200), 2)
            expiry   = START_DATE + timedelta(days=random.randint(30, 365))
            days_exp = (expiry - END_DATE).days
            rows.append({
                "Inventory_ID"    : f"INV_{idx:04d}",
                "Product_ID"      : pid,
                "Warehouse"       : wh,
                "Stock_Qty"       : stock,
                "Reorder_Point"   : random.randint(50, 300),
                "Unit_Cost"       : cost,
                "Inventory_Value" : round(stock * cost, 2),
                "Expiry_Date"     : expiry.strftime("%Y-%m-%d"),
                "Days_To_Expiry"  : days_exp,
                "Last_Count_Date" : (END_DATE - timedelta(days=random.randint(0, 30))).strftime("%Y-%m-%d"),
                "ABC_Class"       : random.choice(["A", "A", "B", "B", "C"]),
                "XYZ_Class"       : random.choice(["X", "X", "Y", "Z"]),
            })
            idx += 1
    return pd.DataFrame(rows)


def make_visits():
    dates    = rand_dates(N_VISITS)
    rows = []
    for i, d in enumerate(dates):
        visited   = random.random() > 0.1
        order_taken = visited and (random.random() > 0.35)
        rows.append({
            "Visit_ID"          : f"VIS_{i+1:05d}",
            "Date"              : d,
            "Month"             : d.month,
            "Quarter"           : f"Q{((d.month-1)//3)+1}",
            "Customer_ID"       : random.choice(CUSTOMERS),
            "SalesRep_ID"       : random.choice(SALESREPS),
            "Route_ID"          : random.choice(ROUTES),
            "Visited"           : int(visited),
            "Order_Taken"       : int(order_taken),
            "Visit_Duration_min": random.randint(5, 60) if visited else 0,
            "SKUs_Checked"      : random.randint(1, 20) if visited else 0,
            "Compliance_Score"  : round(random.uniform(0.5, 1.0), 2) if visited else 0,
            "Visit_Cost"        : round(random.uniform(5, 50), 2),
        })
    return pd.DataFrame(rows).sort_values("Date").reset_index(drop=True)


def make_ar_aging():
    rows = []
    for i in range(N_AR):
        inv_date  = rand_dates(1)[0]
        terms     = random.choice([30, 45, 60, 90])
        due_date  = inv_date + timedelta(days=terms)
        amount    = round(random.uniform(500, 50000), 2)
        paid_pct  = random.uniform(0, 1)
        paid      = round(amount * paid_pct, 2)
        outstanding = round(amount - paid, 2)
        overdue   = (END_DATE - due_date).days
        if overdue <= 0:
            bucket = "Current"
        elif overdue <= 30:
            bucket = "1-30"
        elif overdue <= 60:
            bucket = "31-60"
        elif overdue <= 90:
            bucket = "61-90"
        else:
            bucket = "90+"
        rows.append({
            "Invoice_ID"    : f"INV_{i+1:04d}",
            "Customer_ID"   : random.choice(CUSTOMERS),
            "Invoice_Date"  : inv_date.strftime("%Y-%m-%d"),
            "Due_Date"      : due_date.strftime("%Y-%m-%d"),
            "Payment_Terms" : terms,
            "Invoice_Amount": amount,
            "Paid_Amount"   : paid,
            "Outstanding"   : outstanding,
            "Days_Overdue"  : max(overdue, 0),
            "Aging_Bucket"  : bucket,
        })
    return pd.DataFrame(rows)


def make_external():
    rows = []
    for m in range(1, 13):
        market_size   = round(random.uniform(5_000_000, 8_000_000), 0)
        company_sales = round(market_size * random.uniform(0.08, 0.18), 0)
        rows.append({
            "Period"               : f"2024-{m:02d}",
            "Month"                : m,
            "Market_Size"          : market_size,
            "Company_Sales"        : company_sales,
            "Competitor_Sales"     : round(market_size * random.uniform(0.10, 0.25), 0),
            "NPS_Score"            : round(random.uniform(20, 70), 1),
            "Brand_Health_Index"   : round(random.uniform(50, 90), 1),
            "CSAT_Score"           : round(random.uniform(3.0, 5.0), 2),
            "Competitor_Price_Index": round(random.uniform(0.85, 1.15), 3),
            "Category_Growth_Pct"  : round(random.uniform(-0.05, 0.15), 4),
            "Inflation_Rate_Pct"   : round(random.uniform(0.02, 0.08), 4),
            "Planogram_Compliance_Pct": round(random.uniform(0.60, 0.95), 3),
        })
    return pd.DataFrame(rows)


def make_targets():
    rows = []
    for m in range(1, 13):
        for rep in SALESREPS:
            rows.append({
                "Month"          : m,
                "SalesRep_ID"    : rep,
                "Sales_Target"   : round(random.uniform(50_000, 200_000), 0),
                "Visit_Target"   : random.randint(80, 200),
                "NewCust_Target" : random.randint(2, 10),
                "GP_Target_Pct"  : round(random.uniform(0.25, 0.45), 4),
                "OTD_Target_Pct" : round(random.uniform(0.90, 0.98), 4),
            })
    return pd.DataFrame(rows)


def make_assumptions():
    return pd.DataFrame([
        {"Parameter": "WACC_Pct",              "Value": 0.12,  "Description": "Weighted Average Cost of Capital"},
        {"Parameter": "Tax_Rate_Pct",          "Value": 0.25,  "Description": "Corporate Tax Rate"},
        {"Parameter": "Risk_Free_Rate_Pct",    "Value": 0.05,  "Description": "Risk-Free Rate"},
        {"Parameter": "Inflation_Rate_Pct",    "Value": 0.04,  "Description": "Annual Inflation Rate"},
        {"Parameter": "Target_GP_Margin_Pct",  "Value": 0.35,  "Description": "Target Gross Profit Margin"},
        {"Parameter": "Target_NP_Margin_Pct",  "Value": 0.12,  "Description": "Target Net Profit Margin"},
        {"Parameter": "Target_Growth_Pct",     "Value": 0.15,  "Description": "Target Annual Growth Rate"},
        {"Parameter": "OpEx_Rate_Pct",         "Value": 0.18,  "Description": "Operating Expense as % of Net Sales"},
        {"Parameter": "DA_Rate_Pct",           "Value": 0.03,  "Description": "Depreciation & Amortization Rate"},
        {"Parameter": "Interest_Rate_Pct",     "Value": 0.07,  "Description": "Interest Rate on Debt"},
        {"Parameter": "Debt_Ratio",            "Value": 0.40,  "Description": "Debt to Total Assets Ratio"},
        {"Parameter": "Total_Assets",          "Value": 5_000_000, "Description": "Total Assets ($)"},
        {"Parameter": "Total_Equity",          "Value": 3_000_000, "Description": "Total Equity ($)"},
        {"Parameter": "Invested_Capital",      "Value": 4_000_000, "Description": "Invested Capital ($)"},
        {"Parameter": "Num_Stores_Total",      "Value": 500,   "Description": "Total Stores in Market"},
        {"Parameter": "Num_SKUs_Total",        "Value": 50,    "Description": "Total Active SKUs"},
        {"Parameter": "Avg_Inventory_Days",    "Value": 30,    "Description": "Average Inventory Holding Days"},
        {"Parameter": "Carrying_Cost_Rate",    "Value": 0.20,  "Description": "Annual Inventory Carrying Cost Rate"},
        {"Parameter": "Current_Liabilities",   "Value": 1_500_000, "Description": "Current Liabilities ($)"},
        {"Parameter": "Variable_Cost_Rate",    "Value": 0.40,  "Description": "Variable Cost Rate"},
        {"Parameter": "Carbon_Per_Unit",       "Value": 0.5,   "Description": "Carbon Emission per Unit (kg)"},
    ])


# ─────────────────────────────────────────────
# 2. EXCEL HELPERS
# ─────────────────────────────────────────────

HEADER_FILL   = PatternFill("solid", fgColor="1F3864")
SUBHEAD_FILL  = PatternFill("solid", fgColor="2E75B6")
ALT_FILL      = PatternFill("solid", fgColor="EBF3FB")
KPI_CAT_FILL  = PatternFill("solid", fgColor="375623")
KPI_VAL_FILL  = PatternFill("solid", fgColor="E2EFDA")
DASH_FILL     = PatternFill("solid", fgColor="1F3864")

WHITE_FONT    = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
HEADER_FONT   = Font(name="Calibri", bold=True, color="FFFFFF", size=10)
BODY_FONT     = Font(name="Calibri", size=10)
BOLD_FONT     = Font(name="Calibri", bold=True, size=10)
KPI_CODE_FONT = Font(name="Calibri", bold=True, color="1F3864", size=10)
DASH_FONT     = Font(name="Calibri", bold=True, color="FFFFFF", size=12)

THIN = Side(style="thin", color="BDD7EE")
THICK = Side(style="medium", color="2E75B6")
THIN_BORDER  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
THICK_BORDER = Border(left=THICK, right=THICK, top=THICK, bottom=THICK)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)
RIGHT  = Alignment(horizontal="right",  vertical="center")


def thin_border():
    return THIN_BORDER


def create_data_sheet(wb, name):
    """Create a new sheet and return it."""
    return wb.create_sheet(name)


def populate_data_sheet(ws, df):
    """Write DataFrame to worksheet with proper styling."""
    # Write headers
    for c_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=c_idx, value=col_name)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = THIN_BORDER
    
    # Write data
    for r_idx, row in enumerate(df.itertuples(index=False), 2):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font = BODY_FONT
            cell.alignment = LEFT
            cell.border = THIN_BORDER
            if r_idx % 2 == 0:
                cell.fill = ALT_FILL
    
    # Auto-fit columns
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = min(max_len + 2, 35)


def style_header_row(ws, row, cols, fill=HEADER_FILL, font=HEADER_FONT):
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill  = fill
        cell.font  = font
        cell.alignment = CENTER
        cell.border = THIN_BORDER


def write_df_to_sheet(ws, df, start_row=1):
    """Write DataFrame with styled header."""
    for c_idx, col_name in enumerate(df.columns, 1):
        cell = ws.cell(row=start_row, column=c_idx, value=col_name)
        cell.fill      = HEADER_FILL
        cell.font      = HEADER_FONT
        cell.alignment = CENTER
        cell.border    = THIN_BORDER

    for r_idx, row in enumerate(df.itertuples(index=False), start_row + 1):
        fill = ALT_FILL if r_idx % 2 == 0 else PatternFill()
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.font      = BODY_FONT
            cell.alignment = LEFT
            cell.border    = THIN_BORDER
            cell.fill      = fill

    # Auto-width
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 40)


def kpi_row(ws, row, code, name_en, name_fa, formula, fmt, category_fill=None):
    """Write one KPI row: Code | Name EN | Name FA | Value (formula) | Format."""
    cells_data = [
        (1, code,    KPI_CODE_FONT, CENTER),
        (2, name_en, BOLD_FONT,     LEFT),
        (3, name_fa, BODY_FONT,     LEFT),
        (4, formula, BODY_FONT,     RIGHT),
        (5, fmt,     BODY_FONT,     CENTER),
    ]
    fill = category_fill or KPI_VAL_FILL
    for col, val, font, align in cells_data:
        cell = ws.cell(row=row, column=col, value=val)
        cell.font      = font
        cell.alignment = align
        cell.border    = THIN_BORDER
        if col == 4:
            cell.fill = KPI_VAL_FILL
        else:
            cell.fill = fill


def kpi_section_header(ws, row, code_prefix, title_en, title_fa):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=5)
    cell = ws.cell(row=row, column=1,
                   value=f"  {code_prefix}  |  {title_en}  |  {title_fa}")
    cell.fill      = KPI_CAT_FILL
    cell.font      = WHITE_FONT
    cell.alignment = LEFT
    cell.border    = THICK_BORDER


def set_kpi_sheet_columns(ws):
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 32
    ws.column_dimensions["D"].width = 22
    ws.column_dimensions["E"].width = 14
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "A3"


def kpi_sheet_header(ws, title_en, title_fa):
    ws.merge_cells("A1:E1")
    cell = ws["A1"]
    cell.value     = f"{title_en}  |  {title_fa}"
    cell.fill      = DASH_FILL
    cell.font      = Font(name="Calibri", bold=True, color="FFFFFF", size=14)
    cell.alignment = CENTER

    headers = ["Code", "KPI Name (EN)", "KPI Name (FA)", "Value", "Format"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(row=2, column=i, value=h)
        c.fill      = SUBHEAD_FILL
        c.font      = HEADER_FONT
        c.alignment = CENTER
        c.border    = THIN_BORDER


# ─────────────────────────────────────────────
# 3. KPI SHEET BUILDERS
# ─────────────────────────────────────────────

S  = "Data_Sales"
V  = "Data_Visits"
IN = "Data_Inventory"
AR = "Data_AR_Aging"
EX = "Data_External"
TG = "Data_Targets"
AS = "Assumptions"
CU = "Data_Customers"
PR = "Data_Products"


def build_kpi_sales(ws):
    kpi_sheet_header(ws, "KPI — Sales & Marketing", "فروش و بازاریابی")
    set_kpi_sheet_columns(ws)
    r = 3

    kpi_section_header(ws, r, "SM", "Core Sales Metrics", "شاخص‌های اصلی فروش")
    r += 1

    kpis = [
        ("SM01", "Net Sales ($)",
         "فروش خالص",
         f"=SUM({S}!U2:U5001)",
         "$#,##0"),

        ("SM02", "Gross Sales ($)",
         "فروش ناخالص",
         f"=SUM({S}!Q2:Q5001)",
         "$#,##0"),

        ("SM03", "Total Volume (Units)",
         "حجم فروش (واحد)",
         f"=SUM({S}!N2:N5001)",
         "#,##0"),

        ("SM04", "Average Selling Price ($)",
         "میانگین قیمت فروش",
         f"=IFERROR(SM01_val/SM03_val,0)",
         "$#,##0.00"),

        ("SM05", "Hit Rate (%)",
         "نرخ موفقیت سفارش",
         f"=IFERROR(SUMIF({V}!I2:I4001,1,{V}!I2:I4001)/COUNTA({V}!A2:A4001),0)",
         "0.00%"),

        ("SM06", "Average Drop Size ($)",
         "میانگین ارزش هر سفارش",
         f"=IFERROR(SUM({S}!U2:U5001)/SUMIF({V}!I2:I4001,1,{V}!I2:I4001),0)",
         "$#,##0.00"),

        ("SM07", "Active Customers",
         "مشتریان فعال",
         f"=COUNTA(UNIQUE({S}!F2:F5001))",
         "#,##0"),

        ("SM08", "Revenue per Customer ($)",
         "درآمد به ازای هر مشتری",
         f"=IFERROR(SUM({S}!U2:U5001)/COUNTA(UNIQUE({S}!F2:F5001)),0)",
         "$#,##0.00"),

        ("SM09", "Revenue per Rep ($)",
         "درآمد به ازای هر نماینده",
         f"=IFERROR(SUM({S}!U2:U5001)/COUNTA(UNIQUE({S}!I2:I5001)),0)",
         "$#,##0.00"),

        ("SM10", "Discount Rate (%)",
         "نرخ تخفیف",
         f"=IFERROR(SUM({S}!R2:R5001)/SUM({S}!Q2:Q5001),0)",
         "0.00%"),

        ("SM11", "Return Rate (%)",
         "نرخ مرجوعی",
         f"=IFERROR(SUM({S}!T2:T5001)/SUM({S}!U2:U5001),0)",
         "0.00%"),

        ("SM12", "Gross Profit ($)",
         "سود ناخالص",
         f"=SUM({S}!X2:X5001)",
         "$#,##0"),

        ("SM13", "Gross Profit Margin (%)",
         "حاشیه سود ناخالص",
         f"=IFERROR(SUM({S}!X2:X5001)/SUM({S}!U2:U5001),0)",
         "0.00%"),

        ("SM14", "COGS ($)",
         "بهای تمام‌شده کالای فروش‌رفته",
         f"=SUM({S}!W2:W5001)",
         "$#,##0"),

        ("SM15", "COGS as % of Net Sales",
         "بهای تمام‌شده به درصد فروش خالص",
         f"=IFERROR(SUM({S}!W2:W5001)/SUM({S}!U2:U5001),0)",
         "0.00%"),

        ("SM16", "Operating Expense ($)",
         "هزینه‌های عملیاتی",
         f"=SUM({S}!U2:U5001)*VLOOKUP(\"OpEx_Rate_Pct\",{AS}!$A:$B,2,0)",
         "$#,##0"),

        ("SM17", "EBITDA ($)",
         "سود قبل از بهره، مالیات، استهلاک",
         f"=SUM({S}!X2:X5001)"
         f"-SUM({S}!U2:U5001)*VLOOKUP(\"OpEx_Rate_Pct\",{AS}!$A:$B,2,0)"
         f"+SUM({S}!U2:U5001)*VLOOKUP(\"DA_Rate_Pct\",{AS}!$A:$B,2,0)",
         "$#,##0"),

        ("SM18", "EBITDA Margin (%)",
         "حاشیه EBITDA",
         f"=IFERROR(D{r+14}/SUM({S}!U2:U5001),0)",
         "0.00%"),

        ("SM19", "Net Profit ($)",
         "سود خالص",
         f"=SUM({S}!X2:X5001)"
         f"-SUM({S}!U2:U5001)*VLOOKUP(\"OpEx_Rate_Pct\",{AS}!$A:$B,2,0)"
         f"-SUM({S}!U2:U5001)*VLOOKUP(\"DA_Rate_Pct\",{AS}!$A:$B,2,0)"
         f"-VLOOKUP(\"Total_Assets\",{AS}!$A:$B,2,0)*VLOOKUP(\"Debt_Ratio\",{AS}!$A:$B,2,0)*VLOOKUP(\"Interest_Rate_Pct\",{AS}!$A:$B,2,0)",
         "$#,##0"),

        ("SM20", "Net Profit Margin (%)",
         "حاشیه سود خالص",
         f"=IFERROR(D{r+15}/SUM({S}!U2:U5001),0)",
         "0.00%"),

        ("SM21", "Sales Growth YoY (%)",
         "رشد فروش سال به سال",
         f"=IFERROR((SUMIF({S}!E2:E5001,2024,{S}!U2:U5001)"
         f"-SUMIF({S}!E2:E5001,2023,{S}!U2:U5001))"
         f"/SUMIF({S}!E2:E5001,2023,{S}!U2:U5001),0)",
         "0.00%"),

        ("SM22", "Sales vs Target (%)",
         "فروش در مقابل هدف",
         f"=IFERROR(SUM({S}!U2:U5001)/SUM({TG}!C2:C1000),0)",
         "0.00%"),

        ("SM23", "Sell-in Volume (Units)",
         "حجم فروش به توزیع‌کننده",
         f"=SUMIF({S}!K2:K5001,\"Modern Trade\",{S}!N2:N5001)"
         f"+SUMIF({S}!K2:K5001,\"Traditional Trade\",{S}!N2:N5001)",
         "#,##0"),

        ("SM24", "Sell-out Volume (Units)",
         "حجم فروش به مصرف‌کننده نهایی",
         f"=SUMIF({S}!K2:K5001,\"E-Commerce\",{S}!N2:N5001)"
         f"+SUMIF({S}!K2:K5001,\"HoReCa\",{S}!N2:N5001)",
         "#,##0"),

        ("SM25", "Sell-in / Sell-out Ratio",
         "نسبت فروش به توزیع‌کننده به مصرف‌کننده",
         f"=IFERROR(D{r+19}/D{r+20},0)",
         "0.00"),

        ("SM26", "Revenue by Modern Trade ($)",
         "درآمد کانال مدرن",
         f"=SUMIF({S}!K2:K5001,\"Modern Trade\",{S}!U2:U5001)",
         "$#,##0"),

        ("SM27", "Revenue by Traditional Trade ($)",
         "درآمد کانال سنتی",
         f"=SUMIF({S}!K2:K5001,\"Traditional Trade\",{S}!U2:U5001)",
         "$#,##0"),

        ("SM28", "Revenue by E-Commerce ($)",
         "درآمد کانال آنلاین",
         f"=SUMIF({S}!K2:K5001,\"E-Commerce\",{S}!U2:U5001)",
         "$#,##0"),

        ("SM29", "Revenue by HoReCa ($)",
         "درآمد کانال هتل/رستوران",
         f"=SUMIF({S}!K2:K5001,\"HoReCa\",{S}!U2:U5001)",
         "$#,##0"),

        ("SM30", "Top Region Revenue ($)",
         "درآمد برترین منطقه",
         f"=IFERROR(MAX(SUMIF({S}!J2:J5001,UNIQUE({S}!J2:J5001),{S}!U2:U5001)),0)",
         "$#,##0"),

        ("SM31", "Avg Monthly Revenue ($)",
         "میانگین درآمد ماهانه",
         f"=IFERROR(SUM({S}!U2:U5001)/COUNTA(UNIQUE({S}!C2:C5001)),0)",
         "$#,##0"),
    ]

    for code, name_en, name_fa, formula, fmt in kpis:
        kpi_row(ws, r, code, name_en, name_fa, formula, fmt)
        r += 1

    return r


def build_kpi_customer(ws):
    kpi_sheet_header(ws, "KPI — Customer Analytics", "تحلیل مشتریان")
    set_kpi_sheet_columns(ws)
    r = 3

    kpi_section_header(ws, r, "CA", "Customer Analytics", "تحلیل مشتریان")
    r += 1

    kpis = [
        ("CA01", "Total Active Customers",
         "تعداد مشتریان فعال",
         f"=COUNTA(UNIQUE({S}!F2:F5001))",
         "#,##0"),

        ("CA02", "New Customers (Period)",
         "مشتریان جدید دوره",
         f"=SUMPRODUCT((COUNTIF({S}!F2:F5001,{CU}!A2:A201)>0)"
         f"*(DATEVALUE({CU}!I2:I201)>=DATE(2024,1,1)))",
         "#,##0"),

        ("CA03", "Churned Customers",
         "مشتریان از دست‌رفته",
         f"=SUMPRODUCT(({CU}!A2:A201<>\"\")"
         f"*(COUNTIF({S}!F2:F5001,{CU}!A2:A201)=0))",
         "#,##0"),

        ("CA04", "Customer Retention Rate (%)",
         "نرخ نگهداشت مشتری",
         f"=IFERROR(1-D{r+2}/D{r},0)",
         "0.00%"),

        ("CA05", "Customer Churn Rate (%)",
         "نرخ ریزش مشتری",
         f"=IFERROR(D{r+2}/D{r},0)",
         "0.00%"),

        ("CA06", "Average Revenue per Customer ($)",
         "میانگین درآمد هر مشتری",
         f"=IFERROR(SUM({S}!U2:U5001)/D{r},0)",
         "$#,##0.00"),

        ("CA07", "Customer Lifetime Value ($)",
         "ارزش طول عمر مشتری",
         f"=IFERROR(D{r+5}*(1/D{r+4}),0)",
         "$#,##0.00"),

        ("CA08", "Customer Acquisition Cost ($)",
         "هزینه جذب مشتری",
         f"=IFERROR(SUM({V}!M2:M4001)/MAX(D{r+1},1),0)",
         "$#,##0.00"),

        ("CA09", "LTV / CAC Ratio",
         "نسبت ارزش عمر به هزینه جذب",
         f"=IFERROR(D{r+6}/D{r+7},0)",
         "0.00"),

        ("CA10", "Purchase Frequency",
         "تناوب خرید",
         f"=IFERROR(COUNTA({S}!A2:A5001)/COUNTA(UNIQUE({S}!F2:F5001)),0)",
         "0.00"),

        ("CA11", "Average Order Value ($)",
         "میانگین ارزش سفارش",
         f"=IFERROR(SUM({S}!U2:U5001)/COUNTA({S}!A2:A5001),0)",
         "$#,##0.00"),

        ("CA12", "NPS Score",
         "شاخص خالص ترویج‌کنندگان",
         f"=IFERROR(AVERAGE({EX}!F2:F13),0)",
         "0.0"),

        ("CA13", "CSAT Score",
         "رضایت مشتری",
         f"=IFERROR(AVERAGE({EX}!H2:H13),0)",
         "0.00"),

        ("CA14", "Platinum Segment Revenue ($)",
         "درآمد بخش پلاتینیوم",
         f"=IFERROR(SUMPRODUCT(({S}!U2:U5001)*"
         f"(VLOOKUP({S}!F2:F5001,{CU}!$A:$F,6,0)=\"Platinum\")),0)",
         "$#,##0"),

        ("CA15", "Gold Segment Revenue ($)",
         "درآمد بخش طلایی",
         f"=IFERROR(SUMPRODUCT(({S}!U2:U5001)*"
         f"(VLOOKUP({S}!F2:F5001,{CU}!$A:$F,6,0)=\"Gold\")),0)",
         "$#,##0"),

        ("CA16", "Top 10% Customer Revenue Share (%)",
         "سهم درآمد ۱۰٪ مشتریان برتر",
         f"=IFERROR(SUMPRODUCT(LARGE(SUMIF({S}!F2:F5001,{CU}!A2:A201,{S}!U2:U5001),"
         f"ROW(INDIRECT(\"1:\"&ROUND(COUNTA({CU}!A2:A201)*0.1,0)))))/SUM({S}!U2:U5001),0)",
         "0.00%"),

        ("CA17", "Repeat Purchase Rate (%)",
         "نرخ خرید مجدد",
         f"=IFERROR(SUMPRODUCT((COUNTIF({S}!F2:F5001,UNIQUE({S}!F2:F5001))>1)*1)"
         f"/COUNTA(UNIQUE({S}!F2:F5001)),0)",
         "0.00%"),

        ("CA18", "Avg Payment Terms (Days)",
         "میانگین شرایط پرداخت",
         f"=IFERROR(AVERAGE({CU}!H2:H201),0)",
         "0.0"),

        ("CA19", "Credit Utilization Rate (%)",
         "نرخ استفاده از اعتبار",
         f"=IFERROR(SUM({AR}!H2:H401)/SUM({CU}!G2:G201),0)",
         "0.00%"),

        ("CA20", "Customer Coverage Rate (%)",
         "نرخ پوشش مشتریان",
         f"=IFERROR(COUNTA(UNIQUE({V}!E2:E4001))/COUNTA({CU}!A2:A201),0)",
         "0.00%"),
    ]

    for code, name_en, name_fa, formula, fmt in kpis:
        kpi_row(ws, r, code, name_en, name_fa, formula, fmt)
        r += 1

    return r


def build_kpi_inventory(ws):
    kpi_sheet_header(ws, "KPI — Inventory Management", "مدیریت موجودی")
    set_kpi_sheet_columns(ws)
    r = 3

    kpi_section_header(ws, r, "IN", "Inventory KPIs", "شاخص‌های موجودی")
    r += 1

    kpis = [
        ("IN01", "Total Inventory Value ($)",
         "ارزش کل موجودی",
         f"=SUM({IN}!G2:G1001)",
         "$#,##0"),

        ("IN02", "Inventory Turnover Ratio",
         "نسبت گردش موجودی",
         f"=IFERROR(SUM({S}!W2:W5001)/AVERAGE({IN}!G2:G1001),0)",
         "0.00"),

        ("IN03", "Days Inventory Outstanding (DIO)",
         "روزهای موجودی",
         f"=IFERROR(365/D{r+1},0)",
         "0.0"),

        ("IN04", "Stockout Rate (%)",
         "نرخ کمبود موجودی",
         f"=IFERROR(COUNTIF({IN}!D2:D1001,0)/COUNTA({IN}!A2:A1001),0)",
         "0.00%"),

        ("IN05", "Overstock Rate (%)",
         "نرخ موجودی مازاد",
         f"=IFERROR(SUMPRODUCT(({IN}!D2:D1001>{IN}!E2:E1001*3)*1)/COUNTA({IN}!A2:A1001),0)",
         "0.00%"),

        ("IN06", "Near-Expiry Stock (%)",
         "موجودی نزدیک به انقضا",
         f"=IFERROR(SUMPRODUCT(({IN}!I2:I1001>=0)*({IN}!I2:I1001<=30)*{IN}!G2:G1001)"
         f"/SUM({IN}!G2:G1001),0)",
         "0.00%"),

        ("IN07", "Expired Stock Value ($)",
         "ارزش موجودی منقضی‌شده",
         f"=SUMPRODUCT(({IN}!I2:I1001<0)*{IN}!G2:G1001)",
         "$#,##0"),

        ("IN08", "Inventory Carrying Cost ($)",
         "هزینه نگهداری موجودی",
         f"=SUM({IN}!G2:G1001)*VLOOKUP(\"Carrying_Cost_Rate\",{AS}!$A:$B,2,0)",
         "$#,##0"),

        ("IN09", "Fill Rate (%)",
         "نرخ تأمین سفارش",
         f"=IFERROR(SUMIF({IN}!D2:D1001,\">\"&{IN}!E2:E1001,{IN}!D2:D1001)"
         f"/SUM({IN}!D2:D1001),0)",
         "0.00%"),

        ("IN10", "ABC-A Stock Value ($)",
         "ارزش موجودی کلاس A",
         f"=SUMIF({IN}!K2:K1001,\"A\",{IN}!G2:G1001)",
         "$#,##0"),

        ("IN11", "ABC-B Stock Value ($)",
         "ارزش موجودی کلاس B",
         f"=SUMIF({IN}!K2:K1001,\"B\",{IN}!G2:G1001)",
         "$#,##0"),

        ("IN12", "ABC-C Stock Value ($)",
         "ارزش موجودی کلاس C",
         f"=SUMIF({IN}!K2:K1001,\"C\",{IN}!G2:G1001)",
         "$#,##0"),
    ]

    for code, name_en, name_fa, formula, fmt in kpis:
        kpi_row(ws, r, code, name_en, name_fa, formula, fmt)
        r += 1

    return r


def build_kpi_distribution(ws):
    kpi_sheet_header(ws, "KPI — Distribution & Route", "توزیع و مسیر")
    set_kpi_sheet_columns(ws)
    r = 3

    kpi_section_header(ws, r, "DR", "Distribution KPIs", "شاخص‌های توزیع")
    r += 1

    kpis = [
        ("DR01", "Numeric Distribution (%)",
         "توزیع عددی",
         f"=IFERROR(COUNTA(UNIQUE({S}!F2:F5001))"
         f"/VLOOKUP(\"Num_Stores_Total\",{AS}!$A:$B,2,0),0)",
         "0.00%"),

        ("DR02", "Weighted Distribution (%)",
         "توزیع وزنی",
         f"=IFERROR(SUM({S}!U2:U5001)"
         f"/(SUM({S}!U2:U5001)/IFERROR(COUNTA(UNIQUE({S}!F2:F5001)),1)"
         f"*VLOOKUP(\"Num_Stores_Total\",{AS}!$A:$B,2,0)),0)",
         "0.00%"),

        ("DR03", "Route Productivity ($)",
         "بهره‌وری مسیر",
         f"=IFERROR(SUM({S}!U2:U5001)/COUNTA(UNIQUE({S}!H2:H5001)),0)",
         "$#,##0.00"),

        ("DR04", "Calls per Day",
         "تعداد مراجعه روزانه",
         f"=IFERROR(COUNTA({V}!A2:A4001)/COUNTA(UNIQUE({V}!B2:B4001)),0)",
         "0.0"),

        ("DR05", "Effective Calls (%)",
         "نرخ مراجعه مؤثر",
         f"=IFERROR(SUMIF({V}!H2:H4001,1,{V}!H2:H4001)/COUNTA({V}!A2:A4001),0)",
         "0.00%"),

        ("DR06", "Strike Rate (%)",
         "نرخ تبدیل بازدید به سفارش",
         f"=IFERROR(SUMIF({V}!I2:I4001,1,{V}!I2:I4001)"
         f"/SUMIF({V}!H2:H4001,1,{V}!H2:H4001),0)",
         "0.00%"),

        ("DR07", "Visit Compliance (%)",
         "انطباق بازدید با برنامه",
         f"=IFERROR(AVERAGE({V}!L2:L4001),0)",
         "0.00%"),

        ("DR08", "SKU Distribution Coverage (%)",
         "پوشش توزیع SKU",
         f"=IFERROR(COUNTA(UNIQUE({S}!G2:G5001))"
         f"/VLOOKUP(\"Num_SKUs_Total\",{AS}!$A:$B,2,0),0)",
         "0.00%"),

        ("DR09", "Out-of-Stock Incidents",
         "تعداد موارد کمبود موجودی",
         f"=COUNTIF({IN}!D2:D1001,0)",
         "#,##0"),

        ("DR10", "On-Time Delivery Rate (%)",
         "نرخ تحویل به‌موقع",
         f"=IFERROR(AVERAGE({TG}!G2:G1000),0)",
         "0.00%"),

        ("DR11", "Perfect Order Rate (%)",
         "نرخ سفارش کامل",
         f"=IFERROR(D{r+9}*(1-D{r+2})*(1-D{r+4}),0)",
         "0.00%"),

        ("DR12", "Planogram Compliance (%)",
         "انطباق با پلانوگرام",
         f"=IFERROR(AVERAGE({EX}!L2:L13),0)",
         "0.00%"),

        ("DR13", "Market Share (%)",
         "سهم بازار",
         f"=IFERROR(AVERAGE({EX}!D2:D13)/AVERAGE({EX}!C2:C13),0)",
         "0.00%"),

        ("DR14", "Share of Shelf (%)",
         "سهم قفسه",
         f"=IFERROR(COUNTA(UNIQUE({S}!G2:G5001))"
         f"/(VLOOKUP(\"Num_SKUs_Total\",{AS}!$A:$B,2,0)*1.5),0)",
         "0.00%"),

        ("DR15", "Cost per Visit ($)",
         "هزینه هر بازدید",
         f"=IFERROR(SUM({V}!M2:M4001)/COUNTA({V}!A2:A4001),0)",
         "$#,##0.00"),

        ("DR16", "Revenue per Visit ($)",
         "درآمد هر بازدید",
         f"=IFERROR(SUM({S}!U2:U5001)/COUNTA({V}!A2:A4001),0)",
         "$#,##0.00"),
    ]

    for code, name_en, name_fa, formula, fmt in kpis:
        kpi_row(ws, r, code, name_en, name_fa, formula, fmt)
        r += 1

    return r


def build_kpi_finance(ws):
    kpi_sheet_header(ws, "KPI — Financial Performance", "عملکرد مالی")
    set_kpi_sheet_columns(ws)
    r = 3

    kpi_section_header(ws, r, "FI", "Financial KPIs", "شاخص‌های مالی")
    r += 1

    kpis = [
        ("FI01", "Return on Sales / Net Margin (%)",
         "بازده فروش",
         f"=IFERROR((SUM({S}!X2:X5001)"
         f"-SUM({S}!U2:U5001)*VLOOKUP(\"OpEx_Rate_Pct\",{AS}!$A:$B,2,0)"
         f"-SUM({S}!U2:U5001)*VLOOKUP(\"DA_Rate_Pct\",{AS}!$A:$B,2,0))"
         f"/SUM({S}!U2:U5001),0)",
         "0.00%"),

        ("FI02", "Return on Assets (%)",
         "بازده دارایی‌ها",
         f"=IFERROR((SUM({S}!X2:X5001)"
         f"-SUM({S}!U2:U5001)*VLOOKUP(\"OpEx_Rate_Pct\",{AS}!$A:$B,2,0))"
         f"/VLOOKUP(\"Total_Assets\",{AS}!$A:$B,2,0),0)",
         "0.00%"),

        ("FI03", "Return on Equity (%)",
         "بازده حقوق صاحبان سهام",
         f"=IFERROR((SUM({S}!X2:X5001)"
         f"-SUM({S}!U2:U5001)*VLOOKUP(\"OpEx_Rate_Pct\",{AS}!$A:$B,2,0))"
         f"/VLOOKUP(\"Total_Equity\",{AS}!$A:$B,2,0),0)",
         "0.00%"),

        ("FI04", "ROCE (%)",
         "بازده سرمایه به‌کاررفته",
         f"=IFERROR((SUM({S}!X2:X5001)"
         f"-SUM({S}!U2:U5001)*VLOOKUP(\"OpEx_Rate_Pct\",{AS}!$A:$B,2,0))"
         f"/VLOOKUP(\"Invested_Capital\",{AS}!$A:$B,2,0),0)",
         "0.00%"),

        ("FI05", "EVA ($)",
         "ارزش افزوده اقتصادی",
         f"=IFERROR((SUM({S}!X2:X5001)"
         f"-SUM({S}!U2:U5001)*VLOOKUP(\"OpEx_Rate_Pct\",{AS}!$A:$B,2,0))"
         f"*(1-VLOOKUP(\"Tax_Rate_Pct\",{AS}!$A:$B,2,0))"
         f"-VLOOKUP(\"Invested_Capital\",{AS}!$A:$B,2,0)"
         f"*VLOOKUP(\"WACC_Pct\",{AS}!$A:$B,2,0),0)",
         "$#,##0"),

        ("FI06", "Days Sales Outstanding (DSO)",
         "روزهای وصول مطالبات",
         f"=IFERROR(SUM({AR}!H2:H401)"
         f"/(SUM({S}!U2:U5001)/365),0)",
         "0.0"),

        ("FI07", "Days Payable Outstanding (DPO)",
         "روزهای پرداخت بدهی",
         f"=IFERROR(AVERAGE({CU}!H2:H201),0)",
         "0.0"),

        ("FI08", "Cash Conversion Cycle (Days)",
         "چرخه تبدیل وجه نقد",
         f"=D{r+5}+VLOOKUP(\"Avg_Inventory_Days\",{AS}!$A:$B,2,0)-D{r+6}",
         "0.0"),

        ("FI09", "AR Overdue Amount ($)",
         "مطالبات معوق",
         f"=SUMIF({AR}!I2:I401,\">0\",{AR}!H2:H401)",
         "$#,##0"),

        ("FI10", "AR Overdue Rate (%)",
         "نرخ مطالبات معوق",
         f"=IFERROR(D{r+8}/SUM({AR}!H2:H401),0)",
         "0.00%"),

        ("FI11", "Bad Debt Provision ($)",
         "ذخیره مطالبات مشکوک‌الوصول",
         f"=SUMIF({AR}!J2:J401,\"90+\",{AR}!H2:H401)*0.5"
         f"+SUMIF({AR}!J2:J401,\"61-90\",{AR}!H2:H401)*0.25",
         "$#,##0"),

        ("FI12", "Working Capital ($)",
         "سرمایه در گردش",
         f"=SUM({IN}!G2:G1001)+SUM({AR}!H2:H401)",
         "$#,##0"),

        ("FI13", "Working Capital Ratio",
         "نسبت سرمایه در گردش",
         f"=IFERROR((SUM({IN}!G2:G1001)+SUM({AR}!H2:H401))"
         f"/VLOOKUP(\"Current_Liabilities\",{AS}!$A:$B,2,0),0)",
         "0.00"),

        ("FI14", "Trade Spend ROI",
         "بازده هزینه‌های تجاری",
         f"=IFERROR((SUM({S}!U2:U5001)-SUM({S}!R2:R5001))"
         f"/SUM({S}!R2:R5001),0)",
         "0.00"),

        ("FI15", "Trade Spend as % of Revenue",
         "هزینه تجاری به درصد درآمد",
         f"=IFERROR(SUM({S}!R2:R5001)/SUM({S}!U2:U5001),0)",
         "0.00%"),

        ("FI16", "Contribution Margin ($)",
         "حاشیه مشارکت",
         f"=SUM({S}!X2:X5001)"
         f"-SUM({S}!U2:U5001)*VLOOKUP(\"Variable_Cost_Rate\",{AS}!$A:$B,2,0)",
         "$#,##0"),

        ("FI17", "Contribution Margin (%)",
         "درصد حاشیه مشارکت",
         f"=IFERROR(D{r+15}/SUM({S}!U2:U5001),0)",
         "0.00%"),
    ]

    for code, name_en, name_fa, formula, fmt in kpis:
        kpi_row(ws, r, code, name_en, name_fa, formula, fmt)
        r += 1

    return r


def build_kpi_advanced(ws):
    kpi_sheet_header(ws, "KPI — Advanced & Enterprise", "شاخص‌های پیشرفته")
    set_kpi_sheet_columns(ws)
    r = 3

    kpi_section_header(ws, r, "AE", "Advanced KPIs", "شاخص‌های پیشرفته")
    r += 1

    kpis = [
        ("AE01", "Market Share (%)",
         "سهم بازار",
         f"=IFERROR(AVERAGE({EX}!D2:D13)/AVERAGE({EX}!C2:C13),0)",
         "0.00%"),

        ("AE02", "Market Share Growth (pp)",
         "رشد سهم بازار",
         f"=IFERROR(AVERAGE({EX}!D2:D13)/AVERAGE({EX}!C2:C13)"
         f"-AVERAGE({EX}!E2:E13)/AVERAGE({EX}!C2:C13),0)",
         "0.00%"),

        ("AE03", "Relative Market Share",
         "سهم بازار نسبی",
         f"=IFERROR((AVERAGE({EX}!D2:D13)/AVERAGE({EX}!C2:C13))"
         f"/(AVERAGE({EX}!E2:E13)/AVERAGE({EX}!C2:C13)),0)",
         "0.00"),

        ("AE04", "Brand Awareness (%)",
         "آگاهی از برند",
         f"=IFERROR(AVERAGE({EX}!G2:G13),0)",
         "0.00%"),

        ("AE05", "Brand Health Index",
         "شاخص سلامت برند",
         f"=IFERROR((AVERAGE({EX}!G2:G13)+AVERAGE({EX}!H2:H13)/10)/2,0)",
         "0.00"),

        ("AE06", "Forecast Accuracy (%)",
         "دقت پیش‌بینی",
         f"=IFERROR(1-AVERAGE({EX}!I2:I13),0)",
         "0.00%"),

        ("AE07", "MAPE (%)",
         "میانگین درصد خطای مطلق",
         f"=IFERROR(AVERAGE({EX}!I2:I13),0)",
         "0.00%"),

        ("AE08", "Demand Volatility (CV)",
         "نوسان تقاضا",
         f"=IFERROR(STDEV({S}!N2:N5001)/AVERAGE({S}!N2:N5001),0)",
         "0.00"),

        ("AE09", "Promotional Uplift (%)",
         "افزایش فروش ناشی از تبلیغات",
         f"=IFERROR(SUMIF({S}!P2:P5001,\">0\",{S}!U2:U5001)"
         f"/SUMIF({S}!P2:P5001,0,{S}!U2:U5001)-1,0)",
         "0.00%"),

        ("AE10", "Promo ROI",
         "بازده سرمایه‌گذاری تبلیغاتی",
         f"=IFERROR((SUMIF({S}!P2:P5001,\">0\",{S}!U2:U5001)"
         f"-SUMIF({S}!P2:P5001,0,{S}!U2:U5001))"
         f"/SUM({S}!R2:R5001),0)",
         "0.00"),

        ("AE11", "Category Growth Rate (%)",
         "نرخ رشد دسته‌بندی",
         f"=IFERROR(AVERAGE({EX}!J2:J13),0)",
         "0.00%"),

        ("AE12", "Price Index vs Competition",
         "شاخص قیمت در مقابل رقبا",
         f"=IFERROR(AVERAGE({S}!O2:O5001)/AVERAGE({EX}!K2:K13),0)",
         "0.00"),

        ("AE13", "Price Elasticity",
         "کشش قیمتی",
         f"=IFERROR((SUMIF({S}!P2:P5001,\">0\",{S}!N2:N5001)"
         f"/SUMIF({S}!P2:P5001,0,{S}!N2:N5001)-1)"
         f"/(AVERAGE({S}!O2:O5001)/AVERAGE({S}!O2:O5001)-1+0.001),0)",
         "0.00"),

        ("AE14", "Innovation Revenue Share (%)",
         "سهم درآمد محصولات جدید",
         f"=IFERROR(SUMPRODUCT(({S}!U2:U5001)*"
         f"(VLOOKUP({S}!G2:G5001,{PR}!$A:$I,9,0)=\"New\")),0)",
         "0.00%"),

        ("AE15", "SKU Productivity ($)",
         "بهره‌وری SKU",
         f"=IFERROR(SUM({S}!U2:U5001)/COUNTA(UNIQUE({S}!G2:G5001)),0)",
         "$#,##0"),

        ("AE16", "Long Tail SKU Count",
         "تعداد SKU کم‌فروش",
         f"=COUNTIF({IN}!K2:K1001,\"C\")",
         "#,##0"),

        ("AE17", "Revenue Concentration (Top 5 SKU %)",
         "تمرکز درآمد ۵ SKU برتر",
         f"=IFERROR(SUMPRODUCT(LARGE(SUMIF({S}!G2:G5001,{PR}!A2:A51,{S}!U2:U5001),"
         f"{{1,2,3,4,5}}))/SUM({S}!U2:U5001),0)",
         "0.00%"),

        ("AE18", "Customer Concentration (Top 5 %)",
         "تمرکز مشتری ۵ مشتری برتر",
         f"=IFERROR(SUMPRODUCT(LARGE(SUMIF({S}!F2:F5001,{CU}!A2:A201,{S}!U2:U5001),"
         f"{{1,2,3,4,5}}))/SUM({S}!U2:U5001),0)",
         "0.00%"),

        ("AE19", "RFM — Avg Recency (Days)",
         "RFM — میانگین تازگی خرید",
         f"=IFERROR(AVERAGE(TODAY()-MAXIFS({S}!B2:B5001,{S}!F2:F5001,{CU}!A2:A201)),0)",
         "0.0"),

        ("AE20", "RFM — Avg Frequency",
         "RFM — میانگین تناوب",
         f"=IFERROR(AVERAGE(COUNTIF({S}!F2:F5001,{CU}!A2:A201)),0)",
         "0.0"),

        ("AE21", "RFM — Avg Monetary ($)",
         "RFM — میانگین ارزش پولی",
         f"=IFERROR(AVERAGE(SUMIF({S}!F2:F5001,{CU}!A2:A201,{S}!U2:U5001)),0)",
         "$#,##0"),

        ("AE22", "Sustainability Score",
         "امتیاز پایداری",
         f"=IFERROR(AVERAGE({EX}!M2:M13),0)",
         "0.00"),

        ("AE23", "Carbon Footprint Index",
         "شاخص ردپای کربن",
         f"=IFERROR(SUM({S}!N2:N5001)*VLOOKUP(\"Carbon_Per_Unit\",{AS}!$A:$B,2,0),0)",
         "#,##0.0"),

        ("AE24", "Digital Channel Share (%)",
         "سهم کانال دیجیتال",
         f"=IFERROR(SUMIF({S}!K2:K5001,\"E-Commerce\",{S}!U2:U5001)"
         f"/SUM({S}!U2:U5001),0)",
         "0.00%"),

        ("AE25", "Omnichannel Revenue ($)",
         "درآمد همه‌کانالی",
         f"=SUM({S}!U2:U5001)",
         "$#,##0"),

        ("AE26", "Cross-sell Rate (%)",
         "نرخ فروش متقاطع",
         f"=IFERROR(SUMPRODUCT((COUNTIFS({S}!F2:F5001,{S}!F2:F5001,"
         f"{S}!L2:L5001,\">\"&{S}!L2:L5001)>0)*1)"
         f"/COUNTA({S}!A2:A5001),0)",
         "0.00%"),

        ("AE27", "Upsell Rate (%)",
         "نرخ فروش ارتقایی",
         f"=IFERROR(SUMPRODUCT(({S}!O2:O5001>AVERAGE({S}!O2:O5001))*1)"
         f"/COUNTA({S}!A2:A5001),0)",
         "0.00%"),

        ("AE28", "Supply Chain Resilience Score",
         "امتیاز تاب‌آوری زنجیره تأمین",
         f"=IFERROR((1-D{r+3})*(1-D{r+4})*100,0)",
         "0.0"),

        ("AE29", "Vendor Fill Rate (%)",
         "نرخ تأمین تأمین‌کننده",
         f"=IFERROR(AVERAGE({TG}!H2:H1000),0)",
         "0.00%"),

        ("AE30", "Overall Business Health Score",
         "امتیاز کلی سلامت کسب‌وکار",
         f"=IFERROR(("
         f"IFERROR(SUM({S}!X2:X5001)/SUM({S}!U2:U5001),0)*20+"
         f"IFERROR(1-COUNTIF({IN}!D2:D1001,0)/COUNTA({IN}!A2:A1001),0)*20+"
         f"IFERROR(SUMIF({V}!I2:I4001,1,{V}!I2:I4001)/SUMIF({V}!H2:H4001,1,{V}!H2:H4001),0)*20+"
         f"IFERROR(AVERAGE({EX}!D2:D13)/AVERAGE({EX}!C2:C13),0)*20+"
         f"IFERROR(1-AVERAGE({EX}!I2:I13),0)*20),0)",
         "0.0"),
    ]

    for code, name_en, name_fa, formula, fmt in kpis:
        kpi_row(ws, r, code, name_en, name_fa, formula, fmt)
        r += 1

    return r


# ─────────────────────────────────────────────
# DASHBOARD SHEET
# ─────────────────────────────────────────────
def build_dashboard(wb):
    ws = wb.create_sheet("Dashboard")
    ws.sheet_view.rightToLeft = False

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 28
    ws.column_dimensions["C"].width = 28
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 18
    ws.column_dimensions["F"].width = 14
    ws.column_dimensions["G"].width = 8

    # Title
    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = "FMCG Analytics Dashboard | داشبورد تحلیلی"
    c.font = Font(bold=True, size=16, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor="1F3864")
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    headers = ["Code", "KPI (English)", "KPI (فارسی)", "Value", "Target", "vs Target", "Status"]
    cols = ["A", "B", "C", "D", "E", "F", "G"]
    for col, h in zip(cols, headers):
        c = ws[f"{col}2"]
        c.value = h
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="2E75B6")
        c.alignment = Alignment(horizontal="center")
        c.border = THIN_BORDER

    # Pull top KPIs from each category sheet
    summary_kpis = [
        ("SM01", "Net Revenue ($)", "درآمد خالص", "KPI_Sales", 4),
        ("SM13", "Gross Profit Margin (%)", "حاشیه سود ناخالص", "KPI_Sales", 16),
        ("SM21", "Sales Growth YoY (%)", "رشد فروش سالانه", "KPI_Sales", 24),
        ("CA01", "Active Customers", "مشتریان فعال", "KPI_Customer", 4),
        ("CA04", "Retention Rate (%)", "نرخ نگهداشت", "KPI_Customer", 7),
        ("CA07", "CLV ($)", "ارزش طول عمر مشتری", "KPI_Customer", 10),
        ("IN02", "Inventory Turnover", "گردش موجودی", "KPI_Inventory", 5),
        ("IN04", "Stockout Rate (%)", "نرخ کمبود موجودی", "KPI_Inventory", 7),
        ("DR06", "Strike Rate (%)", "نرخ تبدیل بازدید", "KPI_Dist", 9),
        ("DR10", "On-Time Delivery (%)", "تحویل به‌موقع", "KPI_Dist", 13),
        ("FI01", "Net Margin (%)", "حاشیه سود خالص", "KPI_Finance", 4),
        ("FI06", "DSO (Days)", "روزهای وصول مطالبات", "KPI_Finance", 9),
        ("AE01", "Market Share (%)", "سهم بازار", "KPI_Advanced", 4),
        ("AE06", "Forecast Accuracy (%)", "دقت پیش‌بینی", "KPI_Advanced", 9),
        ("AE30", "Business Health Score", "امتیاز سلامت کسب‌وکار", "KPI_Advanced", 33),
    ]

    for i, (code, en, fa, sheet, row) in enumerate(summary_kpis, start=1):
        r = i + 2
        ws[f"A{r}"] = code
        ws[f"B{r}"] = en
        ws[f"C{r}"] = fa
        ws[f"D{r}"] = f"={sheet}!D{row}"
        ws[f"E{r}"] = f"=VLOOKUP(\"{code}\",Assumptions!$A:$C,2,0)*{ws[f'D{r}'].value}"
        ws[f"F{r}"] = f"=IFERROR(D{r}/E{r}-1,\"N/A\")"
        ws[f"G{r}"] = f"=IF(ISNUMBER(F{r}),IF(F{r}>=0,\"✅\",IF(F{r}>=-0.1,\"⚠️\",\"🔴\")),\"-\")"

        for col in ["A", "B", "C", "D", "E", "F", "G"]:
            cell = ws[f"{col}{r}"]
            cell.border = THIN_BORDER
            if col in ["A", "F", "G"]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
            
            if i % 2 == 0 and col != "G":
                cell.fill = ALT_FILL

        ws[f"D{r}"].number_format = "#,##0.00"
        ws[f"F{r}"].number_format = "0.0%"
        ws.row_dimensions[r].height = 20

    return ws


# ─────────────────────────────────────────────
# MAIN BUILD
# ─────────────────────────────────────────────
def build_workbook():
    wb = Workbook()
    wb.remove(wb.active)

    print("📊 Generating data...")
    
    # Generate all data
    df_sales = make_sales()
    df_customers = make_customers()
    df_products = make_products()
    df_inventory = make_inventory()
    df_visits = make_visits()
    df_ar = make_ar_aging()
    df_external = make_external()
    df_targets = make_targets()
    df_assumptions = make_assumptions()

    print("📝 Creating data sheets...")
    
    # Create and populate data sheets
    ws_sales = create_data_sheet(wb, "Data_Sales")
    populate_data_sheet(ws_sales, df_sales)
    
    ws_customers = create_data_sheet(wb, "Data_Customers")
    populate_data_sheet(ws_customers, df_customers)
    
    ws_products = create_data_sheet(wb, "Data_Products")
    populate_data_sheet(ws_products, df_products)
    
    ws_inventory = create_data_sheet(wb, "Data_Inventory")
    populate_data_sheet(ws_inventory, df_inventory)
    
    ws_visits = create_data_sheet(wb, "Data_Visits")
    populate_data_sheet(ws_visits, df_visits)
    
    ws_ar = create_data_sheet(wb, "Data_AR_Aging")
    populate_data_sheet(ws_ar, df_ar)
    
    ws_external = create_data_sheet(wb, "Data_External")
    populate_data_sheet(ws_external, df_external)
    
    ws_targets = create_data_sheet(wb, "Data_Targets")
    populate_data_sheet(ws_targets, df_targets)
    
    ws_assume = create_data_sheet(wb, "Assumptions")
    populate_data_sheet(ws_assume, df_assumptions)

    print("📈 Creating KPI sheets...")
    
    # Create KPI sheets
    ws_sm = wb.create_sheet("KPI_Sales")
    ws_ca = wb.create_sheet("KPI_Customer")
    ws_in = wb.create_sheet("KPI_Inventory")
    ws_dr = wb.create_sheet("KPI_Dist")
    ws_fi = wb.create_sheet("KPI_Finance")
    ws_ae = wb.create_sheet("KPI_Advanced")

    build_kpi_sales(ws_sm)
    build_kpi_customer(ws_ca)
    build_kpi_inventory(ws_in)
    build_kpi_distribution(ws_dr)
    build_kpi_finance(ws_fi)
    build_kpi_advanced(ws_ae)

    print("📊 Creating dashboard...")
    build_dashboard(wb)

    # Save the workbook
    path = "FMCG_Analysis_Template.xlsx"
    wb.save(path)
    print(f"✅ Successfully saved: {path}")
    print(f"📁 File size: {__import__('os').path.getsize(path) / 1024:.1f} KB")
    
    return wb


if __name__ == "__main__":
    try:
        build_workbook()
    except Exception as e:
        print(f"❌ Error: {e}")
        import traceback
        traceback.print_exc()