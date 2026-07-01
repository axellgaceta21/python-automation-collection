"""
sales_dashboard_generator.py

Project:
Sales Dashboard Generator

Purpose:
Clean an orders export and generate a formatted sales dashboard workbook.

Expected folder structure:

Python/
    datasets/
        sales_orders_messy.xlsx
    scripts/
        sales_dashboard_generator.py
    outputs/
    archive/

Run from the Python folder:
    py scripts/sales_dashboard_generator.py

Output format:
    outputs/sales_dashboard_06.30_1.xlsx
    outputs/sales_dashboard_06.30_2.xlsx

Output sheets:
    Dashboard
    Orders_Clean
    Orders_Invalid
    Monthly_Sales
    Category_Sales
    Channel_Sales
    Top_Products

Main features:
    - Cleans and validates order data
    - Calculates revenue, net sales, AOV, orders, refunds, and customer counts
    - Generates summary tables for monthly sales, categories, channels, and products
    - Creates a formatted Excel dashboard with KPI cards and charts
    - Moves invalid rows to Orders_Invalid with a final Issues column
"""

import re
from pathlib import Path
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, LineChart, PieChart, Reference


BASE_DIR = Path(__file__).resolve().parent.parent
DATASETS_DIR = BASE_DIR / "datasets"
OUTPUTS_DIR = BASE_DIR / "outputs"
ARCHIVE_DIR = BASE_DIR / "archive"

INPUT_FILE = DATASETS_DIR / "sales_orders_messy.xlsx"

DATASETS_DIR.mkdir(exist_ok=True)
OUTPUTS_DIR.mkdir(exist_ok=True)
ARCHIVE_DIR.mkdir(exist_ok=True)


def get_next_output_file(output_dir: Path, base_name: str = "sales_dashboard", extension: str = ".xlsx") -> Path:
    today = datetime.now().strftime("%m.%d")
    existing_files = list(output_dir.glob(f"{base_name}_{today}_*{extension}"))

    versions = []
    for file in existing_files:
        match = re.search(rf"^{re.escape(base_name)}_{re.escape(today)}_(\d+)$", file.stem)
        if match:
            versions.append(int(match.group(1)))

    next_version = max(versions, default=0) + 1
    return output_dir / f"{base_name}_{today}_{next_version}{extension}"


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def parse_date(value):
    if pd.isna(value) or str(value).strip() == "":
        return pd.NaT
    return pd.to_datetime(value, errors="coerce")


def clean_int(value):
    if pd.isna(value) or str(value).strip() == "":
        return None
    try:
        return int(value)
    except ValueError:
        return None


def clean_float(value):
    if pd.isna(value) or str(value).strip() == "":
        return None
    cleaned = str(value).replace(",", "").strip()
    try:
        return float(cleaned)
    except ValueError:
        return None


def normalize_status(value):
    return clean_text(value).title()


def add_issue(issues_by_index, row_index, issue):
    if row_index not in issues_by_index:
        issues_by_index[row_index] = []
    issues_by_index[row_index].append(issue)


def count_total_issues(issues_by_index):
    return sum(len(issues) for issues in issues_by_index.values())


def split_valid_invalid(df, issues_by_index):
    invalid_indexes = sorted(issues_by_index.keys())
    valid_df = df.drop(index=invalid_indexes, errors="ignore").copy()
    invalid_df = df.loc[df.index.intersection(invalid_indexes)].copy()

    if len(invalid_df) > 0:
        invalid_df["Issues"] = invalid_df.index.map(
            lambda idx: "; ".join(sorted(set(issues_by_index.get(idx, []))))
        )
        invalid_df["_original_index"] = invalid_df.index
        invalid_df = invalid_df.sort_values(
            by=["Order ID", "_original_index"],
            kind="stable"
        ).drop(columns=["_original_index"])
    else:
        invalid_df["Issues"] = []

    return valid_df.reset_index(drop=True), invalid_df.reset_index(drop=True)


def clean_orders(df):
    issues_by_index = {}

    df = df.copy()
    df.columns = [str(col).strip() for col in df.columns]

    required_columns = [
        "Order ID", "Order Date", "Customer ID", "Customer Name", "Country",
        "Channel", "Product", "Category", "Quantity", "Unit Price",
        "Discount", "Shipping Fee", "Payment Status", "Fulfillment Status"
    ]

    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {missing_columns}")

    df["Order ID"] = df["Order ID"].apply(clean_text)
    df["Order Date"] = df["Order Date"].apply(parse_date)
    df["Customer ID"] = df["Customer ID"].apply(clean_text)
    df["Customer Name"] = df["Customer Name"].apply(lambda x: clean_text(x).title())
    df["Country"] = df["Country"].apply(lambda x: clean_text(x).title())
    df["Channel"] = df["Channel"].apply(clean_text)
    df["Product"] = df["Product"].apply(lambda x: clean_text(x).title())
    df["Category"] = df["Category"].apply(lambda x: clean_text(x).title())
    df["Quantity"] = df["Quantity"].apply(clean_int)
    df["Unit Price"] = df["Unit Price"].apply(clean_float)
    df["Discount"] = df["Discount"].apply(clean_float).fillna(0)
    df["Shipping Fee"] = df["Shipping Fee"].apply(clean_float).fillna(0)
    df["Payment Status"] = df["Payment Status"].apply(normalize_status)
    df["Fulfillment Status"] = df["Fulfillment Status"].apply(normalize_status)

    for index, row in df.iterrows():
        if not row["Order ID"]:
            add_issue(issues_by_index, index, "Missing order ID")

        if pd.isna(row["Order Date"]):
            add_issue(issues_by_index, index, "Missing or invalid order date")

        if not row["Customer ID"]:
            add_issue(issues_by_index, index, "Missing customer ID")

        if not row["Product"]:
            add_issue(issues_by_index, index, "Missing product")

        if not row["Category"]:
            add_issue(issues_by_index, index, "Missing category")

        if pd.isna(row["Quantity"]):
            add_issue(issues_by_index, index, "Missing or invalid quantity")
        elif row["Quantity"] <= 0:
            add_issue(issues_by_index, index, "Quantity is zero or negative")

        if pd.isna(row["Unit Price"]):
            add_issue(issues_by_index, index, "Missing or invalid unit price")
        elif row["Unit Price"] <= 0:
            add_issue(issues_by_index, index, "Unit price is zero or negative")

    duplicate_mask = df.duplicated(
        subset=["Order ID", "Customer ID", "Product", "Quantity", "Unit Price"],
        keep=False
    ) & df["Order ID"].ne("")

    for index in df[duplicate_mask].index:
        add_issue(issues_by_index, index, "Duplicate order line")

    clean_df, invalid_df = split_valid_invalid(df, issues_by_index)

    clean_df["Gross Sales"] = clean_df["Quantity"] * clean_df["Unit Price"]
    clean_df["Net Sales"] = clean_df["Gross Sales"] - clean_df["Discount"] + clean_df["Shipping Fee"]
    clean_df.loc[clean_df["Payment Status"].eq("Refunded"), "Net Sales"] = 0
    clean_df["Month"] = clean_df["Order Date"].dt.to_period("M").astype(str)

    return clean_df, invalid_df, issues_by_index


def build_summaries(clean_df):
    revenue_df = clean_df[clean_df["Payment Status"].eq("Paid")].copy()

    monthly = (
        revenue_df.groupby("Month", as_index=False)
        .agg(
            Orders=("Order ID", "nunique"),
            Revenue=("Net Sales", "sum"),
            Units=("Quantity", "sum"),
            Customers=("Customer ID", "nunique")
        )
        .sort_values("Month")
    )
    monthly["AOV"] = monthly["Revenue"] / monthly["Orders"]

    category = (
        revenue_df.groupby("Category", as_index=False)
        .agg(
            Orders=("Order ID", "nunique"),
            Revenue=("Net Sales", "sum"),
            Units=("Quantity", "sum")
        )
        .sort_values("Revenue", ascending=False)
    )

    channel = (
        revenue_df.groupby("Channel", as_index=False)
        .agg(
            Orders=("Order ID", "nunique"),
            Revenue=("Net Sales", "sum"),
            Units=("Quantity", "sum")
        )
        .sort_values("Revenue", ascending=False)
    )

    top_products = (
        revenue_df.groupby(["Product", "Category"], as_index=False)
        .agg(
            Orders=("Order ID", "nunique"),
            Revenue=("Net Sales", "sum"),
            Units=("Quantity", "sum")
        )
        .sort_values("Revenue", ascending=False)
        .head(15)
    )

    return monthly, category, channel, top_products


def safe_currency(value):
    return float(value) if pd.notna(value) else 0.0


def format_workbook(output_file):
    wb = load_workbook(output_file)

    dark_fill = PatternFill("solid", fgColor="111827")
    blue_fill = PatternFill("solid", fgColor="1D4ED8")
    green_fill = PatternFill("solid", fgColor="166534")
    red_fill = PatternFill("solid", fgColor="991B1B")
    gold_fill = PatternFill("solid", fgColor="FEF3C7")
    issue_fill = PatternFill("solid", fgColor="F4CCCC")
    invalid_cell_fill = PatternFill("solid", fgColor="FCE4E4")
    summary_fill = PatternFill("solid", fgColor="E0F2FE")

    white_font = Font(color="FFFFFF", bold=True)
    title_font = Font(color="111827", bold=True, size=18)
    kpi_font = Font(color="111827", bold=True, size=16)
    thin_gray = Side(style="thin", color="D1D5DB")
    border = Border(left=thin_gray, right=thin_gray, top=thin_gray, bottom=thin_gray)

    def get_col(ws, name):
        for cell in ws[1]:
            if cell.value == name:
                return cell.column
        return None

    def add_table(ws, table_name):
        if ws.max_row >= 2 and ws.max_column >= 1:
            ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
            table = Table(displayName=table_name, ref=ref)
            table.tableStyleInfo = TableStyleInfo(
                name="TableStyleMedium2",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False,
            )
            try:
                ws.add_table(table)
            except ValueError:
                pass

    # General formatting
    for ws in wb.worksheets:
        if ws.title == "Dashboard":
            continue

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        for cell in ws[1]:
            cell.fill = dark_fill
            cell.font = white_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border

        if "Clean" in ws.title or "Valid" in ws.title:
            for cell in ws[1]:
                cell.fill = green_fill

        if "Invalid" in ws.title:
            for cell in ws[1]:
                cell.fill = red_fill

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(vertical="center", wrap_text=True)

        for col_name in ["Order Date"]:
            col = get_col(ws, col_name)
            if col:
                for row in range(2, ws.max_row + 1):
                    ws.cell(row, col).number_format = "yyyy-mm-dd"

        for col_name in ["Unit Price", "Discount", "Shipping Fee", "Gross Sales", "Net Sales", "Revenue", "AOV"]:
            col = get_col(ws, col_name)
            if col:
                for row in range(2, ws.max_row + 1):
                    ws.cell(row, col).number_format = "$#,##0.00"

        for col_name in ["Quantity", "Orders", "Units", "Customers"]:
            col = get_col(ws, col_name)
            if col:
                for row in range(2, ws.max_row + 1):
                    ws.cell(row, col).number_format = "#,##0"

        if "Invalid" in ws.title:
            issues_col = get_col(ws, "Issues")
            if issues_col:
                ws.cell(1, issues_col).fill = PatternFill("solid", fgColor="D99694")
                ws.cell(1, issues_col).font = Font(color="000000", bold=True)
                issue_map = {
                    "Missing order ID": ["Order ID"],
                    "Missing or invalid order date": ["Order Date"],
                    "Missing customer ID": ["Customer ID"],
                    "Missing product": ["Product"],
                    "Missing category": ["Category"],
                    "Missing or invalid quantity": ["Quantity"],
                    "Quantity is zero or negative": ["Quantity"],
                    "Missing or invalid unit price": ["Unit Price"],
                    "Unit price is zero or negative": ["Unit Price"],
                    "Duplicate order line": ["Order ID", "Customer ID", "Product"],
                }
                for row in range(2, ws.max_row + 1):
                    issues = str(ws.cell(row, issues_col).value or "")
                    ws.cell(row, issues_col).fill = issue_fill
                    for issue, cols in issue_map.items():
                        if issue in issues:
                            for col_name in cols:
                                col = get_col(ws, col_name)
                                if col:
                                    ws.cell(row, col).fill = invalid_cell_fill

        # Widths
        for col_idx in range(1, ws.max_column + 1):
            header = str(ws.cell(1, col_idx).value or "")
            width = 14
            if header in ["Customer Name", "Product", "Fulfillment Status"]:
                width = 24
            elif header in ["Issues"]:
                width = 55
            elif header in ["Order Date", "Customer ID", "Payment Status"]:
                width = 18
            ws.column_dimensions[get_column_letter(col_idx)].width = width

        add_table(ws, re.sub(r"[^A-Za-z0-9_]", "", ws.title)[:20] + "Table")

    # Dashboard formatting
    ws = wb["Dashboard"]
    ws.sheet_view.showGridLines = False

    ws["A1"] = "Sales Dashboard"
    ws["A1"].font = title_font
    ws["A2"] = "Generated summary from cleaned sales order data"
    ws["A2"].font = Font(color="6B7280", italic=True)

    # KPI cards
    card_ranges = ["A4:C6", "D4:F6", "G4:I6", "J4:L6", "A8:C10", "D8:F10"]
    for card in card_ranges:
        rng = ws[card]
        for row in rng:
            for cell in row:
                cell.fill = summary_fill
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")

    for cell in ["A4", "D4", "G4", "J4", "A8", "D8"]:
        ws[cell].font = Font(color="374151", bold=True, size=11)

    for cell in ["A5", "D5", "G5", "J5", "A9", "D9"]:
        ws[cell].font = kpi_font

    for col in range(1, 13):
        ws.column_dimensions[get_column_letter(col)].width = 15

    # Charts
    if wb["Monthly_Sales"].max_row >= 2:
        mws = wb["Monthly_Sales"]
        chart = LineChart()
        chart.title = "Monthly Revenue"
        chart.y_axis.title = "Revenue"
        chart.x_axis.title = "Month"
        data = Reference(mws, min_col=3, min_row=1, max_row=mws.max_row)
        cats = Reference(mws, min_col=1, min_row=2, max_row=mws.max_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.height = 8
        chart.width = 18
        ws.add_chart(chart, "A13")

    if wb["Category_Sales"].max_row >= 2:
        cws = wb["Category_Sales"]
        chart2 = BarChart()
        chart2.title = "Revenue by Category"
        chart2.y_axis.title = "Revenue"
        data = Reference(cws, min_col=3, min_row=1, max_row=min(cws.max_row, 8))
        cats = Reference(cws, min_col=1, min_row=2, max_row=min(cws.max_row, 8))
        chart2.add_data(data, titles_from_data=True)
        chart2.set_categories(cats)
        chart2.height = 8
        chart2.width = 18
        ws.add_chart(chart2, "G13")

    if wb["Channel_Sales"].max_row >= 2:
        chws = wb["Channel_Sales"]
        chart3 = PieChart()
        chart3.title = "Revenue by Channel"
        data = Reference(chws, min_col=3, min_row=1, max_row=min(chws.max_row, 8))
        cats = Reference(chws, min_col=1, min_row=2, max_row=min(chws.max_row, 8))
        chart3.add_data(data, titles_from_data=True)
        chart3.set_categories(cats)
        chart3.height = 8
        chart3.width = 14
        ws.add_chart(chart3, "A29")

    # Freeze and sheet order
    ws.freeze_panes = "A4"

    wb.save(output_file)


def main():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Input file not found:\n{INPUT_FILE}\n\n"
            "Place sales_orders_messy.xlsx inside the datasets folder."
        )

    output_file = get_next_output_file(OUTPUTS_DIR)

    print("Reading sales orders...")
    print(f"Input: {INPUT_FILE}")

    df = pd.read_excel(INPUT_FILE, sheet_name="Orders_Raw")

    print("Cleaning orders...")
    clean_df, invalid_df, issues = clean_orders(df)

    print("Building summaries...")
    monthly, category, channel, top_products = build_summaries(clean_df)

    paid_df = clean_df[clean_df["Payment Status"].eq("Paid")].copy()
    total_revenue = paid_df["Net Sales"].sum()
    total_orders = paid_df["Order ID"].nunique()
    total_units = paid_df["Quantity"].sum()
    unique_customers = paid_df["Customer ID"].nunique()
    aov = total_revenue / total_orders if total_orders else 0
    refunds = clean_df[clean_df["Payment Status"].eq("Refunded")]["Order ID"].nunique()

    dashboard = pd.DataFrame([
        ["Total Revenue", total_revenue, "", "Total Orders", total_orders, "", "Units Sold", total_units, "", "AOV", aov, ""],
        ["Unique Customers", unique_customers, "", "Refunded Orders", refunds, "", "", "", "", "", "", ""],
    ])

    print("Writing dashboard workbook...")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        dashboard.to_excel(writer, sheet_name="Dashboard", index=False, header=False, startrow=3, startcol=0)
        clean_df.to_excel(writer, sheet_name="Orders_Clean", index=False)
        invalid_df.to_excel(writer, sheet_name="Orders_Invalid", index=False)
        monthly.to_excel(writer, sheet_name="Monthly_Sales", index=False)
        category.to_excel(writer, sheet_name="Category_Sales", index=False)
        channel.to_excel(writer, sheet_name="Channel_Sales", index=False)
        top_products.to_excel(writer, sheet_name="Top_Products", index=False)

    print("Applying Excel formatting and charts...")
    format_workbook(output_file)

    print("\nSales dashboard complete.")
    print(f"Created: {output_file.name}")
    print(f"Saved in: {OUTPUTS_DIR}")
    print(f"Clean order rows: {len(clean_df)}")
    print(f"Invalid rows moved: {len(invalid_df)}")
    print(f"Issue count: {count_total_issues(issues)}")
    print(f"Total revenue: ${total_revenue:,.2f}")
    print(f"Total orders: {total_orders:,}")


if __name__ == "__main__":
    main()
