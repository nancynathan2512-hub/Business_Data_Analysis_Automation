import pandas as pd
from pathlib import Path

# --------------------------------------------------
# PATHS
# --------------------------------------------------

BASE_DIR = Path(__file__).resolve().parent.parent

INPUT_FILE = (
    BASE_DIR
    / "Data"
    / "Cleaned Data"
    / "transformed_retail_data.xlsx"
)

OUTPUT_DIR = BASE_DIR / "Data" / "Business Reports"
OUTPUT_FILE = OUTPUT_DIR / "business_report.xlsx"


# --------------------------------------------------
# START
# --------------------------------------------------

print("=" * 60)
print("BUSINESS REPORT GENERATION STARTED")
print("=" * 60)

print("\nLoading transformed data...")
print(INPUT_FILE)

data = pd.read_excel(INPUT_FILE)

print(f"\nRows loaded: {len(data):,}")
print(f"Columns loaded: {len(data.columns)}")


# --------------------------------------------------
# KPI SUMMARY
# --------------------------------------------------

print("\nCalculating KPI summary...")

total_revenue = data["Revenue"].sum()
total_quantity = data["Quantity"].sum()
total_invoices = data["Invoice"].nunique()
total_customers = data["Customer ID"].nunique()
total_products = data["StockCode"].nunique()

kpi_summary = pd.DataFrame({
    "Metric": [
        "Total Revenue",
        "Total Quantity Sold",
        "Total Invoices",
        "Total Customers",
        "Total Products"
    ],
    "Value": [
        total_revenue,
        total_quantity,
        total_invoices,
        total_customers,
        total_products
    ]
})


# --------------------------------------------------
# MONTHLY REVENUE
# --------------------------------------------------

print("Creating monthly revenue...")

monthly_revenue = (
    data.groupby("Year_Month", as_index=False)
    .agg(
        Revenue=("Revenue", "sum"),
        Quantity=("Quantity", "sum"),
        Invoices=("Invoice", "nunique")
    )
    .sort_values("Year_Month")
)


# --------------------------------------------------
# COUNTRY REVENUE
# --------------------------------------------------

print("Creating country revenue...")

country_revenue = (
    data.groupby("Country", as_index=False)
    .agg(
        Revenue=("Revenue", "sum"),
        Quantity=("Quantity", "sum"),
        Invoices=("Invoice", "nunique"),
        Customers=("Customer ID", "nunique")
    )
    .sort_values("Revenue", ascending=False)
)


# --------------------------------------------------
# TOP PRODUCTS
# --------------------------------------------------

print("Creating top products...")

top_products = (
    data.groupby(
        ["StockCode", "Description"],
        as_index=False
    )
    .agg(
        Revenue=("Revenue", "sum"),
        Quantity=("Quantity", "sum"),
        Invoices=("Invoice", "nunique")
    )
    .sort_values("Revenue", ascending=False)
    .head(20)
)


# --------------------------------------------------
# TOP CUSTOMERS
# --------------------------------------------------

print("Creating top customers...")

top_customers = (
    data.groupby("Customer ID", as_index=False)
    .agg(
        Revenue=("Revenue", "sum"),
        Quantity=("Quantity", "sum"),
        Invoices=("Invoice", "nunique")
    )
    .sort_values("Revenue", ascending=False)
    .head(20)
)


# --------------------------------------------------
# CREATE OUTPUT FOLDER
# --------------------------------------------------

OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# --------------------------------------------------
# SAVE BUSINESS REPORT
# --------------------------------------------------

print("\nSaving business report...")

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    kpi_summary.to_excel(
        writer,
        sheet_name="KPI Summary",
        index=False
    )

    monthly_revenue.to_excel(
        writer,
        sheet_name="Monthly Revenue",
        index=False
    )

    country_revenue.to_excel(
        writer,
        sheet_name="Country Revenue",
        index=False
    )

    top_products.to_excel(
        writer,
        sheet_name="Top Products",
        index=False
    )

    top_customers.to_excel(
        writer,
        sheet_name="Top Customers",
        index=False
    )


# --------------------------------------------------
# COMPLETE
# --------------------------------------------------

print("\n" + "=" * 60)
print("BUSINESS REPORT COMPLETE")
print("=" * 60)

print("\nReport contains:")
print("- KPI Summary")
print("- Monthly Revenue")
print("- Country Revenue")
print("- Top Products")
print("- Top Customers")

print("\nReport created at:")
print(OUTPUT_FILE)
