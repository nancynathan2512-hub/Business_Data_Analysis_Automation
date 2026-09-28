import pandas as pd

REPORT_FILE = r"..\Data\Business Reports\business_report.xlsx"

# Read the business report
kpi = pd.read_excel(REPORT_FILE, sheet_name="KPI Summary")

print("KPI Summary:")
print(kpi)

total_revenue = kpi.loc[0, "Value"]
total_quantity = kpi.loc[1, "Value"]
total_invoices = kpi.loc[2, "Value"]
total_customers = kpi.loc[3, "Value"]
total_products = kpi.loc[4, "Value"]

print("\nExtracted Business KPIs:")
print("Revenue:", total_revenue)
print("Quantity:", total_quantity)
print("Invoices:", total_invoices)
print("Customers:", total_customers)
print("Products:", total_products)
business_data = {
    "total_revenue": total_revenue,
    "total_quantity_sold": total_quantity,
    "total_invoices": total_invoices,
    "total_customers": total_customers,
    "total_products": total_products
}

print("\nBusiness data prepared for AI:")
print(business_data)
ai_prompt = f"""
You are a business data analyst.

Analyze the following business KPIs:

Total Revenue: {total_revenue}
Total Quantity Sold: {total_quantity}
Total Invoices: {total_invoices}
Total Customers: {total_customers}
Total Products: {total_products}

Provide:
1. A short business summary.
2. Three key findings.
3. Three practical business recommendations.

Keep the analysis concise and business-focused.
"""

print("\nAI Prompt prepared:")
print(ai_prompt)
# Generate data-driven business insights

monthly = pd.read_excel(
    REPORT_FILE,
    sheet_name="Monthly Revenue"
)


country = pd.read_excel(
    REPORT_FILE,
    sheet_name="Country Revenue"
)

products = pd.read_excel(
    REPORT_FILE,
    sheet_name="Top Products"
)

# Find top-performing month
top_month = monthly.loc[
    monthly["Revenue"].idxmax()
]

# Find top-performing country
top_country = country.loc[
    country["Revenue"].idxmax()
]

# Find top-performing product
top_product = products.loc[
    products["Revenue"].idxmax()
]

ai_response = {
    "summary": (
        f"The business generated total revenue of "
        f"{total_revenue:,.2f} across {total_customers:,.0f} customers. "
        f"The highest-revenue country was {top_country['Country']}, "
        f"while the highest-revenue product was "
        f"{top_product['Description']}."
    ),

    "key_findings": [
        f"Highest-revenue month: {top_month['Year_Month']}"
        f"with revenue of {top_month['Revenue']:,.2f}.",

        f"Highest-revenue country: {top_country['Country']} "
        f"with revenue of {top_country['Revenue']:,.2f}.",

        f"Top product by revenue: {top_product['Description']} "
        f"with revenue of {top_product['Revenue']:,.2f}."
    ],

    "recommendations": [
        f"Focus sales and marketing efforts on the strongest market: "
        f"{top_country['Country']}.",

        f"Review the performance drivers of the top product: "
        f"{top_product['Description']}.",

        "Monitor monthly revenue trends to identify future growth opportunities."
    ]
}

print("\nAI Insights:")
print(ai_response)
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Open the business report workbook
wb = load_workbook(REPORT_FILE)

# Remove old AI Insights sheet if it already exists
if "AI Insights" in wb.sheetnames:
    del wb["AI Insights"]

# Create AI Insights sheet
ws = wb.create_sheet("AI Insights")

# Add business summary
ws["A1"] = "AI Business Insights"
ws["A3"] = "Business Summary"
ws["A4"] = ai_response["summary"]
ws["A4"].alignment = Alignment(wrap_text=True, vertical="top")
ws.row_dimensions[4].height = 45

# Add key findings
ws["A6"] = "Key Findings"

for i, finding in enumerate(ai_response["key_findings"], start=7):
    ws[f"A{i}"] = f"{i-6}. {finding}"

# Add recommendations
ws["A11"] = "Business Recommendations"

for i, recommendation in enumerate(ai_response["recommendations"], start=12):
    ws[f"A{i}"] = f"{i-11}. {recommendation}"

# Adjust column width
ws.column_dimensions["A"].width = 100

# Save workbook
wb.save(REPORT_FILE)

print("\nAI Insights sheet created successfully!")
