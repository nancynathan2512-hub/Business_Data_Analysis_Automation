from openpyxl import load_workbook
from pathlib import Path

# Project paths
project_root = Path(__file__).resolve().parent.parent
input_file = project_root/"Data/Raw Data/online_retail_II.xlsx/online_retail_II.xlsx"

print("=" * 60)
print("LARGE SCALE DATA INSPECTION")
print("=" * 60)

print(f"\nFile: {input_file}")
print("Loading workbook in read-only mode...")

# Read without modifying the original file
wb = load_workbook(
    filename=input_file,
    read_only=True,
    data_only=True
)

print("\nSheets found:")
for sheet in wb.sheetnames:
    print(f"  - {sheet}")

# Inspect every sheet
for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]

    print("\n" + "-" * 60)
    print(f"SHEET: {sheet_name}")
    print("-" * 60)

    print(f"Rows reported by Excel: {ws.max_row:,}")
    print(f"Columns reported by Excel: {ws.max_column:,}")

    # Read first 5 rows only
    print("\nFirst 5 rows:")

    for row_number, row in enumerate(
        ws.iter_rows(values_only=True),
        start=1
    ):
        print(row)

        if row_number == 5:
            break

print("\n" + "=" * 60)
print("INSPECTION COMPLETE")
print("=" * 60)

import hashlib

print("\n" + "=" * 60)
print("DATA QUALITY CHECK")
print("=" * 60)

for sheet_name in wb.sheetnames:

    ws = wb[sheet_name]

    print(f"\nChecking sheet: {sheet_name}")

    # Counters
    total_rows = 0
    missing_values = 0
    negative_quantity = 0
    invalid_price = 0
    cancelled_invoices = 0
    duplicate_rows = 0

    # Store fingerprints of rows to detect exact duplicates
    seen_rows = set()

    # Read header
    rows = ws.iter_rows(values_only=True)
    headers = next(rows)

    print(f"Columns: {headers}")

    for row in rows:

        total_rows += 1

        # Missing values
        missing_values += sum(
            1 for value in row if value is None
        )

        # Exact duplicate detection
        row_text = repr(row).encode("utf-8")
        fingerprint = hashlib.md5(row_text).hexdigest()

        if fingerprint in seen_rows:
            duplicate_rows += 1
        else:
            seen_rows.add(fingerprint)

        # Quantity check
        quantity = row[3]

        if isinstance(quantity, (int, float)):
            if quantity < 0:
                negative_quantity += 1

        # Price check
        price = row[5]

        if isinstance(price, (int, float)):
            if price <= 0:
                invalid_price += 1

        # Invoice check
        invoice = row[0]

        if invoice is not None:
            if str(invoice).startswith("C"):
                cancelled_invoices += 1

    print(f"Rows scanned: {total_rows:,}")
    print(f"Missing values: {missing_values:,}")
    print(f"Duplicate rows: {duplicate_rows:,}")
    print(f"Negative quantities: {negative_quantity:,}")
    print(f"Zero/negative prices: {invalid_price:,}")
    print(f"Cancelled invoices: {cancelled_invoices:,}")

print("\n" + "=" * 60)
print("DATA QUALITY CHECK COMPLETE")
print("=" * 60)
