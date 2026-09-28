from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter


# ============================================================
# BUSINESS DATA AUTOMATION - DATA CLEANING
# ============================================================

# Project folders
PROJECT_ROOT = Path(__file__).resolve().parent.parent

RAW_FILE = PROJECT_ROOT / "Data" / "Raw Data" / "online_retail_II.xlsx"
OUTPUT_DIR = PROJECT_ROOT / "Data" / "cleaned Data"
OUTPUT_FILE = OUTPUT_DIR / "cleaned_retail_data.xlsx"


# Create output folder if it doesn't exist
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


print("=" * 70)
print("DATA CLEANING STARTED")
print("=" * 70)

print(f"\nInput file:")
print(RAW_FILE)

print(f"\nOutput file:")
print(OUTPUT_FILE)


# ------------------------------------------------------------
# Check input file
# ------------------------------------------------------------

if not RAW_FILE.exists():
    print("\nERROR: Raw Excel file was not found.")
    print("Check the file name and folder location.")
    raise SystemExit


# ------------------------------------------------------------
# Open source workbook in read-only mode
# ------------------------------------------------------------

print("\nLoading raw workbook...")

source_wb = load_workbook(
    RAW_FILE,
    read_only=True,
    data_only=True
)

print("\nSheets found:")

for sheet_name in source_wb.sheetnames:
    print(f" - {sheet_name}")


# ------------------------------------------------------------
# Create output workbook
# ------------------------------------------------------------

output_wb = Workbook(write_only=True)


# ------------------------------------------------------------
# Process each sheet
# ------------------------------------------------------------

for sheet_name in source_wb.sheetnames:

    print("\n" + "=" * 70)
    print(f"CLEANING SHEET: {sheet_name}")
    print("=" * 70)

    source_ws = source_wb[sheet_name]

    # Create matching output sheet
    output_ws = output_wb.create_sheet(title=sheet_name)

    rows_read = 0
    rows_kept = 0

    missing_essential = 0
    cancelled_removed = 0
    negative_quantity_removed = 0
    invalid_price_removed = 0
    duplicate_removed = 0

    # --------------------------------------------------------
    # Read header
    # --------------------------------------------------------

    rows = source_ws.iter_rows()

    try:
        header = next(rows)
    except StopIteration:
        print("Empty sheet. Skipping.")
        continue

    headers = [
        str(cell.value).strip() if cell.value is not None else ""
        for cell in header
    ]

    print("\nColumns:")
    print(tuple(headers))

    # Map column names to positions
    column_map = {
        name.lower(): index
        for index, name in enumerate(headers)
    }

    # Required columns
    required_columns = [
        "invoice",
        "stockcode",
        "description",
        "quantity",
        "invoicedate",
        "price",
        "country"
    ]

    missing_columns = [
        col for col in required_columns
        if col not in column_map
    ]

    if missing_columns:
        print("\nERROR: Required columns missing:")
        print(missing_columns)
        continue

    # Write header to cleaned sheet
    output_ws.append(headers)

    # --------------------------------------------------------
    # Track duplicate rows using fingerprints
    # --------------------------------------------------------

    seen_rows = set()

    # --------------------------------------------------------
    # Process rows
    # --------------------------------------------------------

    for row in rows:

        rows_read += 1

        values = [cell.value for cell in row]

        # Make sure row has same number of values as header
        if len(values) < len(headers):
            values += [None] * (len(headers) - len(values))

        # ----------------------------------------------------
        # Get important fields
        # ----------------------------------------------------

        invoice = values[column_map["invoice"]]
        stockcode = values[column_map["stockcode"]]
        description = values[column_map["description"]]
        quantity = values[column_map["quantity"]]
        invoice_date = values[column_map["invoicedate"]]
        price = values[column_map["price"]]
        country = values[column_map["country"]]

        # ----------------------------------------------------
        # 1. Missing essential values
        # ----------------------------------------------------

        essential_values = [
            invoice,
            stockcode,
            description,
            quantity,
            invoice_date,
            price,
            country
        ]

        if any(value is None for value in essential_values):
            missing_essential += 1
            continue

        # ----------------------------------------------------
        # 2. Remove cancelled invoices
        # ----------------------------------------------------

        if str(invoice).startswith("C"):
            cancelled_removed += 1
            continue

        # ----------------------------------------------------
        # 3. Remove negative quantities
        # ----------------------------------------------------

        if isinstance(quantity, (int, float)):

            if quantity < 0:
                negative_quantity_removed += 1
                continue

        # ----------------------------------------------------
        # 4. Remove zero/negative prices
        # ----------------------------------------------------

        if isinstance(price, (int, float)):

            if price <= 0:
                invalid_price_removed += 1
                continue

        # ----------------------------------------------------
        # 5. Clean text fields
        # ----------------------------------------------------

        if isinstance(description, str):
            description = " ".join(description.split())

        if isinstance(stockcode, str):
            stockcode = stockcode.strip()

        if isinstance(country, str):
            country = " ".join(country.split())

        # Put cleaned values back
        values[column_map["description"]] = description
        values[column_map["stockcode"]] = stockcode
        values[column_map["country"]] = country

        # ----------------------------------------------------
        # 6. Remove exact duplicate rows
        # ----------------------------------------------------

        row_text = "|".join(
            "" if value is None else str(value)
            for value in values
        )

        fingerprint = hash(row_text)

        if fingerprint in seen_rows:
            duplicate_removed += 1
            continue

        seen_rows.add(fingerprint)

        # ----------------------------------------------------
        # 7. Keep cleaned row
        # ----------------------------------------------------

        output_ws.append(values)

        rows_kept += 1

        # Progress message every 100,000 rows
        if rows_read % 100000 == 0:
            print(
                f"Processed: {rows_read:,} | "
                f"Kept: {rows_kept:,}"
            )

    # --------------------------------------------------------
    # Sheet summary
    # --------------------------------------------------------

    print("\nSheet cleaning complete.")

    print(f"Rows read:                 {rows_read:,}")
    print(f"Rows kept:                 {rows_kept:,}")
    print(f"Missing essential data:    {missing_essential:,}")
    print(f"Cancelled invoices:        {cancelled_removed:,}")
    print(f"Negative quantities:       {negative_quantity_removed:,}")
    print(f"Zero/negative prices:      {invalid_price_removed:,}")
    print(f"Duplicate rows removed:    {duplicate_removed:,}")


# ------------------------------------------------------------
# Save cleaned workbook
# ------------------------------------------------------------

print("\n" + "=" * 70)
print("SAVING CLEANED DATA")
print("=" * 70)

output_wb.save(OUTPUT_FILE)

source_wb.close()

print("\nCLEANING COMPLETE")
print("=" * 70)
print(f"Cleaned file created at:")
print(OUTPUT_FILE)
print("=" * 70)
