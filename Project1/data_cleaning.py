print("decode projects")
print("python is working")
import pandas as pd
import openpyxl
print("pandas version:", pd.__version__)
print("openpyxl is working")

import pandas as pd
import numpy as np

file_path = "Dataset for Data Analytics.xlsx"
df = pd.read_excel(file_path)
print(df.head())
print("\n--- DATASET SHAPE ---")
print(df.shape)

print("\n--- COLUMN NAMES ---")
print(df.columns.tolist())

print("\n--- DATA TYPES ---")
print(df.dtypes)

print("\n--- MISSING VALUES ---")
print(df.isnull().sum())

print("\n--- DUPLICATE ROWS ---")
print(df.duplicated().sum())

df["CouponCode"] = df["CouponCode"].fillna("No Coupon")
df["CouponCode"].isnull().sum()

print("\n--- MISSING VALUES ---")
print(df.isnull().sum())

df["Date"] = pd.to_datetime(df["Date"]).dt.strftime("%Y-%m-%d")


# ORDER ID VALIDATION
print("\n--- ORDER ID VALIDATION ---")

print("Missing OrderIDs:",
df["OrderID"].isnull().sum())
print("Duplicate OrderIDs:",
df["OrderID"].duplicated().sum())
print("Unique OrderIDs:",
df["OrderID"].nunique())
print("Total Records:", len(df))

# DATE VALIDATION
print("\n--- DATE VALIDATION ---")
df["Date"] = pd.to_datetime(df["Date"], errors="coerce")

print("Invalid Dates:",
df["Date"].isnull().sum())
print("Earliest Date:",
df["Date"].min())
print("Latest Date:",
df["Date"].max())

#  TEXT/CATEGORICAL VALUE CHECK
categorical_columns = [
    "Product",
    "PaymentMethod",
    "OrderStatus",
    "CouponCode",
    "ReferralSource"
]

for column in categorical_columns:
    print(f"\n--- {column} ---")

print(df[column].value_counts(dropna=False))

# CATEGORICAL VALUE CHECK
columns_to_check = [
    "Product",
    "PaymentMethod",
    "OrderStatus",
    "CouponCode",
]

for column in columns_to_check:
    print(f"\n--- {column} ---")

print(df[column].value_counts(dropna=False))

# STANDARDIZE COUPON CODES
df["CouponCode"] = df["CouponCode"].str.strip().str.upper()

print("\n--- STANDARDIZED COUPON CODES ---")
print(df["CouponCode"].value_counts())

# CHECKING PAYMENT COLUMN
print("\n--- PAYMENT METHOD ---")
print(df["PaymentMethod"].value_counts(dropna=False))

# CHECKING ORDER STATUS
print("\n--- ORDER STATUS ---")
print(df["OrderStatus"].value_counts(dropna=False))

# CHECKING PRODUCT COLUMN
print("\n--- PRODUCT STATUS ---")
print(df["Product"].value_counts(dropna=False))

# NUMERIC DATA CHECK
numeric_columns = [
    "Quantity", 
    "UnitPrice", 
    "ItemsInCart", 
    "TotalPrice"  
]
print("\n--- NUMERIC DATA TYPES ---")
print(df[numeric_columns].dtypes)
print("\n--- NUMERIC SUMMARY ---")
print(df[numeric_columns].describe()
      )

# NUMERIC VALIDITY CHECK
print("\n--- NUMERIC VALIDITY CHECK ---")

print("Quantity <= 0:",
(df["Quantity"] <= 0).sum())
print("UnitPrice <= 0:",
(df["UnitPrice"] <= 0).sum())
print("ItemsInCart: < 0",
(df["ItemsInCart"] < 0).sum())
print("TotalPrice: <= 0",
(df["TotalPrice"] <= 0).sum())

# TOTAL PRICE VALIDATION
df["CalculatedTotal"] = df["Quantity"] * df["UnitPrice"]

df["PriceDifference"] = (df["TotalPrice"] - df["CalculatedTotal"].abs())

print("\n--- TOTAL PRICE VALIDATION ---")
print("Maximum difference:",
df["PriceDifference"].max())
print("Rows with difference:",
(df["PriceDifference"] > 0.01).sum())

# REMOVING TEMPORARY VALIDATION COLUMNS
df.drop(columns=["CalculatedTotal", "PriceDifference"], inplace=True)

print("\nTemporary validation columns removed.")
print("Final columns:", df.columns.tolist())

print("\n--- FINAL DATA QUALITY REPORT ---")

print("Total Records:", len(df))
print("Total Columns:", len(df.columns))
print("Duplicate Rows:", df.duplicated().sum())
print("Duplicate OrderIDs:", df["OrderID"].duplicated().sum())
print("Missing OrderIDs:", df["OrderID"].isnull().sum())
print("Invalid Dates:", df["Date"].isnull().sum())
print("Missing Values:", df.isnull().sum().sum())
print("Invalid Quantity:", (df["Quantity"] <= 0).sum())
print("Invalid UnitPrice:", (df["UnitPrice"] <= 0).sum())
print("Invalid ItemsInCart:", (df["ItemsInCart"] < 0).sum())
print("Invalid TotalPrice:", (df["TotalPrice"] <= 0).sum())

print("\nData cleaning and validation completed successfully.")


# CREATING FINAL DECODELABS EXCEL WORKBOOK
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

# Reloading the original data for the Raw Data sheet
raw_df = pd.read_excel(file_path)

# Creating the Data Quality Report
quality_report = pd.DataFrame({
    "Quality Check": [
        "Total Records",
        "Total Columns",
        "Missing Values",
        "Duplicate Rows",
        "Duplicate OrderIDs",
        "Missing OrderIDs",
        "Invalid Dates",
        "Invalid Quantity",
        "Invalid UnitPrice",
        "Invalid ItemsInCart",
        "Invalid TotalPrice"
    ],
    "Result": [
        len(df),
        len(df.columns),
        df.isnull().sum().sum(),
        df.duplicated().sum(),
        df["OrderID"].duplicated().sum(),
        df["OrderID"].isnull().sum(),
        df["Date"].isnull().sum(),
        (df["Quantity"] <= 0).sum(),
        (df["UnitPrice"] <= 0).sum(),
        (df["ItemsInCart"] < 0).sum(),
        (df["TotalPrice"] <= 0).sum()
    ]
})

# Creating the Change Log
change_log = pd.DataFrame({
    "Field": [
        "CouponCode",
        "CouponCode",
        "OrderID",
        "Date"
    ],
    "Change Made": [
        "Replaced missing values with NO COUPON",
        "Standardized coupon codes to uppercase",
        "Checked for duplicate IDs",
        "Validated date values"
    ],
    "Result": [
        "309 missing values handled",
        "All coupon codes standardized",
        "0 duplicate OrderIDs found",
        "0 invalid dates found"
    ]
})

# Final output filename
output_file = "DecodeLabs_Project1_Final.xlsx"

# Export all sheets
with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    raw_df.to_excel(writer, sheet_name="Raw Data", index=False)
    df.to_excel(writer, sheet_name="Cleaned Data", index=False)
    quality_report.to_excel(writer, sheet_name="Data Quality Report", index=False)
    change_log.to_excel(writer, sheet_name="Change Log", index=False)

# Formating the workbook
wb = load_workbook(output_file)

for ws in wb.worksheets:

    # Bold header row
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center")

    # Freezing top row
    ws.freeze_panes = "A2"

    # Adjusting column widths
    for column_cells in ws.columns:
        max_length = 0
        column_letter = get_column_letter(column_cells[0].column)

        for cell in column_cells:
            if cell.value is not None:
                max_length = max(max_length, len(str(cell.value)))

        ws.column_dimensions[column_letter].width = min(max_length + 2, 35)

# Formating Cleaned Data
ws = wb["Cleaned Data"]

# Date column
date_column = None

for cell in ws[1]:
    if cell.value == "Date":
        date_column = cell.column
        break

if date_column:
    for row in range(2, ws.max_row + 1):
        ws.cell(row, date_column).number_format = "yyyy-mm-dd"

# Numeric formatting
for column_name in ["UnitPrice", "TotalPrice"]:
    for cell in ws[1]:
        if cell.value == column_name:
            column = cell.column
            for row in range(2, ws.max_row + 1):
                ws.cell(row, column).number_format = "0.00"

# Saving workbook
wb.save(output_file)

print("\n==========================================")
print("FINAL WORKBOOK CREATED SUCCESSFULLY")
print("==========================================")
print("File:", output_file)
print("Sheets:")
print("1. Raw Data")
print("2. Cleaned Data")
print("3. Data Quality Report")
print("4. Change Log")






