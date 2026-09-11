# Data Cleaning Project - DecodeLabs Internship

## 📋 Project Overview

This project performs comprehensive data cleaning and validation on a dataset containing e-commerce order information. The script processes raw data, identifies and fixes data quality issues, and generates a detailed report with a cleaned dataset and validation summary.

## 📁 Files Included

- **`data_cleaning.py`** - Main Python script for data cleaning and validation
- **`Dataset for Data Analytics.xlsx`** - Raw input data file
- **`DecodeLabs_Project1_Final.xlsx`** - Output Excel workbook with cleaned data and reports

## 🎯 Features

### Data Validation Checks
✅ **Order ID Validation**
- Check for missing OrderIDs
- Identify duplicate OrderIDs
- Count unique OrderIDs

✅ **Date Validation**
- Convert dates to standard format (YYYY-MM-DD)
- Detect and report invalid dates
- Show date range (earliest to latest)

✅ **Categorical Data Validation**
- Validate values in: Product, PaymentMethod, OrderStatus, CouponCode, ReferralSource
- Identify value distributions
- Standardize text (uppercase, trim whitespace)

✅ **Numeric Data Validation**
- Validate Quantity, UnitPrice, ItemsInCart, TotalPrice
- Check for negative or zero values where inappropriate
- Verify total price calculations (Quantity × UnitPrice)

✅ **Duplicate Detection**
- Identify duplicate rows in dataset
- Remove redundant records

### Data Cleaning Operations
- **Fill missing values**: CouponCode nulls replaced with "No Coupon"
- **Standardize text**: Coupon codes converted to uppercase
- **Format dates**: Standardized to YYYY-MM-DD format
- **Validate calculations**: Check price accuracy
- **Remove temporary columns**: Clean up validation helpers

## 📊 Output Excel Workbook

The script generates **`DecodeLabs_Project1_Final.xlsx`** with 4 sheets:

### 1. **Raw Data**
   - Original unmodified dataset for reference

### 2. **Cleaned Data**
   - Processed dataset with all corrections applied
   - Proper formatting (dates, currency, numbers)
   - Frozen header row for easy navigation

### 3. **Data Quality Report**
   - Summary of quality checks performed:
     - Total records and columns count
     - Missing values count
     - Duplicate rows and OrderIDs
     - Invalid entries by field
   
### 4. **Change Log**
   - Detailed record of all transformations made
   - Changes applied to: CouponCode, OrderID, Date
   - Results and impact of each change

## 🛠️ Requirements

### Python Libraries
```
pandas
numpy
openpyxl
```

### Installation
```bash
pip install pandas numpy openpyxl
```

## 🚀 How to Run

1. Ensure the input file `Dataset for Data Analytics.xlsx` is in the same directory as the script
2. Run the script:
   ```bash
   python data_cleaning.py
   ```
3. The script will:
   - Load and analyze the raw data
   - Perform validation checks (printed to console)
   - Clean and standardize the data
   - Generate the final Excel workbook

## 📈 Data Quality Metrics

The script provides a comprehensive quality report including:
- **Total Records**: Row count in dataset
- **Total Columns**: Number of fields
- **Missing Values**: Total null entries
- **Duplicate Rows**: Identical record count
- **Duplicate OrderIDs**: Duplicate order identifiers
- **Invalid Dates**: Malformed or unreachable date values
- **Invalid Quantities**: Non-positive quantity values
- **Invalid Prices**: Non-positive price values
- **Invalid Cart Items**: Negative item counts

## 📝 Column Details

### Expected Columns
- `OrderID` - Unique order identifier
- `Date` - Order date
- `Product` - Product name
- `Quantity` - Number of units ordered
- `UnitPrice` - Price per unit
- `TotalPrice` - Total order price
- `PaymentMethod` - Payment type
- `OrderStatus` - Order status
- `CouponCode` - Applied coupon (if any)
- `ReferralSource` - Source of order
- `ItemsInCart` - Items in shopping cart

## ✨ Key Improvements Made

| Issue | Solution |
|-------|----------|
| Missing CouponCode values | Filled with "No Coupon" |
| Inconsistent coupon formatting | Standardized to uppercase |
| Date format inconsistencies | Converted to YYYY-MM-DD |
| Unvalidated numeric data | Added range and calculation checks |
| Unclear data quality | Generated comprehensive QA report |

## 📌 Notes

- All validations are logged in the console during execution
- Temporary validation columns are removed from final output
- Excel formatting includes bold headers, frozen panes, and auto-adjusted column widths
- The change log provides full traceability of all modifications

## 🎓 Purpose

This project demonstrates best practices for:
- Data validation and cleaning workflows
- Error identification and handling
- Data quality reporting
- Professional data documentation
- Excel workbook creation with pandas and openpyxl

---

*DecodeLabs Internship Project*
