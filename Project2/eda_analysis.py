import pandas as pd

# Loading Project 2 dataset
file_path = r"C:\Users\admin\Desktop\decode projects\project2\Dataset for Data Analytics (4).xlsx"
df = pd.read_excel(file_path)


# BASIC INFO OF THE DATASET
print("DATASET SHAPE")
print(df.shape)

print("\nCOLUMN NAMES")
print(df.columns.tolist())

print("\nDATA TYPES")
print(df.dtypes)

print("\nFIRST 5 ROWS")
print(df.head())

print("\nMISSING VALUES")
print(df.isnull().sum())

print("\nDUPLICATE ROWS")
print(df.duplicated().sum())

# DESCRIPTIVE STATISTICS
print("\nDESCRIPTIVE STATISTICS")

numeric_columns = df.select_dtypes(include="number").columns

print("\nCOUNT")
print(df[numeric_columns].count())

print("\nMEAN")
print(df[numeric_columns].mean())

print("\nMEDIAN")
print(df[numeric_columns].median())

print("\nMINIMUM")
print(df[numeric_columns].min())

print("\nMAXIMUM")
print(df[numeric_columns].max())


# DISTRIBUTION SUMMARY
print("\nDISTRIBUTION SUMMARY")

for column in numeric_columns:
    print(f"\n{column}")
    print(df[column].describe())

# OUTLIER DETECTION USING IQR
print("\nOUTLIER DETECTION")

for column in numeric_columns:
    Q1 = df[column].quantile(0.25)
    Q3 = df[column].quantile(0.75)
    IQR = Q3 - Q1

    lower_bound = Q1 - 1.5 * IQR
    upper_bound = Q3 + 1.5 * IQR

    outliers = df[(df[column] < lower_bound) | (df[column] > upper_bound)]

    print(f"\n{column}")
    print(f"Q1: {Q1:.2f}")
    print(f"Q3: {Q3:.2f}")
    print(f"IQR: {IQR:.2f}")
    print(f"Lower Bound: {lower_bound:.2f}")
    print(f"Upper Bound: {upper_bound:.2f}")
    print(f"Number of Outliers: {len(outliers)}")

# VIEWING THE TOTAL PRICE OUTLIERS
print("\nTOTAL PRICE OUTLIERS")

Q1 = df["TotalPrice"].quantile(0.25)
Q3 = df["TotalPrice"].quantile(0.75)
IQR = Q3 - Q1

upper_bound = Q3 + 1.5 * IQR

totalprice_outliers = df[df["TotalPrice"] > upper_bound]

print(totalprice_outliers[
    ["OrderID", "Date", "Product", "Quantity", "UnitPrice", "ItemsInCart", "TotalPrice"]
])

# PRODUCT ANALYSIS
print("\nPRODUCT ANALYSIS")

product_summary = df.groupby("Product").agg(
    Total_Orders=("OrderID", "count"),
    Total_Quantity=("Quantity", "sum"),
    Total_Revenue=("TotalPrice", "sum"),
    Average_Order_Value=("TotalPrice", "mean")
).sort_values("Total_Revenue", ascending=False)

print(product_summary)

# YEARLY SALES TREND
print("\nYEARLY SALES TREND")

df["Year"] = df["Date"].dt.year

yearly_sales = df.groupby("Year").agg(
    Total_Orders=("OrderID", "count"),
    Total_Revenue=("TotalPrice", "sum"),
    Average_Order_Value=("TotalPrice", "mean")
)

print(yearly_sales)

# MONTHLY SALES TREND
print("\nMONTHLY SALES TREND")

df["Month"] = df["Date"].dt.to_period("M")

monthly_sales = df.groupby("Month").agg(
    Total_Orders=("OrderID", "count"),
    Total_Revenue=("TotalPrice", "sum"),
    Average_Order_Value=("TotalPrice", "mean")
)

print(monthly_sales)

# PRODUCT PERFORMANCE ANALYSIS
print("\nPRODUCT PERFORMANCE ANALYSIS")

product_performance = df.groupby("Product").agg(
    Total_Orders=("OrderID", "count"),
    Total_Quantity=("Quantity", "sum"),
    Total_Revenue=("TotalPrice", "sum"),
    Average_Order_Value=("TotalPrice", "mean")
).sort_values("Total_Revenue", ascending=False)

print(product_performance)

# LOOKING AT THE CORRELATION ANALYSIS
print("\nCORRELATION ANALYSIS")

correlation_matrix = df[numeric_columns].corr()

print(correlation_matrix.round(2))


# CHART 1: MONTHLY REVENUE TREND

import matplotlib.pyplot as plt

df["YearMonth"] = df["Date"].dt.to_period("M")

monthly_revenue = df.groupby("YearMonth")["TotalPrice"].sum()

plt.figure(figsize=(12, 6))

plt.plot(
    monthly_revenue.index.astype(str),
    monthly_revenue.values,
    marker="o"
)

plt.title("Monthly Revenue Trend")
plt.xlabel("Month")
plt.ylabel("Total Revenue")

plt.xticks(rotation=45)
plt.tight_layout()

plt.show()

# CHART 2: REVENUE BY PRODUCT

import matplotlib.pyplot as plt

product_revenue = (
    df.groupby("Product")["TotalPrice"]
    .sum()
    .sort_values(ascending=True)
)

plt.figure(figsize=(10, 6))

plt.barh(product_revenue.index, product_revenue.values)

plt.title("Total Revenue by Product")
plt.xlabel("Total Revenue")
plt.ylabel("Product")

plt.tight_layout()

plt.savefig("revenue_by_product.png", dpi=300, bbox_inches="tight")

plt.show()
plt.close()


# CHART 3: TOTAL PRICE DISTRIBUTION

plt.figure(figsize=(10, 6))

plt.hist(df["TotalPrice"], bins=30)

plt.title("Distribution of Total Order Price")
plt.xlabel("Total Price")
plt.ylabel("Number of Orders")

plt.tight_layout()

plt.savefig("total_price_distribution.png", dpi=300, bbox_inches="tight")

plt.show()
plt.close()

# CHART 4: CORRELATION HEATMAP

import seaborn as sns
import matplotlib.pyplot as plt

correlation_matrix = df[numeric_columns].corr()

plt.figure(figsize=(8, 6))

sns.heatmap(
    correlation_matrix,
    annot=True,
    cmap="Blues",
    fmt=".2f",
    linewidths=0.5
)

plt.title("Correlation Matrix of Numerical Variables")
plt.tight_layout()

plt.savefig("correlation_heatmap.png", dpi=300, bbox_inches="tight")

plt.show()
plt.close()