import pandas as pd

# Load the raw Excel file from the subfolder
df = pd.read_excel("raw_data/raw_data.xlsx")

# Quick look at what you're working with
print(df.shape)          # rows, columns
print(df.head())         # first 5 rows
print(df.dtypes)         # column data types
print(df.isnull().sum()) # missing values per column

import pandas as pd
import numpy as np

# 1. Load data
df = pd.read_excel("raw_data/raw_data.xlsx")
print("Original shape:", df.shape)

# 2. Drop exact duplicates only
df = df.drop_duplicates()

# 3. Strip extra whitespace safely
text_cols = df.select_dtypes(include=["object", "string"]).columns
for col in text_cols:
    df[col] = df[col].astype(str).str.strip()

# 4. Standardize text columns casing
for col in ["Payment_Method", "Order_Status", "Delivery_Status", "Supplier", "Service_Category", "Branch"]:
    if col in df.columns:
        df[col] = df[col].str.title()

# 5. Handle numerical conversions
numeric_cols = ["Quantity", "Unit_Price_NGN", "Total_Amount_NGN", "Stock_Level", "Reorder_Level"]
for col in numeric_cols:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors="coerce")

# 6. Parse dates safely without dropping rows on failure
if "Order_Date" in df.columns:
    df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce", format="mixed")

# 7. Fill missing values instead of dropping rows
df["Quantity"] = df["Quantity"].fillna(0)
df["Unit_Price_NGN"] = df["Unit_Price_NGN"].fillna(0)
df["Total_Amount_NGN"] = df["Total_Amount_NGN"].fillna(0)

# 8. Reset index
df = df.reset_index(drop=True)

# 9. Save cleaned dataset
df.to_excel("raw_data/cleaned_inventory.xlsx", index=False)

print("Data cleaning complete! Final shape:", df.shape)