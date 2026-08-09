import pandas as pd

# 1. Load the cleaned inventory dataset
df = pd.read_excel("raw_data/cleaned_inventory.xlsx")

# Ensure Order_Date is datetime for time-based groupings
if "Order_Date" in df.columns:
    df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce")
    df["Month_Year"] = df["Order_Date"].dt.to_period("M").astype(str)

print("generating comprehensive pivot tables...\n")

# --- 1. Revenue by Branch and Service Category ---
branch_service_pivot = pd.pivot_table(
    df,
    values="Total_Amount_NGN",
    index="Branch",
    columns="Service_Category",
    aggfunc="sum",
    fill_value=0
)

# --- 2. Payment Method Breakdown ---
payment_pivot = pd.pivot_table(
    df,
    values=["Quantity", "Total_Amount_NGN"],
    index="Payment_Method",
    aggfunc={"Quantity": "sum", "Total_Amount_NGN": ["sum", "mean"]}
)

# --- 3. Staff Performance (Order Status Breakdown) ---
staff_pivot = pd.pivot_table(
    df,
    values="Item_ID",
    index="Staff_Name",
    columns="Order_Status",
    aggfunc="count",
    fill_value=0
)

# --- 4. Delivery Status by Branch ---
delivery_pivot = pd.pivot_table(
    df,
    values="Item_ID",
    index="Branch",
    columns="Delivery_Status",
    aggfunc="count",
    fill_value=0
)

# --- 5. Service Type Volume and Revenue ---
service_type_pivot = pd.pivot_table(
    df,
    values=["Quantity", "Total_Amount_NGN"],
    index="Service_Type",
    aggfunc={"Quantity": "sum", "Total_Amount_NGN": "sum"}
)

# --- 6. Customer Type Analysis (Corporate vs Individual) ---
customer_type_pivot = pd.pivot_table(
    df,
    values=["Quantity", "Total_Amount_NGN"],
    index="Customer_Type",
    aggfunc={"Quantity": "sum", "Total_Amount_NGN": ["sum", "mean"]}
)

# --- 7. Stock Levels & Reorder Metrics by Supplier ---
supplier_pivot = pd.pivot_table(
    df,
    values=["Stock_Level", "Reorder_Level"],
    index="Supplier",
    aggfunc={"Stock_Level": "mean", "Reorder_Level": "mean"}
)

# --- 8. Monthly Revenue Trend ---
monthly_trend_pivot = pd.pivot_table(
    df,
    values="Total_Amount_NGN",
    index="Month_Year",
    columns="Branch",
    aggfunc="sum",
    fill_value=0
)

# --- Export All Pivot Tables to a Multi-Sheet Excel Workbook ---
output_path = "raw_data/inventory_pivot_summary.xlsx"
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    branch_service_pivot.to_excel(writer, sheet_name="Branch_Service_Revenue")
    payment_pivot.to_excel(writer, sheet_name="Payment_Summary")
    staff_pivot.to_excel(writer, sheet_name="Staff_Performance")
    delivery_pivot.to_excel(writer, sheet_name="Delivery_Status")
    service_type_pivot.to_excel(writer, sheet_name="Service_Type_Metrics")
    customer_type_pivot.to_excel(writer, sheet_name="Customer_Type_Analysis")
    supplier_pivot.to_excel(writer, sheet_name="Supplier_Stock_Levels")
    monthly_trend_pivot.to_excel(writer, sheet_name="Monthly_Revenue_Trend")

print(f"All 8 pivot tables have been successfully generated and saved to: {output_path}")
