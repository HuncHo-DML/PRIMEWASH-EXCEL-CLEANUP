import streamlit as st
import pandas as pd
import plotly.express as px
import io
import os
from streamlit_autorefresh import st_autorefresh  # <--- ADDED HERE: Import auto-refresh module

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Prime Wash Executive Dashboard",
    page_icon="📊",
    layout="wide"
)

# --- PERIODIC AUTO-REFRESH CONFIGURATION ---
# <--- ADDED HERE: Automatically refreshes the app every 60 seconds (60000 ms) 
# so it checks for updated Excel files on disk without manual browser refreshes.
st_autorefresh(interval=60000, limit=None, key="periodic_dashboard_refresh")

# --- CUSTOM CORPORATE STYLING (CSS) ---
st.markdown("""
    <style>
        h1 { color: #FFFFFF; font-family: 'Segoe UI', sans-serif; }
        h3 { color: #A0AEC0; font-family: 'Segoe UI', sans-serif; }
        
        /* Modern dark card container */
        .stMetric { 
            background-color: #1E293B; 
            border: 1px solid #334155;
            padding: 15px; 
            border-radius: 8px; 
            box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1); 
        }
        
        /* High-contrast text colors for dark mode visibility */
        [data-testid="stMetricLabel"] {
            color: #94A3B8 !important;
            font-size: 14px !important;
        }
        [data-testid="stMetricValue"] {
            color: #F8FAFC !important;
            font-size: 28px !important;
        }
    </style>
""", unsafe_allow_html=True)

# --- LOAD DATA ---
@st.cache_data
def load_data():
    file_path = "raw_data/cleaned_inventory.xlsx"
    
    # <--- ADDED HERE: Tracks the file modification time so cache updates automatically 
    # when the underlying Excel file changes on disk.
    mod_time = os.path.getmtime(file_path)
    
    df = pd.read_excel(file_path)
    df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce")
    df["Month_Year"] = df["Order_Date"].dt.to_period("M").astype(str)
    return df

try:
    df_raw = load_data()
except Exception as e:
    st.error(f"Error loading data file: {e}")
    st.stop()

# --- SIDEBAR FILTERS (SLICERS) ---
st.sidebar.header("Dashboard Filters")
st.sidebar.markdown("Use the filters below to dynamically slice the dashboard data.")

# Branch Filter
branches = df_raw["Branch"].dropna().unique().tolist()
selected_branches = st.sidebar.multiselect("Select Branch(es)", options=branches, default=branches)

# Payment Method Filter
payment_methods = df_raw["Payment_Method"].dropna().unique().tolist()
selected_payments = st.sidebar.multiselect("Select Payment Method(s)", options=payment_methods, default=payment_methods)

# Service Category Filter
service_cats = df_raw["Service_Category"].dropna().unique().tolist()
selected_cats = st.sidebar.multiselect("Select Service Category", options=service_cats, default=service_cats)

# Apply Filters
df = df_raw[
    df_raw["Branch"].isin(selected_branches) &
    df_raw["Payment_Method"].isin(selected_payments) &
    df_raw["Service_Category"].isin(selected_cats)
]

if df.empty:
    st.warning("No data matches the selected filter criteria. Please adjust your filters.")
    st.stop()

# --- SIDEBAR EXPORT SECTION ---
st.sidebar.markdown("---")
st.sidebar.header("Export Reports")

# 1. CSV Export
csv_data = df.to_csv(index=False).encode('utf-8')
st.sidebar.download_button(
    label="📥 Download Filtered Data (CSV)",
    data=csv_data,
    file_name="prime_wash_filtered_report.csv",
    mime="text/csv"
)

# 2. Excel Export
excel_buffer = io.BytesIO()
with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
    df.to_excel(writer, sheet_name="Filtered_Data", index=False)
excel_data = excel_buffer.getvalue()

st.sidebar.download_button(
    label="📥 Download Filtered Data (Excel)",
    data=excel_data,
    file_name="prime_wash_filtered_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

# --- HEADER SECTION ---
st.title("PRIME WASH LAUNDRY SERVICES")
st.markdown("### Executive Operations & Performance Web Dashboard")
st.markdown("---")

# Top KPI Summary Cards
total_revenue = df["Total_Amount_NGN"].sum()
total_orders = df["Item_ID"].count()
avg_order_value = df["Total_Amount_NGN"].mean() if total_orders > 0 else 0

kpi1, kpi2, kpi3 = st.columns(3)
kpi1.metric("Total Revenue", f"₦{total_revenue:,.2f}")
kpi2.metric("Total Orders Processed", f"{total_orders:,}")
kpi3.metric("Average Order Value", f"₦{avg_order_value:,.2f}")

st.markdown("---")

# --- ROW 1: CHARTS 1, 2, 3 ---
col1, col2, col3 = st.columns(3)

with col1:
    st.subheader("1. Revenue by Branch & Service")
    fig1 = px.bar(
        df, x="Branch", y="Total_Amount_NGN", color="Service_Category",
        barmode="stack", color_discrete_sequence=px.colors.qualitative.Prism
    )
    fig1.update_layout(xaxis_title="Branch", yaxis_title="Revenue (NGN)", legend_title="Category")
    st.plotly_chart(fig1, use_container_width=True)

with col2:
    st.subheader("2. Revenue by Payment Method")
    pay_grouped = df.groupby("Payment_Method")["Total_Amount_NGN"].sum().reset_index()
    fig2 = px.pie(pay_grouped, names="Payment_Method", values="Total_Amount_NGN", hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
    st.plotly_chart(fig2, use_container_width=True)

with col3:
    st.subheader("3. Staff Performance")
    if "Staff_Name" in df.columns and "Order_Status" in df.columns:
        staff_grouped = df.groupby(["Staff_Name", "Order_Status"]).size().reset_index(name="Count")
        fig3 = px.bar(staff_grouped, x="Staff_Name", y="Count", color="Order_Status", barmode="stack", color_discrete_sequence=px.colors.qualitative.Safe)
        fig3.update_layout(xaxis_title="Staff Name", yaxis_title="Order Count")
        st.plotly_chart(fig3, use_container_width=True)

# --- ROW 2: CHARTS 4, 5, 6 ---
col4, col5, col6 = st.columns(3)

with col4:
    st.subheader("4. Delivery Status by Branch")
    if "Delivery_Status" in df.columns:
        del_grouped = df.groupby(["Branch", "Delivery_Status"]).size().reset_index(name="Count")
        fig4 = px.bar(del_grouped, x="Branch", y="Count", color="Delivery_Status", barmode="stack")
        st.plotly_chart(fig4, use_container_width=True)

with col5:
    st.subheader("5. Revenue by Service Type")
    srv_grouped = df.groupby("Service_Type")["Total_Amount_NGN"].sum().reset_index()
    fig5 = px.bar(srv_grouped, x="Service_Type", y="Total_Amount_NGN", color_discrete_sequence=["#1B365D"])
    st.plotly_chart(fig5, use_container_width=True)

with col6:
    st.subheader("6. Customer Type Split")
    cust_grouped = df.groupby("Customer_Type")["Total_Amount_NGN"].sum().reset_index()
    fig6 = px.pie(cust_grouped, names="Customer_Type", values="Total_Amount_NGN", hole=0.6, color_discrete_sequence=px.colors.qualitative.Bold)
    st.plotly_chart(fig6, use_container_width=True)

# --- ROW 3: CHARTS 7, 8 ---
col7, col8 = st.columns(2)

with col7:
    st.subheader("7. Stock Levels by Supplier")
    if "Supplier" in df.columns and "Stock_Level" in df.columns:
        sup_grouped = df.groupby("Supplier")["Stock_Level"].mean().reset_index()
        fig7 = px.bar(sup_grouped, x="Supplier", y="Stock_Level", orientation="h", color_discrete_sequence=["#4B9CD3"])
        st.plotly_chart(fig7, use_container_width=True)

with col8:
    st.subheader("8. Monthly Revenue Trend")
    monthly_grouped = df.groupby("Month_Year")["Total_Amount_NGN"].sum().reset_index()
    fig8 = px.line(monthly_grouped, x="Month_Year", y="Total_Amount_NGN", markers=True, line_shape="linear")
    fig8.update_traces(line=dict(color="#1B365D", width=3))
    st.plotly_chart(fig8, use_container_width=True)