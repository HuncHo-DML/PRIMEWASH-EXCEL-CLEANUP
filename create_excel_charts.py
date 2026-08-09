import pandas as pd

# 1. Load the cleaned dataset
df = pd.read_excel("raw_data/cleaned_inventory.xlsx")
df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce")
df["Month_Year"] = df["Order_Date"].dt.to_period("M").astype(str)

print("Generating polished dark-mode Excel dashboard...")

# 2. Build the 8 Data Summaries (Backends)
branch_service = pd.pivot_table(df, values="Total_Amount_NGN", index="Branch", columns="Service_Category", aggfunc="sum", fill_value=0).reset_index()
payment_df = pd.pivot_table(df, values="Total_Amount_NGN", index="Payment_Method", aggfunc="sum").reset_index()
staff_df = pd.pivot_table(df, values="Item_ID", index="Staff_Name", columns="Order_Status", aggfunc="count", fill_value=0).reset_index()
delivery_df = pd.pivot_table(df, values="Item_ID", index="Branch", columns="Delivery_Status", aggfunc="count", fill_value=0).reset_index()
service_type_df = pd.pivot_table(df, values="Total_Amount_NGN", index="Service_Type", aggfunc="sum").reset_index()
customer_df = pd.pivot_table(df, values="Total_Amount_NGN", index="Customer_Type", aggfunc="sum").reset_index()
supplier_df = pd.pivot_table(df, values="Stock_Level", index="Supplier", aggfunc="mean").reset_index()
monthly_df = pd.pivot_table(df, values="Total_Amount_NGN", index="Month_Year", aggfunc="sum").reset_index()

# 3. Write using xlsxwriter with professional dark theme styling
output_path = "raw_data/Prime_Wash_Executive_Dashboard.xlsx"
with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
    workbook = writer.book
    
    # --- FIXED: DARK-MODE TITLE FORMATS WITH MATCHING BACKGROUND (#0E1117) ---
    title_format = workbook.add_format({
        'bold': True,
        'font_size': 18,
        'font_color': '#FFFFFF',
        'font_name': 'Segoe UI',
        'bg_color': '#0E1117'
    })
    
    subtitle_format = workbook.add_format({
        'italic': True,
        'font_size': 11,
        'font_color': '#94A3B8',
        'font_name': 'Segoe UI',
        'bg_color': '#0E1117'
    })

    # Write backends to data sheet
    branch_service.to_excel(writer, sheet_name="Backend_Data", index=False, startrow=0, startcol=0)
    payment_df.to_excel(writer, sheet_name="Backend_Data", index=False, startrow=0, startcol=10)
    staff_df.to_excel(writer, sheet_name="Backend_Data", index=False, startrow=0, startcol=15)
    delivery_df.to_excel(writer, sheet_name="Backend_Data", index=False, startrow=0, startcol=22)
    service_type_df.to_excel(writer, sheet_name="Backend_Data", index=False, startrow=0, startcol=29)
    customer_df.to_excel(writer, sheet_name="Backend_Data", index=False, startrow=0, startcol=33)
    supplier_df.to_excel(writer, sheet_name="Backend_Data", index=False, startrow=0, startcol=37)
    monthly_df.to_excel(writer, sheet_name="Backend_Data", index=False, startrow=0, startcol=41)

    # --- DASHBOARD TAB SETUP ---
    dashboard = workbook.add_worksheet("Dashboard")
    dashboard.hide_gridlines(2)

    # Set background color to match web dark mode (#0E1117)
    dark_bg_format = workbook.add_format({'bg_color': '#0E1117'})
    dashboard.set_column('A:AZ', 15, dark_bg_format)

    # Corporate Title Block (using the new background-matched formats)
    dashboard.write("B2", "PRIME WASH LAUNDRY SERVICES", title_format)
    dashboard.write("B3", "Executive Operations & Performance Report", subtitle_format)
    dashboard.set_row(1, 25)
    dashboard.set_row(2, 18)

    # Reusable function to style charts with dark containers matching the web UI
    def apply_dark_theme_to_chart(chart, title_text):
        chart.set_title({
            'name': title_text,
            'name_font': {'color': '#FFFFFF', 'name': 'Segoe UI', 'size': 11, 'bold': True}
        })
        chart.set_chartarea({
            'fill': {'color': '#1E293B'},
            'border': {'color': '#334155', 'width': 1}
        })
        chart.set_plotarea({
            'fill': {'color': '#1E293B'}
        })
        chart.set_x_axis({
            'name_font': {'color': '#94A3B8', 'size': 9},
            'num_font': {'color': '#94A3B8', 'size': 8},
            'line': {'color': '#334155'}
        })
        chart.set_y_axis({
            'name_font': {'color': '#94A3B8', 'size': 9},
            'num_font': {'color': '#94A3B8', 'size': 8},
            'line': {'color': '#334155'},
            'major_gridlines': {'visible': True, 'line': {'color': '#334155', 'dash_type': 'dash'}}
        })
        chart.set_legend({
            'font': {'color': '#94A3B8', 'size': 8},
            'position': 'right'
        })

    # --- CHART 1: Revenue by Branch & Service Category (Stacked Column) ---
    c1 = workbook.add_chart({"type": "column", "subtype": "stacked"})
    for i, col in enumerate(branch_service.columns[1:], start=1):
        c1.add_series({
            "name": [ "Backend_Data", 0, i ],
            "categories": [ "Backend_Data", 1, 0, len(branch_service), 0 ],
            "values": [ "Backend_Data", 1, i, len(branch_service), i ],
        })
    apply_dark_theme_to_chart(c1, "Revenue by Branch & Service")
    dashboard.insert_chart("B5", c1, {"x_scale": 1.3, "y_scale": 1.3})

    # --- CHART 2: Revenue Share by Payment Method (Pie) ---
    c2 = workbook.add_chart({"type": "pie"})
    c2.add_series({
        "name": "Payment Methods",
        "categories": [ "Backend_Data", 1, 10, len(payment_df), 10 ],
        "values": [ "Backend_Data", 1, 11, len(payment_df), 11 ],
    })
    apply_dark_theme_to_chart(c2, "Revenue by Payment Method")
    dashboard.insert_chart("J5", c2, {"x_scale": 1.3, "y_scale": 1.3})

    # --- CHART 3: Staff Performance by Order Status (Stacked Bar) ---
    c3 = workbook.add_chart({"type": "bar", "subtype": "stacked"})
    for i, col in enumerate(staff_df.columns[1:], start=1):
        c3.add_series({
            "name": [ "Backend_Data", 0, i+15 ],
            "categories": [ "Backend_Data", 1, 15, len(staff_df), 15 ],
            "values": [ "Backend_Data", 1, i+15, len(staff_df), i+15 ],
        })
    apply_dark_theme_to_chart(c3, "Staff Performance")
    dashboard.insert_chart("R5", c3, {"x_scale": 1.3, "y_scale": 1.3})

    # --- CHART 4: Delivery Status by Branch (Column) ---
    c4 = workbook.add_chart({"type": "column", "subtype": "stacked"})
    for i, col in enumerate(delivery_df.columns[1:], start=1):
        c4.add_series({
            "name": [ "Backend_Data", 0, i+22 ],
            "categories": [ "Backend_Data", 1, 22, len(delivery_df), 22 ],
            "values": [ "Backend_Data", 1, i+22, len(delivery_df), i+22 ],
        })
    apply_dark_theme_to_chart(c4, "Delivery Status by Branch")
    dashboard.insert_chart("B24", c4, {"x_scale": 1.3, "y_scale": 1.3})

    # --- CHART 5: Service Type Revenue (Column) ---
    c5 = workbook.add_chart({"type": "column"})
    c5.add_series({
        "name": "Revenue",
        "categories": [ "Backend_Data", 1, 29, len(service_type_df), 29 ],
        "values": [ "Backend_Data", 1, 30, len(service_type_df), 30 ],
        "fill": {"color": "#4B9CD3"}
    })
    apply_dark_theme_to_chart(c5, "Revenue by Service Type")
    dashboard.insert_chart("J24", c5, {"x_scale": 1.3, "y_scale": 1.3})

    # --- CHART 6: Customer Type Analysis (Doughnut) ---
    c6 = workbook.add_chart({"type": "doughnut"})
    c6.add_series({
        "name": "Customer Type",
        "categories": [ "Backend_Data", 1, 33, len(customer_df), 33 ],
        "values": [ "Backend_Data", 1, 34, len(customer_df), 34 ],
    })
    apply_dark_theme_to_chart(c6, "Customer Type Split")
    dashboard.insert_chart("R24", c6, {"x_scale": 1.3, "y_scale": 1.3})

    # --- CHART 7: Average Stock Levels by Supplier (Bar) ---
    c7 = workbook.add_chart({"type": "bar"})
    c7.add_series({
        "name": "Avg Stock Level",
        "categories": [ "Backend_Data", 1, 37, len(supplier_df), 37 ],
        "values": [ "Backend_Data", 1, 38, len(supplier_df), 38 ],
        "fill": {"color": "#4B9CD3"}
    })
    apply_dark_theme_to_chart(c7, "Stock Levels by Supplier")
    dashboard.insert_chart("B43", c7, {"x_scale": 1.3, "y_scale": 1.3})

    # --- CHART 8: Monthly Revenue Trend (Line) ---
    c8 = workbook.add_chart({"type": "line"})
    c8.add_series({
        "name": "Monthly Revenue",
        "categories": [ "Backend_Data", 1, 41, len(monthly_df), 41 ],
        "values": [ "Backend_Data", 1, 42, len(monthly_df), 42 ],
        "line": {"color": "#4B9CD3", "width": 3}
    })
    apply_dark_theme_to_chart(c8, "Monthly Revenue Trend")
    dashboard.insert_chart("J43", c8, {"x_scale": 1.3, "y_scale": 1.3})

print("Polished dark-mode Excel dashboard successfully generated!")
