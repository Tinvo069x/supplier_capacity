import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO

st.set_page_config(page_title="Supplier Capacity Dashboard", layout="wide")
st.title("📊 Supplier Capacity Dashboard")

# Upload file Excel
uploaded_file = st.file_uploader("Upload Excel Input", type=["xlsx"])
if uploaded_file:
    # Đọc dữ liệu
    capacity_df = pd.read_excel(uploaded_file, sheet_name="Capacity")
    demand_df = pd.read_excel(uploaded_file, sheet_name="Demand")

    # ===== Tính Capacity =====
    capacity_df["Capacity"] = (
        capacity_df["Lines"] *
        capacity_df["HoursPerDay"] *
        capacity_df["OutputPerHourPerLine"] *
        capacity_df["WorkingDays"]
    )

    # ===== Demand reshape =====
    demand_long = demand_df.melt(
        id_vars=["Vendor","Item","Process"],
        var_name="Month", value_name="Demand"
    )
    demand_sum = demand_long.groupby(["Vendor","Process","Month"])["Demand"].sum().reset_index()

    # Merge với capacity
    merged = demand_sum.merge(
        capacity_df[["Vendor","Process","Capacity"]],
        on=["Vendor","Process"], how="left"
    )
    merged["Fulfillment_%"] = (merged["Capacity"] / merged["Demand"] * 100).round(2)
    merged["Status"] = merged.apply(lambda r: "OK" if r["Capacity"] >= r["Demand"] else "Shortage", axis=1)

    # ===== Summary theo Vendor =====
    summary_vendor = merged.groupby(["Vendor","Month"]).agg({
        "Capacity":"min",
        "Demand":"sum"
    }).reset_index()
    summary_vendor["Fulfillment_%"] = (summary_vendor["Capacity"] / summary_vendor["Demand"] * 100).round(2)

    # ===== Summary toàn bộ =====
    summary_total = merged.groupby("Month").agg({
        "Capacity":"sum","Demand":"sum"
    }).reset_index()
    summary_total["Fulfillment_%"] = (summary_total["Capacity"] / summary_total["Demand"] * 100).round(2)

    # ===== Chuẩn hóa Month sang yyyy-mm =====
    month_map = {
        "Jan":"2025-01","Feb":"2025-02","Mar":"2025-03","Apr":"2025-04","May":"2025-05",
        "Jun":"2025-06","Jul":"2025-07","Aug":"2025-08","Sep":"2025-09","Oct":"2025-10",
        "Nov":"2025-11","Dec":"2025-12"
    }
    summary_vendor["Month"] = summary_vendor["Month"].map(month_map)
    summary_total["Month"] = summary_total["Month"].map(month_map)

    summary_vendor["Month"] = pd.to_datetime(summary_vendor["Month"])
    summary_total["Month"] = pd.to_datetime(summary_total["Month"])

    summary_vendor = summary_vendor.sort_values(["Vendor","Month"])
    summary_total = summary_total.sort_values("Month")

    # ===== Slicer theo tháng =====
    months_available = sorted(summary_total["Month"].dt.strftime("%Y-%m").unique())
    months_selected = st.multiselect("📅 Chọn tháng:", months_available, default=months_available)

    if months_selected:
        summary_vendor = summary_vendor[summary_vendor["Month"].dt.strftime("%Y-%m").isin(months_selected)]
        summary_total = summary_total[summary_total["Month"].dt.strftime("%Y-%m").isin(months_selected)]

    # ===== Hiển thị bảng =====
    st.subheader("🔎 Vendor Summary")
    filter_mode = st.radio("Chọn chế độ xem:", ["All Vendors", "Shortage Only"])
    if filter_mode == "Shortage Only":
        st.dataframe(summary_vendor[summary_vendor["Fulfillment_%"] < 100])
    else:
        st.dataframe(summary_vendor)

    st.subheader("🌍 Total Supply Chain Summary")
    st.dataframe(summary_total)

    # ===== Chart Demand vs Capacity =====
    st.subheader("📈 Demand vs Capacity")
    vendor_selected = st.selectbox("Chọn Vendor", ["ALL"] + sorted(summary_vendor["Vendor"].unique()))

    if vendor_selected == "ALL":
        chart_data = summary_total.melt(id_vars="Month", value_vars=["Demand","Capacity"], var_name="Type", value_name="Value")
        fig = px.bar(chart_data, x="Month", y="Value", color="Type", barmode="group",
                     title="Total Demand vs Capacity", text="Value")
    else:
        vendor_data = summary_vendor[summary_vendor["Vendor"] == vendor_selected]
        chart_data = vendor_data.melt(id_vars=["Month"], value_vars=["Demand","Capacity"], var_name="Type", value_name="Value")
        fig = px.bar(chart_data, x="Month", y="Value", color="Type", barmode="group",
                     title=f"Vendor {vendor_selected} - Demand vs Capacity", text="Value")

    fig.update_traces(textposition="outside")  # hiển thị value trên bar
    fig.update_layout(
        xaxis=dict(
            tickformat="%Y-%m",
            type="category"
        ),
        uniformtext_minsize=8, uniformtext_mode="hide"
    )
    st.plotly_chart(fig, use_container_width=True)

    # ===== Xuất Excel =====
    out_file = BytesIO()
    with pd.ExcelWriter(out_file, engine="openpyxl") as writer:
        capacity_df.to_excel(writer, sheet_name="Capacity_Input", index=False)
        demand_df.to_excel(writer, sheet_name="Demand_Input", index=False)
        merged.to_excel(writer, sheet_name="Process_Result", index=False)
        summary_vendor.to_excel(writer, sheet_name="Vendor_Summary", index=False)
        summary_total.to_excel(writer, sheet_name="Total_Summary", index=False)

    st.download_button(
        label="⬇️ Download Result Excel",
        data=out_file.getvalue(),
        file_name="Supplier_Capacity_Result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
