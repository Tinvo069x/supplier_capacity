import streamlit as st
import pandas as pd
import plotly.express as px
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

st.title("📊 Supplier Capacity Dashboard")

# Upload file
uploaded_file = st.file_uploader("Upload Excel Input", type=["xlsx"])
if uploaded_file:
    # Đọc dữ liệu
    capacity_df = pd.read_excel(uploaded_file, sheet_name="Capacity")
    demand_df = pd.read_excel(uploaded_file, sheet_name="Demand")

    # Tính capacity
    capacity_df["Capacity"] = (
        capacity_df["Lines"] *
        capacity_df["HoursPerDay"] *
        capacity_df["OutputPerHourPerLine"] *
        capacity_df["WorkingDays"]
    )

    # Reshape demand
    demand_long = demand_df.melt(
        id_vars=["Vendor","Item","Process"],
        var_name="Month", value_name="Demand"
    )
    demand_sum = demand_long.groupby(["Vendor","Process","Month"])["Demand"].sum().reset_index()

    # Merge
    merged = demand_sum.merge(
        capacity_df[["Vendor","Process","Capacity"]],
        on=["Vendor","Process"], how="left"
    )
    merged["Fulfillment_%"] = (merged["Capacity"] / merged["Demand"] * 100).round(2)
    merged["Status"] = merged.apply(lambda r: "OK" if r["Capacity"] >= r["Demand"] else "Shortage", axis=1)

    # Summary per Vendor
    summary_vendor = merged.groupby(["Vendor","Month"]).agg({
        "Capacity":"min",  # bottleneck
        "Demand":"sum"
    }).reset_index()
    summary_vendor["Fulfillment_%"] = (summary_vendor["Capacity"] / summary_vendor["Demand"] * 100).round(2)

    # Summary toàn bộ
    summary_total = merged.groupby("Month").agg({
        "Capacity":"sum","Demand":"sum"
    }).reset_index()
    summary_total["Fulfillment_%"] = (summary_total["Capacity"] / summary_total["Demand"] * 100).round(2)

    # ================== UI trên web ==================
    st.subheader("🔎 Vendor Summary")
    filter_mode = st.radio("Chọn chế độ xem:", ["All Vendors", "Shortage Only"])
    if filter_mode == "Shortage Only":
        shortage_vendors = summary_vendor[summary_vendor["Fulfillment_%"] < 100]
        st.dataframe(shortage_vendors)
    else:
        st.dataframe(summary_vendor)

    st.subheader("🌍 Total Supply Chain Summary")
    st.dataframe(summary_total)

    # ================== Chart ==================
    st.subheader("📈 Demand vs Capacity")
    vendor_selected = st.selectbox("Chọn Vendor để xem chart", ["ALL"] + sorted(summary_vendor["Vendor"].unique()))
    
    if vendor_selected == "ALL":
        fig = px.bar(
            summary_total, x="Month", y=["Demand","Capacity"],
            barmode="group", title="Total Demand vs Capacity"
        )
    else:
        vendor_data = summary_vendor[summary_vendor["Vendor"] == vendor_selected]
        fig = px.bar(
            vendor_data, x="Month", y=["Demand","Capacity"],
            barmode="group", title=f"Vendor {vendor_selected} - Demand vs Capacity"
        )
    st.plotly_chart(fig, use_container_width=True)

    # ================== Export Excel ==================
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
