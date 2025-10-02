import streamlit as st
import pandas as pd
from io import BytesIO

# ==============================
#  BRAND LOGIC FUNCTIONS
# ==============================

# --- Savage Functions ---
def transform_savage(uploaded_file):
    """Transform Savage file into MCU format (placeholder logic)."""
    df = pd.read_excel(uploaded_file, sheet_name=None, skiprows=2)  # skip 2 header rows

    # combine all sheets
    combined = pd.concat(df.values(), keys=df.keys(), names=["Sheet Names"])
    combined = combined.reset_index(level=0).reset_index(drop=True)

    # Example: keep MCU columns (adjust logic as needed)
    mcu_columns = [
        "Sheet Names", "Season", "Style", "BOM", "Cycle", "Article", "Type of Const 1",
        "Supplier", "UOM", "Composition", "Measurement", "Supplier Country", "Avg YY",
        "NOV", "DEC"  # dynamic month cols will be added automatically
    ]
    combined = combined[[c for c in mcu_columns if c in combined.columns]]

    return combined


def savage_page():
    st.title("🦍 Savage MCU Tool")
    uploaded_file = st.file_uploader("Upload Savage Excel File", type=["xlsx", "xls"], key="savage")
    if uploaded_file:
        try:
            transformed_df = transform_savage(uploaded_file)

            st.subheader("🔎 MCU Preview")
            st.dataframe(transformed_df.head())

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                transformed_df.to_excel(writer, index=False, sheet_name="Savage MCU")
            output.seek(0)

            st.download_button(
                label="📥 Download Savage MCU File",
                data=output,
                file_name="savage_mcu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")


# --- HugoBoss Functions (placeholder for now) ---
def hugoboss_page():
    st.title("👔 HugoBoss MCU Tool")
    st.info("Upload logic for HugoBoss will go here. 🚧")


# --- VSPINK Functions ---
def transform_vspink(uploaded_file):
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    static_cols = [
        "Customer", "Supplier", "Supplier COO", "Production Plant (region)", "Program",
        "Construction", "Article", "# of repeats in Article ( optional)",
        "Composition", "If Yarn Dyed/ Piece Dyed"
    ]
    required_cols = static_cols + ["Qty (m)", "EX-mill"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Missing required column: {col}")

    df["EX-mill_dt"] = pd.to_datetime(df["EX-mill"], errors="coerce", dayfirst=True)

    month_map = {
        1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MAY", 6: "JUNE",
        7: "JULY", 8: "AUG", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC"
    }
    df["MONTH"] = df["EX-mill_dt"].dt.month.map(month_map)
    df = df[df["MONTH"].notna()]

    grouped = (
        df.groupby(static_cols + ["MONTH"], dropna=False, as_index=False)["Qty (m)"]
        .sum()
    )

    pivot_df = grouped.pivot_table(
        index=static_cols,
        columns="MONTH",
        values="Qty (m)",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    month_order = ["JAN","FEB","MAR","APR","MAY","JUNE",
                   "JULY","AUG","SEP","OCT","NOV","DEC"]
    non_month_cols = [c for c in pivot_df.columns if c not in month_order]
    ordered_cols = non_month_cols + [m for m in month_order if m in pivot_df.columns]
    pivot_df = pivot_df[ordered_cols]

    return pivot_df


def vspink_page():
    st.title("🩲 VSPINK Brief MCU Tool")
    uploaded_file = st.file_uploader("Upload VSPINK Excel File", type=["xlsx", "xls"], key="vspink")
    if uploaded_file:
        try:
            transformed_df = transform_vspink(uploaded_file)

            st.subheader("🔎 MCU Preview")
            st.dataframe(transformed_df.head())

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                transformed_df.to_excel(writer, index=False, sheet_name="VSPINK MCU")
            output.seek(0)

            st.download_button(
                label="📥 Download VSPINK MCU File",
                data=output,
                file_name="vspink_mcu.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"❌ Error processing file: {e}")


# ==============================
#   NAVIGATION SETUP
# ==============================
savage = st.Page("Savage", page=savage_page)
hugoboss = st.Page("HugoBoss", page=hugoboss_page)
vspink = st.Page("VSPINK Brief", page=vspink_page)

pg = st.navigation(
    {
        "Accounts": [savage, hugoboss, vspink]
    }
)

pg.run()
