import streamlit as st
import pandas as pd
from io import BytesIO

# -----------------------------
# Savage: Buy → PLM Upload
# -----------------------------
def transform_style_units(uploaded_file):
    """
    Reads the Buy Excel file, keeps only Design Style, XFD, and Global Units,
    extracts months from XFD, and pivots Global Units into monthly columns.
    """
    df = pd.read_excel(uploaded_file, header=2)

    # Clean headers
    df.columns = df.columns.str.replace(r'[\n"]+', ' ', regex=True).str.strip()

    # Keep required columns
    required_cols = ['DESIGN STYLE', 'XFD', 'GLOBAL UNITS']
    df = df[[c for c in required_cols if c in df.columns]]

    # Convert XFD to datetime
    df['XFD_dt'] = pd.to_datetime(df['XFD'], errors='coerce', dayfirst=True)
    if df['XFD_dt'].isna().all() and pd.api.types.is_numeric_dtype(df['XFD']):
        df['XFD_dt'] = pd.to_datetime(df['XFD'], errors='coerce', unit='D', origin='1899-12-30')

    # Extract month names
    month_map = {
        1: 'JAN', 2: 'FEB', 3: 'MAR', 4: 'APR', 5: 'MAY', 6: 'JUNE',
        7: 'JULY', 8: 'AUG', 9: 'SEP', 10: 'OCT', 11: 'NOV', 12: 'DEC'
    }
    df['MONTH'] = df['XFD_dt'].dt.month.map(month_map)

    df = df[df['MONTH'].notna()]

    # Pivot table
    pivot_df = df.pivot_table(
        index='DESIGN STYLE',
        columns='MONTH',
        values='GLOBAL UNITS',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    # Reorder columns
    month_order = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUNE',
                   'JULY', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
    non_month_cols = [c for c in pivot_df.columns if c not in month_order]
    ordered_cols = non_month_cols + [m for m in month_order if m in pivot_df.columns]
    pivot_df = pivot_df[ordered_cols]

    return pivot_df


# -----------------------------
# Savage: PLM → MCU Format
# -----------------------------
def transform_plm_to_mcu(uploaded_file):
    """
    Reads PLM download file with multiple sheets, combines them, 
    removes 'Sum' columns, and outputs MCU format.
    """
    sheet_names = [
        "Fabrics", "Strip Cut", "Laces", "Embriodery/Printing", "Elastics",
        "Tapes", "Trim/Component", "Label/ Transfer", "Foam Cup", "Packing Trim"
    ]

    dfs = []
    xls = pd.ExcelFile(uploaded_file)

    for sheet in sheet_names:
        if sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)

            # Remove columns starting with "Sum"
            df = df.loc[:, ~df.columns.str.startswith("Sum")]

            # Add sheet name as a column
            df.insert(0, "Sheet Names", sheet)

            dfs.append(df)

    if not dfs:
        return pd.DataFrame()  # No valid sheets found

    combined_df = pd.concat(dfs, ignore_index=True)

    # Define standard MCU columns
    standard_cols = [
        "Sheet Names", "Season", "Style", "BOM", "Cycle", "Article",
        "Type of Const 1", "Supplier", "UOM", "Composition",
        "Measurement", "Supplier Country", "Avg YY"
    ]

    # Ensure all standard columns exist
    for col in standard_cols:
        if col not in combined_df.columns:
            combined_df[col] = ""

    # Keep standard + dynamic month columns
    month_cols = [c for c in combined_df.columns if "-" in c and not c.startswith("Sum")]
    final_cols = standard_cols + month_cols
    combined_df = combined_df[final_cols]

    return combined_df


# -----------------------------
# Streamlit App
# -----------------------------
st.sidebar.title("Navigation")
page = st.sidebar.radio("Select Brand / Project", ["Savage", "HugoBoss", "VSPINK Brief"])

# -----------------------------
# Page: Savage
# -----------------------------
if page == "Savage":
    st.title("📊 Savage Automation Tool")

    # Step 1: Buy → PLM
    st.header("Step 1️⃣: Buy File → PLM Upload")
    buy_file = st.file_uploader("Upload Buy File", type=["xlsx", "xls"], key="buy_savage")

    if buy_file:
        try:
            transformed_df = transform_style_units(buy_file)
            st.subheader("🔎 Preview of PLM Upload Data")
            st.dataframe(transformed_df.head())

            # Export
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                transformed_df.to_excel(writer, index=False, sheet_name="PLM Upload")
            output.seek(0)

            st.download_button(
                label="📥 Download PLM Upload File",
                data=output,
                file_name="plm_upload_savage.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"❌ Error processing Buy File: {e}")

    # Step 2: PLM → MCU
    st.header("Step 2️⃣: PLM Download → MCU Format")
    plm_file = st.file_uploader("Upload PLM Download File", type=["xlsx", "xls"], key="plm_savage")

    if plm_file:
        try:
            mcu_df = transform_plm_to_mcu(plm_file)
            st.subheader("🔎 Preview of MCU Data")
            st.dataframe(mcu_df.head())

            # Export
            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                mcu_df.to_excel(writer, index=False, sheet_name="MCU Upload")
            output.seek(0)

            st.download_button(
                label="📥 Download MCU File",
                data=output,
                file_name="mcu_savage.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"❌ Error processing PLM File: {e}")

# -----------------------------
# Page: HugoBoss
# -----------------------------
elif page == "HugoBoss":
    st.title("🧑‍💼 HugoBoss Automation Tool")
    st.info("👉 HugoBoss logic will be added here later.")

# -----------------------------
# Page: VSPINK Brief
# -----------------------------
elif page == "VSPINK Brief":
    st.title("🩲 VSPINK Brief Automation Tool")
    st.info("👉 VSPINK Brief logic will be added here later.")
