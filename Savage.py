import streamlit as st
import pandas as pd
from io import BytesIO

# -----------------------------
# Function 1: Buy File → PLM Upload
# -----------------------------
def transform_style_units(uploaded_file):
    """
    Reads the Buy Excel file, keeps only Design Style, XFD, and Global Units,
    extracts months from XFD, and pivots Global Units into monthly columns.
    """
    # --- Load Excel with headers from row 3 (index=2) ---
    df = pd.read_excel(uploaded_file, header=2)

    # --- Clean headers ---
    df.columns = df.columns.str.replace(r'[\n"]+', ' ', regex=True).str.strip()

    # --- Keep only required columns ---
    required_cols = ['DESIGN STYLE', 'XFD', 'GLOBAL UNITS']
    df = df[[c for c in required_cols if c in df.columns]]

    # --- Convert XFD to datetime ---
    df['XFD_dt'] = pd.to_datetime(df['XFD'], errors='coerce', dayfirst=True)

    # Handle Excel serial numbers if text parse failed
    if df['XFD_dt'].isna().all() and pd.api.types.is_numeric_dtype(df['XFD']):
        df['XFD_dt'] = pd.to_datetime(df['XFD'], errors='coerce', unit='D', origin='1899-12-30')

    # --- Extract month name ---
    month_map = {
        1: 'JAN', 2: 'FEB', 3: 'MAR', 4: 'APR', 5: 'MAY', 6: 'JUNE',
        7: 'JULY', 8: 'AUG', 9: 'SEP', 10: 'OCT', 11: 'NOV', 12: 'DEC'
    }
    df['MONTH'] = df['XFD_dt'].dt.month.map(month_map)

    # --- Drop rows without valid month ---
    df = df[df['MONTH'].notna()]

    # --- Pivot: Design Style as rows, Months as columns, Global Units as values ---
    pivot_df = df.pivot_table(
        index='DESIGN STYLE',
        columns='MONTH',
        values='GLOBAL UNITS',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    # --- Reorder month columns ---
    month_order = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUNE',
                   'JULY', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
    non_month_cols = [c for c in pivot_df.columns if c not in month_order]
    ordered_cols = non_month_cols + [m for m in month_order if m in pivot_df.columns]
    pivot_df = pivot_df[ordered_cols]

    return pivot_df


# -----------------------------
# Function 2: PLM Download → MCU Format
# -----------------------------
def transform_plm_to_mcu(uploaded_file):
    """
    Reads PLM download file with multiple sheets,
    drops "Sum*" columns, and combines into MCU format.
    """
    # --- Define expected sheet names ---
    sheet_names = [
        "Fabrics", "Strip Cut", "Laces", "Embriodery/Printing",
        "Elastics", "Tapes", "Trim/Component", "Label/ Transfer",
        "Foam Cup", "Packing Trim"
    ]

    # --- Standard columns to always include ---
    base_cols = [
        "Sheet Names", "Season", "Style", "BOM", "Cycle", "Article",
        "Type of Const 1", "Supplier", "UOM", "Composition",
        "Measurement", "Supplier Country", "Avg YY"
    ]

    all_data = []

    # --- Read the Excel file with all sheets ---
    xls = pd.ExcelFile(uploaded_file)

    for sheet in sheet_names:
        if sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)

            # Drop columns starting with "Sum"
            df = df[[c for c in df.columns if not c.startswith("Sum")]]

            # Add sheet name column
            df.insert(0, "Sheet Names", sheet)

            # Identify month/dynamic columns (everything not in base_cols except "Sheet Names")
            dynamic_cols = [c for c in df.columns if c not in base_cols]

            # Ensure we keep only base_cols + dynamic_cols
            keep_cols = [c for c in base_cols if c in df.columns] + dynamic_cols
            df = df[keep_cols]

            all_data.append(df)

    # --- Combine all sheets into one ---
    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)
    else:
        combined_df = pd.DataFrame(columns=base_cols)

    return combined_df


# -----------------------------
# Streamlit App
# -----------------------------
st.title("📊 Savage Automation Tool")

# --- Section 1: Buy File → PLM Upload ---
st.header("Step 1️⃣: Buy File → PLM Upload")
buy_file = st.file_uploader("Upload Buy File", type=["xlsx", "xls"], key="buy")

if buy_file:
    st.success("✅ Buy File uploaded successfully. Processing...")
    try:
        transformed_df = transform_style_units(buy_file)

        st.subheader("🔎 Preview of PLM Upload")
        st.dataframe(transformed_df.head())

        # Export to Excel in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            transformed_df.to_excel(writer, index=False, sheet_name="PLM Upload")
        output.seek(0)

        st.download_button(
            label="📥 Download PLM Upload File",
            data=output,
            file_name="plm upload - savage.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Error processing Buy File: {e}")


# --- Section 2: PLM Download → MCU Format ---
st.header("Step 2️⃣: PLM Download → MCU Format")
plm_file = st.file_uploader("Upload PLM Download File", type=["xlsx", "xls"], key="plm")

if plm_file:
    st.success("✅ PLM Download File uploaded successfully. Processing...")
    try:
        mcu_df = transform_plm_to_mcu(plm_file)

        st.subheader("🔎 Preview of MCU Output")
        st.dataframe(mcu_df.head())

        # Export MCU file
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            mcu_df.to_excel(writer, index=False, sheet_name="MCU")
        output.seek(0)

        st.download_button(
            label="📥 Download MCU File",
            data=output,
            file_name="MCU - savage.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Error processing PLM Download File: {e}")
