import streamlit as st
import pandas as pd
from io import BytesIO

def transform_style_units(uploaded_file):
    """
    Reads the Excel file, keeps only Design Style, XFD, and Global Units,
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


# --- Streamlit App ---
st.title("📊 PLM Upload - Savage Tool")
st.markdown("Upload your **buy file** and download the transformed output for PLM upload.")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

if uploaded_file:
    st.success("✅ File uploaded successfully. Processing...")
    
    try:
        transformed_df = transform_style_units(uploaded_file)

        # Show preview
        st.subheader("🔎 Preview of Transformed Data")
        st.dataframe(transformed_df.head())

        # Export to Excel in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            transformed_df.to_excel(writer, index=False, sheet_name="PLM Upload")
        output.seek(0)

        # Download button
        st.download_button(
            label="📥 Download PLM Upload File",
            data=output,
            file_name="plm upload - savage.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
