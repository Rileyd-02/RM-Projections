import streamlit as st
import pandas as pd
from io import BytesIO

# ======================
# Savage Functions
# ======================
def transform_style_units(uploaded_file):
    """Buy → PLM Upload"""
    df = pd.read_excel(uploaded_file, header=2)
    df.columns = df.columns.str.replace(r'[\n"]+', ' ', regex=True).str.strip()

    required_cols = ['DESIGN STYLE', 'XFD', 'GLOBAL UNITS']
    df = df[[c for c in required_cols if c in df.columns]]

    df['XFD_dt'] = pd.to_datetime(df['XFD'], errors='coerce', dayfirst=True)
    if df['XFD_dt'].isna().all() and pd.api.types.is_numeric_dtype(df['XFD']):
        df['XFD_dt'] = pd.to_datetime(df['XFD'], errors='coerce', unit='D', origin='1899-12-30')

    month_map = {
        1: 'JAN', 2: 'FEB', 3: 'MAR', 4: 'APR', 5: 'MAY', 6: 'JUNE',
        7: 'JULY', 8: 'AUG', 9: 'SEP', 10: 'OCT', 11: 'NOV', 12: 'DEC'
    }
    df['MONTH'] = df['XFD_dt'].dt.month.map(month_map)
    df = df[df['MONTH'].notna()]

    pivot_df = df.pivot_table(
        index='DESIGN STYLE',
        columns='MONTH',
        values='GLOBAL UNITS',
        aggfunc='sum',
        fill_value=0
    ).reset_index()

    month_order = ['JAN','FEB','MAR','APR','MAY','JUNE','JULY','AUG','SEP','OCT','NOV','DEC']
    non_month_cols = [c for c in pivot_df.columns if c not in month_order]
    ordered_cols = non_month_cols + [m for m in month_order if m in pivot_df.columns]
    return pivot_df[ordered_cols]

def transform_plm_to_mcu(uploaded_file):
    """PLM Download → MCU"""
    sheet_names = [
        "Fabrics", "Strip Cut", "Laces", "Embriodery/Printing", "Elastics",
        "Tapes", "Trim/Component", "Label/ Transfer", "Foam Cup", "Packing Trim"
    ]

    dfs = []
    xls = pd.ExcelFile(uploaded_file)

    for sheet in sheet_names:
        if sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            df = df.loc[:, ~df.columns.str.startswith("Sum")]  # remove Sum cols
            df.insert(0, "Sheet Names", sheet)
            dfs.append(df)

    if not dfs:
        return pd.DataFrame()

    combined_df = pd.concat(dfs, ignore_index=True)

    standard_cols = [
        "Sheet Names","Season","Style","BOM","Cycle","Article",
        "Type of Const 1","Supplier","UOM","Composition",
        "Measurement","Supplier Country","Avg YY"
    ]
    for col in standard_cols:
        if col not in combined_df.columns:
            combined_df[col] = ""

    month_cols = [c for c in combined_df.columns if "-" in c and not c.startswith("Sum")]
    return combined_df[standard_cols + month_cols]

# ======================
# Savage Page
# ======================
def savage_page():
    st.title("📊 Savage Automation Tool")

    # Step 1: Buy → PLM
    st.header("Step 1️⃣: Buy File → PLM Upload")
    buy_file = st.file_uploader("Upload Buy File", type=["xlsx", "xls"], key="buy_savage")
    if buy_file:
        try:
            transformed_df = transform_style_units(buy_file)
            st.subheader("🔎 PLM Upload Preview")
            st.dataframe(transformed_df.head())

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                transformed_df.to_excel(writer, index=False, sheet_name="PLM Upload")
            output.seek(0)

            st.download_button("📥 Download PLM Upload", output,
                               file_name="plm_upload_savage.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"❌ Error processing Buy File: {e}")

    # Step 2: PLM → MCU
    st.header("Step 2️⃣: PLM Download → MCU Format")
    plm_file = st.file_uploader("Upload PLM Download File", type=["xlsx", "xls"], key="plm_savage")
    if plm_file:
        try:
            mcu_df = transform_plm_to_mcu(plm_file)
            st.subheader("🔎 MCU Preview")
            st.dataframe(mcu_df.head())

            output = BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                mcu_df.to_excel(writer, index=False, sheet_name="MCU Upload")
            output.seek(0)

            st.download_button("📥 Download MCU File", output,
                               file_name="mcu_savage.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"❌ Error processing PLM File: {e}")

# ======================
# HugoBoss Page
# ======================
def hugoboss_page():
    st.title("🧑‍💼 HugoBoss Automation Tool")
    st.info("👉 HugoBoss logic will be added here later.")

# ======================
# VSPINK Page
# ======================
def vspink_page():
    st.title("🩲 VSPINK Brief Automation Tool")
    st.info("👉 VSPINK Brief logic will be added here later.")

# ======================
# Navigation
# ======================
pg = st.navigation(
    {
        "Brands": [
            st.Page(savage_page, title="Savage", icon="📊"),
            st.Page(hugoboss_page, title="HugoBoss", icon="🧑‍💼"),
            st.Page(vspink_page, title="VSPINK Brief", icon="🩲"),
        ]
    }
)

pg.run()
