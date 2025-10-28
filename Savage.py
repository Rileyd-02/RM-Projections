import streamlit as st
import pandas as pd
from io import BytesIO

# ----------------------------
# Helper utilities
# ----------------------------
def excel_to_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1"):
    """Return bytes buffer of an Excel file for download."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output

# ----------------------------
# SAVAGE logic
# ----------------------------
def transform_style_units(uploaded_file):
    df = pd.read_excel(uploaded_file, header=2)
    df.columns = df.columns.str.replace(r'[\n"]+', ' ', regex=True).str.strip()

    required = ["DESIGN STYLE", "XFD", "GLOBAL UNITS"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns for Savage Buy file: {missing}")

    df = df[required].copy()
    df["XFD_dt"] = pd.to_datetime(df["XFD"], errors="coerce", dayfirst=True)
    if df["XFD_dt"].isna().all() and pd.api.types.is_numeric_dtype(df["XFD"]):
        df["XFD_dt"] = pd.to_datetime(df["XFD"], errors="coerce", unit="D", origin="1899-12-30")

    month_map = {1:"JAN",2:"FEB",3:"MAR",4:"APR",5:"MAY",6:"JUNE",
                 7:"JULY",8:"AUG",9:"SEP",10:"OCT",11:"NOV",12:"DEC"}
    df["MONTH"] = df["XFD_dt"].dt.month.map(month_map)
    df = df[df["MONTH"].notna()]

    pivot_df = df.pivot_table(
        index="DESIGN STYLE",
        columns="MONTH",
        values="GLOBAL UNITS",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    month_order = ["JAN","FEB","MAR","APR","MAY","JUNE","JULY","AUG","SEP","OCT","NOV","DEC"]
    non_month_cols = [c for c in pivot_df.columns if c not in month_order]
    ordered_cols = non_month_cols + [m for m in month_order if m in pivot_df.columns]
    pivot_df = pivot_df[ordered_cols]
    return pivot_df


def transform_plm_to_mcu(uploaded_file):
    expected_sheet_names = [
        "Fabrics", "Strip Cut", "Laces", "Embriodery/Printing",
        "Elastics", "Tapes", "Trim/Component", "Label/ Transfer",
        "Foam Cup", "Packing Trim"
    ]
    base_cols = [
        "Sheet Names", "Season", "Style", "BOM", "Cycle", "Article",
        "Type of Const 1", "Supplier", "UOM", "Composition",
        "Measurement", "Supplier Country", "Avg YY"
    ]
    xls = pd.ExcelFile(uploaded_file)
    collected = []

    for sheet in expected_sheet_names:
        if sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet)
            df.columns = df.columns.str.strip()
            mask_keep = ~df.columns.str.strip().str.lower().str.startswith("sum")
            df = df.loc[:, mask_keep]
            df.insert(0, "Sheet Names", sheet)
            for col in base_cols:
                if col not in df.columns:
                    df[col] = ""
            dynamic_cols = [c for c in df.columns if c not in base_cols]
            keep_cols = base_cols + dynamic_cols
            df = df.loc[:, keep_cols]
            collected.append(df)

    if not collected:
        return pd.DataFrame(columns=base_cols)

    combined = pd.concat(collected, ignore_index=True)
    dynamic_cols = [c for c in combined.columns if c not in base_cols]
    final_cols = base_cols + dynamic_cols
    combined = combined.loc[:, final_cols]
    return combined

# ----------------------------
# VSPINK logic
# ----------------------------
def transform_vspink_data(df):
    df.columns = df.columns.str.strip().str.replace("\n", " ").str.replace("\r", " ")

    exmill_col = [c for c in df.columns if "ex-mill" in c.lower()][0]
    qty_col = [c for c in df.columns if "qty" in c.lower()][0]
    article_col = [c for c in df.columns if "article" in c.lower()][0]

    metadata_cols = [
        "Customer", "Supplier", "Supplier COO", "Production Plant (region)",
        "Program", "Construction", article_col,
        "# of repeats in Article ( optional)", "Composition",
        "If Yarn Dyed/ Piece Dyed", exmill_col
    ]
    metadata_cols = [c for c in metadata_cols if c in df.columns]

    df["EX-mill_dt"] = pd.to_datetime(df[exmill_col], errors="coerce")
    df["Month-Year"] = df["EX-mill_dt"].dt.strftime("%b-%y")

    df[qty_col] = (
        df[qty_col].astype(str).str.replace(",", "", regex=False).str.strip()
    )
    df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce").fillna(0)

    grouped = df.groupby([article_col, "Month-Year"], as_index=False)[qty_col].sum()

    pivot_df = grouped.pivot_table(
        index=article_col,
        columns="Month-Year",
        values=qty_col,
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    parsed_months = pd.to_datetime(pivot_df.columns[1:], format="%b-%y", errors="coerce")
    month_order = parsed_months.sort_values().strftime("%b-%y").tolist()
    pivot_df = pivot_df[[article_col] + month_order]

    meta = df.groupby(article_col, as_index=False)[metadata_cols].first()
    final_df = pd.merge(meta, pivot_df, on=article_col, how="left")

    return final_df

# ----------------------------
# HUGO BOSS logic
# ----------------------------
def transform_hugoboss_buy_to_plm(df):
    df.columns = df.columns.str.strip()
    material_idx = df.columns.get_loc("Material Number")
    month_cols = df.columns[material_idx+1:]
    final_df = df[["Material Number"] + list(month_cols)]
    return final_df

def transform_hugoboss_plm_to_mcu(df):
    df.columns = df.columns.str.strip()
    mask_keep = ~df.columns.str.lower().str.startswith("sum")
    df = df.loc[:, mask_keep]
    return df

# ----------------------------
# Page functions
# ----------------------------
def page_home():
    st.title("📦 MCU Projections tool")
    st.markdown("""
    **⚠️ Important Notice**
    - Please do not change the sheet names in the uploaded file.
    - Make sure you are using the correct template format before uploading.
    - Any changes may cause errors in processing.

    **Quick guide**
    - *Savage*: Upload Buy Sheet → PLM upload | PLM upload → MCU Format
    - *VSPINK Brief*:VSPINK Brief Sheet → MCU Format
    - *HugoBoss*: Buy Sheet → PLM download | PLM upload → MCU Format
    """)

def page_savage():
    st.header("Savage — Buy File → PLM Upload")
    st.subheader("Bucket 02")
    buy_file = st.file_uploader("Upload Buy file (Savage)", type=["xlsx","xls"], key="buy_file")
    if buy_file:
        try:
            df_out = transform_style_units(buy_file)
            st.subheader("Preview — PLM Upload")
            st.dataframe(df_out.head())
            out_bytes = excel_to_bytes(df_out, sheet_name="PLM Upload")
            st.download_button("📥 Download PLM Upload - savage.xlsx", out_bytes,
                               file_name="plm_upload_savage.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error processing buy file: {e}")

    st.markdown("---")
    st.header("Savage — PLM Download → MCU")
    plm_file = st.file_uploader("Upload PLM Download file (Savage)", type=["xlsx","xls"], key="plm_file")
    if plm_file:
        try:
            mcu = transform_plm_to_mcu(plm_file)
            st.subheader("Preview — MCU Combined")
            st.dataframe(mcu.head())
            out_bytes = excel_to_bytes(mcu, sheet_name="MCU")
            st.download_button("📥 Download MCU - savage.xlsx", out_bytes,
                               file_name="MCU_savage.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error processing PLM download file: {e}")

def page_hugoboss():
    st.header("Hugo Boss — Buy Sheet → PLM Download")
    st.subheader("Bucket 02")
    buy_file = st.file_uploader("Upload Buy Sheet (HugoBoss)", type=["xlsx","xls"], key="hb_buy")
    if buy_file:
        try:
            df = pd.read_excel(buy_file)
            df_out = transform_hugoboss_buy_to_plm(df)
            st.subheader("Preview — PLM Download")
            st.dataframe(df_out.head())
            out_bytes = excel_to_bytes(df_out, sheet_name="PLM Download")
            st.download_button("📥 Download PLM Download - hugoboss.xlsx", out_bytes,
                               file_name="plm_download_hugoboss.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error processing HugoBoss Buy file: {e}")

    st.markdown("---")
    st.header("Hugo Boss — PLM Upload → MCU")
    plm_file = st.file_uploader("Upload PLM Upload file (HugoBoss)", type=["xlsx","xls"], key="hb_plm")
    if plm_file:
        try:
            df = pd.read_excel(plm_file)
            df_out = transform_hugoboss_plm_to_mcu(df)
            st.subheader("Preview — MCU")
            st.dataframe(df_out.head())
            out_bytes = excel_to_bytes(df_out, sheet_name="MCU")
            st.download_button("📥 Download MCU - hugoboss.xlsx", out_bytes,
                               file_name="MCU_hugoboss.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error processing HugoBoss PLM upload: {e}")

def page_vspink():
    st.header("VSPINK Brief")
    st.subheader("Bucket 03")
    up = st.file_uploader("Upload VSPINK file", type=["xlsx","xls"], key="vspink_file")
    if up:
        try:
            df = pd.read_excel(up)
            df_v = transform_vspink_data(df)
            st.subheader("Preview — VSPINK MCU")
            st.dataframe(df_v.head())
            out_bytes = excel_to_bytes(df_v, sheet_name="VSPINK MCU")
            st.download_button("📥 Download VSPINK MCU", out_bytes,
                               file_name="vspink_mcu.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error processing VSPINK file: {e}")

# ----------------------------
# Sidebar Navigation
# ----------------------------
st.sidebar.title("Accounts")
page_choice = st.sidebar.radio("Choose page", ["Home", "Savage", "HugoBoss", "VSPINK Brief"])

if page_choice == "Home":
    page_home()
elif page_choice == "Savage":
    page_savage()
elif page_choice == "HugoBoss":
    page_hugoboss()
elif page_choice == "VSPINK Brief":
    page_vspink()


