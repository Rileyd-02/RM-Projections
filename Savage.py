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
    """Buy File -> PLM Upload (Savage)
    Keeps DESIGN STYLE, XFD, GLOBAL UNITS, converts XFD -> MONTH and pivots.
    """
    # read with header row 3 (index=2)
    df = pd.read_excel(uploaded_file, header=2)
    # normalize headers
    df.columns = df.columns.str.replace(r'[\n"]+', ' ', regex=True).str.strip()

    # required columns
    required = ["DESIGN STYLE", "XFD", "GLOBAL UNITS"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns for Savage Buy file: {missing}")

    df = df[required].copy()

    # parse XFD robustly (text dates or excel serials)
    df["XFD_dt"] = pd.to_datetime(df["XFD"], errors="coerce", dayfirst=True)
    if df["XFD_dt"].isna().all() and pd.api.types.is_numeric_dtype(df["XFD"]):
        df["XFD_dt"] = pd.to_datetime(df["XFD"], errors="coerce", unit="D", origin="1899-12-30")

    # map month
    month_map = {
        1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MAY", 6: "JUNE",
        7: "JULY", 8: "AUG", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC"
    }
    df["MONTH"] = df["XFD_dt"].dt.month.map(month_map)

    # filter valid months
    df = df[df["MONTH"].notna()]

    # pivot DESIGN STYLE x MONTH
    pivot_df = df.pivot_table(
        index="DESIGN STYLE",
        columns="MONTH",
        values="GLOBAL UNITS",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    # reorder months if present
    month_order = ["JAN","FEB","MAR","APR","MAY","JUNE","JULY","AUG","SEP","OCT","NOV","DEC"]
    non_month_cols = [c for c in pivot_df.columns if c not in month_order]
    ordered_cols = non_month_cols + [m for m in month_order if m in pivot_df.columns]
    pivot_df = pivot_df[ordered_cols]

    return pivot_df


def transform_plm_to_mcu(uploaded_file):
    """PLM Download -> Combine sheets -> MCU format for Savage.
    - Drops columns that start with 'Sum' (case-insensitive).
    - Adds 'Sheet Names' column.
    - Keeps standard MCU columns and any dynamic month columns found.
    """
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
            # normalize headers
            df.columns = df.columns.str.strip()

            # drop columns that start with "sum" or "sum of" (case-insensitive)
            mask_keep = ~df.columns.str.strip().str.lower().str.startswith("sum")
            df = df.loc[:, mask_keep]

            # add sheet name
            df.insert(0, "Sheet Names", sheet)

            # ensure base cols exist (fill missing with empty string)
            for col in base_cols:
                if col not in df.columns:
                    df[col] = ""

            # dynamic columns are any columns not in base_cols
            dynamic_cols = [c for c in df.columns if c not in base_cols]
            keep_cols = base_cols + dynamic_cols
            df = df.loc[:, keep_cols]

            collected.append(df)

    if not collected:
        # return empty DataFrame with base columns if nothing found
        return pd.DataFrame(columns=base_cols)

    combined = pd.concat(collected, ignore_index=True)

    # final ordering: base cols then any dynamic columns (in original order)
    dynamic_cols = [c for c in combined.columns if c not in base_cols]
    final_cols = base_cols + dynamic_cols
    combined = combined.loc[:, final_cols]

    return combined


# ----------------------------
# VSPINK logic
# ----------------------------
def transform_vspink(uploaded_file):
    """Transform VSPINK Brief to MCU format preserving metadata and pivoting months."""
    df = pd.read_excel(uploaded_file)
    df.columns = df.columns.str.strip()

    static_cols = [
        "Customer", "Supplier", "Supplier COO", "Production Plant (region)", "Program",
        "Construction", "Article", "# of repeats in Article ( optional)",
        "Composition", "If Yarn Dyed/ Piece Dyed"
    ]
    required_cols = static_cols + ["Qty (m)", "EX-mill"]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns for VSPINK: {missing}")

    # parse EX-mill robustly
    df["EX-mill_dt"] = pd.to_datetime(df["EX-mill"], errors="coerce", dayfirst=True)
    # Excel serial fallback
    if df["EX-mill_dt"].isna().all() and pd.api.types.is_numeric_dtype(df["EX-mill"]):
        df["EX-mill_dt"] = pd.to_datetime(df["EX-mill"], errors="coerce", unit="D", origin="1899-12-30")

    month_map = {
        1: "JAN", 2: "FEB", 3: "MAR", 4: "APR", 5: "MAY", 6: "JUNE",
        7: "JULY", 8: "AUG", 9: "SEP", 10: "OCT", 11: "NOV", 12: "DEC"
    }
    df["MONTH"] = df["EX-mill_dt"].dt.month.map(month_map)
    df = df[df["MONTH"].notna()]

    # group by static metadata + month and sum Qty (m)
    group_cols = static_cols + ["MONTH"]
    grouped = df.groupby(group_cols, dropna=False, as_index=False)["Qty (m)"].sum()

    # pivot
    pivot_df = grouped.pivot_table(
        index=static_cols,
        columns="MONTH",
        values="Qty (m)",
        aggfunc="sum",
        fill_value=0
    ).reset_index()

    month_order = ["JAN","FEB","MAR","APR","MAY","JUNE","JULY","AUG","SEP","OCT","NOV","DEC"]
    non_month = [c for c in pivot_df.columns if c not in month_order]
    ordered = non_month + [m for m in month_order if m in pivot_df.columns]
    pivot_df = pivot_df.loc[:, ordered]

    return pivot_df

# ----------------------------
# Page functions
# ----------------------------
def page_home():
    st.title("📦 MCU / PLM Tools Dashboard")
    st.markdown("""
    **Quick guide**
    - Select your account from the sidebar.
    - Each account page has uploaders for the files the workflow needs.
    - *Savage* supports both: **Buy file → PLM upload** and **PLM download → MCU**.
    - *VSPINK Brief* converts an input file to MCU format preserving metadata and month columns.
    - *HugoBoss* is a placeholder — add transform logic in the code where indicated.
    """)
    st.markdown("**Where to add new accounts:**\n- Create a transform function (like `transform_vspink`) and a `page_xxx()` function.\n- Add the page name to the sidebar options list below.")

def page_savage():
    st.header("Savage — Buy File → PLM Upload")
    st.markdown("Upload the buy file (headers on row 3). The output will be the PLM upload pivot with months as columns.")
    buy_file = st.file_uploader("Upload Buy file (Savage)", type=["xlsx","xls"], key="buy_file")
    if buy_file:
        try:
            df_out = transform_style_units(buy_file)
            st.subheader("Preview — PLM Upload")
            st.dataframe(df_out.head())
            out_bytes = excel_to_bytes(df_out, sheet_name="PLM Upload")
            st.download_button("📥 Download PLM Upload - savage.xlsx", out_bytes, file_name="plm upload - savage.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error processing buy file: {e}")

    st.markdown("---")
    st.header("Savage — PLM Download → MCU")
    st.markdown("Upload the PLM download workbook (multiple sheets). The app will combine supported sheets and produce the MCU-format single sheet.")
    plm_file = st.file_uploader("Upload PLM Download file (Savage)", type=["xlsx","xls"], key="plm_file")
    if plm_file:
        try:
            mcu = transform_plm_to_mcu(plm_file)
            st.subheader("Preview — MCU Combined")
            st.dataframe(mcu.head())
            out_bytes = excel_to_bytes(mcu, sheet_name="MCU")
            st.download_button("📥 Download MCU - savage.xlsx", out_bytes, file_name="MCU - savage.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error processing PLM download file: {e}")

def page_hugoboss():
    st.header("HugoBoss")
    st.info("HugoBoss logic not implemented yet. Add transform + page function in the code.")

def page_vspink():
    st.header("VSPINK Brief")
    st.markdown("Upload your VSPINK file. EX-mill will be converted to months and Qty (m) will be grouped/aggregated by Article and month, while preserving metadata columns.")
    up = st.file_uploader("Upload VSPINK file", type=["xlsx","xls"], key="vspink_file")
    if up:
        try:
            df_v = transform_vspink(up)
            st.subheader("Preview — VSPINK MCU")
            st.dataframe(df_v.head())
            out_bytes = excel_to_bytes(df_v, sheet_name="VSPINK MCU")
            st.download_button("📥 Download VSPINK MCU", out_bytes, file_name="vspink_mcu.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        except Exception as e:
            st.error(f"Error processing VSPINK file: {e}")

# ----------------------------
# App navigation (Sidebar radio - safe for all Streamlit versions)
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

# ----------------------------
# Helpful comment to future devs:
# To add a new account:
# 1) create transform function(s) above (data logic),
# 2) create page_<account>() that uses file_uploader -> calls the transform -> preview & download,
# 3) add the new name to the radio list and an elif branch calling page_<account>().
# ----------------------------
