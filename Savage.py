import streamlit as st
import pandas as pd
import calendar

# ---------------- Savage Logic ----------------
def process_savage(file):
    try:
        df = pd.read_excel(file, header=2)  # skip first 2 rows, header is row 3
        df = df[['DESIGN STYLE', 'XFD', 'GLOBAL UNITS']]

        # Convert XFD to datetime, extract month name
        df['XFD'] = pd.to_datetime(df['XFD'], errors='coerce')
        df['Month'] = df['XFD'].dt.month.map(lambda m: calendar.month_abbr[m] if pd.notnull(m) else None)

        # Pivot months as columns
        pivot_df = df.pivot_table(
            index='DESIGN STYLE',
            columns='Month',
            values='GLOBAL UNITS',
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        return pivot_df

    except Exception as e:
        st.error(f"❌ Error processing Savage file: {e}")
        return None


# ---------------- VSPINK Brief Logic ----------------
def process_vspink(file):
    try:
        df = pd.read_excel(file)

        # Keep relevant columns
        df = df[['Customer', 'Supplier', 'Supplier COO', 'Production Plant (region)',
                 'Program', ' Qty (m) ', ' Construction  ', ' Article ',
                 '# of repeats in Article ( optional)', ' Composition ',
                 ' If Yarn Dyed/ Piece Dyed ', ' EX-mill ']]

        # Clean column names
        df.columns = df.columns.str.strip()

        # Convert EX-mill to month
        df['EX-mill'] = pd.to_datetime(df['EX-mill'], errors='coerce')
        df['Month'] = df['EX-mill'].dt.month.map(lambda m: calendar.month_abbr[m] if pd.notnull(m) else None)

        # Pivot so months become columns
        pivot_df = df.pivot_table(
            index=['Customer', 'Supplier', 'Supplier COO', 'Production Plant (region)',
                   'Program', 'Construction', 'Article',
                   '# of repeats in Article (optional)', 'Composition',
                   'If Yarn Dyed/ Piece Dyed'],
            columns='Month',
            values='Qty (m)',
            aggfunc='sum',
            fill_value=0
        ).reset_index()

        return pivot_df

    except Exception as e:
        st.error(f"❌ Error processing VSPink file: {e}")
        return None


# ---------------- Hugo Boss Logic (placeholder for now) ----------------
def process_hugoboss(file):
    try:
        df = pd.read_excel(file)
        # Placeholder: Just preview uploaded file
        return df.head()
    except Exception as e:
        st.error(f"❌ Error processing Hugo Boss file: {e}")
        return None


# ---------------- Streamlit App ----------------
st.set_page_config(layout="wide", page_title="MCU Converter App")

st.sidebar.title("Navigation")
page = st.sidebar.radio("Select Customer", ["Savage", "Hugo Boss", "VSPINK Brief"])

st.title(f"📂 {page} MCU Converter")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    if page == "Savage":
        result = process_savage(uploaded_file)
    elif page == "VSPINK Brief":
        result = process_vspink(uploaded_file)
    elif page == "Hugo Boss":
        result = process_hugoboss(uploaded_file)
    else:
        result = None

    if result is not None:
        st.subheader("✅ Processed Preview")
        st.dataframe(result)

        # Download button
        output_file = f"{page}_MCU_Output.xlsx"
        result.to_excel(output_file, index=False)
        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Download MCU File",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
