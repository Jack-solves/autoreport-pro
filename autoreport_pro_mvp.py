import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="AutoReport Pro", layout="centered")
st.title("📊 AutoReport Pro - Spreadsheet Report Generator")

st.markdown("Upload your spreadsheet (.xlsx)")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx"])

def clean_data(df):
    initial_rows = df.shape[0]
    df = df.dropna(how="all")              # Remove rows where all cells are empty
    df = df.drop_duplicates()              # Remove duplicate rows
    removed = initial_rows - df.shape[0]
    return df, removed

def generate_xlsx_download(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        st.success("✅ File successfully uploaded and read.")

        cleaned_df, removed_rows = clean_data(df)
        st.subheader("🧼 Cleaned Preview")
        st.dataframe(cleaned_df)

        st.info(f"Removed {removed_rows} empty or duplicate row(s).")

        xlsx_data = generate_xlsx_download(cleaned_df)
        st.download_button(
            label="⬇️ Download Cleaned Excel File",
            data=xlsx_data,
            file_name="cleaned_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
