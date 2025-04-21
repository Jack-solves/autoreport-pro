
import streamlit as st
import pandas as pd
import io
from datetime import datetime

st.set_page_config(page_title="AutoReport Pro", layout="wide")
st.title("📊 AutoReport Pro - Spreadsheet Report Generator")

uploaded_file = st.file_uploader("Upload your spreadsheet (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        st.subheader("Original Data")
        st.dataframe(df)

        st.markdown("### 🔧 Data Cleaning in Progress...")
        df_cleaned = df.drop_duplicates().dropna()
        st.success("Cleaned successfully!")
        st.write(f"Removed {len(df) - len(df_cleaned)} rows (duplicates or missing).")

        st.subheader("🔢 Summary KPIs")
        st.write(f"🧾 Rows: {df_cleaned.shape[0]}")
        st.write(f"📁 Columns: {df_cleaned.shape[1]}")

        numeric_cols = df_cleaned.select_dtypes(include='number').columns
        if len(numeric_cols) > 0:
            st.write("📈 Averages:")
            st.dataframe(df_cleaned[numeric_cols].mean().to_frame("Average"))
        else:
            st.info("No numeric columns to summarize.")

        # Download buttons
        st.subheader("⬇️ Download Options")

        # Excel export
        output_excel = io.BytesIO()
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            df_cleaned.to_excel(writer, index=False, sheet_name='CleanedData')
        st.download_button(
            label="Download Cleaned Excel",
            data=output_excel.getvalue(),
            file_name="cleaned_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Simple PDF output (text only for MVP)
        import pdfkit
        html_string = f"<h1>AutoReport Summary</h1><p>Generated: {datetime.now()}</p>"
        html_string += f"<p>Rows: {df_cleaned.shape[0]}</p><p>Columns: {df_cleaned.shape[1]}</p>"
        html_string += "<ul>"
        for col in numeric_cols:
            html_string += f"<li>{col} Average: {df_cleaned[col].mean():.2f}</li>"
        html_string += "</ul>"

        try:
            pdf_bytes = pdfkit.from_string(html_string, False)
            st.download_button("Download PDF Report", data=pdf_bytes, file_name="summary_report.pdf", mime="application/pdf")
        except:
            st.warning("PDF generation skipped (local wkhtmltopdf may be required).")

    except Exception as e:
        st.error(f"An error occurred: {e}")
