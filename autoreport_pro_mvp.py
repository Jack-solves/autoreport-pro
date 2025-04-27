import streamlit as st
import pandas as pd
from io import BytesIO
import openai
import os

# OpenAI setup
client = openai.OpenAI()
openai.api_key = st.secrets.get("OPENAI_API_KEY", os.getenv("OPENAI_API_KEY"))

# Streamlit app setup
st.set_page_config(page_title="AutoReport Pro - Spreadsheet Cleaner + GPT Summary", layout="centered")
st.image("your_image.png", use_container_width=True)
st.title("📊 AutoReport Pro - Spreadsheet Cleaner + GPT Summary")

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

# Functions
def clean_data(df):
    initial_rows = df.shape[0]
    df_cleaned = df.dropna(how="all").drop_duplicates()
    removed_rows = initial_rows - df_cleaned.shape[0]
    return df_cleaned, initial_rows, removed_rows

def generate_xlsx_download(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer

def gpt_summary(df, initial_rows, removed_rows):
    prompt = f"""
You are a professional data analyst.

Here is the summary you need to create:
- Rows before cleaning: {initial_rows}
- Rows after cleaning: {df.shape[0]}
- Rows removed (empty or duplicates): {removed_rows}

Please also analyze:
- Average Salary (if exists).
- Most Frequent Department (if exists).
- Max and Min Salary.
- Any visible trends or anomalies.
Use bullet points and keep it short and professional.

Spreadsheet sample (first few rows):
{df.head(10).to_string(index=False)}
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a helpful and concise data analyst."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=500,
            temperature=0.5,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error generating summary: {e}"

# App main logic
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        st.success("✅ File successfully uploaded and processed.")

        cleaned_df, initial_rows, removed_rows = clean_data(df)

        st.subheader("🧼 Cleaned Data Preview")
        st.dataframe(cleaned_df, use_container_width=True)

        st.info(f"Removed {removed_rows} empty or duplicate row(s).")

        st.subheader("🏷️ Suggested Report Title")
        st.write("Suggested Title:")

        st.subheader("🧠 AI Report Summary")
        if "summary" not in st.session_state:
            with st.spinner("Generating insights..."):
                summary = gpt_summary(cleaned_df, initial_rows, removed_rows)
                st.session_state.summary = summary

        st.markdown(st.session_state.summary)

        # Download button
        xlsx_data = generate_xlsx_download(cleaned_df)
        st.download_button(
            label="⬇️ Download Cleaned Excel File",
            data=xlsx_data,
            file_name="cleaned_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
