import streamlit as st
import pandas as pd
import openai
from io import BytesIO
import os
import random

# ========== Setup ==========
# Set Streamlit page config
st.set_page_config(page_title="AutoReport Pro", layout="centered")

# Title
st.title("📊 AutoReport Pro - Spreadsheet Cleaner + GPT Summary")

# Random image from Unsplash
topics = ["data", "spreadsheet", "analytics", "report", "business", "office", "technology"]
topic = random.choice(topics)
image_url = f"https://source.unsplash.com/800x250/?{topic}"
st.image(image_url, use_container_width=True)

# Friendly welcome message
st.markdown(
    "<div style='text-align: center; font-size: 18px; padding: 10px;'>"
    "✨ Upload your spreadsheet and let AI clean and summarize it for you! ✨"
    "</div>",
    unsafe_allow_html=True
)

# ========== OpenAI Setup ==========
# Automatically pick up API key from Streamlit secrets or environment variable
openai_api_key = st.secrets.get("OPENAI_API_KEY", os.getenv("OPENAI_API_KEY"))
client = openai.OpenAI(api_key=openai_api_key)

# ========== Functions ==========

def clean_data(df):
    """Remove completely empty rows and duplicates."""
    initial_rows = df.shape[0]
    df_cleaned = df.dropna(how="all").drop_duplicates()
    removed_rows = initial_rows - df_cleaned.shape[0]
    return df_cleaned, removed_rows

def generate_xlsx_download(df):
    """Prepare cleaned dataframe for download."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer

def gpt_summary(df, removed_rows):
    """Ask OpenAI to summarize the cleaned spreadsheet."""
    prompt = f"""
You are a professional data analyst. Analyze the following spreadsheet and provide:
1. Number of rows before and after cleaning.
2. Duplicates or empty rows removed.
3. Key statistics (e.g. average salary, frequent departments, max/min values).
4. Any anomalies or trends.
Use short bullet points.

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

# ========== Main App Logic ==========

uploaded_file = st.file_uploader("📂 Upload an Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        st.success("✅ File successfully uploaded and processed.")

        cleaned_df, removed_rows = clean_data(df)

        st.subheader("🧼 Cleaned Data Preview")
        st.dataframe(cleaned_df, use_container_width=True)

        st.info(f"🔍 Removed {removed_rows} empty or duplicate row(s).")

        st.subheader("🏷️ Suggested Report Title")
        st.markdown("Suggested Title:")

        st.subheader("🧠 AI Report Summary")
        with st.spinner("Generating insights..."):
            summary = gpt_summary(cleaned_df, removed_rows)
            st.markdown(summary)

        xlsx_data = generate_xlsx_download(cleaned_df)
        st.download_button(
            label="⬇️ Download Cleaned Excel File",
            data=xlsx_data,
            file_name="cleaned_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
