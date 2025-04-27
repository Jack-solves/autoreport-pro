import streamlit as st
import pandas as pd
import openai
from io import BytesIO
import os
import random
import requests
from PIL import Image

# ========== Setup ==========
st.set_page_config(page_title="AutoReport Pro", layout="centered")

# Title
st.title("📊 AutoReport Pro - Spreadsheet Cleaner + Smart Summary")

# Random reliable Unsplash image
image_pool = [
    "https://images.unsplash.com/photo-1559027615-2a9b5e00b5ec?auto=format&fit=crop&w=800&q=60",  # Office laptop
    "https://images.unsplash.com/photo-1556155092-490a1ba16284?auto=format&fit=crop&w=800&q=60",  # Data screen
    "https://images.unsplash.com/photo-1581091870627-3b6cc6c7f243?auto=format&fit=crop&w=800&q=60",  # Analytics chart
    "https://images.unsplash.com/photo-1556742400-b5a63d574047?auto=format&fit=crop&w=800&q=60",  # Desk with laptop
    "https://images.unsplash.com/photo-1603791452906-c5d31b4f59f5?auto=format&fit=crop&w=800&q=60",  # Minimal office
]
image_url = random.choice(image_pool)

try:
    response = requests.get(image_url, timeout=10)
    img = Image.open(BytesIO(response.content))
    st.image(img, use_container_width=True)
except Exception as e:
    st.warning("🖼️ Couldn't load image. Continuing without header image.")

# Friendly welcome message
st.markdown(
    "<div style='text-align: center; font-size: 18px; padding: 10px;'>"
    "✨ Upload your spreadsheet and let AI clean and summarize it for you! ✨"
    "</div>",
    unsafe_allow_html=True
)

# ========== OpenAI Setup ==========
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

def gpt_summary(df, removed_rows, filename):
    """Ask OpenAI to summarize the cleaned spreadsheet."""
    prompt = f"""
You are a professional data analyst. Strictly reply with:
- Number of rows before and after cleaning
- Duplicates or empty rows removed
- Key quick statistics (average salary, common departments, max/min values)
- 1-3 short bullet points on anomalies or interesting trends
**Do not explain that you cannot interact with the file, just answer with direct results.**

Spreadsheet sample (first few rows):
{df.head(10).to_string(index=False)}
"""

    try:
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a helpful and highly concise data analyst."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=500,
            temperature=0.2,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"Error generating summary: {e}"

def suggest_title(filename):
    """Generate a title based on filename."""
    base = os.path.splitext(filename)[0].replace("_", " ").replace("-", " ").title()
    return f"Data Quality Report for {base}"

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
        suggested_title = suggest_title(uploaded_file.name)
        st.markdown(f"**{suggested_title}**")

        st.subheader("🧠 AI Report Summary")
        with st.spinner("Generating insights..."):
            summary = gpt_summary(cleaned_df, removed_rows, uploaded_file.name)
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
