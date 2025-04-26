import streamlit as st
import pandas as pd
import openai
from io import BytesIO
import os

# Correct OpenAI client setup
client = openai.OpenAI()  # Automatically uses secret

def chat_with_openai(messages, model="gpt-4"):
    response = client.chat.completions.create(
        model=model,
        messages=messages
    )
    return response.choices[0].message.content

# Set up Streamlit
st.set_page_config(page_title="AutoReport Pro", layout="centered")
st.title("📊 AutoReport Pro - Spreadsheet Cleaner + GPT Summary")
st.image("https://images.unsplash.com/photo-1603791440384-56cd371ee9a7?crop=entropy&cs=tinysrgb&fit=crop&h=150&q=80", use_column_width=True)

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

def clean_data(df):
    initial_rows = df.shape[0]
    df_cleaned = df.dropna(how="all").drop_duplicates()
    removed_rows = initial_rows - df_cleaned.shape[0]
    return df_cleaned, removed_rows

def generate_xlsx_download(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer

def gpt_summary(df, removed_rows):
    prompt = f"""
You are a professional data analyst. Analyze the following spreadsheet and provide:
1. Number of rows before and after cleaning.
2. Duplicates or empty rows removed.
3. Key statistics (e.g. average salary, frequent departments, max/min values).
4. Any anomalies or trends.
Use bullet points.

Spreadsheet sample (first few rows):
{df.head(10).to_string(index=False)}
"""
    try:
        summary = chat_with_openai([
            {"role": "system", "content": "You are a helpful and concise data analyst."},
            {"role": "user", "content": prompt}
        ])
        return summary
    except Exception as e:
        return f"Error generating summary: {e}"

def suggest_title(df):
    prompt = f"""
Based on the following spreadsheet data, suggest a short, professional report title (maximum 10 words). 
Be smart and specific if you can.

Spreadsheet sample (first few rows):
{df.head(10).to_string(index=False)}
"""
    try:
        title = chat_with_openai([
            {"role": "system", "content": "You are a creative and professional report title generator."},
            {"role": "user", "content": prompt}
        ])
        return title.strip()
    except Exception as e:
        return "Untitled Report"

# Main app logic
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        st.success("✅ File successfully uploaded and processed.")

        cleaned_df, removed_rows = clean_data(df)
        st.subheader("🧼 Cleaned Data Preview")
        st.dataframe(cleaned_df)

        st.info(f"Removed {removed_rows} empty or duplicate row(s).")

        # Title suggestion
        st.subheader("🏷️ Suggested Report Title")
        with st.spinner("Thinking of a professional title..."):
            report_title = suggest_title(cleaned_df)
            st.text_input("Suggested Title:", value=report_title)

        # GPT summary
        st.subheader("🧠 AI Report Summary")
        with st.spinner("Generating insights..."):
            summary = gpt_summary(cleaned_df, removed_rows)
            st.markdown(summary)

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
