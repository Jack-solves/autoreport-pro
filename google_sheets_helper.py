import streamlit as st
import json
import os
import pickle
import pandas as pd
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from collections import OrderedDict

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]

def get_gsheets_service():
    creds = None

    # Check if token already exists
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    # If no valid credentials, trigger Streamlit-compatible auth flow
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Load Google OAuth credentials from secrets
            secrets = st.secrets["google_oauth_credentials"]
            client_secrets = {
                "installed": OrderedDict([
                    ("client_id", secrets["client_id"]),
                    ("client_secret", secrets["client_secret"]),
                    ("auth_uri", secrets["auth_uri"]),
                    ("token_uri", secrets["token_uri"]),
                    ("redirect_uris", ["http://localhost"])
                ])
            }

            # Save to file temporarily
            with open("client_secret.json", "w") as f:
                json.dump(client_secrets, f)

            flow = InstalledAppFlow.from_client_secrets_file("client_secret.json", SCOPES)
            auth_url, _ = flow.authorization_url(prompt="consent", redirect_uri="http://localhost")

            st.info("🔐 Please authorize access to your Google account:")
            st.markdown(f"[Click here to authorize]({auth_url})")

            auth_code = st.text_input("Paste the authorization code here:")

            if auth_code:
                try:
                    flow.fetch_token(code=auth_code)
                    creds = flow.credentials

                    # Save the credentials to token.pickle
                    with open("token.pickle", "wb") as token:
                        pickle.dump(creds, token)

                    st.success("✅ Authorization complete!")
                except Exception as e:
                    st.error(f"❌ Failed to fetch token: {e}")
                    return None
            else:
                st.stop()  # Wait until code is entered

    service = build("sheets", "v4", credentials=creds)
    return service

def read_sheet(sheet_id, range_name="Sheet1"):
    service = get_gsheets_service()
    if not service:
        return pd.DataFrame()
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
    values = result.get("values", [])
    if not values:
        return pd.DataFrame()
    df = pd.DataFrame(values[1:], columns=values[0])
    return df

def write_sheet(df, title="AutoReport Export"):
    service = get_gsheets_service()
    if not service:
        return "❌ Authorization failed"
    spreadsheet = {"properties": {"title": title}}
    spreadsheet = service.spreadsheets().create(body=spreadsheet, fields="spreadsheetId").execute()
    sheet_id = spreadsheet.get("spreadsheetId")
    body = {
        "values": [df.columns.tolist()] + df.values.tolist()
    }
    service.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range="Sheet1",
        valueInputOption="RAW",
        body=body
    ).execute()
    return f"https://docs.google.com/spreadsheets/d/{sheet_id}"
