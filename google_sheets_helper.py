import streamlit as st
import json
import os
import pickle
import pandas as pd
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file'
]

def get_gsheets_service():
    creds = None

    # Store token in memory (or in file in prod)
    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            # Read credentials from Streamlit secrets
            secrets = st.secrets["google_oauth_credentials"]
            with open("client_secret.json", "w") as f:
                 json.dump(dict(secrets), f)  # 👈 convert AttrDict to dict

            flow = InstalledAppFlow.from_client_secrets_file(
                "client_secret.json", SCOPES)
            creds = flow.run_console()

            with open("token.pickle", "wb") as token:
                pickle.dump(creds, token)

    service = build("sheets", "v4", credentials=creds)
    return service

def read_sheet(sheet_id, range_name="Sheet1"):
    service = get_gsheets_service()
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
    values = result.get("values", [])
    if not values:
        return pd.DataFrame()
    df = pd.DataFrame(values[1:], columns=values[0])
    return df

def write_sheet(df, title="AutoReport Export"):
    service = get_gsheets_service()
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
