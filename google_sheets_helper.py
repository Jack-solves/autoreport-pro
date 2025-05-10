import streamlit as st
import json
import os
import pickle
import pandas as pd
from google_auth_oauthlib.flow import Flow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from collections import OrderedDict

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file"
]

def get_gsheets_service():
    creds = None

    if os.path.exists("token.pickle"):
        with open("token.pickle", "rb") as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            secrets = st.secrets["google_oauth_credentials"]
            redirect_uri = "https://autoreport-pro.streamlit.app"

            client_secrets = {
                "web": OrderedDict([
                    ("client_id", secrets["client_id"]),
                    ("client_secret", secrets["client_secret"]),
                    ("auth_uri", secrets["auth_uri"]),
                    ("token_uri", secrets["token_uri"]),
                    ("redirect_uris", [redirect_uri])
                ])
            }

            with open("client_secret.json", "w") as f:
                json.dump(client_secrets, f)

            flow = Flow.from_client_config(
                client_secrets,
                scopes=SCOPES,
                redirect_uri=redirect_uri
            )

            auth_url, _ = flow.authorization_url(prompt="consent")

            st.info("🔐 Please authorize access to your Google account:")
            st.markdown(f"[Click here to authorize]({auth_url})")

            auth_code = st.text_input("Paste the authorization code here:")

            if auth_code:
                try:
                    flow.fetch_token(code=auth_code)
                    creds = flow.credentials
                    with open("token.pickle", "wb") as token:
                        pickle.dump(creds, token)
                    st.success("✅ Authorization complete!")
                except Exception as e:
                    st.error(f"❌ Failed to fetch token: {e}")
                    return None
            else:
                st.stop()

    return build("sheets", "v4", credentials=creds)
