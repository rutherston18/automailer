import streamlit as st
import pandas as pd
import json
# ... other standard imports ...

# Google Cloud & Auth Libraries
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# This scope allows the app to send emails on the user's behalf.
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

def authenticate_gmail():
    """
    Handles Google Authentication for each user session.
    Stores credentials in st.session_state instead of a token.json file.
    """
    # Check if credentials are in the session state
    if "google_creds" in st.session_state and st.session_state.google_creds.valid:
        return build("gmail", "v1", credentials=st.session_state.google_creds)

    # If not, or if they are invalid, start the OAuth flow
    try:
        # Load credentials from Streamlit's secrets manager
        # IMPORTANT: 'google_credentials' must match the key you set in Streamlit Cloud secrets
        client_config = st.secrets["google_credentials"]
        
        flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
        
        # This will run a local server to get the authorization code.
        # Streamlit Cloud handles the redirect URI automatically.
        creds = flow.run_local_server(port=0)
        
        # Save the valid credentials in the user's session state for reuse
        st.session_state.google_creds = creds
        
        # Refresh the page to show the logged-in state
        st.experimental_rerun()
        
        return build("gmail", "v1", credentials=creds)

    except Exception as e:
        st.error(f"Authentication failed: {e}")
        st.write("Please ensure you have configured your secrets correctly in Streamlit Community Cloud.")
        return None

# --- Main Streamlit UI ---

st.title("📧 Web-Based Email Campaign Tool")

# Check if the user is logged in by looking for credentials in the session state
if "google_creds" not in st.session_state or not st.session_state.google_creds.valid:
    st.header("Step 1: Log in with Google")
    st.write("You need to authorize this application to send emails on your behalf.")
    if st.button("Login with Google"):
        # This button click will trigger the authentication flow
        authenticate_gmail()
else:
    # User is logged in, show the main application
    st.success(f"Logged in as: {st.session_state.google_creds.id_token.get('email')}")
    st.header("Step 2: Prepare Your Campaign")

    subject_input = st.text_input("Enter Email Subject", "Following up on our conversation")
    uploaded_csv = st.file_uploader("Upload Contacts (CSV)", type=["csv"])
    uploaded_template = st.file_uploader("Upload Template (HTML)", type=["html"])

    if st.button("Send Email Campaign"):
        if not all([uploaded_csv, uploaded_template]):
            st.warning("Please upload both a CSV and an HTML template.")
        else:
            # --- Get the authenticated Gmail service ---
            gmail_service = authenticate_gmail()
            if gmail_service:
                with st.spinner("Sending emails... This may take a moment."):
                    # --- File processing and email sending logic ---
                    # (This logic remains the same: read pandas df, loop, send email)
                    # ...
                    st.success("Email campaign sent successfully!")
                    st.balloons()

    if st.button("Logout"):
        del st.session_state.google_creds
        st.experimental_rerun()