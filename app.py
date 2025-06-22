import streamlit as st
import pandas as pd
import base64
import re
from email.message import EmailMessage
from datetime import datetime
import pytz

# Google Cloud & Auth Libraries
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# We need permission to send Gmail AND read/write Sheets.
SCOPES = ["https://www.googleapis.com/auth/gmail.send", "https://www.googleapis.com/auth/spreadsheets"]

# --- AUTHENTICATION & SERVICE SETUP (UNCHANGED) ---
@st.cache_resource
def get_preauthorized_services():
    """Authenticates and returns service objects for both Gmail and Google Sheets."""
    gmail_service = None
    sheets_service = None
    try:
        creds_data = st.secrets["preauthorized_account"]
        creds = Credentials.from_authorized_user_info(
            info={
                "refresh_token": creds_data["refresh_token"],
                "client_id": creds_data["client_id"],
                "client_secret": creds_data["client_secret"],
                "token_uri": "https://oauth2.googleapis.com/token"
            },
            scopes=SCOPES
        )
        gmail_service = build("gmail", "v1", credentials=creds)
        sheets_service = build("sheets", "v4", credentials=creds)
        st.success("Successfully authenticated with Google services.")
    except Exception as e:
        st.error("Failed to authenticate with pre-authorized account. Check secrets/scopes.")
        st.exception(e)
    return gmail_service, sheets_service

# --- DATA & HELPER FUNCTIONS ---
def get_sheet_data(sheets_service, spreadsheet_id):
    """Reads the first visible sheet and returns its data as a DataFrame."""
    try:
        sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', '')
        first_sheet_name = sheets[0].get("properties", {}).get("title", "Sheet1")
        
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=first_sheet_name
        ).execute()
        values = result.get('values', [])
        
        if not values:
            return pd.DataFrame(), [], ""
        
        headers = values[0]
        data = values[1:]
        return pd.DataFrame(data, columns=headers), headers, first_sheet_name
    except Exception as e:
        st.error(f"Failed to read Google Sheet. Check link and permissions. Error: {e}")
        return pd.DataFrame(), [], ""

def update_google_sheet(sheets_service, spreadsheet_id, sheet_name, row_index, col_index, value):
    """Updates a single cell in the Google Sheet."""
    try:
        # Excel-style cell notation, e.g., A1, B2
        # chr(65) is 'A'. So col_index 0 becomes 'A', 1 becomes 'B', etc.
        column_letter = chr(65 + col_index)
        range_to_update = f"{sheet_name}!{column_letter}{row_index + 1}"
        body = {'values': [[value]]}
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=range_to_update,
            valueInputOption="USER_ENTERED", body=body
        ).execute()
    except Exception as e:
        st.warning(f"Failed to update cell at {range_to_update}. Error: {e}")

# --- EMAIL SENDING FUNCTIONS (UNCHANGED) ---
def send_initial_email(service, to_email, subject, html_body_template, row_data):
    """Creates and sends a new personalized HTML email."""
    try:
        final_html_body = html_body_template.format(**row_data)
        final_subject = subject.format(**row_data)
        message = EmailMessage()
        message.add_alternative(final_html_body, subtype='html')
        message["To"] = to_email
        message["From"] = "me"
        message["Subject"] = final_subject
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        create_message = {"raw": encoded_message}
        sent_message = service.users().messages().send(userId="me", body=create_message).execute()
        return sent_message 
    except Exception as e:
        st.error(f"An error occurred sending initial email to {to_email}: {e}")
        return None

def send_reply_email(service, to_email, subject, thread_id, original_msg_id, html_body_template, row_data):
    """Sends a reply within an existing email thread."""
    try:
        final_html_body = html_body_template.format(**row_data)
        final_subject = f"Re: {subject}" 
        
        message = EmailMessage()
        message.add_alternative(final_html_body, subtype='html')
        message["To"] = to_email
        message["From"] = "me"
        message["Subject"] = final_subject
        message["In-Reply-To"] = original_msg_id
        message["References"] = original_msg_id
        
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        create_message = {"raw": encoded_message, "threadId": thread_id}
        
        sent_message = service.users().messages().send(userId="me", body=create_message).execute()
        return sent_message
    except Exception as e:
        st.error(f"An error occurred sending reply to {to_email}: {e}")
        return None

# --- MAIN STREAMLIT UI ---
st.set_page_config(layout="wide")
st.title("📧 Google Sheet Campaign & Reply Tool")
st.info("This tool uses a Google Sheet for contacts, logs sent emails, and can send threaded replies.")

gmail_service, sheets_service = get_preauthorized_services()

if gmail_service and sheets_service:
    st.header("Step 1: Link Your Google Sheet")
    sheet_url = st.text_input("Paste the full URL of your Google Sheet here")

    if sheet_url:
        try:
            spreadsheet_id = re.search('/d/([a-zA-Z0-9-_]+)', sheet_url).group(1)
            df, headers, sheet_name = get_sheet_data(sheets_service, spreadsheet_id)
        except Exception:
            st.error("Invalid Google Sheet URL. Please paste the full URL from your browser's address bar.")
            st.stop()

        if not df.empty:
            st.success(f"Successfully loaded {len(df)} rows from sheet: '{sheet_name}'")
            st.dataframe(df.head())

            tab1, tab2 = st.tabs(["🚀 Send Initial Campaign", "✉️ Send Reminders / Replies"])

            with tab1:
                st.subheader("Send First-Time Emails")
                st.info("This will send an email to every contact in the sheet and log the results back to the sheet.")
                
                subject_input = st.text_input("Initial Email Subject", key="initial_subject")
                uploaded_template = st.file_uploader("Upload Initial Email Template (HTML)", type=["html"], key="initial_template")
                
                if st.button("Start Initial Campaign"):
                    if not all([uploaded_template, subject_input]):
                        st.warning("Please provide a subject and an HTML template.")
                    else:
                        html_template = uploaded_template.getvalue().decode("utf-8")
                        with st.spinner(f"Sending initial emails to {len(df)} contacts..."):
                            # --- FIX: Ensure log columns exist before looping ---
                            log_headers = ["Timestamp", "Status", "Subject", "Thread ID", "Message ID"]
                            for header in log_headers:
                                if header not in headers:
                                    headers.append(header)
                            # Update the header row in the sheet if new columns were added
                            sheets_service.spreadsheets().values().update(
                                spreadsheetId=spreadsheet_id, range=f"{sheet_name}!A1",
                                valueInputOption="USER_ENTERED", body={'values': [headers]}
                            ).execute()
                            # --- END FIX ---

                            for i, row in df.iterrows():
                                row_data = row.to_dict()
                                email = row_data.get('email')
                                if not email or pd.isna(email):
                                    continue
                                
                                final_subject = subject_input.format(**row_data)
                                result = send_initial_email(gmail_service, email, final_subject, html_template, row_data)
                                if result:
                                    # Update the sheet with log data
                                    update_google_sheet(sheets_service, spreadsheet_id, sheet_name, i + 1, headers.index("Timestamp"), datetime.now(pytz.timezone("Asia/Kolkata")).strftime("%Y-%m-%d %H:%M:%S"))
                                    update_google_sheet(sheets_service, spreadsheet_id, sheet_name, i + 1, headers.index("Status"), "Sent")
                                    update_google_sheet(sheets_service, spreadsheet_id, sheet_name, i + 1, headers.index("Subject"), final_subject) # <-- FIX: LOG THE SUBJECT
                                    update_google_sheet(sheets_service, spreadsheet_id, sheet_name, i + 1, headers.index("Thread ID"), result.get('threadId'))
                                    update_google_sheet(sheets_service, spreadsheet_id, sheet_name, i + 1, headers.index("Message ID"), result.get('id'))
                            st.success("Initial campaign sent and logs updated in your Google Sheet!")
                            st.balloons()
            
            with tab2:
                st.subheader("Send a Follow-up or Reminder Email")
                st.info("This will send a threaded reply to contacts who have a 'Thread ID' in the sheet.")
                
                reminder_subject = st.text_input("Original Subject Line (for context)", key="reminder_subject", help="Enter the original subject to filter, or leave blank to target all.")
                uploaded_reminder_template = st.file_uploader("Upload Reminder/Reply Template (HTML)", type=["html"], key="reminder_template")

                if st.button("Start Reminder Campaign"):
                    if not uploaded_reminder_template:
                        st.warning("Please upload a reminder template.")
                    else:
                        reminder_template = uploaded_reminder_template.getvalue().decode("utf-8")
                        # Filter for contacts that have been successfully emailed before
                        reply_df = df[df['Thread ID'].notna() & (df['Thread ID'] != '')].copy()
                        
                        # --- FIX: Check if Subject column exists before filtering ---
                        if reminder_subject and 'Subject' in reply_df.columns:
                            reply_df = reply_df[reply_df['Subject'] == reminder_subject]
                        elif reminder_subject:
                            st.warning("Cannot filter by subject because the 'Subject' column was not found in your sheet.")
                        # --- END FIX ---

                        if reply_df.empty:
                            st.warning("No contacts found with a Thread ID to reply to (or matching the specified subject). Please run an initial campaign first.")
                        else:
                            with st.spinner(f"Sending reminders to {len(reply_df)} contacts..."):
                                for i, row in reply_df.iterrows():
                                    row_data = row.to_dict()
                                    send_reply_email(
                                        gmail_service,
                                        row_data.get('email'),
                                        row_data.get('Subject'), # The original subject
                                        row_data.get('Thread ID'),
                                        row_data.get('Message ID'),
                                        reminder_template,
                                        row_data
                                    )
                                st.success("Reminder campaign sent!")
                                st.balloons()
else:
    st.error("Application is offline. Could not authenticate to Google.")
