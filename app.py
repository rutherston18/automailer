import streamlit as st
import pandas as pd
import base64
import re
import time 
from email.message import EmailMessage
from datetime import datetime
import pytz

# Google Cloud & Auth Libraries
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# All three scopes are required for the app's full functionality
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send", 
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.readonly"
]

# --- AUTHENTICATION & SERVICE SETUP ---
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
        # Pad rows with empty strings to match header length
        data = [row + [''] * (len(headers) - len(row)) for row in data]
        return pd.DataFrame(data, columns=headers), headers, first_sheet_name
    except Exception as e:
        st.error(f"Failed to read Google Sheet. Check link and permissions. Error: {e}")
        return pd.DataFrame(), [], ""

def update_google_sheet_batch(sheets_service, spreadsheet_id, sheet_name, start_row, start_col, data_values):
    """Updates a range of cells in the Google Sheet in one batch call."""
    try:
        start_col_letter = chr(65 + start_col)
        end_col_letter = chr(65 + start_col + len(data_values[0]) - 1)
        range_to_update = f"{sheet_name}!{start_col_letter}{start_row}:{end_col_letter}{start_row + len(data_values) - 1}"

        body = {'values': data_values}
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=range_to_update,
            valueInputOption="USER_ENTERED", body=body
        ).execute()
    except Exception as e:
        st.warning(f"Failed to perform batch update on sheet. Error: {e}")

# --- EMAIL SENDING FUNCTIONS ---
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
        final_subject = f"Re: {subject}" if not subject.lower().startswith("re:") else subject
        
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
                        
                        log_headers = ["Timestamp", "Status", "Subject", "Thread ID", "Message ID"]
                        original_header_count = len(headers)
                        new_headers_to_add = [h for h in log_headers if h not in headers]
                        
                        if new_headers_to_add:
                            headers.extend(new_headers_to_add)
                            sheets_service.spreadsheets().values().update(
                                spreadsheetId=spreadsheet_id, range=f"{sheet_name}!A1",
                                valueInputOption="USER_ENTERED", body={'values': [headers]}
                            ).execute()
                        
                        update_data_values = []
                        
                        with st.expander("Live Send Status", expanded=True):
                            for i, row in df.iterrows():
                                row_data = row.to_dict()
                                email = row_data.get('email')
                                if not email or pd.isna(email): 
                                    st.write(f"_Row {i+2}: Skipping, no email address._")
                                    continue
                                
                                st.write(f"Row {i+2}: Sending to **{email}**...")
                                final_subject = subject_input.format(**row_data)
                                result = send_initial_email(gmail_service, email, final_subject, html_template, row_data)
                                
                                row_update_values = [''] * len(log_headers)

                                if result:
                                    msg_id_header = ""
                                    for attempt in range(4): 
                                        try:
                                            time.sleep(0.5 + attempt * 0.5) 
                                            full_message = gmail_service.users().messages().get(
                                                userId='me', id=result['id'], format='metadata', metadataHeaders=['Message-ID']
                                            ).execute()
                                            msg_headers = full_message.get('payload', {}).get('headers', [])
                                            msg_id_header = next((h['value'] for h in msg_headers if h['name'] == 'Message-ID'), '')
                                            if msg_id_header:
                                                st.write(f"&nbsp;&nbsp;&nbsp;↳ Success: Message-ID fetched.")
                                                break
                                        except Exception:
                                            if attempt == 3:
                                                st.warning(f"&nbsp;&nbsp;&nbsp;↳ Warning: Could not fetch Message-ID for {email} after several attempts.")
                                    
                                    row_update_values = [
                                        datetime.now(pytz.timezone("Asia/Kolkata")).strftime("%Y-%m-%d %H:%M:%S"),
                                        "Sent",
                                        final_subject,
                                        result.get('threadId', ''),
                                        msg_id_header
                                    ]
                                else:
                                    st.error(f"&nbsp;&nbsp;&nbsp;↳ Failed to send email to {email}.")
                                    row_update_values[1] = "Failed"
                                
                                update_data_values.append(row_update_values)
                        
                        if update_data_values:
                            st.info("Updating Google Sheet with logs...")
                            update_google_sheet_batch(
                                sheets_service, spreadsheet_id, sheet_name, 
                                start_row=2,
                                start_col=original_header_count,
                                data_values=update_data_values
                            )
                            st.success("Google Sheet updated!")
                        st.balloons()
            
            with tab2:
                # --- FIX: RESTORED REMINDER UI AND LOGIC ---
                st.subheader("Send a Follow-up or Reminder Email")
                st.info("This will send a threaded reply to contacts who have a 'Message ID' in the sheet.")
                
                uploaded_reminder_template = st.file_uploader("Upload Reminder/Reply Template (HTML)", type=["html"], key="reminder_template")

                if st.button("Start Reminder Campaign"):
                    if not uploaded_reminder_template:
                        st.warning("Please upload a reminder template.")
                    else:
                        reminder_template = uploaded_reminder_template.getvalue().decode("utf-8")
                        
                        # Check if required columns exist before trying to filter
                        if 'Message ID' not in df.columns or 'Thread ID' not in df.columns:
                            st.error("Cannot send reminders. 'Message ID' or 'Thread ID' column not found in the sheet.")
                        else:
                            reply_df = df[df['Message ID'].notna() & (df['Message ID'] != '')].copy()

                            if reply_df.empty:
                                st.warning("No contacts found with a valid 'Message ID' to reply to. Please run an initial campaign first.")
                            else:
                                with st.spinner(f"Sending reminders to {len(reply_df)} contacts..."):
                                    for i, row in reply_df.iterrows():
                                        row_data = row.to_dict()
                                        send_reply_email(
                                            gmail_service,
                                            row_data.get('email'),
                                            row_data.get('Subject'),
                                            row_data.get('Thread ID'),
                                            row_data.get('Message ID'),
                                            reminder_template,
                                            row_data
                                        )
                                    st.success("Reminder campaign sent!")
                                    st.balloons()
else:
    st.error("Application is offline. Could not authenticate to Google.")
