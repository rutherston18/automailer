import streamlit as st
import pandas as pd
import base64
import re
import time 
import os
from pathlib import Path
from email.message import EmailMessage
from datetime import datetime
import pytz

# Google Cloud & Auth Libraries
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# All four scopes are required for the app's full functionality now
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send", 
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.modify" # Required to add labels
]

# --- TEMPLATE MANAGEMENT FUNCTIONS ---
def get_available_templates():
    """Get list of available HTML templates from the templates folder."""
    templates_dir = Path("templates")
    if not templates_dir.exists():
        templates_dir.mkdir(exist_ok=True)
        return []
    
    template_files = list(templates_dir.glob("*.html"))
    return [f.stem for f in template_files]

def load_template(template_name):
    """Load HTML template content from file."""
    template_path = Path("templates") / f"{template_name}.html"
    if template_path.exists():
        return template_path.read_text(encoding='utf-8')
    return None

def save_template(template_name, html_content):
    """Save HTML template to file."""
    templates_dir = Path("templates")
    templates_dir.mkdir(exist_ok=True)
    template_path = templates_dir / f"{template_name}.html"
    template_path.write_text(html_content, encoding='utf-8')

def template_selector_ui(template_type="initial"):
    """Create UI for template selection with option to use saved or upload new."""
    available_templates = get_available_templates()
    
    if available_templates:
        template_option = st.radio(
            f"Choose {template_type} template source:",
            ["Use saved template", "Upload new template"],
            key=f"{template_type}_template_option"
        )
        
        if template_option == "Use saved template":
            selected_template = st.selectbox(
                "Select template:",
                available_templates,
                key=f"{template_type}_template_select"
            )
            
            if selected_template:
                return load_template(selected_template)
        else:
            uploaded_file = st.file_uploader(
                f"Upload {template_type} Template (HTML)", 
                type=["html"], 
                key=f"{template_type}_upload"
            )
            
            if uploaded_file:
                template_content = uploaded_file.getvalue().decode("utf-8")
                
                save_option = st.checkbox(f"Save this template for future use", key=f"save_{template_type}")
                if save_option:
                    template_name = st.text_input(
                        "Template name (no spaces):", 
                        value=uploaded_file.name.replace('.html', '').replace(' ', '_'),
                        key=f"name_{template_type}"
                    )
                    if st.button(f"Save Template", key=f"save_btn_{template_type}"):
                        if template_name:
                            save_template(template_name, template_content)
                            st.success(f"Template '{template_name}' saved successfully!")
                            st.rerun()
                        else:
                            st.warning("Please enter a template name.")
                return template_content
    else:
        st.info("No saved templates found. Upload your first template below.")
        uploaded_file = st.file_uploader(
            f"Upload {template_type} Template (HTML)", 
            type=["html"], 
            key=f"{template_type}_upload_first"
        )
        if uploaded_file:
            template_content = uploaded_file.getvalue().decode("utf-8")
            save_option = st.checkbox(f"Save this template for future use", key=f"save_first_{template_type}")
            if save_option:
                template_name = st.text_input(
                    "Template name (no spaces):", 
                    value=uploaded_file.name.replace('.html', '').replace(' ', '_'),
                    key=f"name_first_{template_type}"
                )
                if st.button(f"Save Template", key=f"save_btn_first_{template_type}"):
                    if template_name:
                        save_template(template_name, template_content)
                        st.success(f"Template '{template_name}' saved successfully!")
                        st.rerun()
                    else:
                        st.warning("Please enter a template name.")
            return template_content
    return None

# --- AUTHENTICATION & SERVICE SETUP ---
@st.cache_resource
def get_preauthorized_services():
    """Authenticates and returns service objects for Gmail and Google Sheets."""
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
        return gmail_service, sheets_service
    except Exception as e:
        st.error("Failed to authenticate. Check secrets/scopes. Did you re-generate your token with the new 'gmail.modify' scope?")
        st.exception(e)
        return None, None

# --- NEW FUNCTION TO GET LABELS ---
@st.cache_data(ttl=600) # Cache labels for 10 minutes
def get_gmail_labels(_gmail_service):
    """Fetches all user-created labels from Gmail."""
    try:
        results = _gmail_service.users().labels().list(userId='me').execute()
        labels = results.get('labels', [])
        # Filter for user-created labels (system labels like INBOX, SPAM are excluded)
        user_labels = {label['name']: label['id'] for label in labels if label['type'] == 'user'}
        return user_labels
    except Exception as e:
        st.error(f"Could not fetch Gmail labels: {e}")
        return {}

# --- DATA & HELPER FUNCTIONS ---
def get_sheet_data(sheets_service, spreadsheet_id, sheet_name=None):
    try:
        if sheet_name is None:
            sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheets = sheet_metadata.get('sheets', '')
            sheet_name = sheets[0].get("properties", {}).get("title", "Sheet1")
        
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=f"{sheet_name}!A:Z"
        ).execute()
        values = result.get('values', [])
        
        if not values: return pd.DataFrame(), [], sheet_name
        
        headers = values[0]
        data = values[1:]
        df = pd.DataFrame(data, columns=headers)
        return df, headers, sheet_name
    except Exception as e:
        st.error(f"Failed to read Google Sheet '{sheet_name}'. Check link/permissions. Error: {e}")
        return pd.DataFrame(), [], ""

# --- EMAIL SENDING FUNCTIONS (UPDATED) ---
def send_initial_email(service, to_email, subject, html_body_template, row_data, label_ids=[]):
    """Creates, sends, and applies labels to a new email."""
    try:
        final_html_body = html_body_template.format(**row_data)
        final_subject = subject.format(**row_data)
        message = EmailMessage()
        message.add_alternative(final_html_body, subtype='html')
        message["To"] = to_email
        message["From"] = "me"
        message["Subject"] = final_subject
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        # Add labelIds to the message body for the send request
        create_message = {"raw": encoded_message, "labelIds": label_ids}
        sent_message = service.users().messages().send(userId="me", body=create_message).execute()
        return sent_message 
    except Exception as e:
        st.error(f"An error occurred sending initial email to {to_email}: {e}")
        return None

def get_message_id_with_retry(service, gmail_message_id):
    """Retrieves Message-ID header with exponential backoff retry logic."""
    for attempt in range(5):
        try:
            full_message = service.users().messages().get(userId='me', id=gmail_message_id, format='full').execute()
            message_id = next((h['value'] for h in full_message['payload']['headers'] if h['name'].lower() == 'message-id'), None)
            if message_id:
                st.write(f"&nbsp;&nbsp;&nbsp;↳ Success: Found Message-ID: `{message_id}`")
                return message_id
        except Exception:
            time.sleep(2 ** attempt) # Exponential backoff
    st.warning("&nbsp;&nbsp;&nbsp;↳ Warning: Could not retrieve Message-ID after all retries")
    return ""

def send_reply_email(service, to_email, subject, thread_id, original_msg_id, html_body_template, row_data, label_ids=[]):
    """Sends a reply and applies labels."""
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
        # Add threadId and labelIds to the request body
        create_message = {"raw": encoded_message, "threadId": thread_id, "labelIds": label_ids}
        sent_message = service.users().messages().send(userId="me", body=create_message).execute()
        return sent_message
    except Exception as e:
        st.error(f"An error occurred sending reply to {to_email}: {e}")
        return None

# --- MAIN STREAMLIT UI ---
st.set_page_config(layout="wide")
st.title("CR Mailing Scenes")

gmail_service, sheets_service = get_preauthorized_services()

if gmail_service and sheets_service:
    # --- FETCH LABELS ---
    gmail_labels = get_gmail_labels(gmail_service)
    
    st.header("Step 1: Put sheet link")
    sheet_url = st.text_input("Make sure all the columns are filled!")

    if sheet_url:
        try:
            spreadsheet_id = re.search('/d/([a-zA-Z0-9-_]+)', sheet_url).group(1)
            sheet_names = [s['properties']['title'] for s in sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()['sheets']]
            
            if sheet_names:
                selected_sheet = st.selectbox("Select which sheet to load:", options=sheet_names, key="sheet_selector")
                df, headers, sheet_name = get_sheet_data(sheets_service, spreadsheet_id, selected_sheet)
            else:
                st.error("No sheets found in the workbook."); st.stop()
        except:
            st.error("Invalid Google Sheet URL."); st.stop()

        if not df.empty:
            st.success(f"Successfully loaded {len(df)} rows from sheet: '{sheet_name}'")
            st.dataframe(df.head())

            tab1, tab2 = st.tabs(["🚀 Send Initial Campaign", "✉️ Send Reminders / Replies"])

            with tab1:
                st.subheader("Send First-Time Emails")
                subject_input = st.text_input("Initial Email Subject", value="Following up", key="initial_subject")
                
                # --- NEW LABEL SELECTOR ---
                selected_label_name = st.selectbox(
                    "Apply this Gmail label (optional):", 
                    options=[None] + list(gmail_labels.keys()),
                    key="initial_label"
                )
                label_id_to_apply = [gmail_labels[selected_label_name]] if selected_label_name else []
                
                html_template = template_selector_ui("initial")
                
                if st.button("Start Initial Campaign"):
                    if not all([html_template, subject_input]):
                        st.warning("Please provide a subject and select/upload an HTML template.")
                    else:
                        # ... (Rest of the sending logic) ...
                        # In the send_initial_email call, pass the label_id_to_apply
                        result = send_initial_email(gmail_service, row.get('email'), subject_input, html_template, row.to_dict(), label_id_to_apply)

            with tab2:
                st.subheader("Send a Follow-up or Reminder Email")
                
                # --- NEW LABEL SELECTOR FOR REPLIES ---
                selected_reply_label_name = st.selectbox(
                    "Apply this Gmail label to replies (optional):", 
                    options=[None] + list(gmail_labels.keys()),
                    key="reply_label"
                )
                reply_label_id_to_apply = [gmail_labels[selected_reply_label_name]] if selected_reply_label_name else []
                
                reminder_template = template_selector_ui("reminder")
                
                if st.button("Start Reminder Campaign"):
                    # ... (Rest of the reminder logic) ...
                    # In the send_reply_email call, pass the reply_label_id_to_apply
                    send_reply_email(gmail_service, row.get('email'), row.get('Subject'), row.get('Thread ID'), row.get('Message ID'), reminder_template, row.to_dict(), reply_label_id_to_apply)
else:
    st.error("Application is offline. Could not authenticate to Google.")
