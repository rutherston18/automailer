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
def send_initial_email(service, to_email, subject, html_body_template, row_data):
    """Creates and sends a new email, returns the sent message object."""
    try:
        final_html_body = html_body_template.format(**row_data)
        final_subject = subject.format(**row_data)
        message = EmailMessage()
        message.add_alternative(final_html_body, subtype='html')
        message["To"] = to_email
        message["From"] = "me"
        message["Subject"] = final_subject
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        create_message = {"raw": encoded_message} # Labels are applied after sending
        sent_message = service.users().messages().send(userId="me", body=create_message).execute()
        return sent_message 
    except Exception as e:
        st.error(f"An error occurred sending initial email to {to_email}: {e}")
        return None

def apply_label_to_message(service, msg_id, label_ids_to_add):
    """Applies a list of labels to a specific message."""
    try:
        modify_request = {'addLabelIds': label_ids_to_add, 'removeLabelIds': []}
        service.users().messages().modify(userId='me', id=msg_id, body=modify_request).execute()
        return True
    except Exception as e:
        st.warning(f"Could not apply label to message {msg_id}. Error: {e}")
        return False


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
            time.sleep(2 ** attempt)
    st.warning("&nbsp;&nbsp;&nbsp;↳ Warning: Could not retrieve Message-ID after all retries")
    return ""

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
        create_message = {"raw": encoded_message, "threadId": thread_id} # Labels applied after
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
                subject_input = st.text_input("Initial Email Subject", value="Association with Entrepreneurship Cell, IIT Madras || {company}", key="initial_subject")
                
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
                        sent_emails_info = []
                        
                        with st.expander("Live Send Status", expanded=True):
                            st.write("--- Phase 1: Sending Emails ---")
                            for i, row in df.iterrows():
                                if pd.isna(row.get('email')) or not row.get('email'): continue
                                st.write(f"Row {i+2}: Sending to **{row.get('email')}**...")
                                result = send_initial_email(gmail_service, row.get('email'), subject_input, html_template, row.to_dict())
                                if result:
                                    temp_id = result['id']
                                    # --- FIX: Apply label immediately after sending ---
                                    if label_id_to_apply:
                                        apply_label_to_message(gmail_service, temp_id, label_id_to_apply)
                                        st.write(f"&nbsp;&nbsp;&nbsp;↳ Label '{selected_label_name}' applied.")
                                    # --- END FIX ---
                                    sent_emails_info.append({"row_index": i, "temp_id": temp_id, "thread_id": result['threadId'], "subject": subject_input.format(**row.to_dict())})
                                    st.write(f"&nbsp;&nbsp;&nbsp;↳ Success: Email sent.")
                                else:
                                    st.error(f"&nbsp;&nbsp;&nbsp;↳ Failed to send email to {row.get('email')}.")
                        
                        update_log = {}
                        if sent_emails_info:
                            with st.expander("Live Log Status", expanded=True):
                                st.write("\n--- Phase 2: Fetching Message IDs ---")
                                time.sleep(5)
                                for sent_item in sent_emails_info:
                                    i = sent_item["row_index"]
                                    msg_id_header = get_message_id_with_retry(gmail_service, sent_item['temp_id'])
                                    update_log[i] = {"Timestamp": datetime.now(pytz.timezone("Asia/Kolkata")).strftime("%Y-%m-%d %H:%M:%S"), "Status": "Sent", "Subject": sent_item["subject"], "Thread ID": sent_item["thread_id"], "Message ID": msg_id_header}

                        st.info("--- Phase 3: Updating Google Sheet with logs ---")
                        # This logic updates the original DataFrame and writes it back
                        log_headers = ["Timestamp", "Status", "Subject", "Thread ID", "Message ID"]
                        for header in log_headers:
                            if header not in df.columns:
                                df[header] = ''
                        for row_index, log_data in update_log.items():
                            for col_name, value in log_data.items():
                                df.loc[row_index, col_name] = value
                        try:
                            update_values = [df.columns.values.tolist()] + df.values.tolist()
                            sheets_service.spreadsheets().values().clear(spreadsheetId=spreadsheet_id, range=sheet_name).execute()
                            sheets_service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=f"{sheet_name}!A1", valueInputOption="USER_ENTERED", body={'values': update_values}).execute()
                            st.success("Google Sheet updated successfully!")
                            st.balloons()
                        except Exception as e:
                            st.error(f"Failed to update Google Sheet. Error: {e}")

            with tab2:
                st.subheader("Send a Follow-up or Reminder Email")
                
                selected_reply_label_name = st.selectbox(
                    "Apply this Gmail label to replies (optional):", 
                    options=[None] + list(gmail_labels.keys()),
                    key="reply_label"
                )
                reply_label_id_to_apply = [gmail_labels[selected_reply_label_name]] if selected_reply_label_name else []
                
                reminder_template = template_selector_ui("reminder")
                
                if st.button("Start Reminder Campaign"):
                    if not reminder_template:
                        st.warning("Please select/upload a reminder template.")
                    else:
                        if 'Message ID' not in df.columns or 'Thread ID' not in df.columns:
                            st.error("Cannot send reminders. 'Message ID' or 'Thread ID' column not found in the sheet.")
                        else:
                            reply_df = df[df['Message ID'].notna() & (df['Message ID'] != '')].copy()
                            if reply_df.empty:
                                st.warning("No contacts found with a valid 'Message ID' to reply to.")
                            else:
                                with st.spinner(f"Sending reminders to {len(reply_df)} contacts..."):
                                    for i, row in reply_df.iterrows():
                                        result = send_reply_email(
                                            gmail_service,
                                            row.get('email'),
                                            row.get('Subject'),
                                            row.get('Thread ID'),
                                            row.get('Message ID'),
                                            reminder_template,
                                            row.to_dict()
                                        )
                                        # --- FIX: Apply label to replies as well ---
                                        if result and reply_label_id_to_apply:
                                            apply_label_to_message(gmail_service, result['id'], reply_label_id_to_apply)
                                        # --- END FIX ---
                                    st.success("Reminder campaign sent!")
                                    st.balloons()
else:
    st.error("Application is offline. Could not authenticate to Google.")
