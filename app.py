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

# All three scopes are required for the app's full functionality
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send", 
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.readonly"
]

# --- TEMPLATE MANAGEMENT FUNCTIONS ---
def get_available_templates():
    """Get list of available HTML templates from the templates folder."""
    templates_dir = Path("templates")
    if not templates_dir.exists():
        templates_dir.mkdir(exist_ok=True)
        return []
    
    template_files = list(templates_dir.glob("*.html"))
    return [f.stem for f in template_files]  # Return names without .html extension

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
                template_content = load_template(selected_template)
                if template_content:
                    with st.expander("Preview Template"):
                        st.code(template_content[:500] + "..." if len(template_content) > 500 else template_content, language="html")
                    return template_content
                else:
                    st.error("Could not load selected template.")
                    return None
        else:
            # Upload new template
            uploaded_file = st.file_uploader(
                f"Upload {template_type} Template (HTML)", 
                type=["html"], 
                key=f"{template_type}_upload"
            )
            
            if uploaded_file:
                template_content = uploaded_file.getvalue().decode("utf-8")
                
                # Option to save this template
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
                            st.rerun()  # Refresh to show in saved templates
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
            
            # Option to save this template
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
    """Authenticates and returns service objects for both Gmail and Google Sheets."""
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
        st.error("Failed to authenticate with pre-authorized account. Check secrets/scopes.")
        st.exception(e)
        return None, None

# --- DATA & HELPER FUNCTIONS ---
def get_all_sheet_names(sheets_service, spreadsheet_id):
    """Returns a list of all sheet names in the spreadsheet."""
    try:
        sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', [])
        sheet_names = [sheet.get("properties", {}).get("title", "Sheet1") for sheet in sheets]
        return sheet_names
    except Exception as e:
        st.error(f"Failed to get sheet names. Error: {e}")
        return []

def get_sheet_data(sheets_service, spreadsheet_id, sheet_name=None):
    """Reads the specified sheet and returns its data as a DataFrame."""
    try:
        if sheet_name is None:
            # Get the first sheet if no sheet name specified
            sheet_metadata = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
            sheets = sheet_metadata.get('sheets', '')
            sheet_name = sheets[0].get("properties", {}).get("title", "Sheet1")
        
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=f"{sheet_name}!A:Z"
        ).execute()
        values = result.get('values', [])
        
        if not values:
            return pd.DataFrame(), [], sheet_name
        
        headers = values[0]
        data = values[1:]
        df = pd.DataFrame(data, columns=headers)
        return df, headers, sheet_name
    except Exception as e:
        st.error(f"Failed to read Google Sheet '{sheet_name}'. Check link and permissions. Error: {e}")
        return pd.DataFrame(), [], ""

# --- IMPROVED EMAIL SENDING FUNCTIONS ---
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

def get_message_id_with_retry(service, gmail_message_id, max_retries=5, base_delay=2):
    """
    Retrieve Message-ID header with exponential backoff retry logic.
    Gmail API sometimes needs time to process and index sent messages.
    """
    for attempt in range(max_retries):
        try:
            # Use 'full' format to get complete message data including all headers
            full_message = service.users().messages().get(
                userId='me', 
                id=gmail_message_id, 
                format='full'  # Changed from 'metadata' to 'full'
            ).execute()
            
            # Extract headers from the message payload
            payload = full_message.get('payload', {})
            headers = payload.get('headers', [])
            
            # Look for Message-ID header (case-insensitive)
            message_id = None
            for header in headers:
                if header.get('name', '').lower() == 'message-id':
                    message_id = header.get('value', '')
                    break
            
            if message_id:
                st.write(f"&nbsp;&nbsp;&nbsp;↳ Success: Found Message-ID: `{message_id}`")
                return message_id
            else:
                st.write(f"&nbsp;&nbsp;&nbsp;↳ Attempt {attempt + 1}: Message-ID not found in headers")
                if attempt < max_retries - 1:
                    delay = base_delay * (2 ** attempt)  # Exponential backoff
                    st.write(f"&nbsp;&nbsp;&nbsp;↳ Waiting {delay} seconds before retry...")
                    time.sleep(delay)
                    
        except HttpError as e:
            if e.resp.status == 404:
                st.write(f"&nbsp;&nbsp;&nbsp;↳ Attempt {attempt + 1}: Message not found (404)")
            else:
                st.write(f"&nbsp;&nbsp;&nbsp;↳ Attempt {attempt + 1}: HTTP Error {e.resp.status}")
            
            if attempt < max_retries - 1:
                delay = base_delay * (2 ** attempt)
                st.write(f"&nbsp;&nbsp;&nbsp;↳ Waiting {delay} seconds before retry...")
                time.sleep(delay)
                
        except Exception as e:
            st.write(f"&nbsp;&nbsp;&nbsp;↳ Attempt {attempt + 1}: Unexpected error: {str(e)}")
            if attempt < max_retries - 1:
                delay = base_delay * (2 ** attempt)
                st.write(f"&nbsp;&nbsp;&nbsp;↳ Waiting {delay} seconds before retry...")
                time.sleep(delay)
    
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

# Show template management info
with st.expander("📁 Template Management Info"):
    templates_dir = Path("templates")
    if templates_dir.exists():
        template_files = list(templates_dir.glob("*.html"))
        if template_files:
            st.success(f"Found {len(template_files)} saved templates:")
            for template_file in template_files:
                st.write(f"- {template_file.stem}")
        else:
            st.info("No templates found in the templates folder.")
    else:
        st.info("Templates folder will be created automatically when you save your first template.")
    
    st.write("**To add templates manually:** Place .html files in the `templates/` folder in your app directory.")

gmail_service, sheets_service = get_preauthorized_services()

if gmail_service and sheets_service:
    st.header("Step 1: Link Your Google Sheet")
    sheet_url = st.text_input("Paste the full URL of your Google Sheet here")

    if sheet_url:
        try:
            spreadsheet_id = re.search('/d/([a-zA-Z0-9-_]+)', sheet_url).group(1)
            
            # Get all sheet names first
            sheet_names = get_all_sheet_names(sheets_service, spreadsheet_id)
            
            if sheet_names:
                st.success(f"Found {len(sheet_names)} sheet(s) in the workbook")
                
                # Let user select which sheet to use
                selected_sheet = st.selectbox(
                    "Select which sheet to load:",
                    options=sheet_names,
                    index=0,
                    key="sheet_selector"
                )
                
                # Load the selected sheet
                df, headers, sheet_name = get_sheet_data(sheets_service, spreadsheet_id, selected_sheet)
            else:
                st.error("No sheets found in the workbook.")
                st.stop()
                
        except Exception:
            st.error("Invalid Google Sheet URL. Please paste the full URL from your browser's address bar.")
            st.stop()

        if not df.empty:
            st.success(f"Successfully loaded {len(df)} rows from sheet: '{sheet_name}'")
            st.dataframe(df.head())

            tab1, tab2 = st.tabs(["🚀 Send Initial Campaign", "✉️ Send Reminders / Replies"])

            with tab1:
                st.subheader("Send First-Time Emails")
                subject_input = st.text_input("Initial Email Subject", key="initial_subject")
                
                # Use template selector instead of file uploader
                html_template = template_selector_ui("initial")
                
                if st.button("Start Initial Campaign"):
                    if not all([html_template, subject_input]):
                        st.warning("Please provide a subject and select/upload an HTML template.")
                    else:
                        sent_emails_info = []
                        
                        # --- PHASE 1: SEND ALL EMAILS ---
                        with st.expander("Live Send Status", expanded=True):
                            st.write("--- Phase 1: Sending Emails ---")
                            for i, row in df.iterrows():
                                if pd.isna(row.get('email')) or not row.get('email'): 
                                    continue
                                st.write(f"Row {i+2}: Sending to **{row.get('email')}**...")
                                result = send_initial_email(gmail_service, row.get('email'), subject_input, html_template, row.to_dict())
                                if result:
                                    sent_emails_info.append({
                                        "row_index": i, 
                                        "temp_id": result['id'], 
                                        "thread_id": result['threadId'], 
                                        "subject": subject_input.format(**row.to_dict()),
                                        "email": row.get('email')
                                    })
                                    st.write(f"&nbsp;&nbsp;&nbsp;↳ Success: Email sent (Gmail ID: {result['id']}).")
                                else:
                                    st.error(f"&nbsp;&nbsp;&nbsp;↳ Failed to send email to {row.get('email')}.")
                        
                        # --- PHASE 2: WAIT & FETCH MESSAGE IDS WITH IMPROVED RETRY LOGIC ---
                        update_log = {}
                        if sent_emails_info:
                            with st.expander("Live Log Status", expanded=True):
                                st.write("\n--- Phase 2: Fetching Message IDs with Retry Logic ---")
                                st.write("⏳ Waiting 10 seconds for Gmail to process sent messages...")
                                time.sleep(10)  # Increased initial wait time
                                
                                for sent_item in sent_emails_info:
                                    i = sent_item["row_index"]
                                    st.write(f"Row {i+2}: Fetching Message-ID for {sent_item['email']}...")
                                    
                                    # Use the improved retry function
                                    msg_id_header = get_message_id_with_retry(
                                        gmail_service, 
                                        sent_item['temp_id']
                                    )
                                    
                                    update_log[i] = {
                                        "Timestamp": datetime.now(pytz.timezone("Asia/Kolkata")).strftime("%Y-%m-%d %H:%M:%S"), 
                                        "Status": "Sent", 
                                        "Subject": sent_item["subject"], 
                                        "Thread ID": sent_item["thread_id"], 
                                        "Message ID": msg_id_header
                                    }
                        
                        # --- PHASE 3: UPDATE GOOGLE SHEET RELIABLY ---
                        st.info("--- Phase 3: Updating Google Sheet with logs ---")
                        log_headers = ["Timestamp", "Status", "Subject", "Thread ID", "Message ID"]
                        for header in log_headers:
                            if header not in df.columns:
                                df[header] = ''  # Add new columns to the DataFrame if they don't exist
                        
                        # Update the DataFrame in memory with the new log data
                        for row_index, log_data in update_log.items():
                            for col_name, value in log_data.items():
                                df.loc[row_index, col_name] = value

                        try:
                            # Convert the entire updated DataFrame to a list of lists
                            update_values = [df.columns.values.tolist()] + df.values.tolist()
                            # Clear the sheet and write the entire updated data back in one go
                            sheets_service.spreadsheets().values().clear(
                                spreadsheetId=spreadsheet_id, range=selected_sheet
                            ).execute()
                            sheets_service.spreadsheets().values().update(
                                spreadsheetId=spreadsheet_id, range=f"{selected_sheet}!A1",
                                valueInputOption="USER_ENTERED", body={'values': update_values}
                            ).execute()
                            st.success("Google Sheet updated successfully!")
                            st.balloons()
                        except Exception as e:
                            st.error(f"Failed to update Google Sheet. Error: {e}")

            with tab2:
                st.subheader("Send a Follow-up or Reminder Email")
                st.info("This will send a threaded reply to contacts who have a 'Message ID' in the sheet.")
                
                # Use template selector instead of file uploader
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
                                st.warning("No contacts found with a valid 'Message ID' to reply to. Please run an initial campaign first.")
                            else:
                                with st.spinner(f"Sending reminders to {len(reply_df)} contacts..."):
                                    for i, row in reply_df.iterrows():
                                        send_reply_email(
                                            gmail_service,
                                            row.get('email'),
                                            row.get('Subject'),
                                            row.get('Thread ID'),
                                            row.get('Message ID'),
                                            reminder_template,
                                            row.to_dict()
                                        )
                                    st.success("Reminder campaign sent!")
                                    st.balloons()
else:
    st.error("Application is offline. Could not authenticate to Google.")