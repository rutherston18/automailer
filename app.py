import streamlit as st
import pandas as pd
import base64
from email.message import EmailMessage

# Google Cloud & Auth Libraries
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# This scope must match the one used to generate your original token.json
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

@st.cache_resource
def get_preauthorized_gmail_service():
    """
    Authenticates using a refresh_token stored in Streamlit Secrets.
    This bypasses the need for any interactive user login.
    """
    try:
        # Load all necessary credentials from the single secrets section
        creds_data = st.secrets["preauthorized_account"]
        
        # Create a Credentials object directly from the secrets
        creds = Credentials.from_authorized_user_info(
            info={
                "refresh_token": creds_data["refresh_token"],
                "client_id": creds_data["client_id"],
                "client_secret": creds_data["client_secret"],
                "token_uri": "https://oauth2.googleapis.com/token" # Standard token URI
            },
            scopes=SCOPES
        )
        
        # The credentials object will handle refreshing the access token automatically
        # if it's expired, using the refresh_token.
        
        # Build and return the Gmail service object
        gmail_service = build("gmail", "v1", credentials=creds)
        st.success("Successfully authenticated with the pre-authorized account.")
        return gmail_service

    except Exception as e:
        st.error("Failed to authenticate with the pre-authorized account. Please double-check your Streamlit Secrets.")
        st.exception(e)
        return None

def send_email(service, to_email, subject, html_body_template, row_data):
    """Creates and sends a personalized HTML email."""
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
        send_message = service.users().messages().send(userId="me", body=create_message).execute()
        return send_message
    except HttpError as error:
        st.error(f"An API error occurred sending to {to_email}: {error}")
        return None
    except KeyError as e:
        st.error(f"Placeholder error for {to_email}: Missing column {e} in your CSV or Subject.")
        return None

# --- Main Streamlit UI ---
st.title("📧 Web-Based Email Campaign Tool")
st.info("This app sends emails immediately from a single, pre-authorized company account.")
st.markdown("---")

# Authenticate automatically on app load
gmail_service = get_preauthorized_gmail_service()

# Only show the UI if authentication was successful
if gmail_service:
    st.header("Step 1: Prepare Your Campaign")
    subject_input = st.text_input("Enter Email Subject (placeholders like {name} are okay)")
    uploaded_csv = st.file_uploader("Upload Contacts (CSV)", type=["csv"])
    uploaded_template = st.file_uploader("Upload Template (HTML)", type=["html"])
    
    st.markdown("---")

    # --- NEW PREVIEW SECTION ---
    if uploaded_csv is not None and uploaded_template is not None:
        st.header("Step 2: Preview Your Campaign")
        
        try:
            # Read data for preview
            df = pd.read_csv(uploaded_csv)
            html_template = uploaded_template.getvalue().decode("utf-8")
            
            st.subheader("CSV Data Preview (First 5 Rows)")
            st.dataframe(df.head())
            
            # Use the first row of data to generate a live preview
            if not df.empty:
                st.subheader("Live Email Preview")
                st.info("This is how the email will look for the first contact in your list.")
                
                first_row_data = df.iloc[0].to_dict()
                
                # Preview Subject
                preview_subject = subject_input.format(**first_row_data)
                st.text_input("Rendered Subject:", preview_subject, disabled=True)
                
                # Preview Body
                preview_html = html_template.format(**first_row_data)
                with st.container(border=True):
                    st.markdown(preview_html, unsafe_allow_html=True)
            else:
                st.warning("Your CSV file is empty. Cannot generate a preview.")

        except KeyError as e:
            st.error(f"⚠️ Placeholder Error: The placeholder {e} in your template or subject does not match any column in your CSV file. Please check your files.")
        except Exception as e:
            st.error(f"An error occurred while generating the preview: {e}")
    # --- END OF PREVIEW SECTION ---

    st.markdown("---")
    
    # --- SENDING SECTION ---
    st.header("Step 3: Send Email Campaign")
    if st.button("Send Email Campaign"):
        if not all([uploaded_csv, uploaded_template, subject_input]):
            st.warning("Please provide a subject, a CSV, and an HTML template before sending.")
        else:
            try:
                # We need to re-read the files in case the user has changed them
                # since the preview was generated.
                uploaded_csv.seek(0)
                df_send = pd.read_csv(uploaded_csv)
                
                uploaded_template.seek(0)
                html_template_send = uploaded_template.getvalue().decode("utf-8")

                total_emails = len(df_send)
                
                with st.spinner(f"Sending {total_emails} emails..."):
                    progress_bar = st.progress(0)
                    success_count = 0
                    
                    for i, row in df_send.iterrows():
                        row_data = row.to_dict()
                        recipient_email = row_data.get('email')

                        if not recipient_email or pd.isna(recipient_email):
                            st.warning(f"Skipping row {i+2} in CSV: No 'email' column found or value is empty.")
                            continue
                        
                        result = send_email(gmail_service, recipient_email, subject_input, html_template_send, row_data)
                        
                        if result:
                            success_count += 1
                        
                        progress_bar.progress((i + 1) / total_emails)
                
                st.success(f"Finished! Successfully sent {success_count} out of {total_emails} emails.")
                st.balloons()
                
            except Exception as e:
                st.error("An error occurred during file processing or sending.")
                st.exception(e)
else:
    st.error("Application is offline. Could not authenticate to Google. Please check the logs or secrets configuration.")
