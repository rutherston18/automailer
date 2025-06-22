import streamlit as st
import pandas as pd
import base64
from email.message import EmailMessage

# Google Cloud & Auth Libraries
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# This scope allows the app to send emails on the logged-in user's behalf.
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

def authenticate_gmail():
    """
    Handles Google Authentication for each user session using st.secrets.
    Stores credentials in st.session_state instead of a token.json file.
    """
    # Check if credentials are already in the session state and are valid
    if "google_creds" in st.session_state and st.session_state.google_creds.valid:
        # If yes, build and return the service object
        return build("gmail", "v1", credentials=st.session_state.google_creds)

    # If not, or if they are invalid, start the OAuth flow
    try:
        # --- THIS IS THE CORRECTED PART FOR READING SECRETS ---
        # Re-create the client_config dictionary from the individual secrets.
        # The key 'google_credentials_oauth' must match the section header in your secrets.
        client_config = {
            "web": {
                "client_id": st.secrets["google_credentials_oauth"]["client_id"],
                "project_id": st.secrets["google_credentials_oauth"]["project_id"],
                "auth_uri": st.secrets["google_credentials_oauth"]["auth_uri"],
                "token_uri": st.secrets["google_credentials_oauth"]["token_uri"],
                "auth_provider_x509_cert_url": st.secrets["google_credentials_oauth"]["auth_provider_x509_cert_url"],
                "client_secret": st.secrets["google_credentials_oauth"]["client_secret"],
                "redirect_uris": st.secrets["google_credentials_oauth"]["redirect_uris"]
            }
        }
        # ---------------------------------------------------------
        
        flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
        
        # This will run a local server to get the authorization code.
        # Streamlit Cloud handles the redirect URI automatically.
        creds = flow.run_local_server(port=0)
        
        # Save the valid credentials in the user's session state for reuse
        st.session_state.google_creds = creds
        
        # Refresh the page to show the logged-in state
        st.experimental_rerun()
        
    except Exception as e:
        st.error("Authentication failed. Please ensure secrets are configured correctly in Streamlit Cloud.")
        st.exception(e) # This shows the full error for debugging
        return None

def send_email(service, to_email, subject, html_body_template, row_data):
    """Creates and sends a personalized HTML email."""
    try:
        # Use .format(**row_data) to fill in all placeholders from the CSV row
        final_html_body = html_body_template.format(**row_data)
        final_subject = subject.format(**row_data)

        message = EmailMessage()
        # Set the HTML content
        message.add_alternative(final_html_body, subtype='html')
        
        message["To"] = to_email
        message["From"] = "me" # "me" refers to the authenticated user
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
st.info("This app sends emails immediately from your logged-in Google account.")

# Check if the user is logged in
if "google_creds" not in st.session_state or not st.session_state.google_creds.valid:
    st.header("Step 1: Log in with Google")
    st.write("You need to authorize this application to send emails on your behalf.")
    if st.button("Login with Google"):
        # This will trigger the authentication flow and rerun the script on success
        authenticate_gmail()
else:
    # --- USER IS LOGGED IN ---
    # Display logged-in user's email
    user_email = st.session_state.google_creds.id_token.get('email', 'Unknown User')
    st.success(f"Logged in as: {user_email}")
    st.markdown("---")
    
    st.header("Step 2: Prepare and Send Your Campaign")

    subject_input = st.text_input("Enter Email Subject", "Following up on our conversation")
    uploaded_csv = st.file_uploader("Upload Contacts (CSV)", type=["csv"])
    uploaded_template = st.file_uploader("Upload Template (HTML)", type=["html"])

    if st.button("Send Email Campaign"):
        if not all([uploaded_csv, uploaded_template, subject_input]):
            st.warning("Please provide a subject, a CSV, and an HTML template.")
        else:
            # Get the authenticated Gmail service. Should already be valid.
            gmail_service = authenticate_gmail()
            if gmail_service:
                try:
                    df = pd.read_csv(uploaded_csv)
                    html_template = uploaded_template.getvalue().decode("utf-8")
                    total_emails = len(df)
                    
                    with st.spinner(f"Sending {total_emails} emails..."):
                        progress_bar = st.progress(0)
                        success_count = 0
                        
                        for i, row in df.iterrows():
                            row_data = row.to_dict()
                            recipient_email = row_data.get('email')

                            if not recipient_email:
                                st.warning(f"Skipping row {i+2} in CSV: No 'email' column found or value is empty.")
                                continue
                            
                            result = send_email(gmail_service, recipient_email, subject_input, html_template, row_data)
                            
                            if result:
                                success_count += 1
                            
                            # Update progress bar
                            progress_bar.progress((i + 1) / total_emails)
                    
                    st.success(f"Finished! Successfully sent {success_count} out of {total_emails} emails.")
                    st.balloons()
                except Exception as e:
                    st.error("An error occurred during file processing or sending.")
                    st.exception(e)

    st.markdown("---")
    if st.button("Logout"):
        # Clear the credentials from the session state and rerun
        del st.session_state.google_creds
        st.experimental_rerun()
