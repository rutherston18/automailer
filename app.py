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
    Handles Google Authentication for a web-based Streamlit app using the
    correct authorization code flow, which is required for cloud environments.
    """
    creds_key = "google_creds"
    
    # If we already have valid credentials in the session, build and return the service
    if creds_key in st.session_state and st.session_state[creds_key].valid:
        return build("gmail", "v1", credentials=st.session_state[creds_key])

    # --- Step 1: Get client configuration from secrets ---
    try:
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
        # The first URI in your list is used as the redirect URI.
        # This MUST be the URL of your deployed Streamlit app.
        # It must also be listed in the "Authorized redirect URIs" in your Google Cloud Console.
        redirect_uri = client_config["web"]["redirect_uris"][0]
        
    except Exception as e:
        st.error("Your secrets are not configured correctly in Streamlit Cloud.")
        st.exception(e)
        return None

    # --- Step 2: Handle the return from Google with the authorization code ---
    if "code" in st.query_params:
        # User has returned from Google's login page with a code.
        
        # Verify the state to prevent CSRF attacks
        if st.session_state.get("oauth_state") != st.query_params.get("state"):
            st.error("State mismatch. Authentication failed. Please try logging in again.")
            return None
        
        try:
            flow = InstalledAppFlow.from_client_config(client_config, scopes=SCOPES, state=st.session_state["oauth_state"])
            flow.redirect_uri = redirect_uri
            
            # Use the authorization code to fetch the token
            flow.fetch_token(code=st.query_params["code"])
            
            st.session_state[creds_key] = flow.credentials
            
            # Clean up URL and state, then rerun the script for a clean UI
            del st.session_state["oauth_state"]
            st.query_params.clear()
            st.rerun()

        except Exception as e:
            st.error("Failed to fetch authorization token.")
            st.exception(e)
            return None
            
    else:
        # --- Step 3: If no code, start the authorization flow by generating a login link ---
        flow = InstalledAppFlow.from_client_config(client_config, scopes=SCOPES)
        flow.redirect_uri = redirect_uri

        authorization_url, state = flow.authorization_url(
            access_type='offline',
            include_granted_scopes='true'
        )
        
        # Store state in session to verify after redirect
        st.session_state["oauth_state"] = state
        
        return None # Indicate that the login process is not complete

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
st.info("This app sends emails immediately from your logged-in Google account.")

# Attempt to authenticate at the start of the script run
gmail_service = authenticate_gmail()

if gmail_service:
    # --- USER IS LOGGED IN AND AUTHENTICATED ---
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
                        
                        progress_bar.progress((i + 1) / total_emails)
                
                st.success(f"Finished! Successfully sent {success_count} out of {total_emails} emails.")
                st.balloons()
            except Exception as e:
                st.error("An error occurred during file processing or sending.")
                st.exception(e)

    st.markdown("---")
    if st.button("Logout"):
        del st.session_state.google_creds
        if "oauth_state" in st.session_state:
            del st.session_state["oauth_state"]
        st.rerun()

else:
    # --- USER IS NOT LOGGED IN ---
    st.header("Step 1: Log in with Google")
    st.write("You need to authorize this application to send emails on your behalf.")
    # The login link is now displayed by the authenticate_gmail() function
    st.info("Click the 'Login with Google' link above to proceed. A new tab may open.")

