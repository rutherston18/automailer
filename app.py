import streamlit as st
import os.path
import base64
from email.message import EmailMessage
import pandas as pd

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# --- Core Gmail Functions ---

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

@st.cache_resource
def authenticate_gmail():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    
    st.success("Authentication Successful!")
    return build("gmail", "v1", credentials=creds)

# CHANGED: This function now sends HTML emails
def send_email(service, to_email, subject, html_body_template, row_data):
    """Creates and sends a personalized HTML email."""
    try:
        final_html_body = html_body_template.format(**row_data)
        final_subject = subject.format(**row_data)

        message = EmailMessage()
        # SETS THE HTML CONTENT. This is the main change.
        message.add_alternative(final_html_body, subtype='html')
        
        message["To"] = to_email
        message["From"] = "me"
        message["Subject"] = final_subject

        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
        create_message = {"raw": encoded_message}
        
        send_message = service.users().messages().send(userId="me", body=create_message).execute()
        return send_message
    except HttpError as error:
        st.error(f"An error occurred sending to {to_email}: {error}")
        return None
    except KeyError as e:
        st.error(f"Placeholder error for {to_email}: Missing column {e} in your CSV or Subject.")
        return None


# --- Streamlit User Interface ---

st.title("CR MAIL SCENES")
st.write("This app sends personalized HTML emails using a CSV file and an HTML template.")

# --- File Uploaders and Subject Input ---
st.header("1. Provide Your Data")

# CHANGED: A dedicated field for the subject line
subject_input = st.text_input("Enter Email Subject (you can use placeholders like {name})", "Following up regarding {company}")

# CHANGED: Uploader now accepts .html files
uploaded_csv = st.file_uploader("Upload your contact list (CSV)", type=["csv"])
uploaded_template = st.file_uploader("Upload your email template (HTML)", type=["html"])

# --- Main App Logic ---
if uploaded_csv is not None and uploaded_template is not None:
    
    st.header("2. Preview Your Data")
    
    try:
        df = pd.read_csv(uploaded_csv)
        st.dataframe(df.head())
        
        html_template = uploaded_template.getvalue().decode("utf-8")
        
        st.subheader("HTML Template Preview:")
        # Use st.markdown to render the HTML preview
        st.markdown(html_template.format(**df.iloc[0].to_dict()), unsafe_allow_html=True)

    except Exception as e:
        st.error(f"Error reading files or generating preview: {e}")
        st.stop()

    st.header("3. Authenticate and Send")
    st.write("Click the button below to authenticate and start sending.")

    if st.button("Start Sending Emails"):
        with st.spinner("Authenticating with Gmail..."):
            service = authenticate_gmail()

        if service:
            total_emails = len(df)
            st.info(f"Starting to send {total_emails} emails...")
            
            progress_bar = st.progress(0)
            success_count = 0

            for i, row in df.iterrows():
                row_data = row.to_dict()
                recipient_email = row_data.get('email')

                if not recipient_email:
                    st.warning(f"Skipping row {i+1}: No 'email' column found.")
                    continue
                
                result = send_email(service, recipient_email, subject_input, html_template, row_data)
                
                if result:
                    success_count += 1
                
                progress_bar.progress((i + 1) / total_emails)
            
            st.success(f"Finished! Successfully sent {success_count} out of {total_emails} emails.")