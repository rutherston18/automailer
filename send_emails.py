import os.path
import base64
from email.message import EmailMessage

import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

def authenticate_gmail():
    """Shows basic usage of the Gmail API.
    Logs the user in and returns the service object.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())
            
    return build("gmail", "v1", credentials=creds)

def create_and_send_email(service, to_email, name, company, product):
    """Creates and sends a personalized email."""
    
    # --- YOUR EMAIL TEMPLATE ---
    subject = f"Following up regarding {company}"
    
    body = f"""
    Hi {name},

    This is a follow-up email regarding our services at {company}.
    We believe that our premier product, the '{product}', could be a great fit for your team.

    Thank you for your time.

    Best regards,
    {name}
    """
    # -------------------------

    try:
        message = EmailMessage()
        message.set_content(body)
        message["To"] = to_email
        message["From"] = "me"  # 'me' refers to the authenticated user
        message["Subject"] = subject

        # Encode the message in a URL-safe base64 format
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

        create_message = {"raw": encoded_message}
        
        # Send the email
        send_message = (
            service.users().messages().send(userId="me", body=create_message).execute()
        )
        print(f"Message Id: {send_message['id']} sent to {to_email}")
    except HttpError as error:
        print(f"An error occurred: {error}")
        send_message = None
    return send_message


def main():
    """Main function to run the mail merge."""
    
    # 1. Authenticate with Gmail
    service = authenticate_gmail()
    
    # 2. Read data from the CSV file
    try:
        df = pd.read_csv("data.csv")
    except FileNotFoundError:
        print("Error: data.csv not found. Please create it in the same directory.")
        return

    # 3. Loop through each person and send an email
    for index, row in df.iterrows():
        # Get data from the current row
        recipient_email = row['email']
        recipient_name = row['name']
        recipient_company = row['company']
        recipient_product = row['product']
        
        print(f"Preparing to send email to {recipient_name} at {recipient_email}...")
        
        # Create and send the email
        create_and_send_email(service, recipient_email, recipient_name, recipient_company, recipient_product)

if __name__ == "__main__":
    main()