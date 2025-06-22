from google_auth_oauthlib.flow import Flow

# This scope must match what's in your app.py
SCOPES = [
    "https://www.googleapis.com/auth/gmail.send",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.modify" # <-- ADD THIS LINE
]
# Point this to the credentials.json file you downloaded for your "Web application"
# We are now using the base `Flow` object for more control.
flow = Flow.from_client_secrets_file(
    'credentials.json',
    scopes=SCOPES,
    # This redirect_uri must exactly match one of the URIs in the Google Console
    redirect_uri='http://localhost:8080'
)

# Step 1: Generate the authorization URL that prompts for consent every time
auth_url, _ = flow.authorization_url(prompt='consent')

print('--- MANUAL TOKEN GENERATION ---')
print('\nStep 1: Please go to this URL in your browser and log in:')
print(auth_url)
print('\n---')
print('Step 2: After you authorize, your browser will show a "This site can’t be reached" error page.')
print('         THIS IS EXPECTED.')
print('\nStep 3: COPY THE ENTIRE URL from your browser\'s address bar.')
print('         It will look like: http://localhost:8080/?state=...&code=4/0A...&scope=...')
print('\n---')

# Step 4: Get the full redirect URL from the user
redirect_response_url = input('Step 4: Paste the full URL you copied here and press Enter: ')

# Step 5: The script will now exchange the code for a token
try:
    flow.fetch_token(authorization_response=redirect_response_url)
    
    # Step 6: Save the credentials
    credentials = flow.credentials
    with open('token.json', 'w') as f:
        f.write(credentials.to_json())

    print("\nSUCCESS! A new token.json file has been created.")
    print("You can now copy the refresh_token from this new file into your Streamlit Secrets.")

except Exception as e:
    print(f"\nAn error occurred: {e}")
    print("Please make sure you copied the entire URL, starting with 'http://localhost:8080/'.")