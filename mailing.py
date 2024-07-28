import json
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
import googleapiclient.discovery
from email.mime.text import MIMEText
import base64

# Path to the stored credentials file
CREDENTIALS_FILE = "stored_credentials.json"

def send_email(subject, recipient, message_text):
    # Load stored credentials
    with open(CREDENTIALS_FILE, 'r') as f:
        stored_credentials = json.load(f)

    credentials = Credentials(**stored_credentials)
    
    # Refresh the token if it has expired
    if credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())

    # Create the Gmail API service
    service = googleapiclient.discovery.build('gmail', 'v1', credentials=credentials)

    # Create the email message
    message = MIMEText(message_text)
    message['to'] = recipient
    message['subject'] = subject
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode()
    message = {'raw': raw_message}

    # Send the email
    sent_message = service.users().messages().send(userId='me', body=message).execute()

    # Update stored credentials if refreshed
    with open(CREDENTIALS_FILE, 'w') as f:
        json.dump(credentials_to_dict(credentials), f)

    return str(sent_message["id"])

def credentials_to_dict(credentials):
    return {
        'token': credentials.token,
        'refresh_token': credentials.refresh_token,
        'token_uri': credentials.token_uri,
        'client_id': credentials.client_id,
        'client_secret': credentials.client_secret,
        'scopes': credentials.scopes
    }
