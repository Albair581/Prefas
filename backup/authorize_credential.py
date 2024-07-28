from flask import Flask, redirect, url_for, session, request
from google_auth_oauthlib.flow import Flow
import os
import json

app = Flask(__name__)
app.secret_key = 'your_secret_key'
os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"  # Only for local development (http)

CLIENT_SECRETS_FILE = "client_secret.json"
SCOPES = ['https://www.googleapis.com/auth/gmail.send']
REDIRECT_URI = "http://localhost:5000/callback"
CREDENTIALS_FILE = "stored_credentials.json"

@app.route('/')
def index():
    return 'Welcome to the OAuth 2.0 demo! <a href="/authorize">Authorize</a>'

@app.route('/authorize')
def authorize():
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE, scopes=SCOPES, redirect_uri=REDIRECT_URI)
    authorization_url, state = flow.authorization_url(
        access_type='offline', include_granted_scopes='true')
    session['state'] = state
    return redirect(authorization_url)

@app.route('/callback')
def callback():
    state = session['state']
    flow = Flow.from_client_secrets_file(
        CLIENT_SECRETS_FILE, scopes=SCOPES, state=state, redirect_uri=REDIRECT_URI)
    flow.fetch_token(authorization_response=request.url)
    credentials = flow.credentials
    session['credentials'] = credentials_to_dict(credentials)

    # Store credentials in a file
    with open(CREDENTIALS_FILE, 'w') as f:
        json.dump(session['credentials'], f)

    return 'Credentials have been stored successfully.'

def credentials_to_dict(credentials):
    return {
        'token': credentials.token,
        'refresh_token': credentials.refresh_token,
        'token_uri': credentials.token_uri,
        'client_id': credentials.client_id,
        'client_secret': credentials.client_secret,
        'scopes': credentials.scopes
    }

if __name__ == '__main__':
    app.run('localhost', 5000, debug=True)
