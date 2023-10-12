from flask import Flask, request, jsonify
from googleapiclient.discovery import build
from oauth2client import OAuth2Credentials
import google.auth
import google_auth_oauthlib
import google.auth.transport.requests
from google.oauth2 import service_account
import pickle
import os
import datetime
from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from google.auth.transport.requests import Request


def Create_Service(client_secret_file, api_name, api_version, *scopes):
    print(client_secret_file, api_name, api_version, scopes, sep='-')
    CLIENT_SECRET_FILE = client_secret_file
    API_SERVICE_NAME = api_name
    API_VERSION = api_version
    SCOPES = [scope for scope in scopes[0]]
    print(SCOPES)

    cred = None

    pickle_file = f'token_{API_SERVICE_NAME}_{API_VERSION}.pickle'
    # print(pickle_file)

    if os.path.exists(pickle_file):
        with open(pickle_file, 'rb') as token:
            cred = pickle.load(token)

    if not cred or not cred.valid:
        try:
            if cred and cred.expired and cred.refresh_token:
                cred.refresh(Request())
        except Exception as e:
            print("Error refreshing token:", e)
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            cred = flow.run_local_server()

        with open(pickle_file, 'wb') as token:
            pickle.dump(cred, token)
    else:
            # Diagnostic print statements if credentials are loaded from pickle file
            print(f"Access token (from pickle file): {cred.token}")
            print(f"Refresh token (from pickle file): {cred.refresh_token}")
    try:
        service = build(API_SERVICE_NAME, API_VERSION, credentials=cred)
        print(API_SERVICE_NAME, 'service created successfully')
        return service
    except Exception as e:
        print('Unable to connect.')
        print(e)
        return None

def convert_to_RFC_datetime(year=1900, month=1, day=1, hour=0, minute=0):
    dt = datetime.datetime(year, month, day, hour, minute, 0).isoformat() + 'Z'
    return dt

# Load your OAuth2 credentials and create a Gmail API service client

CLIENT_SECRET_FILE = 'wbrcredentials.json'  # Replace with the path to your credentials.json file
API_NAME = 'gmail'
API_VERSION = 'v1'
SCOPES = ['https://www.googleapis.com/auth/gmail.send', 'https://www.googleapis.com/auth/gmail.modify', 'https://www.googleapis.com/auth/gmail.readonly']

gmail_service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)


# Set up a watch on the mailbox
request = {
    'labelIds': ['INBOX'],  # You can specify filters, labels, etc.
    'topicName': 'projects/gmail-project-394016/topics/Base_Topic'
}

result = gmail_service.users().watch(userId='me', body=request).execute()
print(result)



app = Flask(__name__)

@app.route('/notifications', methods=['POST'])
def gmail_notification():
    # Extract information from the Pub/Sub message
    envelope = request.get_json()
    print(f"Received envelope: {envelope}")

    # ... perform your custom logic ...

    # Always return a 200 response to acknowledge receipt
    return jsonify(success=True), 200
