from flask import Flask, request, jsonify
from googleapiclient.discovery import build
import google.auth
import google_auth_oauthlib
import google.auth.transport.requests
from google.oauth2 import service_account
import pickle, os, datetime
from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from google.auth.transport.requests import Request
import base64, json
from photos_timesheets import download_repair_photos, download_attachments, save_info_as_pdf, print_file_with_ghostscript, save_info_with_photos, sanitize_subject, wrap_text, mark_as_read


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

watch_request = {
    'labelIds': ['INBOX'],  # You can specify filters, labels, etc.
    'topicName': 'projects/gmail-project-394016/topics/Base_Topic'
}

result = gmail_service.users().watch(userId='me', body=watch_request).execute()

print(result)



app = Flask(__name__)

@app.route('/notifications', methods=['POST'])
def gmail_notification():
    # Extract information from the Pub/Sub message
    envelope = request.get_json()
    print(f"Received envelope: {envelope}")

    # decode the incoming data
    decoded_data = base64.b64decode(envelope['message']['data']).decode('utf-8')
    # print(decoded_data)
    # headers = request.headers
    # print(headers)
    decoded_data_json = json.loads(decoded_data)
    history_id = decoded_data_json.get('historyId')
    print(f"History ID: {history_id}")

    # Get changes since the last known historyId
    response = gmail_service.users().history().list(userId='me', startHistoryId=history_id).execute()
    
    changes = response.get('history', [])
    print(f"Changes received: {changes}")
    
    # there's nothing in these lists
    
    for change in changes:
        # Get message IDs from the change
        for message in change.get('messages', []):
            msg_id = message['id']

            msg = gmail_service.users().messages().get(userId='me', id=msg_id).execute()
            print(f"Message details: {msg}")

            for header in msg['payload']['headers']:
                print(header['name'], ":", header['value'])


            subject_header = next((header for header in msg['payload']['headers'] if header['name'] == 'Subject'), None)
            subject = subject_header['value'] if subject_header else 'no subject'
            payload = msg['payload']
            headers = payload['headers']
            sender = [h['value'] for h in headers if h['name'] == 'From'][0]  # Retrieve sender
            if sender == "throwod@gmail.com":
                print("we got him") 
            # print(subject, sender)

    # ... perform your custom logic ...

    # Always return a 200 response to acknowledge receipt
    return jsonify(success=True), 200

if __name__ == "__main__":
    app.run(debug=True)