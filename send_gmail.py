from download_repair_photos import send_message, CLIENT_SECRET_FILE, SCOPES, API_NAME, API_VERSION
from Google import Create_Service

# Setup the Gmail service
service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

# Define your message here (this is just an example structure, you'd fill it with relevant data)
message_body = {
    'raw': 'your_encoded_message_here'
}

# Use the function
send_message(service, 'me', message_body)