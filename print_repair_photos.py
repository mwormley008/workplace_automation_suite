# trying to get this gmail thing going we'll see 
from Google import Create_Service
import base64, os, datetime, pickle, time, tkinter, re
from time import sleep
import win32print, subprocess
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta, time

from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askopenfilename, Frame, Button


CLIENT_SECRET_FILE = 'wbrcredentials.json'  # Replace with the path to your credentials.json file
API_NAME = 'gmail'
API_VERSION = 'v1'
SCOPES = ['https://www.googleapis.com/auth/gmail.send', 'https://www.googleapis.com/auth/gmail.modify', 'https://www.googleapis.com/auth/gmail.readonly']

def create_message(sender, to, subject, message_text):
  message = MIMEText(message_text)
  message['to'] = to
  message['from'] = sender
  message['subject'] = subject
  raw_message = base64.urlsafe_b64encode(message.as_string().encode("utf-8"))
  return {
    'raw': raw_message.decode("utf-8")
  }

def create_draft(service, user_id, message_body):
  try:
    message = {'message': message_body}
    draft = service.users().drafts().create(userId=user_id, body=message).execute()
    print("Draft id: %s\nDraft message: %s" % (draft['id'], draft['message']))
    return draft
  except Exception as e:
    print('An error occurred: %s' % e)
    return None
  
def send_message(service, user_id, message):
  try:
    message = service.users().messages().send(userId=user_id, body=message).execute()
    print('Message Id: %s' % message['id'])
    return message
  except Exception as e:
    print('An error occurred: %s' % e)
    return None
  
def download_attachments(service, user_id, msg_id, store_dir, desired_sender, days_ago):
    attachments_paths = []
    
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()
        payload = message['payload']
        headers = payload['headers']
        sender = [h['value'] for h in headers if h['name'] == 'From'][0]  # Retrieve sender
        subject = [h['value'] for h in headers if h['name'] == 'Subject'][0]  # Retrieve subject
        safe_subject = "".join([c for c in subject if c.isalpha() or c.isdigit() or c==' ']).rstrip()  # Cleaned subject line for directory


        # Extract the email address from sender
        sender_email = re.search(r'<(.*)>', sender)
        if sender_email:
            sender = sender_email.group(1)

        received_at = int(message['internalDate']) / 1000  # Convert to seconds
        received_date = datetime.utcfromtimestamp(received_at).date()
        today = datetime.utcnow().date()
        threshold_date = today - timedelta(days_ago)

        if received_date < threshold_date:
            return  # Skip processing if the email was not received today

# Retrieve the body of the email
        if 'parts' in payload:
            parts = payload.get('parts', [])
            for part in parts:
                mimeType = part.get('mimeType')
                body = part.get('body', {})
                data = body.get('data')
                if part.get('parts'):
                    # If the email part has parts inside it, recursively extract the text from it.
                    inner_parts = part.get('parts')
                    for inner_part in inner_parts:
                        inner_body = inner_part.get('body')
                        data = inner_body.get('data')
                        if inner_part.get('mimeType') == "text/plain":
                            # Decode the plain text body
                            text = base64.urlsafe_b64decode(data).decode()
                if mimeType == "text/plain":
                    # Decode the plain text body
                    text = base64.urlsafe_b64decode(data).decode()
        else:
            if payload['mimeType'] == "text/plain":
                text = base64.urlsafe_b64decode(payload['body']['data']).decode()
            else:
                text = ""

        if desired_sender == sender:
            parts = payload.get('parts', [])
            for part in parts:
                if part.get('filename'):
                    filename = part['filename']
                    attachment_id = part['body']['attachmentId']
                    attachment = service.users().messages().attachments().get(userId=user_id, messageId=msg_id, id=attachment_id).execute()
                    file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))
                    subject = [h['value'] for h in headers if h['name'] == 'Subject'][0]  # Retrieve subject
                    safe_subject = "".join([c for c in subject if c.isalpha() or c.isdigit() or c==' ']).rstrip()
                    unique_email_dir = os.path.join(store_dir, safe_subject)

                    
                    os.makedirs(unique_email_dir, exist_ok=True)

                    filepath = os.path.join(unique_email_dir, filename)  # Use original filename
                    with open(filepath, 'wb') as f:
                        f.write(file_data)

                    print(f"Attachment '{filename}' downloaded to: {filepath}")

                    attachments_paths.append(filepath)
            
            details = f"Subject: {subject}\n\nFrom: {sender}\n\nReceived Date: {received_date}\n\nBody:{text}"
            details_file = os.path.join(unique_email_dir, "00000000details.txt")
            with open(details_file, 'w') as f:
                f.write(details)

        return attachments_paths

    except Exception as error:
        print(f"An error occurred: {error}")



def print_file(filepath):
    printer_name = win32print.GetDefaultPrinter()
    if not printer_name:
        print("No default printer found. Please set a default printer.")
        return

    try:
        with open(filepath, 'rb') as f:
            data = f.read()

        hPrinter = win32print.OpenPrinter(printer_name)
        try:
            hJob = win32print.StartDocPrinter(hPrinter, 1, ("Attachment", None, "RAW"))
            try:
                win32print.StartPagePrinter(hPrinter)
                win32print.WritePrinter(hPrinter, data)
                win32print.EndPagePrinter(hPrinter)
            finally:
                win32print.EndDocPrinter(hPrinter)
        finally:
            win32print.ClosePrinter(hPrinter)

        print(f"Attachment '{filepath}' printed successfully.")
    except Exception as e:
        print(f"An error occurred while printing: {e}")

def print_file_with_ghostscript(filepath):
    ghostscript_path = r"C:\Program Files\gs\gs10.01.2\bin\gswin64c.exe"  # Replace with the path to Ghostscript executable
    printer_name = win32print.GetDefaultPrinter()

    if not os.path.exists(filepath):
        print(f"File '{filepath}' does not exist.")
        return

    command = [
        ghostscript_path,
        "-dNOPAUSE",
        "-dBATCH",
        "-dPrinted",
        f"-sDEVICE=mswinpr2",  # Use the Windows printer device
        f"-sOutputFile=%printer%{printer_name}",
        filepath
    ]

    try:
        subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
        print(f"File '{filepath}' printed successfully.")
    except Exception as e:
        print(f"An error occurred while printing: {e}")

service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
user_email = "wbrroof@gmail.com"  # Replace with the email address you want to send the message from
store_directory = r"C:\Users\Michael\Desktop\python-work\repair_photos"


# Tells how many days back to check
desired_date = simpledialog.askinteger("Desired Dates", "How many days into the past do you want to select emails?")


today = datetime.utcnow().date()
start_of_today = datetime.combine(today, time.min)
start_of_tomorrow = start_of_today + timedelta(days=1)

start_of_today_timestamp = int(start_of_today.timestamp()) * 1000
start_of_tomorrow_timestamp = int(start_of_tomorrow.timestamp()) * 1000


email_addresses = ["oblivion969.dm@gmail.com"]  # List of email addresses

# print_status = messagebox.askyesno("Confirmation", "Do you want to print matching invoices?")

for email_address in email_addresses:
    
    query = f"from:{email_address}"
    try:
        response = service.users().messages().list(userId=user_email, q=query).execute()
        
        messages = response.get('messages', [])
        all_attachments = []
        for message in messages:
            
            msg_id = message['id']
            attachments = download_attachments(service, user_email, msg_id, store_directory, email_address, desired_date)
            if attachments:
                all_attachments.extend(attachments)
        sleep(1)
        # if print_status:
        #     for attachment in all_attachments:
        #         print_file_with_ghostscript(attachment)
    except Exception as e:
        print('An error occurred: %s' % e)


# recipient_email = "throwod@gmail.com"  # Replace with the recipient's email address
# message_subject = "Testing Gmail API"
# message_text = "Bruh, this is a message sent via the Gmail API! Let's go!"

# message = create_message(user_email, recipient_email, message_subject, message_text)
# if message:
#     send_message(service, user_email, message)