# trying to get this gmail thing going we'll see 
from Google import Create_Service
import base64, os, datetime, pickle, time, tkinter
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
  
def download_attachments(service, user_id, msg_id, store_dir, desired_subject, days_ago):
    
    attachments_paths = []
    try:
        message = service.users().messages().get(userId=user_id, id=msg_id).execute()
        payload = message['payload']
        headers = payload['headers']
        subject = [h['value'] for h in headers if h['name'] == 'Subject'][0]

        received_at = int(message['internalDate']) / 1000  # Convert to seconds
        received_date = datetime.utcfromtimestamp(received_at).date()
        today = datetime.utcnow().date()
        threshold_date = today - timedelta(days_ago)
        
        print(received_date)
        print(threshold_date)

        if received_date < threshold_date:
            return []  # Skip processing if the email was not received today

        if desired_subject.lower() in subject.lower():
            parts = payload.get('parts', [])
            print(f"Number of parts in the email: {len(parts)}")
            print(f"Parts: {parts}")

            found_valid_attachment = False
            for part in parts:
                print('hello')
                if part.get('filename'):
                    print("something here")
                    filename = part['filename']  # Fix the filename extraction here
                    print(f"Extracted filename: {filename}")  # Add this line to print the extracted filename

                    attachment_id = part['body']['attachmentId']
                    attachment = service.users().messages().attachments().get(userId=user_id, messageId=msg_id, id=attachment_id).execute()
                    file_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))
                    os.makedirs(store_dir, exist_ok=True)
                    filepath = os.path.join(store_dir, filename)
                    with open(filepath, 'wb') as f:
                        f.write(file_data)

                    print(f"Attachment '{filename}' downloaded to: {filepath}")

                    attachments_paths.append(filepath)
                    print(f"Size of 'file_data': {len(file_data)}")  # Add this line to print the size of 'file_data'

                    # Additional debugging: Check the first few bytes of the 'file_data'
                    print(f"First few bytes of 'file_data': {file_data[:100]}")
                # elif 'parts' in part:
                #         print('hallelujah')
                #         for sub_part in part['parts']:
                #             print(f"Subpart: {sub_part}")
                #             if sub_part.get('filename'):
                #                 print("something")
                #                 # Extract sub-part information
                #                 sub_filename = sub_part['filename']
                #                 sub_attachment_id = sub_part['body']['attachmentId']
                                
                #                 # Download the sub-attachment using its ID
                #                 sub_attachment = service.users().messages().attachments().get(userId=user_id, messageId=msg_id, id=sub_attachment_id).execute()
                                
                #                 # Decode the attachment data
                #                 sub_file_data = base64.urlsafe_b64decode(sub_attachment['data'].encode('UTF-8'))

                #                 # Additional debugging: Check the size and first few bytes of 'sub_file_data'
                #                 print(f"Size of sub 'file_data': {len(sub_file_data)}")
                #                 print(f"First few bytes of sub 'file_data': {sub_file_data[:100]}")

                #                 # Check if the sub-file is a PDF
                #                 _, file_extension = os.path.splitext(sub_filename)
                #                 if file_extension.lower() == ".pdf":
                #                     # Save the sub-attachment to the specified 'store_dir'
                #                     os.makedirs(store_dir, exist_ok=True)
                #                     filepath = os.path.join(store_dir, sub_filename)
                #                     with open(filepath, 'wb') as f:
                #                         f.write(sub_file_data)

                #                     print(f"Sub-attachment '{sub_filename}' downloaded to: {filepath}")

                #                     attachments_paths.append(filepath)
                #                 else:
                #                     print(f"Skipped sub-attachment '{sub_filename}' because it is not a PDF or is empty.")
                #             else:
                #                 print("Skipping a sub-part with no 'filename' attribute.")

                #         print(f"Attachment '{filename}' downloaded to: {filepath}")

                #         attachments_paths.append(filepath)
                #         found_valid_attachment = True

            if not found_valid_attachment:
                print("No valid attachments found in the email.")
        else:
            print(f"Skipping email with subject: {subject}")
    except Exception as e:
        print('An error occurred: %s' % e)

    return attachments_paths

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
store_directory = r"C:\Users\Michael\Desktop\python-work\Invoices"

today = datetime.utcnow().date()
start_of_today = datetime.combine(today, time.min)
start_of_tomorrow = start_of_today + timedelta(days=1)

start_of_today_timestamp = int(start_of_today.timestamp()) * 1000
start_of_tomorrow_timestamp = int(start_of_tomorrow.timestamp()) * 1000



query_list = [
    "from:carolyn@profastening.net subject:'Invoice'", 
    "from:Sales@gemcoroofingsupply.com subject:'Invoice'", 
    "from:april@sheetmetalsupplyltd.com subject:'Invoice'",
    "amy@profastening.net subject:'Invoice'",
    "dawn@sheetmetalsupplyltd.com subject:'Invoice'",
    "lia@stevensoncrane.com subject:'invoice'"
    ]

# Tells how many days back to check
desired_date = simpledialog.askinteger("Desired Dates", "How many days into the past do you want to select emails?")

print_status = messagebox.askyesno("Confirmation", "Do you want to print matching invoices?")

for query in query_list:
    print(query)
    try:
        response = service.users().messages().list(userId=user_email, q=query).execute()
        messages = response.get('messages', [])
        all_attachments = []
        for message in messages:
            msg_id = message['id']
            attachments = download_attachments(service, user_email, msg_id, store_directory, "Invoice", desired_date)
            if attachments:
                all_attachments.extend(attachments)
            print(attachments)
            
        sleep(1)
        if print_status:
            for attachment in all_attachments:
                print_file_with_ghostscript(attachment)
    except Exception as e:
        print('An error occurred: %s' % e)


def main():
    # ... (Your existing code for authentication and obtaining the service object)
    service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

    user_email = "wbrroof@gmail.com"  # Replace with the email address you want to fetch messages from
    store_directory = r"C:\Users\Michael\Desktop\python-work\Invoices"

    today = datetime.utcnow().date()
    start_of_today = datetime.combine(today, time.min)
    start_of_tomorrow = start_of_today + timedelta(days=1)

    start_of_today_timestamp = int(start_of_today.timestamp()) * 1000
    start_of_tomorrow_timestamp = int(start_of_tomorrow.timestamp()) * 1000

    query_list = [
        "from:carolyn@profastening.net subject:'Invoice'",
        "from:Sales@gemcoroofingsupply.com subject:'Invoice'",
        "from:april@sheetmetalsupplyltd.com subject:'Invoice'"
    ]

    print_status = messagebox.askyesno("Confirmation", "Do you want to print matching invoices?")

    for query in query_list:
        try:
            response = service.users().messages().list(userId=user_email, q=query).execute()
            messages = response.get('messages', [])
            all_attachments = []
            for message in messages:
                msg_id = message['id']
                attachments = download_attachments(service, user_email, msg_id, store_directory, "Invoice")
                if attachments:
                    all_attachments.extend(attachments)
                print(attachments)

            sleep(1)
            if print_status:
                for attachment in all_attachments:
                    print_file_with_ghostscript(attachment)
        except Exception as e:
            print('An error occurred: %s' % e)

# if __name__ == "__main__":
#     main()
# recipient_email = "throwod@gmail.com"  # Replace with the recipient's email address
# message_subject = "Testing Gmail API"
# message_text = "Bruh, this is a message sent via the Gmail API! Let's go!"

# message = create_message(user_email, recipient_email, message_subject, message_text)
# if message:
#     send_message(service, user_email, message)