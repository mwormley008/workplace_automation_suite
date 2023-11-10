# trying to get this gmail thing going we'll see 
# Also that problem with downloading invoices only occurs from April from SMS, which I don't have time to fix right now but 
# maybe was already solved 

"""
The totem invoices are structured kind of weirdly, so I need special logic for it.
jillian.schoedel@industrialandwholesalelumber.com subject: 'Invoice
2023-10-26
2023-10-24
Number of parts in the email: 2
Parts: [{'partId': '0', 'mimeType': 'multipart/alternative', 'filename': '', 'headers': [{'name': 'Content-Type', 'value': 'multipart/alternative; boundary=--boundary_105_8a0c0c4f-eae8-4679-873f-7b917e116be6'}], 'body': {'size': 0}, 'parts': [{'partId': '0.0', 
'mimeType': 'text/plain', 'filename': '', 'headers': [{'name': 'Content-Type', 'value': 'text/plain; charset=utf-8'}, {'name': 'Content-Transfer-Encoding', 'value': 'base64'}], 'body': {'size': 231, 'data': 'UGxlYXNlIGZpbmQgeW91ciBkb2N1bWVudCBhdHRhY2hlZC4gVGhhbmsgeW91Lg0KwqANCllvdXIgZG9jdW1lbnQgaXMgaW4gQWRvYmUgQWNyb2JhdCBQb3J0YWJsZSBEb2N1bWVudCBGb3JtYXQgKFBERikuIElmIHlvdSBkb24ndCBoYXZlIEFkb2JlIFJlYWRlciwgeW91IGNhbiBkb3dubG9hZCBpdCBmcmVlIG9mIGNoYXJnZSBieSBjbGlja2luZyBoZXJlOg0KaHR0cDovL2dldC5hZG9iZS5jb20vcmVhZGVy'}}, {'partId': '0.1', 'mimeType': 'text/html', 'filename': '', 'headers': [{'name': 'Content-Type', 'value': 'text/html; charset=us-ascii'}, {'name': 'Content-Transfer-Encoding', 
'value': 'quoted-printable'}], 'body': {'size': 821, 'data': 'PCFET0NUWVBFIGh0bWwgUFVCTElDICItLy9XM0MvL0RURCBYSFRNTCAxLjAgVHJhbnNpdGlvbmFsLy9FTiIgImh0dHA6Ly93d3cudzMub3JnL1RSL3hodG1sMS9EVEQveGh0bWwxLXRyYW5zaXRpb25hbC5kdGQiPjxodG1sIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hodG1sIj48aGVhZD48c3R5bGUgdHlwZT0idGV4dC9jc3MiPi5UZWxlcmlrTm9ybWFsIHtmb250LWZhbWlseTogQ2FsaWJyaTtmb250LXNpemU6IDE0LjY2NjY2NjY2NjY2NjdweDttYXJnaW4tdG9wOiAwcHg7bWFyZ2luLWJvdHRvbTogMTJweDtsaW5lLWhlaWdodDogMTE1JTt9dGFibGUuVGVsZXJpa1RhYmxlTm9ybWFsIHtib3JkZXItY29sbGFwc2U6IGNvbGxhcHNlO3RhYmxlLWxheW91dDogYXV0bzt9PC9zdHlsZT48dGl0bGU-PC90aXRsZT48L2hlYWQ-PGJvZHk-PHAgY2xhc3M9IlRlbGVyaWtOb3JtYWwiPjxzcGFuPlBsZWFzZSBmaW5kIHlvdXIgZG9jdW1lbnQgYXR0YWNoZWQuIFRoYW5rIHlvdS48L3NwYW4-PC9wPjxwIGNsYXNzPSJUZWxlcmlrTm9ybWFsIj4mbmJzcDs8L3A-PHAgY2xhc3M9IlRlbGVyaWtOb3JtYWwiPjxzcGFuPllvdXIgZG9jdW1lbnQgaXMgaW4gQWRvYmUgQWNyb2JhdCBQb3J0YWJsZSBEb2N1bWVudCBGb3JtYXQgKFBERikuIElmIHlvdSBkb24ndCBoYXZlIEFkb2JlIFJlYWRlciwgeW91IGNhbiBkb3dubG9hZCBpdCBmcmVlIG9mIGNoYXJnZSBieSBjbGlja2luZyBoZXJlOjwvc3Bhbj48L3A-PHAgY2xhc3M9IlRlbGVyaWtOb3JtYWwiPjxzcGFuPmh0dHA6Ly9nZXQuYWRvYmUuY29tL3JlYWRlcjwvc3Bhbj48L3A-PC9ib2R5PjwvaHRtbD4='}}]}, {'partId': '1', 'mimeType': 'multipart/mixed', 'filename': '', 'headers': [{'name': 'Content-Type', 'value': 'multipart/mixed; boundary=--boundary_107_ab57cd0a-6c7e-4950-bb37-ae534c703a2b'}], 'body': {'size': 0}, 'parts': [{'partId': '1.0', 'mimeType': 'application/octet-stream', 'filename': 'Invoice 314863.pdf', 'headers': [{'name': 'Content-Type', 'value': 'application/octet-stream; name="Invoice 314863.pdf"'}, {'name': 'Content-Transfer-Encoding', 'value': 'base64'}, {'name': 'Content-Disposition', 'value': 'attachment; filename="Invoice 314863.pdf"'}], 'body': {'attachmentId': 'ANGjdJ948jpb6Txqjga9dlj38veWLPS6qjQ2MjQjsAiI-3y39XD6KmKG23jkgTjhGOBMOHNQ6eENYU8WxgwQJpMSa0r1yF_0gfg8F5TnsLwjqskBql_vNh1TUyfCIUJpTA5YV-BLiZt_Eom6_Ftd9XtR7IiTZ2Bv-vn20O5mE-zpwnp1tK0GpQzI-4zykMEhTt9T993v_JdGk0lHB4C-dPOG6wfzLnlUIpeWyLEpvCpqSVIPf5XJyY37wDHB1xs1oxeXAilD0ygOFY6b4AW0G9tVUFFnWqZAiRTdmuesBSGB1V4LIvKup78jRMLSQAv5XAKveULif_yutTRzG6_8iZD2Kzn3jsvfiFJoY2LQbiaxuCiQgeZ9b-HGBmeO8v2uE7B858MI6i7-DoptFgOr', 'size': 67183}}]}]
hello
hello
No valid attachments found in the email.
[]
read list: []
smaller list: []"""

from Google import Create_Service
import base64, os, datetime, pickle, time, tkinter, shutil
from time import sleep
import win32print, subprocess
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta, time
import pickle 
from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askopenfilename, Frame, Button


from flask import Flask, request
import requests
from photos_timesheets import extract_first_page_and_overwrite

# from download_repair_photos import mark_as_read

# TODO: Add logic to only check things with the UNREAD tag 


CLIENT_SECRET_FILE = r"C:\Users\Michael\Desktop\python-work\wbrcredentials.json"  # Replace with the path to your credentials.json file
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
  
def download_attachments(service, user_id, msg_id, store_dir, desired_subject, days_ago, nested):
    
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
                    if filename.endswith('.zip'):
                        continue
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
                elif nested == "yes":
                    print('hallelujah')
                    for sub_part in part.get('parts', []):  # Safely get 'parts'
                        if sub_part.get('filename'):
                            filepath = process_nested_attachment(sub_part, store_dir, service, user_id, msg_id)
                            if filepath:  # Only append if the file was actually saved
                                attachments_paths.append(filepath)
                        else:
                            print("Skipping a sub-part with no 'filename' attribute.")

            if not found_valid_attachment:
                print("No valid attachments found in the email.")
        else:
            print(f"Skipping email with subject: {subject}")
    except Exception as e:
        print('An error occurred: %s' % e)
    print(f'attachment_paths: {attachments_paths}')
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
    sumatra_path = r"C:\Users\Michael\AppData\Local\SumatraPDF\SumatraPDF.exe"
    printer_name = win32print.GetDefaultPrinter()

    if not os.path.exists(filepath):
        print(f"File '{filepath}' does not exist.")
        return

    # command = [
    #     ghostscript_path,
    #     "-q"
    #     "-dNOPAUSE",
    #     "-dBATCH",
    #     "-dPrinted",
    #     f"-sDEVICE=mswinpr2",  # Use the Windows printer device
    #     f"-sOutputFile=%printer%{printer_name}",
    #     filepath
    # ]
    command = [sumatra_path, '-print-to-default', filepath]
    
    try:
        subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
        print(f"File '{filepath}' printed successfully.")
    except Exception as e:
        print(f"An error occurred while printing: {e}")

def get_subject_from_message(service, user_email, msg_id):
    """Get the subject of an email based on its id"""
    message = service.users().messages().get(userId=user_email, id=msg_id).execute()
    headers = message['payload']['headers']
    for header in headers:
        if header['name'] == 'Subject':
            return header['value']
    return None

def mark_as_read(service, user_id, msg_id):
    try:
        # Remove 'UNREAD' label
        service.users().messages().modify(userId=user_id, id=msg_id, body={'removeLabelIds': ['UNREAD']}).execute()
        print(f'Message {msg_id} marked as read.')
    except Exception as e:
        print(f'An error occurred: {e}')
        return None
    
def clear_directory(folder_path):
    """
    Deletes all files and subfolders in the given directory without deleting the directory itself.
    """
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f"Failed to delete {file_path}. Reason: {e}")

def process_nested_attachment(sub_part, store_dir, service, user_id, msg_id):
    # Extract sub-part information
    sub_filename = sub_part['filename']
    sub_attachment_id = sub_part['body']['attachmentId']
    
    # Download the sub-attachment using its ID
    sub_attachment = service.users().messages().attachments().get(userId=user_id, messageId=msg_id, id=sub_attachment_id).execute()
    
    # Decode the attachment data
    sub_file_data = base64.urlsafe_b64decode(sub_attachment['data'].encode('UTF-8'))

    # Additional debugging: Check the size and first few bytes of 'sub_file_data'
    print(f"Size of sub 'file_data': {len(sub_file_data)}")
    print(f"First few bytes of sub 'file_data': {sub_file_data[:100]}")

    # Check if the sub-file is a PDF
    _, file_extension = os.path.splitext(sub_filename)
    if file_extension.lower() == ".pdf":
        # Save the sub-attachment to the specified 'store_dir'
        os.makedirs(store_dir, exist_ok=True)
        filepath = os.path.join(store_dir, sub_filename)
        with open(filepath, 'wb') as f:
            f.write(sub_file_data)

        print(f"Sub-attachment '{sub_filename}' downloaded to: {filepath}")
        return filepath
    else:
        print(f"Skipped sub-attachment '{sub_filename}' because it is not a PDF or is empty.")
        return None




if __name__ =="__main__":
    service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
    user_email = "wbrroof@gmail.com"  # Replace with the email address you want to send the message from
    store_directory = r"C:\Users\Michael\Desktop\python-work\Invoices"
    clear_directory(store_directory)

    print("print invoice")
    today = datetime.utcnow().date()
    start_of_today = datetime.combine(today, time.min)
    start_of_tomorrow = start_of_today + timedelta(days=1)

    start_of_today_timestamp = int(start_of_today.timestamp()) * 1000
    start_of_tomorrow_timestamp = int(start_of_tomorrow.timestamp()) * 1000

    # Load access token from pickle file
    pickle_file = f'C:\\Users\\Michael\\Desktop\\python-work\\token_{API_NAME}_{API_VERSION}.pickle'
    with open(pickle_file, 'rb') as token_file:
        access_token = pickle.load(token_file)
        print('pickle')
        print(access_token.token)
    print(pickle_file)

    email_list = [
        "from:carolyn@profastening.net subject:'Invoice'", 
        "from:Sales@gemcoroofingsupply.com subject:'Invoice'", 
        "from:april@sheetmetalsupplyltd.com subject:'Invoice'",
        "from:april@sheetmetalsupplyltd.com subject:'Invoices'",
        "amy@profastening.net subject:'Invoice'",
        "dawn@sheetmetalsupplyltd.com subject:'Invoice'",
        "lia@stevensoncrane.com subject:'invoice'",
        "kris@stevensoncrane.com subject:'Invoice",
        "customercareBT@becn.com subject: 'invoice'",
        "donotreply@waterinvoice.com subject: 'eInvoice'",
        "jillian.schoedel@industrialandwholesalelumber.com subject: 'Invoice'",
        "kathy@dandpconstruction.com subject: 'invoice'",
        "sales@duro-last.com subject: 'Order'",
        ]

    # Test email list
    # email_list = [
    #     #     "from:carolyn@profastening.net subject:'Invoice'", 
    #     #     "from:Sales@gemcoroofingsupply.com subject:'Invoice'", 
    #     #     "from:april@sheetmetalsupplyltd.com subject:'Invoice'",
    #     #     "from:april@sheetmetalsupplyltd.com subject:'Invoices'",
    #     #     "amy@profastening.net subject:'Invoice'",
    #     #     "dawn@sheetmetalsupplyltd.com subject:'Invoice'",
    #         "lia@stevensoncrane.com subject:'invoice'",
    #     #     "customercareBT@becn.com subject: 'invoice'",
    #     #     "donotreply@waterinvoice.com subject: 'eInvoice'",
    #         ]


    # Asks how many days back to check
    desired_date = 1
    print_status = True
    query_date = datetime.now() - timedelta(days=desired_date)  # Using the desired_date variable instead of fixed 7
    query_date_str = query_date.strftime('%Y-%m-%d')

    # Define the watch reuest body
    request = {
        'labelIds': ['INBOX'],
        'topicName': 'projects/gmail-project-394016/topics/Base_Topic',
        'labelFilterBehavior': 'INCLUDE'
    }
    # Set up the watch
    watch_response = service.users().watch(userId='me', body=request).execute()

    # Print the response
    print('Watch response:', watch_response)

    # app = Flask(__name__)


    # @app.route('/', methods=['GET'])
    # def home():
    #     return 'Server is running!'


    # @app.route('/subscribe', methods=['POST'])
    # def subscribe():
    #     print('hw')
    #     watch_payload = {
    #         "topicName": "projects/gmail-project-394016/topics/Base_Topic",
    #         "labelIds": ["INBOX"],
    #         "labelFilterBehavior": "INCLUDE"
    #     }
        
    #     headers = {
    #         "Authorization": f"Bearer {access_token}",
    #         "Content-Type": "application/json"
    #     }
        
    #     response = requests.post(
    #         "https://www.googleapis.com/gmail/v1/users/me/watch",
    #         json=watch_payload,
    #         headers=headers
    #     )
        
    #     if response.status_code == 200:
    #         return "Subscribed to Gmail notifications."
    #     else:
    #         return f"Failed to subscribe. Error: {response.text}", 400


    # @app.route('/notification', methods=['POST'])
    # def notification():
    #     try:
    #         notification_data = request.json  # Assuming Gmail sends JSON notifications
    #         resource_state = notification_data['message']['data']['emailHistoryId']  # Extract the resource state
    #         print(f'Received notification with resource state: {resource_state}')

    #         if resource_state == 'exists':
    #             # Fetch email details using historyId or message.id
    #             history_id = notification_data['historyId']  # Extract the historyId
    #             message_id = notification_data['message']['data']['message']['id']  # Extract the message ID

    #             # Use historyId or message_id to fetch email details using Gmail API
    #             # Implement your logic to process the email here
    #             # ...

    #             print(f'Processing email with historyId: {history_id} and message ID: {message_id}')
    #         else:
    #             print('Received notification for resource state other than "exists"')

    #         return 'OK', 200
    #     except Exception as e:
    #         print(f'An error occurred: {e}')
    #         return 'Error', 500

    # if __name__ == '__main__':
    #     app.run(debug=True)

    check_unread = True

    read_list = []
    smaller_list = []

    for email_query in email_list:
        if check_unread:
            query = f"is:unread {email_query} after:{query_date_str}"
        else:
            query = f"{email_query} after:{query_date_str}"
        print(email_query)
        try:
            response = service.users().messages().list(userId=user_email, q=query).execute()
            messages = response.get('messages', [])
            all_attachments = []
            for message in messages:
                msg_id = message['id']
                if email_query == "from:april@sheetmetalsupplyltd.com subject:'Invoice'" or email_query == "from:april@sheetmetalsupplyltd.com subject:'Invoices'":
                    nested = "yes"
                    attachments = download_attachments(service, user_email, msg_id, store_directory, "Invoice", desired_date, nested)
                elif email_query == "jillian.schoedel@industrialandwholesalelumber.com subject: 'Invoice'":
                    nested = "yes"
                    attachments = download_attachments(service, user_email, msg_id, store_directory, "Invoice", desired_date, nested)
                else:    
                    nested = "no"
                    attachments = download_attachments(service, user_email, msg_id, store_directory, "Invoice", desired_date, nested)

                subject = get_subject_from_message(service, user_email, msg_id)  # Get the subject of the message

                if attachments:
                    all_attachments.extend(attachments)
                    smaller_list.append({'id': subject})  # Append subject instead of ID
                    mark_as_read(service, user_email, msg_id)

                print(f"List of attachments: {attachments}")
                print(f"List of attachments: {all_attachments}")
                


            sleep(1)
            if print_status and email_query == "donotreply@waterinvoice.com subject: 'eInvoice'":
                for attachment in all_attachments:
                    extract_first_page_and_overwrite(attachment)
                    print_file_with_ghostscript(attachment)
            elif print_status:
                # print('print status')
                for attachment in all_attachments:
                    print_file_with_ghostscript(attachment)

        except Exception as e:
            print('An error occurred: %s' % e)
        print(f"read list: {read_list}")
        print(f"smaller list: {smaller_list}")


    # def main():
    #     # ... (Your existing code for authentication and obtaining the service object)
    #     service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

    #     user_email = "wbrroof@gmail.com"  # Replace with the email address you want to fetch messages from
    #     store_directory = r"C:\Users\Michael\Desktop\python-work\Invoices"

    #     today = datetime.utcnow().date()
    #     start_of_today = datetime.combine(today, time.min)
    #     start_of_tomorrow = start_of_today + timedelta(days=1)

    #     start_of_today_timestamp = int(start_of_today.timestamp()) * 1000
    #     start_of_tomorrow_timestamp = int(start_of_tomorrow.timestamp()) * 1000

    #     query_list = [
    #         "from:carolyn@profastening.net subject:'Invoice'",
    #         "from:Sales@gemcoroofingsupply.com subject:'Invoice'",
    #         "from:april@sheetmetalsupplyltd.com subject:'Invoice'",
    #         "from:wbrroof@aol.com"
    #     ]

    #     print_status = messagebox.askyesno("Confirmation", "Do you want to print matching invoices?")

    #     for query in query_list:
    #         try:
    #             response = service.users().messages().list(userId=user_email, q=query).execute()
    #             messages = response.get('messages', [])
    #             all_attachments = []
    #             for message in messages:
    #                 msg_id = message['id']
    #                 attachments = download_attachments(service, user_email, msg_id, store_directory, "Invoice")
    #                 if attachments:
    #                     all_attachments.extend(attachments)
    #                 print(attachments)

    #             sleep(1)
    #             if print_status:
    #                 for attachment in all_attachments:
    #                     print_file_with_ghostscript(attachment)
    #         except Exception as e:
    #             print('An error occurred: %s' % e)

    # if __name__ == "__main__":
    #     main()
    # recipient_email = "throwod@gmail.com"  # Replace with the recipient's email address
    # message_subject = "Testing Gmail API"
    # message_text = "Bruh, this is a message sent via the Gmail API! Let's go!"

    # message = create_message(user_email, recipient_email, message_subject, message_text)
    # if message:
    #     send_message(service, user_email, message)