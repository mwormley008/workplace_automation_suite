# trying to get this gmail thing going we'll see 

# TODO: Would be nice to see if i could avoid printing blank pdfs 
# close windows after using them
# delete files once used

from Google import Create_Service
import base64, os, datetime, pickle, time, tkinter, re
from time import sleep
import win32print, subprocess
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta, time

from tkinter import Tk, simpledialog, messagebox
from tkinter.filedialog import askopenfilename, Frame, Button
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, landscape

from pyautogui import press, write, hotkey

from pywinauto import Desktop, Application

import shutil, PyPDF2


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
                    
                    if 'time' in subject.lower() or 'expense' in subject.lower() or desired_sender == 'edinc99@gmail.com':  # Check if the word "time" is in the subject
                        print(subject.lower()+'expense')
                        unique_email_dir = os.path.join(r"C:\Users\Michael\Desktop\python-work\time_sheets", safe_subject)
                    else:
                        print(subject.lower()+'repair')
                        unique_email_dir = os.path.join(store_dir, safe_subject)
                    
                    os.makedirs(unique_email_dir, exist_ok=True)

                    filepath = os.path.join(unique_email_dir, filename)  # Use original filename
                    with open(filepath, 'wb') as f:
                        f.write(file_data)

                    print(f"Attachment '{filename}' downloaded to: {filepath}")

                    attachments_paths.append(filepath)
            
            save_info_as_pdf(subject, sender, received_date, text, unique_email_dir)

            # details = f"Subject: {subject}\n\nFrom: {sender}\n\nReceived Date: {received_date}\n\nBody:{text}"
            # details_file = os.path.join(unique_email_dir, "00000000details.txt")
            # with open(details_file, 'w') as f:
            #     f.write(details)

        return attachments_paths

    except Exception as error:
        print(f"An error occurred: {error}")

def save_info_as_pdf(subject, sender, received_date, body, directory):
    filename = os.path.join(directory, "0000000details.pdf")
    c = canvas.Canvas(filename, pagesize=landscape(letter))
    width, height = landscape(letter)  # Get dimensions

    textobject = c.beginText()
    textobject.setFont('Helvetica', 20)  # Set font and size
    textobject.setTextOrigin(30, height - 50)  # Start near top-left corner

    lines = [
        f"Subject: {subject}\n\n",
        f"From: {sender}\n\n",
        f"Received Date: {received_date}\n\n",
        f"Body: {body}"
    ]

    for line in lines:
        for sub_line in line.split('\n'):
            textobject.textLine(sub_line)

    c.drawText(textobject)
    c.save()

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

def mark_as_read(service, user_id, msg_id):
    try:
        # Remove 'UNREAD' label
        service.users().messages().modify(userId=user_id, id=msg_id, body={'removeLabelIds': ['UNREAD', 'INBOX']}).execute()
        print(f'Message {msg_id} marked as read.')
    except Exception as e:
        print(f'An error occurred: {e}')
        return None

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

def print_first_page_with_ghostscript(filepath):
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
        filepath,
        "-dLastPage=1",
        "-sPageList=1",
    ]

    try:
        subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
        print(f"File '{filepath}' printed successfully.")
        print(f'{command}')
    except Exception as e:
        print(f"An error occurred while printing: {e}")


def only_pdfs_in_folder(folder):
    """
    Return True if all files in the folder have a .pdf extension, else False.
    """
    files = [f for f in os.listdir(folder) if os.path.isfile(os.path.join(folder, f))]
    return all(f.lower().endswith('.pdf') for f in files)

def print_pdfs_from_folder(folder_path):
    # List all files in the directory
    files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]

    # Check if all files are PDFs
    if all(f.lower().endswith('.pdf') for f in files):
        for pdf_file in files:
            full_path = os.path.join(folder_path, pdf_file)
            print_file_with_ghostscript(full_path)
            sleep(.5)
            os.remove(full_path)  # Delete the PDF after printing
            sleep(.5)
        # Check if the folder is empty
        if not os.listdir(folder_path):
            try:
                shutil.rmtree(folder_path)
                print(f"Folder '{folder_path}' has been deleted.")
            except Exception as e:
                print(f"An error occurred while deleting the folder: {e}")
        else:
            print(f"Folder '{folder_path}' still contains some files. Not deleted.")
    else:
        print(f"The folder '{folder_path}' contains non-PDF files. Skipping...")

def print_first_pdfs_from_folder(folder_path):
    # List all files in the directory
    files = [f for f in os.listdir(folder_path) if os.path.isfile(os.path.join(folder_path, f))]

    # Check if all files are PDFs
    if all(f.lower().endswith('.pdf') for f in files):
        for pdf_file in files:
            full_path = os.path.join(folder_path, pdf_file)
            extract_first_page_and_overwrite(full_path)
            print_first_page_with_ghostscript(full_path)
            sleep(.5)
            os.remove(full_path)  # Delete the PDF after printing
            sleep(.5)
        # Check if the folder is empty
        if not os.listdir(folder_path):
            try:
                shutil.rmtree(folder_path)
                print(f"Folder '{folder_path}' has been deleted.")
            except Exception as e:
                print(f"An error occurred while deleting the folder: {e}")
        else:
            print(f"Folder '{folder_path}' still contains some files. Not deleted.")
    else:
        print(f"The folder '{folder_path}' contains non-PDF files. Skipping...")

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

def extract_first_page_and_overwrite(source_pdf):
    with open(source_pdf, 'rb') as src_file:
        reader = PyPDF2.PdfReader(src_file)

        if len(reader.pages) == 0:
            print("The source PDF has no pages.")
            return

        writer = PyPDF2.PdfWriter()
        writer.add_page(reader.pages[0])

        # Temporarily store content
        temp_data = writer

    # Overwrite the source file
    with open(source_pdf, 'wb') as out_file:
        temp_data.write(out_file)

    print(f"Overwritten {source_pdf} with just its first page.")


CLIENT_SECRET_FILE = 'wbrcredentials.json'  # Replace with the path to your credentials.json file
API_NAME = 'gmail'
API_VERSION = 'v1'
SCOPES = ['https://www.googleapis.com/auth/gmail.send', 'https://www.googleapis.com/auth/gmail.modify', 'https://www.googleapis.com/auth/gmail.readonly']

if __name__ == "__main__":
    
    service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
    user_email = "wbrroof@gmail.com"  # Replace with the email address you want to send the message from
    store_directory = r"C:\Users\Michael\Desktop\python-work\repair_photos"
    timesheets_directory = r"C:\Users\Michael\Desktop\python-work\time_sheets"
    clear_directory(store_directory)
    clear_directory(timesheets_directory)

    print('repair photos')
    # Tells how many days back to check
    desired_date = simpledialog.askinteger("Desired Dates", "How many days into the past do you want to select emails?")
    check_unread = messagebox.askyesno("Check Unread?", "Would you like to only check the unread emails after your search date?")


    today = datetime.utcnow().date()
    start_of_today = datetime.combine(today, time.min)
    start_of_tomorrow = start_of_today + timedelta(days=1)

    start_of_today_timestamp = int(start_of_today.timestamp()) * 1000
    start_of_tomorrow_timestamp = int(start_of_tomorrow.timestamp()) * 1000



    email_addresses = ["oblivion969.dm@gmail.com", "fespitia76@gmail.com", "mmblidy92@gmail.com", "tawormley@aol.com", "edinc99@gmail.com"]  # List of email addresses


    email_to_label_mapping = {
        "oblivion969.dm@gmail.com": "Label_7", # PICS-JR - Dave Myers
        "fespitia76@gmail.com": "Label_6",     # PICS-Frank Espitia
        "mmblidy92@gmail.com": "Label_8",      # PICS-Mike Blidy
        "tawormley@aol.com": "Label_9",        # PICS-Troy
        "edinc99@gmail.com": "Label_11",       # Time Sheets - Expenses
        # Add more mappings if needed
    }


    print_status = messagebox.askyesno("Confirmation", "Do you want to print matching repair photos and time sheets?")

    for email_address in email_addresses:
        query_date = datetime.now() - timedelta(days=desired_date)  # Using the desired_date variable instead of fixed 7
        query_date_str = query_date.strftime('%Y-%m-%d')

        if check_unread:
            query = f"is:unread from:{email_address} after:{query_date_str}"
        else:
            query = f"from:{email_address} after:{query_date_str}"
            

        try:
            response = service.users().messages().list(userId=user_email, q=query).execute()
            
            messages = response.get('messages', [])
            all_attachments = []
            for message in messages:
                
                msg_id = message['id']
                msg = service.users().messages().get(userId=user_email, id=msg_id).execute()
                subject = [header['value'] for header in msg['payload']['headers'] if header['name'] == 'Subject'][0]
                attachments = download_attachments(service, user_email, msg_id, store_directory, email_address, desired_date)
                
                label_id_to_add = email_to_label_mapping[email_address]

                if attachments:
                    all_attachments.extend(attachments)
                
                # Check if subject contains special terms
                if any(term in subject.lower() for term in ['time', 'expense', 'fwd']):
                    label_id_to_add = 'Label_11'  # Replace with appropriate label ID
                
                mark_as_read(service, user_email, msg_id)

            # Modify labels for the message
                # if 'special_word' in subject:
                    # label_id_to_add = 'special_label_id'
                # else:
                    # label_id_to_add = 'regular_label_id_for_' + email_address  # You should define these label IDs

                service.users().messages().modify(
                    userId=user_email,
                    id=msg_id,
                    body={'addLabelIds': [label_id_to_add]}
                ).execute()

                if email_address ==  "edinc99@gmail.com":
                    path_to_print = (os.path.join(timesheets_directory, subject))
                    path_to_print = path_to_print.replace('\\\\', '\\')
                    print_first_pdfs_from_folder(path_to_print)
                    sleep(.5)
                    # print_first_pdfs_from_folder(os.path.join(store_directory, subject))

            sleep(1)
            # if print_status:
            #     for attachment in all_attachments:
            #         print_file_with_ghostscript(attachment)
        except Exception as e:
            print('An error occurred: %s' % e)
    # TODO: I'd like to add some printing properties here but it looks like I'm going to need to figure out how to do a right click
    print_status = messagebox.askyesno("Confirmation", "Do you want to print?")
    if print_status:

        print_folder1 = r"C:\Users\Michael\Desktop\python-work\repair_photos"
        print_folder2 = r"C:\Users\Michael\Desktop\python-work\time_sheets"
        
        print_folders = [print_folder1, print_folder2]

        for folder in print_folders:
            subfolders = [os.path.join(folder, subfolder) for subfolder in os.listdir(folder) if os.path.isdir(os.path.join(folder, subfolder))]

            sleep(5)
            # if only pdfs
            for subfolder in subfolders:
                if only_pdfs_in_folder(subfolder):
                    print_pdfs_from_folder(subfolder)
                    # Branch for subfolders with only .pdf files
                    print("This subfolder only contains .pdf files:", subfolder)
                    # Add your desired logic for when the subfolder only contains .pdf files here 
            print_items = os.listdir(folder)
            subprocess.run(['explorer', os.path.realpath(folder)])
            sleep(2)
            # Selects the first item in the list
            press('home')
            sleep(.2)
            press('down')
            sleep(.2)
            press('up')
            sleep(1)
            for item in print_items:
                sleep(1)
                press('enter')
                sleep(1)
                press('0')
                sleep(3)
                hotkey('shift', 'F10')
                sleep(1)
                press('p')
                sleep(5)
                hotkey('shift', 'F10')
                sleep(3)
                press('d')
                sleep(3)
                hotkey('ctrl', 'a')
                sleep(2)
                hotkey('shift', 'F10')
                sleep(2)
                press('p')
                sleep(3)
                press('tab', presses=5)
                sleep(2)
                if folder is print_folder1:
                    press('down', presses=6)
                sleep(2)
                press('enter')
                sleep(10)
                if folder is print_folder2:
                    sleep(10)
                hotkey('alt', 'left')
                sleep(3)
                press('down')
                sleep(2)
    print("download_repair_photos.py is complete.")


        

