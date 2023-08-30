# This is a program designed to associate sent quickbooks invoices with 
# repair photo emails so that it can forward silently in the gmail api

from download_repair_photos import send_message, CLIENT_SECRET_FILE, SCOPES, API_NAME, API_VERSION
from Google import Create_Service
import base64
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from send_gmail import initialize_service
import tkinter as tk
from tkinter import ttk
from datetime import datetime, timedelta


def display_emails(emails):
    root = tk.Tk()
    root.title("Select Emails")

    selected_emails = []

    for email in emails:
        match = re.search(r"Invoice (\d+)", email['subject'])
        if match:
            invoice_number = match.group(1)
        else:
            invoice_number = "N/A"  # Default value if no match found

        email_display_text = f"Invoice: {invoice_number} | To: {email['to']}"
        var = tk.BooleanVar()
        c = tk.Checkbutton(root, text=email_display_text, variable=var)
        c.pack(anchor=tk.W)
        selected_emails.append((var, email))

    def on_submit():
        # Handle the selected emails and forwarding logic here
        for var, email in selected_emails:
            if var.get():
                print(f"Selected email: Invoice: {invoice_number} | To: {email['to']}")

    btn = ttk.Button(root, text="Submit", command=on_submit)
    btn.pack()

    root.mainloop()

def fetch_quickbooks_emails(service, user_email):
    # Calculate the date for 24 hours ago
    one_day_ago = (datetime.now() - timedelta(days=1)).strftime('%Y/%m/%d')
    
    query = f'from:wbrroof@gmail.com subject:"Invoice" after:{one_day_ago}'
    response = service.users().messages().list(userId=user_email, q=query).execute()
    messages = response.get('messages', [])
    emails = []

    for message in messages:
        msg_id = message['id']
        msg = service.users().messages().get(userId=user_email, id=msg_id).execute()
        subject = [header['value'] for header in msg['payload']['headers'] if header['name'] == 'Subject'][0]
        to_email = [header['value'] for header in msg['payload']['headers'] if header['name'] == 'To'][0]
        emails.append({'id': msg_id, 'subject': subject, 'to': to_email})


    return emails


def forward_email(service, user_email, msg_id, to_email):
    # Fetch the original email content
    email_data = service.users().messages().get(userId=user_email, id=msg_id, format='raw').execute()
    raw_email_data = base64.urlsafe_b64decode(email_data['raw'].encode('ASCII'))

    # Create a new email for forwarding
    mime_msg = MIMEText(raw_email_data, _subtype='plain')
    mime_msg['Subject'] = "Forwarded Invoice"
    mime_msg['From'] = user_email
    mime_msg['To'] = to_email
    raw_msg = base64.urlsafe_b64encode(mime_msg.as_bytes()).decode('utf-8')

    body = {
        'raw': raw_msg
    }

    response = service.users().messages().send(userId=user_email, body=body).execute()
    return response

if __name__ == "__main__":
    service = initialize_service()
    emails = fetch_quickbooks_emails(service, 'me')
    # Checking if emails were fetched
    if emails:
        print(f"Found {len(emails)} emails from Quickbooks with the desired subject.")
        # Uncomment the next line if you want to see the list of emails
        # print(emails)

        # For a more interactive experience, you could use the display_emails function
        display_emails(emails)
    else:
        print("No emails found from Quickbooks with the desired subject.")