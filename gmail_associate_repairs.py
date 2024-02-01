# This is a program designed to associate sent quickbooks invoices with 
# repair photo emails so that it can forward silently in the gmail api

from photos_timesheets import send_message, CLIENT_SECRET_FILE, SCOPES, API_NAME, API_VERSION
from Google import Create_Service
import base64
import re
from email import message_from_bytes
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
    repairmen = ["Frank E", "Jr.", "Mike Blidy"]
    email_to_label_mapping = {
                "Jr.": "pics-jr---dave-myers", # PICS-JR - Dave Myers
                "Frank E": "pics-frank-espitia",     # PICS-Frank Espitia
                "Mike Blidy": "pics-mike-blidy",      # PICS-Mike Blidy
                # Add more mappings if needed
            }

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
        def create_on_repairman_submit_function(email_id, email_to, email_subject, repairman_window):
            def on_repairman_submit():
                selected_repairman = repairman_var.get()
                keyword = keyword_entry.get()

                print(f"Selected Repairman: {selected_repairman}")
                print(f"Keyword: {keyword}")

                label_id = email_to_label_mapping[selected_repairman]

                repairman_emails = fetch_repairman_emails(service, 'me', label_id, keyword)
                print(repairman_emails)

                # If no emails are found, maybe inform the user
                if not repairman_emails:
                    tk.messagebox.showinfo("Info", "No emails found for the given criteria.")
                    return

                display_repairman_emails(repairman_emails, email_id, email_to, email_subject)
                repairman_window.destroy()

            return on_repairman_submit

        selected_email_details = [(email['id'], email['to'], email['subject']) for var, email in selected_emails if var.get()]

        # If there are no selected emails, return early
        if not selected_email_details:
            return

        # # Assuming only one email is selected
        # selected_email_id, selected_email_to, selected_email_subject = selected_email_details[0]
        # Process each email one-by-one
        for email_id, email_to, email_subject in selected_email_details:
            # Create a new window for selecting repairman and keyword
            match = re.search(r"Invoice (\d+)", email_subject)
            if match:
                invoice_number = match.group(1)
                new_subject = f"Invoice {invoice_number}"
            
            repairman_window = tk.Toplevel(root)
            repairman_window.title(f"Select Repairman for {new_subject}")  # To give context to the user
            # Set window size
            repairman_window.geometry("400x300")
            # Radio buttons for repairman
            repairman_var = tk.StringVar(value="")  # Default value as empty string
            tk.Label(repairman_window, text="Choose a repairman:").pack(pady=10)
            for repairman in repairmen:
                tk.Radiobutton(repairman_window, text=repairman, variable=repairman_var, value=repairman).pack(anchor=tk.W)

            # Entry for keyword
            tk.Label(repairman_window, text="Enter a keyword for the repair photo invoice:").pack(pady=10)
            keyword_entry = ttk.Entry(repairman_window)
            keyword_entry.pack(pady=5)
            
            # Assuming you have a button or something to confirm the selection
            # The loop will pause for each email until the user takes an action on the 'repairman_window' before moving to the next email
            configured_on_repairman_submit = create_on_repairman_submit_function(email_id, email_to, email_subject, repairman_window)

            btn = ttk.Button(repairman_window, text="Submit", command=configured_on_repairman_submit)
            repairman_window.bind('<Return>', lambda event=None: configured_on_repairman_submit())

            btn.pack(pady=5)
            print("Button created and packed.")

            
            repairman_window.wait_window()  # Waits for the repairman_window to close before processing the next email


      
        # # Submit button
        # ttk.Button(repairman_window, text="Submit", command=on_repairman_submit).pack(pady=10)
    # The submit button for the initial email selection window
    btn = ttk.Button(root, text="Submit", command=on_submit)
    btn.pack()
    
    # ttk.Button(root, text="Submit", command=on_repairman_submit).pack(pady=10)
    root.bind('<Return>', lambda event=None: on_submit())

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


def forward_email(service, user_email, msg_id):
    # Fetch the original email content
    email_data = service.users().messages().get(userId=user_email, id=msg_id, format='raw').execute()
    raw_email_data = base64.urlsafe_b64decode(email_data['raw'].encode('ASCII'))

    # Convert the raw email data into an email.mime object
    original_msg = message_from_bytes(raw_email_data)

    # Extract recipient and subject from the original email
    to_email = original_msg['To']  # Assuming the original 'From' will be the 'To' for the forwarded email
    print(to_email)
    subject = original_msg['Subject']
    match = re.search(r"Invoice (\d+)", subject)
    if match:
        invoice_number = match.group(1)
        new_subject = f"Repair photos for Invoice {invoice_number}"
    else:
        new_subject = "Repair photos"

    # Create a new MIMEMultipart email to house the forwarded email
    mime_msg = MIMEMultipart()
    mime_msg['Subject'] = new_subject
    mime_msg['To'] = 'throwod@gmail.com'
    

    # Create the custom body for the email
    body_text = """
    Hello,

    Please see attached photos, thank you.

    Version 2
    Company
    O: 123-456-7890
    email@email.com
    """
    mime_msg.attach(MIMEText(body_text, 'plain'))
    
    # Attach the original email to the new MIMEMultipart email
    mime_msg.attach(original_msg)

    raw_msg = base64.urlsafe_b64encode(mime_msg.as_bytes()).decode('utf-8')

    body = {
        'raw': raw_msg
    }

    response = service.users().messages().send(userId=user_email, body=body).execute()
    return response

def forward_combined_email(service, user_email, original_to_email, original_subject, repairman_email_id):
    # 1. Extract invoice number from the original subject
    match = re.search(r"Invoice (\d+)", original_subject)
    if match:
        invoice_number = match.group(1)
        new_subject = f"Repair photos for Invoice {invoice_number}"
    else:
        new_subject = "Repair photos"

    # 2. Fetch the repairman's email content and convert it
    repairman_email_data = service.users().messages().get(userId=user_email, id=repairman_email_id, format='raw').execute()
    raw_repairman_email_data = base64.urlsafe_b64decode(repairman_email_data['raw'].encode('ASCII'))
    repairman_msg = message_from_bytes(raw_repairman_email_data)

    # 3. Create a new MIMEMultipart email
    mime_msg = MIMEMultipart()
    mime_msg['Subject'] = new_subject
    mime_msg['To'] = original_to_email   # If you want to set it to a fixed email, replace `original_to_email` with the desired email

    # Custom body for the email
    body_text = """
    Hello,

    Please see attached photos, thank you.

    Michael Wormley
    WBR Roofing
    ​O: 847-487-8787​
    ​wbrroof@aol.com
    """
    mime_msg.attach(MIMEText(body_text, 'plain'))
    
    # 4. Attach specific parts from repairman's email
    for part in repairman_msg.walk():
        if part.get_content_type() in ["image/jpeg", "image/png", "application/pdf"]:  
            mime_msg.attach(part)

    # 5. Convert and send
    raw_msg = base64.urlsafe_b64encode(mime_msg.as_bytes()).decode('utf-8')
    body = {'raw': raw_msg}

    response = service.users().messages().send(userId=user_email, body=body).execute()
    return response

def fetch_repairman_emails(service, user_email, label_id, keyword):
    one_month_ago = (datetime.now() - timedelta(weeks=4)).strftime('%Y/%m/%d')
    query = f"label:{label_id} {keyword} after:{one_month_ago}"  # Adjust as needed
    
    print(query)
    response = service.users().messages().list(userId=user_email, q=query).execute()
    return response.get('messages', [])

def display_repairman_emails(repairman_emails, original_email_id, original_email_to, original_email_subject):
    root = tk.Tk()
    root.title("Select Repairman Emails")

    selected_emails = []

    for email_data in repairman_emails:
        email_id = email_data['id']
        full_email = service.users().messages().get(userId='me', id=email_id).execute()
        headers = full_email['payload']['headers']
        
        # Extract subject from headers
        subject = next((header['value'] for header in headers if header['name'] == 'Subject'), 'No Subject')
        
        var = tk.BooleanVar(root)

        c = tk.Checkbutton(root, text=subject, variable=var)
        c.pack(anchor=tk.W)
        selected_emails.append((var, email_data))

    def on_email_select_submit():
        # Capture the IDs of the selected repairman emails
        selected_repairman_email_ids = [email['id'] for var, email in selected_emails if var.get()]

        # If no repairman emails are selected, notify the user and return
        if not selected_repairman_email_ids:
            tk.messagebox.showinfo("Info", "No repairman emails selected.")
            return

        # Loop through the selected repairman emails for forwarding
        for email_id in selected_repairman_email_ids:
            # Assuming forward_combined_email is the function you're using to handle the email modifications and forwarding
            forward_combined_email(service, 'me', original_email_to, original_email_subject, email_id)

        # Provide feedback to the user
        tk.messagebox.showinfo("Success", f"{len(selected_repairman_email_ids)} emails forwarded successfully!")

        # Close the display_repairman_emails window
        root.destroy()

    btn = ttk.Button(root, text="Submit", command=on_email_select_submit)
    btn.pack()
    root.mainloop()

def display_repairman_emails_and_associate(repairmain_emails, original_email_id, original_email_to, original_email_subject, root):
    """I started working on this but I think it's currently abandoned"""
    
    # Create a new window for selecting repairman and keyword
    repairman_window = tk.Toplevel()
    repairman_window.title("Select Repairman and Keyword for " + original_email_subject)

    # Display repairman emails for the user to select which email to forward to
    repairman_email_list = repairmain_emails
    repairman_email_vars = []
    for email in repairman_email_list:
        email_var = tk.BooleanVar(value=False)
        repairman_email_vars.append(email_var)
        tk.Checkbutton(repairman_window, text=email['subject'], variable=email_var).pack()

    # Entry for the user to input keyword
    tk.Label(repairman_window, text="Keyword:").pack(pady=(20, 5))
    keyword_entry = tk.Entry(repairman_window)
    keyword_entry.pack(pady=(0, 20))

    def associate_and_forward():
        # Here you would handle forwarding the original email to the selected repairman email addresses
        for idx, email_var in enumerate(repairman_email_vars):
            if email_var.get():
                selected_repairman_email = repairman_email_list[idx]
                # Use both the original_email_id and selected_repairman_email to handle forwarding
                forward_email(original_email_id, selected_repairman_email, keyword_entry.get())
        repairman_window.destroy()

    tk.Button(repairman_window, text="Associate and Forward", command=associate_and_forward).pack()

    # Make sure the repairman_window is modal, i.e., user has to close it before interacting with other windows
    repairman_window.transient(repairman_window.master)
    repairman_window.grab_set()
    repairman_window.wait_window()
    
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