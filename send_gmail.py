from photos_timesheets import send_message, CLIENT_SECRET_FILE, SCOPES, API_NAME, API_VERSION
from Google import Create_Service
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

def create_message(to, subject, message_text):
    """Create a message for an email.

    Args:
      sender: Email address of the sender.
      to: Email address of the receiver.
      subject: The subject of the email message.
      message_text: The text of the email message.

    Returns:
      An object containing a base64url encoded email object.
    """
    message = MIMEText(message_text)
    message['to'] = to
    message['subject'] = subject
    # Return the message object encoded in base64 and formatted as ASCII
    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")}

def create_message_with_attachment(to, subject, message_text, file_path):
    """Create a message with an attachment for an email.

    Args:
      sender: Email address of the sender.
      to: Email address of the receiver.
      subject: The subject of the email message.
      message_text: The text of the email message.
      file_path: Path to the file to be attached.

    Returns:
      An object containing a base64url encoded email object.
    """
    message = MIMEMultipart()
    message['to'] = to
    message['subject'] = subject
    
    # Add the plain text to the email
    message.attach(MIMEText(message_text, 'plain'))
    
    # Process the attachment and add it to the email
    with open(file_path, 'rb') as attachment_file:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment_file.read())
        encoders.encode_base64(part)
        
        # Use the filename as the attachment name
        part.add_header('Content-Disposition', 'attachment; filename="{}"'.format(file_path.split('\\')[-1]))
        message.attach(part)
    
    # Return the message object encoded in base64 and formatted as ASCII
    return {'raw': base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")}

def initialize_service():
    return Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

if __name__ == "__main__":
    # sender_email = "wbrroof@gmail.com"
    receiver_email = "throwod@gmail.com"
    subject = "Hello from Gmail API"
    body = "This is a test email sent from the Gmail API."
    attachment_path = r"C:\Users\Michael\Desktop\python-work\Vendors.xlsx"  # Replace with your file path

    message = create_message_with_attachment(receiver_email, subject, body, attachment_path)


    # Setup the Gmail service
    service = initialize_service()
    #  service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)

    # Define your message here (this is just an example structure, you'd fill it with relevant data)
    # message_body = {
    #     'raw': 'your_encoded_message_here'
    # }

    # Use the function
    send_message(service, 'me', message)