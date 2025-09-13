import smtplib
from email.message import EmailMessage
import os

def sendEmail(emailData, receiverEmail, attachment_path=None):
    """
    Sends an email with optional attachment.
    emailData format: [receiver_email, subject, body, closing, sender_email]

    receiverEmail: If provided, overrides the receiver_email from emailData.
    attachment_path: Optional path to a file (e.g., the Excel timesheet) to attach.
    """

    data_receiver, subject, body, closing, sender = emailData
    final_receiver = receiverEmail if receiverEmail else data_receiver

    msg = EmailMessage()
    msg['From'] = sender
    msg['To'] = final_receiver
    msg['Subject'] = subject

    full_message = f"{body}\n\n{closing}"
    msg.set_content(full_message)

    # Attach the Excel file if provided and exists
    if attachment_path and os.path.exists(attachment_path):
        with open(attachment_path, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(attachment_path)
            # The maintype and subtype can be set to something generic like
            # 'application'/'octet-stream' for arbitrary binary data.
            msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)
    else:
        print("No attachment found or invalid path. Sending email without attachment.")

    smtp_host = 'smtp.gmail.com'
    smtp_port = 587
    sender_email = sender
    sender_password = "<YOUR_EMAIL_PASSWORD>"  # Secure properly, do not hardcode in production

    try:
        with smtplib.SMTP(smtp_host, smtp_port) as server:
            server.starttls()
            server.login(sender_email, sender_password)
            server.send_message(msg)
            return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return "email not sent"
