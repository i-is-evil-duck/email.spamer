import os
import json
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from dotenv import load_dotenv

# Load environment variables from .env
load_dotenv()

# Get sender info from environment variables
sender_email = os.environ.get('EMAIL')
sender_password = os.environ.get('Key')

if not sender_email or not sender_password:
    raise ValueError("EMAIL or Key environment variable is missing.")

# Load email subject and message from JSON
with open('email.json', 'r') as f:
    email_data = json.load(f)

subject = email_data['subject']
message_template = email_data['message']

# Excel file with recipient info
file_path = "example.xlsx"

# PDF files to attach
pdf_files = ['./package.pdf', './letter.pdf']

# Connect to SMTP server
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(sender_email, sender_password)

# Load Excel workbook
workbook = openpyxl.load_workbook(file_path)
worksheet = workbook.active

# Loop through recipients
for row in worksheet.iter_rows(min_row=2, values_only=True):
    name = row[0]
    recipient_email = row[1]

    # Skip row if email or name is missing
    if not name or not recipient_email:
        continue

    # Create email
    email_message = MIMEMultipart()
    email_message['From'] = sender_email
    email_message['To'] = recipient_email
    email_message['Subject'] = subject

    # Add personalized body
    email_body = message_template.format(name=name)
    email_message.attach(MIMEText(email_body, 'plain'))

    # Attach PDF files
    for file in pdf_files:
        try:
            with open(file, 'rb') as f:
                attach_file = MIMEApplication(f.read(), _subtype="pdf")
                attach_file.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file))
                email_message.attach(attach_file)
        except FileNotFoundError:
            print(f"Warning: {file} not found, skipping attachment.")

    # Send email
    try:
        server.sendmail(sender_email, recipient_email, email_message.as_string())
        print(f"Email sent to {name} ({recipient_email})")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}: {e}")

# Quit server
server.quit()
print("All emails processed.")
