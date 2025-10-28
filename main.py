import os
import openpyxl
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# Excel file path and name
file_path = "example.xlsx"

# Email sender information
sender_email = my_secret = os.environ['EMAIL']
sender_password = my_secret = os.environ['Key']

# Email subject
subject = "Sponsorship Request"

# Email message
message = """Hello {name},

On behalf of the Alpha Robotics team, we thank you for reading our sponsorship request.

Our names are Aaron, Aiden, Adam, Ethan, Steven and Wellington and we are the only Burnaby team to qualify for the VEX Robotics World Championships held in Dallas, Texas. Attached below is our sponsorship letter and package.

Thanks again for considering a sponsorship.

Regards,

Alpha Robotics 502W"""

# PDF files to attach
pdf_files = ['./package.pdf', './letter.pdf']

# Connect to SMTP server
server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()

# Log in to email account
server.login(sender_email, sender_password)

# Load workbook
workbook = openpyxl.load_workbook(file_path)

# Load worksheet
worksheet = workbook.active

# Iterate through rows
for row in worksheet.iter_rows(min_row=2, values_only=True):
    name = row[0]
    recipient_email = row[1]

    # Create a message
    email_message = MIMEMultipart()
    email_message['From'] = sender_email
    email_message['To'] = recipient_email
    email_message['Subject'] = subject

    # Add message body
    email_body = message.format(name=name)
    email_message.attach(MIMEText(email_body, 'plain'))

    # Add PDF attachments
    for file in pdf_files:
        with open(file, 'rb') as f:
            attach_file = MIMEApplication(f.read(),_subtype = "pdf")
            attach_file.add_header('Content-Disposition', 'attachment', filename = file)
            email_message.attach(attach_file)

    # Send email
    server.sendmail(sender_email, recipient_email, email_message.as_string())

# Quit server
server.quit()
print("All emails sent successfully!")
