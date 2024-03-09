#Checkout email automation by Tushar Anand
#Keep this file and an updated excel template(with proper name) in the same folder
#example - if today's date 16th February, then file name should be - Tushar_2023-02-14_Checkout(any date that is behind today's date)
#fill in proper details in this code
#install all python dependencies in your local
#run python checkout_template.py

import os
import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from openpyxl import load_workbook

# Set the directory where the Excel files are located
directory = ''

# Get today's date
today = datetime.date.today()

# Find the highest date smaller than today's date for which an Excel file exists
previous_day = None
for days_ago in range(1, (today - datetime.date(2022, 1, 1)).days + 1):
    date = today - datetime.timedelta(days=days_ago)
    file_name = '<your first name>_{}_Checkout.xlsx'.format(date)
    file_path = os.path.join(directory, file_name)
    if os.path.isfile(file_path):
        previous_day = date
        break

# If no Excel file was found for a previous day, raise an error
if previous_day is None:
    raise ValueError('No previous Excel file found')

# Get the filename for the Excel file to rename
old_file_name = '<your first name>_{}_Checkout.xlsx'.format(previous_day)

# Set the path of the Excel file to rename
old_path = os.path.join(directory, old_file_name)

# Load the Excel file using openpyxl
wb = load_workbook(filename=old_path)

# Get the current date in the format YYYY-MM-DD
current_date = datetime.date.today().strftime('%Y-%m-%d')

# Set the new file name in the format <your first name>_YYYY-MM-DD_Checkout.xlsx
new_file_name = '<your first name>_{}_Checkout.xlsx'.format(current_date)

# Get the directory where the old file is located
old_directory = os.path.dirname(old_path)

# Set the path of the new file
new_path = os.path.join(old_directory, new_file_name)

# Create a copy of the old file with the new file name
wb.save(new_path)

# Set up the email message
msg = MIMEMultipart()
msg['From'] = 'Tushar Anand <tanand7@asu.edu>'

# Set the email recipients as a list
to_recipients = ['recipient@email.com']
cc_recipients = ['cc@email.com']
msg['To'] = to_recipients[0] #if multiple recipients, use join to create a comma separated list
msg['Cc'] = cc_recipients[0]

all_recipients = to_recipients + cc_recipients

# Set the subject in the format Checkout Form YYYY-MM-DD <your first name>
subject = 'Checkout Form {} <your first name>'.format(current_date)
msg['Subject'] = subject

# Set the email body
body = """Hi Julian/Julia,

<email body>

Thanks,
<your name>"""

# Attach the email body to the email
msg.attach(MIMEText(body))

# Attach the new file to the email
with open(new_path, 'rb') as f:
    attachment = MIMEApplication(f.read(), _subtype='xlsx')
    attachment.add_header('Content-Disposition', 'attachment', filename=new_file_name)
    msg.attach(attachment)

# Send the email using Gmail's SMTP server
smtp_server = 'smtp.gmail.com'
smtp_port = 587
smtp_username = '<sender email>'
smtp_password = '<sender app password>' #enable 2-factor authentication and set app password in your gmail account and paste that here
#go to settings > security > enable 2-factor > set app password > copy the password

server = smtplib.SMTP(smtp_server, smtp_port)
server.ehlo()
server.starttls()
server.login(smtp_username, smtp_password)
server.sendmail(msg['From'], all_recipients, msg.as_string())
server.quit()