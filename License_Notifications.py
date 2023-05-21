#!/usr/bin/env python3

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import pandas as pd
import smtplib
import datetime
import sys
import automationassets

from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File

# From, To, and sender, receiver info + subject and message email contend
your_name = "<enter a name>"
sender = "<enter an email>"
receiver = "<enter an email>"
subjects = 'License will expire soon'
messages = 'Please make sure to re-purchase the licence if needed before'
url_message = "Link to <file-name>.xlsx:"
url = MIMEText('<a href="<enter excel url>">LINK</a>','html')

# Get Credentials from Automation Account
cred = automationassets.get_automation_credential("<enter the name id>")

# SMTP Office 365 server
server = smtplib.SMTP('smtp.office365.com', 587)
server.ehlo()
server.starttls()
email = cred['username']
passwd = cred['password']
server.login(email, passwd)

# Present the date of the current day
today = datetime.date.today()

# Calculates 90 days, 30 days and 7 days
ninety_days = today + datetime.timedelta(days=90)
thirty_days = today + datetime.timedelta(days=30)
seven_days = today + datetime.timedelta(days=7)

# Filename for licenses
licenses_path = "./<enter your file name>"

# Download the <file-name>.xlsx from Sharepoint
sharepoint_url = '<enter sharepoint site url>'
ctx_auth = AuthenticationContext(sharepoint_url)
ctx_auth.acquire_token_for_user(email, passwd)   
ctx = ClientContext(sharepoint_url, ctx_auth)
response = File.open_binary(ctx, "<enter excel path>")

# Read the .xlsx file
email_list = pd.read_excel(response.content)

# Output with pandas dates converted to YYYY-MM-DD (in order to convert it python's datetime)
email_list["Expiration Date (DD/MM/YYYY or D/M/YYYY)"] = pd.to_datetime(email_list["Expiration Date (DD/MM/YYYY or D/M/YYYY)"]).dt.strftime("%Y-%m-%d")

# Get all the names and the dates from <file-name>.xlsx file
all_names = email_list['Software Name']
all_dates = email_list['Expiration Date (DD/MM/YYYY or D/M/YYYY)']
# print(all_dates)
# datetime instance
dt = datetime.datetime

# Loop through the names
for idx in range(len(all_names)):

    # Get each records name and date
    name = all_names[idx]
    # date = all_dates[idx]
    #    print(date)
    try:
        date = datetime.date.fromisoformat(all_dates[idx])
        # print(date)
        if date == ninety_days or date == thirty_days or date == seven_days or date == today:
            full_email = MIMEMultipart('alternative')
            full_email['Subject'] = name + " " + subjects
            full_email['From'] = sender
            full_email['To'] = receiver

            html = """
            <html>
            <head></head>
            <body>
                <h2>Please make sure to re-purchase the licence if needed before {date}</h2>
                <a href="<enter excel url>">Link to <file-name>.xlsx</a>
            </body>
            </html>
            """.format(date = str(date))

            part1 = MIMEText(html,'html')
            part2 = MIMEText("Link to <file-name>.xlsx:\n<file-name>.xlsx", 'text')
            full_email.attach(part1)
            full_email.attach(part2)

            # In the email field, you can add multiple other emails if you want
            # all of them to receive the same text
            try:
                server.sendmail(sender, [receiver], full_email.as_string())
                print('Email to {} successfully sent!\n\n'.format(receiver))
            except Exception as e:
                print('Email to {} could not be sent :( because {}\n\n'.format(receiver, str(e)))
    finally:
        continue
# Close the smtp server
server.close()
sys.exit()
