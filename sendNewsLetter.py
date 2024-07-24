#! python3
# sendNewsLetter.py

# remember you need to install the modules on the local system, that way when the computer builds the execution environment
# it has the code to use to run the program in the interpreter

import openpyxl, sys, smtplib, email
import pandas as pd
import re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# This loads your current file into a variable, then finds the sheet and the columns you're utilizing

df = pd.read_excel('C:/Users/avisk/OneDrive/Documents/Life/Marriage/Newsletter/FamilyEmailAddresses.xlsx', sheet_name='Sheet1')

# this will reference the above Excel File, you just need to edit the below index in the email variable to include all the people in the email column you want to email
email = df.loc[0:5,'Email']
print(email)
email2 = email.to_numpy()
# print(email2)

recipients = email2
sender_email = "aviskant@gmail.com"
recipient_email = ", ".join(recipients)
subject = "Hello from Python"
body = "Hello all, I was working on automating some things around the house and thought it would be fun to write a family newsletter every once in a while to catch up.  I hope everyone is open to it, just finished writing the Python Program."

# uncomment the below if you would like to add an attachment

# with open("C:/Users/avisk/OneDrive/Documents/Life/Marriage/Newsletter/attachment.txt", "rb") as attachment:
#     # Add the attachment to the message
#     part = MIMEBase("application", "octet-stream")
#     part.set_payload(attachment.read())
# encoders.encode_base64(part)
# part.add_header(
#     "Content-Disposition",
#     f"attachment; filename= 'attachment.txt'",
# )

message = MIMEMultipart()
message['Subject'] = subject
message['From'] = sender_email
message['To'] = recipient_email
html_part = MIMEText(body)

# for attachments uncomment the below

# message.attach(html_part)
# message.attach(part)


smtpObj = smtplib.SMTP()
with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp_server:
    smtpObj.set_debuglevel(1)
    smtp_server.login("aviskant@gmail.com", "vkka vgap tpxm exqn")
    smtp_server.sendmail(sender_email, recipient_email, message.as_string())
    print("Message sent!")
    smtp_server.quit()


# How to Send Emails in Python with Gmail SMTP and API - https://mailtrap.io/blog/python-send-email-gmail/
# Automate the Boring Stuff with Python - Al Sweigart
# Sending email in python (MIMEmultipart) - https://stackoverflow.com/questions/55036268/sending-email-in-python-mimemultipart
# How to send email to multiple recipients using python smtplib? - https://stackoverflow.com/questions/8856117/how-to-send-email-to-multiple-recipients-using-python-smtplib
# Mail Trap - https://mailtrap.io/blog/python-send-email-gmail/
# Mail - https://www.geeksforgeeks.org/send-mail-gmail-account-using-python/
# PyEnv - https://github.com/pyenv/pyenv/wiki#suggested-build-environment
# pandas - https://stackoverflow.com/questions/16729574/how-can-i-get-a-value-from-a-cell-of-a-dataframe