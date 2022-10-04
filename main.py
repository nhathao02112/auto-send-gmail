import configparser
import imghdr
import os
import smtplib
import ssl
import sys
import time
from email.message import EmailMessage
import pandas as pd

from datetime import date, datetime, timedelta
hnay = datetime.now()
x = hnay.strftime('%X' + ' ' + ' ''%x')

baseDir = os.path.dirname(os.path.realpath(sys.argv[0])) + os.path.sep

config = configparser.RawConfigParser()
config.read(baseDir + 'mail_acount.cfg')
username = config.get('CREDS', 'mail_username')
password = config.get('CREDS', 'mail_password')
df = pd.read_excel("content.xlsx")
for i in range(0, len(df)):
    content = df['To'][i]
    title = df['Subject'][i]
    noidung = df['Body'][i]
    hinhanh = df['Image'][i]
    message = EmailMessage()
    subject = title
    body = noidung
    df.loc[i, 'Date & time'] = datetime.now()
    sender_email = username
    receiver_email = content
    message["From"] = sender_email
    message["To"] = receiver_email
    message["Subject"] = subject
    message.set_content(body)

    with open(hinhanh, 'rb') as m:
        file_data = m.read()
        file_type = imghdr.what(m.name)
        file_name = m.name
    message.add_attachment(file_data, maintype = 'image', subtype = file_type, filename = file_name)

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
         server.login(sender_email, password)
         server.sendmail(sender_email, receiver_email, message.as_string())
         print("Đã gửi Gmail", i + 1)
         time.sleep(10)

df.drop(df.filter(regex="Unnamed"),axis=1, inplace=True)

with pd.ExcelWriter('content.xlsx', engine="openpyxl", mode='a', if_sheet_exists='overlay') as writer:
    df.to_excel(writer, sheet_name="Sheet1")
    writer.save()
print("Đã gửi tất cả mail")
