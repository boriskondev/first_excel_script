# https://realpython.com/python-send-email/
# tutorial until message attachment

import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

port = 465
smtp_server = "smtp.gmail.com"
sender_email = "pythonemailtests@gmail.com"
receiver_email = "b.kondev@publicis-dialog.bg"
password = "Pythonemailtests!"

message = MIMEMultipart("alternative")
message["Subject"] = "multipart test"
message["From"] = sender_email
message["To"] = receiver_email

text = """\
Hi,
How are you?
Real Python has many great tutorials:
www.realpython.com"""
html = """\
<html>
  <body style="font-family:verdana;">
    <p>Hi,<br>
       How are you?<br>
       <a href="http://www.realpython.com">Real Python</a> has many great tutorials.
    </p>
  </body>
</html>
"""

part_1 = MIMEText(text, "plain")
part_2 = MIMEText(html, "html")

message.attach(part_1)
message.attach(part_2)

context = ssl.create_default_context()

with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
    server.login(sender_email, password)
    server.sendmail(sender_email, receiver_email, message.as_string())

'''
# This one is working OK

import smtplib

sender = "pythonemailtests@gmail.com"
password = "Pythonemailtests!"
recipient = "b.kondev@publicis-dialog.bg"
subject = "Test email"

with open("email.txt", "r") as f:
    text = f.read()

smtp_server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
smtp_server.login(sender, password)
message = f"Subject: {subject}\n\n{text}"
smtp_server.sendmail(sender, recipient, message)
smtp_server.close()
'''