import email
import smtplib
import ssl

from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

subject = "This is an email with attachment"
body = "I am tryingto send an email with attachment using Python"
sender_email = "pythonemailtests@gmail.com"
password = input("Enter password: ")
receiver_email = "b.kondev@publicis-dialog.bg"

message = MIMEMultipart()
message["From"] = sender_email
message["To"] = receiver_email
message["Subject"] = subject
message["Bcc"] = receiver_email

message.attach(MIMEText(body, "plain"))

filename = "Notification Adress Change Publicis AD 2019.pdf"

with open(filename, "rb") as attachment:
    part = MIMEBase("application", "octet-stream")
    part.set_payload(attachment.read())

encoders.encode_base64(part)

part.add_header(
    "Content-Disposition",
    f"attachment; filename= {filename}",
)

message.attach(part)
text = message.as_string()

context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(sender_email, password)
    server.sendmail(sender_email, receiver_email, text)