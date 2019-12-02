import smtplib
import ssl


port = 587
smtp_server = "mail.publicismarcgroup.com"
sender_email = "b.kondev@publicis-dialog.bg"
password = input("Enter password: ")
receiver_email = "b.kondev@publicis-dialog.bg"
message = """\
Subject: Hi there

This message is send from Python.
"""

context = ssl.create_default_context()

with smtplib.SMTP(smtp_server, port) as server:
    server.starttls(context=context)
    server.login(sender_email, password)
    server.sendmail(sender_email, receiver_email, message)
