import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

port = 465
smtp_server = "smtp.gmail.com"
sender_email = "pythonemailtests@gmail.com"
password = input("Enter password: ")
receiver_email = "b.kondev@publicis-dialog.bg"

message = MIMEMultipart("alternative")
message["Subject"] = "Test email ;)"
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
    <p>Здравей,<br>
      Как си?<br>
       <a href="http://www.realpython.com">Real Python</a> има много яки уроци по Python. Виж ги, яки са :)
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

