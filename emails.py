import smtplib
from email.mime.text import MIMEText

sender = "boriskondev@gmail.com"
password = "Painkiller27081986!"
recipient = "b.kondev@publicis-dialog.bg"
subject = "Test email"
text_enc = MIMEText("Ако получаваш това, значи скриптът работи.".encode("utf-8"))

smtp_server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
smtp_server.login(sender, password)
message = f"Subject: {subject}\n\n{text_enc}"
smtp_server.sendmail(sender, recipient, message)
smtp_server.close()
