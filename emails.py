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
