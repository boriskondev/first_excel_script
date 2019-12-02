import smtplib
from email.message import EmailMessage

sender_email = "pythonemailtests@gmail.com"
password = input()
receiver_email = ["b.kondev@publicis-dialog.bg"]

message = EmailMessage()
message["Subject"] = "Corey Schafer email test"
message["From"] = sender_email
message["To"] = receiver_email
# message["Cc"] = receiver_email
message.set_content("Hi, this is an email test based on the tutorial of Corey Schafer.")

# For sending PDF files
excels = ["test.xlsx"]

for excel in excels:
    with open(excel, "rb") as file:
        file_data = file.read()
        file_name = file.name

    message.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)


with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
    server.login(sender_email, password)
    server.send_message(message)