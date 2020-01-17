# plain email and html alternative
# can attach images / PDF files

import smtplib
import imghdr
from email.message import EmailMessage
import csv

sender_email = "pythonemailtests@gmail.com"
password = input()
# receiver_email = "b.kondev@publicis-dialog.bg"
# contacts = ["b.kondev@publicis-dialog.bg", ] Putting emails in a list will send one (not separate) email to all of them

with open("contacts.csv") as file:
    reader = csv.reader(file)
    next(reader)
    for name, email in reader:
        message = EmailMessage()
        message["Subject"] = "My favorite profile photos"
        message["From"] = sender_email
        message["To"] = email
        # message["Cc"] = email This one works OK too.
        message.set_content(f"Hi, {name},\nPlease find attached a selection of my favorite profile images."
                            "\nYou can choose whichever you want.\nThank you,\nBoris")

        message.add_alternative(f"""\
        <!DOCTYPE html>
        <html lang="en">
        <head>
        </head>
        <body>
        <p>Hi, {name},</p>
        <p>Please find attached a selection of my favorite profile photos.</p>
        <p>You can choose whichever you want.</p>
        <p>Thank you,</p>
        <p>Boris</p>
        <p>PS. In case you want to see more photos of mine, you can check my <a href="https://www.facebook.com/boris.kondev">Facebook</a> page.</p>
        </body>
        </html>
        """, subtype="html")

        # For sending images
        elements = ["profile1.jpg", "profile2.jpg", "profile3.jpg", "profile4.jpg"]

        for element in elements:
            with open(element, "rb") as image:
                file_data = image.read()
                file_type = imghdr.what(image.name)
                file_name = image.name

            message.add_attachment(file_data, maintype="image", subtype=file_type, filename=file_name)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, password)
            server.send_message(message)

'''
# For sending PDF files
# The same principle applies for Excel files :)
pdfs = ["file.pdf"]

for pdf in pdfs:
    with open(pdf, "rb") as file:
        file_data = file.read()
        file_name = file.name

    message.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)
'''

