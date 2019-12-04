# -*- coding: utf-8 -*-

import csv
import smtplib
import os
from pathlib import Path
from email.message import EmailMessage
from datetime import datetime

simple_explanation = ""
html_explanation = ""

folder = Path("C:/Users/pmg23_b.kondev/Desktop/Results/Week_01/Banks")

week_to_process = int(input("Week to process: "))
password = input("Email password: ")
sender_email = "emails.tasting@gmail.com"

time_start = datetime.now()

current_contacts = open("all_contacts.csv")
reader = csv.reader(current_contacts)
next(reader)
for bank, name, family, email in reader:
    for file in folder.iterdir():
        winner_or_winners = file.name.split("_")[2]
        winning_bank = file.name.split("_")[3].split(".")[0]
        if bank == winning_bank:
            email = ', '.join(email.split())
            message = EmailMessage()
            message["Subject"] = f"Visa winter promotion 2019 {winner_or_winners}, week {week_to_process}"
            message["From"] = sender_email
            message["To"] = email
            message["Cc"] = sender_email

            if winner_or_winners == "winner":
                simple_explanation = f"\nПриложено изпращам печелившия ви за тази седмица." \
                    f"\nМоля да проверите дали трансакцията му е валидна и да ни върнете потвърждение."
                html_explanation = f"<p>Приложено изпращам печелившия ви за тази седмица.</p>" \
                    f"<p>Моля да проверите дали трансакцията му е валидна и да ни върнете потвърждение.</p>"
            elif winner_or_winners == "winners":
                simple_explanation = f"\nПриложено изпращам печелившите ви за тази седмица." \
                    f"\nМоля да проверите дали трансакциите им са валидни и да ни върнете потвърждение."
                html_explanation = f"<p>Приложено изпращам печелившите ви за тази седмица.</p>" \
                    f"<p>Моля да проверите дали трансакциите им са валидни и да ни върнете потвърждение.</p>"

            message.set_content(f"Здравей, {name},{simple_explanation}\nПоздрави,\nБорис")

            message.add_alternative(f"""\
            <!DOCTYPE html>
            <html lang="en">
            <head>
            </head>
            <body>
            <p>Здравей, {name},</p>
            {html_explanation}
            <p>Поздрави,</p>
            <p>Борис</p>
            </body>
            </html>
            """, subtype="html")

            os.chdir(folder)

            with open(file.name, "rb") as f:
                file_data = f.read()
                file_name = f.name

            message.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)

            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
                server.login(sender_email, password)
                server.send_message(message)

current_contacts.close()

time_end = datetime.now()
time_took = time_end - time_start

print(f"All the emails were sent for {time_took.seconds} seconds.")




