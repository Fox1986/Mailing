#!/usr/bin/env python
# -*- coding: utf-8 -*-

# ----------------------------------------------------------------------------------------------------------------------#

# Titel:            mailing.py
# Beschreibung:     Versenden von E-Mails mittels Python
# Autor:            Hinrik Taeger
# Version:          1.0
# Kategorie:        Support
# Ziel:             MacOS

# Funktion zum Senden von Mails über Python
# Mail-Konfigurationen müssen hinterlegt werden
# Kontakte werden in einer JSON-Datei hinterlegt
# Hauptfunktion benötigt einen Empfänger, einen Betreff, eine Nachricht und ggf. Anhänge

# ----------------------------------------------------------------------------------------------------------------------#

#                           Importe

import json
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# ----------------------------------------------------------------------------------------------------------------------#

#                           Funktionen


# Erstellen einer Vorlagen-Datei für Kontakte
def create_contacts():
    if not os.path.isfile("contacts.json"):
        with open("contacts.json", "w") as new_contacts:
            new_contacts.write("{\n")
            new_contacts.write('"contacts": [\n')
            new_contacts.write("{")
            new_contacts.write('"name":"John",\n')
            new_contacts.write('"mail":"johndoe@xyz.com"\n')
            new_contacts.write("}\n")
            new_contacts.write("]\n")
            new_contacts.write("}\n")
        return 1
    else:
        return 0


# Auslesen der Kontakte
def read_contacts():
    file = open("contacts.json")
    data = json.load(file)
    for contact in data["contacts"]:
        yield contact["name"], contact["mail"]


# Erstellen einer Vorlage-Datei für die notwendigen SMTP-Anmelde-Daten
def create_smtp_config():
    if not os.path.isfile("smtp.conf"):
        with open("smtp.conf", "w") as new_config:
            new_config.write("server smtp.XYZ.de\n")
            new_config.write("port 587\n")
            new_config.write("login_name janedoe@XYZ.com\n")
            new_config.write("login_pass 123456\n")
        return 1
    else:
        return 0


# Hauptfunktion für das Versenden von Mails über Python
def send_mail(receiver_mail, subject, message, attachment):
    # Testen ob die notwendigen Config-Daten vorliegen
    test_config = create_smtp_config()
    if test_config == 0:
        # Einlesen der Conig-Daten, um sich mit dem SMTP-Server zu verbinden
        with open("smtp.conf", "r") as config:
            # Serverdaten
            smtpServer = config.readline().replace("server", "").strip()
            smtpPort = config.readline().replace("port", "").strip()

            # Zugangsdaten
            username = config.readline().replace("login_name", "").strip()
            password = config.readline().replace("login_pass", "").strip()
        # Sender & Empfänger
        sender = username
        receiver = receiver_mail

        # MSG Objekt erzeugen
        msg = MIMEMultipart()
        msg['Subject'] = subject
        msg['From'] = sender
        msg['To'] = receiver

        # HTML Text
        html_part = MIMEText(message, 'html')
        msg.attach(html_part)

        # Die Attachment-Liste wird hier Datei für Datei angehängt
        if len(attachment) > 0:
            for file in attachment:
                file_part = MIMEBase('application', "octet-stream")
                file_part.set_payload(open(file, "rb").read())
                encoders.encode_base64(file_part)
                file_part.add_header('Content-Disposition', 'attachment; filename="' + file + '"')
                msg.attach(file_part)

        # Erzeugen einer Mail Session
        smtpObj = smtplib.SMTP(smtpServer, int(smtpPort))
        # Debuginformationen auf der Konsole ausgeben
        smtpObj.set_debuglevel(1)
        # Wenn der Server eine Authentifizierung benötigt dann...
        smtpObj.starttls()
        smtpObj.login(username, password)

        # absenden der E-Mail
        smtpObj.sendmail(sender, receiver, msg.as_string())
    else:
        print("No config data found. Please fill smtp.conf")


# ----------------------------------------------------------------------------------------------------------------------#

#                           MAIN


if __name__ == '__main__':
    # Liste für Dateianhänge
    attach = ["test.txt"]
    # Betreff
    subj = "Hallo Welt!"
    # HTML-Text
    text = "<html><body>" \
           "<h1>Eine Überschrift</h1>" \
           "<p style='color:red'>Ein Text in der Schriftfarbe rot</p>" \
           "<div style='background-color:green'>Hier ist ein Text mit einer Hintergrundfarbe.</div>" \
           "</body></html>"
    # Prüfen, ob es Kontakte gibt
    test_contact_file = create_contacts()
    # Senden der Mails an alle Kontakte
    if test_contact_file == 0:
        contacts = read_contacts()
        for c in contacts:
            send_mail(c[1], subj, text, attach)
    else:
        print("No contacts found. Please fill contacts.json")
