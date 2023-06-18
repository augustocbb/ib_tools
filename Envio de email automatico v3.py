#####################################################
# Title: Email Sender with Attachments
# Date: 18/05/23
# Author: Augusto Bastos
# Description: This code prompts the user to enter
#              values for the sender, recipients,
#              password, and signature image path
#              using input dialog boxes. It then
#              sends an email with attachments using
#              the entered values and displays the
#              attachment number for each email sent.
#####################################################

import smtplib
import os
import rarfile
import random
import tkinter as tk
from tkinter import simpledialog
from tkinter import Tk
from tkinter.filedialog import askopenfilenames
from urllib.parse import quote
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders

root = tk.Tk()
root.withdraw()

# Prompt for sender
sender = simpledialog.askstring("Input", "Enter sender:", parent=root)

# Prompt for recipients
recipients = simpledialog.askstring("Input", "Enter recipients (comma-separated):", parent=root)
recipients = recipients.split(",") if recipients else []

# Prompt for password
password = simpledialog.askstring("Input", "Enter password:", parent=root)

faturamento = 1500
qtde_produtos = 10
ticket_medio = faturamento / qtde_produtos

# Create the email message using the MIMEMultipart object
# and set the desired values for the Subject, Body, Sender, and Recipients fields
subject = "Email automático do Python"
# Add a placeholder for the signature image in the HTML body of the message
body = f"""
<p>Olá Lira, aqui é o código Python</p>

<p>O faturamento da loja foi de R${faturamento}</p>
<p>Vendemos {qtde_produtos} produtos</p>
<p>O ticket Médio foi de R${ticket_medio}</p>

<p>Abs,</p>
<p>Código Python</p>
<br>
<img src="cid:assinatura">
"""

assinatura = "C://Users/joaop/Downloads/assinatura.png"

# Create a helper function to get the size of a file in MB
def get_file_size(file):
    return os.path.getsize(file) / (1024 * 1024)

def load_connection_statement():
    connection_data = {}
    with open("connection_statement.txt", "r") as file:
        for line in file:
            key, value = line.strip().split("=")
            connection_data[key] = value
    return connection_data

def send_email(subject, body, sender, recipients, password):
    # Create a hidden window to select the files to be attached
    root = Tk()
    root.withdraw()
    # Use the askopenfilenames function to get a list of file names selected by the user
    anexos = askopenfilenames()
  
    # Create an empty list to store the names of the attached files
    attached_files = []
    # Create a variable to store the total size of the attachments in MB
    total_size = 0

    # Load connection statement from file
    connection_data = load_connection_statement()

    # Iterate through the list of attachments selected by the user
    for i, anexo in enumerate(anexos):
        # Check if the individual file size exceeds the limit (5 MB)
        file_size = get_file_size(anexo)
        if file_size > 5:
            # Create a new email message for each large attachment
            msg = MIMEMultipart()
            msg["Subject"] = f"{subject} [Attachment {i+1}]"
            msg["From"] = sender
            msg["To"] = ", ".join(recipients)
            msg.attach(MIMEText(body, "html"))

            with open(anexo, "rb") as attachment:
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(anexo)}")
                msg.attach(part)

            smtp_server = smtplib.SMTP(connection_data["SERVER"], connection_data["PORT"])
            smtp_server.starttls()
            smtp_server.login(sender, password)
            smtp_server.sendmail(sender, recipients, msg.as_string())
            smtp_server.quit()
            print(f"Email with {anexo} (Attachment {i+1}) sent")
        else:
            # Add the file name to the list of attached files
            attached_files.append(anexo)
            # Update the total size of the attachments with the size of the current file in MB
            total_size += file_size
            # If the total size of the attachments exceeds 5 MB or it is the last attachment in the list
            if total_size > 5 or i == len(anexos) - 1:
                # Create a rar file with the attached files list using the rarfile module and a random name
                rar_name = f"archive{random.randint(1000, 9999)}.rar"
                with rarfile.RarFile(rar_name, "w") as rar:
                    for file in attached_files:
                        rar.add(file)

                # Create a message with the same body and original subject
                msg = MIMEMultipart()
                msg["Subject"] = f"{subject} [Attachments {i+1-len(attached_files)}-{i+1}]"
                msg["From"] = sender
                msg["To"] = ", ".join(recipients)
                msg.attach(MIMEText(body, "html"))

                with open(assinatura, "rb") as img:
                    img_part = MIMEImage(img.read())
                    img_part.add_header("Content-ID", "<assinatura>")
                    msg.attach(img_part)

                with open(rar_name, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(rar_name)}")
                    msg.attach(part)

                smtp_server = smtplib.SMTP(connection_data["SERVER"], connection_data["PORT"])
                smtp_server.starttls()
                smtp_server.login(sender, password)
                smtp_server.sendmail(sender, recipients, msg.as_string())
                smtp_server.quit()
                print(f"Email with {rar_name} (Attachments {i+1-len(attached_files)}-{i+1}) sent")

                # Clear the list of attached files
                attached_files = []
                # Reset the total size of the attachments
                total_size = 0

send_email(subject, body, sender, recipients, password)
