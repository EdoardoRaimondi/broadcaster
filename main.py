"""
Main program file.
From a .xls file, it extracts contact references.
Then it send them a custom message created by the user via UI.
"""

import pandas as pd
import smtplib
from email.message import EmailMessage

"""
Hyperparameters
"""
colNames = ["Cliente", "Acquisto", "Comm"]


def extractContacts(file_name, columns):
    """
    Function to extract the interested contact fields
    :param file_name
    :param columns: containing meaningful info
    """
    df = pd.read_excel(file_name, header=4, usecols=columns, names=colNames)
    df.dropna(inplace=True)
    df.groupby("Cliente")
    print(df.head(50))


def send_message(sender, password, receiver, subject, message):
    """
    Function to send an email
    :param sender: email address. Must be a gmail one
    :param receiver: email address
    :param password: of sender
    :param subject: of the message to send
    :param message: to send
    """
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = receiver
    msg.set_content(message)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:  # Added Gmails SMTP Server
        smtp.login(sender, password)  # This command Login SMTP Library using your GMAIL
        smtp.send_message(msg)  # This Sends the message


send_message("edoardoraimondi3@gmail.com", "tumiosa1998", "oscar.raimondi@libero.it", "prova", "prova")