""" General utility functions (data transformation & email sending) """

from typing import Dict, List
from dotenv import load_dotenv
import os
import smtplib
from email.message import EmailMessage

# loads environment variables from .env file
load_dotenv(encoding='utf8')

SENDER_EMAIL = os.getenv('SENDER_EMAIL')
APP_PASSWORD = os.getenv('APP_PASSWORD')
CSDESIGN_TRESURER = os.getenv('CSD_TRESURER_NAME')
CSD_TRESURER_PHONE = os.getenv('CSD_TRESURER_PHONE')

def filter_orders(data: Dict[str, str]) -> Dict[str, str]:
    filtered_data = []
    # remove lines that have a receipt & that are not orders.
    for line in data:
        if line["Type"] == "Prestation" and line["№ facture"] == "" and line["Encaissement"] == "":
            filtered_data.append(line)
    return filtered_data


def get_asso_lines(data_dicts, asso_name):
    """Filters the lines to return only those corresponding to an association"""
    asso_lines = []
    for line in data_dicts:
        if line["Bénéficiaire"].lower() == asso_name.lower():
            asso_lines.append(line)
    return asso_lines


def send_receipts_by_mail(recipient_first_name: str, recipient_email: str, asso_name:str, receipts_paths: List[str], orders_data: List[Dict[str, str]]):
    """Sends the receipts by email to an association"""
    
    subject = "Facture(s) CS Design"
    content = f"Hello {recipient_first_name},\n\n{len(orders_data)} prestation(s) ont été réalisées par CS Design pour l'association {asso_name} :\n"
    
    # add all receipts details
    for order in orders_data:
        content += f"- {order['Date']} : {order['Description']}, {order['Prix total']}\n"

    content += "\nTu trouveras en pièces jointes les factures correspondantes.\n\n"\
            + f"Bonne journée,\n{CSDESIGN_TRESURER}\nTrésorier de CS Design\n{CSD_TRESURER_PHONE}"
    msg = EmailMessage()
    msg['Subject'] = subject
    msg['From'] = SENDER_EMAIL
    msg['To'] = recipient_email
    msg.set_content(content)

    # add pdf receipts as attachments
    for receipt in receipts_paths:
        with open(receipt, 'rb') as f:
            file_data = f.read()
        msg.add_attachment(file_data, maintype="application", subtype="pdf", filename=receipt.split("\\")[-1])

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(SENDER_EMAIL, APP_PASSWORD)
        smtp.send_message(msg)