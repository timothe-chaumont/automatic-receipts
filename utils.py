""" General utility functions (data transformation & email sending) """

from typing import Dict, List
from dotenv import load_dotenv
import os
import smtplib
from email.message import EmailMessage
import re

import receipt_utils as ru


# loads environment variables from .env file
load_dotenv(encoding='utf8')

SENDER_EMAIL = os.getenv('SENDER_EMAIL')
APP_PASSWORD = os.getenv('APP_PASSWORD')
CSDESIGN_TRESURER = os.getenv('CSD_TRESURER_NAME')
CSD_TRESURER_PHONE = os.getenv('CSD_TRESURER_PHONE')

# data processing

def filter_processed_orders(data: Dict[str, str]) -> List[Dict[str, str]]:
    """Returns order lines that have no receipt and that have not been paid"""
    filtered_data = []
    # remove lines that have a receipt & that are not orders.
    for line in data:
        if line["Type"] == "Prestation" and line["№ facture"] == "" and line["Encaissement"] == "":
            filtered_data.append(line)
    return filtered_data


def filter_individuals_orders(filtered_data:List[Dict[str, str]]) -> List[Dict[str, str]]:
    """Given the orders that could be processed, returns the ones that correspond to individuals
        Args:
            filtered_data (List[Dict[str, str]]): a list of dictionaries containing the data of the orders to process
    """
    # 1. Filter the orders that can be processed (for which we have an email address)
    can_be_processed = []
    # for each line,
    for line in filtered_data:
        # print(line['Inté / Exté'], line['Bénéficiaire'])
        # if it is not an association
        if line['Inté / Exté'] == 'Inté':
            # check if an email address is provided
            contact_data = line['Contact eventuel']
            # regular expression supposed to match only email addresses
            if re.match(
                    r"^[A-Za-z0-9\.\+_-]+@[A-Za-z0-9\._-]+\.[a-zA-Z]*$", contact_data
            ):
                # add the line to the list that can be processed
                can_be_processed.append(line)
    return can_be_processed


def filter_individuals_orders(filtered_data:List[Dict[str, str]]) -> List[Dict[str, str]]:
    """Given the orders that could be processed, returns the ones that correspond to individuals

        Args:
            filtered_data (List[Dict[str, str]]): a list of dictionaries containing the data of the orders to process
    """
    # 1. Filter the orders that can be processed (for which we have an email address)
    can_be_processed = []
    # for each line,
    for line in filtered_data:
        # print(line['Inté / Exté'], line['Bénéficiaire'])
        # if it is not an association
        if line['Inté / Exté'] == 'Inté':
            # check if an email address is provided
            contact_data = line['Contact eventuel']
            # regular expression supposed to match only email addresses
            if re.match(
                    r"^[A-Za-z0-9\.\+_-]+@[A-Za-z0-9\._-]+\.[a-zA-Z]*$", contact_data
            ):
                # add the line to the list that can be processed
                can_be_processed.append(line)
    return can_be_processed


def filter_assos_orders(filtered_data:List[Dict[str, str]]) -> List[Dict[str, str]]:
    """Given the orders that could be processed, returns the ones that correspond to associations
    
        Args:
            filtered_data (List[Dict[str, str]]): a list of dictionaries containing the data of the orders to process
    """
    can_be_processed = []
    # for each line,
    for line in filtered_data:
        # if it is not an association
        if line['Inté / Exté'] != 'Asso':
            # skip to the next order
            continue
        # check if the association is already in the list of associations
        # try to get the information about the association
        try:
            # asso_official_name, asso_address, tresurer_first_name, tresurer_email =
            ru.get_asso_address(line['Bénéficiaire'])
            # add the association to the list that can be processed
            can_be_processed.append(line)
        # if the information was not found, just pass to the next line
        except Exception as e:
            continue

    return can_be_processed


def get_asso_lines(data_dicts, asso_name):
    """Filters the lines to return only those corresponding to an association"""
    asso_lines = []
    for line in data_dicts:
        if line["Bénéficiaire"].lower() == asso_name.lower():
            asso_lines.append(line)
    return asso_lines


def send_receipts_by_mail(recipient_first_name: str, recipient_email: str, asso_name: str, receipts_paths: List[str], orders_data: List[Dict[str, str]], recipient_type: str = 'asso'):
    """Sends the receipts by email to an association"""

    subject = "Facture(s) CS Design"
    content = f"Hello {recipient_first_name},\n\n{len(orders_data)} prestation(s) ont été réalisées par CS Design pour"

    if recipient_type == "asso":
        content += f" l'association {asso_name} :\n"
    else:
        content += f" toi :\n"

    # add all receipts details
    for order in orders_data:
        content += f"- {order['Date']} : {order['Description']}, {order['Prix total']}\n"

    content += "\nTu trouveras en pièces jointes les factures correspondantes.\nMerci de confirmer le paiement de ces facture(s) en répondant à ce mail.\n\n"\
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
        msg.add_attachment(file_data, maintype="application",
                           subtype="pdf", filename=receipt.split("\\")[-1])

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(SENDER_EMAIL, APP_PASSWORD)
        smtp.send_message(msg)

