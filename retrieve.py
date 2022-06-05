"""  Class for data orders data retrieval & cleanup
"""

import os
import re
from typing import Dict, List, Tuple, Optional
import json

import pandas as pd
from dotenv import load_dotenv
from typing_extensions import Literal
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from pyparsing import Optional


load_dotenv(encoding='utf8')

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')

# data ranges (the columns names span between A2 and R2)
FEATURES_LINE_RANGE = 'A2:R2'
ALL_RANGE = "A:S" 

ASSO_DETAILS_FEATURES = ['official name', 'address', 'tresurer first name', 'tresurer mail']


def connect_to_spreadsheet(scopes=SCOPES):
    """Returns the credentials to connect to the accounting spreadsheet."""
    # auth
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', scopes)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        # if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', scopes)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds


class Retriever():
    def __init__(self) -> None:
        self.creds = connect_to_spreadsheet(SCOPES)
        self.fetch_orders_data(self.creds)
        self.fetch_asso_details()

    def fetch_orders_data(self, creds) -> pd.DataFrame:
        """Returns all the orders data as a panda Dataframe
        """
        # fetch the spreadsheet object
        service = build('sheets', 'v4', credentials=creds)
        sheet = service.spreadsheets()

        # fetch the data & convert it to a panda Dataframe
        data = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                        range=ALL_RANGE).execute()['values']
        df = pd.DataFrame(data[2:], columns=data[1])

        # later : remove empty lines at the end 
        self.orders = df
        return self.orders

    def fetch_asso_details(self) -> pd.DataFrame:
        with open("associations_addresses.json", "r", encoding='utf-8') as f:
            data_dict = json.load(f)
        # convert to df (later : store as csv instead of json)
        self.asso_details = pd.DataFrame([data_dict[a].values() for a in data_dict], columns=ASSO_DETAILS_FEATURES, index=data_dict.keys())
        return self.asso_details

    def get_unprocessed_orders(self, orders = None) -> pd.DataFrame:
        """Returns a DataFrame containing only the service (='prestations') lines that are either unpaid or that don't have any receipts
        """
        if orders is None:
            orders = self.orders

        # get boolean series corresponding to 
        presta_orders = orders["Type"] == "Prestation"
        orders_wo_receipt =  orders["№ facture"].isnull() | orders["№ facture"].eq("")
        unpaid_orders = ~ orders["Encaissement"].isin(['Virement','Lydia Pro', 'Chèque'])

        return orders[presta_orders & orders_wo_receipt & unpaid_orders]

    def filter_unknown_assos(self, orders:pd.DataFrame) -> pd.DataFrame:
        """Drops the order lines of assos for which we don't know the information.

            Args:
                orders (pd.DataFrame) : dataframe containing associations orders data
        """
        known_assos = self.asso_details.index.tolist()
        orders_assos = orders['Bénéficiaire'].str.lower()

        return orders[orders_assos.isin(known_assos)]

    def filter_invalid_mails(self, orders:pd.DataFrame) -> pd.DataFrame:
        """Returns the orders dataframe with only the lines that have valid email addresses.
        """
        # use a regular expression to match valid email addresses
        pattern=r"^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$"
        return orders[orders['Contact eventuel'].map(lambda i: bool(re.match(pattern, i)))]

    def filter_by_client_type(self, orders:pd.DataFrame, recipient_type:Literal["Asso", "Inté", "Exté"]) -> pd.DataFrame:
        """Returns the orders dataframe with only the lines that have the given recipient type.
        """
        return orders[orders["Inté / Exté"] == recipient_type]

    def filter_by_recipient_name(self, orders:pd.DataFrame, name:str) -> pd.DataFrame:
        """Returns the orders dataframe with only the lines that correspond to the given recipient.
        """
        return orders[orders["Bénéficiaire"].str.lower() == name.lower()]

    def get_asso_details(self, name:str) -> pd.Series:
        """Returns the details of an association.

            Args:
                name (str) : name of the association
        """
        if name.lower() in self.asso_details.index:
            return self.asso_details.loc[name.lower(), ASSO_DETAILS_FEATURES]
        
        raise Exception(f"{name} not found in the list of associations")


if __name__ == "__main__":
    r = Retriever()
    orders = r.orders
    u_orders = r.get_unprocessed_orders()
