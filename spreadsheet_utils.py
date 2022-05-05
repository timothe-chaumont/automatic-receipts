import os
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from typing import List, Dict
import string
from dotenv import load_dotenv

# loads environment variables from .env file
load_dotenv(encoding='utf8')

# permits to read, edit, create, and delete the spreadsheets in Google Drive
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')

# data ranges
FEATURES_LINE_RANGE = 'A2:R2'
ALL_RANGE = "A:S"

# CSD services
SERVICES_DATA = [{"designation": "Impression affiche A1", "price": 4.0},
                 {"designation": "Impression affiche A2", "price": 2.0},
                 {"designation": "Impression affiche A3", "price": 1.0},
                 {"designation": "Impression sticker", "price": 0.15},
                 {"designation": "Impression t-shirt", "price": 6.0}, ]


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


def get_spreadsheet(creds):
    """Returns the spreadsheet object"""
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    return sheet


# data fetching

def fetch_all_data(sheet, column_idx: Dict[str, int]) -> List[Dict[str, str]]:
    """Retrieves all the relevant data of the spreadsheet to limit the number of requests to a minimum.
       Later : return a numpy array

    Args:
        sheet (_type_): googleapiclient spreadsheets object
        column_idx (Dict[str, int]): dictionnary of column names -> column indexes

    Returns:
        A list of dictionnaries that represent lines 
    """
    # fetch all the lines data
    data = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                              range=ALL_RANGE).execute()['values']

    # remove headers / columns names
    col_names = data[1]

    # strip the headers and empty end lines
    i = 3
    while i < len(data) and data[i][0]:
        i += 1
    first_empty_line = i
    data = data[2:first_empty_line]

    # convert to dictionnary, with the relevant data
    data_dicts = []
    for line_nb, line in enumerate(data):
        line_dict = {'line': line_nb+3}
        for col_name in column_idx:
            col_id = column_idx[col_name]
            # if a cell is empty for a col :
            if col_id >= len(line):
                line_dict[col_name] = ""
            else:
                line_dict[col_name] = line[col_id]
        data_dicts.append(line_dict)

    return data_dicts


def find_column_index(columns_names, column_name: str) -> int:
    """Returns the index of the column corresponding to the input name"""
    for col_nb, col_name in enumerate(columns_names):
        if col_name == column_name:
            return col_nb
    raise Exception(f"Column '{column_name}' not found.")


def get_all_col_indexes(sheet) -> Dict[str, int]:
    """Returns a dictionnary of all the used columns in the spreadsheet"""
    columns_names = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                       range=FEATURES_LINE_RANGE).execute()['values'][0]
    col_indexes = {}
    for col_name in ["Date", "Type", "Inté / Exté", "Bénéficiaire", "Contact eventuel", "Description", "A1", "A2", "A3", "Sticker", "T-shirt", "Prix total", "№ facture", "Encaissement"]:
        col_indexes[col_name] = find_column_index(columns_names, col_name)
    return col_indexes


def find_last_line(sheet, date_col_idx: int):
    """Find the last order line
       We assume that the date is only filled out if there is an order
    """
    # express the recipient lines range
    date_col_letter = string.ascii_uppercase[date_col_idx]
    date_col_range = f"{date_col_letter}:{date_col_letter}"
    dates_list = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                    range=date_col_range).execute()['values']
    return len(dates_list)


def find_lines(recipient_name: str, sheet, recipient_col_idx: int) -> List[int]:
    """Returns the list of lines whitch correspond to an entry where it is the recipient

        ex : find_lines("Hyris") returns [5, 7, 21, 29, 32, 34, 60, 67, 72]
    """
    # express the recipient lines range
    asso_col_letter = string.ascii_uppercase[recipient_col_idx]
    asso_name_col_range = f"{asso_col_letter}:{asso_col_letter}"
    # get the list of entry's recipient names
    recipients_list = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                         range=asso_name_col_range).execute()['values']
    recipient_lines = []
    for line, recip in enumerate(recipients_list):
        if recip and recip[0]. lower() == recipient_name.lower():
            recipient_lines.append(line+1)
    return recipient_lines


def get_no_receipt_lines(sheet, columns_idx) -> List[int]:
    """Returns the list of lines where there is no receipt number"""
    no_receipt_lines = []
    nb_lines = find_last_line(sheet, columns_idx["Date"])
    # express the receipts number range
    receipt_col_letter = string.ascii_uppercase[columns_idx['№ facture']]
    receipt_col_range = f"{receipt_col_letter}3:{receipt_col_letter}{nb_lines}"
    # retreive the data from the spreadsheet
    receipt_nb_list = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                         range=receipt_col_range).execute()['values']
    # add the last orders with no receipts (because the list stops at the last order number)
    receipt_nb_list += [[] for _ in range(nb_lines - 2 - len(receipt_nb_list))]
    # list all the lines where there is no receipt number
    for line, receipt_nb in enumerate(receipt_nb_list):
        if not receipt_nb:
            no_receipt_lines.append(line + 3)
    return no_receipt_lines


def format_price_string(price: float) -> str:
    """Returns a string corresponding to the input price """
    if int(10 * price) == 10 * price:
        return f"{str(float(price)).replace('.', ',')}0€"
    else:
        return f"{str(float(price)).replace('.', ',')}€"


def has_receipt(line_nb: int, sheet, columns_idx) -> bool:
    """Returns True if the entry has a receipt number, False otherwise"""
    line_data = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                   range=f'A{line_nb}:T{line_nb}').execute()['values'][0]
    if columns_idx["№ facture"] < len(line_data):
        return line_data[columns_idx["№ facture"]] != ""
    return False


def get_order_data(line_nb: int, sheet, column_idx):
    """Returns a dictionnary with the needed data to generate a receipt for a given line in the sheet
        and the toal price of the order
    """
    line_data = sheet.values().get(spreadsheetId=SPREADSHEET_ID,
                                   range=f'A{line_nb}:T{line_nb}').execute()['values'][0]
    assert line_data[column_idx["Type"]
                     ] == 'Prestation', "This line does not correspond to an order"

    # get various order data
    recipient_name = line_data[column_idx["Bénéficiaire"]]
    total_print_price = line_data[column_idx["Prix total"]] + " TTC"

    orders_list = []
    # iterate over each type of print (A1, A2, A3, sticker, t-shirt)
    for q_id, qtity_st in enumerate(line_data[column_idx["A1"]:column_idx["Prix total"]]):
        # if there is an order
        if qtity_st:
            # compute the total price of each service type
            element_price = int(qtity_st) * SERVICES_DATA[q_id]["price"]
            order_dict = {"quantity": qtity_st, "designation": SERVICES_DATA[q_id]["designation"],
                          "unit price": format_price_string(SERVICES_DATA[q_id]["price"]) + " HT",
                          "line total price": format_price_string(element_price) + " TTC"}
            orders_list.append(order_dict)
    return orders_list, total_print_price, recipient_name


# write to the spreadsheet
def write_receipt_number(receipt_nb: str, line: int, sheet, columns_idx):
    """Writes the receipt number in the spreadsheet
        TODO : change scope, (remove readonly) and rerun ...

    """
    receipt_col_letter = string.ascii_uppercase[columns_idx['№ facture']]
    receipt_nb_cell_id = f"{receipt_col_letter}{line}"
    sheet.values().update(spreadsheetId=SPREADSHEET_ID,
                          range=receipt_nb_cell_id,
                          valueInputOption="RAW",
                          body={"values": [[receipt_nb]]}).execute()


if __name__ == '__main__':
    creds = connect_to_spreadsheet()
    sheet = get_spreadsheet(creds)
