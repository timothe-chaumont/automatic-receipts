import datetime as dt
import os
import json
from typing import Set
from dotenv import load_dotenv

# loads environment variables from .env file
load_dotenv(encoding='utf8')

RECEIPTS_PATH = os.getenv('RECEIPTS_PATH')


# receipt content uitility functions

def get_asso_address(asso_name: str):
    """Returns the address of the association"""
    # read the json file
    with open("associations_addresses.json", "r", encoding='utf-8') as f:
        data_dict = json.load(f)
    if asso_name.lower() not in data_dict:
        raise Exception(
            f"{asso_name}'s official name and address not found in the json file")
    asso_data = data_dict[asso_name.lower()]
    return asso_data["official name"], asso_data["address"], asso_data["tresurer first name"], asso_data["tresurer mail"]


def get_this_months_dir_name():
    """Returns the name of the directory for this month"""
    today = dt.date.today()
    return today.strftime("%Y-%m")


def get_receipt_name(dir_name, receipt_number):
    """Builds and returns the receipt's name : yyyy-mm-receipt_number """
    assert receipt_number <= 9999, "The receipt number is too large"
    str_number = "0" * (4 - len(str(receipt_number))) + str(receipt_number)
    return dir_name + "-" + str_number


def get_receipt_number(sheet_receipts_names: Set[str], receipts_dir: str = RECEIPTS_PATH) -> str:
    """Checks how many receipts have been created this month (if any)
        and returns the number (/ name) of the next receipt to be done

        Now also checks the numbers written on the sheets of the receipts

        Args:
            sheet_receipts_names (Set[str]): the list of the names of the receipts written in the online sheet
    """
    # list all the receipts directories in the path
    month_dirs = os.listdir(receipts_dir)

    # get this month's directory name
    todays_dir_name = get_this_months_dir_name()
    if todays_dir_name in month_dirs:
        # list the receipts already created
        receipt_files = os.listdir(os.path.join(receipts_dir, todays_dir_name))
        # find max number of all receipt files in the dir
        max_number = 0
        for file_name in receipt_files:
            if ".pdf" in file_name:  # do not use .docx files
                if int(file_name.split(".pdf")[0][-2:]) > max_number:
                    max_number += 1
        receipt_number = max_number + 1
        next_r_name = get_receipt_name(todays_dir_name, receipt_number)
    else:
        # create a new directory
        os.mkdir(os.path.join(RECEIPTS_PATH, todays_dir_name))
        print(f"Directory {todays_dir_name} created")
        # get the first receipt's name
        receipt_number = 1
        next_r_name = get_receipt_name(todays_dir_name, receipt_number)

    # check if this receipt number is not already used in the sheet
    # (find the first one which number is not used)
    while next_r_name in sheet_receipts_names:
        receipt_number += 1
        next_r_name = get_receipt_name(
            todays_dir_name, receipt_number)

    return next_r_name


# receipt doc creation utility functions

def make_rows_bold(*rows) -> None:
    """Set text weight to bold for each cell of the input rows"""
    for row in rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.bold = True


def set_col_width(col, width) -> None:
    """Sets the width of a table's column by setting each cell's width """
    for cell in col.cells:
        cell.width = width


def set_style(style, font_name: str, font_size, font_color: str, bold: bool = False):
    """Sets the style properties """
    font = style.font
    font.name = font_name
    font.size = font_size
    font.color.rgb = font_color
    font.bold = bold
