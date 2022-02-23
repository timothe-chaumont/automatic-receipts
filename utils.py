""" General utility functions (data transformation) """

from typing import Dict


def filter_orders(data: Dict[str, str]) -> Dict[str, str]:
    filtered_data = []
    # remove lines that have a receipt & that are not orders.
    for line in data:
        if line["Type"] == "Prestation" and line["№ facture"] == "" and line["Encaissement"] == "":
            filtered_data.append(line)
    return filtered_data