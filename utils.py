""" General utility functions (data transformation) """

from typing import Dict


def filter_orders(data: Dict[str, str]) -> Dict[str, str]:
    # remove lines that have a receipt & that are not orders.
    return list(line for line in data if line["Type"] == "Prestation" and line["№ facture"] == "")