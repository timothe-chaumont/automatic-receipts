""" Main code that reads from the spreadsheet and produces a receipt. """

import argparse
from typing import Dict, List, Union

import pandas as pd

import receipt_creation as rc
import receipt_utils as ru
import spreadsheet_utils as su
import utils as ut

# TODO:
# - when counting the lines for an association, only count those without a receipt
# - add a function to list all the associations and their orders with no receipt (or not paid)
# - add space before euro symbol receipt


def process_everything():
    from retrieve import Retriever
    retriever = Retriever()
    all_orders = retriever.orders

    # get unprocessed receipts
    to_be_processed = retriever.get_unprocessed_orders()
    print(f"{len(to_be_processed)} commande(s) sans facture ni paiement.")

    # retrieve the asso orders that could be processed
    asso_orders = retriever.filter_by_client_type(to_be_processed, "Asso")
    can_be_processed_asso = retriever.filter_unknown_assos(asso_orders)
    # Same for individual & extern orders
    valid_email_orders = retriever.filter_invalid_mails(to_be_processed)
    can_be_processed_indiv = retriever.filter_by_client_type(valid_email_orders, "Inté")
    can_be_processed_etern = retriever.filter_by_client_type(valid_email_orders, "Exté")

    # group them into one array
    can_be_processed = pd.concat([can_be_processed_asso, can_be_processed_indiv, can_be_processed_etern])
    # list all the different recipientsof the orders that will be processed (to group the orders by recipient)
    unique_recipient_names = can_be_processed['Bénéficiaire'].unique()

    # print the overview and ask the user to confirm
    print(f"{len(can_be_processed)} prestations peuvent être traitées, dont:")
    print(f" - {len(can_be_processed_asso)} pour des associations,")
    print(f" - {len(can_be_processed_indiv)} pour des étudiants,")
    print(f" - {len(can_be_processed_etern)} pour des clients extérieurs.")

    answer = input("Voulez-vous continuer? (o/n)\n")
    if answer.lower() == 'o':
        print("Let's go!")
    else:
        print("Ok!")
        return

    # for each asso, process the orders
    for recip_name in unique_recipient_names:
        
        # get the orders for this recipient
        recip_orders = retriever.filter_by_recipient_name(can_be_processed, recip_name)

        # check if it is an association or other
        recip_type = recip_orders["Inté / Exté"].iloc[0]
        if recip_type == "Asso":
            # get asso details (address, email, etc.)
            asso_details = retriever.get_asso_details(recip_name)
            # get recipient data that will be used to create the receipt
            # is different depending on the recipient type
            receipt_recipient_info = asso_details["official name"] + "\n" + asso_details["address"]

        # if it is an individual or an extern order
        else:
            receipt_recipient_info = recip_orders["Bénéficiaire"].iloc[0]

        # will store the pdf receipts paths to attach them to emails
        receipts_paths = []

        # process the orders
        for order_idx, order in recip_orders.iterrows():

            # get the order(s) details
            total_print_price = order["Prix total"] + " TTC"

            # will store the details of the order
            orders_list = []

            # iterate over each type of print (A1, A2, A3, sticker, t-shirt)
            for q_id, qtity_st in enumerate(list(order[order_key] for order_key in ("A1", "A2", "A3", "Sticker", "T-shirt"))):
                # if there is an order
                if qtity_st:
                    # compute the total price of each service type
                    element_price = int(qtity_st) * \
                        su.SERVICES_DATA[q_id]["price"]
                    order_dict = {"quantity": qtity_st, "designation": su.SERVICES_DATA[q_id]["designation"],
                                  "unit price": su.format_price_string(su.SERVICES_DATA[q_id]["price"]) + " HT",
                                  "line total price": su.format_price_string(element_price) + " TTC"}
                    orders_list.append(order_dict)

            # get the already created receipt numbers
            sheet_receipt_names = all_orders["№ facture"].unique()
            receipt_nb = ru.get_receipt_number(sheet_receipt_names)

            docx_file_name = rc.build_receipt_path(
                ru.RECEIPTS_PATH, ru.get_this_months_dir_name(), receipt_nb + ".docx")

            rc.create_receipt_docx(
                receipt_recipient_info, orders_list, receipt_nb, total_print_price, docx_file_name)
            pdf_file_name = rc.build_receipt_path(
                ru.RECEIPTS_PATH, ru.get_this_months_dir_name(), receipt_nb + ".pdf")

            # export to pdf
            rc.export_receipt_to_pdf(docx_file_name, pdf_file_name)
            print(f" - Gerenated receipt {receipt_nb}.")
            receipts_paths.append(pdf_file_name)

            # update the spreadsheet
            retriever.write_receipt_number(receipt_nb, order_idx)

        # send an email with the receipts attached
        if order["Inté / Exté"] == "Asso":
            recip_mail = asso_details["email"]
            recipient_first_name = asso_details["tresurer first name"]

        else:
            recip_mail = order["Contact eventuel"]
            recipient_first_name = None

        ut.send_receipts_by_mail(
            recip_name, 
            recip_mail, 
            recip_type, 
            receipts_paths, 
            recip_orders, 
            recipient_first_name
        )


def main(args):
    """ Main function that reads from the spreadsheet and produces a receipt. """

    creds = su.connect_to_spreadsheet()
    spreadsheet = su.get_spreadsheet(creds)
    # retrieve column indexes
    col_indexes = su.get_all_col_indexes(spreadsheet)
    # retrieve all the data from the spreadsheet
    data = su.fetch_all_data(spreadsheet, col_indexes)
    # filter out the lines that already have a receipt or have been paid
    filtered_data = ut.filter_processed_orders(data)

    if args.everything_association:
        process_all_associations(spreadsheet, col_indexes)

    elif args.everything_individual:
        process_all_individuals(spreadsheet, col_indexes)

    elif args.summary:
        print("----- Summary of the spreadsheet -----\n")

        print(f"{len(filtered_data)} commandes sans factures ni paiement:")
        # count the number of ~unpaid receipts for associations and individuals
        nb_asso_filtered_receipt = len(list(1 for i in range(
            len(filtered_data)) if filtered_data[i]['Inté / Exté'] == 'Asso'))
        nb_person_filtered_receipt = len(list(1 for i in range(
            len(filtered_data)) if filtered_data[i]['Inté / Exté'] == 'Inté'))
        print(f" - {nb_asso_filtered_receipt} pour des assos.")
        print(f" - {nb_person_filtered_receipt} pour des personnes.")

        # print the top associations with the most unpaid orders
        print("\n----- Top associations -----\n")

        # get the top associations
        assos = dict()
        for i, line in enumerate(filtered_data):
            if line['Inté / Exté'] == 'Asso':
                assos[line['Bénéficiaire']] = assos.get(
                    line['Bénéficiaire'], 0) + 1
        assos = sorted(assos.items(), key=lambda x: x[1], reverse=True)
        for i in range(min(5, len(assos))):
            print(f"{i+1}. {assos[i][0]} ({assos[i][1]} commandes)")

        # print the top individuals with the most unpaid orders
        print("\n----- Top individuals -----\n")

        # get the top individuals
        individuals = dict()
        for i, line in enumerate(filtered_data):
            if line['Inté / Exté'] == 'Inté':
                individuals[line['Bénéficiaire']] = individuals.get(
                    line['Bénéficiaire'], 0) + 1
        individuals = sorted(individuals.items(),
                             key=lambda x: x[1], reverse=True)
        for i in range(min(5, len(individuals))):
            print(
                f"{i+1}. {individuals[i][0]} ({individuals[i][1]} commandes)")

        # print the top externs with the most unpaid orders
        print("\n----- Top extern -----\n")

        # get the top externs
        externs = dict()
        for i, line in enumerate(filtered_data):
            if line['Inté / Exté'] == 'Exté':
                externs[line['Bénéficiaire']] = externs.get(
                    line['Bénéficiaire'], 0) + 1
        externs = sorted(externs.items(),
                             key=lambda x: x[1], reverse=True)
        for i in range(min(5, len(externs))):
            print(
                f"{i+1}. {externs[i][0]} ({externs[i][1]} commandes)")

    # if an association was given
    elif args.association:
        asso_name = " ".join(args.association)

        # find lines corresponding to the association orders
        # then keep only the orders of the given association
        asso_data = ut.get_asso_lines(filtered_data, asso_name)

        entry_lines = su.find_lines(
            asso_name, spreadsheet, col_indexes["Bénéficiaire"])

        print(f"{len(entry_lines)} ligne(s) trouvée(s) pour {asso_name}.")

        # store the paths to send them by email
        receipts_paths = []

        for line in entry_lines:
            # check if the receipt has already been created
            if not su.has_receipt(line, spreadsheet, col_indexes):

                # later refactor to use the data already fetched
                orders_list, total_print_price, recipient_name = su.get_order_data(
                    line, spreadsheet, col_indexes)
                asso_official_name, asso_address, tresurer_first_name, tresurer_email = ru.get_asso_address(
                    recipient_name)
                # get the already created receipt numbers
                sheet_receipt_names = set(
                    data[i]["№ facture"] for i in range(len(data)))
                receipt_nb = ru.get_receipt_number(sheet_receipt_names)

                recipient_info = asso_official_name + "\n" + asso_address
                docx_file_name = rc.build_receipt_path(
                    ru.RECEIPTS_PATH, ru.get_this_months_dir_name(), receipt_nb + ".docx")

                rc.create_receipt_docx(
                    recipient_info, orders_list, receipt_nb, total_print_price, docx_file_name)
                pdf_file_name = rc.build_receipt_path(
                    ru.RECEIPTS_PATH, ru.get_this_months_dir_name(), receipt_nb + ".pdf")
                # export to pdf
                rc.export_receipt_to_pdf(docx_file_name, pdf_file_name)
                print(f"Facture {receipt_nb} générée.")
                receipts_paths.append(pdf_file_name)

                # update the spreadsheet
                su.write_receipt_number(
                    receipt_nb, line, spreadsheet, col_indexes)

        # if there are receipts to send
        if args.mail and asso_data != []:
            # send an email with the receipts attached
            ut.send_receipts_by_mail(
                tresurer_first_name, tresurer_email, recipient_name, receipts_paths, asso_data)
            print(f"Email sent to {tresurer_email}")

    elif args.individual:
        individual_name = " ".join(args.individual)
        # get the lines corresponding to individuals
        individual_data = list(
            line for line in filtered_data if line['Inté / Exté'] == 'Inté')
        # find the line corresponding to the individual
        individual_lines = list(i for i, line in enumerate(data)
                                if line['Bénéficiaire'] == individual_name and line['Inté / Exté'] == 'Inté')
        print(individual_lines)

        # find the lines corresponding to the individual orders


if __name__ == '__main__':

    # use a parser to get the arguments
    parser = argparse.ArgumentParser()
    parser.add_argument("-a", "--association", help="Process all entries for the given association.",
                        type=str, nargs="+")  # at least one word
    parser.add_argument("-ea", "--everything_association", help="Process all entries of type Association.",
                        action='store_true')  # at least one word
    parser.add_argument("-ei", "--everything_individual", help="Process all entries of type Inté.",
                        action='store_true')  # at least one word
    parser.add_argument("-i", "--individual", help="Process all entries for a given individual.",
                        type=str, nargs="+")  # at least one word (firstname, lastname)
    parser.add_argument("-m", "--mail", help="Send automatically the receipts by email.",
                        action='store_true')  # no arguments
    parser.add_argument("-s", "--summary", help="Prints a description of the current state of the spreadsheet.",
                        action='store_true')  # no arguments needed
    args = parser.parse_args()

    # main(args)
    process_everything()
