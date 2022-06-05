""" Main code that reads from the spreadsheet and produces a receipt. """

import argparse
from typing import Dict, List, Union

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

    # get unprocessed receipts
    to_be_processed = retriever.get_unprocessed_orders()
    print(len(to_be_processed), "unprocessed orders")

    ## 1. process the associations

    # retrieve the asso orders that could be processed
    asso_orders = retriever.filter_by_client_type(to_be_processed, "Asso")
    can_be_processed = retriever.filter_unknown_assos(asso_orders)
    unprocessed_assos_names = can_be_processed['Bénéficiaire'].unique()
    print("assos", len(can_be_processed))

    # for each asso, process the orders
    for asso_name in unprocessed_assos_names:
        # get asso details (address, email, etc.)
        asso_details = retriever.get_asso_details(asso_name)
        
        # get the orders for this asso
        asso_orders = retriever.filter_by_recipient_name(can_be_processed, asso_name)

        # process the orders
        for order_idx, order in asso_orders.iterrows():
            pass
        
            # create the receipt
            # receipt = rc.create_receipt(order, asso_details)
            # # save the receipt
            # rc.save_receipt(receipt)

    # # 2. process the individuals
    # indiv_orders = retriever.filter_by_client_type(to_be_processed, "Inté")
    # can_be_processed = retriever.filter_invalid_mails(indiv_orders)
    # print("inté", len(can_be_processed))

    # # 3. process the extern orders
    # extern_orders = retriever.filter_by_client_type(to_be_processed, "Exté")
    # can_be_processed = retriever.filter_invalid_mails(extern_orders)
    # print("exté", len(can_be_processed))




def process_all_associations(spreadsheet, col_indexes:  Dict[str, int]):
    """ For all associations-related unpaid orders, see which ones can be processed (we have enough information)
        Print an overview of what will be done
        If the user wants to continue, process the orders

        Args:
            filtered_data (List[Dict[str, Union[str, int]]]): List of the associations lines from the spreadsheet that are unpaid or don't have a receipt
            col_indexes ( Dict[str, int]): List of the column indexes of the spreadsheet
    """

    # retrieve all the data from the spreadsheet
    data = su.fetch_all_data(spreadsheet, col_indexes)

    # filter out the lines that already have a receipt or have been paid
    filtered_data = ut.filter_processed_orders(data)

    # 1. Filter the orders that can be processed
    processable = ut.filter_assos_orders(filtered_data)

    if len(processable) == 0:
        print("No association can be processed")
        return None

    # 2. Print the overview
    answer = input(
        f"{len(processable)} orders can be processed. Would you like to continue? (y/n)\n"
    )

    if answer.lower() == 'y':
        print("Let's go!")
    else:
        print("Ok!")
        return

    # 3. Process each order

    # for each asso in the list of processable orders,
    for asso_name in set(line['Bénéficiaire'] for line in processable):
        print("\nProcessing association:", asso_name)

        # find all the lines that correspond to the association
        asso_orders = [order for order in processable if order['Bénéficiaire'].lower(
        ) == asso_name.lower()]

        # get the association details
        asso_official_name, asso_address, tresurer_first_name, tresurer_email = ru.get_asso_address(
            asso_name)

        # store the paths to send them by email
        receipts_paths = []

        for asso_order in asso_orders:
            # get the order(s) details
            total_print_price = asso_order["Prix total"] + " TTC"
            orders_list = []

            # iterate over each type of print (A1, A2, A3, sticker, t-shirt)
            for q_id, qtity_st in enumerate(list(asso_order[order_key] for order_key in ("A1", "A2", "A3", "Sticker", "T-shirt"))):
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
            print(f" - Gerenated receipt {receipt_nb}.")
            receipts_paths.append(pdf_file_name)

            # update the spreadsheet
            su.write_receipt_number(
                receipt_nb, asso_order['line'], spreadsheet, col_indexes)

        # send an email with the receipts attached
        ut.send_receipts_by_mail(
            tresurer_first_name, tresurer_email, asso_name, receipts_paths, asso_orders)
        print(f"Email sent to {tresurer_email}")


def process_all_individuals(spreadsheet, col_indexes:  Dict[str, int]):
    # retrieve all the data from the spreadsheet
    data = su.fetch_all_data(spreadsheet, col_indexes)

    # filter out the lines that already have a receipt or have been paid
    filtered_data = ut.filter_processed_orders(data)

    # 1. Filter the orders that can be processed (for which we have an email address)
    processable = ut.filter_individuals_orders(filtered_data)

    if len(processable) == 0:
        print("No association can be processed")
        return None

    # 2. Print the overview
    answer = input(
        f"{len(processable)} orders can be processed. Would you like to continue? (y/n)\n"
    )

    if answer.lower() == 'y':
        print("Let's go!")
    else:
        print("Ok!")
        return

    # 3. process all orders

    # retrieve the set of beneficiaries
    beneficiaries = set(line['Bénéficiaire'] for line in processable)

    for indiv_name in beneficiaries:
        print("\nProcessing individual:", indiv_name)

        # find all the lines that correspond to the association
        indiv_orders = [order for order in processable if order['Bénéficiaire'].lower(
        ) == indiv_name.lower()]

        # store the generated receipts paths to send them by email
        receipts_paths = []

        for order in indiv_orders:

            # get the order(s) details
            total_print_price = order["Prix total"] + " TTC"
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
            sheet_receipt_names = set(
                data[i]["№ facture"] for i in range(len(data)))
            receipt_nb = ru.get_receipt_number(sheet_receipt_names)

            recipient_info = order['Bénéficiaire']
            docx_file_name = rc.build_receipt_path(
                ru.RECEIPTS_PATH, ru.get_this_months_dir_name(), receipt_nb + ".docx")

            rc.create_receipt_docx(
                recipient_info, orders_list, receipt_nb, total_print_price, docx_file_name)
            pdf_file_name = rc.build_receipt_path(
                ru.RECEIPTS_PATH, ru.get_this_months_dir_name(), receipt_nb + ".pdf")

            # export to pdf
            rc.export_receipt_to_pdf(docx_file_name, pdf_file_name)
            print(f" - Gerenated receipt {receipt_nb}.")
            receipts_paths.append(pdf_file_name)

            # update the spreadsheet
            su.write_receipt_number(
                receipt_nb, order['line'], spreadsheet, col_indexes)

        # send an email with the receipts attached
        email_address = order['Contact eventuel']
        ut.send_receipts_by_mail(
            recipient_info, email_address, None, receipts_paths, indiv_orders, 'individual')
        print(f"Email sent to {email_address}")


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
