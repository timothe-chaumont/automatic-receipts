""" Main code that reads from the spreadsheet and produces a receipt. """

import argparse
from tqdm import tqdm

import receipt_creation as rc
import receipt_utils as ru
import spreadsheet_utils as su

# TODO:
# - when counting the lines for an association, only count those without a receipt
# - add a function to list all the associations and their orders with no receipt (or not paid)
# - add space before euro symbol receipt
# - send email automatically


def main(args):
    """ Main function that reads from the spreadsheet and produces a receipt. """

    creds = su.connect_to_spreadsheet()
    spreadsheet = su.get_spreadsheet(creds)
    # retrieve column indexes
    col_indexes = su.get_all_col_indexes(spreadsheet)

    if args.summary:
        print("----- Summary of the spreadsheet -----\n")
        nor_rec_lines = su.get_no_receipt_lines(spreadsheet, col_indexes)
        nb_lines_by_recipient = {}
        for line in tqdm(nor_rec_lines):
            orders_list, total_print_price, recipient_name = su.get_order_data(
                line, spreadsheet, col_indexes)
            if recipient_name not in nb_lines_by_recipient:
                nb_lines_by_recipient[recipient_name] = 1
            else:
                nb_lines_by_recipient[recipient_name] += 1
            # print(f"Line {line}: {recipient_name}, {total_print_price}")
        for recipient in nb_lines_by_recipient:
            print(f"{recipient} ({nb_lines_by_recipient[recipient]} orders)")

    # if an association was given
    elif args.association:
        asso_name = " ".join(args.association)
        # find lines corresponding to Hyris's orders
        entry_lines = su.find_lines(
            asso_name, spreadsheet, col_indexes["Bénéficiaire"])
        print(f"{len(entry_lines)} ligne(s) trouvée(s) pour {asso_name}.")

        for line in entry_lines:
            # check if the receipt has already been created
            if not su.has_receipt(line, spreadsheet, col_indexes):

                orders_list, total_print_price, recipient_name = su.get_order_data(
                    line, spreadsheet, col_indexes)
                asso_name, asso_address = ru.get_asso_address(recipient_name)
                receipt_nb = ru.get_receipt_number()

                recipient_info = asso_name + "\n" + asso_address
                docx_file_name = rc.build_receipt_path(
                    ru.RECEIPTS_PATH, ru.get_this_months_dir_name(), receipt_nb + ".docx")

                rc.create_receipt_docx(
                    recipient_info, orders_list, receipt_nb, total_print_price, docx_file_name)
                pdf_file_name = rc.build_receipt_path(
                    ru.RECEIPTS_PATH, ru.get_this_months_dir_name(), receipt_nb + ".pdf")
                # export to pdf
                rc.export_receipt_to_pdf(docx_file_name, pdf_file_name)
                print(f"Facture {receipt_nb} générée.")

                # update the spreadsheet
                su.write_receipt_number(
                    receipt_nb, line, spreadsheet, col_indexes)


if __name__ == '__main__':
    # use a parser to get the arguments
    parser = argparse.ArgumentParser()
    parser.add_argument("-a", "--association", help="Process all entries for the given association.",
                        type=str, nargs="+")  # at least one word
    parser.add_argument("-s", "--summary", help="Prints a description of the current state of the spreadsheet.",
                        action='store_true')  # no arguments needed
    args = parser.parse_args()

    main(args)