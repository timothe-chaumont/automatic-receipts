""" Reads, filters data from the spreadsheet and produces receipts for the unprocessed orders. """

import warnings
warnings.simplefilter(action='ignore')

import pandas as pd

import receipt_creation as rc
import receipt_utils as ru
import spreadsheet_utils as su
import utils as ut
from retrieve import Retriever



def main():
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
    print(f"\n{len(can_be_processed)} prestations peuvent être traitées, dont:")
    print(f" - {len(can_be_processed_asso)} pour des associations,")
    print(f" - {len(can_be_processed_indiv)} pour des étudiants,")
    print(f" - {len(can_be_processed_etern)} pour des clients extérieurs.")

    answer = input("\nVoulez-vous continuer? (O/N)\n")
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
            print(f" - Facture {receipt_nb} exportée.")
            receipts_paths.append(pdf_file_name)

            # update the spreadsheet
            retriever.write_receipt_number(receipt_nb, order_idx)

        # send an email with the receipts attached
        if order["Inté / Exté"] == "Asso":
            recip_mail = asso_details["tresurer mail"]
            recipient_first_name = asso_details["tresurer first name"]

        else:
            recip_mail = order["Contact eventuel"]
            recipient_first_name = None

        ut.send_receipts_by_mail(
            recip_name, 
            recip_mail, 
            recip_type, 
            receipts_paths, 
            recip_orders.to_dict(orient="records"), 
            recipient_first_name
        )
        print(f"Email envoyé à {recip_name}.")

if __name__ == '__main__':
    main()
