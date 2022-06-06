import unittest
from retrieve import Retriever
import pandas as pd


class TestRetrieverMethods(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        """
        This method is called once before running all the test methods in this class.
        """
        cls.retriever = Retriever()
    
    def test_get_unprocessed_orders(self):

        not_order_lines = pd.DataFrame({'Type': ['Commande', 'Note de frais'], '№ facture': ['', ''], 'Encaissement': ['', '']})
        paid_orders = pd.DataFrame({'Type': ['Prestation', 'Prestation', 'Prestation'], '№ facture': ['', '', ''], 'Encaissement': ['Lydia Pro', 'Virement', 'Chèque']})
        orders_with_receipt = pd.DataFrame({'Type': ['Prestation',], '№ facture': ['2022-01-0001', ], 'Encaissement': ['', ]})
        unprocessed_orders = pd.DataFrame({'Type': ['Prestation', 'Prestation', 'Prestation', 'Prestation'], '№ facture': ['', '', '', None], 'Encaissement': ['', 'Non Payé', 'En attente', None]})

        total_lines = pd.concat([not_order_lines, paid_orders, orders_with_receipt, unprocessed_orders])
        retrieved_unprocessed_orders = self.retriever.get_unprocessed_orders(total_lines)

        # create error message
        excess_lines = retrieved_unprocessed_orders[~retrieved_unprocessed_orders.index.isin(unprocessed_orders.index)]
        missing_lines = unprocessed_orders[~unprocessed_orders.index.isin(retrieved_unprocessed_orders.index)]
        msg = f"{len(missing_lines)} missing line(s) and {len(excess_lines)} excess line(s).\n\nGot: {retrieved_unprocessed_orders}\n\nExpected: {unprocessed_orders}"

        self.assertTrue(retrieved_unprocessed_orders.equals(unprocessed_orders), msg=msg)

    def test_filter_by_client_type(self):

        # create test data
        asso_orders = pd.DataFrame({'Inté / Exté': ['Asso'], 'Bénéficiaire': ['adr']})
        indiv_orders = pd.DataFrame({'Inté / Exté': ['Inté', 'Inté'], 'Bénéficiaire': ['bde', 'bde']})
        extern_orders = pd.DataFrame({'Inté / Exté': ['Exté', 'Exté', 'Exté'], 'Bénéficiaire': ['bds', 'bds', 'bds']})

        total_orders = pd.concat([asso_orders, indiv_orders, extern_orders])

        retrieved_asso_orders = self.retriever.filter_by_client_type(total_orders, 'Asso')
        retrieved_indiv_orders = self.retriever.filter_by_client_type(total_orders, 'Inté')
        retrieved_extern_orders = self.retriever.filter_by_client_type(total_orders, 'Exté')

        self.assertTrue(retrieved_asso_orders.equals(asso_orders))
        self.assertTrue(retrieved_indiv_orders.equals(indiv_orders))
        self.assertTrue(retrieved_extern_orders.equals(extern_orders))
    
if __name__ == '__main__':
    unittest.main()