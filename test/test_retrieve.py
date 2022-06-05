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
    
    def test_filter_by_client_type(self):

        # create test data
        asso_orders = pd.DataFrame({'Inté / Exté': ['Asso'], 'Bénéficiaire': ['adr']}, )
        indiv_orders = pd.DataFrame({'Inté / Exté': ['Inté', 'Inté'], 'Bénéficiaire': ['bde', 'bde']})
        extern_orders = pd.DataFrame({'Inté / Exté': ['Exté', 'Exté', 'Exté'], 'Bénéficiaire': ['bds', 'bds', 'bds']})

        total_orders = pd.concat([asso_orders, indiv_orders, extern_orders])

        retrieved_asso_orders = self.retriever.filter_by_client_type(total_orders, 'Asso')
        retrieved_indiv_orders = self.retriever.filter_by_client_type(total_orders, 'Inté')
        retrieved_extern_orders = self.retriever.filter_by_client_type(total_orders, 'Exté')

        self.assertTrue(retrieved_asso_orders.equals(asso_orders))
        self.assertTrue(retrieved_indiv_orders.equals(indiv_orders))
        self.assertTrue(retrieved_extern_orders.equals(extern_orders))