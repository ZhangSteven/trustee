"""
Test the read_holding() method from open_holding.py

"""

import unittest2
from datetime import datetime
from xlrd import open_workbook
from trustee.utility import get_current_directory
from trustee.transaction import read_bond_section, get_report_name, \
                                    get_portfolio_id, read_transaction
from os.path import join



class TestTransaction(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestTransaction, self).__init__(*args, **kwargs)

    def setUp(self):
        """
            Run before a test function
        """
        pass

    def tearDown(self):
        """
            Run after a test finishes
        """
        pass



    def test_get_report_name(self):
        filename = join(get_current_directory(), 'samples', 'new_nav_sample2.xls')
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_index(0)
        self.assertEqual(get_report_name(ws), 'SECURITIES TRANSACTION IMPLEMENTED')
        


    def test_get_portfolio_id(self):
        filename = join(get_current_directory(), 'samples', 'new_nav_sample2.xls')
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_index(0)
        self.assertEqual(get_portfolio_id(ws), ('12229', 6))
        


    def test_read_bond_section(self):
        filename = join(get_current_directory(), 'samples', 'new_nav_sample2.xls')
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_index(0)
        port_values = {}
        port_values['portfolio_id'] = '12229'
        row = read_bond_section(ws, 26, 'buy', port_values)
        self.assertEqual(row, 41)   # row# of the equities section begins
        transactions = port_values['bond_transactions']
        self.assertEqual(len(transactions), 2)
        self.verify_transaction1(transactions[0])
        self.verify_transaction2(transactions[1])



    def test_read_transaction(self):
        filename = join(get_current_directory(), 'samples', 'new_nav_sample2.xls')
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_index(0)
        port_values = {}
        read_transaction(ws, port_values)
        transactions = port_values['bond_transactions']
        self.assertEqual(len(transactions), 3)
        self.verify_transaction1(transactions[0])
        self.verify_transaction2(transactions[1])
        self.verify_transaction3(transactions[2])



    def verify_transaction1(self, t):
        """
        First bond trade in new_nav_sample2.xls
        """
        self.assertEqual(t['portfolio_id'], '12229')
        self.assertEqual(t['action'], 'buy')
        self.assertEqual(t['trade date'], datetime(2015,1,28))
        self.assertEqual(t['value date'], datetime(2015,2,4))
        self.assertEqual(t['security_id'], 'XS1163722587')
        self.assertEqual(t['description'], 'SINO OCEAN LD TRS FIN II 5.95%')
        self.assertEqual(t['reference code'], '17482')
        self.assertEqual(t['currency'], 'USD')
        self.assertEqual(t['amount'], 40000000)
        self.assertAlmostEqual(t['price'], 98.737)
        self.assertAlmostEqual(t['fx rate'], 7.75190)
        self.assertAlmostEqual(t['effective yield'], 6.099958)



    def verify_transaction2(self, t):
        """
        First bond trade in new_nav_sample2.xls
        """
        self.assertEqual(t['portfolio_id'], '12229')
        self.assertEqual(t['action'], 'buy')
        self.assertEqual(t['trade date'], datetime(2015,2,5))
        self.assertEqual(t['value date'], datetime(2015,2,12))
        self.assertEqual(t['security_id'], 'XS1189103382')
        self.assertEqual(t['description'], 'HONG KONG INT\'L QINGDAO 5.95%')
        self.assertEqual(t['reference code'], '17606')
        self.assertEqual(t['broker'], 'BNPPAB')
        self.assertEqual(t['currency'], 'USD')
        self.assertEqual(t['amount'], 30000000)
        self.assertAlmostEqual(t['price'], 98.523)
        self.assertAlmostEqual(t['cost'], 29556900)
        self.assertAlmostEqual(t['cost HKD equivalent'], 229136911.56)



    def verify_transaction3(self, t):
        """
        First bond trade in new_nav_sample2.xls
        """
        self.assertEqual(t['portfolio_id'], '12229')
        self.assertEqual(t['action'], 'sell')
        self.assertEqual(t['trade date'], datetime(2016,2,28))
        self.assertEqual(t['value date'], datetime(2016,3,4))
        self.assertEqual(t['security_id'], 'XS123')
        self.assertEqual(t['description'], 'Some Test Sec 88%')
        self.assertEqual(t['reference code'], '88888')
        self.assertEqual(t['currency'], 'USD')
        self.assertEqual(t['amount'], 40000000)
        self.assertAlmostEqual(t['price'], 98.737)
        self.assertAlmostEqual(t['fx rate'], 7.75190)
        self.assertAlmostEqual(t['effective yield'], 6.099958)