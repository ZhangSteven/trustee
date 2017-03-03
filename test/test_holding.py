"""
Test the read_holding() method from open_holding.py

"""

import unittest2
from datetime import datetime
from xlrd import open_workbook
from trustee.utility import get_current_directory
from trustee.holding import read_sub_section, read_section, read_holding
from os.path import join



class TestHolding(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestHolding, self).__init__(*args, **kwargs)

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



    def test_read_sub_section(self):
        filename = join(get_current_directory(), 'samples', 'nav_sample1.xls')
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 70    # the bond sub section starts at A71
        accounting_treatment = 'HTM'
        fields = ['par_amount', 'is_listed', 'listed_location', 
                    'fx_on_trade_day', 'coupon_rate', 'coupon_start_date', 
                    'maturity_date', 'average_cost', 'amortized_cost', 
                    'book_cost', 'interest_bought', 'amortized_value', 
                    'accrued_interest', 'amortized_gain_loss', 'fx_gain_loss',
                    'fund_percentage']
        asset_class = 'bond'
        currency = 'USD'
        bond_holding = []

        row = read_sub_section(ws, row, accounting_treatment, fields, asset_class, currency, bond_holding)
        self.assertEqual(row, 102)
        self.assertEqual(len(bond_holding), 17)
        self.verify_bond_position1(bond_holding[0])
        self.verify_bond_position2(bond_holding[16])



    def test_read_section(self):
        filename = join(get_current_directory(), 'samples', 'nav_sample1.xls')
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        row = 61    # the bond section starts at A62
        port_values = {}
        currency = 'USD'
        fields = ['par_amount', 'is_listed', 'listed_location', 
                    'fx_on_trade_day', 'coupon_rate', 'coupon_start_date', 
                    'maturity_date', 'average_cost', 'amortized_cost', 
                    'book_cost', 'interest_bought', 'amortized_value', 
                    'accrued_interest', 'amortized_gain_loss', 'fx_gain_loss',
                    'fund_percentage']
        row = read_section(ws, row, fields, 'bond', currency, port_values)
        self.assertEqual(row, 133)  # reading stops at A134
        bond_holding = port_values['bond']
        self.assertEqual(len(bond_holding), 25)
        self.verify_bond_position1(bond_holding[0])
        self.verify_bond_position2(bond_holding[16])
        self.verify_bond_position3(bond_holding[24])



    def test_read_holding(self):
        filename = join(get_current_directory(), 'samples', 'nav_sample1.xls')
        wb = open_workbook(filename=filename)
        ws = wb.sheet_by_name('Portfolio Val.')
        port_values = {}
        read_holding(ws, port_values)
        bond_holding = port_values['bond']
        self.assertEqual(len(bond_holding), 37)
        self.verify_bond_position1(bond_holding[0])
        self.verify_bond_position2(bond_holding[16])
        self.verify_bond_position3(bond_holding[24])
        self.verify_bond_position4(bond_holding[36])



    def verify_bond_position1(self, bond):
        """
        Bond position at A73, nav_sample1.xls
        """
        self.assertEqual(bond['isin'], 'US78387GAP81')
        self.assertEqual(bond['name'], 'AT&T INC 5.1%')
        self.assertEqual(bond['currency'], 'USD')
        self.assertEqual(bond['accounting_treatment'], 'HTM')
        self.assertAlmostEqual(bond['par_amount'], 1000000)
        self.assertEqual(bond['is_listed'], 'Y')
        self.assertEqual(bond['listed_location'], 'TBC')
        self.assertAlmostEqual(bond['fx_on_trade_day'], 7.7694)
        self.assertAlmostEqual(bond['coupon_rate'], 5.1/100)
        self.assertEqual(bond['coupon_start_date'], datetime(2011,9,15))
        self.assertEqual(bond['maturity_date'], datetime(2014,9,15))
        self.assertAlmostEqual(bond['average_cost'], 97.733735)
        self.assertAlmostEqual(bond['amortized_cost'], 99.16551)   
        self.assertAlmostEqual(bond['book_cost'], 977337.35)
        self.assertAlmostEqual(bond['interest_bought'], 0)
        self.assertAlmostEqual(bond['amortized_value'], 991655.1)
        self.assertAlmostEqual(bond['accrued_interest'], 15016.67)
        self.assertAlmostEqual(bond['amortized_gain_loss'], 14317.75)
        self.assertAlmostEqual(bond['fx_gain_loss'], -0.00999999977648258)



    def verify_bond_position2(self, bond):
        """
        Bond position at A97, nav_sample1.xls
        """
        self.assertEqual(bond['isin'], 'USG46715AC56')
        self.assertEqual(bond['name'], 'HUTCHISON WHAMPOA 7.5%')
        self.assertEqual(bond['currency'], 'USD')
        self.assertEqual(bond['accounting_treatment'], 'HTM')
        self.assertAlmostEqual(bond['par_amount'], 500000)
        self.assertEqual(bond['is_listed'], 'Y')
        self.assertEqual(bond['listed_location'], 'TBC')
        self.assertAlmostEqual(bond['fx_on_trade_day'], 7.7694)
        self.assertAlmostEqual(bond['coupon_rate'], 7.5/100)
        self.assertEqual(bond['coupon_start_date'], datetime(2011,8,1))
        self.assertEqual(bond['maturity_date'], datetime(2027,8,1))
        self.assertAlmostEqual(bond['average_cost'], 111.30955)
        self.assertAlmostEqual(bond['amortized_cost'], 110.122788)   
        self.assertAlmostEqual(bond['book_cost'], 556547.75)
        self.assertAlmostEqual(bond['interest_bought'], 0)
        self.assertAlmostEqual(bond['amortized_value'], 550613.94)
        self.assertAlmostEqual(bond['accrued_interest'], 15625)
        self.assertAlmostEqual(bond['amortized_gain_loss'], -5933.81000000005)
        self.assertAlmostEqual(bond['fx_gain_loss'], 0)



    def verify_bond_position3(self, bond):
        """
        Bond position at A130, nav_sample1.xls
        """
        self.assertEqual(bond['isin'], 'XS0545110354')
        self.assertEqual(bond['name'], 'Franshion 6.8% Perpetual Subordinated Convertible Securities Callable 2015')
        self.assertEqual(bond['currency'], 'USD')
        self.assertEqual(bond['accounting_treatment'], 'AFS')
        self.assertAlmostEqual(bond['par_amount'], 10200000)
        self.assertEqual(bond['is_listed'], 'TBC')
        self.assertEqual(bond['listed_location'], 'TBC')
        self.assertAlmostEqual(bond['fx_on_trade_day'], 7.7694)
        self.assertAlmostEqual(bond['coupon_rate'], 6.8/100)
        self.assertEqual(bond['coupon_start_date'], datetime(2011,10,12))
        self.assertEqual(bond['maturity_date'], 'N/A')
        self.assertAlmostEqual(bond['average_cost'], 99.9950980392157)
        self.assertAlmostEqual(bond['amortized_cost'], 81.75)   
        self.assertAlmostEqual(bond['book_cost'], 10199500)
        self.assertAlmostEqual(bond['interest_bought'], 0)
        self.assertAlmostEqual(bond['amortized_value'], 8338500)
        self.assertAlmostEqual(bond['accrued_interest'], 152206.67)
        self.assertAlmostEqual(bond['amortized_gain_loss'], -1861000)
        self.assertAlmostEqual(bond['fx_gain_loss'], 0)



    def verify_bond_position4(self, bond):
        """
        Bond position at A186, nav_sample1.xls
        """
        self.assertEqual(bond['isin'], 'HK0000096856')
        self.assertEqual(bond['name'], 'Far East Horizon Ltd 6.95%')
        self.assertEqual(bond['currency'], 'CNY')
        self.assertEqual(bond['accounting_treatment'], 'HTM')
        self.assertAlmostEqual(bond['par_amount'], 69700000)
        self.assertEqual(bond['is_listed'], 'TBC')
        self.assertEqual(bond['listed_location'], 'TBC')
        self.assertAlmostEqual(bond['fx_on_trade_day'], 1.2309)
        self.assertAlmostEqual(bond['coupon_rate'], 6.95/100)
        self.assertEqual(bond['coupon_start_date'], datetime(2011,12,21))
        self.assertEqual(bond['maturity_date'], datetime(2016,12,21))
        self.assertAlmostEqual(bond['average_cost'], 99.6429072453372)
        self.assertAlmostEqual(bond['amortized_cost'], 99.6429072453372)   
        self.assertAlmostEqual(bond['book_cost'], 69410048)
        self.assertAlmostEqual(bond['interest_bought'], 0)
        self.assertAlmostEqual(bond['amortized_value'], 69451106.35)
        self.assertAlmostEqual(bond['accrued_interest'], 145988.08)
        self.assertAlmostEqual(bond['amortized_gain_loss'], 41058.349999994)
        self.assertAlmostEqual(bond['fx_gain_loss'], 0.01)

