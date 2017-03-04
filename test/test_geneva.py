"""
Test the read_holding() method from open_holding.py

"""

import unittest2
from datetime import datetime
from trustee.utility import get_current_directory
from trustee.geneva import read_line, filter_maturity
from small_program.read_file import read_file
from os.path import join



class TestGeneva(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestGeneva, self).__init__(*args, **kwargs)

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



    def test_read_file(self):
        file = join(get_current_directory(), 'samples', '12229_local_appraisal_sample1.xlsx')
        holding, row_in_error = read_file(file, read_line)
        self.assertEqual(len(row_in_error), 0)
        self.assertEqual(len(holding), 58)

        self.verify_position1(holding[0])
        self.verify_position2(holding[2])
        self.verify_position3(holding[57])
        self.assertEqual(len(filter_maturity(holding)), 27)



    def test_read_file2(self):
        file = join(get_current_directory(), 'samples', '12229_local_appraisal_sample2.xlsx')
        holding, row_in_error = read_file(file, read_line)
        self.assertEqual(len(row_in_error), 0)
        self.assertEqual(len(holding), 58)
        self.assertEqual(len(filter_maturity(holding)), 28)
        self.verify_position4(filter_maturity(holding)[1])



    def test_read_file3(self):
        file = join(get_current_directory(), 'samples', '12229_local_appraisal_sample3.xlsx')
        holding, row_in_error = read_file(file, read_line)
        self.assertEqual(len(row_in_error), 0)
        self.assertEqual(len(holding), 6)
        self.assertEqual(len(filter_maturity(holding)), 2)
        self.verify_position5(filter_maturity(holding))



    def test_read_file4(self):
        file = join(get_current_directory(), 'samples', '12732_local_appraisal_sample4.xlsx')
        holding, row_in_error = read_file(file, read_line)
        self.assertEqual(len(row_in_error), 0)
        self.assertEqual(len(holding), 3)
        self.assertEqual(len(filter_maturity(holding)), 1)



    def verify_position1(self, position):
        self.assertEqual(position['Portfolio'], '12229')
        self.assertEqual(position['InvestID'], 'CNY')
        self.assertAlmostEqual(position['Quantity'], -67698863.01)
        self.assertEqual(position['Group1'], 'Cash and Equivalents')
        self.assertFalse(position['MaturityDate'], '')



    def verify_position2(self, position):
        self.assertEqual(position['Portfolio'], '12229')
        self.assertEqual(position['InvestID'], 'HK0000083706 HTM')
        self.assertEqual(position['Group1'], 'Corporate Bond')
        self.assertAlmostEqual(position['Quantity'], 20000000)
        self.assertEqual(position['MaturityDate'], datetime(2016,6,30))
        self.assertAlmostEqual(position['UnitCost'], 100)



    def verify_position3(self, position):
        self.assertEqual(position['Portfolio'], '12229')
        self.assertEqual(position['InvestID'], 'US912803AY96 HTM')
        self.assertEqual(position['Group2'], 'Government Bond')
        self.assertAlmostEqual(position['PercentAssets'], 0.02)
        self.assertEqual(position['MaturityDate'], datetime(2021,11,15))
        self.assertAlmostEqual(position['UnitCost'], 29.117)



    def verify_position4(self, position):
        self.assertEqual(position['Portfolio'], '12229')
        self.assertEqual(position['InvestID'], 'US02580ECN13 HTM')
        self.assertEqual(position['Group2'], 'Corporate Bond')
        self.assertAlmostEqual(position['PercentAssets'], 11.03)
        self.assertEqual(position['MaturityDate'], '')
        self.assertAlmostEqual(position['UnitCost'], 100.0657)



    def verify_position5(self, holding):
        self.assertEqual(holding[0]['InvestID'], 'US02580ECC57 HTM')
        self.assertEqual(holding[1]['InvestID'], 'US459200AG65 HTM')