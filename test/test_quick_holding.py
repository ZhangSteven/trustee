"""
Test the read_holding() method from open_holding.py

"""

import unittest2
from xlrd import open_workbook
from trustee.utility import get_current_directory
from small_program.read_file import read_file
from trustee.quick_holding import read_line_trustee, update_amortized_cost
from trustee.geneva import read_line
from os.path import join



class TestQuickHolding(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestQuickHolding, self).__init__(*args, **kwargs)

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



    def test_read_line_trustee(self):
        input_file = join(get_current_directory(), 'samples', 'new_12229.xlsx')
        holding, row_in_error = read_file(input_file, read_line_trustee, starting_row=2)

        self.assertEqual(len(holding), 85)
        self.verify_trustee_position1(holding[0])
        self.verify_trustee_position2(holding[3])
        self.verify_trustee_position3(holding[84])



    def test_update_amortized_cost(self):
        input_file = join(get_current_directory(), 'samples', 'new_12229.xlsx')
        trustee_holding, row_in_error = read_file(input_file, read_line_trustee, starting_row=2)

        input_file = join(get_current_directory(), 'samples', '12229_local_appraisal_sample5.xlsx')
        geneva_holding, row_in_error = read_file(input_file, read_line)

        self.assertEqual(len(geneva_holding), 88)
        update_amortized_cost(geneva_holding, trustee_holding)
        self.verify_geneva_position1(geneva_holding[1])
        self.verify_geneva_position2(geneva_holding[6])
        self.verify_geneva_position3(geneva_holding[87])



    def verify_trustee_position1(self, position):
        # fist position in new_12229.xlsx
        # HK0000134780 FarEast Horizon5.75%
        self.assertEqual(len(position), 2)
        self.assertEqual(position['Identifier'], 'HK0000134780 HTM')



    def verify_trustee_position2(self, position):
        # 4th position in new_12229.xlsx
        # DBANFB12014 Dragon Days Ltd 6.0%
        self.assertEqual(len(position), 2)
        self.assertEqual(position['Identifier'], 'HK0000175916 HTM')



    def verify_trustee_position3(self, position):
        # last position in new_12229.xlsx
        # XS1600847666 LEGAL & GENERAL GROU
        self.assertEqual(len(position), 2)
        self.assertEqual(position['Identifier'], 'XS1600847666 HTM')



    def verify_geneva_position1(self, position):
        # first non-cash position in 12229_local_appraisal_sample5.xlsx
        # HK0000134780 HTM
        self.assertEqual(position['Group1'], 'Chinese Renminbi Yuan')
        self.assertEqual(position['Quantity'], 100000000)
        self.assertAlmostEqual(position['Amortized Cost'], 99.88211324)



    def verify_geneva_position2(self, position):
        # 7th position in 12229_local_appraisal_sample5.xlsx
        # HK0000175916 HTM (Dragon Days in trustee)
        self.assertEqual(position['Group1'], 'Hong Kong Dollar')
        self.assertEqual(position['Quantity'], 84000000)
        self.assertAlmostEqual(position['Amortized Cost'], 100)



    def verify_geneva_position3(self, position):
        # last position in 12229_local_appraisal_sample5.xlsx
        # US912803AY96 HTM
        self.assertEqual(position['Group1'], 'United States Dollar')
        self.assertEqual(position['Quantity'], 250000)
        self.assertAlmostEqual(position['Amortized Cost'], 79.812716)

        
