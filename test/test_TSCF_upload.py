"""
Test the read_holding() method from open_holding.py

"""

import unittest2
from xlrd import open_workbook
from trustee.utility import get_current_directory
from small_program.read_file import read_file
from trustee.TSCF_upload import read_line_jones, update_position
from trustee.geneva import read_line
from os.path import join



class TestTSCFUpload(unittest2.TestCase):

    def __init__(self, *args, **kwargs):
        super(TestTSCFUpload, self).__init__(*args, **kwargs)

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



    def test_read_line_jones(self):
        input_file = join(get_current_directory(), 'samples', 'Jones Holding 2017.12.20.xlsx')
        holding, row_in_error = read_file(input_file, read_line_jones)

        self.assertEqual(len(holding), 105)
        self.verify_jones_position1(holding[0])
        self.verify_jones_position2(holding[4])
        self.verify_jones_position3(holding[104])



    def test_update_position(self):
        input_file = join(get_current_directory(), 'samples', 'Jones Holding 2017.12.20.xlsx')
        jones_holding, row_in_error = read_file(input_file, read_line_jones)
        
        input_file = join(get_current_directory(), 'samples', '12229 local appraisal 20180103.xlsx')
        geneva_holding, row_in_error = read_file(input_file, read_line)

        self.assertEqual(len(geneva_holding), 88)
        update_position(geneva_holding, jones_holding)
        self.verify_geneva_position1(geneva_holding[1])
        self.verify_geneva_position2(geneva_holding[4])
        self.verify_geneva_position3(geneva_holding[87])



    def verify_jones_position1(self, position):
        # fist position in Jones Holding 2017.12.20.xlsx
        # FR0013101599 CNP ASSURANCES (CNPFP 6 01/22/49 FIXED)
        self.assertEqual(len(position), 3)
        self.assertEqual(position['ISIN'], 'FR0013101599')
        self.assertAlmostEqual(position['Purchase Cost'], 98.233)
        self.assertAlmostEqual(position['Yield at Cost'], 6.125)



    def verify_jones_position2(self, position):
        # 5th position in Jones Holding 2017.12.20.xlsx
        self.assertEqual(len(position), 3)
        self.assertEqual(position['ISIN'], 'HK0000171949')
        self.assertAlmostEqual(position['Purchase Cost'], 100)
        self.assertAlmostEqual(position['Yield at Cost'], 6.15)




    def verify_jones_position3(self, position):
        # last position in Jones Holding 2017.12.20.xlsx
        self.assertEqual(len(position), 3)
        self.assertEqual(position['ISIN'], 'XS1736887099')
        self.assertAlmostEqual(position['Purchase Cost'], 100)
        self.assertAlmostEqual(position['Yield at Cost'], 4.8)



    def verify_geneva_position1(self, position):
        # first CNY position in 12229 local appraisal 20180103.xlsx
        # HK0000171949 HTM
        self.assertEqual(position['Group1'], 'Chinese Renminbi Yuan')
        self.assertAlmostEqual(position['Yield at Cost'], 6.15)
        self.assertAlmostEqual(position['Purchase Cost'], 100)



    def verify_geneva_position2(self, position):
        # first HKD position in 12229 local appraisal 20180103.xlsx
        # HK0000163607 HTM
        self.assertEqual(position['Group1'], 'Hong Kong Dollar')
        self.assertAlmostEqual(position['Yield at Cost'], 6.193005)
        self.assertAlmostEqual(position['Purchase Cost'], 99.259)



    def verify_geneva_position3(self, position):
        # last position in 12229 local appraisal 20180103.xlsx
        # US912803AY96 HTM
        self.assertEqual(position['Group1'], 'United States Dollar')
        self.assertAlmostEqual(position['Yield at Cost'], 6.234)
        self.assertAlmostEqual(position['Purchase Cost'], 39.9258742)

        
