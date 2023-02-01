
'''
PSEUDOCODE
    * Create a copy of the p9 template, give it a person's name
    * Open the file
    * Save the data in the file
'''

import unittest
import openpyxl
import main

class TestGetPayeWorksheet(unittest.TestCase):

    def setUp(self):
        self.workbook = openpyxl.Workbook()
        self.worksheet = self.workbook.create_sheet('PAYE2022')

    def test_worksheet_isInstance(self):
        sheet = main.get_paye_worksheet(self.workbook)
        self.assertIsInstance(sheet, 
            openpyxl.worksheet.worksheet.Worksheet)

if __name__ == '__main__':
    unittest.main()
