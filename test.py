
'''
PSEUDOCODE
    * Create a copy of the p9 template, give it a person's name
    * Open the file
    * Save the data in the file
'''

import unittest
from unittest.mock import MagicMock
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

class WorkSheet:

    def iter_rows():
        pass

class TestIterateThroughRows(unittest.TestCase):

    def setUp(self):
        self.worksheet = WorkSheet()
        self.worksheet.iter_rows = MagicMock()

    def test_isCalled_withargs(self):
        min_row, max_row, min_col, max_col =5, 468, 1, 28
        main.iterate_through_rows(self.worksheet, min_row=min_row, 
            max_row=max_row, min_col=min_col, max_col=max_col)
        self.worksheet.iter_rows.assert_called_with(min_row=min_row, 
            max_row=max_row, min_col=min_col, max_col=max_col)

if __name__ == '__main__':
    unittest.main()
