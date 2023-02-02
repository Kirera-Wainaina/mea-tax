
'''
PSEUDOCODE
    * Create a copy of the p9 template, give it a person's name
    * Open the file
    * Save the data in the file
'''

import unittest
from unittest.mock import MagicMock, patch
import openpyxl
import main # main.py
import os

class TestGetPayeWorksheet(unittest.TestCase):

    def setUp(self):
        self.workbook = openpyxl.Workbook()
        self.worksheet = self.workbook.create_sheet('PAYE2022')
        self.sheet = main.get_paye_worksheet(self.workbook)

    def test_worksheet_isInstance(self):
        self.assertIsInstance(self.sheet, 
            openpyxl.worksheet.worksheet.Worksheet)

    def test_worksheet_name(self):
        self.assertEqual(self.sheet.title, 'PAYE2022')

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

class TestHandleEmployeeDetails(unittest.TestCase):
    
    def setUp(self):
        self.details = MagicMock()

    def test_returnValue_true(self):
        self.assertTrue(main.handle_employee_details(self.details))

class TestCreateEmployeeTaxFile(unittest.TestCase):

    def setUp(self):
        pass

    def test_path_isReturned(self):
        path = main.create_employee_tax_file('richard')
        self.assertEqual(path, '{cwd}/repo/{name}.xlsx'.format(
            cwd=os.getcwd(), name='richard'))


if __name__ == '__main__':
    unittest.main()
