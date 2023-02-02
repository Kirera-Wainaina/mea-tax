import openpyxl
import os
import shutil

def open_workbook():
    workbook = openpyxl.load_workbook(filename='./static/records.xlsx')
    return workbook

def get_paye_worksheet(workbook):
    return workbook['PAYE2022']

def iterate_through_rows(worksheet, 
        min_row, max_row,
        min_col, max_col):
    for row in worksheet.iter_rows(
            min_row=min_row, max_row=max_row, 
            min_col=min_col, max_col=max_col):
        handle_employee_details(row)

def handle_employee_details(details):
    employee_name = details[1].value
    tax_file = create_employee_tax_file(employee_name)
    return True

def create_employee_tax_file(name):
    template_path = '{cwd}/static/p9.xlsx'.format(cwd=os.getcwd())
    employee_file_path = '{cwd}/repo/{name}.xlsx'.format(cwd=os.getcwd(), name=name)
    shutil.copy(template_path, employee_file_path)
    return employee_file_path

if __name__ == '__main__':
    workbook = open_workbook()
    worksheet = get_paye_worksheet(workbook)
    iterate_through_rows(
        worksheet=worksheet, min_row=5, max_row=468,
        min_col=1, max_col=28)
