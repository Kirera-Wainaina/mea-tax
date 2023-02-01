import openpyxl

def open_workbook ():
    workbook = openpyxl.load_workbook(filename='./static/records.xlsx')
    return workbook

def get_paye_worksheet (workbook):
    return workbook['PAYE2022']

if __name__ == '__main__':
    workbook = open_workbook();
