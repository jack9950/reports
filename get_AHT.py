import sys
import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException

def get_AHT(filename):
# first open using xlrd    book = xlrd.open_workbook(filename)
    try:
        book = xlrd.open_workbook(filename)
    except FileNotFoundError:
        print("File: ", filename)
        print("\nFile not found...Exiting...")
        raise
        #raise

    sheet = book.sheet_by_index(1)
    nrows, ncols = sheet.nrows, sheet.ncols

    values = []

    for row in range(6, nrows):
        if sheet.cell_value(row,4) != '':
            #The format is [agent ID, Agent Name, 
            #               Sign In Time, Calls Handled, 
            #               AHT]
            values.append([int(sheet.cell_value(row,4)), sheet.cell_value(row,5), 
                           sheet.cell_value(row,8),int(sheet.cell_value(row,11)),
                           int(sheet.cell_value(row,18))])

    return values
