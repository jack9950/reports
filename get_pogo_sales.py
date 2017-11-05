import sys
import xlrd
import time
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException
from data_files import homeFolder, callsHandledReportLocation, pogoSalesReportLocation

def get_pogo_sales(filename):
# first open using xlrd    book = xlrd.open_workbook(filename)
    # currentHour = time.strftime('%H')
    # filename = homeFolder + 'bounce_energy_iqor_report_' + currentHour  + '.xls'
    try:
        book = xlrd.open_workbook(filename)
    except FileNotFoundError:
        print("File: ", filename)
        print("\nFile not found...Exiting...")
        raise

    sheet = book.sheet_by_index(0)
    nrows, ncols = sheet.nrows, sheet.ncols

    values = []

    for row in range(1, nrows):
        if sheet.cell_value(row,6) != '':
            #The format is [agent ID]
            values.append(sheet.cell_value(row,6))

    return values
