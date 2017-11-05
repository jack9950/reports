import sys
import xlrd
import time
from teams import agent_ids_to_names
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException
from data_files import homeFolder, callsHandledReportLocation, pogoSalesReportLocation

#Sample return:
# [agent_id, [Acct #, Order #, order status], [Acct #, Order #, order status]]
# [2062062, [2092985, 1443822, "Deposit due"], [2092021, 1444496, "Ercot/ISO Processing"] ]

def get_pogo_sales_breakdown(filename):
# first open using xlrd    book = xlrd.open_workbook(filename)
    # currentHour = time.strftime('%H')
    # filename = homeFolder + 'bounce_energy_iqor_report_' + currentHour  + '.xls'

    try:
        book = xlrd.open_workbook(filename)
    except FileNotFoundError:
        print("File: ", filename)
        print("\nFile not found...Exiting...")
        raise
        #raise

    sheet = book.sheet_by_index(0)
    nrows, ncols = sheet.nrows, sheet.ncols

    values = []

    for row in range(1, nrows):
        if sheet.cell_value(row,6) != '':

            #print(agent_id, ": ", agent_ids_to_names[agent_id])
            try:
                agent_id = sheet.cell_value(row,6)
                values.append([agent_ids_to_names[agent_id],
                               sheet.cell_value(row,1),
                               sheet.cell_value(row,0),
                               sheet.cell_value(row,3)])
            except:
                pass

    return values
