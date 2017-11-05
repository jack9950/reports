import sys
import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException
from data_files import agent_ids_to_names

#Sample return:
# [agent_id, [Acct #, Order #, order status], [Acct #, Order #, order status]]
# [2062062, [2092985, 1443822, "Deposit due"], [2092021, 1444496, "Ercot/ISO Processing"] ]

def get_DEPP_sales_breakdown(filename):
# first open using xlrd    book = xlrd.open_workbook(filename)
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
        agent_id = sheet.cell_value(row, 16)
        DEPP_name = sheet.cell_value(row, 5)
        if(agent_id != None and (DEPP_name == "Surge Protection Plan" or
                                 DEPP_name == "Electric Repair Essentials" or
                                 DEPP_name == "Surge Protection Plan (20% Off)" or
                                 DEPP_name == "Cooling Maintenance Essentials (6 Month Free Trial - Nest Bundle)" or
                                 DEPP_name == "Cooling Repair & Maintenance Essentials" or
                                 DEPP_name == "Electric Repair Essentials (20% Off)") or
                                 DEPP_name == "Heating & Cooling Repair Essentials"):
            try:
                agent_name = agent_ids_to_names[agent_id]
                pogo_account_number = sheet.cell_value(row,0)
                pogo_order_number = sheet.cell_value(row,1)
                DEPP_name = sheet.cell_value(row,5)
                bounce_status = sheet.cell_value(row,10)
                print("Inside get_DEPP_sales_breakdown:")
                print(agent_name, pogo_account_number, pogo_order_number,
                               DEPP_name, bounce_status)
                values.append([agent_name,
                               pogo_account_number,
                               pogo_order_number,
                               DEPP_name,
                               bounce_status])
            except:
                pass

    return values
