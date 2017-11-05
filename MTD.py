import sys
import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook, InvalidFileException

homeFolder = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\'
callsHandledReportLocation = homeFolder + 'MTD\\Bounce_Engery_Agent_Performance_Rollup.xls'
pogoSalesReportLocation = homeFolder + 'MTD\\NOPR.xls'
fcpReportLocation = homeFolder + 'MTD\\FCP.xls'
DEPPreportLocation = homeFolder + 'MTD\\products_sonar.xls'
hiveNewServiceReportLocation = homeFolder + 'MTD\\products_sonar.xls'
hiveRenewalsReportLocation = homeFolder + 'MTD\\hive_renewals.xls'

def get_calls_handled(filename):
# first open using xlrd    book = xlrd.open_workbook(filename)
    try:
        book = xlrd.open_workbook(filename)
    except FileNotFoundError:
        print("File: ", filename)
        print("\nFile not found...Exiting...")
        raise

    sheet = book.sheet_by_name('MTD ')
    nrows, ncols = sheet.nrows, sheet.ncols

    values = []

    for row in range(6, nrows):
        if sheet.cell_value(row,4) != '':
            #The format is [agent ID, Calls Handled, Sales Calls Handled]
            values.append([sheet.cell_value(row,4), sheet.cell_value(row,11), sheet.cell_value(row,47)])
            # print("loop: ", sheet.cell_value(row,4), sheet.cell_value(row,11), sheet.cell_value(row,47))

    return values

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
    mylist = []
    for row in range(1, nrows):
        agent_id = sheet.cell_value(row,47)
        transfer_order = sheet.cell_value(row,13)
        order_status = sheet.cell_value(row,9)
        if (agent_id != ''
            and transfer_order == 'N'
            and order_status != 'Test order'
            and order_status != 'Duplicate Order'):
            #The format is [agent ID]
            values.append(agent_id)
            # print(customer_id, order_id, order_status, transfer_order, agent_id)

    return values

def get_DEPP_sales(filename):
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

    #Collect up the warranty sales add them to the value arra and return the array.
    for row in range(1, nrows):
            # agent_id_cell = Column 16 (Column Q)
            # product_name_cell = Column 5 (Column F)
            # bounce_status_cell = Column 10 (Column K)
        agent_id = sheet.cell_value(row, 16)
        product_name = sheet.cell_value(row, 5)
        if(agent_id != None and (product_name == "Surge Protection Plan" or
                                 product_name == "Electric Repair Essentials" or
                                 product_name == "Surge Protection Plan (20% Off)" or
                                 product_name == "Cooling Maintenance Essentials (6 Month Free Trial - Nest Bundle)" or
                                 product_name == "Cooling Repair & Maintenance Essentials" or
                                 product_name == "Electric Repair Essentials (20% Off)") or
                                 product_name == "Heating & Cooling Repair Essentials"):
            #The format is [agent ID, Product Name, Bounce Status]
            values.append(agent_id)
            # print (agent_id, product_name)

    return values

def get_fcp_sales(filename):
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

        #Format returned in [agent_id]
        if sheet.cell_value(row, 61) != "":
            agent_id = sheet.cell_value(row,61)
            values.append(agent_id)
            # print(agent_id)

    return values

def get_HIVE_new_service(filename):
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

    #Collect up the warranty sales add them to the value arra and return the array.
    for row in range(1, nrows):
            # agent_id_cell = Column 16 (Column Q)
            # product_name_cell = Column 5 (Column F)
            # bounce_status_cell = Column 10 (Column K)
        agent_id = sheet.cell_value(row, 16)
        product_name = sheet.cell_value(row, 5)
        bounce_status = sheet.cell_value(row, 10)
        account_number = sheet.cell_value(row, 0)
        order_number = sheet.cell_value(row, 1)

        if (agent_id != None and
          (product_name == "Home Hero 24 - ONC" or
           product_name == "Home Hero 24 - CNP" or
           product_name == "Home Hero 24 - AEPC" or
           product_name == "Home Hero 24 - AEPN" or
           product_name == "Home Hero 24 - TNMP") and
          (bounce_status == "Accepted" or
	       bounce_status == "Scheduled" or
	       bounce_status == "No deposit due" or
	       bounce_status == "Ercot/ISO Processing" or
	       bounce_status == "Deposit due in first bill" or
	       bounce_status == "Deposit paid" or
	       bounce_status == "Deposit waiver accepted")):
            #The format is [agent ID]
            product_name = "Home Hero 24" #force the plan name so that we can remove duplicates later
            values.append([agent_id, account_number, order_number, product_name])
            # print([agent_id, account_number, order_number, product_name])
            # print (sheet.cell_value(row,17), sheet.cell_value(row,6), sheet.cell_value(row,11))

    return values

def get_HIVE_renewals(filename):
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

    #Collect up the warranty sales add them to the value array and return the array.
    for row in range(1, nrows):
            # agent_id_cell = Column 16 (Column Q)
            # product_name_cell = Column 5 (Column F)
            # bounce_status_cell = Column 10 (Column K)
        agent_id = sheet.cell_value(row, 19)
        product_name = sheet.cell_value(row, 11)
        bounce_status = sheet.cell_value(row, 3)
        account_number = sheet.cell_value(row, 1)
        order_number = sheet.cell_value(row, 0)

        if (agent_id != None and
          (product_name == "Home Hero 24" or
           product_name == "Home Hero 24 - ONC" or
           product_name == "Home Hero 24 - CNP" or
           product_name == "Home Hero 24 - AEPC" or
           product_name == "Home Hero 24 - AEPN" or
           product_name == "Home Hero 24 - TNMP") and
          (bounce_status == "Accepted" or
	       bounce_status == "Scheduled" or
	       bounce_status == "No deposit due" or
	       bounce_status == "Ercot/ISO Processing" or
	       bounce_status == "Deposit due in first bill" or
	       bounce_status == "Deposit paid" or
	       bounce_status == "Deposit waiver accepted")):
            #The format is [agent ID]
            product_name = "Home Hero 24" #force the plan name so that we can remove duplicates later
            values.append([agent_id, account_number, order_number, product_name])
            # print (sheet.cell_value(row,17), sheet.cell_value(row,6), sheet.cell_value(row,11))

    return values
