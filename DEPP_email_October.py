import sys
import os
import traceback
import csv
import time
import shutil
from datetime import datetime
import win32com.client as win32
from selenium import webdriver
from selenium.webdriver.support.ui import Select

from get_DEPP_sales2 import get_DEPP_sales, get_DEPP_sales_breakdown
from get_October_missing_DEPPs import get_missing_DEPPs, get_missing_DEPPs_breakdown  

from data_files import homeFolder
from data_files import DEPPreportLocation
from data_files import tableNames
from data_files import jaelesiaTeam, tekTeam, antwonTeam, jacksonTeam

from DEPPformat import topOfTable
from DEPPformat import agentRowStart, agentRowEnd
from DEPPformat import agentIDStart, agentIDEnd
from DEPPformat import agentNameStart, agentNameEnd
from DEPPformat import DEPPSalesStart, DEPPSalesEnd
from DEPPformat import DEPPSalesStartGreen, DEPPSalesStartNoColor
from DEPPformat import supRowStart, supRowEnd
from DEPPformat import grandTotalRowStart, grandTotalRowEnd
from DEPPformat import supIDStart, supNameStart
from DEPPformat import supDEPPSalesStart
from DEPPformat import gTotalIDStart, gTotalNameStart
from DEPPformat import gTotalDEPPSalesStart

from DEPPbreakdownTableFormat import emailStartHtml, emailEndHtml
from DEPPbreakdownTableFormat import rowOpenTag, rowCloseTag
from DEPPbreakdownTableFormat import salesDEPPTableOpenTag
from DEPPbreakdownTableFormat import tableCloseTag
from DEPPbreakdownTableFormat import agentNameOpenTag, agentNameCloseTag
from DEPPbreakdownTableFormat import acctNumOpenTag, acctNumCloseTag
from DEPPbreakdownTableFormat import orderNumOpenTag, orderNumCloseTag
from DEPPbreakdownTableFormat import orderStatusOpenTag, orderStatusCloseTag
from DEPPbreakdownTableFormat import DEPPNameOpenTag, DEPPNameCloseTag

arguments = []
for arg in sys.argv:
    arguments.append(arg)
arguments = arguments[1:]

try:
    int(arguments[0])
    reportDate = arguments[0]
except:
    reportDate = ''

currentDate = datetime.now().strftime("%m-%d-%y")
currentTime = time.strftime("%#I:%M %p")
fileNameDate = datetime.now().strftime("%m%d%y")
fileNameTime = time.strftime("%#I%M%p")

#********************************************************************************
#This will open the Bounce Energy Sonar page, log into the site and download the NOPR data
#********************************************************************************

#Auto download the Excel file to the current working directory
profile = webdriver.ChromeOptions()
prefs = {"download.default_directory" : homeFolder}
profile.add_experimental_option("prefs",prefs)

#Open Bounce Sonar page
try:
  browser = webdriver.Chrome(chrome_options=profile)
except Exception as ex:
  traceback.print_exception()

#Open Bounce Sonar page
browser.get('https://apps.bounceenergy.com/sonar/')

#Find the username and password elements and log-in to Sonar
try:
  usernameElem = browser.find_element_by_id('UserUsername')
  usernameElem.send_keys('jndiho')
  passwordElem = browser.find_element_by_name('login_pass')
  passwordElem.send_keys('NyamoYa78&*')
  passwordElem.submit()
except: 
  print('Login to Sonar Failed!')

#Select the New Orders Placed Report from the Report dropdown.
selectReportCategory = Select(browser.find_element_by_id('category_id'))  
time.sleep(5)
selectReportCategory.select_by_value('21')
select_report_type = Select(browser.find_element_by_id('report_id'))
time.sleep(5)
select_report_type.select_by_value('242')
time.sleep(5)

#browser.find_element_by_xpath("//input[@name='username']").click()
#Find the csv checkbox and click it
browser.find_element_by_xpath(".//input[@type='checkbox' and @name='report[report_type]']").click()
#Find the "Today" radio button and click it
browser.find_element_by_xpath(".//input[@type='radio' and @value='last_month']").click()

try:
  #Find the "Generate Report" submit button and click it
  browser.find_element_by_xpath(".//input[@type='submit' and @value='Generate Report']").click()
  time.sleep(10)
finally:
  i = 0
  while (not os.path.isfile(homeFolder + 'report.csv') and i < 120):
    i+=1
    print('i = ', i)
    time.sleep(1)
  print("File downloaded to ", 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\')
  # browser.close()

if (os.path.isfile(homeFolder + 'report.csv') != True):
  print('Failed to download file!')
  sys.exit()

browser.quit()
#*******************************************************************************
#Process the DEPP File and send email
#*******************************************************************************
# Cell Background and Font Styles (to be used to conditionally format cells)
below_goal_text = "9C0006"
below_goal_bg = "FFC7CE"
close_to_goal_text = "9C6500"
close_to_goal_bg = "FFEB9C"
at_or_above_goal_text = "006100"
at_or_above_goal_bg = "C6EFCE"

(jaelesiaDEPPsales, tekDEPPsales, antwonDEPPsales, jacksonDEPPsales,
 totalDEPPsales) = 0, 0, 0, 0, 0

supervisorIDs = {"aervin": 2062007, "jnickerson": 2062001, "tlevon": 2062007,
                 "jacksonn": 2062047, "jabram": 2062017,
                 "iqr_acollins": 2062072, "jmoore": 206223, "mayala": 2062002}



# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Gather up the DEPP sales from the Products report
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
DEPPFileName = homeFolder + 'report.csv'
# DEPPFile = open(DEPPFileName)
with open(DEPPFileName) as DEPPFile:
  DEPPReader = csv.reader(DEPPFile)
  DEPPData = list(DEPPReader)
  DEPPData = DEPPData[1:]
  # print("We made it..************************")

DEPPfilePath = homeFolder + 'report.csv'
missingDEPPfilePath = homeFolder + 'editsOctober.csv'

DEPP_sales = get_DEPP_sales(DEPPfilePath)

missing_DEPPs = get_missing_DEPPs(missingDEPPfilePath)

DEPP_sales_all = [*DEPP_sales, *missing_DEPPs]

print("*************************************************************************")
print("*************************************************************************")
print("*************************************************************************")
# remove any duplicates - there's gotta be a better way to do this!

DUPs_removed = []
for DEPP in DEPP_sales_all:
    if DEPP not in DUPs_removed:
          DUPs_removed.append(DEPP)

DEPP_sales_all = DUPs_removed

print("*************************************************************************")
print("*************************************************************************")
print("*************************************************************************")

DEPP_sales = []

for sale in DEPP_sales_all:
    DEPP_sales.append(sale[0])

# print(DEPP_sales_all)

for id in DEPP_sales:
    if (type(id) == str):
        try:
            DEPP_sales[DEPP_sales.index(id)] = supervisorIDs[id]
        except:
            pass

# Sum up the DEPP sales for each supervisor and for the whole of iQor
for agentID in DEPP_sales:
    if agentID in jaelesiaTeam:
        jaelesiaDEPPsales += 1
        totalDEPPsales += 1
    if agentID in tekTeam:
        tekDEPPsales += 1
        totalDEPPsales += 1
    if agentID in antwonTeam:
        antwonDEPPsales += 1
        totalDEPPsales += 1
    if agentID in jacksonTeam:
        jacksonDEPPsales += 1
        totalDEPPsales += 1

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Run through each entry in the tableNames and build the HTML string to be
# attached to the body of the email
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
html = topOfTable

for agentRow in tableNames:
    agentID = agentRow[0]
    agentName = agentRow[1]
    DEPPSales = ""
    DEPPSalesStart = DEPPSalesStartNoColor

    # This is executed if it is an agent and not a supervisor
    if (type(agentID) == int):  # only agents have numeric IDs

        # Get the agent DEPP Sales
        DEPPSales = DEPP_sales.count(agentID)

        # if (DEPPSales > 0):
        #     print(agentName, "printing green")
        #     DEPPSalesStart = DEPPSalesStartGreen

        DEPPSales = str(DEPP_sales.count(agentID))

        # Add the HTML string for the agent row
        agentID = str(agentID)
        html += (agentRowStart
                 + agentIDStart + agentID + agentIDEnd
                 + agentNameStart + agentName + agentNameEnd
                 + DEPPSalesStart + DEPPSales + DEPPSalesEnd
                 + agentRowEnd)

    # This is executed if it is a supervisor
    if (agentID == 'jaelesia' or agentID == 'tek' or
            agentID == 'antwon' or agentID == 'jackson'):

        if (agentID == 'jaelesia'):            
            DEPPSales = str(jaelesiaDEPPsales)


        elif (agentID == 'tek'):
            DEPPSales = str(tekDEPPsales) 

        elif (agentID == 'antwon'):
            DEPPSales = str(antwonDEPPsales)

        elif (agentID == 'jackson'):
            DEPPSales = str(jacksonDEPPsales)

        # Add the HTMl string for the supervisor
        agentID = "&nbsp;"
        html += (supRowStart
                 + supIDStart + agentID + agentIDEnd
                 + supNameStart + agentName + agentNameEnd
                 + supDEPPSalesStart + DEPPSales + DEPPSalesEnd
                 + supRowEnd)

    # This is executed if it is grand Total
    if agentID == 'grandTotal':
        DEPPSales = str(totalDEPPsales)

        # Add the HTML string for the Grand Total
        agentID = "&nbsp;"
        html += (grandTotalRowStart
                 + gTotalIDStart + agentID + agentIDEnd
                 + gTotalNameStart + agentName + agentNameEnd
                 + gTotalDEPPSalesStart + DEPPSales + DEPPSalesEnd
                 + grandTotalRowEnd + "</table> <br> <br>")

# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# DEPP Sales Breakdown
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------

# The list format that will be returned by get_DEPP_sales_breakdown is:
# [agent_name, pogo_account_number, pogo_order_number,
#  DEPP_name, bounce_status]
DEPP_sales = get_DEPP_sales_breakdown(DEPPFileName)

# remove any duplicates - there is probably a better way to do this!
DUPs_removed = []
for DEPP in DEPP_sales:
    if DEPP not in DUPs_removed:
          DUPs_removed.append(DEPP)
DEPP_sales = DUPs_removed

DEPP_sales.sort()

# for print('DEPPSales', DEPPSales)

html += salesDEPPTableOpenTag

for DEPP in DEPP_sales:
    # format will be [bounce_sale, DEPP_sales]
    # an empty [] means that it is a partially blank row, and
    # one of the two, bounce_sales or DEPP_sales has more rows than the other
    # we will test for this unevenness by checking the length
    # bounceSale = row[0]
    # DEPPSale = row[1]
    
    # if len(row) > 0:
    agentName2 = DEPP[0]
    accountNumber2 = str(int(DEPP[1]))
    orderNumber2 = str(int(DEPP[2]))
    DEPPName = DEPP[3]
    orderStatus2 = DEPP[4]

    print(agentName2, accountNumber2, orderNumber2, DEPPName, orderStatus2)

    html += (rowOpenTag
             + agentNameOpenTag + agentName2 + agentNameCloseTag
             + acctNumOpenTag + accountNumber2 + acctNumCloseTag
             + orderNumOpenTag + orderNumber2 + orderNumCloseTag
             + DEPPNameOpenTag + DEPPName + DEPPNameCloseTag
             + orderStatusOpenTag + orderStatus2 + orderStatusCloseTag
             + rowCloseTag)

html += tableCloseTag + emailEndHtml

# ------------------------------------------------------------------------------
# send email
# ------------------------------------------------------------------------------
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

try:
  # subject = 'iQor DEPP MTD as of ' + currentDate + ' ' + currentTime
  # print("Arguments[0] is: ", arguments[0])
  subject = 'iQor DEPP October Final'
  additionalEmailList = "; ".join(arguments[0:])
  mail.To = additionalEmailList + '; jackson.ndiho@iqor.com'
  mail.Subject = subject
  mail.HtmlBody = subject + ":" + html
  mail.send
except:
  # subject = 'iQor DEPP MTD as of ' + currentDate + ' ' + currentTime
  subject = 'iQor DEPP October Final'
  mail.To = 'jackson.ndiho@iqor.com'
  mail.Subject = subject
  mail.HtmlBody = subject + ":" + html
  mail.send

currentName = homeFolder + 'report.csv'
newName = homeFolder + 'report_MTD_' + fileNameDate + '_' + fileNameTime +'.csv'
shutil.move(currentName, newName)

print("\nDEPP Sales email sent to: " + additionalEmailList
      + "; jackson.ndiho@iqor.com \nat " + currentDate + " " + currentTime 
      + "\n\nDone.......")