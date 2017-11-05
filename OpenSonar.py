#This will open the Bounce Energy Sonar page, log into the site and download the NOPR data

from selenium import webdriver
from selenium.webdriver.support.ui import Select
import time
import os

homeFolder = 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\'
#Auto download the Excel file to the current working directory
profile = webdriver.ChromeOptions()
prefs = {"download.default_directory" : homeFolder}
profile.add_experimental_option("prefs",prefs)

#Open Bounce Sonar page
browser = webdriver.Chrome(chrome_options=profile)
browser.get('https://apps.bounceenergy.com/sonar/')

#Find the username and password elements and log-in to Sonar
try:
	usernameElem = browser.find_element_by_id('UserUsername')
	usernameElem.send_keys('jndiho')
	passwordElem = browser.find_element_by_name('login_pass')
	passwordElem.send_keys('Muguero78&*')
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

#Find the "Today" radio button and click it
browser.find_element_by_xpath(".//input[@type='radio' and @value='today']").click()

#Find the "Generate Report" submit button and click it
try:
	browser.find_element_by_xpath(".//input[@type='submit' and @value='Generate Report']").click()
	time.sleep(10)
finally:
	print("File downloaded to ", 'C:\\Users\\Jackson.Ndiho\\Documents\\Sales\\')
	browser.close()