# -*- coding: utf-8 -*-
"""
Created on Fri Dec 13 10:11:27 2019

@author: MY_EMEDUser07
"""

import pandas as pd
import sys
import time
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.wait import WebDriverWait

sys.path.insert(1, r"M:\Macros\07_EMEDPython3Lib")
#import the module from the path 
from setup_chrome_driver import setup_chrome_driver

	
df = pd.read_excel(r"M:\stock.xlsm")
driver = setup_chrome_driver("M:\Stocks")
df["Filename"] = ""

driver.get("https://www.investing.com/")
driver.find_element_by_link_text("Sign In").click()

time.sleep(30)

for row, column in df.iterrows():
	stockname = column["Name"]
	stocklink = column["Link"]
	
	try:
		driver.get(stocklink)
		time.sleep(2)
		driver.find_element_by_link_text("Historical Data").click()
		dateelement = driver.find_element_by_id("widgetFieldDateRange")
		dateelement2 = driver.find_element_by_id("picker")
		
		time.sleep(3)
		driver.execute_script("arguments[0].innerText = '01/01/2019 - 12/12/2019';", dateelement)
		driver.execute_script("arguments[0].value = '01/01/2019 - 12/12/2019';", dateelement2)
		time.sleep(1)
		driver.find_element_by_id("widgetFieldDateRange").click()
		time.sleep(1)
		driver.find_element_by_id("applyBtn").click()
		time.sleep(3)
		driver.find_element_by_xpath('//*[@id="column-content"]/div[4]/div/a').click()
		
		filename = driver.find_element_by_xpath('//*[@id="leftColumn"]/div[7]/h2').text
		print(filename)
		df.loc[df["Link"] == stocklink, ["Filename"]] = filename
	except:
		continue

df.to_excel(r"M:\stock.xlsx")