# import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait

from openpyxl import Workbook

firefox_options = webdriver.FirefoxOptions()
firefox_options.set_preference("zoom.minPercent", 1)
firefox_options.set_preference("toolkit.zoomManager.zoomValues", ".1,.2,.3,.5,.67,.8,.9,1,1.1,1.2,1.33,1.5,1.7,2,2.4,3,4,5")

browser = webdriver.Firefox(options=firefox_options)
browser.maximize_window()

browser.get("https://airtable.com/shrSKpx3IEyBLa3h6")

browser.set_context("chrome")

win = browser.find_element(By.TAG_NAME, "html")

while True:
    win.send_keys(Keys.CONTROL + Keys.SUBTRACT)

    browser.set_context("content")

    try:
        verticalScrollBar = browser.find_element(By.CLASS_NAME, "antiscroll-scrollbar-vertical")
    except:
        break

    browser.set_context("chrome")

browser.set_context("content")

# time.sleep(3)
wait = WebDriverWait(browser, 10)

# Declaring and Initializing Table Data
data = []

leftPane = browser.find_element(By.CLASS_NAME, "leftPaneWrapper")
rightPane = browser.find_element(By.CLASS_NAME, "rightPaneWrapper")

# Get Left Header
leftHeaderPane = leftPane.find_element(By.CLASS_NAME, "headerLeftPane")
leftHeaderRow = leftHeaderPane.find_element(By.CSS_SELECTOR, "div.paneInnerContent > div.headerRow")
leftHeaderCells = leftHeaderRow.find_elements(By.CSS_SELECTOR, "div.cell.header")

# Get Right Header
rightHeaderPane = rightPane.find_element(By.CLASS_NAME, "headerRightPane")
rightHeaderRow = rightHeaderPane.find_element(By.CSS_SELECTOR, "div.paneInnerContent > div.headerRow")
rightHeaderCells = rightHeaderRow.find_elements(By.CSS_SELECTOR, "div.cell.header")

# Get Header Data
header = []

for cell in leftHeaderCells:
    header.append(cell.text)

for cell in rightHeaderCells:
    header.append(cell.text)

data.append(header)

# Get Rows Data
dataLeftPane = leftPane.find_element(By.CLASS_NAME, "dataLeftPane")
leftDataRows = dataLeftPane.find_elements(By.CSS_SELECTOR, "div.dataLeftPaneInnerContent > div.dataRow")

dataRightPane = rightPane.find_element(By.CLASS_NAME, "dataRightPane")
rightDataRows = dataRightPane.find_elements(By.CSS_SELECTOR, "div.dataRightPaneInnerContent > div.dataRow")

leftDataRowsLength = len(leftDataRows)
rightDataRowsLength = len(rightDataRows)

if leftDataRowsLength != rightDataRowsLength:
    raise Exception("Invalid Number of Rows")

for i in range(leftDataRowsLength):
    row = []

    leftDataRow = leftDataRows[i]
    rightDataRow = rightDataRows[i]

    leftCells = leftDataRow.find_elements(By.CLASS_NAME, "cell")
    rightCells = rightDataRow.find_elements(By.CLASS_NAME, "cell")

    for cell in leftCells:
        row.append(cell.text)

    for cell in rightCells:
        row.append(cell.text)

    data.append(row)


# Creating Excel File
workbook = Workbook()

worksheet = workbook.active

for row in data:
    worksheet.append(row)

workbook.save("data.xlsx")

browser.quit()
