import XLUtils
from selenium import webdriver
import time

driver = webdriver.Chrome(executable_path="C:\\WebDrivers\\chromedriver_win32\\chromedriver.exe")
driver.get("http://demo.guru99.com/test/newtours/")
driver.maximize_window()
driver.implicitly_wait(5)

path = "..\\DDT\\Excel\\docs\\Login1.xlsx"

rows = XLUtils.getRowCount(path, 'Sheet1')

for r in range(2, rows+1):
    username = XLUtils.readData(path, 'Sheet1', r, 1)
    password = XLUtils.readData(path, 'Sheet1', r, 2)

    driver.find_element_by_name("userName").send_keys(username)
    driver.find_element_by_name("password").send_keys(password)
    driver.find_element_by_name('submit').click()

    if driver.title == "Login: Mercury Tours":
        print('Test Passed')
        XLUtils.writeData(path, 'Sheet1', r, 3, 'Test Passed')

    else:
        print('Test Failed')
        XLUtils.writeData(path, 'Sheet1', r, 3, 'Test Failed')
    driver.find_element_by_link_text('Home').click()

driver.implicitly_wait(5)
driver.close()