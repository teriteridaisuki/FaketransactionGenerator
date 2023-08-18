from selenium.webdriver.chrome.options import Options
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import openpyxl
import sys

exepath = os.path.dirname(os.path.realpath(sys.argv[0]))
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
chrome_driver = r"C:\Program Files\Google\Chrome\Application\chromedriver.exe"
driver = webdriver.Chrome(chrome_driver, chrome_options=chrome_options)
#driver.implicitly_wait(10)
def switchtitle(title):
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if title in driver.title:
            # print(driver.title)
            return driver.title