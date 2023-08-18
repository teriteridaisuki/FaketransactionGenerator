from selenium.webdriver.chrome.options import Options
import os
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
import pyautogui
import re
import openpyxl
import sys
import cv2


exepath = os.path.dirname(os.path.realpath(sys.argv[0]))
chrome_options = Options()
chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
chrome_driver = r"C:\Program Files\Google\Chrome\Application\chromedriver.exe"
driver = webdriver.Chrome(chrome_driver, chrome_options=chrome_options)
driver.implicitly_wait(10)
rightclick = ActionChains(driver)


def switchtitle(title):
    for handle in driver.window_handles:
        driver.switch_to.window(handle)
        if title in driver.title:
            return driver.title

def checkerselect(place,name):#选择审批人并提交
    driver.find_element(By.XPATH, "//div[text()= '"+place+"']/../../td[2]").click()
    driver.find_elements(By.XPATH, "//div[@class='EI6JOB-u-l' and @style='margin-top:-8px;right:2px;']")[1].click()
    driver.find_element(By.XPATH, "//div[text()= '"+name+"']").click()
    time.sleep(2)
    driver.find_elements(By.XPATH, "//div[text()= '确定']")[1].click()
def duplicate(formnum):#复制之前单据
    switchtitle("财务一体化")
    driver.find_element(By.XPATH, "//img[contains(@src,'bz.svg')]").click()
    driver.find_element(By.XPATH, "//span[text()= '应付及付款']").click()
    driver.find_element(By.XPATH, "//span[text()= '支付内部往来清算款']").click()
    driver.find_element(By.XPATH, "//span[text()= '复制']").click()
    driver.find_element(By.XPATH, "//label[text()= '单据编号：']/../following-sibling::div[1]/div[1]/input[1]").send_keys(formnum)
    driver.find_elements(By.XPATH, "//img[contains(@src,'查询_chaxun_U_B.png')]")[1].click()
    ActionChains(driver).double_click(driver.find_element(By.XPATH, "//div[text()= '"+formnum+"']")).perform()
    time.sleep(5)
    driver.find_element(By.XPATH, "//span[text()= '保存']").click()
    driver.find_element(By.XPATH, "//div[text()= '确定']").click()
    time.sleep(5)
    driver.find_elements(By.XPATH, "//span[text()= '提交']")[1].click()
    driver.find_element(By.XPATH, "//div[text()= '提交']").click()

def submitprocess():
    checkerselect("分公司部门经理", "10034872")
    checkerselect("分公司财务经理", "10034872")
    checkerselect("分公司总会计师", "20000723")
    checkerselect("分公司经理", "20073578")
    driver.find_element(By.XPATH, "//div[text()= '确定']").click()
    driver.find_element(By.XPATH, "//div[text()= '确认']").click()
    driver.find_element(By.XPATH, "//div[text()= '确定']").click()
    driver.find_element(By.XPATH, "//div[text()= '确定']").click()
def main(runtimes,formnum):
    for i in range(1,runtimes+1):
        duplicate(formnum)
        submitprocess()
    pass

def wait(func,*para,clickstat=False,clicknum=None):
    while True:
        try:
            result = func(*para)  # 调用提供的函数
            if clickstat==True:
                if clicknum==None:
                    result.click()
                else:
                    result[clicknum].click()
            return result    # 如果成功运行，返回结果并退出函数
        except Exception:   # 捕获所有异常
            pass  # 如果出现异常，继续循环
main(3,"FFK2011202308003694")



