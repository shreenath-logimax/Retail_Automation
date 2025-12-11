from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import unittest
from Utils.Excel import ExcelUtils
from Utils.Function import Function_Call
from openpyxl import load_workbook
from time import sleep

FILE_PATH = ExcelUtils.file_path
class Boardrate(unittest.TestCase):
    def __init__(self,driver):
        self.driver =driver   
        self.wait = WebDriverWait(driver, 30)

    def Todayrate(self):
            driver = self.driver
            wait = self.wait
            Board_Rate=[]
            Function_Call.click(self,"//span[@class='header_rate']/b[contains(text(),'INR')]")
            rate_text1 = wait.until(EC.presence_of_element_located((By.XPATH, "//li[@class='user-body rate_block_body']//tr[th[contains(text(),'Gold 22KT 1gm')]]/td"))).text
            rate_text2 = wait.until(EC.presence_of_element_located((By.XPATH, "//li[@class='user-body rate_block_body']//tr[th[contains(text(),'Gold 18KT 1gm')]]/td"))).text
            rate_text3 = wait.until(EC.presence_of_element_located((By.XPATH, "//li[@class='user-body rate_block_body']//tr[th[contains(text(),'Silver 1gm')]]/td"))).text
            # Example: "INR 9500"
            gold_rate22KT = int(float(rate_text1.replace("INR", "").strip()))
            Board_Rate.append(gold_rate22KT)
            print(gold_rate22KT)  
            gold_rate18KT = int(float(rate_text2.replace("INR", "").strip()))
            Board_Rate.append(gold_rate18KT)
            print(gold_rate18KT)  
            Silver_rate = int(float(rate_text3.replace("INR", "").strip()))
            Board_Rate.append(Silver_rate)
            print(Silver_rate)     
            return gold_rate22KT,gold_rate18KT,Silver_rate
        