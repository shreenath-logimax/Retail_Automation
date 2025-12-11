from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import  sleep
import unittest
from Utils.Excel import ExcelUtils
from Utils.Board_rate import Boardrate
from Test_EST.EST_Tag import ESTIMATION_TAG
from Test_EST.EST_Nontag import ESTIMATION_NonTag
from Test_EST.EST_Homebill import ESTIMATION_Homebill
from Test_EST.EST_oldmetal import ESTIMATION_Oldmetal
from openpyxl.drawing.image import Image
from openpyxl import load_workbook
from openpyxl.styles import Font
import re

FILE_PATH = ExcelUtils.file_path
class Billing(unittest.TestCase):
    def __init__(self,driver):
        self.driver =driver   
        self.wait = WebDriverWait(driver, 30)

    def test_Billing(self):
        driver = self.driver
        wait = self.wait 
        
        Rate=Boardrate.Todayrate(self)
        print(Rate)
        
        
        wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT,"Toggle navigation"))).click()
        module=wait.until(EC.invisibility_of_element_located((By.XPATH,"//span[contains(text(), 'Billing')]")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'nearest', inline: 'center'});", module)
        module.click()
        New_Bill=wait.until(EC.invisibility_of_element_located((By.XPATH,"//span[contains(text(), 'New Bill')]")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'nearest', inline: 'center'});", New_Bill)
        New_Bill.click()
        
        Sheet_name = "Billing"                                        
        valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, Sheet_name)
        print(f"'{valid_rows}': valid rows")
        workbook = load_workbook(FILE_PATH)
        sheet = workbook[Sheet_name]
        for row_num in range(2, valid_rows):  
            data = {
                    "Test Case Id": 1,
                    "TestStatus": 2,
                    "ActualStatus": 3,
                    "Cost Centre": 4,
                    "Billing To": 5,
                    "Employee": 6,
                    "Customer Number": 7,
                    "Customer Name": 8,
                    "Delivery Location": 9,
                	"Bill Type": 10,
                }
            row_data = {key: sheet.cell(row=row_num, column=col).value 
                            for key, col in data.items()}
            print(row_data)
            # Call you 'create' method
            Create_data = self.create(row_data, row_num, Sheet_name)
            print(Create_data)
    
    
     