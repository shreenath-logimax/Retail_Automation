from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import  sleep
import unittest
from Utils.Excel import ExcelUtils
from Utils.Function import Function_Call
from Utils.Board_rate import Boardrate
from Test_Bill.Sales import SALES
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
            Create_data = self.create(row_data, row_num, Sheet_name,Rate)
            print(Create_data)
            
    def create(self,row_data, row_num, Sheet_name,Board_Rate):
        driver = self.driver
        wait = self.wait
        driver.refresh()
        Mandatory_field=[]        
        #Cost Centre
        if row_data['Cost Centre']:
            Function_Call.select_visible_text(self, '//select[@id="id_branch"]',value=row_data['Cost Centre'])
        else:
            msg = f"'{None}' → Cost Centre field is mandatory ⚠️"
            Mandatory_field.append("Cost Centre"); print(msg); Function_Call.Remark(self,row_num, msg,Sheet_name)
        
        #Billing To
        Bill_To = {
                "Customer": '//input[@id="billing_for1"]',
                "Company": '//input[@id="billing_for2"]',
                "Supplier":'//input[@id="billing_for3"]'
            }
        print(Bill_To[row_data["Billing To"]])
        a = Bill_To[row_data["Billing To"]]
        Function_Call.click(self,a)
        
        #Employes
        if row_data["Employee"] is not None:
            Function_Call.dropdown_select(self,f"//span[@id='select2-emp_select-container']", row_data["Employee"],'//span[@class="select2-search select2-search--dropdown"]/input')
        else:
            msg = f"'{None}' → Employee field is mandatory ⚠️"
            Mandatory_field.append("Employee"); print(msg); Function_Call.Remark(self,row_num, msg,Sheet_name)
        
        # Customer
        if row_data["Customer Number"]:
            Function_Call.fill_autocomplete_field(self,"bill_cus_name", row_data["Customer Number"])
        else:
            msg = f"'{None}' → Customer field is mandatory ⚠️"
            Mandatory_field.append(msg)
            print(msg)
            Function_Call.Remark(row_num, msg)
            sleep(3)
        Function_Call.click(self,'(//button[@class="btn btn-close btn-warning"])[11]')
        Test_id=row_data["Test Case Id"]
        
       
        bill=row_data["Bill Type"]   
        print(bill)
        match  bill:
            case"SALES":
                Function_Call.click(self,'//input[@id="bill_typesales"]')
                SALES.test_Sales(self,Test_id)
                
                
                
                # if row_data["EstNo"]:
                #     Function_Call.fill_input2(self,'//input[@id="filter_est_no"]',row_data["EstNo"])
                #     Function_Call.click(self,'//button[@id="search_est_no"]')
            
            
            
            
        