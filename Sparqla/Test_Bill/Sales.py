from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import  sleep
import unittest
from Utils.Excel import ExcelUtils
from Utils.Function import Function_Call
from Utils.Board_rate import Boardrate
from openpyxl.drawing.image import Image
from openpyxl import load_workbook
from openpyxl.styles import Font
import re

FILE_PATH = ExcelUtils.file_path
class SALES(unittest.TestCase):
    def __init__(self,driver):
        self.driver =driver   
        self.wait = WebDriverWait(driver, 30)

    def test_Sales(self,test_case_id):
        driver = self.driver
        wait = self.wait 
        Rate=Boardrate.Todayrate(self)
        print(Rate)        
        Sheet_name = "SALES"                                        
        valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, Sheet_name)
        print(f"'{valid_rows}': valid rows")
        workbook = load_workbook(FILE_PATH)
        sheet = workbook[Sheet_name]
        for row_num in range(2, valid_rows):  
            current_id = sheet.cell(row=row_num, column=1).value  # Column 1 = Test Case Id
            if current_id == test_case_id:
                data = {
                            "Test Case Id": 1,
                            "Test Status": 2,
                            "Actual Status": 3,
                            "EstNo": 4,
                            "SGST": 5,
                            "CGST": 6,
                            "Total": 7,
                            "TagNo": 8,
                            "Old TagNo": 9,
                            "Home Bill": 10,
                            "Non_tagged": 11,
                            "Employee": 12,
                            "Is Partly": 13,
                            "Section": 14,
                            "Product": 15,
                            "Design": 16,
                            "Sub Design": 17,
                            "Pcs": 18,
                            "Purity": 19,
                            "Size": 20,
                            "G.Wt": 21,
                            "Wast(%)": 22,
                            "Wast Wt(g)": 23,
                            "MC Type": 24,
                            "MC": 25,
                            "Rate": 26,
                            "Discount": 27,
                            "Taxable Amt": 28,
                            "Charges": 29
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
        Function_Call.fill_input2(self,'//input[@id="filter_est_no"]',row_data["EstNo"])
        Function_Call.click(self,'//button[@id="search_est_no"]')