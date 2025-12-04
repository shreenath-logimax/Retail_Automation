# from selenium import webdriver
# from selenium.webdriver.common.by import By
# from selenium.webdriver.common.keys import Keys
# from selenium.webdriver.support.ui import Select
# from selenium.common.exceptions import NoSuchElementException
# from selenium.common.exceptions import NoAlertPresentException
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from Test_lot.othermetal import Othermetal
# from Test_lot.Stone import Stone
# from time import sleep
# import unittest
# from Utils.Excel import ExcelUtils
# from Utils.Function import Function_Call
# from openpyxl import load_workbook
# from time import sleep

# FILE_PATH = ExcelUtils.file_path
# class Lot(unittest.TestCase):  
#     def __init__(self,driver):
#         self.driver =driver   
#         self.wait = WebDriverWait(driver, 30)
#     def test_lot(self):
#         driver = self.driver
#         wait = self.wait
#         wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT,"Toggle navigation"))).click()
#         wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT,"Inventory"))).click()
#         wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT,"Lot Inward"))).click()
#         wait.until(EC.element_to_be_clickable((By.ID,"add_lot"))).click()
#         function_name = "Lot"
#         valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, function_name)
#         workbook = load_workbook(FILE_PATH)
#         sheet = workbook[function_name]
#         # Call the function
#         Lot=ExcelUtils.Lot_details(FILE_PATH, function_name)
#         Window=1
#         beforelist=''
#         row_count=2
#         pcs=[0]
#         for row_num in range(2, valid_rows):   
#             for list in Lot:
#                 lot_no=list
#             # Define columns and dynamically fetch their values   
#                 data = {
#                         "Test Case Id": 1,
#                         "Test Status": 2,
#                         "Actual Status": 3,
#                         "Lot": 4,
#                         "Lot Received": 5,
#                         "Smith": 6,
#                         "StockType": 7,
#                         "Category": 8,
#                         "Purity": 9,
#                         "Section": 10,
#                         "Product": 11,
#                         "Design": 12,
#                         "Sub Design": 13,
#                         "Pcs": 14,
#                         "GWT": 15,
#                         "LWT": 16,           # Adjust if actual LWT column is different
#                         "Other metal": 17,   # shifted to match correct position
#                         "Charge Name": 18,
#                         "Type": 19,
#                         "Charge": 20,
#                         "Purchase MC": 21,
#                         "Purchase MC Type": 22,
#                         "Purchase Wastage": 23,
#                         "Purchase Rate": 24,
#                         "Purchase Rate Type": 25,
#                         "Metal Type": 26,
#                         "Employee": 27
#                 } 
#                 row_data = {key: sheet.cell(row=row_num, column=col).value 
#                                 for key, (col) in data.items()}
#                 print(row_data)
#                 row_lot_data = self.Lotdetails() 
#                 row_no=row_num+1
#                 Next_Lot = sheet.cell(row=row_no, column=4).value  # Column 1 = Test Case Id
                
#                 Create_data=self.create(row_data,lot_no,Next_Lot,row_num,beforelist,Window)
#                 print(Create_data)
#                 Lot.pop(0)
#                 try: 
#                     lot,pcs_count =Create_data
#                     if lot == lot_no:
#                         pcs[0] = pcs[0] +int(pcs_count)
#                         beforelist =lot_no
#                         print(beforelist)
#                         break
#                 except:
#                     if Create_data:
#                         Test_Status,Actual_Status,Lot_id,pcs_count= Create_data
#                         pcs[0] = pcs[0] +int(pcs_count)
#                         Lot_id=Lot_id
#                         print(Lot_id)
#                         sheet.cell(row=row_num, column=2).value = Test_Status
#                         sheet.cell(row=row_num, column=3).value = Actual_Status
#                         workbook.save(FILE_PATH)
#                         workbook.close()
#                         Status = ExcelUtils.get_Status(FILE_PATH,function_name)  
#                         print(Status)   
#                         Pcs_count = sum(map(int, pcs))   # convert each string to int
#                         print(Pcs_count)  
#                         print(type(pcs_count))
#                 data = ExcelUtils.update_Lot_id(FILE_PATH,Lot_id,row_count,Pcs_count,workbook)
#                 pcs_count,message=data
#                 row_count =pcs_count+row_count 
#                 print(row_count)  
#                 print(message)  
#                 pcs[0]=0
#                 print(pcs)
#                 workbook.save(FILE_PATH)
#                 Update_master = ExcelUtils.update_master_status(FILE_PATH,Status,function_name)
#                 workbook.save(FILE_PATH) 
#                 break
        
#     def create(self,row_data,lot_no,Next_Lot,row_num,beforelist,Window): 
#         driver = self.driver
#         wait = self.wait    
#         if beforelist != lot_no:
#             Function_Call.dropdown_select(
#                 self,'//span[@id="select2-lt_rcvd_branch_sel-container"]', 
#                 row_data["Lot Received"],"//input[@type='search']"
#                 )
#             # wait.until(EC.visibility_of_element_located((By.ID,"select2-lt_rcvd_branch_sel-container"))).click()     
#             print("yes1")
#             wait.until(EC.visibility_of_element_located((By.XPATH,"//span[@id='select2-lt_gold_smith-container']/span"))).click()
#             wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).clear()
#             wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(row_data["Smith"])
#             wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(Keys.ENTER)
#             print(row_data["StockType"])
#             if row_data["StockType"]=="Tagged" :
#                 wait.until(EC.element_to_be_clickable((By.XPATH,"//form[@id='lot_form']/div/div/div[3]/div/input[1]"))).click()
#                 print("Added")
#             else: 
#                 wait.until(EC.element_to_be_clickable((By.XPATH,"//form[@id='lot_form']/div/div/div[3]/div/input[2]"))).click()
#                 print("Added")
#             wait.until(EC.visibility_of_element_located((By.XPATH,"//span[@id='select2-category-container']/span"))).click()
#             wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).clear()
#             wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(row_data["Category"],Keys.ENTER)
#             wait.until(EC.visibility_of_element_located((By.XPATH,"//span[@id='select2-purity-container']/span"))).click()
#             wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).clear()
#             wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(row_data["Purity"],Keys.ENTER)
#         else:
#             print("Same to lot")    
#         if row_data["StockType"]=="Non-Tagged" :
#             wait.until(EC.visibility_of_element_located((By.ID,"select2-select_section-container"))).click()
#             wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).clear()
#             wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(row_data["Section"],Keys.ENTER)     
#         else: 
#             print("Tagged Items")
#         wait.until(EC.visibility_of_element_located((By.ID,"select2-select_product-container"))).click()
#         wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).clear()
#         wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(row_data["Product"],Keys.ENTER)
#         sleep(5)
#         driver.find_element(By.ID,"select2-select_design-container").click()
#         print(row_data["Design"])
#         driver.find_element(By.XPATH,"//input[@type='search']").clear()
#         driver.find_element(By.XPATH,"//input[@type='search']").send_keys(row_data["Design"],Keys.ENTER)
#         print('Data Entered Successfully')
#         sleep(5)
#         driver.find_element(By.ID,"select2-select_sub_design-container").click()
#         wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(row_data["Sub Design"])
#         wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(Keys.ENTER)
#         wait.until(EC.visibility_of_element_located((By.ID,"lot_pcs"))).click()
#         wait.until(EC.visibility_of_element_located((By.ID,"lot_pcs"))).clear()
#         wait.until(EC.visibility_of_element_located((By.ID,"lot_pcs"))).send_keys(row_data["Pcs"])
#         wait.until(EC.visibility_of_element_located((By.ID,"lot_gross_wt"))).click()
#         wait.until(EC.visibility_of_element_located((By.ID,"lot_gross_wt"))).clear()
#         wait.until(EC.visibility_of_element_located((By.ID,"lot_gross_wt"))).send_keys(row_data["GWT"])
#         print(row_data["Type"])
#         test_case_id =row_data["Test Case Id"]
#         if row_data["LWT"]=="Yes" :
#             wait.until(EC.element_to_be_clickable((By.XPATH, '//span[@class="input-group-addon input-sm add_tag_lwt"]'))).click()
#             Sheet_name = "Lot_Lwt"
#             LessWeight=Stone.test_tagStone(self,Sheet_name,test_case_id)
#             print(LessWeight)
#             Lwt,Wt_gram,TotalAmount=LessWeight
#             print(LessWeight)
#             print(Lwt)
#             print(Wt_gram)
#             print(TotalAmount)
#         else:
#             print("There is no Less Weight in this product")
#         if row_data["Other metal"]=="Yes":
#                 wait.until(EC.element_to_be_clickable((By.ID,"other_metal_amount"))).click()
#                 Sheet_name = "Lot_othermetal"
#                 Data=Othermetal.test_othermetal(self,Sheet_name,test_case_id)
#                 OtherMetal,OtherMetalAmount =Data
#                 print(OtherMetal)
#                 print(OtherMetalAmount)
#         else:
#                 print("There is no Other Metal in this product")    
#         wait.until(EC.element_to_be_clickable((By.XPATH,"//div[@id='item_details']/div[2]/div/div[6]/div/div/span"))).click() 
#         wait.until(EC.element_to_be_clickable((By.XPATH,"//table[@id='table_charges']/tbody/tr/td[2]/select"))).click()
#         print(row_data["Charge Name"])
#         Select(wait.until(EC.element_to_be_clickable((By.XPATH,"//table[@id='table_charges']/tbody/tr/td[2]/select")))).select_by_visible_text(row_data["Charge Name"])
#         wait.until(EC.element_to_be_clickable((By.XPATH,"//table[@id='table_charges']/tbody/tr/td[3]/select"))).click()
#         Select( wait.until(EC.element_to_be_clickable((By.XPATH,"//table[@id='table_charges']/tbody/tr/td[3]/select")))).select_by_visible_text(row_data["Type"])
#         wait.until(EC.element_to_be_clickable((By.ID,"update_charge_details"))).click()
#         wait.until(EC.element_to_be_clickable((By.ID,"mc_value"))).click()
#         wait.until(EC.element_to_be_clickable((By.ID,"mc_value"))).clear()
#         wait.until(EC.element_to_be_clickable((By.ID,"mc_value"))).send_keys(row_data["Purchase MC"])
#         wait.until(EC.element_to_be_clickable((By.XPATH,"//form[@id='lot_form']/div/div[3]/div/div"))).click()
#         wait.until(EC.element_to_be_clickable((By.ID,"mc_type"))).click()
#         Select(wait.until(EC.element_to_be_clickable((By.ID,"mc_type")))).select_by_visible_text(row_data["Purchase MC Type"])
#         wait.until(EC.element_to_be_clickable((By.ID,"lot_wastage"))).clear()
#         wait.until(EC.element_to_be_clickable((By.ID,"lot_wastage"))).send_keys(row_data["Purchase Wastage"])
#         wait.until(EC.element_to_be_clickable((By.ID,"rate_per_gram"))).click()
#         wait.until(EC.element_to_be_clickable((By.ID,"rate_per_gram"))).clear()
#         wait.until(EC.element_to_be_clickable((By.ID,"rate_per_gram"))).send_keys(row_data["Purchase Rate"])
#         wait.until(EC.element_to_be_clickable((By.ID,"rate_calc_type"))).click()
#         Select(wait.until(EC.element_to_be_clickable((By.ID,"rate_calc_type")))).select_by_visible_text(row_data["Purchase Rate Type"])
#         wait.until(EC.element_to_be_clickable((By.ID,"add_lot_items"))).click()
#         if Next_Lot == lot_no: 
#             pcs=row_data["Pcs"]
#             return lot_no,pcs
#         else:    
#             wait.until(EC.element_to_be_clickable((By.ID,"save_all"))).click()
#             windows = driver.window_handles
#             driver.switch_to.window(windows[1])
#             sleep(5)
#             try:
#                 message = wait.until(EC.element_to_be_clickable((By.XPATH,"//div[@class='alert alert-success alert-dismissable']"))).text
#                 message = message.replace("\n", " ").strip()
#                 message = message.replace("×", "").strip()
#                 print(message)                    
#                 expected_message = "Add Lot! Lot added successfully"
#                 driver.save_screenshot('Lot.png.png')
#                 if message == expected_message:
#                         Test_Status="Pass"
#                         Actual_Status= message
#                 else:
#                     Test_Status="Fail"
#                     Actual_Status= message
#             except:
#                 driver.save_screenshot('Loterror.png.png')
#                 Test_Status="Fail"
#                 Actual_Status="Lot Not Add Successfully"           
#             wait.until(EC.element_to_be_clickable((By.ID,"ltInward-dt-btn"))).click()
#             wait.until(EC.element_to_be_clickable((By.XPATH,"(.//*[normalize-space(text()) and normalize-space(.)='Sa'])[2]/following::li[1]"))).click()
#             wait.until(EC.element_to_be_clickable((By.ID,"select2-metal-container"))).click()
#             wait.until(EC.element_to_be_clickable((By.XPATH,"//span/input"))).clear()
#             wait.until(EC.element_to_be_clickable((By.XPATH,"//span/input"))).send_keys(row_data["Metal Type"],Keys.ENTER)
#             # wait.until(EC.element_to_be_clickable((By.XPATH,"//span[@id='select2-select_emp-container']/span"))).click()
#             # wait.until(EC.element_to_be_clickable((By.XPATH,"//span/input"))).clear()
#             # wait.until(EC.element_to_be_clickable((By.XPATH,"//span/input"))).send_keys(row_data["Employee"],Keys.ENTER)
#             # wait.until(EC.element_to_be_clickable((By.XPATH,"//span[@id='select2-lot_type-container']/span"))).click()
#             # wait.until(EC.element_to_be_clickable((By.XPATH,"//span/input"))).clear()
#             # wait.until(EC.element_to_be_clickable((By.XPATH,"//span/input"))).send_keys(row_data["StockType"],Keys.ENTER)
#             wait.until(EC.element_to_be_clickable((By.XPATH,"//button[@id='lot_inward_search']/i"))).click()
#             sleep(3)
#             wait.until(EC.element_to_be_clickable((By.XPATH,"//input[@type='search']"))).clear()
#             wait.until(EC.element_to_be_clickable((By.XPATH,"//input[@type='search']"))).send_keys(row_data["Product"],Keys.ENTER)
#             sleep(1) 
#             Lot_id = driver.find_element(By.XPATH,"//table[@id='lot_inward_list']/tbody/tr[1]/td[1]").text
#             print(Lot_id)
#             print(type(Lot_id))
#             pcs=row_data["Pcs"]
            
#             windows = driver.window_handles
#             driver.switch_to.window(windows[1])   # switch to window 1 (second window)
#             driver.close()
#             driver.switch_to.window(windows[0])
#             return Test_Status,Actual_Status,Lot_id,pcs
            
        
 
#     def update_Lot_id(self, Lot_id, row_count, pcs):
#         Pcs_count = sum(map(int, pcs))   # convert each string to int
#         print(Pcs_count) 
       
#         workbook = load_workbook(FILE_PATH)
#         sheet = workbook["Tag"]

#         try:
#             Pcs_count = int(Pcs_count)  # ensure it's an integer
#         except:
#             Pcs_count = 1         # fallback if something wrong
#         for i in range(Pcs_count):   
#             row_Num=row_count+i
#             sheet.cell(row=row_Num, column=5).value=Lot_id

#         workbook.save(FILE_PATH)
#         return Pcs_count,"Lot ID Added in Tag sheet successfully"


#     def Lotdetails(self):
#         function_name = "Lot"
#         valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, function_name)
#         workbook = load_workbook(FILE_PATH)
#         sheet = workbook[function_name]
#         sleep(10)
#         for row_num in range(2, valid_rows):
#             # Define columns and dynamically fetch their values
#             data = {
#                 "Test Case Id":1,
#                 "Pcs": 13,
#                 "GWT": 14,
#                 "Wt": 20,
#             }
#             row_Lotdata = {key: sheet.cell(row=row_num, column=col).value 
#                             for key, col in data.items()}
#         print(row_Lotdata)
#         # Pcs= row_data["Pcs"],
#         # GWT = row_data["GWT"],
#         # DWT = row_data["Wt"],
#         # datas = Pcs,GWT,DWT
#         #print(datas)
#         return row_Lotdata


        

#     def is_element_present(self, how, what):
#         try: self.wait.until(EC.element_to_be_clickable(by=how, value=what))
#         except NoSuchElementException as e: return False
#         return True
    
#     def is_alert_present(self):
#         try: self.driver.switch_to_alert()
#         except NoAlertPresentException as e: return False
#         return True
    
#     def close_alert_and_get_its_text(self):
#         try:
#             alert = self.driver.switch_to_alert()
#             alert_text = alert.text
#             if self.accept_next_alert:
#                 alert.accept()
#             else:
#                 alert.dismiss()
#             return alert_text
#         finally: self.accept_next_alert = True

# if __name__ == "__main__":
#     unittest.main()

# -*- coding: utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from Test_lot.othermetal import Othermetal
from Test_lot.Stone import Stone
import win32com.client as win32
from time import sleep
import unittest
from Utils.Excel import ExcelUtils
from Utils.Function import Function_Call
from openpyxl import load_workbook
from time import sleep

FILE_PATH = ExcelUtils.file_path
class Lot(unittest.TestCase):  
    def __init__(self,driver):
        self.driver =driver   
        self.wait = WebDriverWait(driver, 30)
    def test_lot(self):
        driver = self.driver
        wait = self.wait
        wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT,"Toggle navigation"))).click()
        wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT,"Inventory"))).click()
        wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT,"Lot Inward"))).click()
        wait.until(EC.element_to_be_clickable((By.ID,"add_lot"))).click()
        function_name = "Lot"
        valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, function_name)
        workbook = load_workbook(FILE_PATH)
        sheet = workbook[function_name]
        # Call the function
        Lot=ExcelUtils.Lot_details(FILE_PATH, function_name)
        Window=1
        Products=[]
        beforelist=''
        row_count=2
        pcs=[0]
        for row_num in range(2, valid_rows):   
            for list in Lot:
                lot_no=list
            # Define columns and dynamically fetch their values   
                data = {
                        "Test Case Id": 1,
                        "Test Status": 2,
                        "Actual Status": 3,
                        "Lot": 4,
                        "Lot Received": 5,
                        "Smith": 6,
                        "StockType": 7,
                        "Category": 8,
                        "Purity": 9,
                        "Section": 10,
                        "Product": 11,
                        "Design": 12,
                        "Sub Design": 13,
                        "Pcs": 14,
                        "GWT": 15,
                        "LWT": 16,           # Adjust if actual LWT column is different
                        "Other metal": 17,   # shifted to match correct position
                        "Charge Name": 18,
                        "Type": 19,
                        "Charge": 20,
                        "Purchase MC": 21,
                        "Purchase MC Type": 22,
                        "Purchase Wastage": 23,
                        "Purchase Rate": 24,
                        "Purchase Rate Type": 25,
                        "Metal Type": 26,
                        "Employee": 27
                } 
                row_data = {key: sheet.cell(row=row_num, column=col).value 
                                for key, (col) in data.items()}
                print(row_data)
                row_lot_data = self.Lotdetails() 
                row_no=row_num+1
                Next_Lot = sheet.cell(row=row_no, column=4).value  # Column 1 = Test Case Id
                
                Create_data=self.create(row_data,lot_no,Next_Lot,row_num,beforelist,Window,Products)
                print(Create_data)
                Lot.pop(0)
                try: 
                    lot,pcs_count,Product =Create_data
                    if lot == lot_no:
                        pcs[0] = pcs[0] +int(pcs_count)
                        beforelist =lot_no
                        Products.append(Product)
                        print(beforelist)
                        break
                except:
                    if Create_data:
                        Test_Status,Actual_Status,Lot_id,pcs_count= Create_data
                        pcs[0] = pcs[0] +int(pcs_count)
                        Lot_id=Lot_id
                        print(Lot_id)
                        sheet.cell(row=row_num, column=2).value = Test_Status
                        sheet.cell(row=row_num, column=3).value = Actual_Status
                        workbook.save(FILE_PATH)
                        workbook.close()
                        if Products:
                           Products.clear()
                        else:
                            pass      
                        Status = ExcelUtils.get_Status(FILE_PATH,function_name)  
                        print(Status)   
                if row_data["StockType"]=="Tagged":         
                    data = self.update_Lot_id(Lot_id,row_count,pcs,workbook)
                    pcs_count,message=data
                    row_count =pcs_count+row_count 
                    print(row_count)  
                    print(message)  
                pcs[0]=0
                print(pcs)
                Update_master = ExcelUtils.update_master_status(FILE_PATH,Status,function_name) 
                break
        
    def create(self,row_data,lot_no,Next_Lot,row_num,beforelist,Window,Products): 
        driver = self.driver
        wait = self.wait    
        if beforelist != lot_no:
            # Function_Call.click(self,'//span[@id="select2-lt_rcvd_branch_sel-container"]')
            Function_Call.dropdown_select(
                self,'//span[@id="select2-lt_rcvd_branch_sel-container"]', 
                row_data["Lot Received"],"//input[@type='search']"
                )
            #wait.until(EC.visibility_of_element_located((By.ID,"select2-lt_rcvd_branch_sel-container"))).click()     
            # wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).clear()
            # wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(row_data["Lot Received"])
            # wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(Keys.ENTER)
            print("yes1")
            wait.until(EC.visibility_of_element_located((By.XPATH,"//span[@id='select2-lt_gold_smith-container']/span"))).click()
            wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).clear()
            wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(row_data["Smith"])
            wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(Keys.ENTER)
            print(row_data["StockType"])
            if row_data["StockType"]=="Tagged" :
                wait.until(EC.element_to_be_clickable((By.XPATH,"//form[@id='lot_form']/div/div/div[3]/div/input[1]"))).click()
                print("Added")
            else: 
                wait.until(EC.element_to_be_clickable((By.XPATH,"//form[@id='lot_form']/div/div/div[3]/div/input[2]"))).click()
                print("Added")
            wait.until(EC.visibility_of_element_located((By.XPATH,"//span[@id='select2-category-container']/span"))).click()
            wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).clear()
            wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(row_data["Category"],Keys.ENTER)
            wait.until(EC.visibility_of_element_located((By.XPATH,"//span[@id='select2-purity-container']/span"))).click()
            wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).clear()
            wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(row_data["Purity"],Keys.ENTER)
        else:
            print("Same to lot")    
        if row_data["StockType"]=="Non-Tagged" :
            wait.until(EC.visibility_of_element_located((By.ID,"select2-select_section-container"))).click()
            wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).clear()
            wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(row_data["Section"],Keys.ENTER)     
        else: 
            print("Tagged Items")
        wait.until(EC.visibility_of_element_located((By.ID,"select2-select_product-container"))).click()
        wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).clear()
        wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(row_data["Product"],Keys.ENTER)
        sleep(5)
        driver.find_element(By.ID,"select2-select_design-container").click()
        print(row_data["Design"])
        driver.find_element(By.XPATH,"//input[@type='search']").clear()
        driver.find_element(By.XPATH,"//input[@type='search']").send_keys(row_data["Design"],Keys.ENTER)
        print('Data Entered Successfully')
        sleep(5)
        driver.find_element(By.ID,"select2-select_sub_design-container").click()
        wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']"))).send_keys(row_data["Sub Design"],Keys.ENTER)
        wait.until(EC.visibility_of_element_located((By.ID,"lot_pcs"))).click()
        wait.until(EC.visibility_of_element_located((By.ID,"lot_pcs"))).clear()
        wait.until(EC.visibility_of_element_located((By.ID,"lot_pcs"))).send_keys(row_data["Pcs"])
        wait.until(EC.visibility_of_element_located((By.ID,"lot_gross_wt"))).click()
        wait.until(EC.visibility_of_element_located((By.ID,"lot_gross_wt"))).clear()
        wait.until(EC.visibility_of_element_located((By.ID,"lot_gross_wt"))).send_keys(row_data["GWT"])
        print(row_data["Type"])
        test_case_id =row_data["Test Case Id"]
        if row_data["LWT"]=="Yes" :
            wait.until(EC.element_to_be_clickable((By.XPATH, '//span[@class="input-group-addon input-sm add_tag_lwt"]'))).click()
            Sheet_name = "Lot_Lwt"
            LessWeight=Stone.test_tagStone(self,Sheet_name,test_case_id)
            print(LessWeight)
            Lwt,Wt_gram,TotalAmount=LessWeight
            print(LessWeight)
            print(Lwt)
            print(Wt_gram)
            print(TotalAmount)
        else:
            print("There is no Less Weight in this product")
        if row_data["Other metal"]=="Yes":
                wait.until(EC.element_to_be_clickable((By.ID,"other_metal_amount"))).click()
                Sheet_name = "Lot_othermetal"
                Data=Othermetal.test_othermetal(self,Sheet_name,test_case_id)
                OtherMetal,OtherMetalAmount =Data
                print(OtherMetal)
                print(OtherMetalAmount)
        else:
                print("There is no Other Metal in this product")    
        wait.until(EC.element_to_be_clickable((By.XPATH,"//div[@id='item_details']/div[2]/div/div[6]/div/div/span"))).click() 
        wait.until(EC.element_to_be_clickable((By.XPATH,"//table[@id='table_charges']/tbody/tr/td[2]/select"))).click()
        print(row_data["Charge Name"])
        Select(wait.until(EC.element_to_be_clickable((By.XPATH,"//table[@id='table_charges']/tbody/tr/td[2]/select")))).select_by_visible_text(row_data["Charge Name"])
        wait.until(EC.element_to_be_clickable((By.XPATH,"//table[@id='table_charges']/tbody/tr/td[3]/select"))).click()
        Select( wait.until(EC.element_to_be_clickable((By.XPATH,"//table[@id='table_charges']/tbody/tr/td[3]/select")))).select_by_visible_text(row_data["Type"])
        wait.until(EC.element_to_be_clickable((By.ID,"update_charge_details"))).click()
        wait.until(EC.element_to_be_clickable((By.ID,"mc_value"))).click()
        wait.until(EC.element_to_be_clickable((By.ID,"mc_value"))).clear()
        wait.until(EC.element_to_be_clickable((By.ID,"mc_value"))).send_keys(row_data["Purchase MC"])
        wait.until(EC.element_to_be_clickable((By.XPATH,"//form[@id='lot_form']/div/div[3]/div/div"))).click()
        wait.until(EC.element_to_be_clickable((By.ID,"mc_type"))).click()
        Select(wait.until(EC.element_to_be_clickable((By.ID,"mc_type")))).select_by_visible_text(row_data["Purchase MC Type"])
        wait.until(EC.element_to_be_clickable((By.ID,"lot_wastage"))).clear()
        wait.until(EC.element_to_be_clickable((By.ID,"lot_wastage"))).send_keys(row_data["Purchase Wastage"])
        wait.until(EC.element_to_be_clickable((By.ID,"rate_per_gram"))).click()
        wait.until(EC.element_to_be_clickable((By.ID,"rate_per_gram"))).clear()
        wait.until(EC.element_to_be_clickable((By.ID,"rate_per_gram"))).send_keys(row_data["Purchase Rate"])
        wait.until(EC.element_to_be_clickable((By.ID,"rate_calc_type"))).click()
        Select(wait.until(EC.element_to_be_clickable((By.ID,"rate_calc_type")))).select_by_visible_text(row_data["Purchase Rate Type"])
        wait.until(EC.element_to_be_clickable((By.ID,"add_lot_items"))).click()
        if Next_Lot == lot_no: 
            Product=row_data["Product"]
            pcs=row_data["Pcs"]
            return lot_no,pcs,Product
        else:    
            wait.until(EC.element_to_be_clickable((By.ID,"save_all"))).click()
            windows = driver.window_handles
            driver.switch_to.window(windows[1])
            sleep(5)
            try:
                message = wait.until(EC.element_to_be_clickable((By.XPATH,"//div[@class='alert alert-success alert-dismissable']"))).text
                message = message.replace("\n", " ").strip()
                message = message.replace("×", "").strip()
                print(message)                    
                expected_message = "Add Lot! Lot added successfully"
                driver.save_screenshot('Lot.png.png')
                if message == expected_message:
                        Test_Status="Pass"
                        Actual_Status= message
                else:
                    Test_Status="Fail"
                    Actual_Status= message
            except:
                driver.save_screenshot('Loterror.png.png')
                Test_Status="Fail"
                Actual_Status="Lot Not Add Successfully"           
            wait.until(EC.element_to_be_clickable((By.ID,"ltInward-dt-btn"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH,"(.//*[normalize-space(text()) and normalize-space(.)='Sa'])[2]/following::li[1]"))).click()
            wait.until(EC.element_to_be_clickable((By.ID,"select2-metal-container"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH,"//span/input"))).clear()
            wait.until(EC.element_to_be_clickable((By.XPATH,"//span/input"))).send_keys(row_data["Metal Type"],Keys.ENTER)
            # wait.until(EC.element_to_be_clickable((By.XPATH,"//span[@id='select2-select_emp-container']/span"))).click()
            # wait.until(EC.element_to_be_clickable((By.XPATH,"//span/input"))).clear()
            # wait.until(EC.element_to_be_clickable((By.XPATH,"//span/input"))).send_keys(row_data["Employee"],Keys.ENTER)
            # wait.until(EC.element_to_be_clickable((By.XPATH,"//span[@id='select2-lot_type-container']/span"))).click()
            # wait.until(EC.element_to_be_clickable((By.XPATH,"//span/input"))).clear()
            # wait.until(EC.element_to_be_clickable((By.XPATH,"//span/input"))).send_keys(row_data["StockType"],Keys.ENTER)
            wait.until(EC.element_to_be_clickable((By.XPATH,"//button[@id='lot_inward_search']/i"))).click()
            sleep(3)
            wait.until(EC.element_to_be_clickable((By.XPATH,"//input[@type='search']"))).clear()
            if Products:
                Entered_Product=Products[0]
            else:
                Entered_Product=row_data["Product"]
            wait.until(EC.element_to_be_clickable((By.XPATH,"//input[@type='search']"))).send_keys(Entered_Product)
            sleep(5) 
            Lot_id = driver.find_element(By.XPATH,"//table[@id='lot_inward_list']/tbody/tr[1]/td[1]").text
            
            print(Lot_id)
            print(type(Lot_id))
            pcs=row_data["Pcs"]
            
            windows = driver.window_handles
            driver.switch_to.window(windows[1])   # switch to window 1 (second window)
            driver.close()
            driver.switch_to.window(windows[0])
            driver.execute_script("window.scrollBy(0, -300);") 
            return Test_Status,Actual_Status,Lot_id,pcs
            
    def update_Lot_id(self, Lot_id, row_count, pcs,workbook):
        Pcs_count = sum(map(int, pcs))   # convert each string to int
        print(Pcs_count) 
        sheet = workbook["Tag"]

        try:
            Pcs_count = int(Pcs_count)  # ensure it's an integer
        except:
            Pcs_count = 1         # fallback if something wrong
        for i in range(Pcs_count):   
            row_Num=row_count+i
            sheet.cell(row=row_Num, column=5).value=Lot_id

        workbook.save(FILE_PATH)
        workbook.close()
        return Pcs_count,"Lot ID Added in Tag sheet successfully"    
 
    

      



    def Lotdetails(self,TestCaseId):
        print(TestCaseId)
        function_name = "Lot"
        valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, function_name)
        workbook = load_workbook(FILE_PATH)
        sheet = workbook[function_name]
        sleep(3)
        for row_num in range(2, valid_rows):
            current_id = sheet.cell(row=row_num, column=1).value  # Column 1 = Test Case Id
            if current_id == TestCaseId:
            # Define columns and dynamically fetch their values
                data = {
                    "Test Case Id":1,
                    "Pcs": 14,
                    "GWT": 15,
                    "LWt": 16,
                }
                row_Lotdata = {key: sheet.cell(row=row_num, column=col).value 
                                for key, col in data.items()}
                print(row_Lotdata)
                # Pcs= row_data["Pcs"],
                # GWT = row_data["GWT"],
                # DWT = row_data["Wt"],
                # datas = Pcs,GWT,DWT
                #print(datas)
                return row_Lotdata
      


        

    def is_element_present(self, how, what):
        try: self.wait.until(EC.element_to_be_clickable(by=how, value=what))
        except NoSuchElementException as e: return False
        return True
    
    def is_alert_present(self):
        try: self.driver.switch_to_alert()
        except NoAlertPresentException as e: return False
        return True
    
    def close_alert_and_get_its_text(self):
        try:
            alert = self.driver.switch_to_alert()
            alert_text = alert.text
            if self.accept_next_alert:
                alert.accept()
            else:
                alert.dismiss()
            return alert_text
        finally: self.accept_next_alert = True

if __name__ == "__main__":
    unittest.main()
