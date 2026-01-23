from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from time import  sleep
from Utils.Excel import ExcelUtils
from Utils.Function import Function_Call
from Utils.Board_rate import Boardrate
from Test_Bill.Sales import SALES
from Test_Bill.Credit_Card import CreditCard
from Test_Bill.Cheque import Cheque
from Test_Bill.NetBanking import NetBanking
from openpyxl.drawing.image import Image
from openpyxl import load_workbook
from openpyxl.styles import Font
from datetime import datetime
import re
import unittest

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
        Function_Call.click(self,"//span[contains(text(), 'Billing')]")
        Function_Call.click(self,"//span[contains(text(), 'New Bill')]")
        
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
                    "driect":11,
                    "EstNo":12,
                    "SGST":13,
                    "CGST":14,
                    "Total":15,
                    "Discount":16,
                    "Handling_Charges":17,
                    "Return_Charges":18,
                    "Is Credit": 19,
                    "Is Tobe":20,
                    "Received":21,
                    "Credit Due Date":22,
                    "Gift Voucher":23,
                    "Cash":24,
                    "Creditcard":25,
                    "Cheque":26,
                    "NetBanking":27,
                    "BillNo":28,
                    "Keep_it_As":29,
                    "Store As":30,
                    "OrderNo":31,
                    "Amount":32
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
            Function_Call.dropdown_select(self, '//select[@id="id_branch"]',value=row_data['Cost Centre'])
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
        Function_Call.click2(self,'(//button[@class="btn btn-close btn-warning"])[11]')
        Test_id=row_data["Test Case Id"]
        if row_data["Bill Type"] =="SALES":
           Function_Call.click(self,'//input[@id="bill_typesales"]')
        if row_data["Bill Type"]=="SALES & PURCHASE":
           Function_Call.click(self,'//input[@id="bill_type_salesPurch"]')
        if row_data["Bill Type"]=="SALES PURCHASE & RETURN":
           Function_Call.click(self,'//input[@id="bill_type_saleRet"]') 
           #Function_Call.click(self,'//select[@id="filter_bill_no"]') 
           # Locate the dropdown
           dropdown = Select(wait.until(EC.element_to_be_clickable((By.ID, "filter_bill_no"))))
           # Select the last option
           dropdown.select_by_index(len(dropdown.options) - 1)
           Function_Call.click(self,'//button[@id="search_bill_no"]') 
           Function_Call.click(self,'//input[@class="select_est_details"]') 
           Function_Call.click(self,'//a[@id="update_bill_return"]') 
        if row_data["Bill Type"]== "PURCHASE":
           Function_Call.click(self,'//input[@id="bill_type_purchase"]')
        if row_data["Bill Type"]== "ORDER ADVANCE":
           Function_Call.click(self,'//input[@id="bill_type_order_advance"]')   
           Function_Call.fill_input2(self,'//input[@id="filter_order_no"]', row_data["OrderNo"])
           Function_Call.click(self,'//button[@id="search_order_no"]')
           Function_Call.click2(self,'(//button[@class="btn btn-close btn-warning"])[11]')
           Function_Call.fill_input2(self,'//input[@name="billing[bill_amount]"]', row_data["Amount"])
           Function_Call.click(self,'//li[@id="tab_make_pay"]')
        if row_data["Bill Type"]== "SALES RETURN":
           Function_Call.click(self,'//input[@id="bill_type_sales_return"]')   
           dropdown = Select(wait.until(EC.element_to_be_clickable((By.ID, "filter_bill_no"))))
           # Select the last option
           dropdown.select_by_index(len(dropdown.options) - 1)
           Function_Call.click(self,'//button[@id="search_bill_no"]') 
           Function_Call.click(self,'//input[@class="select_est_details"]') 
           Function_Call.click(self,'//a[@id="update_bill_return"]') 
           
           
           
           
            
        if row_data["Bill Type"]== "CREDIT COLLECTION":
           Function_Call.click(self,'//input[@id="bill_type_credit_bill"]')
        if row_data["Bill Type"]== "ORDER DELIVERY":
           Function_Call.click(self,'//input[@id="bill_type_order_del"]')           
        if row_data["Bill Type"]== "Repair Order Delivery":
           Function_Call.click(self,'//input[@id="repair_order_delivery"]')
           
           
           
        print(row_data["Bill Type"])
        if row_data["driect"]=='No' and row_data["Bill Type"]!="ORDER ADVANCE":
            Function_Call.fill_input2(self,'//input[@id="filter_est_no"]',row_data["EstNo"])
            Function_Call.click(self,'//button[@id="search_est_no"]')
            sleep(2)
            Function_Call.click(self,'(//button[@class="btn btn-close btn-warning"])[11]')
            sleep(3)
            Function_Call.click(self,"//a[normalize-space()='Total Summary']")
            Taxable_Sale_Amount = Function_Call.get_text(self, "//span[@class='summary_lbl summary_sale_amt']")
            print(Taxable_Sale_Amount)
            CGST=Function_Call.get_text(self,'//span[@class="summary_lbl sales_cgst"]')
            print(CGST)
            SGST=Function_Call.get_text(self,'//span[@class="summary_lbl sales_sgst"]') 
            print(SGST)
            IGST=Function_Call.get_text(self,'//span[@class="summary_lbl sales_igst"]') 
            print(IGST)
            Sale_Amount=Function_Call.get_text(self,'//span[@class="summary_lbl sale_amt_with_tax"]')
            print(Sale_Amount)
            Purchase_Amount =Function_Call.get_text(self,'//span[@class="summary_lbl summary_pur_amt"]')
            print(Purchase_Amount)

            if Purchase_Amount == '0.00':
                if row_data["Discount"]:
                    errors=Function_Call.fill_input(
                        self,wait,
                        locator=(By.XPATH, '//input[@id="summary_discount_amt"]'),
                        value=row_data["Discount"],
                        pattern = r"^(\d{1,7}(\.\d{1,2})?)?$",
                        field_name="Discount",
                        screenshot_prefix="Discount",
                        row_num=row_num,
                        Sheet_name=Sheet_name)   
                else:
                    pass
            
                if row_data["Handling_Charges"]:
                    errors=Function_Call.fill_input(
                        self,wait,
                        locator=(By.XPATH, '//input[@id="handling_charges"]'),
                        value=row_data["Handling_Charges"],
                        pattern = r"^(\d{1,3}(\.\d{1,2})?)?$",
                        field_name="Handling_Charges",
                        screenshot_prefix="Handling_Charges",
                        row_num=row_num,
                        Sheet_name=Sheet_name)   
                else:
                    pass
                
                if row_data["Return_Charges"]:
                    errors=Function_Call.fill_input(
                        self,wait,
                        locator=(By.XPATH, '//input[@id="return_charges"]'),
                        value=row_data["Return_Charges"],
                        pattern = r"^(\d{1,3}(\.\d{1,2})?)?$",
                        field_name="Return_Charges",
                        screenshot_prefix="Return_Charges",
                        row_num=row_num,
                        Sheet_name=Sheet_name)   
                else:
                    pass
                Function_Call.click(self,'(//button[@class="btn btn-warning next-tab"])[2]')
        
        
        received = Function_Call.get_value(self, '//input[@name="billing[tot_amt_received]"]')
        received_value = float(received)
        cash=row_data["Cash"]
        if row_data['Received']:
            
            value = float(row_data['Received'])  # percentage value

            Result = (received_value * value) / 100

            credit =  Result

            print("Received:", received_value)
            print("Percentage:", value)
            print("Result:", Result)
            print("Final Credit:", credit)
            Function_Call.click(self,'//input[@id="is_credit_yes"]')
            
            if row_data['Is Credit'] == 'Yes':
                errors=Function_Call.fill_input(
                    self,wait,
                    locator=(By.XPATH, '//input[@name="billing[tot_amt_received]"]'),
                    value=credit,
                    pattern = r"^(\d{1,7}(\.\d{1,2})?)?$",
                    field_name="Received",
                    screenshot_prefix="Received",
                    row_num=row_num,
                    Sheet_name=Sheet_name)   
            
                    
                #Function_Call.click(self,'//li[@id="tab_make_pay"]')
            if row_data['Is Tobe'] == 'Yes':
                Function_Call.click(self,'//input[@id="is_to_be_yes"]')
            
            
            if row_data["Credit Due Date"]:
                errors=Function_Call.fill_input(
                    self,wait,
                    locator=(By.XPATH, '//input[@id="credit_due_date"]'),
                    value=row_data["Credit Due Date"],
                    pattern=r"^(0[1-9]|[12][0-9]|3[01])-(0[1-9]|1[0-2])-\d{4}$",
                    field_name="Credit_Due_Date",
                    screenshot_prefix="Credit_Due_Date",
                    row_num=row_num,
                    Sheet_name=Sheet_name,
                    extra_keys = Keys.TAB,
                    Date_range="Yes"
                    )     
            else:
                pass   
                
            if row_data["Cash"]:        
                Received = credit-cash   
            else:
                Received = received_value  
        
        if received_value == 0:
            pay=Function_Call.get_value(self,'//input[@name="billing[pay_to_cus]"]')
            pay_value = float(pay)
            Received = pay_value-cash
        else:
                Received = received_value-cash   
        
        
        print(row_data["Cash"])
        if row_data["Cash"]:
            errors=Function_Call.fill_input(
                self,wait,
                locator=(By.XPATH, '//input[@id="make_pay_cash"]'),
                value=row_data["Cash"],
                pattern = r"\d{1,3}?$",
                field_name="Cash",
                screenshot_prefix="Cash",
                row_num=row_num,
                Sheet_name=Sheet_name)
        else:
            pass
        
        if row_data['Creditcard']=='Yes':
            print(Received)
            test_case_id=row_data['Test Case Id']
            pay = CreditCard.test_Credit_Card(self,test_case_id,Received)
            
        if row_data['Cheque']=='Yes':
            print(Received)
            test_case_id=row_data['Test Case Id']
            pay = Cheque.test_Cheque(self,test_case_id,Received)
        
        if row_data['NetBanking']=='Yes':
            print(Received)
            test_case_id=row_data['Test Case Id']
            pay = NetBanking.test_NetBanking(self,test_case_id,Received)
                 
            
            
            
           