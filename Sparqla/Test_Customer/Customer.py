from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import random
from selenium.webdriver.support.ui import Select
from time import sleep
import re
import time
import unittest
from Utils.Excel import ExcelUtils
from Test_gettag.getttag import GetTag
from Test_Customer.less import Stone
from openpyxl.drawing.image import Image
from openpyxl import load_workbook
import math
from selenium.webdriver import ActionChains
from openpyxl.styles import Font


FILE_PATH = ExcelUtils.file_path
class CustomerOrderTR(unittest.TestCase):
    def __init__(self, driver):
        self.driver = driver
        self.wait = WebDriverWait(driver, 30)
        self.Mandatory_field = []
    def test_customer_order_t_r(self):
        driver = self.driver
        wait = self.wait
        function_name = "Customer"

        # Read Excel data
        valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, function_name)
        Customer = ExcelUtils.customer_details(FILE_PATH, function_name)
        print(Customer)
        count = ExcelUtils.Tag_reserve(FILE_PATH, function_name)
        print(count)
        if count != 0:
            TAG = GetTag.test_gettag(self, count)
            print(TAG)
            Update_Tag = ExcelUtils.update_tag_id(FILE_PATH, function_name, TAG)
            print(Update_Tag)
        else:
            print("All tags available")
        # Get Gold rate
        rate_text = wait.until(EC.presence_of_element_located( (By.XPATH, "(//span[@class='header_rate']/b)[3]"))).text
        gold_rate = int(rate_text.replace("INR", "").strip())
        print(gold_rate)
        # Navigate to Customer Orders page
        wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Toggle navigation"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "(.//*[normalize-space(text()) and normalize-space(.)='All Requests'])[1]/following::span[1]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "(.//*[normalize-space(text()) and normalize-space(.)='Customer Orders'])[1]/following::span[1]"))).click()
        # Load Excel sheet
        workbook = load_workbook(FILE_PATH)
        sheet = workbook[function_name]
        beforelist = ''
        row = 1
        for row_num in range(2, valid_rows):
            for customer_No in Customer:
                Cus_No = customer_No
                # Map Excel columns to keys
                data_map = {
                    "TestCaseId": 1,
                    "TestStatus": 2,
                    "ActualStatus": 3,
                    "Customer Number": 4,
                    "Customer Name": 5,
                    "OrderBranch": 6,
                    "Employee": 7,
                    "BalanceType": 8,
                    "OrderType": 9,
                    "RateType": 10,
                    "TagScan": 11,
                    "Product": 12,
                    "Design": 13,
                    "SubDesign": 14,
                    "Purity": 15,
                    "GrossWt": 16,
                    "LessWt": 17,
                    "Size": 18,
                    "Pcs": 19,
                    "Wast%": 20,
                    "Wast_Wgt": 21,
                    "MC_Type": 22,
                    "MC_Value": 23,
                    "OtherCharge": 24,
                    "ChargeName": 25,
                    "Rate": 26,
                    "Description": 27,
                    "OrderEdit": 28,
                    "UpdateGwt": 29,
                    "Remove": 30,
                    "Field_validation_satus": 31
                }

                row_data = {key: sheet.cell(row=row_num, column=col).value 
                            for key, col in data_map.items()}
                #wchich is empty
                keys_to_check = ["Customer Number", "OrderBranch", "Employee", "TagScan",
                                 "Purity", "GrossWt", "Pcs", "Rate", "Description"]
                print(row_data)
                row_no = row_num + 1
                #Take the after column cus No 
                Next_No = sheet.cell(row=row_no, column=4).value
                
                # Call your 'create' method
                Create_data = self.create(row_data, Cus_No, Next_No, beforelist, row, gold_rate, keys_to_check, row_num)
                print(Create_data)
                beforelist=Cus_No

                # Remove processed customer from the list
                Customer.pop(0)
                customer_No =Create_data
                if customer_No == Cus_No:
                    row=row+1
                    break
                else:
                    if Create_data:
                        Test_Status,Actual_Status= Create_data
                        row=1
                        self.update_excel_status(row_num, Test_Status, Actual_Status, function_name)
                        break
                
                
    def create(self, row_data, Cus_No, Next_No, beforelist, row, gold_rate, keys_to_check, row_num):
        driver, wait = self.driver, self.wait
        test_case_id = row_data["TestCaseId"]
        Mandatory_field = []
        # ------------------------------
        # Step 1: Open Order if Customer is new
        # ------------------------------
        if beforelist != Cus_No:
            wait.until(EC.element_to_be_clickable((By.ID, "add_Order"))).click()
            self.fill_customer_branch_employee(row_data, row_num, Mandatory_field)
            self.select_order_balance_type_rate(row_data)

            # Scroll up if needed
            if row >= 5:
                driver.execute_script(f"window.scrollBy(0, -{row*100});")
        else:
            print("Same Customer, skipping header")

        # ------------------------------
        # Step 2: Handle Order Types
        # ------------------------------
        if row_data["OrderType"] == "Tag Reserve":
            return self.handle_tag_reserve(row_data, row, row_num, keys_to_check,Mandatory_field,gold_rate, Next_No, Cus_No)

        if row_data["OrderType"] == "Customized Order":
            return self.handle_customized_order(row_data, row, row_num, test_case_id, keys_to_check,gold_rate, Next_No, Cus_No)

        # ------------------------------
        # Step 3: Post-processing / Create Order
        # ------------------------------
        

    
    def fill_autocomplete_field(self, field_id, value):
        driver, wait = self.driver, self.wait
        field = wait.until(EC.element_to_be_clickable((By.ID, field_id)))
        field.click()
        field.clear()
        field.send_keys(value)
        time.sleep(2)
        field.send_keys(Keys.BACKSPACE)
        wait.until(EC.presence_of_element_located((By.XPATH, f"//li[contains(text(),'{value}')]"))).click()

    
    
    def fill_customer_branch_employee(self, row_data, row_num, Mandatory_field):
        driver, wait = self.driver, self.wait

        # Customer
        if row_data["Customer Number"]:
            self.fill_autocomplete_field("cus_name", row_data["Customer Number"])
        else:
            msg = f"'{None}' → Customer field is mandatory ⚠️"
            Mandatory_field.append(msg)
            print(msg)
            CustomerOrderTR.Remark(row_num, msg)

        # Order Branch
        if row_data["OrderBranch"]:
            self.select1_fill(wait, "//span[@id='select2-branch_select-container']/span", row_data["OrderBranch"])
        else:
            msg = f"'{None}' → Order Branch field is mandatory ⚠️"
            Mandatory_field.append(msg)
            print(msg)
            CustomerOrderTR.Remark(row_num, msg)

        # Employee
        if row_data["Employee"]:
            self.select1_fill(wait, "//*[@id='order_submit']/div[1]/div[1]/div/div[6]/div/span/span[1]/span", row_data["Employee"])
        else:
            msg = f"'{None}' → Employee field is mandatory ⚠️"
            Mandatory_field.append(msg)
            print(msg)
            CustomerOrderTR.Remark(row_num, msg)

    def select_order_balance_type_rate(self, row_data):
        driver, wait = self.driver, self.wait

        # Balance Type
        balance_map = {"Metal Balance": "metal_bal_type", "Cash Balance": "cash_bal_type"}
        if row_data["BalanceType"] in balance_map:
            wait.until(EC.element_to_be_clickable((By.ID, balance_map[row_data["BalanceType"]]))).click()

        # Order Type
        order_map = {"Tag Reserve": "tag_order", "Customized Order": "customer_order"}
        if row_data["OrderType"] in order_map:
            wait.until(EC.element_to_be_clickable((By.ID, order_map[row_data["OrderType"]]))).click()

        # Rate Type
        rate_map = {"Order Rate(Fixed)": "order_rate", "Delivery Rate": "delivery_rate"}
        if row_data["RateType"] in rate_map:
            wait.until(EC.element_to_be_clickable((By.ID, rate_map[row_data["RateType"]]))).click()

    
    def handle_customized_order(self, row_data, row, row_num, test_case_id, keys_to_check,gold_rate, Next_No, Cus_No):
        driver, wait = self.driver, self.wait
        Mandatory_field= []

        # --- Customer Number check ---
        if row_data["Customer Number"] is None:
            error = CustomerOrderTR.alert1(self)
            Test_Status="Pass"
            Actual_Status= (f"⚠️ Found the message:'{error}'") # prints: Select Order Branch
            print(error)
            wait.until(EC.element_to_be_clickable((By.XPATH,'//button[@class="btn btn-default btn-cancel"]'))).click()
            return Test_Status,Actual_Status 

        # --- Missing keys validation ---
        missing_keys = [k for k in self.Missing_data(row_data, keys_to_check)
                        if k not in ["Customer Number","TagScan","Purity","GrossWt","Pcs","Rate","Description"]]
        if missing_keys:   
            # Add Item Button     
            wait.until(EC.element_to_be_clickable((By.ID,"add_order_item"))).click()           
            try:
                alert_text = CustomerOrderTR.alert(self)
                Test_Status="Pass"
                Actual_Status= (f"⚠️ Found the message:'{alert_text}'") # prints: Select Order Branch
                print(Actual_Status)
            except:
                # No alert → check field value
                print(f'//span[@id="select2-prod_{row}-container"]/span')
                Product =wait.until(EC.presence_of_element_located((By.XPATH,f'//span[@id="select2-prod_{row}-container"]/span'))) 
                Showing_Product = Product.text
                if Showing_Product:
                    Test_Status="Fail"
                    driver.save_screenshot(f"{missing_keys}_{test_case_id}.png")
                    Actual_Status=(f'{missing_keys} → These Field are None to Add successfully ❌')
                    print(Actual_Status)   
            #Cancel Button     
            wait.until(EC.element_to_be_clickable((By.XPATH,'//button[@class="btn btn-default btn-cancel"]'))).click()              
            return Test_Status,Actual_Status 
        else:
            # Add Item Button     
            wait.until(EC.element_to_be_clickable((By.ID,"add_order_item"))).click() 

        if row >= 4:
            driver.execute_script(f"window.scrollBy(0, {row*100});")
            sleep(2)

        # --- Select2 dropdowns ---
        for field, value in [("prod", row_data["Product"]),
                            ("dsgn", row_data["Design"]),
                            ("sub_design", row_data["SubDesign"])]:
            self.select2_fill(wait, f"//span[@id='select2-{field}_{row}-container']", value)

        # --- Purity ---
        if row_data["Purity"]:
            Select(wait.until(EC.element_to_be_clickable((By.ID,f"purity{row}")))).select_by_visible_text(row_data["Purity"])
        else:
            msg = f"'{None}' → Purity field is mandatory ⚠️"
            Mandatory_field.append(msg); print(msg); CustomerOrderTR.Remark(row_num,msg)

        # --- Inputs ---
       # --- GrossWt (max 3 decimals) ---
        Error_field_val = []
        if row_data["GrossWt"]:
            errors = self.fill_input(wait,
            locator=(By.ID, f"weight_{row}"), 
            pattern=r"\d+(\.\d{1,3})?",
            value=row_data["GrossWt"], 
            field_name="GrossWt",
            row_num=row_num,
            screenshot_prefix="GrossWt"
            )
            Error_field_val.extend(errors)
            print(Error_field_val)  # collect errors if any
        else:
            msg = f"'{None}' → Gross weight field is mandatory ⚠️"
            Mandatory_field.append(msg); print(msg); CustomerOrderTR.Remark(row_num, msg)
        
            
            
        # Less WT field     
        if row_data["LessWt"]=="Yes":
            wait.until(EC.element_to_be_clickable((By.XPATH,f'//input[@name="o_item[{row}][less_wt]"]'))).click()
            Sheet_name = "Customer_LWT"
            LessWeight=Stone.test_tagStone(self,Sheet_name,test_case_id) 
        else:
            print("less weight not there")
        time.sleep(2)

        # --- Pcs (integer only) ---
        if row_data["Pcs"]:
            errors = self.fill_input(wait,
                locator=(By.ID, f"qty_{row}"), 
                pattern=r"^\d+$",
                value=row_data["Pcs"], 
                field_name="Pcs",
                screenshot_prefix="Pcs",
                row_num=row_num,
                extra_keys=Keys.ENTER
            )
            Error_field_val.extend(errors)
            print(Error_field_val)
        else:
            msg = f"'{None}' → PCS field is mandatory ⚠️"
            Mandatory_field.append(msg); print(msg); CustomerOrderTR.Remark(row_num, msg)
            
        # Wast% (0–100, 2 decimals max)
        errors =self .fill_input(
                wait,
                locator=(By.NAME, f"o_item[{row}][wast_percent]"),
                value=row_data["Wast%"],
                pattern=r"^\d+(\.\d{1,2})?$",
                field_name="Wast%",
                screenshot_prefix="Wast%",
                range_check=lambda v: 0 <= float(v) <= 100,
                row_num=row_num
            )
        Error_field_val.extend(errors)
        print(Error_field_val)  # collect errors if any
        
        
        # Wait for the MC_Type field    
        wait.until(EC.element_to_be_clickable((By.XPATH,f"(//span[contains(@id, 'id_mc_type')])[{row}]"))).click()
        mc_field=wait.until(EC.element_to_be_clickable((By.XPATH,"/html/body/span[2]/span/span[1]/input")))
        mc_field.clear()
        mc_field.send_keys(row_data["MC_Type"])
        mc_field.send_keys(Keys.ENTER)
        
    
        errors =self.fill_input(
            wait,
            locator=(By.NAME,f"o_item[{row}][mc]"), 
            pattern=r"^\d+(\.\d{1,2})?$",
            value=row_data["MC_Value"], 
            field_name="MC_Value", 
            row_num=row_num,
            screenshot_prefix="MC"
            )
        Error_field_val.extend(errors)
        print(Error_field_val)
        
         # Other Charge        
        if row_data["OtherCharge"]=="Yes":
            self.handle_other_charges(row_data, row)

        # --- Rate ---
        if row_data["Rate"]:
            errors = self.fill_input(
            wait,
            locator=(By.XPATH,f"//*[@id='detail{row}']/td[20]/input"), 
            pattern=r"^\d+(\.\d{1,2})?$",
            value=row_data["Rate"], 
            field_name="Rate", 
            row_num=row_num,
            screenshot_prefix="Rate"
            )
            Error_field_val.extend(errors)
            print(Error_field_val)   # collect errors if any
        
        else:
            rate_field = wait.until(EC.element_to_be_clickable((By.XPATH,f"//*[@id='detail{row}']/td[20]/input")))
            rate_field.clear(); rate_field.send_keys(Keys.TAB)
            driver.save_screenshot(f"Rate_{test_case_id}.png")
            msg = f"'{None}' → Rate field is mandatory ⚠️"
            Mandatory_field.append(msg); print(msg); CustomerOrderTR.Remark(row_num,msg)
            
        # Description Field    
        if row_data["Description"]is not None:
            self.fill_description(row_data, row)
        else:
            mandatory=(f"'{None}' → Description field is mandatory ⚠️")
            Mandatory_field.append(mandatory)
            print(mandatory)
            CustomerOrderTR.Remark(row_num,mandatory)
                
        remove_map = {
            "Product": f'//span[@id="select2-prod_{row}-container"]/span',
            "Design": f'//span[@id="select2-dsgn_{row}-container"]/span',
            "SubDesign": f'//span[@id="select2-sub_design_{row}-container"]/span'
        }
        
        remove=row_data["Remove"]
        if remove:
            xpath = remove_map[remove]
            if xpath:
                close=wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
                # Scroll element into view (works for left/right as well as up/down)
                driver.execute_script("arguments[0].scrollIntoView({block: 'nearest', inline: 'center'});", close)
                close.click()
                placeholder_text = wait.until(EC.presence_of_element_located((By.XPATH, xpath)))
                placeholder_text.click()
                placeholder_name=placeholder_text.text
        else:
            placeholder_name=''
            print('No Data Remove')
            
        missing_keys=self.Missing_data(row_data,keys_to_check)
        print(missing_keys)   
        # Remove 'Customer' and 'OrderBranch' from missing_keys
        keys_to_remove = ["Customer Number", "OrderBranch", "Employee","TagScan",]
        missing_keys = [key for key in missing_keys if key not in keys_to_remove]
        print(Error_field_val)
        remove_items = ["Wast%", "MC_Value"]
        Errors_field_value = [x for x in Error_field_val if x not in remove_items]
        print(Errors_field_value)
        
        if missing_keys or placeholder_name or Errors_field_value:
            try:
                # Add Item Button    
                wait.until(EC.element_to_be_clickable((By.ID,"create_order"))).click()
                alert_text = CustomerOrderTR.alert(self)
                alert_text = CustomerOrderTR.alert(self)

                Test_Status1="Pass"
                Actual_Status= (f"⚠️ Found the message:'{alert_text}'") # prints: Select Order Branch
                print(Actual_Status)
            except:
                # Collect all failed fields in one list
                failed_fields = []
                if missing_keys:
                    failed_fields.extend(missing_keys)
                if Errors_field_value:
                    failed_fields.extend(Errors_field_value)
                if placeholder_name:
                    failed_fields.append(placeholder_name)
                Test_Status1="Fail"
                Actual_Status=(f'{failed_fields} → These Field are None to Add successfully ❌')
                print(Actual_Status)    
                time.sleep(3)
            # Find all matching cancel buttons
            if self.driver.title == "Order - Add | LOGIMAX":
                wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@class="btn btn-default btn-cancel"]'))).click()
                wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@class="btn btn-default btn-cancel"]'))).click()
                print("✅ Cancel button double-clicked")
            else:
                print("ℹ️ Cancel button not found, skipping...")
            return Test_Status1,Actual_Status  
        else:
            pass

        # return "Pass", "Customized Order processed ✅"
        result = self.check_calculation_and_submit(row_data, row, gold_rate, Next_No, Cus_No)
        return result
    
    def handle_tag_reserve(self, row_data, row, row_num, keys_to_check,Mandatory_field,gold_rate, Next_No, Cus_No):
        """
        Handles the Tag Reserve order type.
        Fills the TagScan field, checks for missing mandatory keys,
        validates alerts, and returns Test_Status and Actual_Status.
        """
        driver, wait = self.driver, self.wait
        # Fill TagScan field
        tag_field = wait.until(EC.element_to_be_clickable((By.ID, "est_tag_scan")))
        tag_field.click()
        tag_field.clear()
        tag_field.send_keys(row_data["TagScan"])

        # Get missing keys and remove optional ones
        missing_keys = self.Missing_data(row_data, keys_to_check)
        keys_to_remove = ["TagScan", "Purity", "GrossWt", "Pcs", "Rate", "Description"]
        missing_keys = [key for key in missing_keys if key not in keys_to_remove]

        if missing_keys:
            # Click Tag Search button
            wait.until(EC.element_to_be_clickable((By.ID, "tag_search"))).click()
            try:
                # Look for toaster alert
                toaster_element = wait.until(EC.presence_of_element_located((By.XPATH, "//span[text()='Please select branch']")))
                alert = toaster_element.text
                Test_Status = "Pass"
                Actual_Status = f"⚠️ Found the message:'{alert}'"
                print(Actual_Status)
            except:
                # No alert → check field value
                tag_input = wait.until(EC.element_to_be_clickable((By.NAME, f"o_item[{row}][tag_name]")))
                showing_value = tag_input.get_attribute("value")
                if showing_value == row_data["TagScan"]:
                    Test_Status = "Fail"
                    Actual_Status = f"{missing_keys} → These Field are None to Add successfully ❌"
                    print(Actual_Status)
            # Cancel the form
            wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@class="btn btn-default btn-cancel"]'))).click()
            return Test_Status, Actual_Status
        else:
            wait.until(EC.element_to_be_clickable((By.ID, "tag_search"))).click()
        # Description Field    
        if row_data["Description"]is not None:
            self.fill_description(row_data, row)
        else:
            mandatory=(f"'{None}' → Description field is mandatory ⚠️")
            Mandatory_field.append(mandatory)
            print(mandatory)
            CustomerOrderTR.Remark(row_num,mandatory)
        result = self.check_calculation_and_submit(row_data, row, gold_rate, Next_No, Cus_No)
        return result
                
                
                
    def fill_description(self, row_data, row):
        """Fill Description field or log it as mandatory."""
        driver, wait = self.driver, self.wait
        # Click the edit icon
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//tr[@id='detail{row}']/td[24]/a/i"))).click()
        # Fill description
        desc_field = wait.until(EC.element_to_be_clickable((By.ID, "description")))
        desc_field.clear()
        desc_field.send_keys(str(row_data["Description"]))  # ensure it's string
        # Click Add button
        sleep(2)
        wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Add"))).click()
        print(f"✅ Description filled: {row_data['Description']}")
        
    
        
    def handle_other_charges(self, row_data, row):
        """
        Handles adding 'Other Charges' if applicable.
        Supports multiple charges (comma-separated) and random auto-values.
        """
        driver, wait = self.driver, self.wait

        if row_data.get("OtherCharge") != "Yes":
            print("ℹ️ No other charges needed")
            return

        # Open Other Charge section
        wait.until(EC.element_to_be_clickable((By.XPATH, f"//*[@id='detail{row}']/td[17]/a"))).click()

        charges_raw = row_data["ChargeName"]
        if not charges_raw:
            print("⚠️ OtherCharge flag is Yes but no ChargeName provided")
            return

        charges_list = [s.strip() for s in charges_raw.split(",")]

        for idx, charge in enumerate(charges_list):
            # For the 2nd, 3rd, ... charges → click +Add
            if idx > 0:
                wait.until(EC.element_to_be_clickable((By.ID, "add_new_charge"))).click()

            # Select charge type
            charge_dropdown = wait.until(EC.element_to_be_clickable(
                (By.XPATH, f'(//select[@name="est_stones_item[id_charge][]"])[{idx+1}]')
            ))
            Select(charge_dropdown).select_by_visible_text(charge)

            # Locate corresponding value field (same row as idx+1)
            value_input = wait.until(EC.presence_of_element_located(
                (By.XPATH, f"(//input[@name='est_stones_item[value_charge][]'])[{idx+1}]")
            ))

            current_value = value_input.get_attribute("value").strip()

            # If empty or "0.00" → auto-fill random multiple of 100
            if current_value == "0.00":
                random_value = random.randint(1, 10) * 100
                time.sleep(1)
                value_input.clear()
                value_input.send_keys(str(random_value))
                print(f"⚡ Added random value {random_value} for {charge}")
            else:
                print(f"✅ Auto value {current_value} kept for {charge}")

        # Save button
        wait.until(EC.element_to_be_clickable((By.ID, "update_charge_details"))).click()
        print("Field✅ OtherCharges added:", charges_list)

        # Scroll down if row is beyond 4
        if row >= 4:
            scroll = row * 100
            driver.execute_script(f"window.scrollBy(0, {scroll});")
            time.sleep(2)

    
    
    def Remark(row_num,Field_validation_satus): 
        # Load the workbook
        workbook = load_workbook(FILE_PATH)
        sheet = workbook.active  # or workbook["SheetName"]
        if Field_validation_satus:
            sheet.cell(row=row_num, column=31, value=Field_validation_satus).font = Font(bold=True, color="FF8000")
        # Save workbook
        workbook.save(FILE_PATH) 
        
    def update_excel_status(self,row_num, Test_Status, Actual_Status, function_name):
        print(function_name)
        # Load the workbook
        workbook = load_workbook(FILE_PATH)
        sheet = workbook[function_name]  # or workbook["SheetName"]
        
        if Test_Status== 'Pass':
            # Write Test_Status into column 2
            sheet.cell(row=row_num, column=2, value=Test_Status).font=Font(bold=True, color="00B050")
            
            # Write Actual_Status in col 3 
            sheet.cell(row=row_num, column=3, value=Actual_Status).font = Font(bold=True, color="00B050")
        if Test_Status=='Fail':
            # Write Test_Status into column 2
            sheet.cell(row=row_num, column=2, value=Test_Status).font=Font(bold=True, color="FF0000")
            # Write Actual_Status in col 3 
            sheet.cell(row=row_num, column=3, value=Actual_Status).font = Font(bold=True, color="FF0000")
        # Save workbook
        workbook.save(FILE_PATH)
        # Get status from ExcelUtils
        Status = ExcelUtils.get_Status(FILE_PATH, function_name)
        # Print and return status
        print(Status)
        return Status
                    
    
                    
    def Missing_data(self,row_data,keys_to_check):
        missing_keys = [key for key in keys_to_check if row_data.get(key) is None]
        print(missing_keys)       
        return missing_keys       
    
    def alert(self):
        driver = self.driver
        # Wait until alert is present
        alert = WebDriverWait(driver, 10).until(lambda d: d.switch_to.alert)
        # Get the text from the alert
        alert_text = alert.text
        # Accept the alert (click OK)
        alert.accept()
        return alert_text
    
    def alert1(self):
        wait = self.wait 
        # Wait for toaster message to appear
        wait.until(EC.element_to_be_clickable((By.ID,"add_order_item"))).click()
        alert_msg = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "#toaster .alert"))
        ).text
        alert_text = re.sub(r"[×\s]*Close", "", alert_msg).replace("\n", "").strip()
        # alert_text = alert_msg.replace("×Close", "").replace("\n", "").strip()
        Test_Status="Pass"
        Actual_Status= (f"⚠️ Found the message:'{alert_text}'") # prints: Select Order Branch
        print(Actual_Status)
        return alert_text

    
    # ---------- Helpers ----------
    def add_mandatory(self, field, row_num):
        msg = f"'{None}' → {field} field is mandatory ⚠️"
        self.Mandatory_field.append(msg)
        print(msg)
        CustomerOrderTR.Remark(row_num, msg)

    def select1_fill(self,wait, click_xpath, value):
        """Reusable function to handle select2 dropdowns"""
        wait = self.wait
        wait.until(EC.element_to_be_clickable((By.XPATH, click_xpath))).click()
        input_box = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/span/span/span[1]/input")))
        input_box.clear()
        input_box.send_keys(value, Keys.ENTER)
        time.sleep(1)
   
   
    def select2_fill(self,wait, click_xpath, value):
        """Reusable function to handle select2 dropdowns"""
        wait = self.wait
        sleep(3)
        wait.until(EC.element_to_be_clickable((By.XPATH, click_xpath))).click()
        sleep(2)
        input_box = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/span[2]/span/span[1]/input")))
        input_box.clear()
        input_box.send_keys(value, Keys.ENTER)
        time.sleep(1)
    
    def fill_input(self, wait,locator, value, field_name, row_num,pattern=None, screenshot_prefix="", extra_keys=None, range_check=None):
        """Generic handler for text/numeric fields with validation and optional range check."""
        errors = []
        driver, wait = self.driver, self.wait
        test_case_id = row_num
        field = wait.until(EC.element_to_be_clickable(locator))
        field.click()
        field.clear()
        if value is not None:
            field.send_keys(value)
            if extra_keys:
                field.send_keys(extra_keys)
        entered_value = field.get_attribute("value")

        if entered_value == "":
            driver.save_screenshot(f"{screenshot_prefix}_{test_case_id}.png")
            msg = f"{value} → Not allowed in {field_name} ⚠️"
            CustomerOrderTR.Remark(row_num, msg)
            errors.append(field_name)
            return errors

        # Regex / Range check
        valid = True
        if pattern:
            valid = re.fullmatch(pattern, entered_value) is not None
        if valid and range_check:
            try:
                valid = range_check(float(entered_value))
            except:
                valid = False

        if not valid:
            driver.save_screenshot(f"{screenshot_prefix}_{test_case_id}.png")
            msg = f"'{entered_value}' → Not allowed in {field_name} ❌"
            CustomerOrderTR.Remark(row_num, msg)
            errors.append(field_name)           
        else:
            print(f"'{entered_value}' → Accepted {field_name} ✅")
        return errors
        
    
    def check_calculation_and_submit(self, row_data, row, gold_rate, Next_No, Cus_No):
        driver, wait = self.driver, self.wait

        # Check the Calculation type
        Cal_value=wait.until(EC.presence_of_element_located((By.XPATH,f'(//input[@name="o_item[{row}][calculation_based_on]"])')))
        Cal_current_value = Cal_value.get_attribute("value").strip()
        print(Cal_current_value)
        
        # Call The Calculation function
        Cal_val=CustomerOrderTR.calculation(self,Cal_current_value,row,gold_rate)
        ceil_value,Taxable_Amt,Cal_type= Cal_val
        if ceil_value == Taxable_Amt:
            print("✅ Taxable Amount calculation is correct")
        else:
            driver.save_screenshot(f"{Cal_type}_{row_data['TestCaseId']}.png")
            Test_Status="Fail"
            Actual_Status = f"❌ {Cal_type} incorrect. Expected: {ceil_value}, Table shows: {Taxable_Amt}"
            return Test_Status,Actual_Status
        
        if Next_No == Cus_No:
            return Cus_No
        else:
            if row >= 3:
                # row=row+2
                # scroll = row * 100
                # driver.execute_script(f"window.scrollBy(0, {scroll});") 
                sleep(3)
            else:
                pass    
            
            create_order_btn = wait.until(EC.element_to_be_clickable((By.ID, "create_order")))
            # Scroll the element into view (up/down, left/right if needed)
            driver.execute_script("arguments[0].scrollIntoView({block: 'nearest', inline: 'center'});",create_order_btn)
            sleep(2)
            # Perform JavaScript click (bypasses overlay issues)
            driver.execute_script("arguments[0].click();", create_order_btn)

            try:
                message = wait.until(EC.element_to_be_clickable((By.XPATH,"/html/body/div[1]/div[1]/section[2]/div/div/div/div[2]/div[1]/div/div"))).text
                message =message.replace("×",'').replace("\n",'')
                print(message)                   
                expected_message = "Add Order!New Order added successfully"
                driver.save_screenshot('cus.png.png')
                if message == expected_message:
                        Test_Status="Pass"
                        Actual_Status= f'✅ {message}'
                else:
                    Test_Status="Fail"
                    Actual_Status= f'❌ {message}'
            except:
                driver.save_screenshot('Cuserror.png.png')
                Test_Status="Fail"
                Actual_Status="❌ New Order not added successfully"  
            return Test_Status,Actual_Status
       
    def calculation(self,Cal_current_value,row,gold_rate):
        wait = self.wait 
        data = {
            "0": "Mc & Wast On Gross",
            "1": "Mc & Wast On Net",
            "2": "Mc on Gross, Wast On Net",
            "3": "Fixed Rate",
            "4": "Fixed Rate based on Weight"
        }
        Cal_type = data[str(Cal_current_value)]   # convert int → str because keys are strings
        print(Cal_type)
        
         # --- Helper to fetch field values safely ---
        def get_val(by, locator, cast=float, default=0):
            el = wait.until(EC.presence_of_element_located((by, locator)))
            val = el.get_attribute("value")
            if not val:  
                return default
            return cast(val)

        # Fetch values with one-liners
        Gwt   = get_val(By.NAME, f"o_item[{row}][weight]")
        Lwt   = get_val(By.NAME, f"o_item[{row}][less_wt]")
        Nwt   = get_val(By.NAME, f"o_item[{row}][net_wt]")
        PCS   = get_val(By.NAME, f"o_item[{row}][totalitems]", cast=int)
        Stone = get_val(By.NAME, f"o_item[{row}][stn_amt]")
        Wast_per = get_val(By.NAME, f"o_item[{row}][wast_percent]")
        Wast  = get_val(By.NAME, f"o_item[{row}][wast_wgt]")
        Mc    = get_val(By.NAME, f"o_item[{row}][mc]")
        other_Amt =get_val(By.NAME, f'o_item[{row}][value_charge]',cast=int)

        # Taxable amount kept as string (not converted to float)
        Taxable_Amt = wait.until(
            EC.presence_of_element_located((By.NAME, f"o_item[{row}][taxable]"))
        ).get_attribute("value")

        # MC type
        Mc_type = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, f'//span[contains(@id,"o_item[{row}][id_mc_type]")]')
            )
        ).get_attribute("title")

        # Debug print all values
        print(f"Gwt={Gwt}, Lwt={Lwt}, Nwt={Nwt}, PCS={PCS}, Stone={Stone}, "
            f"Wast_per={Wast_per}, Wast={Wast}, Mc={Mc}, Mc_type={Mc_type}, "
            f"Taxable={Taxable_Amt}")
        gross_weight=Gwt
        net_weight = Nwt  
        wastage_percentage = Wast_per 
        Making_cost_pergram = Mc 
        diamond_cost =Stone
        Charge_Amt = other_Amt
        
        # initialize
        ceil_value = None
        
        if Cal_type=="Mc on Gross, Wast On Net":
           # calculation making cost on gross Wastage% on Net  
            if Mc_type == 'Piece':
                mc =Making_cost_pergram
            else:    
                Mc=Making_cost_pergram*gross_weight
                mc=float('{:.2f}'.format(math.ceil(Mc)))
            Va = (wastage_percentage/100)*net_weight
            Va = round(Va, 3)
            total = net_weight+Va
            total = round(total, 3)
            Cal = (total*gold_rate)+diamond_cost+mc+Charge_Amt
            ceil_value=("{:.2f}".format(math.ceil(Cal)))
            print(ceil_value)
            
            
        if  Cal_type=="Mc & Wast On Net":
            # calculation making cost & Wastage% on Net  
            if Mc_type == 'Piece':
                mc =Making_cost_pergram
            else:    
                Mc=Making_cost_pergram*net_weight
                mc=float("{:.2f}".format(math.ceil(Mc)))
            Va = (wastage_percentage/100)*net_weight
            Va= round(Va, 3)
            total = net_weight+Va
            total = round(total, 3)
            Cal2 = total*gold_rate+mc+diamond_cost+Charge_Amt
            ceil_value=("{:.2f}".format(math.ceil(Cal2)))

                        
        if  Cal_type == "Mc & Wast On Gross":
            # calculation making cost & Wastage% on Gross 81148.00
            if Mc_type == 'Piece':
                mc =Making_cost_pergram
            else:    
                Mc=Making_cost_pergram*gross_weight
                Mc= gross_weight*Making_cost_pergram
                mc=float("{:.2f}".format(math.ceil(Mc)))
           
            Va = (wastage_percentage/100)*gross_weight
            Va = round(Va, 3)
            total= net_weight+Va
            total = round(total, 3)
            cal3 = total*gold_rate+mc+diamond_cost+Charge_Amt
            ceil_value=("{:.2f}".format(math.ceil(cal3)))
            
        if Cal_type== "Fixed Rate based on Weight":
            if Mc_type=='Piece':
                mc = Making_cost_pergram
            else:
                Mc=Making_cost_pergram*gross_weight
                mc=float('{:.2f}'.format(math.ceil(Mc)))
            Va = (wastage_percentage/100)*gross_weight
            Va = round(Va, 3)
            total= net_weight+Va
            total = round(total, 3)
            cal3 = total*gold_rate+mc+diamond_cost+Charge_Amt
            ceil_value=("{:.2f}".format(math.ceil(cal3)))
        
        if Cal_type == "Fixed Rate":
            ceil_value=Taxable_Amt
            

        print(ceil_value)
        print(type(ceil_value))
        return ceil_value,Taxable_Amt,Cal_type







