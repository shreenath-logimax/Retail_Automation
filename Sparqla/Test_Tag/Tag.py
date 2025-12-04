from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from Test_lot.Lot import Lot
from Test_Tag.Tag_Stone import Tag_Stone
from Test_Tag.Tag_othermetal import  Tag_othermetal
from Utils.Function import Function_Call
from time import sleep
import unittest,math
from Utils.Excel import ExcelUtils
from openpyxl import load_workbook
from time import sleep
import re


FILE_PATH = ExcelUtils.file_path #Give Excel Filepath

class Tag(unittest.TestCase):  
    
    def __init__(self,driver):
        self.driver =driver 
        self.wait = WebDriverWait(driver, 30) 
        
    def test_tag(self):
        driver = self.driver
        wait = self.wait
        
        wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT,"Toggle navigation"))).click()
        wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT,"Inventory"))).click()
        wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT,"Tagging"))).click()
        wait.until(EC.element_to_be_clickable((By.ID,'add_tagging'))).click()
        Function_Call.click(self,"//span[@class='header_rate']/b[contains(text(),'INR')]")
        rate_text1 = wait.until(EC.presence_of_element_located((By.XPATH, "//li[@class='user-body rate_block_body']//tr[th[contains(text(),'Gold 22KT 1gm')]]/td"))).text
        rate_text2 = wait.until(EC.presence_of_element_located((By.XPATH, "//li[@class='user-body rate_block_body']//tr[th[contains(text(),'Gold 18KT 1gm')]]/td"))).text
        rate_text3 = wait.until(EC.presence_of_element_located((By.XPATH, "//li[@class='user-body rate_block_body']//tr[th[contains(text(),'Silver 1gm')]]/td"))).text
        # Example: "INR 9500"
        gold_rate22KT = int(float(rate_text1.replace("INR", "").strip()))
        print(gold_rate22KT)  
        gold_rate18KT = int(float(rate_text2.replace("INR", "").strip()))
        print(gold_rate18KT)  
        Silver_rate = int(float(rate_text3.replace("INR", "").strip()))
        print(Silver_rate)  
        function_name = "Tag"
        valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, function_name)
        workbook = load_workbook(FILE_PATH)
        sheet = workbook[function_name]
        PCS_Count = 0
        for row_num in range(2, valid_rows):
                # Define columns and dynamically fetch their values
            data = {
                    "Test Case Id": 1,
                    "Branch": 4,
                    "Lot No": 5,
                    "Product": 6,
                    "Section": 7,
                    "Design": 8,
                    "Sub Design": 9,
                    "Pieces": 10,
                    "No.of items": 11,
                    "GWT": 12,
                    "Less Weight": 13,
                    "Other Metal": 14,
                    "MC&VA Available": 15,
                    "Wastage%": 16,
                    "Mc Type": 17,
                    "MC": 18,
                    "Size": 19,
                    "Calc Type": 20,
                    "HUID1": 21,
                    "HUID2": 22,
                    "Rate / MRP": 23,
                    "Attribute": 24,
                    "Attribute Name": 25,
                    "Certification":26,
                    "Certification No": 27,
                    "Certification Image": 28,
                    "Button":29
                }
            row_data = {key: sheet.cell(row=row_num, column=col).value 
                        for key, col in data.items()}
            print(row_data)
            if row_num!=2:
                row_no=row_num-1
                before_Lot = sheet.cell(row=row_no, column=5).value
                before_Product= sheet.cell(row=row_no, column=6).value
            else:
                before_Lot=None
                before_Product=None
                
            Current_Lot=row_data["Lot No"]
            Current_Product=row_data["Product"]
    
            if  Current_Lot != before_Lot :
                if row_num !=2:
                   driver.execute_script("window.scrollBy(0, -600);")
                wait.until(EC.element_to_be_clickable((By.ID,"select2-branch_select-container"))).click()
                Branch=wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']")))
                Branch.clear()
                Branch.send_keys(row_data["Branch"],Keys.ENTER)
                sleep(5)
                wait.until(EC.element_to_be_clickable((By.ID,"select2-tag_lot_received_id-container"))).click()
                Lot_No=wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']")))
                Lot_No.clear()
                Lot_No.send_keys(row_data["Lot No"])
                Lot_NO=row_data["Lot No"]
                LOT = wait.until(EC.element_to_be_clickable((By.XPATH, f"//li[normalize-space()='{Lot_NO}']")))
                LOT.click()
                print(Lot_No)
                sleep(2)
            if  Current_Product !=  before_Product: 
                if row_num !=2:
                   driver.execute_script("window.scrollBy(0, -600);")
                wait.until(EC.element_to_be_clickable((By.ID, "select2-tag_lt_prod-container"))).click()
                Product=wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']")))
                Product.send_keys(row_data["Product"],Keys.ENTER)
                sleep(2) 
                wait.until(EC.element_to_be_clickable((By.ID,"select2-section_select-container"))).click()
                Section=wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']")))
                Section.clear()
                Section.send_keys(row_data["Section"],Keys.ENTER)
                
                TestCaseId=row_data["Test Case Id"]
                row_Lotdata=Lot.Lotdetails(self,TestCaseId)
                print (row_Lotdata)
                Pcs= row_Lotdata["Pcs"]
                GWT = row_Lotdata["GWT"]
                DWT = row_Lotdata["LWt"]
                if DWT.upper()=="NO":
                    DWT=0
                    
                Pieces = wait.until(EC.visibility_of_element_located((By.XPATH,"//table[@id='lt-det']/tbody/tr/td[2]"))).text
                Pieces  = int(float(Pieces))
                print(Pieces)
                if Pieces==Pcs :
                    PCS_Count=Pieces
                    print(PCS_Count)
                    print("same pcs")
                else: 
                    print("Pcs not same")
                
                rawValue = wait.until(EC.visibility_of_element_located((By.XPATH,"//table[@id='lt-det']/tbody/tr/td[3]"))).text
                Gross_weigth  = int(float(rawValue))
                print(Gross_weigth)
                if Gross_weigth ==GWT:
                    print("same Gross Weight")
                else: 
                    print("GWT not same")
                
                rawValue2 = wait.until(EC.visibility_of_element_located((By.XPATH,"//table[@id='lt-det']/tbody/tr/td[5]"))).text
                Dia_weight = int(float(rawValue2))
                print(Dia_weight)
                if Dia_weight== DWT:
                    print("same Diamound Weight")
                else: 
                    print("Diamound Wt not same")
                    
                rawValue1 = wait.until(EC.visibility_of_element_located((By.XPATH,"//table[@id='lt-det']/tbody/tr/td[4]"))).text
                Net_weight = int(float(rawValue1))
                print(Net_weight)
                NWT= Gross_weigth-Dia_weight
                if Net_weight== NWT :
                    print("same Net Weight")
                else: 
                    print("Net Wt not same")
                
            else:
                print(f'{row_num} Row Tag runing') 
            if row_num !=2:
                driver.execute_script("window.scrollBy(0, -600);")        

            wait.until(EC.element_to_be_clickable((By.ID,"select2-des_select-container"))).click()
            Design=wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']")))
            Design.clear()
            Design.send_keys(row_data["Design"],Keys.ENTER)
        
            wait.until(EC.element_to_be_clickable((By.ID, "select2-sub_des_select-container"))).click()
            Sub_Design=wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']")))
            Sub_Design.clear() 
            Sub_Design.send_keys(row_data["Sub Design"],Keys.ENTER)
            
            Pieces=wait.until(EC.visibility_of_element_located((By.ID,"tag_pcs")))
            Pieces.clear()
            Pieces.send_keys(row_data["Pieces"])
            
            No_of_items=wait.until(EC.visibility_of_element_located((By.ID,"bulk_tag")))
            No_of_items.clear()
            No_of_items.send_keys(row_data["No.of items"])
            
            GWT=wait.until(EC.visibility_of_element_located((By.ID,"tag_gwt")))
            GWT.clear()
            GWT.send_keys(row_data["GWT"])
            GWT.send_keys(Keys.TAB)
            wait.until(EC.element_to_be_clickable((By.XPATH,'(//button[@id="close_stone_details"])[1]'))).click()
            sleep(2)
            test_case_id =row_data["Test Case Id"]
            if row_data["Less Weight"]=="Yes":
                wait.until(EC.element_to_be_clickable((By.XPATH,"//form[@id='tag_form']/div/div/div[2]/div/div/div[8]/div[2]/div/div/span"))).click()
                Sheet_name='Tag_LWt'
                LessWeight=Tag_Stone.test_tagStone(self,Sheet_name,test_case_id)
                print(LessWeight)
                Lwt,Wt_gram,TotalAmount=LessWeight
                print(LessWeight)
                print(Lwt)
                print(Wt_gram)
                print(TotalAmount)
            else:
                TotalAmount=0
                Wt_gram=0
                print("There is no Less Weight in this product")
            test_case_id =row_data["Test Case Id"]
            if row_data["Other Metal"]=="Yes":
                wait.until(EC.element_to_be_clickable((By.XPATH,"//form[@id='tag_form']/div/div/div[2]/div/div/div[9]/div[2]/div/div/span"))).click()
                Sheet_name='Tag_othermetal'
                Data=Tag_othermetal.test_othermetal(self,Sheet_name,test_case_id)
                OtherMetal,OtherMetalAmount =Data
                print(OtherMetal)
                print(OtherMetalAmount)
            else:
                OtherMetalAmount =0
                print("There is no Other Metal in this product")         
            sleep(3)
            
            if row_data["MC&VA Available"]=="No":
                wait.until(EC.element_to_be_clickable((By.XPATH,"//input[@id='tag_wast_perc']"))).send_keys(row_data["Wastage%"])
            
            else:
                Tab=wait.until(EC.element_to_be_clickable((By.XPATH,"//input[@id='tag_wast_perc']")))
                Tab.send_keys(Keys.TAB)
                
            if row_data["Mc Type"]:
                McType=wait.until(EC.element_to_be_clickable((By.XPATH,'//select[@id="tag_id_mc_type"]')))
                McType.click()
                Select(McType).select_by_visible_text(row_data["Mc Type"])
                wait.until(EC.element_to_be_clickable((By.XPATH,"//input[@id='tag_mc_value']"))).send_keys(row_data["MC"])
            
            if row_data["Size"]:
                wait.until(EC.element_to_be_clickable((By.ID,"select2-tag_size-container"))).click()
                Size=wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']")))
                Size.clear()
                Size.send_keys(row_data["Size"],Keys.ENTER)
           
                   
            if row_data["Calc Type"]:
                CalcType=wait.until(EC.element_to_be_clickable((By.ID,"tag_calculation_based_on")))
                CalcType.click()
                Select(CalcType).select_by_visible_text(row_data["Calc Type"])
                CalculationType=row_data["Calc Type"]
            else:
                select_el = wait.until(EC.presence_of_element_located((By.ID, "tag_calculation_based_on")))
                select = Select(select_el)
                selected_text = select.first_selected_option.text
                print("Calc Type selected ->", selected_text)   # e.g. "Mc & Wast On Gross"
                CalculationType=selected_text
                
               

                metal_text = wait.until(
                    EC.presence_of_element_located((By.XPATH, "//*[@id='lt_metal']"))
                ).text

                print("Metal:", metal_text)
            if metal_text=='GOLD - 75.0000':
                Board_rate=gold_rate18KT
            if metal_text=='GOLD - 916.0000':
                Board_rate=gold_rate22KT
            if metal_text=='Silver - 92.5000':
                Board_rate=Silver_rate
                    
                
            value = Tag.calculation(self,row_data,CalculationType,TotalAmount,Wt_gram,OtherMetalAmount,Board_rate)
            print(value)
            
            value = "{:.2f}".format(float(value))
            print(f'rate{value}')
            HUID1=(row_data["HUID1"])
            print(type(HUID1))
            if HUID1 != None:
                wait.until(EC.element_to_be_clickable((By.ID,"tag_huid"))).click()
                wait.until(EC.visibility_of_element_located((By.ID,"tag_huid"))).send_keys(row_data["HUID1"])
            HUID2=(row_data["HUID2"])
            if HUID2 != None:
                wait.until(EC.element_to_be_clickable((By.ID,"tag_huid2"))).click()
                wait.until(EC.visibility_of_element_located((By.ID,"tag_huid2"))).send_keys(row_data["HUID1"])
            else:
                print('done')
            Attribute = row_data["Attribute Name"]   
            if Attribute != None:
                wait.until(EC.element_to_be_clickable((By.ID,"tag_attribute"))).click()    
                
                wait.until(EC.element_to_be_clickable((By.XPATH,"//span[text()='Select Attribute']"))).click()
                Attribute_Name=wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']")))
                Attribute_Name.clear()
                Attribute_Name.send_keys(row_data["Attribute Name"],Keys.ENTER)
                
                wait.until(EC.element_to_be_clickable((By.XPATH,"//span[text()='Select Attribute Value']"))).click()
                Value=wait.until(EC.visibility_of_element_located((By.XPATH,"//input[@type='search']")))
                Value.clear()
                Value.send_keys(row_data["Value"],Keys.ENTER)
                
                wait.until(EC.element_to_be_clickable((By.ID,"update_attribute_details"))).click()
            else:
                print("There is no Attribute")
            if row_data["Certification"] == "Yes":
                wait.until(EC.element_to_be_clickable((By.ID,"cert_no"))).click()
                wait.until(EC.visibility_of_element_located((By.ID,"cert_no"))).send_keys(row_data["Certification No"])
        
                wait.until(EC.element_to_be_clickable((By.ID,"cert_img"))).click()
                wait.until(EC.element_to_be_clickable((By.ID,"cert_img"))).send_keys(row_data["Certification Image"])
            
            else:
                print('There is no Certification')
            val = wait.until(EC.presence_of_element_located((By.ID,"tag_sale_value")))
            sell_val = val.get_attribute("value")
            print(sell_val)
            if sell_val == value:
                print(f"Tag Calculation Amount Showing correctly ")
            else :
                print(f"Tag Calculation Amount Showing not correctly Actual Amount ")    
            Button=row_data["Button"]  
            print(Button)  
            if Button == "Add":
                Input_Pcs=row_data["Pieces"]
                Input_Pcs=int(Input_Pcs)
                print(type(PCS_Count))
                PCS_Count=PCS_Count-Input_Pcs
                if PCS_Count==1:
                   wait.until(EC.element_to_be_clickable((By.ID,"addTagToPreview"))).click()
                   sleep(3)
                   Function_Call.alert(self)  
                wait.until(EC.element_to_be_clickable((By.ID,"addTagToPreview"))).click()
            else:
                wait.until(EC.element_to_be_clickable((By.ID,"addTagToPreviewAndCopy"))).click() 
            Test_Status = 'Pass'
            Actual_Status="Tagged successfully"
            sheet.cell(row=row_num, column=2).value = Test_Status
            sheet.cell(row=row_num, column=3).value = Actual_Status
            workbook.save(FILE_PATH)
            Status = ExcelUtils.get_Status(FILE_PATH,function_name)  
            print(Status)
            Update_master = ExcelUtils.update_master_status(FILE_PATH,Status,function_name)  
        Tag.update_Tagdetails(self,row_num)
                # driver.find_element(By.XPATH,"(.//*[normalize-space(text()) and normalize-space(.)='Booking Master'])[1]/following::div[3]").click()
    
    
    def update_Tagdetails(self,row_num):
        wait = self.wait
        sheet_name = "Tag_EST"
        workbook = load_workbook(FILE_PATH)
        sheet = workbook[sheet_name]

        # Wait until the table is present
        table = wait.until(EC.presence_of_element_located((By.XPATH, '//table[@id="lt_item_tag_preview"]')))
        table_rows = table.find_elements(By.XPATH, ".//tbody/tr")
        print(f"Total rows found: {len(table_rows)}")

        # Columns you want to copy from the web table
        columns_to_keep = ['Lot', 'Tag No', 'Product', 'Design', 'Sub Design',
                        'Calc Type', 'Pieces', 'Gross Wgt', 'Less Wgt',
                        'Net Wgt', 'Wast %', 'Making Charge']

        # Get all table headers to find indexes
        all_table_headers = [th.text for th in table.find_elements(By.XPATH, ".//thead/tr/th")]
        # Initialize an empty list to store the indexes
        columns_indexes = []
        columns_indexes = [all_table_headers.index(col) for col in columns_to_keep]
        print('oooooo')
        print(columns_indexes)
        # --- Write table headers to Excel (starting from column D) ---
        for col_offset, header in enumerate(columns_to_keep):
            sheet.cell(row=1, column=4 + col_offset, value=header)
        print('IIIIIIIIIII')
        print(col_offset)
        # Loop through each web table row ONCE
        row_idx = 2
        # row_idx = sheet.max_row + 1
        for table_row in table_rows:
            table_cells = table_row.find_elements(By.TAG_NAME, "td")
            row_values = [table_cells[i].text.strip() for i in columns_indexes]
            print("Row Values:", row_values)
            # Clean product name (3rd element → index 2)
            row_values[2] = re.sub(r'-\d+$', '', row_values[2]).strip()
            print(row_values)

            # Write row data to Excel
            for col_offset, value in enumerate(row_values):
                sheet.cell(row=row_idx, column=4 + col_offset, value=value)
            row_idx += 1  # move to next Excel row

        workbook.save(FILE_PATH)
        print("✅ Data merged successfully!")

        
    
    
    
    
    
    
    
    def calculation(self,row_data,CalculationType,TotalAmount,Wt_gram,OtherMetalAmount,Board_rate):
       
        wait = self.wait 
        Nwt_val = wait.until(EC.presence_of_element_located((By.ID,"tag_nwt")))# TAG weight taken form UI
        Nwt = Nwt_val.get_attribute("value")
        Nwt=float(Nwt)
        print(Nwt)
        
        Wast_val = wait.until(EC.presence_of_element_located((By.ID,"tag_wast_perc")))# Wastage % taken form UI
        Wast = Wast_val.get_attribute("value")
        Wast=float(Wast) 
        
        Mc_val = wait.until(EC.presence_of_element_located((By.ID,"tag_mc_value")))# Macking Cost value teken form UI
        Mc = Mc_val.get_attribute("value")
        Mc=float(Mc)  
        
        Mc_type_val = wait.until(EC.presence_of_element_located((By.ID,'tag_id_mc_type')))
        Mc_type = Mc_type_val.get_attribute("value")# Macking Cost value teken form UI
        print(Mc_type) 
        
        gross_weight =row_data["GWT"]# GWT Taken from Excel
        

        gross_weight=float(gross_weight)
        diamond_weight =float(Wt_gram)  
        net_weight = Nwt  
        wastage_percentage = Wast 
        Making_cost_pergram = Mc 
        diamond_cost =float(TotalAmount)
        
        if CalculationType=="Mc on Gross,Wast On Net":
           # calculation making cost on gross Wastage% on Net  
            if Mc_type == '1':
                mc =Making_cost_pergram
            else:    
                Mc=Making_cost_pergram*gross_weight
                mc=float('{:.2f}'.format(Mc))
            Va = (wastage_percentage/100)*net_weight
            Va = round(Va, 3)
            total = net_weight+Va
            total = round(total, 3)
            Cal = (total*Board_rate)+diamond_cost+mc
            ceil_value=("{:.2f}".format(math.ceil(Cal)))
            
            
        if CalculationType=="Mc & Wast On Net":
            # calculation making cost & Wastage% on Net  
            if Mc_type == '1':
                mc =Making_cost_pergram
            else:    
                Mc=Making_cost_pergram*net_weight
                mc=float("{:.2f}".format(math.ceil(Mc)))
            Va = (wastage_percentage/100)*net_weight
            Va= round(Va, 3)
            total = net_weight+Va
            total = round(total, 3)
            Cal2 = total*Board_rate+mc+diamond_cost
            ceil_value=("{:.2f}".format(math.ceil(Cal2)))

                        
        if CalculationType== "Mc & Wast On Gross":
            # calculation making cost & Wastage% on Gross 81148.00
            if Mc_type == '1':
                mc =Making_cost_pergram
            else:    
                Mc=Making_cost_pergram*gross_weight
                Mc= gross_weight*Making_cost_pergram
                mc=float("{:.2f}".format(math.ceil(Mc)))
           
            Va = (wastage_percentage/100)*gross_weight
            Va = round(Va, 3)
            total= net_weight+Va
            total = round(total, 3)
            cal3 = total*Board_rate+mc+diamond_cost
            ceil_value=("{:.2f}".format(math.ceil(cal3)))
            
        if CalculationType == "Fixed Rate based on Weight":
            if Mc_type=='1':
                mc = Making_cost_pergram
            else:
                Mc=Making_cost_pergram*gross_weight
                mc=float('{:.2f}'.format(math.ceil(Mc)))
            Va = (wastage_percentage/100)*gross_weight
            Va = round(Va, 3)
            total= net_weight+Va
            total = round(total, 3)
            cal3 = total*Board_rate+mc+diamond_cost
            ceil_value=("{:.2f}".format(math.ceil(cal3)))
        
        # if row_data["Calc Type"] == "Fixed Rate":
            # ceil_value   
          
        if CalculationType=="Yes":
           Total =float(ceil_value) + float(OtherMetalAmount)
           return Total
                        
        else:
            print(ceil_value)
            print(type(ceil_value))
            return ceil_value
        
        
    
    
    def is_element_present(self, how, what):
        try: self.driver.find_element(by=how, value=what)
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
    
    def tearDown(self):
        self.driver.quit()
        self.assertEqual([], self.verificationErrors)

if __name__ == "__main__":
    unittest.main()
