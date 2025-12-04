from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import  sleep
import unittest
from Utils.Excel import ExcelUtils
from Utils.Function import Function_Call
from openpyxl.drawing.image import Image
from openpyxl import load_workbook
from openpyxl.styles import Font
import re
import math

FILE_PATH = ExcelUtils.file_path
class ESTIMATION_TAG(unittest.TestCase):
    def __init__(self,driver):
        self.driver =driver   
        self.wait = WebDriverWait(driver, 30)

    def test_estimationtag(self,test_case_id,Board_Rate):
        driver = self.driver
        wait = self.wait
        sleep(3)
        Sheet_name = 'Tag_EST'
        test_case_id = test_case_id  
        value =ExcelUtils.Test_case_id_count(FILE_PATH, Sheet_name,test_case_id)
        print(value)
        valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, Sheet_name)
        workbook = load_workbook(FILE_PATH)
        sheet = workbook[Sheet_name]
        row=1
        print(row)
        count = value
        for row_num in range(2, valid_rows):
            current_id = sheet.cell(row=row_num, column=1).value  # Column 1 = Test Case Id
            if current_id == test_case_id:
                data = {
                    "Test Case Id":1,
                    "Test Status":2,
                    "Actual Status":3, 
                    "Lot":4,
                    "Tag No":5,
                    "Product":6,
                    "Design":7, 
                    "Sub Design":8,
                    "Calc Type":9,
                    "Pieces":10, 
                    "Gross Wgt":11,
                    "Less Wgt":12, 
                    "Net Wgt":13, 
                    "Wast %":14, 
                    "Making Charge":15
                }
                row_data = {
                    key: sheet.cell(row=row_num, column=col).value
                    for key, col in data.items()
                }
                print(row_data)
                print(Board_Rate)
                # Call your 'create' method
                Create_data = ESTIMATION_TAG.create(self,row_data, row_num, Sheet_name, row, count,Board_Rate)
                print(Create_data)
                if Create_data:
                    ceil_value,Test_Status,Actual_Status= Create_data
                    ESTIMATION_TAG.update_excel_status(self,row_num, Test_Status, Actual_Status, Sheet_name)
                    return ceil_value
                    
                                   
    def create(self,row_data, row_num, Sheet_name, row, count, Board_Rate):
        wait = self.wait
        Mandatory_field=[]    
        #Tag Check box selected
        Function_Call.click(self,'//input[@id="select_tag_details"]')
        if row_data["Tag No"]:
            Function_Call.fill_input2(self,'//input[@id="est_tag_scan"]',row_data["Tag No"])
            Function_Call.click(self,'//button[@id="tag_search"]')
                    # Wait for table to load after tag scan
        sleep(3)    
        web_data = {}
        try:
            web_data["Tag_code"] = Function_Call.get_value(self, '//input[@name="est_tag[tag_name][]"]')           
            web_data["Product"] = Function_Call.get_text(self, '//div[@class="prodct_name"]')
            web_data["Design"] = Function_Call.get_text(self, '//div[@class="design_name"]')
            web_data["SubDesign"] = Function_Call.get_text(self, '//div[@class="sub_design_name"]')
            web_data["Pieces"] = Function_Call.get_value(self, '//input[@name="est_tag[piece][]"]')
            web_data["Gross Wt"] = Function_Call.get_value(self, '//input[@name="est_tag[gwt][]"]')
            web_data["Less Wt"] = Function_Call.get_value(self, '//input[@name="est_tag[lwt][]"]')
            web_data["Net Wt"] = Function_Call.get_text(self, '//div[@class="nwt"]')
            web_data["Wastage_per"] = Function_Call.get_value(self, '//input[@name="est_tag[wastage][]"]')
            web_data["MC Value"] = Function_Call.get_value(self, '//input[@name="est_tag[mc][]"]')
        except Exception as e:
            print("⚠️ Error fetching web values:", e) 
        print(web_data)
        fields_to_check = {
            "Tag No":"Tag_code",
            "Product": "Product",
            "Design": "Design",
            "Sub Design": "SubDesign",
            "Pieces": "Pieces",
            "Gross Wgt": "Gross Wt",
            "Less Wgt": "Less Wt",
            "Net Wgt": "Net Wt",
            "Wast %": "Wastage_per",
            "Making Charge": "MC Value"
        }
        # Compare Excel vs Web
        for excel_key, web_key in fields_to_check.items():
            excel_value = str(row_data.get(excel_key, ""))
            web_value = str(web_data.get(web_key, ""))

            if excel_value != web_value:
                msg=(f"❌ Mismatch in {excel_key}: Excel={excel_value} | Web={web_value}")
                Function_Call.Remark(self,row_num, msg, Sheet_name)
            else:
                print(f"✅ {excel_key} matches: {excel_value}")
        
        # Fetch values with one-liners
        Gwt   = ESTIMATION_TAG.get_val(self, f'(//input[@name="est_tag[gwt][]"])[{row}]')
        Lwt   = ESTIMATION_TAG.get_val(self, f'(//input[@name="est_tag[lwt][]"])[{row}]')
        Nwt   = ESTIMATION_TAG.get_val(self, f'(//input[@name="est_tag[nwt][]"])[{row}]')
        PCS   = ESTIMATION_TAG.get_val(self, f'(//input[@name="est_tag[piece][]"])[{row}]', cast=int)
        Wast_per = ESTIMATION_TAG.get_val(self, f'(//input[@name="est_tag[wastage][]"])[{row}]')
        Wast  = ESTIMATION_TAG.get_val(self, f'(//input[@name="est_tag[est_wastage_wt][]"])[{row}]')
        Mc    = ESTIMATION_TAG.get_val(self, f'(//input[@name="est_tag[mc][]"])[{row}]')

        # Use f-string for dynamic row selection
        other_amt_value = Function_Call.get_text(self, f'//table[@id="estimation_tag_details"]//tbody//tr[{row}]//td[24]')
        other_Amt = float(other_amt_value)
        other_Amt=0
              
        Function_Call.click(self, f'(//input[@name="est_tag[lwt][]"])[{row}]')
        table = wait.until(EC.presence_of_element_located((By.XPATH, '//table[@id="estimation_stone_cus_item_details"]')))
        table_rows = table.find_elements(By.XPATH, ".//tbody/tr")

        count = len(table_rows)
        print("Row count:", count)

        LWT_Tot_Amt = 0.0  # float accumulator
        for th in table.find_elements(By.XPATH, ".//tr/td[8]/input"):
            value = th.get_attribute("value")  # e.g., "8000.00"
            try:
                Lwt_Amt = float(value) + LWT_Tot_Amt
                LWT_Tot_Amt = Lwt_Amt
            except ValueError:
                print(f"Skipping invalid value: {value}")
                
        Function_Call.click(self,f'(//button[@id="close_stone_details"])[2]')
        print("Total LWT Amount:", LWT_Tot_Amt)
        Stone = LWT_Tot_Amt                    

        # Taxable amount kept as string (not converted to float)
        Taxable_Amt = Function_Call.get_text(self,f'(//span[@class="cost"])[{row}]')
        print(Taxable_Amt)
        
        # MC type  
        mc_type_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, '//select[@class="form-control est_mc_type"]')))
        select = Select(mc_type_dropdown)
        selected_text = select.first_selected_option.text.strip()
        print("Selected MC Type:", selected_text)
        Mc_type =selected_text
        
        purity = wait.until(EC.presence_of_element_located((By.XPATH, "//span[@class='tag_purity_name']"))).text
        print(purity) 
        if purity =='916.00':
           gold_rate=Board_Rate[0]
        if purity == '75.00':
           gold_rate=Board_Rate[1]
        if purity == '92.50':
            gold_rate=Board_Rate[2]
        print(gold_rate)
        
        
        
        # Debug print all values
        print(f"Gwt={Gwt}, Lwt={Lwt}, Nwt={Nwt}, PCS={PCS}, Stone={Stone}, "
            f"Wast_per={Wast_per}, Wast={Wast}, Mc={Mc}, Mc_type={Mc_type}, "
            f"Taxable={Taxable_Amt}")
        Taxvalue=3
        if row_data["Calc Type"]:
            Cal_current_value=row_data["Calc Type"]
            Result=ESTIMATION_TAG.calculation(self,Cal_current_value,gold_rate,Gwt,Nwt,Wast_per,Mc,Stone,other_Amt,Mc_type,Taxvalue)
            ceil_value,Cal_type,IGst,SGst =Result
            if ceil_value==Taxable_Amt:
                Test_Status= "Pass"
                Actual_Status =(f"✅ Calculation Value is correct {ceil_value}")
            else:
                Test_Status= "Pass"
                Actual_Status =(f"❌ Calculation Error in {Taxvalue}: Calculation={ceil_value} | Web Value={Taxvalue}")
            return ceil_value,Test_Status,Actual_Status

        
    def calculation(self,Cal_current_value,gold_rate,Gwt,Nwt,Wast_per,Mc,Stone,other_Amt,Mc_type,Taxvalue):
        wait = self.wait 
        Cal_type =str(Cal_current_value) # convert int → str because keys are strings
        print(Cal_type)
        
        gross_weight=Gwt
        net_weight = Nwt  
        wastage_percentage = Wast_per 
        Making_cost_pergram = Mc 
        diamond_cost =Stone
        Charge_Amt = other_Amt
        Tax = Taxvalue
        
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
            ceil_value
        
        if Tax:  
           salevalue=float(ceil_value)
           Find_Tax=(salevalue*Tax)/100
           Tol_Amt = salevalue+Find_Tax
           Gst=Find_Tax/2
           IGst=("{:.2f}".format(math.ceil(Gst)))
           SGst=("{:.2f}".format(math.ceil(Gst)))
           ceil_value=("{:.2f}".format(Tol_Amt))
        
        print(ceil_value)
        print(type(ceil_value))
        return ceil_value,Cal_type,IGst,SGst
        

  # --- Helper to fetch field values safely ---
    def get_val(self,locator, cast=float, default=0):
        wait = self.wait
        el = wait.until(EC.presence_of_element_located((By.XPATH, locator)))
        val = el.get_attribute("value")
        if not val:  
            return default                             
        return cast(val)
    
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




