from selenium import webdriver
from Utils.Excel import ExcelUtils
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from Test_login.Login import Login
from Test_master.Metal import Metal
from Test_master.Category import Category
from Test_master.Product import Product
from Test_master.Design import Design
from Test_master.Subdesign import Subdesign
from Test_master.Designmapping import Designmapping
from Test_master.Subdesignmapping import Subdesignmapping
from Test_master.MCVA import McVa
from Test_lot.Lot import Lot
from Test_Tag.Tag import Tag
from Test_vendor.Vendor import VendorRegistration
from Test_Customer.Customer import CustomerOrderTR
from Test_EST.EST import ESTIMATION
import datetime     
import os

def create_driver():
    """Create Chrome driver. Headless when running inside Jenkins."""
    options = Options()
    options.add_argument("--log-level=3")
    options.add_experimental_option("excludeSwitches", ["enable-logging"])

    # Block camera popup
    options.add_argument("--use-fake-device-for-media-stream")
    options.add_argument("--use-fake-ui-for-media-stream")

  

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    driver.maximize_window()
    return driver

class main():
    @staticmethod
    def main():
        ExcelUtils.ensure_copy_exists()
        FILE_PATH = ExcelUtils.file_path
        # Step 1: Initialize WebDriver
        ExcelUtils.ExcelClose(FILE_PATH)
        ct1 = datetime.datetime.now()
        print('Automation process Started',ct1)

        driver = create_driver()
        # options = Options()
        # options.add_argument("--log-level=3")
        
        
        # options.add_experimental_option('excludeSwitches', ['enable-logging'])

        # # Block camera popup
        # options.add_argument("--use-fake-device-for-media-stream")
        # options.add_argument("--use-fake-ui-for-media-stream")

        # driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
        # driver.maximize_window()
        # options = Options()
        try:
            sheet_names = ExcelUtils.get_sheet_names(FILE_PATH)
            print(sheet_names)
            # Step 2: Get functions to execute from the Master sheet
            functions_to_execute = ExcelUtils.get_master_sheet_data(FILE_PATH)
            print(functions_to_execute)
        
            for function_name in functions_to_execute:
                print("funct",function_name)
                if function_name in sheet_names:
                    print("sheet",sheet_names)
                    # Initialize LoginAutomation
                    match function_name:
                        case "Login":
                            print("yes")
                            login_automation = Login(driver)
                            Data = login_automation.test_login()
                        case "Metal":
                            print("yes")
                            Metal_automation = Metal(driver)
                            Data = Metal_automation.test_metal()
                        case "CategoryName":
                            print("yes")
                            Category_automation = Category(driver)
                            Data = Category_automation.test_category()    
                        case "Product":
                            print("yes")
                            Product_automation =Product(driver)
                            Data = Product_automation.test_product()
                        case "Design":
                            print("yes")
                            Design_automation =Design(driver)
                            Data = Design_automation.test_design()
                        case "SubDesign":
                            print("yes")
                            Subdesign_automation =Subdesign(driver)
                            Data = Subdesign_automation.test_subdesign()  
                        case "Designmapping":
                            print("yes")
                            designmap_automation =Designmapping(driver)
                            Data = designmap_automation.test_designmapping()      
                        case "Subdesignmapping":
                            print("yes")
                            Subdesignmap_automation =Subdesignmapping(driver)
                            Data = Subdesignmap_automation.test_subdesignmapping()
                        case "MC&VA":
                            print("yes")
                            McVa_automation =McVa(driver)
                            Data = McVa_automation.test_mc_va()    
                        case "Lot":                            
                            print("yes")
                            lot_automation = Lot(driver)
                            Data = lot_automation.test_lot()
                        case "Tag":
                            print("yes")
                            Tag_automation = Tag(driver)
                            Data = Tag_automation.test_tag()
                        case "Vendor":
                            print("yes")
                            Vendor_automation = VendorRegistration(driver) 
                            Data = Vendor_automation.test_vendor_registration()
                        case "Customer":
                            print("yes")
                            Vendor_automation = CustomerOrderTR(driver)
                            Data = Vendor_automation.test_customer_order_t_r()  
                        case "EST":
                            print("yes")
                            EST = ESTIMATION(driver)
                            Data = EST.test_estimation() 
                else:
                    print("Invalid option")
        finally:
            # Close the WebDriver
            driver.close()
            driver.quit()
            ct2 = datetime.datetime.now()
            print('Automation process completed',ct2)
            time_diff = ct2 - ct1  
            print("ct1 =", ct1.strftime("%Y-%m-%d %H:%M:%S.%f"))  # Format output
            print("ct2 =", ct2.strftime("%Y-%m-%d %H:%M:%S.%f"))  
            print("Time difference =", time_diff) 

    if __name__ == "__main__":
        print(__name__)
        main()
        
        




