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
import unittest

FILE_PATH = ExcelUtils.file_path 
class Billing(unittest.TestCase):
    def __init__(self, driver):
        self.driver = driver   
        self.wait = WebDriverWait(driver, 30)

    def test_Billing(self):
        """Main test entry point for Billing automation."""
        driver = self.driver
        wait = self.wait 
        
        rate = Boardrate.Todayrate(self)
        print(f"Today's Rate: {rate}")  
        
        wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, "Toggle navigation"))).click()
        Function_Call.click(self, "//span[contains(text(), 'Billing')]")
        Function_Call.click(self, "//span[contains(text(), 'New Bill')]")
        
        sheet_name = "Billing"                                        
        valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, sheet_name)
        print(f"Processing {valid_rows - 2} rows from {sheet_name}")

        workbook = load_workbook(FILE_PATH)
        sheet = workbook[sheet_name]

        for row_num in range(2, valid_rows):  
            data_map = {
                "Test Case Id": 1, "TestStatus": 2, "ActualStatus": 3, "Cost Centre": 4,
                "Billing To": 5, "Employee": 6, "Customer Number": 7, "Customer Name": 8,
                "Delivery Location": 9, "Bill Type": 10, "driect": 11, "EstNo": 12,
                "SGST": 13, "CGST": 14, "Total": 15, "Discount": 16,
                "Handling_Charges": 17, "Return_Charges": 18, "Is Credit": 19,
                "Is Tobe": 20, "Received": 21, "Credit Due Date": 22,
                "Gift Voucher": 23, "Cash": 24, "Creditcard": 25, "Cheque": 26,
                "NetBanking": 27, "BillNo": 28, "Keep_it_As": 29, "Store As": 30,
                "OrderNo": 31, "Amount": 32
            }
            row_data = {key: sheet.cell(row=row_num, column=col).value for key, col in data_map.items()}
            print(f"Starting test case: {row_data['Test Case Id']}")
            
            self.create(row_data, row_num, sheet_name, rate)

        workbook.close()

    def create(self, row_data, row_num, sheet_name, board_rate):
        """Processes a single billing record."""
        self.driver.refresh()
        
        # 1. Handle Header Information
        if not self._handle_bill_header(row_data, row_num, sheet_name):
            return

        # 2. Handle Bill Type Specifics
        if not self._handle_bill_type(row_data, row_num, sheet_name):
            return

        # 3. Handle Calculations and Summary
        if not self._handle_total_summary(row_data, row_num, sheet_name):
            return

        # 4. Process Payments
        self._handle_payments(row_data, row_num, sheet_name)

    def _handle_bill_header(self, row_data, row_num, sheet_name):
        """Fills Cost Centre, Bill To, Employee, and Customer details."""
        # Cost Centre
        if row_data.get('Cost Centre'):
            Function_Call.select_visible_text(self, '//select[@id="id_branch"]', 
                                          value=row_data['Cost Centre'])
        else:
            msg = "Cost Centre field is mandatory ⚠️"
            Function_Call.Remark(self, row_num, msg, sheet_name)
            return False

        # Billing To
        bill_to_map = {
            "Customer": '//input[@id="billing_for1"]',
            "Company": '//input[@id="billing_for2"]',
            "Supplier": '//input[@id="billing_for3"]'
        }
        target_bill_to = row_data.get("Billing To", "Customer")
        Function_Call.click(self, bill_to_map.get(target_bill_to, bill_to_map["Customer"]))

        # Employee
        if row_data.get("Employee"):
            Function_Call.dropdown_select(self, "//span[@id='select2-emp_select-container']", 
                                         row_data["Employee"], 
                                         '//span[@class="select2-search select2-search--dropdown"]/input')
        else:
            msg = "Employee field is mandatory ⚠️"
            Function_Call.Remark(self, row_num, msg, sheet_name)
            return False

        # Customer
        if row_data.get("Customer Number"):
            Function_Call.fill_autocomplete_field(self, "bill_cus_name", str(row_data["Customer Number"]))
        else:
            msg = "Customer field is mandatory ⚠️"
            Function_Call.Remark(self, row_num, msg, sheet_name)
            return False

        Function_Call.click2(self, '(//button[@class="btn btn-close btn-warning"])[11]')
        return True

    def _handle_bill_type(self, row_data, row_num, sheet_name):
        """Selects and configures the bill type."""
        bill_type = row_data.get("Bill Type")
        
        match bill_type:
            case "SALES":
                Function_Call.click(self, '//input[@id="bill_typesales"]')
            case "SALES & PURCHASE":
                Function_Call.click(self, '//input[@id="bill_type_salesPurch"]')
            case "SALES PURCHASE & RETURN":
                Function_Call.click(self, '//input[@id="bill_type_saleRet"]')
                return self._select_last_filter_bill(row_num, sheet_name)
            case "PURCHASE":
                Function_Call.click(self, '//input[@id="bill_type_purchase"]')
            case "ORDER ADVANCE":
                Function_Call.click(self, '//input[@id="bill_type_order_advance"]')
                Function_Call.fill_input2(self, '//input[@id="filter_order_no"]', row_data.get("OrderNo"))
                Function_Call.click(self, '//button[@id="search_order_no"]')
                Function_Call.click2(self, '(//button[@class="btn btn-close btn-warning"])[11]')
                Function_Call.fill_input2(self, '//input[@name="billing[bill_amount]"]', row_data.get("Amount"))
                Function_Call.click(self, '//li[@id="tab_make_pay"]')
            case "SALES RETURN":
                Function_Call.click(self, '//input[@id="bill_type_sales_return"]')
                return self._select_last_filter_bill(row_num, sheet_name)
            case "CREDIT COLLECTION":
                Function_Call.click(self, '//input[@id="bill_type_credit_bill"]')
            case "ORDER DELIVERY":
                Function_Call.click(self, '//input[@id="bill_type_order_del"]')
            case "Repair Order Delivery":
                Function_Call.click(self, '//input[@id="repair_order_delivery"]')
        return True

    def _select_last_filter_bill(self, row_num, sheet_name):
        """Helper to select the last bill in the filter dropdown."""
        try:
            sleep(5)
            element = self.wait.until(EC.element_to_be_clickable((By.ID, "filter_bill_no")))
            dropdown = Select(element)
            options_count = len(dropdown.options)
            
            if options_count > 1:
                dropdown.select_by_index(options_count - 1)
                Function_Call.click(self, '//button[@id="search_bill_no"]')
                Function_Call.click(self, '//input[@class="select_est_details"]')
                Function_Call.click(self, '//a[@id="update_bill_return"]')
                return True
            else:
                msg = "No bills available to return ⚠️"
                print(msg)
                Function_Call.Remark(self, row_num, msg, sheet_name)
                return False
        except Exception as e:
            print(f"Error in bill selection: {e}")
            return False

    def _handle_total_summary(self, row_data, row_num, sheet_name):
        """Handles estimations, calculations, and charges in the summary tab."""
        if row_data.get("driect") == 'No' and row_data.get("Bill Type") != "ORDER ADVANCE":
            Function_Call.fill_input2(self, '//input[@id="filter_est_no"]', row_data["EstNo"])
            Function_Call.click(self, '//button[@id="search_est_no"]')
            sleep(2)
            Function_Call.click(self, '(//button[@class="btn btn-close btn-warning"])[11]')
            sleep(3)
            Function_Call.click(self, "//a[normalize-space()='Total Summary']")
            
            # Fetch summary details
            summary_labels = {
                "Taxable Sale Amount": "//span[@class='summary_lbl summary_sale_amt']",
                "CGST": '//span[@class="summary_lbl sales_cgst"]',
                "SGST": '//span[@class="summary_lbl sales_sgst"]',
                "IGST": '//span[@class="summary_lbl sales_igst"]',
                "Sale Amount": '//span[@class="summary_lbl sale_amt_with_tax"]',
                "Purchase Amount": '//span[@class="summary_lbl summary_pur_amt"]'
            }
            for label, xpath in summary_labels.items():
                val = Function_Call.get_text(self, xpath)
                print(f"{label}: {val}")

            if Function_Call.get_text(self, summary_labels["Purchase Amount"]) == '0.00':
                self._fill_summary_charges(row_data, row_num, sheet_name)
            
            # Navigate to the next tab (Make Payment)
            try:
                Function_Call.click(self, '(//button[@class="btn btn-warning next-tab"])[2]')
            except:
                Function_Call.click(self, '//li[@id="tab_make_pay"]')
        return True

    def _fill_summary_charges(self, row_data, row_num, sheet_name):
        """Fills discount, handling, and return charges."""
        charge_map = {
            "Discount": ('//input[@id="summary_discount_amt"]', r"^(\d{1,7}(\.\d{1,2})?)?$"),
            "Handling_Charges": ('//input[@id="handling_charges"]', r"^(\d{1,3}(\.\d{1,2})?)?$"),
            "Return_Charges": ('//input[@id="return_charges"]', r"^(\d{1,3}(\.\d{1,2})?)?$")
        }
        for field, (xpath, pattern) in charge_map.items():
            if row_data.get(field):
                Function_Call.fill_input(
                    self, self.wait, locator=(By.XPATH, xpath),
                    value=row_data[field], pattern=pattern, field_name=field,
                    screenshot_prefix=field, row_num=row_num, Sheet_name=sheet_name
                )

    def _handle_payments(self, row_data, row_num, sheet_name):
        """Manages received amount, credit settings, and specific payment methods."""
        # Ensure we are on the Make Payment tab
        try:
            Function_Call.click(self, '//li[@id="tab_make_pay"]')
        except:
            pass

        received_str = Function_Call.get_value(self, '//input[@name="billing[tot_amt_received]"]')
        received_value = float(received_str or 0)
        cash_val = float(row_data.get("Cash") or 0)
        final_received = received_value

        if row_data.get('Received'):
            percentage = float(row_data['Received'])
            credit_val = (received_value * percentage) / 100
            final_received = credit_val
            
            Function_Call.click(self, '//input[@id="is_credit_yes"]')
            if row_data.get('Is Credit') == 'Yes':
                Function_Call.fill_input(
                    self, self.wait, locator=(By.XPATH, '//input[@name="billing[tot_amt_received]"]'),
                    value=credit_val, pattern=r"^(\d{1,7}(\.\d{1,2})?)?$",
                    field_name="Received", screenshot_prefix="Received",
                    row_num=row_num, Sheet_name=sheet_name
                )
            
            if row_data.get('Is Tobe') == 'Yes':
                Function_Call.click(self, '//input[@id="is_to_be_yes"]')
            
            if row_data.get("Credit Due Date"):
                Function_Call.fill_input(
                    self, self.wait, locator=(By.XPATH, '//input[@id="credit_due_date"]'),
                    value=row_data["Credit Due Date"],
                    pattern=r"^(0[1-9]|[12][0-9]|3[01])-(0[1-9]|1[0-2])-\d{4}$",
                    field_name="Credit_Due_Date", screenshot_prefix="Credit_Due_Date",
                    row_num=row_num, Sheet_name=sheet_name, extra_keys=Keys.TAB, Date_range="Yes"
                )

        # Determine adjustable amount
        adjustable_amount = final_received
        if received_value == 0:
            pay_to_cus = float(Function_Call.get_value(self, '//input[@name="billing[pay_to_cus]"]') or 0)
            adjustable_amount = pay_to_cus

        pending_payment = adjustable_amount - cash_val
        
        # Process Cash
        if cash_val > 0:
            Function_Call.fill_input(
                self, self.wait, locator=(By.XPATH, '//input[@id="make_pay_cash"]'),
                value=int(cash_val), pattern=r"^(?:[0-9]{1,5}|1[0-9]{5})$", field_name="Cash",
                screenshot_prefix="Cash", row_num=row_num, Sheet_name=sheet_name
            )

        # Process Other Methods
        test_case_id = row_data.get('Test Case Id')
        payment_methods = {
            'Creditcard': CreditCard.test_Credit_Card,
            'Cheque': Cheque.test_Cheque,
            'NetBanking': NetBanking.test_NetBanking
        }
        for key, method in payment_methods.items():
            if row_data.get(key) == 'Yes':
                method(self, test_case_id, pending_payment)
        
        balance=Function_Call.get_text(self, '//table[@id="payment_modes"]//tfoot//tr[2]//th[3]')
        print(f"Balance: {balance}")
        