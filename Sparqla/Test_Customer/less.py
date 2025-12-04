from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from time import sleep
import unittest, math
from Utils.Excel import ExcelUtils

FILE_PATH = ExcelUtils.file_path


class Stone(unittest.TestCase):
    def __init__(self, driver):
        super().__init__()  # ensure unittest init is called
        self.driver = driver
        self.wait = WebDriverWait(driver, 10)

    # ---------- Main Test ----------
    def test_tagStone(self, sheet_name, test_case_id):
        wait = self.wait
        sleep(2)
        # Excel setup
        count = ExcelUtils.Test_case_id_count(FILE_PATH, sheet_name, test_case_id)
        valid_rows = ExcelUtils.get_valid_rows(FILE_PATH, sheet_name)
        workbook = load_workbook(FILE_PATH)
        sheet = workbook[sheet_name]
        row = 1
        for row_num in range(2, valid_rows + 1):  # include last row
            if sheet.cell(row=row_num, column=1).value != test_case_id:
                # continue
                # Map Excel row to dict
                data_map = {
                    "Test Case Id": 1, "Less Weight": 2, "Type": 3, "Name": 4,
                    "Code": 5, "Pcs": 6, "Wt": 7, "Wt Type": 8,
                    "Cal.Type": 9, "Rate": 10, "Amount": 11
                }
                row_Stonedata = {
                    k: sheet.cell(row=row_num, column=v).value for k, v in data_map.items()
                }
                print("Row Data:", row_Stonedata)

                # --- Fill dropdowns ---
                self.select_dropdown(f"(//select[@name='est_stones_item[stones_type][]'])[{row}]", row_Stonedata["Type"])
                self.select_dropdown(f"(//select[@name='est_stones_item[stone_id][]'])[{row}]", row_Stonedata["Name"])

                # --- Fill inputs ---
                self.fill_input(f"(//input[@name='est_stones_item[stone_pcs][]'])[{row}]", row_Stonedata["Pcs"])
                self.fill_input(f"(//input[@name='est_stones_item[stone_wt][]'])[{row}]",  row_Stonedata["Wt"])
                self.select_dropdown("(//select[@name='est_stones_item[stone_uom_id][]'])[{row}]", row_Stonedata["Wt Type"])

                # --- Cal.Type radio button ---
                stonerow = row - 1
                cal_type = 1 if row_Stonedata["Cal.Type"] == "Wt" else 2
                cal_xpath = f"(//input[@name='est_stones_item[cal_type][{stonerow}]' and @value='{cal_type}'])"
                wait.until(EC.element_to_be_clickable((By.XPATH, cal_xpath))).click()

                # --- Rate & amount validation ---
                self.fill_input(f"(//input[@name='est_stones_item[stone_rate][]'])[{row}]", row_Stonedata["Rate"], clear=False)
                amt_element = wait.until(EC.visibility_of_element_located(
                    (By.XPATH, f'(//input[@name="est_stones_item[stone_price][]"])[{row}]'))
                )
                table_amt = f"{float(amt_element.get_attribute('value')):.2f}"
                self.calc_and_validate(row_Stonedata, table_amt)

                # --- Add multiple rows if needed ---
                if count > 1:
                    wait.until(EC.element_to_be_clickable((By.ID, "create_stone_item_details"))).click()
                    row += 1
                    count -= 1
                else:
                    wt_total = self.test_tablevalue(self.driver.find_elements(By.NAME, "est_stones_item[stone_wt][]"))
                    print("Total Wt:", wt_total)
                    amt_total = self.test_tablevalue(self.driver.find_elements(By.NAME, "est_stones_item[stone_price][]"))
                    print("Total Amount:", amt_total)

            # Save stone details
            wait.until(EC.element_to_be_clickable((By.ID, "update_stone_details"))).click()
            print("✅ Less weight detail added successfully")
            return "Less weight detail added successfully"

    # ---------- Utility Methods ----------
    def select_dropdown(self, locator, row, value):
        """Click and select dropdown value by visible text"""
        if value is None:
            print(f"⚠️ Skipping dropdown {locator}, value is None")
            return
        element = self.wait.until(EC.element_to_be_clickable((By.XPATH, locator.format(row))))
        Select(element).select_by_visible_text(str(value))
        print(f"✅ Dropdown filled with {value}")

    def fill_input(self, locator, row, value, clear=True):
        """Fill input field"""
        if value is None:
            print(f"⚠️ Skipping input {locator}, value is None")
            return
        element = self.wait.until(EC.visibility_of_element_located((By.XPATH, locator.format(row))))
        element.click()
        if clear:
            element.clear()
        element.send_keys(str(value))
        print(f"✅ Input filled with {value}")

    def calc_and_validate(self, row_Stonedata, table_amt):
        """Validate stone amount calculation"""
        try:
            if row_Stonedata["Cal.Type"] == "Wt":
                total = float(row_Stonedata["Wt"]) * float(row_Stonedata["Rate"])
            else:
                total = float(row_Stonedata["Pcs"]) * float(row_Stonedata["Rate"])
            expected_amt = f"{math.ceil(total):.2f}"
            if expected_amt == table_amt:
                print("✅ Stone Rate Calculation correct")
            else:
                print(f"❌ Stone Rate Calculation incorrect → Expected {expected_amt}, Got {table_amt}")
        except Exception as e:
            print(f"⚠️ Calculation skipped due to missing/invalid data: {e}")

    def test_tablevalue(self, rows):
        """Get total from input fields"""
        river = self.driver
        value=[]
        sleep(4)
        for row in rows:
            val = row.get_attribute("value")  
            if val and val.strip():  # Ensure it's not empty
                value.append(float(val.strip()))
            else:
                print("No value found in input.")  
        print("Collected Values:", value)
        if value:
            total_value = round(sum(value), 3)
        else:
            total_value = 0.0
        return total_value