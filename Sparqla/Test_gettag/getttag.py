from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
import unittest
from selenium.webdriver.support.ui import Select



class GetTag(unittest.TestCase):
    def __init__(self,driver):
        self.driver =driver   
        self.wait = WebDriverWait(driver, 30)
            
    def test_gettag(self,count):
        wait = self.wait
        wait.until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT,"Toggle navigation"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH,"(.//*[normalize-space(text()) and normalize-space(.)='Stock Issue'])[1]/following::span[1]"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH,"(.//*[normalize-space(text()) and normalize-space(.)='Re-Order Items'])[1]/following::span[1]"))).click()
        wait.until(EC.element_to_be_clickable((By.ID,"select2-branch_select-container"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH,"//span/input"))).clear()
        wait.until(EC.element_to_be_clickable((By.XPATH,"//span/input"))).send_keys("HEAD OFFICE")
        wait.until(EC.element_to_be_clickable((By.XPATH,"//span/input"))).send_keys(Keys.ENTER)
        sleep(4)
        Metal=wait.until(EC.element_to_be_clickable((By.XPATH,'//input[@placeholder="Select Metal"]')))
        Metal.click()
        Metal.send_keys('GOLD')
        Metal.send_keys(Keys.ENTER)
        wait.until(EC.element_to_be_clickable((By.XPATH,"//button[@id='tag_design_search']/i"))).click()
        Select(wait.until(EC.presence_of_element_located((By.NAME, "tag_items_list_length")))).select_by_visible_text("50")

        sleep(8)
        allTags = []
        rows = count  
        print(type(rows))      
        i = 1
        while i <= rows :
            orderNo = wait.until(EC.element_to_be_clickable((By.XPATH,f"//table[@id='tag_items_list']/tbody/tr[{i}]/td[4]"))).text
            if orderNo== '-':
                tagNo = wait.until(EC.element_to_be_clickable((By.XPATH,f"//table[@id='tag_items_list']/tbody/tr[{i}]/td[7]"))).text.strip()
                allTags.append(tagNo)
                print("Found Tag No: " + str(tagNo))
            else:
                rows=rows+1        
            i=i+1    
        print(allTags)
        return(allTags)
    

if __name__ == "__main__":
    unittest.main()
