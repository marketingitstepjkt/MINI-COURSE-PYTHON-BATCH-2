import time

from openpyxl import load_workbook
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

file_excel = load_workbook(filename="Day 2/data.xlsx")

sheet_range = file_excel["Sheet1"]

# Setup webdriver
driver = webdriver.Chrome()
driver.get("https://demoqa.com/webtables")
driver.maximize_window()
driver.implicitly_wait(10)

# Start looping kita
index = 2

while index <= len(sheet_range["A"]):
    first_name = sheet_range["A" + str(index)].value
    last_name = sheet_range["B" + str(index)].value
    age = sheet_range["C" + str(index)].value
    email = sheet_range["D" + str(index)].value
    salary = sheet_range["E" + str(index)].value
    department = sheet_range["F" + str(index)].value

    # Handle add button
    add_button = driver.find_element(By.ID, "addNewRecordButton")
    add_button.click()

    # Check condition using try
    try:
        #Tunggu modal muncul
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.ID, "registration-form-modal")))
 
        driver.find_element(By.ID, "firstName").send_keys(first_name)
        driver.find_element(By.ID, "lastName").send_keys(last_name)
        driver.find_element(By.ID, "userEmail").send_keys(email)
        driver.find_element(By.ID, "age").send_keys(age)
        driver.find_element(By.ID, "salary").send_keys(salary)
        driver.find_element(By.ID, "department").send_keys(department)
        driver.find_element(By.ID, "submit").click()

    except TimeoutException:
        print("Website error") 
        pass
    
    time.sleep(1)
    print(f"Data index ke - {index} dengan nama {first_name} sudah terinput")
    index = index + 1

print("Data sudah terinput semua, terimakasih!")

    
