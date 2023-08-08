from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import os
import time

load_dotenv()

# excel_url = 'https://rmiteduau-my.sharepoint.com/:x:/r/personal/s3978081_rmit_edu_vn/_layouts/15/Doc.aspx?sourcedoc=%7B308AE816-D32B-4CDB-9965-BD97F6B2BADB%7D&file=Book.xlsx&action=default&mobileredirect=true'
excel_url = 'https://rmiteduau.sharepoint.com/:x:/r/sites/RMITFinTechClub2023/_layouts/15/doc2.aspx?sourcedoc=%7B7665dfe2-01ab-4a69-8a84-e9be1b81f67e%7D&action=editnew'
# excel_url = 'https://rbpc.rmit.edu.vn'

EMAIL = os.getenv('EMAIL')
PASSWORD = os.getenv('PASSWORD')

driver = webdriver.Chrome()
choose_Lam = False

while True:
    driver.get(excel_url)
    if choose_Lam:
        time.sleep(4)

        # email input
        email_field = driver.find_element(By.ID , 'i0116')
        email_field.send_keys(EMAIL)

        next_btn = driver.find_element(By.ID, 'idSIButton9')
        next_btn.send_keys(Keys.ENTER)

        time.sleep(3)

        # password input
        email_field = driver.find_element(By.ID , 'i0118')
        email_field.send_keys(PASSWORD)

        next_btn = driver.find_element(By.ID, 'idSIButton9')
        next_btn.send_keys(Keys.ENTER)

        time.sleep(2)

        auth_number = driver.find_element(By.ID, 'idRichContext_DisplaySign').get_attribute('textContent')
        print(f'Your Authentication number is: {auth_number}')

        try:
            next_btn = WebDriverWait(driver, 60).until(EC.element_to_be_clickable((By.ID, 'idSIButton9')))
        except TimeoutException:
            print('Could not hear anything from users')
            break
        next_btn.send_keys(Keys.ENTER)
        #idSIButton9

    time.sleep(60)
    # break
