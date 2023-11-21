import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import ElementClickInterceptedException

driver = None

try:
    driver = webdriver.Edge()
except WebDriverException as e:
    try:
        driver = webdriver.Chrome()
    except WebDriverException as e:
        try:
            driver = webdriver.Firefox()
        except WebDriverException as e:
            print(f"WebDriver creation failed: {e}")
        
if driver is not None:
    driver.get('https://qcpi.questcdn.com/cdn/posting/?group=1950787&provider=1950787')
    wait = WebDriverWait(driver, 20)

    results_df = pd.DataFrame(columns=['Closing Date', 'Est. Value Notes', 'Description'])

    try:
        time.sleep(1)
        open_detail = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//a[@href="#"]')))[14]
        open_detail.click()
    except ElementClickInterceptedException as e:
        print(f"Element click intercepted: {e}")

    for page in range(5):
        time.sleep(1)
        location_detail_table = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//table[starts-with(@class,"table")]')))[3]
        project_description_table = wait.until(EC.presence_of_all_elements_located((By.XPATH, '//table[starts-with(@class,"table")]')))[4]

        closing_date = location_detail_table.find_elements(By.TAG_NAME, 'td')[1].text
        est_value_notes = location_detail_table.find_elements(By.TAG_NAME, 'td')[7].text
        description = project_description_table.find_elements(By.TAG_NAME, 'td')[7].text

        print('Closing Date: {}\nEst. Value Notes: {}\nDescription: {}'.format(closing_date, est_value_notes, description))
        print('--------------------------------------')

        results_df = pd.concat([results_df, pd.DataFrame({'Closing Date': [closing_date], 'Est. Value Notes': [est_value_notes], 'Description': [description]})], ignore_index=True)
        wait.until(EC.element_to_be_clickable((By.XPATH, '//button[@id="id_prevnext_next"]'))).click()

    results_df.to_excel('output.xlsx', index=False)
    driver.quit()
else:
    print('None of the specificed browers are present on the system.')
