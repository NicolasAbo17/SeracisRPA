# Librerias selenium xlwt pandas xlrd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains

import warnings
import time
from pathlib import Path

warnings.filterwarnings("ignore")

#Settings de Chrome Selenium

options = Options()
#options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--user-data-dir=" + str(Path.home()) + "/AppData/Local/Google/Chrome/User Data/Default")

options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('excludeSwitches', ['enable-logging'])

driver = webdriver.Chrome(executable_path=r'./chromedriver.exe',chrome_options=options)
wait = WebDriverWait(driver, 10)
action = ActionChains(driver)

powerbi_link = "https://app.powerbi.com/groups/467151f2-719f-41c7-b407-786c2143001e/list"

nombre_bd = " Inconsistencias Programación_V2.0 "

driver.get(powerbi_link)
time.sleep(4)

try:
    optionButton = driver.find_element(By.XPATH,"//button[@class = 'mat-tab-link mat-focus-indicator'][2]")
except:
    input("Press enter when you login...")
    optionButton = driver.find_element(By.XPATH,"//button[@class = 'mat-tab-link mat-focus-indicator'][2]")

optionButton.click()
time.sleep(2)

while True:
    actionChain = ActionChains(driver)
    row_database = driver.find_element(By.XPATH,"//a[text()='"+nombre_bd+"']/../..")
    actionChain.move_to_element(row_database).perform()

    time.sleep(3)
    refreshBtn = driver.find_element(By.XPATH,"//a[text()='"+nombre_bd+"']/../button")
    refreshBtn.click()
    
    time.sleep(1800)
