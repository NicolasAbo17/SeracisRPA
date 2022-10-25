#Importar librerias

import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
import warnings
import time
import pandas as pd
import os
import shutil
import datetime
import glob
from datetime import datetime
warnings.filterwarnings("ignore")

#Settings de Chrome Selenium

options = Options()
#options.add_argument("--headless")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")
options.add_experimental_option("excludeSwitches", ["enable-automation"])
options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome_prefs = {"download.default_directory": r"C:\Users\nicol\Downloads"}
options.experimental_options["prefs"] = chrome_prefs

# Descarga del archivo para el personal operativo acreditado:

driver = webdriver.Chrome(executable_path=r'C:\Users\nicol\Downloads\chromedriver_win32\chromedriver.exe',chrome_options=options)
form_url = "https://apo.supervigilancia.gov.co/AcreditaPO/BuscaEmprapo.aspx"
driver.get(form_url)

a = driver.find_element(By.ID, 'ctl00_contentMaster_TxNit')
a.send_keys("811007280")

b = driver.find_element(By.ID, 'ctl00_contentMaster_BtBuscar')
b.click()

c = driver.find_element(By.TAG_NAME,'I')
c.click()

time.sleep(5)
driver.close()
driver.quit()

#Lectura del excel y escritura en BD para el personal operativo acreditado:

filepath = r"C:\Users\nicol\Downloads"
old_file = max([filepath +'\\'+ f for f in os.listdir(filepath)], key=os.path.getctime)
fec_hoy = datetime.now()
fec_hoy = fec_hoy.strftime('%d_%m_%Y_%H_%M')
new_file = r"C:\Users\nicol\Downloads\Acreditados" + fec_hoy + ".xlsx"
os.rename(old_file, new_file)

acre = pd.read_excel(new_file,sheet_name='Sheet1', engine='openpyxl')

acre.columns = acre.iloc[0]
acre = acre.reindex(acre.index.drop(0)).reset_index(drop=True)
acre.columns.name = None

print(acre.head())

# Ingreso a semántica y escritura ciclica acorde al registro en los DF:

#for i in range(0,len(acre.index)):
for i in range(0,3):

    var_1 = str(acre.loc[i,['IdNum']].item())
    var_2 = str(acre.loc[i,['Vigen.Acr']].item())
    var1 = "L.BOTERO"

    driver = webdriver.Chrome(executable_path=r'C:\Users\nicol\Downloads\chromedriver_win32\chromedriver.exe',chrome_options=options)
    form_url = "https://corporativo.seracis.com/seracis/public/index.php/recursohumano/movimiento/recurso/acreditacion/lista"
    driver.get(form_url)
    
    log_sem_1 = driver.find_element(By.NAME, '_username')
    log_sem_1.send_keys(var1)

    log_sem_2 = driver.find_element(By.NAME, '_password')
    log_sem_2.send_keys(var1)

    btn_log = driver.find_element(By.ID, 'js-login-btn')
    btn_log.click()

    empl_bus = driver.find_element(By.ID, 'form_numeroIdentificacion')
    empl_bus.send_keys(var_1)

    empl_filtro = driver.find_element(By.ID, 'form_btnFiltro')
    empl_filtro.click()

    driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

    time.sleep(2)
    driver.close()
    driver.quit()

#Descarga del archivo para el personal operativo en trámite:

driver = webdriver.Chrome(executable_path=r'C:\Users\nicol\Downloads\chromedriver_win32\chromedriver.exe',chrome_options=options)
form_url = "https://apo.supervigilancia.gov.co/Acreditapo/BuscaEnTram.aspx"
driver.get(form_url)

a = driver.find_element(By.ID, 'ctl00_contentMaster_TxNit')
a.send_keys("811007280")

b = driver.find_element(By.ID, 'ctl00_contentMaster_BtBuscar')
b.click()

c = driver.find_element(By.TAG_NAME,'I')
c.click()

time.sleep(5)
driver.close()
driver.quit()

#Lectura del excel y escritura en BD para el personal pendiente de acreditación:

filepath_2 = glob.glob(r"C:\Users\nicol\Downloads\*.xlsx")
old_file_2 = max(filepath_2, key=os.path.getctime)
fec_hoy_2 = datetime.now()
fec_hoy_2 = fec_hoy_2.strftime('%d_%m_%Y_%H_%M')
new_file_2 = r"C:\Users\nicol\Downloads\Pendientes" + fec_hoy + ".xlsx"
os.rename(old_file, new_file_2)

pen = pd.read_excel(new_file_2,sheet_name='Sheet1', engine='openpyxl')

pen.columns = pen.iloc[0]
pen = pen.reindex(pen.index.drop(0)).reset_index(drop=True)
pen.columns.name = None
print(pen.head())

# Ingreso a semántica y escritura ciclica acorde al registro en los DF:

# for i in range(0,len(pen.index)):

#     var_1 = str(acre.loc[i,['IdNum']].item())
#     var_2 = str(acre.loc[i,['Estado']].item())
#     var1 = "L.BOTERO"

#     driver = webdriver.Chrome(executable_path=r'C:\Users\nicol\Downloads\chromedriver_win32\chromedriver.exe',chrome_options=options)
#     form_url = "https://corporativo.seracis.com/seracis/public/index.php/recursohumano/movimiento/recurso/acreditacion/lista"
#     driver.get(form_url)
    
#     log_sem_1 = driver.find_element(By.NAME, '_username')
#     log_sem_1.send_keys(var1)

#     log_sem_2 = driver.find_element(By.NAME, '_password')
#     log_sem_2.send_keys(var1)

#     btn_log = driver.find_element(By.ID, 'js-login-btn')
#     btn_log.click()

#     empl_bus = driver.find_element(By.ID, 'form_numeroIdentificacion')
#     empl_bus.send_keys(var_1)

#     empl_filtro = driver.find_element(By.ID, 'form_btnFiltro')
#     empl_filtro.click()

#     driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

#     time.sleep(2)
#     driver.close()
#     driver.quit()