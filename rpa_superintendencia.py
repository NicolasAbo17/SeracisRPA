# Librerias
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
import warnings
import time
import pandas as pd
import openpyxl
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

driver = webdriver.Chrome(executable_path=r'C:\Users\nicol\Downloads\chromedriver_win32\chromedriver.exe',chrome_options=options)

# Paths
directorioDes = r"C:\Users\nicol\Downloads"

empleadosUrl = "http://192.168.2.95/prueba/public/index.php/recursohumano/administracion/recurso/empleado/lista"
usuarioSemantica = "L.BOTERO"
empleadosArchivo = "empleados.xlsx"
empleadosColumnas = ['NOMBRE', 'IDENTIFICACIÓN', 'FECHA_DESDE', 'CARGO', 'ZONA', 'SUBZONA']
empleadosCargos = ['ESCOLTA', 'MANEJADOR CANINO', 'OPERADOR (A) MEDIOS TECNOLOGICOS', 'SUPERVISOR', 'SUPERVISOR CANINO', 'SUPERVISOR CENTRAL DE MONITOREO', 'SUPERVISOR DE CONTROL DE COMUNICACIONES', 'SUPERVISOR DE PORTERIA', 'SUPERVISOR DE ZONA', 'VIGILANTE']

acreditadosUrl = "https://apo.supervigilancia.gov.co/AcreditaPO/BuscaEmprapo.aspx"
acreditadosArchivo = "Informacion de Companias.xlsx"
seracisNit = "811007280"

enprocesoUrl = "https://apo.supervigilancia.gov.co/Acreditapo/BuscaEnTram.aspx"

# Functions

def descargarEmpleadosSemantica():
    driver.get(empleadosUrl)
    
    # Login
    log_sem_1 = driver.find_element(By.NAME, '_username')
    log_sem_1.send_keys(usuarioSemantica)

    log_sem_2 = driver.find_element(By.NAME, '_password')
    log_sem_2.send_keys(usuarioSemantica)

    btn_log = driver.find_element(By.ID, 'js-login-btn')
    btn_log.click()

    # Filtro
    fil_limite = driver.find_element(By.ID, 'form_limiteRegistros')
    fil_limite.send_keys('00')

    estado_sel = driver.find_element(By.ID, 'form_estadoContrato')
    estado_sel.send_keys('S')

    btn_fil = driver.find_element(By.ID, 'form_btnFiltro')
    btn_fil.click()
    time.sleep(2)

    # Descarga
    btn_down = driver.find_element(By.ID, 'form_btnExcel')
    btn_down.click()
    time.sleep(10)

def leerEmpleados():
    empleados = pd.read_excel(directorioDes + "/" + empleadosArchivo)
    empleados.drop(empleados.columns.difference(empleadosColumnas), 1, inplace=True)
    empleados = empleados[empleados['CARGO'].isin(empleadosCargos)]
    return empleados

def descargarSupervigilancia(): 
    driver.get(acreditadosUrl)

    a = driver.find_element(By.ID, 'ctl00_contentMaster_TxNit')
    a.send_keys(seracisNit)

    b = driver.find_element(By.ID, 'ctl00_contentMaster_BtBuscar')
    b.click()

    c = driver.find_element(By.TAG_NAME,'I')
    c.click()

    time.sleep(5)
    driver.close()

def leerAcreditados():
    return pd.read_excel(directorioDes + "/" + acreditadosArchivo, skiprows=[0])


# Process
# descargarEmpleadosSemantica()
# empleados = leerEmpleados()
# empleados.loc[empleados['CARGO'] == empleadosCargos[2], 'CARGO'] = 'OPERADOR DE MEDIOS TECNOLOGICOS'
# empleados['FECHA.ACR'] = ''
# empleados = empleados.sort_values(by=['IDENTIFICACIÓN', 'CARGO'], ascending=[True, True])

# descargarSupervigilancia()
# acreditados = leerAcreditados()
# acreditados = acreditados.sort_values(by=['IdNum', 'Cargo'], ascending=[True, True])

# i = 0
# j = 0
# while i < empleados.shape[0] and j < acreditados.shape[0]:
#     if empleados.loc[empleados.index[i],'IDENTIFICACIÓN'] == acreditados.loc[acreditados.index[j],'IdNum']:
#         if empleados.loc[empleados.index[i],'CARGO'].strip() == acreditados.loc[acreditados.index[j],'Cargo'].strip():
#             print(empleados.loc[empleados.index[i],'IDENTIFICACIÓN'])
#             empleados.loc[empleados.index[i],'FECHA.ACR'] = acreditados.loc[acreditados.index[j],'Vigen.Acr']
#             print(empleados.loc[empleados.index[i],'FECHA.ACR'])
#             i += 1
#         j += 1
#     else:
#         if(empleados.iloc[i]['IDENTIFICACIÓN'] < acreditados.iloc[j]['IdNum']):
#             i+=1
#         else:
#             j+=1

# fec_hoy = datetime.now()
# fec_hoy = fec_hoy.strftime('%d_%m_%Y_%H_%M')
# new_file = directorioDes + "\Acreditados" + fec_hoy + ".xlsx"
# empleados.to_excel(new_file, index=False, header=True)

empleados = pd.read_excel("Acreditados25_10_2022_13_34.xlsx")

driver.quit()