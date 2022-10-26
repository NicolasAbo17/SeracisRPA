# Librerias
from trace import Trace
from xmlrpc.client import boolean
import selenium
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import warnings
import time
import pandas as pd
import openpyxl
import xlrd
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
wait = WebDriverWait(driver, 10)

# Paths
directorioDes = r"C:\Users\nicol\Downloads"

acreditadosUrl = "https://apo.supervigilancia.gov.co/AcreditaPO/BuscaEmprapo.aspx"
acreditadosArchivo = "Informacion de Companias.xlsx"
seracisNit = "811007280"

enprocesoUrl = "https://apo.supervigilancia.gov.co/Acreditapo/BuscaEnTram.aspx"
enprocesoArchivo = "Informacion de Companias (1).xlsx"

empleadosUrl = "http://192.168.2.95/prueba/public/index.php/recursohumano/administracion/recurso/empleado/lista"
subirAcreditacionUrl = "http://192.168.2.95/prueba/public/index.php/recursohumano/movimiento/recurso/acreditacion/lista"
usuarioSemantica = "L.BOTERO"

empleadosArchivo = "empleados.xlsx"
empleadosColumnas = ['NOMBRE', 'IDENTIFICACIÓN', 'FECHA_DESDE', 'CARGO', 'ZONA', 'SUBZONA']
empleadosCargos = ['ESCOLTA', 'MANEJADOR CANINO', 'OPERADOR (A) MEDIOS TECNOLOGICOS', 'SUPERVISOR', 'SUPERVISOR CANINO', 'SUPERVISOR CENTRAL DE MONITOREO', 'SUPERVISOR DE CONTROL DE COMUNICACIONES', 'SUPERVISOR DE PORTERIA', 'SUPERVISOR DE ZONA', 'VIGILANTE']

global loginBool
loginBool = False

# Functions
def descargarSupervigilancia(url): 
    global loginBool
    loginBool = False

    driver.get(url)

    a = driver.find_element(By.ID, 'ctl00_contentMaster_TxNit')
    a.send_keys(seracisNit)

    b = driver.find_element(By.ID, 'ctl00_contentMaster_BtBuscar')
    b.click()

    c = driver.find_element(By.TAG_NAME,'I')
    c.click()

    time.sleep(5)

def leerUltimo(skipRows, delete, xls):
    old_file = max([directorioDes +'\\'+ f for f in os.listdir(directorioDes)], key=os.path.getctime)
    if xls:
        workbook = xlrd.open_workbook_xls(old_file, ignore_workbook_corruption=True)  
        df = pd.read_excel(workbook)
    elif skipRows:
        df = pd.read_excel(old_file, skiprows=[0])
    else:
        df = pd.read_excel(old_file)
    if delete:
        os.remove(old_file)
    return df

def leerSupervigilancia():
    return leerUltimo(True, False, False)

def loginSemantica():
    global loginBool
    if loginBool == False:
        loginBool = True
        log_sem_1 = driver.find_element(By.NAME, '_username')
        log_sem_1.send_keys(usuarioSemantica)

        log_sem_2 = driver.find_element(By.NAME, '_password')
        log_sem_2.send_keys(usuarioSemantica)

        btn_log = driver.find_element(By.ID, 'js-login-btn')
        btn_log.click()

def subirListaAcreditados():
    driver.get(subirAcreditacionUrl)
    loginSemantica()

    btn_val = driver.find_element(By.LINK_TEXT, 'Cargar validación')
    cargarArchivoNuevaVentana(btn_val, directorioDes + "/" + enprocesoArchivo, True)

    btn_acr = driver.find_element(By.LINK_TEXT, 'Cargar acreditación')
    cargarArchivoNuevaVentana(btn_acr, directorioDes + "/" + acreditadosArchivo, False)

def cargarArchivoNuevaVentana(toClick, filepath, writeNum):
    original_window = driver.current_window_handle
    toClick.click()

    wait.until(EC.number_of_windows_to_be(2))
    for window_handle in driver.window_handles:
        if window_handle != original_window:
            driver.switch_to.window(window_handle)
            break
    
    if writeNum: 
        txt_numero = driver.find_element(By.ID, 'form_numero')
        txt_numero.send_keys('10')
    else:
        cbx_validacion = driver.find_element(By.ID, 'form_omitirValidacion')
        cbx_validacion.click()

    btn_browse = driver.find_element(By.ID, 'form_attachment')
    btn_browse.send_keys(filepath)

    btn_submit = driver.find_element(By.ID, 'form_btnCargar')
    btn_submit.click()

    wait.until(EC.number_of_windows_to_be(1))
    driver.switch_to.window(original_window)
    os.remove(filepath)

def descargarEmpleadosSemantica():
    driver.get(empleadosUrl)
    loginSemantica()

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
    time.sleep(8)

def leerEmpleados():
    df = leerUltimo(False, True, False)
    df.drop(df.columns.difference(empleadosColumnas), 1, inplace=True)
    df = df[df['CARGO'].isin(empleadosCargos)]
    df.loc[df['CARGO'] == empleadosCargos[2], 'CARGO'] = 'OPERADOR DE MEDIOS TECNOLOGICOS'
    df['FECHA.ACR'] = 'Na'
    return df.sort_values(by=['IDENTIFICACIÓN', 'CARGO'], ascending=[True, True])

def aniadirFecha(df1, df2, col):
    i = 0
    j = 0
    while i < df1.shape[0] and j < df2.shape[0]:
        if df1.loc[df1.index[i],'IDENTIFICACIÓN'] == df2.loc[df2.index[j],'IdNum']:
            if df1.loc[df1.index[i],'CARGO'].strip() == df2.loc[df2.index[j],'Cargo'].strip():
                df1.loc[df1.index[i],'FECHA.ACR'] = df2.loc[df2.index[j],col]
                i += 1
            j += 1
        else:
            if(df1.iloc[i]['IDENTIFICACIÓN'] < df2.iloc[j]['IdNum']):
                i+=1
            else:
                j+=1
    return df1

def descargarApos():
    driver.get(subirAcreditacionUrl)
    loginSemantica()

    informe_completo = pd.DataFrame()
    # for i in range(0,len(empleados.index)):
    for i in range(0,len(empleados.index)):
        iden = str(empleados.loc[i,['IDENTIFICACIÓN']].item())
        
        empl_bus = driver.find_element(By.ID, 'form_numeroIdentificacion')
        empl_bus.clear()
        empl_bus.send_keys(iden)

        empl_filtro = driver.find_element(By.ID, 'form_btnFiltro')
        empl_filtro.click()
        time.sleep(2)

        btn_archivos = driver.find_element(By.ID, 'archivos')
        btn_archivos.click()
        btn_informe = driver.find_element(By.ID, 'form_btnInformeApo')
        btn_informe.click()
        time.sleep(2)

        filename = max([directorioDes +'\\'+ f for f in os.listdir(directorioDes)], key=os.path.getctime)
        print(filename)

        informe_apo = leerUltimo(False, True, True)
        informe_completo = pd.concat([informe_completo, informe_apo], ignore_index=True)
        
    print(informe_completo.head())

# # Process
# # Descarga de supervigilancia las personas en proceso de acreditación y ya acreditadas
# descargarSupervigilancia(acreditadosUrl)
# acreditados = leerSupervigilancia()
# acreditados = acreditados.sort_values(by=['IdNum', 'Cargo'], ascending=[True, True])

# descargarSupervigilancia(enprocesoUrl)
# enproceso = leerSupervigilancia()
# enproceso = enproceso.sort_values(by=['IdNum', 'Cargo'], ascending=[True, True])

# subirListaAcreditados()

# # Descarga los empleados como están registrados actualmente en semantica
# descargarEmpleadosSemantica()
# empleados = leerEmpleados()

# # Añade la fecha de acreditación a quienes la tengan
# empleados = aniadirFecha(empleados, enproceso, 'Estado')
# empleados = aniadirFecha(empleados, acreditados, 'Vigen.Acr')

# empleados = empleados[empleados['FECHA.ACR'] == 'Na']
# fec_hoy = datetime.now()
# fec_hoy = fec_hoy.strftime('%d_%m_%Y_%H_%M')
# new_file = directorioDes + "\Acreditados" + fec_hoy + ".xlsx"
# empleados.to_excel(new_file, index=False, header=True)

empleados = pd.read_excel(directorioDes + "\Acreditados26_10_2022_14_30.xlsx")
descargarApos()

driver.quit()