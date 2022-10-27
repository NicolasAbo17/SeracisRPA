# Librerias selenium xlwt pandas xlrd
from calendar import month
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
directorioApo = r"C:\Users\nicol\Downloads\APOS"

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

apoInicio = "APO8110072801"
retirosVacios = pd.DataFrame(columns=['Nit','RazonSocial','TipoDocumento','NoDocumento','FechaRetiro'])
registrosVacios = pd.DataFrame(columns=['Nit','RazonSocial','TipoDocumento','NoDocumento','FechaRetiro','Nombre1','Nombre2','Apellido1','Apellido2',
'FechaNacimiento','Genero','Cargo','FechaVinculacion','CodigoCurso','NitEscuela','Nro','TipoEstablecimiento','TelefonoR','DireccionR','DireccionP',
'departamento','Ciudad','EducacionBM','EducacionS','Discapacidad'])

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
            if df2.loc[df2.index[j],'Cargo'].strip() in df1.loc[df1.index[i],'CARGO'].strip():
                df1.loc[df1.index[i],'FECHA.ACR'] = df2.loc[df2.index[j],col]
                i += 1
            j += 1
        else:
            if(df1.iloc[i]['IDENTIFICACIÓN'] < df2.iloc[j]['IdNum']):
                i+=1
            else:
                j+=1
    return df1

def verificarNro(nro):
    if not nro.startswith('ECSP'):
        for chr in nro:
            if chr.isdigit():
                nro = 'ECSP' + nro[nro.index(chr):]
                break

    for chr in nro:
        if chr == '-':
            if nro[nro.index(chr)+1] == '1':
                nro = nro[:nro.index(chr)+1] + 'I' + nro[nro.index(chr)+2:]
            break

    return nro

def obtenerTextoFecha():
    todayDate = datetime.today()
    todayYear = str(todayDate.year)
    todayMonth = str(todayDate.month)
    todayDay = str(todayDate.day)
    month = todayDate.month
    day = todayDate.day

    if month < 10:
        todayMonth = "0" + str(month)
    if day < 10:
        todayDay = "0" + str(day)
    return todayYear + todayMonth + todayDay

def descargarApos():
    driver.get(subirAcreditacionUrl)
    loginSemantica()

    if not os.path.exists(directorioApo):
        os.makedirs(directorioApo)

    informe_completo = pd.DataFrame()
    faltan = pd.DataFrame()
    apo_num = 1

    # for i in range(0,len(empleados.index)):
    for i in range(0,len(empleados.index)):
        iden = str(empleados.loc[empleados.index[i],['IDENTIFICACIÓN']].item())
        cargo = str(empleados.loc[empleados.index[i],['CARGO']].item())
        
        empl_bus = driver.find_element(By.ID, 'form_numeroIdentificacion')
        empl_bus.clear()
        empl_bus.send_keys(iden)

        empl_filtro = driver.find_element(By.ID, 'form_btnFiltro')
        empl_filtro.click()
        time.sleep(1)

        btn_archivos = driver.find_element(By.ID, 'archivos')
        btn_archivos.click()
        btn_informe = driver.find_element(By.ID, 'form_btnInformeApo')
        btn_informe.click()
        
        index = -1
        rows = len(driver.find_elements(By.XPATH,"//table/tbody/tr"))

        for t_row in range(1, (rows + 1)):
            xPath = "//table/tbody/tr[" + str(t_row) + "]/td[" + str(7) + "]"
            t_cargo = driver.find_element(By.XPATH, xPath).text
            xPath = "//table/tbody/tr[" + str(t_row) + "]/td[" + str(8) + "]"
            fecha = datetime.strptime(driver.find_element(By.XPATH, xPath).text, "%Y-%m-%d")
            if(t_cargo in cargo and (fecha - datetime.today()).days > 30):
                index = t_row - 1
        
        time.sleep(2)

        filename = max(glob.glob(directorioDes + r'\*.xls'), key=os.path.getctime)
        informe_apo = leerUltimo(False, True, True)
        if index != -1:
            informe_apo = informe_apo.iloc[[index]]
            informe_completo = pd.concat([informe_completo, informe_apo], ignore_index=True)
        else:
            faltan = faltan.append(empleados.iloc[[i]], ignore_index=True)
        
        if len(informe_completo.index) >= 10:
            informe_completo.loc[informe_completo['Ciudad'] == 'BOGOTA D.C.', 'Ciudad'] = 'BOGOTA'
            informe_completo.loc[informe_completo['TelefonoR'] == 0, 'TelefonoR'] = 448518
            informe_completo.loc[informe_completo['TelefonoR'] == 'NO FIGURA', 'TelefonoR'] = 448518
            informe_completo['Nro'] = informe_completo['Nro'].apply(verificarNro)

            filename = apoInicio + obtenerTextoFecha()
            if apo_num >= 10:
                filename += str(apo_num)
            else:
                filename += "0" + str(apo_num)
            filename += ".xls"
            apo_num += 1

            print(directorioApo + "/" + filename)
            
            with pd.ExcelWriter(directorioApo + "/" + filename) as writer:
                informe_completo.to_excel(writer, index=False, header=True, sheet_name="ApoDatos")
                retirosVacios.to_excel(writer, index=False, header=True, sheet_name="Retiros")
            informe_completo = pd.DataFrame()
    
    if informe_completo.shape[0] > 0:
        filename = filename[:-6]
        if apo_num >= 10:
            filename += str(apo_num)
        else:
            filename += "0" + str(apo_num)
        filename += ".xlsx"

        informe_completo.to_excel(filename, index=False, header=True)
    faltan.to_excel(directorioDes + "/faltan.xlsx", index=False, header=True)

# Process
# Descarga de supervigilancia las personas en proceso de acreditación y ya acreditadas
descargarSupervigilancia(acreditadosUrl)
acreditados = leerSupervigilancia()
acreditados = acreditados.sort_values(by=['IdNum', 'Cargo'], ascending=[True, True])

descargarSupervigilancia(enprocesoUrl)
enproceso = leerSupervigilancia()
enproceso = enproceso.sort_values(by=['IdNum', 'Cargo'], ascending=[True, True])

subirListaAcreditados()

# Descarga los empleados como están registrados actualmente en semantica
descargarEmpleadosSemantica()
empleados = leerEmpleados()

# Añade la fecha de acreditación a quienes la tengan
empleados = aniadirFecha(empleados, enproceso, 'Estado')
empleados = aniadirFecha(empleados, acreditados, 'Vigen.Acr')

empleados = empleados[empleados['FECHA.ACR'] == 'Na']
fec_hoy = datetime.now()
fec_hoy = fec_hoy.strftime('%d_%m_%Y_%H_%M')
new_file = directorioDes + "\Acreditados" + fec_hoy + ".xlsx"
empleados.to_excel(new_file, index=False, header=True)

# empleados = pd.read_excel(directorioDes + "\Acreditados27_10_2022_08_52.xlsx")
descargarApos()

driver.quit()