import glob
from msilib.schema import File
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from datetime import datetime
import locale

from selenium.webdriver import ActionChains
from selenium.webdriver.common.actions.action_builder import ActionBuilder
from selenium.webdriver.common.actions.mouse_button import MouseButton

import os

import pyautogui

import json

import logging

### Definimos la ruta y el archivo de configuración config.json
with open("C:\\Users\\admrpa\\Documents\\GitHub\\fonafe\\generacion_boletas_sap\\src\\config.json", 'r') as file:   config = json.load(file)

### Establecemos donde se guardarán las imagenes en la variable
ruta_imagenes = str(config['PRODUCCION']['ruta_imagenes'])




def notificaCorreo():
    """ Este método realiza el inicio de sesión sobre el portal de notificación"""

    options = Options()
    options.add_argument("start-maximized")
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    driver.get("https://notifica.electroucayali.com.pe/pages/login")
    time.sleep(2)

    user = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Documento']")
    user.send_keys("42723795")
    time.sleep(2)

    contraseña = driver.find_element(By.CSS_SELECTOR, "input[placeholder='Contraseña']")
    contraseña.send_keys("4272")
    time.sleep(2)

    ingresar = driver.find_element(By.CSS_SELECTOR,"a.btn.btn-rose").click()

    time.sleep(2)

    sidebar = driver.find_element(By.XPATH,"//div[contains(@class,'sidebar-wrapper ps')]")
    time.sleep(1)
    ActionChains(driver)\
        .move_to_element(sidebar)\
        .perform()

    time.sleep(2)

    notificar = driver.find_element(By.XPATH,"//p[text()='Notificar']")
    ActionChains(driver)\
        .move_to_element(notificar)\
        .click(notificar)\
        .perform()

    time.sleep(1.5)

    btn_notificacion = driver.find_element(By.XPATH,"(//span[@class='sidebar-normal'][contains(.,'Notificacion')])[1]")
    ActionChains(driver)\
        .move_to_element(btn_notificacion)\
        .click(btn_notificacion)\
        .perform()

    time.sleep(9)

    locale.setlocale(locale.LC_ALL, 'es-ES') 
    now = datetime.now()

    # Asignar nombre del mes para crear carpeta donde se guardaran las boletas.
    if config['PRODUCCION']['periodo_mes'] == '01':
            carpeta_mes = "ENERO"
    elif config['PRODUCCION']['periodo_mes'] == '02':
            carpeta_mes = "FEBRERO"
    elif config['PRODUCCION']['periodo_mes'] == '03':
            carpeta_mes = "MARZO"
    elif config['PRODUCCION']['periodo_mes'] == '04':
            carpeta_mes = "ABRIL"
    elif config['PRODUCCION']['periodo_mes'] == '05':
            carpeta_mes = "MAYO"
    elif config['PRODUCCION']['periodo_mes'] == '06':
            carpeta_mes = "JUNIO"
    elif config['PRODUCCION']['periodo_mes'] == '07':
            carpeta_mes = "JULIO"
    elif config['PRODUCCION']['periodo_mes'] == '08':
            carpeta_mes = "AGOSTO"
    elif config['PRODUCCION']['periodo_mes'] == '09':
            carpeta_mes = "SETIEMBRE"
    elif config['PRODUCCION']['periodo_mes'] == '10':
            carpeta_mes = "OCTUBRE"
    elif config['PRODUCCION']['periodo_mes'] == '11':
            carpeta_mes = "NOVIEMBRE"
    elif config['PRODUCCION']['periodo_mes'] == '12':
            carpeta_mes = "DICIEMBRE"
    else:
            carpeta_mes = "SIN ASIGNAR"

    asunto = driver.find_element(By.XPATH,"//input[contains(@name,'asunto')]")
    ActionChains(driver)\
        .move_to_element(asunto)\
        .click(asunto)\
        .pause(1)\
        .send_keys("Boleta de remuneraciones de " + carpeta_mes + " del " + config['PRODUCCION']['periodo_anio'])\
        .perform()
    
    mensaje = driver.find_element(By.XPATH,"//textarea[contains(@name,'notificacion_mensaje')]")

    
    Plantilla = "Estimado trabajador %usuario% El presente es para saludarlo y a la vez remitirle adjunto la boleta de remuneraciones del mes de " + carpeta_mes + " " + config['PRODUCCION']['periodo_anio'] + ", puede llamar al departamento de recursos humanos. Puede consultarlo en %url%"
 
    
    ###Estimado trabajador %usuario% El presente es para saludarlo y a la vez remitirle adjunto la boleta de remuneraciones del mes de <MES> <AÑO>, puede llamar al departamento de recursos humanos. Puede consultarlo en %url%"
    ActionChains(driver)\
        .move_to_element(mensaje)\
        .click(mensaje)\
        .pause(1)\
        .send_keys(Plantilla)\
        .perform()

    time.sleep(3)

    ## Seleccionar Categoria
    categoria = driver.find_element(By.XPATH,"//span[@class='ng-tns-c79-3 ng-star-inserted'][contains(.,'Seleccione una Categoría')]")
    ActionChains(driver)\
        .move_to_element(categoria)\
        .click(categoria)\
        .perform()

    time.sleep(1.5)

    cat_boleta = driver.find_element(By.XPATH,"//span[@class='mat-option-text'][contains(.,'Boleta')]")
    ActionChains(driver)\
        .move_to_element(cat_boleta)\
        .click(cat_boleta)\
        .perform()

    time.sleep(1.5)

    ## Seleccionar Canal
    canal = driver.find_element(By.XPATH,"//span[@class='ng-tns-c79-5 ng-star-inserted'][contains(.,'Seleccione Documento')]")
    ActionChains(driver)\
        .move_to_element(canal)\
        .click(canal)\
        .perform()

    time.sleep(1.5)

    canal_email = driver.find_element(By.XPATH,"//span[@class='mat-option-text'][contains(.,'email')]")
    ActionChains(driver)\
        .move_to_element(canal_email)\
        .click(canal_email)\
        .perform()

    time.sleep(1.5)

    ## Seleccionar Indice
    indice = driver.find_element(By.XPATH,"//span[@class='ng-tns-c79-7 ng-star-inserted'][contains(.,'Seleccione Documento')]")
    ActionChains(driver)\
        .move_to_element(indice)\
        .click(indice)\
        .perform()

    time.sleep(1.5)

    indice_codsap = driver.find_element(By.XPATH,"//span[@class='mat-option-text'][contains(.,'CODSAP')]")
    ActionChains(driver)\
        .move_to_element(indice_codsap)\
        .click(indice_codsap)\
        .perform()
    
    time.sleep(3)

    ##Cargar carpeta de boletas
    seleccionar_archivos = driver.find_element(By.XPATH,"//input[contains(@type,'file')]")
    ActionChains(driver)\
        .click(seleccionar_archivos)\
        .perform()
    
    time.sleep(6)

    
    filename = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\file_name.png", confidence=0.9)
    pyautogui.click(filename)

    
    pyautogui.write("C:\\Users\\admrpa\\Documents\\SAP\\SAP GUI\\BOLETAS_CON_FIRMA\\2022\\JUNIO\\Boletas\\Boletas\\Test")
    time.sleep(4)
    pyautogui.hotkey("enter")
    time.sleep(2)
    pyautogui.write("*[R].pdf")
    time.sleep(2)
    pyautogui.hotkey("enter")
    logging.info("Seleccionar todo")
    
    SelecttAllX,SelecttAllY  = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\EspacioBlanco.png", confidence=0.9)
    time.sleep(1)
    SelecttAllY += 80
    pyautogui.click(SelecttAllX, SelecttAllY)
    time.sleep(1)
    pyautogui.hotkey('ctrl','a')
    time.sleep(2)

    pyautogui.hotkey("enter")

    time.sleep(4)
    ReprocesarArchivos = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\ProcesarArchivos.png", confidence=0.9)
    pyautogui.click(ReprocesarArchivos)
    

    seleccionar_archivos = driver.find_element(By.XPATH,"//span[@class='mat-button-wrapper'][contains(.,'Preprocesar Archivos')]")
    ActionChains(driver)\
        .click(seleccionar_archivos)\
        .perform()

    time.sleep(25)


    """"
    Notificar = driver.find_element(By.XPATH,"//span[@class='mat-button-wrapper'][contains(.,'Notificar')]")
    ActionChains(driver)\
        .click(Notificar)\
        .perform()
    """
    time.sleep(1)

    ###pyautogui.alert("El proceso termino con éxito")



    