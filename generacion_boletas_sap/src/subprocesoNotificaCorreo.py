import glob
from msilib.schema import File
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver import ActionChains
from selenium.webdriver.common.actions.action_builder import ActionBuilder
from selenium.webdriver.common.actions.mouse_button import MouseButton

import os

import pyautogui

import json

import logging

### Definimos la ruta y el archivo de configuración config.json
with open("C:\\Proyectos\\proyectoFonafe\\generacion_boletas_sap\\src\\config.json", 'r') as file:   config = json.load(file)

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

    time.sleep(8)

    asunto = driver.find_element(By.XPATH,"//input[contains(@name,'asunto')]")
    ActionChains(driver)\
        .move_to_element(asunto)\
        .click(asunto)\
        .pause(1)\
        .send_keys("Boleta de remuneraciones de <MES> del <AÑO>")\
        .perform()
    
    mensaje = driver.find_element(By.XPATH,"//textarea[contains(@name,'notificacion_mensaje')]")
    ActionChains(driver)\
        .move_to_element(mensaje)\
        .click(mensaje)\
        .pause(1)\
        .send_keys("Estimado trabajador %usuario% El presente es para saludarlo y a la vez remitirle adjunto la boleta de remuneraciones del mes de <MES> <AÑO>, puede llamar al departamento de recursos humanos. Puede consultarlo en %url%")\
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

    
    pyautogui.write("C:\\Users\\admrpa\\Documents\\SAP\\SAP GUI\\BOLETAS_CON_FIRMA\\2022\\JUNIO\\Boletas\\Boletas")
    time.sleep(4)
    pyautogui.hotkey("enter")
    time.sleep(4)
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
    seleccionar_archivos = driver.find_element(By.XPATH,"//span[@class='mat-button-wrapper'][contains(.,'Notificar')]")
    ActionChains(driver)\
        .click(seleccionar_archivos)\
        .perform()
    """
    pyautogui.alert("El proceso termino con éxito")

    