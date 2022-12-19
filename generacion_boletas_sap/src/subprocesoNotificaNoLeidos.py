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

from selenium.webdriver import ActionChains
from selenium.webdriver.common.actions.action_builder import ActionBuilder
from selenium.webdriver.common.actions.mouse_button import MouseButton

import os

import pyautogui
import json
import logging

from tkinter import messagebox

from bs4 import BeautifulSoup

import requests

### Definimos la ruta y el archivo de configuración config.json
with open("C:\\Users\\admrpa\\Documents\\GitHub\\fonafe\\generacion_boletas_sap\\src\\config.json", 'r') as file:   config = json.load(file)

### Establecemos donde se guardarán las imagenes en la variable
ruta_imagenes = str(config['PRODUCCION']['ruta_imagenes'])




def notificaCorreonoleidos():
     
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

    try:
        ingresar = driver.find_element(By.CSS_SELECTOR,"a.btn.btn-rose").click()
    except:
        messagebox.showinfo(message="Mensaje", title="Excepcion")

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

    logging.info("Evento ingresar a Notifica")

    btn_notificacion = driver.find_element(By.XPATH,"//span[@class='sidebar-normal'][contains(.,'Lista Notificaciones')]")
    ActionChains(driver)\
        .move_to_element(btn_notificacion)\
        .click(btn_notificacion)\
        .perform()

    time.sleep(9)

    #Captura datos de tablero y evalua boletas generadas hasta encontrar la primer boleta creada y enviada
    tablero = driver.find_element(By.XPATH,"//div[@class='col-sm-12'][contains(.,'IDFecha CreacionFecha EnvioResultadoCanalCategoriaAccionesIDFecha CreacionFecha EnvioResultadoCanalCategoria Acciones10114/12/2022 11:46:17emailBoleta10014/12/2022 11:38:45emailBoleta9914/12/2022 11:03:49emailBoleta9814/12/2022 10:56:28emailBoleta9714/12/2022 10:40:21emailBoleta9614/12/2022 10:26:05emailBoleta9513/12/2022 18:20:20emailBoleta9413/12/2022 17:57:33emailBoleta9313/12/2022 17:30:03emailBoleta9212/12/2022 17:35:43emailBoleta')]")
    indiceTablero = 1
    EncontroBoletasEnviadas = False
    while EncontroBoletasEnviadas==False:
        try:
            fechaCreacion = driver.find_element(By.XPATH,"//tbody/tr["+ str(indiceTablero) +"]/td[2]")
            fechaEnvio = driver.find_element(By.XPATH,"//tbody/tr["+ str(indiceTablero) +"]/td[3]")
            print(str(indiceTablero)+ " - " + str(fechaCreacion.text) + " - " + str(fechaEnvio.text))
            if fechaCreacion.text != "" and fechaEnvio.text != "":
                print(str(indiceTablero))
                btn_acciones = driver.find_element(By.XPATH,"//tbody/tr["+ str(indiceTablero) +"]/td[7]/p[1]/a[1]/i[1]")
                ActionChains(driver)\
                    .move_to_element(btn_acciones)\
                    .click(btn_acciones)\
                    .perform()
                EncontroBoletasEnviadas = True
            else:
                print(str(indiceTablero))
        except:
            btn_acciones = driver.find_element(By.XPATH,"//a[contains(text(),'Siguiente')]")
            ActionChains(driver)\
                .move_to_element(btn_acciones)\
                .click(btn_acciones)\
                .perform()
            time.sleep(1)
            indiceTablero = 1
        indiceTablero = indiceTablero + 1

    time.sleep(2)

    NuevaVentana = driver.window_handles[1]
    driver.switch_to.window(NuevaVentana)

    indiceNotificaciones = 1
    EncontroBoletasEnviadas = False
    while EncontroBoletasEnviadas==False:
        try: 
            fechaLectura = driver.find_element(By.XPATH,"//tbody/tr["+ str(indiceNotificaciones) +"]/td[3]")
            print(str(indiceTablero)+ " - " + str(fechaLectura.text))
            if fechaLectura.text == "":
                print("Se encontro uno que no lee")
                NombreEmpleado = driver.find_element(By.XPATH,"//tbody/tr["+ str(indiceNotificaciones) +"]/td[4]")
                print(str(NombreEmpleado.text))
                #Enviar correo
            else:
                print(str(indiceNotificaciones))
        except:
            EncontroBoletasEnviadas = True
            btn_acciones = driver.find_element(By.XPATH,"//a[contains(text(),'Siguiente')]")
            ActionChains(driver)\
                .move_to_element(btn_acciones)\
                .click(btn_acciones)\
                .perform()
            time.sleep(1)
            indiceNotificaciones = 1
        indiceNotificaciones = indiceNotificaciones + 1

    logging.info("Evento notificar")

    ###pyautogui.alert("El proceso termino con éxito")

notificaCorreonoleidos()


    