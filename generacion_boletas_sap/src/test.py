import os
import glob

import os

import pyautogui
import time
import json
import json
import logging

from datetime import datetime

import locale

import ctypes

import lackey
from lackey import click 

import subprocesoSap as sps
from win32api import GetSystemMetrics

print("Ancho =", GetSystemMetrics(0))
print("Alto =", GetSystemMetrics(1))



###click('imagenes\SelectAll.png')

try:
    logging.info("Inicia subproceso: Abrir SAP GUI")
    sps.openSAP()
except:
    logging.info("Error al abrir SAP")

time.sleep(5)
click('imagenes\options.png')
logging.info("Evento clic en botón 'Opciones'")

# Maximizar menu princiapl SAP Easy Access (opcional) Puede que ya aparezca maximizado
time.sleep(1)
click('imagenes\\maximiza_sap.png')
logging.info("Evento clic en botón 'Maximizar'")

# Ingresar Numero de transacción (Generacion Boletas de Pago)
time.sleep(1)
click('imagenes\\texbox_trx.png')    
time.sleep(1)


lackey.exit('')




        
"""

Contador = 0
ActividadExitosa = False
while Contador<30 and ActividadExitosa==False:
    try:
        print("Alto =", GetSystemMetrics(1))
        time.sleep(1)
    except:
        logging.error("TRAY")
    Contador += 1


print(os.listdir(path="C:\\Users\\admrpa\\Documents\\SAP\\SAP GUI\\BOLETAS_CON_FIRMA\\2022\\JUNIO"))

print(glob.glob("C:\\Users\\admrpa\\Documents\\SAP\\SAP GUI\\BOLETAS_CON_FIRMA\\2022\\JUNIO\*.pdf"))


### Definimos la ruta y el archivo de configuración config.json
with open("C:\\Proyectos\\proyectoFonafe\\generacion_boletas_sap\\src\\config.json", 'r') as file:   config = json.load(file)

### Establecemos donde se guardarán las imagenes en la variable
ruta_imagenes = str(config['PRODUCCION']['ruta_imagenes'])

SelecttAllX , SelecttAllY = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\SelectAll.png", confidence=0.9)
Recuadro = (SelecttAllX - 19)
pyautogui.click(Recuadro, SelecttAllY)
pyautogui.hotkey("enter")

time.sleep(4)
ReprocesarArchivos = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\ProcesarArchivos.png", confidence=0.9)
pyautogui.click(ReprocesarArchivos)

time.sleep(4)
pyautogui.hotkey('ctrl','a')

locale.setlocale(locale.LC_ALL, 'es-ES') 
now = datetime.now()
Plantilla = "Estimado trabajador %usuario% El presente es para saludarlo y a la vez remitirle adjunto la boleta de remuneraciones del mes de " + now.strftime('%B') + " " + now.strftime('%Y') + ", puede llamar al departamento de recursos humanos. Puede consultarlo en %url%"
    
print(Plantilla)

"""