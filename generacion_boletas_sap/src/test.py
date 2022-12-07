import os
import glob

import os

import pyautogui
import time
import json

from datetime import datetime

import locale

print(os.listdir(path="C:\\Users\\admrpa\\Documents\\SAP\\SAP GUI\\BOLETAS_CON_FIRMA\\2022\\JUNIO"))

print(glob.glob("C:\\Users\\admrpa\\Documents\\SAP\\SAP GUI\\BOLETAS_CON_FIRMA\\2022\\JUNIO\*.pdf"))


### Definimos la ruta y el archivo de configuración config.json
with open("C:\\Proyectos\\proyectoFonafe\\generacion_boletas_sap\\src\\config.json", 'r') as file:   config = json.load(file)

### Establecemos donde se guardarán las imagenes en la variable
ruta_imagenes = str(config['PRODUCCION']['ruta_imagenes'])

"""
SelecttAllX , SelecttAllY = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\SelectAll.png", confidence=0.9)
Recuadro = (SelecttAllX - 19)
pyautogui.click(Recuadro, SelecttAllY)
pyautogui.hotkey("enter")

time.sleep(4)
ReprocesarArchivos = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\ProcesarArchivos.png", confidence=0.9)
pyautogui.click(ReprocesarArchivos)

time.sleep(4)
pyautogui.hotkey('ctrl','a')"""

locale.setlocale(locale.LC_ALL, 'es-ES') 
now = datetime.now()
Plantilla = "Estimado trabajador %usuario% El presente es para saludarlo y a la vez remitirle adjunto la boleta de remuneraciones del mes de " + now.strftime('%B') + " " + now.strftime('%Y') + ", puede llamar al departamento de recursos humanos. Puede consultarlo en %url%"
    
print(Plantilla)