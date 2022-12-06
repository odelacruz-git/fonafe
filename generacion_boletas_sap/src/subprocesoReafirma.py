import subprocess
import time
import sys
from datetime import datetime
import logging
import pyautogui
import os
import json


## Definimos la ruta y el archivo de configuración config.json
with open("C:\\Proyectos\\proyectoFonafe\\generacion_boletas_sap\\src\\config.json", 'r') as file:   config = json.load(file)
ruta_imagenes = str(config['PRODUCCION']['ruta_imagenes'])

carpeta_mes = "SIN ASIGNAR"
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


def abrirReafirma():
    """Esta función abre la aplicación REAFIRMA PDF"""
    
    try:
        path = "C:\\Users\\admrpa\\Desktop\\ReFirmaPDF.exe"
        subprocess.Popen(path)
        time.sleep(20)
    except:
        logging.error(sys.exc_info())
       
  

    

def firmarBoletas():
    """Esta función, ejecuta el proceso de firma digital dentro de cada una de las boletas"""
    
    # Click en opción "Archivo"
    x_btn_archivo, y_btn_archivo = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\r_archivo_01.png", confidence=0.9)
    #x_password = x_password + 60
    pyautogui.click(x_btn_archivo, y_btn_archivo)

    # Click en opción "Firmar por Lote"
    time.sleep(1)
    x_firma_lote, y_firma_lote = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\r_firmaLote_02.png", confidence=0.9)
    #x_password = x_password + 60
    pyautogui.click(x_firma_lote, y_firma_lote)

    # Click para ingresar la ruta donde leeremos las boletas a firmar
    time.sleep(2)
    x_entrada, y_entrada = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\r_entrada_03.png", confidence=0.9)
    print(x_entrada,y_entrada)
    x_entrada = x_entrada + 552
    pyautogui.click(x_entrada,y_entrada)
    time.sleep(2)
    ruta_boletas_sin_firma = str(config['PRODUCCION']['ruta_boletas_sin_firma']) + str(config['PRODUCCION']['periodo_anio']) + "\\"+carpeta_mes+"\\"
    pyautogui.write(ruta_boletas_sin_firma)
    # Enter
    time.sleep(2)
    pyautogui.press('enter')


    #Crear y validar si existe carpeta del MES actual donde se descargaran las boletas
    ruta_boletas_con_firma = str(config['PRODUCCION']['ruta_boletas_con_firma']) + str(config['PRODUCCION']['periodo_anio']) + "\\"+carpeta_mes+"\\Boletas"
    os.makedirs(ruta_boletas_con_firma, exist_ok=True)

    # Click para ingresar la ruta donde se almacenaran las boletas firmadas
    time.sleep(2)
    x_salida, y_salida = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\r_salida_04.png", confidence=0.9)
    print(x_salida,y_salida)
    x_salida = x_salida + 552
    pyautogui.click(x_salida,y_salida)
    time.sleep(2)
    pyautogui.write(ruta_boletas_con_firma)
    # Enter
    time.sleep(2)
    pyautogui.press('enter')

     # Seleccionar el certificado digital generado por reniec
    time.sleep(2)
    x_certificado, y_certificado = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\r_escoge_certificado_digital_05.png", confidence=0.9)
    print(x_certificado,y_certificado)
    pyautogui.click(x_certificado,y_certificado)
    time.sleep(2)

    ##Obviamos la selección de la imagen de la firma.

    ##Presionar Enter para ir al siguiente paso.
    pyautogui.press('enter')
    time.sleep(3)

    pyautogui.keyDown('shift')
    pyautogui.press('a')
    pyautogui.press('c')
    pyautogui.press('e')
    pyautogui.press('p')
    pyautogui.press('t')
    pyautogui.press('o')
    pyautogui.keyUp('shift')

     # Click en el boton firmar
    time.sleep(2)
    x_firmar, y_firmar = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\r_escribir_acepto_firmar.png", confidence=0.9)
    print(x_firmar,y_firmar)
    x_firmar = x_firmar + 140
    pyautogui.click(x_firmar,y_firmar)
    
    time.sleep(3)

    ## Escribir el Private Key definido cuando se instaló el certificado.
    pyautogui.write(config['PRODUCCION']['private_key_cert'])

    # Click en el boton firmar
    time.sleep(2)
    pyautogui.press('enter')

    ## Se asigna un tiempo considerando la cantidad de boletas, en promedio 1s cada 2 boletas. (El tiempo se encuentra en el archivo Config.json)

    time.sleep(config['PRODUCCION']['tiempo_firma_boleta'])

    proceso_finalizado = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\r_proceso_finalizado_12.png", confidence=0.9) 
    if proceso_finalizado != None:
            print("La firma por lote dio como resultado PROCESO FINALIZADO")
            time.sleep(2)
            ##Presionar Enter para ir al siguiente VER REPORTE
            x_reporte, y_reporte = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\ver_reporte.png", confidence=0.9)
            print(x_reporte,y_reporte)
            pyautogui.click(x_reporte,y_reporte)
            time.sleep(2)
    else:
            print("REFIRMA no mostró el mensaje PROCESO FINALIZADO, por favor revisar el log")

    ## Click en cerrar REFIRMA.
    x_cerrar, y_cerrar = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\r_cierra_refirma.png", confidence=0.9)
    print(x_cerrar,y_cerrar)
    pyautogui.click(x_cerrar,y_cerrar)
    time.sleep(2)
