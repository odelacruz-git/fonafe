#!/usr/bin/env python
# -*- coding: cp1252 -*-
# -*- coding: utf-8 -*-

import subprocess
import win32com.client
import time
import sys
from datetime import datetime
import logging
import pyautogui
import os
import json

import subprocesoReafirma as spr

### Definimos la ruta y el archivo de configuración config.json
with open("C:\\Proyectos\\proyectoFonafe\\generacion_boletas_sap\\src\\config.json", 'r') as file:   config = json.load(file)

### Establecemos donde se guardarán las imagenes en la variable
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

def openSAP():
    """Esta función abre la aplicación SAP y la conexión FONAFE_RPA"""
    try:
        path = "C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe"
        subprocess.Popen(path)
        time.sleep(2)

        logging.info("Ejecutando SAP.EXE")

        SapGuiAuto = win32com.client.GetObject('SAPGUI')
        if not type(SapGuiAuto) == win32com.client.CDispatch:
            return

        application = SapGuiAuto.GetScriptingEngine
        if not type(application) == win32com.client.CDispatch:
            SapGuiAuto = None
            return

        connection = application.OpenConnection(config['PRODUCCION']['conexionSAP']['nombre'], True)
        if not type(connection) == win32com.client.CDispatch:
            application = None
            SapGuiAuto = None
            return
        logging.info("Abriendo conexión hacia " + config['PRODUCCION']['conexionSAP']['nombre'])
    except:
        print(sys.exc_info())

    finally:
        connection = None
        application = None
        SapGuiAuto = None
        
    time.sleep(4)

def loginSAP():
    """Esta función ingresa el código de transaccion SAP y las credenciales para el logueo a SAP en la conexion establecida"""
  
    # Click en el menu de opciones de ventana
    time.sleep(1)
    option = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\options.png", confidence=0.9)
    pyautogui.click(option)
    logging.info("Evento clic en botón 'Opciones', con posición " + str(option))

    # Maximizar menu princiapl SAP Easy Access (opcional) Puede que ya aparezca maximizado
    time.sleep(1)
    maximiza_sap = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\maximiza_sap.png", confidence=0.9)
    pyautogui.click(maximiza_sap)
    logging.info("Evento clic en botón 'Maximizar', con posición " + str(maximiza_sap))

    # Ingresar Numero de transacción (Generacion Boletas de Pago)
    time.sleep(1)
    textbox_trx = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\texbox_trx.png", confidence=0.9)
    pyautogui.click(textbox_trx)
    time.sleep(1)
    pyautogui.write(config['PRODUCCION']['conexionSAP']['trx_sap'])
    logging.info("Evento para ingresar número de transacción")

    # Ubicarse y escribir el username
    time.sleep(1)
    x_user, y_user = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\user.png", confidence=0.9)
    x_user = x_user + 60
    pyautogui.click(x_user, y_user)
    pyautogui.write(config['PRODUCCION']['conexionSAP']['cnx_usuario'])
    logging.info("Evento ubicar y escribir el username, en posicion: " + str(x_user) + " - " + str(y_user))

    # Ubicarse y escribir el password
    time.sleep(1)
    x_password, y_password = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\password.png", confidence=0.9)
    x_password = x_password + 60
    pyautogui.click(x_password, y_password)
    pyautogui.write(config['PRODUCCION']['conexionSAP']['cnx_contraseña'])
    time.sleep(1)
    pyautogui.hotkey("enter")
    logging.info("Evento ubicar y escribir password, en posicion: " + str(x_password) + " - " + str(y_password))

def generarBoletaSAP():
    
    """ Proceso en SAP, que interactua con los inputs en el archivo conf.json para generar las boletas"""

    # Ingresar Sociedad
    time.sleep(2)
    x_sociedad, y_sociedad = pyautogui.locateCenterOnScreen(
        ruta_imagenes + "\\texbox_sociedad.png",
        confidence=0.9)
    x_sociedad += 100
    pyautogui.click(x_sociedad, y_sociedad)
    pyautogui.write("0600")
    pyautogui.hotkey("enter")
    logging.info("Evento ingresar a Sociedad, en posicion: " + str(x_sociedad) + " - " + str(y_sociedad))
    # Ingresar Nomina
    time.sleep(2)
    x_nomina, y_nomina = pyautogui.locateCenterOnScreen(
        ruta_imagenes + "\\textbox_nomina.png",
        confidence=0.9)
    x_nomina += 105
    pyautogui.click(x_nomina, y_nomina)
    pyautogui.write("U1")
    logging.info("Evento ingresar a Nomina, en posicion: " + str(x_nomina) + " - " + str(y_nomina))

    time.sleep(1)
    # Ingresar mes
    x_periodo, y_periodo = pyautogui.locateCenterOnScreen(
        ruta_imagenes + "\\textbox_periodo.png",
        confidence=0.9)
    mes = (x_periodo + 56)
    pyautogui.click(mes, y_periodo)
    pyautogui.write(config['PRODUCCION']['periodo_mes'])
    logging.info("Evento ingresar Mes, en posicion: " + str(mes) + " - " + str(y_periodo))


    # Ingresar año
    anio = (x_periodo + 90)
    pyautogui.click(anio, y_periodo)
    pyautogui.write(config['PRODUCCION']['periodo_anio'])


    # Ingresar calidad nomina.
    time.sleep(1)
    x_calidad_nomina, y_calidad_nomina = pyautogui.locateCenterOnScreen(
        ruta_imagenes + "\\calidad_nomina.png",
        confidence=0.9)
    nomina = x_calidad_nomina - 19
    pyautogui.click(nomina, y_calidad_nomina)
    pyautogui.write("U1")
    pyautogui.hotkey("enter")
    logging.info("Evento ingresar Calidad Nomina, en posicion: " + str(nomina) + " - " + str(y_calidad_nomina))


    # Seleccionar radio button "Archivo pdf" y desactivar checkbox "PDF Unico"
    time.sleep(1)
    archivo_pdf = pyautogui.locateCenterOnScreen(
        ruta_imagenes + "\\archivo_pdf.png",
        confidence=0.9)
    pyautogui.click(archivo_pdf)

    logging.info("Evento seleccionar 'Archivo pdf', en posicion: " + str(archivo_pdf))

    time.sleep(1)
    pdf_unico = pyautogui.locateCenterOnScreen(
        ruta_imagenes + "\\pdf_unico.png",
        confidence=0.9)
    pyautogui.click(pdf_unico)

    logging.info("Evento desactivar 'PDF Unico', en posicion: " + str(pdf_unico))

    #Crear y validar si existe carpeta del MES actual donde se descargaran las boletas
    ruta_descarga = str(config['PRODUCCION']['ruta_boletas_sin_firma']) + str(config['PRODUCCION']['periodo_anio']) + "\\"+carpeta_mes+"\\"
    logging.info("ruta de descarga es: " + ruta_descarga)
    os.makedirs(ruta_descarga, exist_ok=True)


    #colocar ruta de descarga boletas y ejecutar proceso con f8
    time.sleep(2)
    x_ruta_descarga, y_ruta_descarga = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\ruta_descarga.png", confidence=0.9)
    x_ruta_descarga = x_ruta_descarga + 109
    pyautogui.click(x_ruta_descarga, y_ruta_descarga)
    pyautogui.write(ruta_descarga)
    time.sleep(4)
    pyautogui.hotkey('f8')
    time.sleep(3)
    logging.info("Evento Colocar ruta de descarga boletas, en posicion: " + str(x_ruta_descarga) + " - " + str(y_ruta_descarga))

    # Reconocer icono OK de culminación del proceso de descarga
    mensaje_final_error = pyautogui.locateCenterOnScreen(
        ruta_imagenes + "\\mensaje_final_error.png",
        confidence=0.9) 
    print(mensaje_final_error)
    if mensaje_final_error != None :
        logging.info("Las boletas NO se descargaron correctamente, se creo la carpeta del MES pero se encuentra vacía, por favor validarl os parametros ingresados")
    else:
        time.sleep(70)
        mensaje_final_ok = pyautogui.locateCenterOnScreen(
        ruta_imagenes + "\\mensaje_final_ok.png",
        confidence=0.9)
        print(mensaje_final_ok)
        if mensaje_final_ok != None:
            logging.info("Las boletas se descargaron correctamente en " + ruta_descarga)
        else:
            logging.warning("No apareció el icono de proceso culminado correctamente en SAP, por favor revisar el log")  


def closeSap():
    # Teclear alt + f4 para cerrar SAP
    ruta_imagenes = str(config['PRODUCCION']['ruta_imagenes'])
    print(ruta_imagenes)
    time.sleep(1)
    pyautogui.hotkey('alt', 'f4')

    time.sleep(1)
    salir = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\salir.png", confidence=0.9)
    pyautogui.click(salir)

    logging.info("Evento cerrar SAP, en posicion: " + str(salir))

    time.sleep(1)
    cerrar_logon = pyautogui.locateCenterOnScreen(ruta_imagenes + "\\cerrar_logon.png", confidence=0.9)
    pyautogui.click(cerrar_logon)

    logging.info("Evento confirmar cerrar SAP, en posicion: " + str(cerrar_logon))

