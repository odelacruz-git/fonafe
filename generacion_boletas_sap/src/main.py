#!/usr/bin/env python
# -*- coding: cp1252 -*-
# -*- coding: utf-8 -*-

import time
import sys
from datetime import datetime
import logging
import pyautogui
import os
import json
import subprocesoReafirma as spr
import subprocesoSap as sps
import subprocesoNotificaCorreo as spn
from win32api import GetSystemMetrics
import locale



### Definimos la ruta y el archivo de configuración config.json
with open("C:\\Users\\admrpa\\Documents\\GitHub\\fonafe\\generacion_boletas_sap\\src\\config.json", 'r') as file:   config = json.load(file)

### Establecemos donde se guardarán las imagenes en la variable
ruta_imagenes = str(config['PRODUCCION']['ruta_imagenes'])


def main():
    print("Alto =", GetSystemMetrics(0))
    print("Ancho =", GetSystemMetrics(1))

    print(sys.getdefaultencoding())

    ancho = config['PRODUCCION']['Ancho_Resolucion']
    alto = config['PRODUCCION']['Alto_Resolucion']

    print('\nLocale from environment:', locale.getlocale())

    if(ancho == GetSystemMetrics(1) and alto == GetSystemMetrics(0)):
        inicio = datetime.now()
        time.sleep(2)
        log_string = inicio.strftime('%Y%m%d_%H%M%S')
        archivo_log = f'genera_notifica_boletas_{log_string}'
        print(archivo_log)
        path_log = str(config['PRODUCCION']['ruta_logs'])
        logging.basicConfig(filename=f'{path_log}\\{archivo_log}.log', filemode='w', level=logging.INFO, encoding='latin1', format='%(asctime)-5s %(name)-5s - %(levelname)-5s - %(message)-5s')
        logging.info("========== Iniciando Proceso Automatizado ==========")
        logging.info("========== Generación y Emisión de Boletas ==========")   

    
        if os.path.exists(ruta_imagenes) == True:
            logging.info(f'La ruta {ruta_imagenes} de las imagenes se validó correctamente')
        else: 
            logging.error(f'La ruta especificada NO existe, por favor verificar')

        ## Inicia subproceso Abrir SAP GUI y Conexión
        try:
            logging.info("Inicia subproceso: Abrir SAP GUI")
            sps.openSAP()
        except:
            logging.error(sys.exc_info())


        ## Inicia loginSAP
        try:
            logging.info("Inicia subproceso: Logueo en SAP")
            sps.loginSAP()
        except:
            logging.error(sys.exc_info())
        else:
            logging.info('Logueo exitoso en SAP')
        finally:
            logging.info('Termina subproceso: Logueo en SAP')

        ## Inicia Generación de Boletas en SAP
        try:
            logging.info("Inicia subproceso: Generar boletas en SAP")
            sps.generarBoletaSAP()
        except:
            logging.error(sys.exc_info())
            logging.error("Subproceso: Generar boletas SAP terminó en error, por favor revisar")
        else:
            logging.info('subproceso culminado correctamente')
        finally:
            logging.info('Termina subproceso: Generar boletas en SAP')
        
        ## Inicia Logout en SAP
        try:
            logging.info("Inicia subproceso: Logout en SAP")
            sps.closeSap()
        except:
            logging.error(sys.exc_info())
            logging.error("Subproceso: Logout SAP terminó en error, por favor revisar")
        else:
            logging.info('Subproceso Logout en SAP ejecutado correctamente')
        finally:
            logging.info('Termina subproceso: Logout en SAP')

        ## Inicia abrirReafirma
        try:
            time.sleep(5)
            logging.info("Inicia subproceso: Abrir Reafirma")
            spr.abrirReafirma()
        except:
            logging.error(sys.exc_info())
            logging.error("Subproceso: Abrir Reafirma terminó en error, por favor revisar")
        
        ## Inicia Firmaboletas
        try:
            logging.info("Inicia subproceso: Firma de Boletas")
            spr.firmarBoletas()
        except:
            logging.error(sys.exc_info())
            logging.error("Subproceso: Firma de boletas terminó en error, por favor revisar")
    
        ## Inicia NotificaBoletas con Selenium
        try:
            time.sleep(5)
            logging.info("Inicia subproceso: Inicia  Notifica boletas por correo")
            spn.notificaCorreo()
        except:
            print(sys.exc_info())
            logging.error(sys.exc_info())
            logging.error("Subproceso: Inicia Notifica boletas por correo terminó en error, por favor revisar")

        fin = datetime.now()

        delta = fin - inicio

        logging.info(f'Tiempo de ejecución: {delta}\n Proceso finalizado!\n')
    else:
        print("No cumple con la resolucion esperada")

if __name__=='__main__':
    main()



