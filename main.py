from logging import exception
import re

import time

import pandas as pd 

from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By

# Lee dos archivos de Excel e imprime un mensaje.
contactos = pd.ExcelFile("Plantilla Excel/contactos.xlsx").parse(0)
datosEnvio = pd.ExcelFile("Plantilla Excel/datosEnvio.xlsx").parse(0)

print("\n---- ---- ---- ----")
print("Exportando contactos.xlsx \nExportando datosEnvio.xlsx")

# Obtener el número de filas en el marco de datos de contactos.
cantidadContactos = contactos.shape[0]

# Eliminación de duplicados del marco de datos.
contactos = contactos.drop_duplicates()

# Obtener el número de filas en el marco de datos.
cantidadContactosDepurados = contactos.shape[0]

# Restar el número de filas en el marco de datos del número de filas en el marco de datos sin duplicados.
totalDuplicados = cantidadContactos - cantidadContactosDepurados

print("\n---- ---- ---- ----")
print(f"En el archivo de contactos ahi {totalDuplicados} contactos duplicados \nSe han removido exitosamente")

"""
    Toma una cadena como entrada y devuelve una cadena con todos los espacios eliminados
    :param texto: el texto a procesar
    :return: el texto sin espacios.
"""
def espacios(texto):
    texto = re.sub(r"[\s]+","",texto)
    return texto

# Convirtiendo la columna 'Contactos' a una cadena y luego aplicándole la función espacios.
contactos["Contactos"] = contactos["Contactos"].astype(str).apply(espacios)

print("\n---- ---- ---- ----")
print("Quitando los espacios encontrados en la plantilla contactos.xlsx")

# Creando una lista de los valores de la columna `Mensaje` en el dataframe `datosEnvio`.
mensaje = list(datosEnvio.Mensaje)
mensaje = [x for x in mensaje if str(x) != 'nan']
# Creando una lista de los valores de la columna `RutaImagen` en el dataframe `datosEnvio`.
rutaImagen = list(datosEnvio.RutaImagen)
rutaImagen = [x for x in rutaImagen if str(x) != 'nan']

# Creando una lista de los valores de la columna `RutaVideo` en el dataframe `datosEnvio`.
rutaVideo = list(datosEnvio.RutaVideo)
rutaVideo = [x for x in rutaVideo if str(x) != 'nan']

# Crear una lista de los valores de la columna `RutaDocumento` en el dataframe `datosEnvio`.
rutaDocumento = list(datosEnvio.RutaDocumento)
rutaDocumento = [x for x in rutaDocumento if str(x) != 'nan']

# Concatenando los valores de las columnas `NumeroPais` y `Contactos` y almacenando el resultado en la columna `Contactos`.
contactos["Contactos"] = contactos["NumeroPais"].astype(str) + contactos["Contactos"]

# Creando un dataframe con dos columnas, `Contactos` y `Estado`.
df = pd.DataFrame(columns=["Contactos", "Estado"])

print("\n---- ---- ---- ----")
print("Abriendo WatsApp Web")

# Creando una nueva instancia de la clase `Opciones`.
chrome_options=Options()

# Una forma de guardar la sesión.
chrome_options.add_argument("--user-data-dir2=chrome-data")

# Creación de una nueva instancia del controlador web de Chrome.
driver = webdriver.Chrome(f"Programas/chromedriver.exe",options=chrome_options)

print("\n---- ---- ---- ----")
print("Escanea el codigo QR\n")

# Apertura de la página Web de WhatsApp.
driver.get("https://web.whatsapp.com")
time.sleep(40)

# Un contador.
a = 0

# Creando una lista de los valores en la columna `Contactos` en el dataframe `contactos`.
cel = contactos["Contactos"]

# Abrir un chat con cada contacto de la lista.
for i in cel:
    print("\n______________________")
    print(f"Creando y Abriendo chat con: {i}")
    link = (f"https://web.whatsapp.com/send?phone={i}")
    driver.get(link)
    time.sleep(30)

    a=a+1
    df.at[a,"Contactos"] = i

    try:
        for j in mensaje:
            print("Escribiendo mensaje")
            input_xpath = '//*[@id="main"]/footer/div[1]/div/div/div[2]/div[1]/div/div[2]'
            input_box = WebDriverWait(driver,40).until(lambda driver: driver.find_element_by_xpath("xpath",input_xpath))
            time.sleep(10)
            input_box.send_keys(j + Keys.ENTER)
            time.sleep(15)
            print("Mensaje enviado")
            df.at[a, 'Estado'] ='Mensaje enviado satisfactoriamente'
            
        if rutaImagen:
            for k in rutaImagen:
                print('Enviando imagen')
                attach_button_xpath = '//div[@title = "Adjuntar"]'
                attach_button = WebDriverWait(driver,20).until(lambda driver: driver.find_element_by_xpath("xpath",attach_button_xpath))
                time.sleep(2)
                attach_button.click()
                image_box_xpath = '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'
                image_box = WebDriverWait(driver,20).until(lambda driver: driver.find_element_by_xpath("xpath",image_box_xpath))
                image_box.send_keys(k)
                time.sleep(3)
                send_button = WebDriverWait(driver,40).until(lambda driver: driver.find_element_by_xpath("xpath",'//*[@id="app"]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/div/div[2]/div[2]/div/div/span'))
                send_button.click()
                time.sleep(8)
                print('Imagen enviada')

        if rutaVideo:
            for k in rutaVideo:
                print('Enviando video')
                attach_button_xpath = '//div[@title = "Adjuntar"]'
                attach_button = WebDriverWait(driver,20).until(lambda driver: driver.find_element_by_xpath("xpath",attach_button_xpath))
                time.sleep(4)
                attach_button.click()
                image_box_xpath = '//input[@accept="image/*,video/mp4,video/3gpp,video/quicktime"]'
                image_box = WebDriverWait(driver,20).until(lambda driver: driver.find_element_by_xpath("xpath",image_box_xpath))
                image_box.send_keys(k)
                time.sleep(4)
                send_button = WebDriverWait(driver,40).until(lambda driver: driver.find_element_by_xpath("xpath",'//*[@id="app"]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/div/div[2]/div[2]/div/div/span'))
                send_button.click()
                time.sleep(20)
                print('Video enviado')
                df.at[a, 'estado'] ='Mensaje enviado satisfactoriamente'
        if rutaDocumento:
            for k in rutaDocumento:
                print('Enviando documento')
                attach_button_xpath = '//div[@title = "Adjuntar"]'
                attach_button = WebDriverWait(driver,20).until(lambda driver: driver.find_element_by_xpath("xpath",attach_button_xpath))
                time.sleep(4)
                attach_button.click()
                doc_box_xpath = '//input[@accept="*"]'
                doc_box = WebDriverWait(driver,20).until(lambda driver: driver.find_element_by_xpath("xpath",doc_box_xpath))
                doc_box.send_keys(k)
                time.sleep(4)
                send_button = WebDriverWait(driver,40).until(lambda driver: driver.find_element_by_xpath("xpath",'//*[@id="app"]/div[1]/div[1]/div[2]/div[2]/span/div[1]/span/div[1]/div/div[2]/div/div[2]/div[2]/div/div/span'))
                send_button.click()
                time.sleep(20)
                print('Documento enviado')
                df.at[a, 'estado'] ='Mensaje enviado satisfactoriamente'      
                
    except Exception as e:
        print(e)
        print(f"Este numero no tiene WhatsApp: {i}")
        df.at[f"{a} Estado"] = "Numero sin WhatsApp"
        print(e)
        print('-----------------------------------------------')

from datetime import datetime
fechahoy=datetime.now()
FFH=fechahoy.strftime("%Y%m%d_%H%M")
nombre=FFH+'_reporte'
df.to_excel('{}.xlsx'.format(nombre),index=False)
print('Reporte exportado') 
print('Cerrando programa')
driver.quit()