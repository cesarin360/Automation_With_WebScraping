from selenium import webdriver
from selenium.webdriver.common.by import By
from onedrive import onedrive
from os import path, remove
from datetime import datetime
import time
import re
import read_and_write as csv

""" Actualiza tabla de datos utilizando Selenium """
class Aplicacion():
    def __init__(self):
        #Se lee el archivo json
        PATH = csv.read_json('config.json', 'login')
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        '''diccionario para configuración del directorio de donde se guardara los archivos
        descargados'''
        prefs = {
            "profile.default_content_settings.popups": 0,
            "download.default_directory": r"C:\Users\josue\Documents\Casos_covid\\",
            "directory_upgrade": True
        }
        options.add_experimental_option("prefs", prefs)
        driver = webdriver.Chrome(PATH['driver'], options=options)
        driver.maximize_window()
        
        #Se obtiene el link de la pagina web
        driver.get(PATH['url1'])
        delay = 60
        try:
            #Fecha de la actualización de los datos 
            dates = ''.join(re.sub(r"((/)?:|\D+)", r"-",
                            driver.find_element_by_class_name("logo").text[:26]))
            
            #Se convierte la fecha al formato YYYY-MM-DD
            date_dt = str(datetime.strptime(dates[1:], '%d-%m-%Y'))
            pathhtm = PATH['path'][:37]+'descarga.htm'
            
            #Se lee el archivo que contiene la fecha de la ultima actualización de datos
            with open('last_date.txt', 'r+') as myfile:
                data = myfile.read()
                '''Si la fehca actual no coincide con la ultima fecha ingresada
                entonces se sustituye por la actual '''
                if date_dt[:10] not in data:
                    pathfile = PATH['path']+' '+data+'.csv'
                    if path.exists(pathfile):
                        remove(pathfile)
                    myfile.seek(0)
                    myfile.write(date_dt[:10])
                    myfile.truncate()
            pathfile = PATH['path']+' '+date_dt[:10]+'.csv'
            
            #Se elimina los archivos anteriores
            if path.exists(pathfile):
                remove(pathfile)
            if path.exists(pathhtm):
                remove(pathhtm)
            time.sleep(10)
            #Se busca el xpath del elemento y se le da un click
            driver.find_element(
                By.XPATH, "//a[contains(@data-value,'descargasTab')]").click()
            time.sleep(10)
            link = driver.find_element(
                By.XPATH, "//a[contains(@id, 'fallecidosFF')]")
            time.sleep(5)
            link.click()
            time.sleep(10)
            
            '''Pasamos el path del elemento recien descargado
            para convertirlo de formato csv a xlsx y agregarlos a una tabla''' 
            csv.substitute(pathfile)
        except:
            print("ERROR 1")
            driver.quit()
            Aplicacion()
        
        finally:
            try:
                #Verificamos si el archivo comvertido se subio correctamente
                if onedrive.upload_file():
                    #os.system('TASKKILL /F /IM chrome.exe')
                    time.sleep(delay)
                    driver.quit()
                    #Iniciamos de nuevo el proceso
                    Aplicacion()
                else:
                    driver.quit()
                    Aplicacion()
            except:
                #Si ha ocurrido algún error se vuelve a ejecutar el proceso
                print("ERROR 1.1")
                driver.quit()
                Aplicacion()
def main():
    mi_app = Aplicacion()
    return 0
if __name__ == '__main__':
    main()
