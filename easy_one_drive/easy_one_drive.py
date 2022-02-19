import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import win32com.client as comclt

class easy_one_drive:
    def __init__(self,email,PASSWORD,driver_path='',chromedriver=''):
        self.email = email
        self.PASSWORD = PASSWORD
        if driver_path!='':
            self.driver_path = driver_path
            options = webdriver.ChromeOptions()
            options.add_argument("--start-maximized")
            self.driver = webdriver.Chrome(self.driver_path, chrome_options=options)
            self.chromedriver = chromedriver
        else:
            self.driver_path = driver_path
            self.driver = chromedriver
    
    def logging_in(self):
        print('Login into account',self.email)
        try:
            self.driver.get("https://onedrive.live.com/about/es-es/signin/")
            time.sleep(10)
            self.driver.switch_to.frame(1)
            time.sleep(10)
            self.driver.find_element_by_class_name('form-control').send_keys(self.email)
            self.driver.find_element_by_class_name('form-control').send_keys(Keys.ENTER)
            time.sleep(10)
            self.driver.find_element(By.XPATH, "//input[contains(@type,'password')]").send_keys(self.PASSWORD)
            time.sleep(10)
            self.driver.find_element(By.ID, "idSIButton9").click()
            time.sleep(5)
            print('Logged in!')
            print('Tip: you can change the uploading or downloading waiting time. Default time: 300 seconds')
        except:
            self.driver.quit()

    def upload_file(self,archivo,espera=90):
        try:
            print("Uploading file: ",archivo)
            self.driver.refresh()
            time.sleep(20)
            # Subir archivos    
            self.driver.find_element(By.XPATH, "//i[contains(@data-icon-name,'upload')]").click()
            time.sleep(2)
            self.driver.find_element(By.XPATH, "//button[contains(@name, 'Archivos')]").click()
            time.sleep(5)
            windowsShell = comclt.Dispatch("WScript.Shell")
            windowsShell.SendKeys("{ESC}")
            self.driver.find_element_by_xpath("//div[@class='ms-ItemsView-backgroundTaskRunner']/input[@type='file']").send_keys(archivo)
            time.sleep(3)
            try:
                self.driver.find_element_by_xpath("//button[contains(@class,'od-Button OperationMonitor-itemButtonAction')]").click()
            except:
                pass
            time.sleep(espera)
            print("File uploaded successfully")
        except:
            print('Error Uploading File')
            pass
            
    def delete_file(self, name, pat, espera=300):
        try:
            self.driver.find_element(By.XPATH, "//div[contains(@class, 'inline-block')]/input[contains(@id, 'idBtn_Back')]").click()
            time.sleep(10)
            self.driver.find_element(By.XPATH, '//button[contains(@title, "'+str(name)+'")]').click()      
            time.sleep(10)
            cajas = []
            if isinstance(pat,list):
                for i in range(0,len(pat)):
                    caja = self.driver.find_element(By.XPATH,'//div[contains(@role, "gridcell")]/div[contains(@title,"'+str(pat[i])+'")]')
                    cajas.append(caja)
                print(cajas)
                for i in range(0,len(cajas)):
                    cajas[i].click()
            else:
                cajas = self.driver.find_elements(By.XPATH,'//div[contains(@role, "gridcell")]/div[contains(@title,"'+str(pat)+'")]')
                cajas[0].click()
            self.driver.find_element(By.XPATH, "//i[contains(@data-icon-name,'delete')]").click()
            time.sleep(5)
            self.driver.find_element_by_xpath('//button[contains(@data-automationid, "confirmbutton")]').click()
            time.sleep(espera)
            print("Delete completed")
        except Exception:
            print("File not found", Exception)
            pass
        
    def logging_out(self):
        print("Logging out")
        try:
            self.driver.find_element(By.CLASS_NAME, "_2KqWkae0FcyhdNhWQ-Cp-M").click()
            time.sleep(2)
            self.driver.find_element(By.ID, "mectrl_body_signOut").click()
            time.sleep(5)
            self.driver.quit()
            time.sleep(3)
        except:
            pass