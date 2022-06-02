from tracemalloc import stop
from unittest import expectedFailure
from selenium import webdriver
import pandas as pd
from h11 import Data
import time
import os
from inspect import classify_class_attrs
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Inches
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from datetime import datetime
from selenium.webdriver.support.ui import Select
import numpy as np
from optparse import Values
from docxtpl import DocxTemplate
from openpyxl import workbook, load_workbook
import os
from os import link
from sqlite3 import adapt
path = os.path.curdir

#------------------------------------------CLICA EM ENTRAR
def botaoEntarSisReg(driver):
    time.sleep(40)
    try:
        user = driver.find_element_by_name('entrar')
        user.click()
    except:
        botaoEntarSisReg(driver)
path = os.path.curdir
continuar = True

#-------------------------------------------ATIVA O SELENIUM
driver = webdriver.Chrome(executable_path="%s/chrome/chromedriver.exe" % os.path.curdir)
driver.get("http://sisregiii.saude.gov.br/")
#--------------------------------------------LOGA
time.sleep(1)
user = driver.find_element_by_id('usuario')
user.send_keys("jplimar")
user = driver.find_element_by_id('senha')
user.send_keys('102030')
botaoEntarSisReg(driver)
time.sleep(1)
#-------------------------------------------ACESSA SISEG
time.sleep(1)
user = driver.find_element(By.XPATH, '//*[contains(text(),"consulta amb")]')
user.click()
user = driver.find_element(By.XPATH, '//*[contains(text(),"Autorizações")]')
user.click() 
user = driver.find_element(By.XPATH, '//*[contains(text(),"Autorizações")]')#acessa menu
user.click()
frame = driver.find_element_by_xpath('/html/body/div/div/div[4]/div[1]/iframe')#acessa o frame
driver.switch_to.frame(frame)
#---------------------------------------------ACESSA O DATA_FRAME

Data_frame = pd.read_excel("AUTOMATICO.xlsx")
Solicitantes = Data_frame['Cod']
Motivo = Data_frame['CANCELAMENTO EM TODAS AS VAGAS!!']
Unidade = Data_frame['Unidade Solicitante:']
Adm = Data_frame['Nome do Funcionário Solicitante']

#-------------------------------------------SCRIPT MAIN
def Teste(item):

    user = driver.find_element(By.NAME,'co_solic')
    user = driver.find_element(By.XPATH,'/html[1]/body[1]/center[1]/div[2]/form[1]/center[1]/table[1]/tbody[1]/tr[2]/td[2]/input[1]').send_keys(item[0])
    time.sleep(1)
        
    user = driver.find_element(By.XPATH,'/html/body/center/div[2]/form/center/table/tbody/tr[5]/td/input[1]')      
    user.click()
        
    time.sleep(1)
            
    user = driver.find_element(By.NAME,'limpar_cmp')
    user.click()

#-------------------------------------------SCRIPT MAIN
def CancelarAgend():
    
    Solicitantes_list = Data_frame['Cod'].to_list()

    Motivo_list = Data_frame['CANCELAMENTO EM TODAS AS VAGAS!!'].to_list()
    Pat = zip(Solicitantes_list, Motivo_list)
    Dict = dict(Pat)


    Lista = list(Dict.items())

    for item in Lista:
    
        
                Teste(item)
        
                #-----------------------------------INSERE MOTIVO E CANCELA DEFINITIVAMENTE   
                try:
                    
                    box = driver.find_element(By.XPATH,'/html/body/center/div[2]/form/center/table[3]/tbody/tr[1]/td/table/tbody/tr[1]/td/input') #checkbox
                    box.click()

        
                    calncelar = driver.find_element(By.XPATH,'/html/body/center/div[2]/form/center/table[3]/tbody/tr[1]/td/table/tbody/tr[3]/td/textarea').send_keys(item[1]) #Motivo
                    print("\nMotivo")
                    time.sleep(1)

                    try:
                        print("\nCancelando") 
                        user = driver.find_element(By.XPATH,'/html/body/center/div[2]/form/center/table[3]/tbody/tr[1]/td/table/tbody/tr[4]/td/input')  #calcela    
                        time.sleep(1)         
                        print("\nCLK")
                        
                        try:
                            user.click()
                            print("\nCLK")
                        except:
                            continue
                        print("\nCancelado")
                        
                        time.sleep(1)
                        frame = driver.find_element_by_xpath('/html/body/div/div/div[4]/div[1]/iframe')
                        
                        
                        box = driver.find_element(By.XPATH,'/html/body/center/div[2]/form/center/table[3]/tbody/tr[1]/td/table/tbody/tr[1]/td/input') #checkbox
                        box.click()
                        
                        time.sleep(1)
                        frame = driver.find_element_by_xpath('/html/body/div/div/div[4]/div[1]/iframe')
                                     

                    
                        user = driver.find_element(By.NAME,'limpar_cmp')
                        user.click()
                    except:
                        continue
                    

                #---------------------------------------------------------------------------
                except:
                    continue
                else:
                    continue
        
                    print("\n\nPassou")
             
    driver.close()
    
    
def Export():
     
    try:
            pd.set_option('display.max_columns', 500)
            doc = DocxTemplate('Export.docx')
            #print(Data_frame)

            print("\nImportando")
            from docx import Document


            Data_frame = pd.read_excel("AUTOMATICO.xlsx")

            data_e_hora_atuais = datetime.now()
            data_e_hora_em_texto = data_e_hora_atuais.strftime("%d/%m/%Ys")
            dth= data_e_hora_atuais.strftime('%d_%m_%Y_%H_%M')


            Loc = 'AgendaCancelada_'+ dth +'.xlsx'

            Data_frame.to_excel( Loc )

    except:
        print("\n\nERRO22")
        pass  
       
CancelarAgend()
Export()