import pandas as pd
from datetime import datetime, timedelta
import glob
import os
import win32com.client
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
import time
import os
import pyautogui as pt

date_now = datetime.now()
diaanterior = date_now - timedelta(1)
diaanterior = diaanterior.strftime('%d/%m/%Y')

meses = {'January' : 1, 'February' : 2, 'March' : 3, 'April' : 4, 'May' : 5, 'June' : 6, 'July' : 7, 'August' : 8, 'September' : 9, 'October' : 10, 'November' : 11, 'December' : 12}

dataMaxima = datetime.strptime(input('Insira uma data:  (formato: yyyy/mm/dd)'), '%Y-%m-%d')
estado = input('Selecione o Estado que gostaria de realizar a consulta de data: ')
# hora, minutos = input('Selecione o horario que gostaria: ').split(':')
# hora = int(hora)
# minutos = int(minutos)

chrome_options = Options()
# chrome_options.add_argument('--headless')
# chrome_options.add_argument('--no-sandbox')
# chrome_options.add_argument('--disable-dev-shm-usage')

# chrome_options.add_experimental_option("prefs", {"download.default_directory": r"Q:\Risco\Novo Risco\pythonrisco\Codigos\data\ValidacaoINOA"})
driver = webdriver.Chrome(r"Q://Risco//Novo Risco//pythonrisco//Codigos//ChromeDriver//chromedriver.exe", chrome_options=chrome_options)
driver.implicitly_wait(5)
driver.maximize_window()

driver.get('https://ais.usvisa-info.com/pt-br/niv/users/sign_in')

user = driver.find_element(By.XPATH, '//*[@id="user_email"]').send_keys('mateus.fc.silva@gmail.com')
senha = driver.find_element(By.XPATH, '//*[@id="user_password"]').send_keys('Mudar@123')

driver.find_element(By.XPATH, '//*[@id="sign_in_form"]/div[3]/label/div').click()

driver.find_element(By.XPATH, '//*[@id="sign_in_form"]/p[1]/input').click()
time.sleep(6)

driver.find_element(By.XPATH, '/html/body/div[4]/main/div[2]/div[3]/div[1]/div/div[1]/div[1]/div[2]/ul/li/a').click()


driver.find_element(By.XPATH, '/html/body/div[4]/main/div[2]/div[2]/div/section/ul/li[4]/a').click()
time.sleep(1)
driver.find_element(By.XPATH, '/html/body/div[4]/main/div[2]/div[2]/div/section/ul/li[4]/div/div/div[2]/p[2]/a').click()
time.sleep(2)


SelectUF = driver.find_element(By.XPATH, '//*[@id="appointments_consulate_appointment_facility_id"]').click()
SelectUF = driver.find_elements(By.CSS_SELECTOR, '#appointments_consulate_appointment_facility_id > *')
for uf in SelectUF:
    distrito = uf.text
    if distrito == estado:
        uf.click()

agendamento = driver.find_element(By.XPATH, '//*[@id="appointments_consulate_appointment_date"]').click()


mesTabela = meses[driver.find_element(By.XPATH, '/html/body/div[5]/div[1]/div/div/span[1]').text]
anoTabela = int(driver.find_element(By.XPATH, '/html/body/div[5]/div[1]/div/div/span[2]').text)
print('Ve mes e ano Tabela 1')

if (mesTabela <= dataMaxima.month and anoTabela == dataMaxima.year) or (anoTabela < dataMaxima.year):
    print('Procura dias Tabela 1')
    div1 = driver.find_element(By.XPATH, '/html/body/div[5]/div[1]')
    a = div1.find_elements(By.CSS_SELECTOR, 'a.ui-state-default')
    if len(a) == 0:
                                    ## INICIO DO LOOP ##
        while True:
            print('Ve mes e ano Tabela 2')

            mesTabela = meses[driver.find_element(By.XPATH, '/html/body/div[5]/div[2]/div/div/span[1]').text]
            anoTabela = int(driver.find_element(By.XPATH, '/html/body/div[5]/div[2]/div/div/span[2]').text)

            if (mesTabela <= dataMaxima.month and anoTabela == dataMaxima.year) or (anoTabela < dataMaxima.year):
                print('Procura dias Tabela 2')
                div2= driver.find_element(By.XPATH, '/html/body/div[5]/div[2]')
                a = div2.find_elements(By.CSS_SELECTOR, 'a.ui-state-default')
                
                if len(a) == 0:
                    print('N??o foi possivel encontrar um agendamento com a data limite solicitada')
                    driver.find_element(By.XPATH, '//*[@id="ui-datepicker-div"]/div[2]/div/a').click()
                    continue
                else:
                    for dia in a:
                        if (int(dia.text) <= dataMaxima.day and mesTabela == dataMaxima.month) or (mesTabela < dataMaxima.month):
                            print(f'ACHEI A DATA {dia.text}/{mesTabela}/{anoTabela}')
            else:
                print('N??o foi possivel encontrar um agendamento com a data limite solicitada')
            break
    else:
        for dia in a:
            if (int(dia.text) <= dataMaxima.day and mesTabela >= dataMaxima.month) or (mesTabela < dataMaxima.month):
                print(f'ACHEI A DATA {dia.text}/{mesTabela}/{anoTabela}')
else:
    print('N??o foi possivel encontrar um agendamento com a data limite solicitada')

try:
    horDisponiveis = driver.find_element(By.XPATH, '/html/body/div[4]/main/div[4]/div/div/form/fieldset[1]/ol/fieldset/div/div[2]/div[3]/li[2]/select').click()
    horDisponiveis = driver.find_elements(By.CSS_SELECTOR, '#appointments_consulate_appointment_time > *')
    for horas in horDisponiveis:
        tempo = horas.text
        print(tempo)
    if tempo == horario:
        horas.click()
except:
    
