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
hora, minutos = input('Selecione o horario que gostaria: ').split(':')
hora = int(hora)
minutos = int(minutos)

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

disponivel = None
while disponivel == None:
    disponivel = pt.locateOnScreen('C:\\Users\\labreu\\Desktop\\dataDisponivel.png', confidence=0.8)
    driver.find_element(By.XPATH, '/html/body/div[5]/div[2]/div/a/span').click()
    driver.find_element(By.XPATH, '/html/body/div[5]/div[2]/table/thead/tr/th[1]').click()
    time.sleep(2)

mesTabela = meses[driver.find_element(By.XPATH, '/html/body/div[5]/div[1]/div/div/span[1]').text]
anoTabela = int(driver.find_element(By.XPATH, '/html/body/div[5]/div[1]/div/div/span[2]').text)
print('Ve mes e ano Tabela 1')
