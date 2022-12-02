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


date_now = datetime.now()
diaanterior = date_now - timedelta(1)
diaanterior = diaanterior.strftime('%d/%m/%Y')


chrome_options = Options()
# chrome_options.add_argument('--headless')
# chrome_options.add_argument('--no-sandbox')
# chrome_options.add_argument('--disable-dev-shm-usage')

chrome_options.add_experimental_option("prefs", {"download.default_directory": r"Q:\Risco\Novo Risco\pythonrisco\Codigos\data\ValidacaoINOA"})
driver = webdriver.Chrome("Q://Risco//Novo Risco//pythonrisco//Codigos//ChromeDriver//chromedriver.exe", chrome_options=chrome_options)
driver.implicitly_wait(5)
driver.maximize_window()

driver.get('https://ais.usvisa-info.com/pt-br/niv/users/sign_in')

user = driver.find_element(By.XPATH, '//*[@id="user_email"]').send_keys('mateus.fc.silva@gmail.com')
senha = driver.find_element(By.XPATH, '//*[@id="user_password"]').send_keys('Mudar@123')

driver.find_element(By.XPATH, '//*[@id="sign_in_form"]/div[3]/label/div').click()

driver.find_element(By.XPATH, '//*[@id="sign_in_form"]/p[1]/input').click()
time.sleep(3)

driver.find_element(By.XPATH, '//*[@id="main"]/div[2]/div[3]/div[1]/div/div[1]/div[1]/div[2]/ul/li/a').click()


driver.find_element(By.XPATH, '//*[@id="tvktw8-accordion-label"]').click()
time.sleep(1)
driver.find_element(By.XPATH, '//*[@id="tvktw8-accordion"]/div/div[2]/p[2]/a').click()
time.sleep(2)

agendamento = driver.find_element(By.XPATH, '//*[@id="appointments_consulate_appointment_date"]').click()
agendamento = driver.find_element(By.XPATH, '//*[@id="appointments_consulate_appointment_date"]').get_attribute('value')

agendamento = datetime.strptime(agendamento, '%Y-%m-%d')
dataselecionada = datetime.strptime(input('Insira uma data:  (formato: yyyy/mm/dd)'), '%Y-%m-%d')

while True:
    try:
        driver.find_element(By.XPATH, '//*[@id="ui-datepicker-div"]/div[1]')
        a = driver.find_elements(By.CSS_SELECTOR, 'a.ui-state-default')
        for i in a:
            i = i.text
            print(i)
    except:
        try:
            driver.find_element(By.XPATH, '//*[@id="ui-datepicker-div"]/div[2]')
            a = driver.find_elements(By.CSS_SELECTOR, 'a.ui-state-default')

            for i in a:
                i = i.text
                print(i)
        except:
            driver.find_element(By.XPATH, '//*[@id="ui-datepicker-div"]/div[2]/div/a').click()
            
    if dataselecionada > agendamento:
        print('show')
    else:
        print('Nok')