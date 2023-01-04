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

meses = {'January' : '01', 'February' : '02', 'March' : '03', 
'April' : '04', 'May' : '05', 'June' : '06', 
'July' : '07', 'August' : '08', 'September' : '09', 
'October' : '10', 'November' : '11', 'December' : '12'}

menoresMeses = ['February', 'April', 'June', 'September', 'November']

# dataMaxima = datetime.strptime(input('Insira uma data:  (formato: yyyy/mm/dd)'), '%Y-%m-%d')
estado = int(input('Selecione o Estado (1-> BRASILIA, 2-> POA, 3-> RECIFE, 4-> RIO DE JANEIRO, 5-> SAO PAULO): '))

if estado == 1:
    estado = 'Brasilia'
elif estado == 2:
    estado = 'Porto Alegre'
elif estado == 3:
    estado = 'Recife'
elif estado == 4:
    estado = 'Rio de Janeiro'
elif estado == 5:
    estado = 'Sao Paulo'

print(f'O estado selecionado foi: {estado}')
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

Bconfirma = driver.find_element(By.XPATH, '/html/body/div[4]/main/div[2]/div[3]/div[1]/div/div[1]/div[1]/div[2]/ul/li/a').click()

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
        print(f'Confirmaçao de Local: O local selecionado foi {distrito}')


dataMaxima = datetime.strptime(input('Insira uma data:  (formato: yyyy-mm-dd)'), '%Y-%m-%d')

agendamento = driver.find_element(By.XPATH, '//*[@id="appointments_consulate_appointment_date"]').click()

mesTabela = meses[driver.find_element(By.XPATH, '/html/body/div[5]/div[1]/div/div/span[1]').text]
anoTabela = driver.find_element(By.XPATH, '/html/body/div[5]/div[1]/div/div/span[2]').text
anoTabelaINT = int(driver.find_element(By.XPATH, '/html/body/div[5]/div[1]/div/div/span[2]').text)

                   ## INICIO DO LOOP ##

print('Procura dias Tabela 1')

if mesTabela in menoresMeses and menoresMeses[mesTabela] and dataMaxima.day == 31:
    if mesTabela == 'February':
        if (anoTabelaINT%4==0 and anoTabelaINT%100!=0) or (anoTabelaINT%400==0):
            diaDataMaxima = '29'
        else:
            diaDataMaxima = '28'
    else:
        diaDataMaxima = '30'
elif mesTabela == 'February' and dataMaxima.day == 30:
    if (anoTabelaINT%4==0 and anoTabelaINT%100!=0) or (anoTabelaINT%400==0):
        diaDataMaxima = '29'
    else:
        diaDataMaxima = '28'
else:
    diaDataMaxima = '0' + str(dataMaxima.day) if dataMaxima.day < 10 else str(dataMaxima.day)

if datetime.strptime(diaDataMaxima + '-' + mesTabela + '-' + anoTabela, '%d-%m-%Y') <= dataMaxima:

    div1 = driver.find_element(By.XPATH, '/html/body/div[5]/div[1]')
    a = div1.find_elements(By.CSS_SELECTOR, 'a.ui-state-default')

    if len(a) == 0:

        while True:
            print('Procura dias Tabela 2')
            mesTabela = meses[driver.find_element(By.XPATH, '/html/body/div[5]/div[2]/div/div/span[1]').text]
            anoTabela = driver.find_element(By.XPATH, '/html/body/div[5]/div[2]/div/div/span[2]').text

            if datetime.strptime(diaDataMaxima + '-' + mesTabela + '-' + anoTabela, '%d-%m-%Y') > dataMaxima:
                print('Nao foram encotrado dias disponiveis com base na data maxima (TABLE 2)')
                break

            div2= driver.find_element(By.XPATH, '/html/body/div[5]/div[2]')
            a = div2.find_elements(By.CSS_SELECTOR, 'a.ui-state-default')
            
            if len(a) == 0:
                print('Não foi possivel encontrar um agendamento com a data limite solicitada')
                driver.find_element(By.XPATH, '//*[@id="ui-datepicker-div"]/div[2]/div/a').click()
                continue
            else:
                if datetime.strptime(a[0].text + '-' + mesTabela + '-' + anoTabela, '%d-%m-%Y') > dataMaxima:
                    print('Nao foram encotrado dias disponiveis com base na data maxima (TABLE 2')
                    break
                a[0].click()
                print(f'ACHEI A DATA {a[0].text}/{mesTabela}/{anoTabela}')
            break
    else:
        if datetime.strptime(a[0].text + '-' + mesTabela + '-' + anoTabela, '%d-%m-%Y') <= dataMaxima:
            a[0].click()
            print(f'ACHEI A DATA {a[0].text}/{mesTabela}/{anoTabela}')
            data = driver.find_element(By.XPATH, '//*[@id="appointments_consulate_appointment_date"]').get_attribute('value')
            try:
                print('Encontrando o primeiro horario')
                horDisponiveis = driver.find_element(By.XPATH, '/html/body/div[4]/main/div[4]/div/div/form/fieldset[1]/ol/fieldset/div/div[2]/div[3]/li[2]/select').click()
                FirstHour = driver.find_element(By.XPATH, '//*[@id="appointments_consulate_appointment_time"]/option[2]').click()
                HoursText = driver.find_element(By.XPATH, '//*[@id="appointments_consulate_appointment_time"]/option[2]').get_attribute('text')
                print(f'Horario encontrado, sua seção sera agendada para o horario {HoursText}')
            except:
                print('Não foi possivel encontrar um primeiro horario')
                
            #outlook = win32com.client.Dispatch('outlook.application')
           # mail = outlook.CreateItem(0)
            #mail.To = "luis.abreu@ativainvestimentos.com.br; marcus.romao@ativainvestimentos.com.br; henrique.barbosa@ativainvestimentos.com.br"
            #mail.Subject = 'Plano Bronze'
            #mail.GetInspector
            # Attachments = r'Q:\Risco\Novo Risco\1 - Rotinas\Middle\Preços\filtroRemuneracao.xlsx'
            # mail.Attachments.Add(Attachments)
            # pathToIMage = 'Q://Risco/Novo Risco/pythonrisco/Codigos/data/CM_teste/Assinatura_Risco.png'
            # attachment = mail.Attachments.Add(pathToIMage)
            # attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F", "MyId1")
            #mail.HTMLBody = "<p>  Boa tarde! <p>" \
           #                 '<p> Identificamos uma data para o seu agendamento <p>' \
           #             f"<p> No estado {distrito}, na data {data} e no horario das {HoursText} <p>" \

           # '<p><p>' \
           # '<p> <figure><img src="cid:MyId1"</figure>'
          #  mail.display()
            # mail.Send()

        else:
            print('Nao foram encotrado dias disponiveis com base na data maxima (TABLE 1')
else:
    print('Nao foram encotrado dias disponiveis com base na data maxima (TABLE 1')
