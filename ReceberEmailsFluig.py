##IMPORTANDO BIBLIOTECAS##
import pyautogui 
import time
import os
import smtplib
from email.message import EmailMessage
from selenium import *
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import sys

##CONFIGURANDO JANELA##
servico = Service(ChromeDriverManager().install())
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)
chrome_options.add_argument("--headless")

chrome_options.add_argument("--window-size=1920,1080")
janela = webdriver.Chrome(service=servico, options=chrome_options)

##LOGANDO NO FLUIG##

janela.get('https://fluig.redeinspiraeducadores.com.br/portal/p/01/home')
action = webdriver.ActionChains(janela)
logar = janela.find_element('xpath', '//*[@id="username"]').send_keys("maycon.henrique")

senha = janela.find_element('xpath', '//*[@id="password"]').send_keys('1q2w3e4r')
apertar = janela.find_element('xpath', '//*[@id="submitLogin"]').send_keys(Keys.ENTER)

# pyautogui.hotkey('enter')
time.sleep(2)

while True:
    conteudotarefas = janela.find_element(By.CLASS_NAME, 'taskChart-legend')
    if conteudotarefas != None:
        verificar = conteudotarefas.text
        break

email = 'automacao@redeinspiraeducadores.com.br'
senha = 'Automacao@2022'
mensagemtotal = verificar
mensagemtarefaatrasadatexto = (verificar[1:17])
mensagemtarefaatrasadanumero = int((verificar[0:1]))

mensagemtarefaprazotexto = (verificar[20:36])
mensagemtarefaprazonumero = int((verificar[19:20]))

mensagemtarefapendentetexto = (verificar[38:62])
mensagemtarefapendentenumero = int((verificar[37:38]))

mensagemtarefaaprovacaotexto = (verificar[73:104])
mensagemtarefaaprovacaonumero = int(verificar[72:73])

mensagemtarefasolicitacaotexto = (verificar[107:125])
mensagemtarefasolicitacaonumero = int(verificar[105:107])

# ##CRIAÇÃO DOS E-MAILS##

msg0= EmailMessage()
msg0['Subject'] = 'Atualização de tarefas pendentes do Fluig'
msg0['From'] = 'automacao@redeinspiraeducadores.com.br'
msg0['To'] = 'maycon.henrique@redeinspiraeducadores.com.br'
msg0.set_content("Resumo de tarefas do Fluig: \n{}:{}\n{}:{}\n{}:{}\n{}:{}\n{}:{}".format(mensagemtarefaatrasadatexto,mensagemtarefaatrasadanumero,
mensagemtarefaprazotexto,mensagemtarefaprazonumero,mensagemtarefapendentetexto,mensagemtarefapendentenumero,mensagemtarefaaprovacaotexto,
mensagemtarefaaprovacaonumero, mensagemtarefasolicitacaotexto,mensagemtarefasolicitacaonumero))

with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login(email, senha)
    smtp.send_message(msg0)

janela.close()
janela.quit()


