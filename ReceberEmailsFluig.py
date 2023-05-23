##IMPORTANDO BIBLIOTECAS##
import pyautogui 
import time
import os
import smtplib
from selenium import *
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import sys
from email.message import EmailMessage
import pandas


arquivo = 'contas.xlsx'
planilha = pandas.read_excel(arquivo)

for index, row in planilha.iterrows():

    print(row["USUARIO"])
    print(row["SENHA"])
    print(row["email"])

    def converter(palavra):
        palavra_certa = ""
        for i in palavra:
            try:
                int(i)
                palavra_certa += i
            except ValueError:
                None
        return palavra_certa
            

    ##CONFIGURANDO JANELA##
    # os.system('pip install --upgrade webdriver-manager')
    servico = Service(ChromeDriverManager().install())
    chrome_options = Options()
    chrome_options.add_experimental_option("detach", True)
    chrome_options.add_argument("--headless")

    chrome_options.add_argument("--window-size=1920,1080")
    janela = webdriver.Chrome(service=servico, options=chrome_options)

    ##LOGANDO NO FLUIG##

    janela.get('https://fluig.redeinspiraeducadores.com.br/portal/p/01/home')
    action = webdriver.ActionChains(janela)
    logar = janela.find_element('xpath', '//*[@id="username"]').send_keys(row["USUARIO"])

    senha = janela.find_element('xpath', '//*[@id="password"]').send_keys(row["SENHA"])
    apertar = janela.find_element('xpath', '//*[@id="submitLogin"]').send_keys(Keys.ENTER)

    # pyautogui.hotkey('enter')
    time.sleep(2)

    while True:
        conteudotarefas = janela.find_element(By.CLASS_NAME, 'taskChart-legend')
        if conteudotarefas != None:
            verificar = conteudotarefas.text
            break

    email = 'automacao@redeinspiraeducadores.com.br'
    senha = 'hjiwfssnejuugumg'


    mensagemtotal = verificar
    mensagemtarefaatrasadatexto = "Tarefas Atrasadas"
    mensagemtarefaatrasadanumero = int((verificar[0:1]))

    mensagemtarefaprazotexto = "Tarefas no prazo"


    mensagemtarefapendentetexto = "Aprovação de documentos pendentes"


    mensagemtarefaaprovacaotexto = "Documentos aguardando aprovação"

    mensagemtarefasolicitacaotexto = "Minhas Solicitações"


    mensagemtarefaatrasadanumero = converter(verificar[0:2])


    mensagemtarefaprazonumero = converter(verificar[17:23])


    mensagemtarefapendentenumero = converter(verificar[33:43])


    mensagemtarefaaprovacaonumero = converter(verificar[68:80])

    mensagemtarefasolicitacaonumero = converter(verificar[96:116])

    # ##CRIAÇÃO DOS E-MAILS##

    msg0= EmailMessage()
    msg0['Subject'] = 'Atualização de tarefas pendentes do Fluig'
    msg0['From'] = 'automacao@redeinspiraeducadores.com.br'
    msg0['To'] = row["email"]
    msg0.set_content("Resumo de tarefas do Fluig: \n{}:{}\n{}:{}\n{}:{}\n{}:{}\n{}:{}".format(mensagemtarefaatrasadatexto,mensagemtarefaatrasadanumero,
    mensagemtarefaprazotexto,mensagemtarefaprazonumero,mensagemtarefapendentetexto,mensagemtarefapendentenumero,mensagemtarefaaprovacaotexto,
    mensagemtarefaaprovacaonumero, mensagemtarefasolicitacaotexto,mensagemtarefasolicitacaonumero))

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(email, senha)
        smtp.send_message(msg0)

    janela.close()
    janela.quit()


