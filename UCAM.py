# -*- coding: ISO-8859-1 -*-
import configparser
import os
import subprocess
import sys

import pymsgbox
from selenium import webdriver
from selenium.webdriver.common.by import By
from win32api import GetSystemMetrics

cfgD = configparser.ConfigParser()
cfgC = configparser.ConfigParser()
cfgD.read('DADOS3.ini', encoding='utf-8')
cfgC.read('CONFIG.ini', encoding='utf-8')

login = cfgD.get('dados', 'login')
senha = cfgD.get('dados', 'senha')
chromedriver = cfgC.get('config', 'driver')

def Acessar():
    print('Abriu o Chrome...')

    navegador.get('https://login.candidomendes.edu.br/login.jsf?client_id=272fac70-fca0-4ca1-b17a-ad1d1b65ff28@ucam')
    print('Abriu o site...')
    navegador.maximize_window()
    print("Preenchendo os dados...")
    Preencher('j_idt9',login,'Login','name')
    Preencher('j_idt11',senha,'Senha','name')
    Clicar('j_idt13','Entrar','name')
    navegador.implicitly_wait(20)
    Clicar('/html/body/div[3]/div/div[1]/span/div[3]/form/a[2]/div','AVA','xpath')

    pymsgbox.alert("Clique OK quando finalizar a aula")

    navegador.close()

def Preencher(campo, form, nomeCampo, type):
    msgOK = f'Preencheu o campo {nomeCampo}'
    msgFalha = f'Falha ao preencher o campo {nomeCampo}...'
    try:
        if type == 'name':
            navegador.find_element(By.NAME, campo).send_keys(form)
        if type == 'xpath':
            navegador.find_element(By.XPATH, campo).send_keys(form)
        print(msgOK)
    except:
        print(msgFalha)

def Clicar(campo, nomeCampo, type):
    msgOK = f'Clicou no botão "{nomeCampo}".'
    msgFalha = f'Falha ao clicar no botão "{nomeCampo}"...'
    try:
        if type == 'name':
            navegador.find_element(By.NAME, campo).click()
        if type == 'xpath':
            navegador.find_element(By.XPATH, campo).click()
        if type == 'class':
            navegador.find_element(By.CLASS_NAME, campo).click()
        if type == 'id':
            navegador.find_element(By.ID, campo).click()
        print(msgOK)
        return True
    except:
        print(msgFalha)
        return False

if os.path.exists(r'C:\\Program Files (x86)\\Google\\Chrome\\Application\\'):
    output = subprocess.check_output(
        r'wmic datafile where name="C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" get Version /value',
        shell=True
    )
    chrome_version = output.decode('utf-8').split('.')
    chrome_version = chrome_version[0].replace('\n', '')
    chrome_version = chrome_version.replace('Version=', '')
    chrome_version = int(chrome_version)
    print('Chrome 32 bits instalado: ', chrome_version)

    if chrome_version != '':
        s = r'' + chromedriver
        print(s)
        navegador = webdriver.Chrome()#s)
        Acessar()
    else:
        print('Please, install the last version of Google Chrome in your system before use this software')
        print('Por favor, instale a última versão do Google Chrome antes de usar esse software')
elif os.path.exists(r'C:\\Program Files\\Google\\Chrome\\Application\\'):
    output = subprocess.check_output(
        r'wmic datafile where name="C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" get Version /value',
        shell=True
    )
    chrome_version = output.decode('utf-8').split('.')
    chrome_version = chrome_version[0].replace('\n', '')
    chrome_version = chrome_version.replace('Version=', '')
    chrome_version = int(chrome_version)
    print('Chrome 64 bits instalado: ', chrome_version)

    if chrome_version != '':
        s = r'' + chromedriver
        print(s)
        navegador = webdriver.Chrome()#s)
        Acessar()
    else:
        print('Please, install the last version of Google Chrome in your system before use this software')
        print('Por favor, instale a última versão do Google Chrome antes de usar esse software')
else:
    print('Google chrome não instalado nesse sistema na pasta padrão...')
    pymsgbox.alert("Google chrome não instalado nesse sistema na pasta padrão...")
