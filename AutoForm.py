# -*- coding: ISO-8859-1 -*-
import configparser
import os.path
import subprocess
import sys

import pymsgbox
from selenium import webdriver
from selenium.webdriver.common.by import By
from win32api import GetSystemMetrics

cfg = configparser.ConfigParser()
cfg.read('CONFIG.ini', encoding='utf-8')

chromedriver = cfg.get('config','driver')
nome = cfg.get('config', 'nome')
cpf = cfg.get('config', 'cpf')
rg = cfg.get('config', 'rg')
genero = cfg.get('config', 'genero')
orientacao = cfg.get('config', 'orientacao')
estadocivil = cfg.get('config', 'estadocivil')
foto = cfg.get('config', 'foto')
curriculo = cfg.get('config', 'curriculo')
email = cfg.get('config', 'email')
senha = cfg.get('config', 'senha')
nascimento = cfg.get('config', 'nascimento')
tel = cfg.get('config', 'tel')
whats = cfg.get('config', 'whats')
pai = cfg.get('config', 'pai')
profpai = cfg.get('config', 'profpai')
mae = cfg.get('config', 'mae')
profmae = cfg.get(u'config', 'profmae')
cep = cfg.get('config', 'cep')
numero = cfg.get('config', 'numero')
comp = cfg.get('config', 'comp')
estado_natur = cfg.get('config', 'estado_natur')
cidade_natur = cfg.get('config', 'cidade_natur')
raca = cfg.get('config', 'raca')
face = cfg.get('config', 'face')
linkedin = cfg.get('config', 'linkedin')
github = cfg.get('config', 'github')
escola = cfg.get('config', 'escola')
situacao = cfg.get('config', 'situacao')
estudando = cfg.get('config', 'estudando')
turno = cfg.get('config', 'turno')
escolaEnt = cfg.get('config', 'escolaEnt')
idioma = cfg.get('config', 'idioma')
nivel = cfg.get('config', 'nivel')
obsIdioma = cfg.get('config', 'obsIdioma')
exp = cfg.get('config', 'exp')
cursos = cfg.get('config', 'cursos')
js = cfg.get('config', 'js')
jsExp = cfg.get('config', 'jsExp')
java = cfg.get('config', 'java')
javaExp = cfg.get('config', 'javaExp')
python = cfg.get('config', 'python')
pythonExp = cfg.get('config', 'pythonExp')
go = cfg.get('config', 'go')
goExp = cfg.get('config', 'goExp')
delphi = cfg.get('config', 'delphi')
delphiExp = cfg.get('config', 'delphiExp')
c = cfg.get('config', 'c')
cExp = cfg.get('config', 'cExp')
cpp = cfg.get('config', 'cpp')
cppExp = cfg.get('config', 'cppExp')
csharp = cfg.get('config', 'csharp')
csharpExp = cfg.get('config', 'csharpExp')
r = cfg.get('config', 'r')
rExp = cfg.get('config', 'rExp')
php = cfg.get('config', 'php')
phpExp = cfg.get('config', 'phpExp')
rust = cfg.get('config', 'rust')
rustExp = cfg.get('config', 'rustExp')
kotlin = cfg.get('config', 'kotlin')
kotlinExp = cfg.get('config', 'kotlinExp')
swift = cfg.get('config', 'swift')
swiftExp = cfg.get('config', 'swiftExp')
sabendo = cfg.get('config', 'sabendo')
motivo = cfg.get('config', 'motivo')

def Acessar():
    print('Abriu o Chrome...')

    navegador.get('https://t-systems.proway.com.br/inscricao/')
    print('Abriu o site...')
    navegador.set_window_size(width=GetSystemMetrics(0) / 1.3, height=GetSystemMetrics(1) / 1.05)
    print('Alterou o tamanho da janela...')

    print("Preenchendo os dados...")

    # preenche o estado primeiro e a cidade por último para não ter problema com delay...
    Preencher('estadoNascimento', estado_natur, 'name')

    Preencher('nome', nome, 'name')
    Preencher('cpf', cpf, 'name')
    Preencher('rg', rg, 'name')
    Preencher('genero', genero, 'name')
    Preencher('orientacaoSexual', orientacao, 'name')
    Preencher('estadoCivil', estadocivil, 'name')
    Preencher('foto', foto, 'name')
    Preencher('file', curriculo, 'name')
    Preencher('email', email, 'name')
    Preencher('email_conf', email, 'name')
    Preencher('senha', senha, 'name')
    Preencher('senha_conf', senha, 'name')
    Preencher('dataNascimento', nascimento, 'name')
    Preencher('telefone', tel, 'name')
    Preencher('celular', whats, 'name')
    Preencher('nomepai', pai, 'name')
    Preencher('profissaopai', profpai, 'name')
    Preencher('nomemae', mae, 'name')
    Preencher('profissaomae', profmae, 'name')
    Preencher('cep', cep, 'name')
    Preencher('numero', numero, 'name')
    Preencher('complemento', comp, 'name')
    Preencher('raca', raca, 'name')
    Preencher('facebook', face, 'name')
    Preencher('linkedin', linkedin, 'name')
    Preencher('github', github, 'name')
    Preencher('escolaridade', escola, 'name')
    Preencher('situacaoEstudo', situacao, 'name')

    if estudando == 'Sim':
        Clicar('/html/body/div/div/div/form/div/div[5]/div[1]/div[3]/div[2]/button[2]', 'xpath')
        Preencher('escolaturno', turno, 'name')
        Preencher('escola', escolaEnt, 'name')

    Preencher('idioma', idioma, 'name')
    Preencher('idiomaNivel', nivel, 'name')
    Preencher('/html/body/div/div/div/form/div/div[5]/div[3]/div[3]/input', obsIdioma, 'xpath')
    Preencher('empregosAnteriores', exp, 'name')
    Preencher('outrosCursosAperfeicoamento', cursos, 'name')

    if js == "Sim":
        Clicar('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[1]/label/input', 'xpath')
        navegador.implicitly_wait(5)
        Preencher('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[1]/select', jsExp, 'xpath')

    if java == "Sim":
        Clicar('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[2]/label/input', 'xpath')
        navegador.implicitly_wait(5)
        Preencher('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[2]/select', javaExp, 'xpath')

    if python == "Sim":
        Clicar('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[3]/label/input', 'xpath')
        navegador.implicitly_wait(5)
        Preencher('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[3]/select', pythonExp, 'xpath')

    if go == "Sim":
        Clicar('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[4]/label/input',  'xpath')
        navegador.implicitly_wait(5)
        Preencher('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[4]/select', goExp,  'xpath')

    if delphi == "Sim":
        Clicar('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[5]/label/input', 'xpath')
        navegador.implicitly_wait(5)
        Preencher('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[5]/select', delphiExp, 'xpath')

    if c == "Sim":
        Clicar('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[6]/label/input', 'xpath')
        navegador.implicitly_wait(5)
        Preencher('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[6]/select', cExp, 'xpath')

    if cpp == "Sim":
        Clicar('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[7]/label/input', 'xpath')
        navegador.implicitly_wait(5)
        Preencher('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[7]/select', cppExp, 'xpath')

    if csharp == "Sim":
        Clicar('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[8]/label/input', 'xpath')
        navegador.implicitly_wait(5)
        Preencher('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[8]/select', csharpExp, 'xpath')

    if r == "Sim":
        Clicar('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[9]/label/input', 'xpath')
        navegador.implicitly_wait(5)
        Preencher('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[9]/select', rExp, 'xpath')

    if php == "Sim":
        Clicar('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[10]/label/input', 'xpath')
        navegador.implicitly_wait(5)
        Preencher('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[10]/select', phpExp, 'xpath')

    if rust == "Sim":
        Clicar('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[11]/label/input', 'xpath')
        navegador.implicitly_wait(5)
        Preencher('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[11]/select', rustExp, 'xpath')

    if kotlin == "Sim":
        Clicar('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[12]/label/input', 'xpath')
        navegador.implicitly_wait(5)
        Preencher('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[12]/select', kotlinExp, 'xpath')

    if swift == "Sim":
        Clicar('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[13]/label/input', 'xpath')
        navegador.implicitly_wait(5)
        Preencher('/html/body/div/div/div/form/div/div[6]/div[3]/div/div/div/div[13]/select', swiftExp, 'xpath')

    Preencher('comoFicouSabendo', sabendo, 'name')
    Preencher('porqueQuerSerSelecionado', motivo, 'name')

    Preencher('codCidadeNascimento', cidade_natur, 'name')

    Clicar('maior', 'name')

    pymsgbox.alert("Processo concluído. Verifique se todos os dados foram preenchidos corretamente e se o envio ocorreu normalmente e clique em OK para finalizar o navegador")
    navegador.close()

def Preencher(campo, form, type):
    msgOK = f'Preencheu o campo "{campo}"'
    msgFalha = f'Falha ao preencher o campo "{campo}"...'
    try:
        if type == 'name':
            navegador.find_element(By.NAME, campo).send_keys(form)
        if type == 'xpath':
            navegador.find_element(By.XPATH, campo).send_keys(form)
        print(msgOK)
    except:
        print(msgFalha)

def Clicar(campo, type):
    msgOK = f'Clicou no campo "{campo}"'
    msgFalha = f'Falha ao clicar no campo "{campo}"...'
    try:
        if type == 'name':
            navegador.find_element(By.NAME, campo).click()
        elif type == 'xpath':
            navegador.find_element(By.XPATH, campo).click()
        print(msgOK)
    except:
        print(msgFalha)

if os.path.exists(r'C:\\Program Files (x86)\\Google\\Chrome\\Application\\'):
    output = subprocess.check_output(
    r'wmic datafile where name="C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe" get Version /value',
    shell=True
)
    chrome_version = output.decode('utf-8').split('.')
    chrome_version = chrome_version[0].replace('\n', '')
    chrome_version = chrome_version.replace('Version=', '')
    chrome_version = int(chrome_version)

    if chrome_version == 103:
        s = r'bin\chromedriver_103.exe'
        print('Google Chrome installed in version 103')
        print('Google Chrome instalado na versão 103')
        navegador = webdriver.Chrome(s)
        Acessar()
    elif chrome_version == 104:
        s = r'bin\chromedriver_104.exe'
        print('Google Chrome installed in version 104')
        print('Google Chrome instalado na versão 104')
        navegador = webdriver.Chrome(s)
        Acessar()
    elif chrome_version == 106:
        s = r'bin\chromedriver_106.exe'
        print('Google Chrome installed in version 106')
        print('Google Chrome instalado na versão 106')
        navegador = webdriver.Chrome(s)
        Acessar()
        sys.exit(0)
    elif chrome_version == 108:
        s = chromedriver
        print('Google Chrome installed in version 108')
        print('Google Chrome instalado na versão 108')
        navegador = webdriver.Chrome(s)
        Acessar()
    else:
        print('Please, install the last version of Google Chrome in your system before use this software')
        print('Por favor, instale a última versão do Google Chrome antes de usar esse software')

else:
    print('Google chrome não instalado na pasta padrão...')