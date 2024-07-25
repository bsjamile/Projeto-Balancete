from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import time
import os

usuario = input('Insira o usuário do AtenaPrev.Net: ')
senha = input('Insira a senha do AtenaPrev.Net: ')
referencia = '07.2024'



diretorio = f"H:\\GCO\\AFAF\\AFAF004\\GESTÃO CONTÁBIL\\Fechamento Mensal\\BALANCETES\\2024\\{referencia}\\"
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

prefs = {"download.default_directory" : diretorio}
chrome_options.add_experimental_option("prefs", prefs)

servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico, options=chrome_options)
navegador.get("http://contabilidade.neos.com/ControleAcesso/login/Login.aspx")

navegador.find_element('xpath','//*[@id="MainContent_lgnLogin_UserName"]').send_keys(usuario) #insere o login
navegador.find_element('xpath','//*[@id="MainContent_lgnLogin_Password"]').send_keys(senha) #insere a senha
navegador.find_element('xpath','//*[@id="MainContent_lgnLogin_LoginButton"]').click() #clica "entrar"

navegador.get("http://contabilidade.neos.com/Contabilidade/Default.aspx")
navegador.get("http://contabilidade.neos.com/Contabilidade")
navegador.get("http://contabilidade.neos.com/Contabilidade/Relatorios/Balancete/contabil/Pesquisar.aspx")

# Lista de opções para escolha dos balancetes
opcoes_xpath = [
    '//*[@id="MainContent_MainContent_fpuBalancete_lbtList_lbxLeft"]/option[9]',# CD BA
    '//*[@id="MainContent_MainContent_fpuBalancete_lbtList_lbxLeft"]/option[8]',# BD BA
    '//*[@id="MainContent_MainContent_fpuBalancete_lbtList_lbxLeft"]/option[7]',# CD RN
    '//*[@id="MainContent_MainContent_fpuBalancete_lbtList_lbxLeft"]/option[5]',# PGA
    '//*[@id="MainContent_MainContent_fpuBalancete_lbtList_lbxLeft"]/option[4]',# BD RN
    '//*[@id="MainContent_MainContent_fpuBalancete_lbtList_lbxLeft"]/option[3]',# CD PE
    '//*[@id="MainContent_MainContent_fpuBalancete_lbtList_lbxLeft"]/option[2]',# BD PE    
    '//*[@id="MainContent_MainContent_fpuBalancete_lbtList_lbxLeft"]/option[1]',# CD Néos
]


# Loop para escolha das opções xpath
for opcoes_xpath in opcoes_xpath:
    #Abre janela de seleção de balancete
    navegador.find_element('xpath','//*[@id="MainContent_MainContent_fpuBalancete_lnkSelecionado"]').click()
    time.sleep (0.75)

    #Limpa possíveis balancetes selecionados
    navegador.find_element('xpath','//*[@id="MainContent_MainContent_fpuBalancete_lbtList_btnRemAll"]').click() #//*[@id="MainContent_MainContent_fpuBalancete_lbtList_lbxRight"]/option
    time.sleep (0.75)

    #Seleciona balancete CD Néos
    elemento = navegador.find_element(By.XPATH, opcoes_xpath)
    acoes = ActionChains(navegador)
    acoes.double_click(elemento).perform()
    time.sleep (0.75)

    #Fecha janela de seleção de balancete
    navegador.find_element('xpath','//*[@id="MainContent_MainContent_fpuBalancete_btnCancelar"]').click()
    time.sleep (1.75)

    #Exporta balancete
    navegador.find_element('xpath','//*[@id="MainContent_MainContent_btnExportar"]').click()
    time.sleep (1.75)

#Abre janela de seleção de balancete
navegador.find_element('xpath','//*[@id="MainContent_MainContent_fpuBalancete_lnkSelecionado"]').click()
time.sleep (0.75)

#Limpa possíveis balancetes selecionados
navegador.find_element('xpath','//*[@id="MainContent_MainContent_fpuBalancete_lbtList_btnRemAll"]').click() #//*[@id="MainContent_MainContent_fpuBalancete_lbtList_lbxRight"]/option
time.sleep (0.75)

#Fecha janela de seleção de balancete
navegador.find_element('xpath','//*[@id="MainContent_MainContent_fpuBalancete_btnCancelar"]').click()
time.sleep (1.75)

opcoes_xpath_perfis = [
    'Bas CD BA',
    'Dif CD BA',
    'Cons CD BA',
    'Cons CD RN',
    'Mod Plus CD RN',
    'Mod CD RN',
    'Agr CD RN',
    'Agr Plus CD RN',
    'Super Conservador CD PE',
    'Conservador CD PE',
    'Moderado CD PE',
    'Agressivo CD PE',
    'Super Agressivo CD PE',
    'Sup Cons CD Néos',
    'Cons CD Néos',
    'Mod CD Néos',
    'Agr CD Néos', 
]

# Loop para escolha dos perfis
for opcoes_xpath_perfis in opcoes_xpath_perfis:

    #Seleciona os perfis
    elemento = navegador.find_element(By.XPATH, '//*[@id="MainContent_MainContent_ddlTipoCota"]').send_keys(opcoes_xpath_perfis)
        
    #Exporta balancete
    navegador.find_element('xpath','//*[@id="MainContent_MainContent_btnExportar"]').click()
    time.sleep (1.75)

time.sleep (120)

nomes_antigos = [
    'Balancete.xlsx',
    'Balancete (1).xlsx',
    'Balancete (2).xlsx',
    'Balancete (3).xlsx',
    'Balancete (4).xlsx',
    'Balancete (5).xlsx',
    'Balancete (6).xlsx',
    'Balancete (7).xlsx',
    'Balancete (8).xlsx',
    'Balancete (9).xlsx',
    'Balancete (10).xlsx',
    'Balancete (11).xlsx',
    'Balancete (12).xlsx',
    'Balancete (13).xlsx',
    'Balancete (14).xlsx',
    'Balancete (15).xlsx',
    'Balancete (16).xlsx',
    'Balancete (17).xlsx',
    'Balancete (18).xlsx',
    'Balancete (19).xlsx',
    'Balancete (20).xlsx',
    'Balancete (21).xlsx',
    'Balancete (22).xlsx',
    'Balancete (23).xlsx',
    'Balancete (24).xlsx',
    'Balancete (25).xlsx',
]

nomes_novos = [
    f'Balancete CD BA - {referencia}.xlsx',
    f'Balancete BD BA - {referencia}.xlsx',
    f'Balancete CD RN - {referencia}.xlsx',
    f'Balancete PGA - {referencia}.xlsx',
    f'Balancete BD RN - {referencia}.xlsx',
    f'Balancete CD PE - {referencia}.xlsx',
    f'Balancete BD PE - {referencia}.xlsx',
    f'Balancete CD NÉOS - {referencia}.xlsx',
    f'Balancete CD BA Básico - {referencia}.xlsx',
    f'Balancete CD BA Diferenciado - {referencia}.xlsx',
    f'Balancete CD BA Conservador - {referencia}.xlsx',
    f'Balancete CD RN Conservador - {referencia}.xlsx',
    f'Balancete CD RN Moderado Plus - {referencia}.xlsx',
    f'Balancete CD RN Moderado - {referencia}.xlsx',
    f'Balancete CD RN Agressivo - {referencia}.xlsx',
    f'Balancete CD RN Agressivo Plus - {referencia}.xlsx',
    f'Balancete CD PE Super Conservador - {referencia}.xlsx',
    f'Balancete CD PE Conservador - {referencia}.xlsx',
    f'Balancete CD PE Moderado - {referencia}.xlsx',
    f'Balancete CD PE Agressivo - {referencia}.xlsx',
    f'Balancete CD PE Super Agressivo - {referencia}.xlsx',
    f'Balancete CD Néos Super Conservador - {referencia}.xlsx',
    f'Balancete CD Néos Conservador - {referencia}.xlsx',
    f'Balancete CD Néos Moderado - {referencia}.xlsx',
    f'Balancete CD Néos Agressivo - {referencia}.xlsx', 
]

# Adicione este código ao seu script
for i in range(len(nomes_antigos)):
    # Construa o caminho completo para os nomes de arquivo antigos e novos
    nome_antigo = os.path.join(diretorio, nomes_antigos[i])
    novo_nome = os.path.join(diretorio, nomes_novos[i])

    # Renomeia o arquivo
    os.rename(nome_antigo, novo_nome)




navegador.quit