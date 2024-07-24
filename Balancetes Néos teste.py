from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
import time

usuario = input('Insira o usuário do AtenaPrev.Net: ')
senha = input('Insira a senha do AtenaPrev.Net: ')


chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

prefs = {"download.default_directory" : "H:\\GCO\\AFAF\\AFAF004\\GESTÃO CONTÁBIL\\Fechamento Mensal\\BALANCETES\\2024\\07.2024"}
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


#A partir daqui o código deveria emitir os balancete por perfil, mas o código não consegue selecionar o perfil, nisso emite sempre o balancete consolidado
opcoes_xpath_perfis = [
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[1]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[2]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[3]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[4]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[5]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[6]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[7]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[8]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[9]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[10]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[11]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[12]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[13]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[14]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[15]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[16]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[17]',
    '//*[@id="MainContent_MainContent_ddlTipoCota"]/option[18]',
    
]

# Loop para escolha dos perfis
for opcoes_xpath in opcoes_xpath:

    #Seleciona os perfis
    elemento = navegador.find_element(By.XPATH, opcoes_xpath_perfis)
    acoes = ActionChains(navegador)
    acoes.double_click(elemento).perform()
    time.sleep (0.75)
    
    #Exporta balancete
    navegador.find_element('xpath','//*[@id="MainContent_MainContent_btnExportar"]').click()
    time.sleep (1.75)