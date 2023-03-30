from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import sys
sys.path.insert(1, "C:\\Users\\davi.costa\\Desktop")
from login import credenciais

servico = Service(ChromeDriverManager().install())
chrome_options = Options()
chrome_options.add_experimental_option("detach", True)

nome_usuario = credenciais.get('NOME_USUARIO')
senha = credenciais.get('SENHA')

navegador = webdriver.Chrome(service=servico, options=chrome_options)

navegador.get('http://sei.funai.gov.br/sei/controlador.php?acao=procedimento_trabalhar&acao_origem=procedimento_controlar&acao_retorno=procedimento_controlar&id_procedimento=5201689&infra_sistema=100000100&infra_unidade_atual=110001002&infra_hash=5c7536526aad967b0691884e9af71e0c4cacc03ccacdf5c475e915d14cd9679b')

navegador.find_element(
    "xpath", '//*[@id="txtUsuario"]').send_keys(nome_usuario)

navegador.find_element(
    "xpath", '//*[@id="pwdSenha"]').send_keys(senha)

navegador.find_element("xpath", '//*[@id="sbmLogin"]').click()

teste = navegador.find_element("xpath", '//*[@id="selInfraUnidades"]')

dropdown = Select(teste)

dropdown.select_by_visible_text("E-PAT-CGETNO")

navegador.find_element("xpath", '//*[@id="P5201689"]/td[3]/a').click()
navegador.implicitly_wait(5)

redimensionar = navegador.find_element(By.ID, "divRedimensionar")

actions = ActionChains(navegador)

actions.drag_and_drop_by_offset(redimensionar, 150, 0).perform()

iframe = navegador.find_element(By.ID, "ifrArvore")

navegador.switch_to.frame(iframe)

navegador.implicitly_wait(3)

WebDriverWait(navegador, 60).until(
    EC.element_to_be_clickable((By.ID, 'anchorAP5201689'))).click()

div = navegador.find_element(By.ID, 'div5201689')

pastas = div.find_elements(By.TAG_NAME, "div")

for pasta in pastas:
    elementos_filho = pasta.find_elements("xpath", "./*")
    substring = "Solicitação de Provisão"

    for elemento in elementos_filho:
        if substring in elemento.text:
            nome = elemento.text
            elemento.click()
            navegador.switch_to.default_content()
            iframe2 = navegador.find_element(By.ID, "ifrVisualizacao")
            navegador.switch_to.frame(iframe2)
            iframe3 = navegador.find_element(By.ID, "ifrArvoreHtml")
            navegador.switch_to.frame(iframe3)
            heading2 = navegador.find_elements(
                By.CLASS_NAME, 'Formulario_texto_Editavel_Alinhado_Esquerda')

            tds = navegador.find_elements(By.TAG_NAME, 'td')
            
            colunas = ['CR', 'OBJETIVO' ,'AÇÃO', 'PTRES', 'FONTE', 'PLANO INTERNO',
                       'TERRA INDÍGENA', 'ETNIA', 'TOTAL']
            
            informacoes = []
            
            for i, td in enumerate(tds):
                informacoes.append(td.text)
                print(td.text)
                if i == 1:
                    break

            for div in heading2:
                informacoes.append(div.text)

            informacoes_escolhidas = [elemento for elemento in informacoes if elemento != ' ' 
                                      if elemento !=
                                      'Para cobrir despesas com:' 
                                      if elemento != 'Conforme o documento de solicitação:']
            
            df = pd.DataFrame(columns=colunas)

            df = df.append(
                pd.Series(informacoes_escolhidas, index=df.columns), ignore_index=True)

            df.to_excel(f"Z:/CGETNO/CGETNO/Estágio/Davi/automacao/{nome}.xlsx", engine='xlsxwriter')
            navegador.switch_to.default_content()
            navegador.switch_to.frame(iframe)


print('fim')
