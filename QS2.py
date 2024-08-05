# Bibliotecas para Web Scraping
import requests
from bs4 import BeautifulSoup

# Bibliotecas para Navegação
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Biblioteca para acesso as credenciais
import os
from dotenv import load_dotenv

# Biblioteca para temporização
from time import sleep

# Resolver Captcha
from anticaptchaofficial.recaptchav2proxyless import *

# Bibliotecas para o Excel
import pandas as pd
from openpyxl import Workbook

# Biblioteca para o envio dos emails
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

# Biblioteca de Avisos
import warnings
warnings.filterwarnings('ignore')

# Carrega as variáveis do arquivo .env
load_dotenv()

# Obtém as credenciais
name = os.getenv("USUARIO")
senha = os.getenv("SENHA")
mail_user = os.getenv("MAIL_USER")
mail_password = os.getenv("MAIL_PASSWORD")
anticaptcha_api_key = os.getenv("ANTICAPTCHA_API_KEY")

# Verifica se as credenciais foram carregadas corretamente
#if not name or not senha or not mail_user or not mail_password or not anticaptcha_api_key:
    #raise ValueError("As credenciais não foram carregadas corretamente do arquivo .env")

# Configurações do Chrome para iniciar no modo incógnito
options = Options()
options.add_argument("--incognito")

# Inicializa o serviço do ChromeDriver
service = Service()

# Inicializa o driver do Selenium com as opções configuradas
driver = webdriver.Chrome(service=service, options=options)

# Navega para a página de login
link1 = "https://cmegroup.quikstrike.net/Account/Login.aspx?ReturnUrl=%2fUser%2fQuikStrikeView.aspx%3finit%3d&init="
driver.get(link1)

# Aguarda a página carregar
sleep(5)

# Clique no botão "Continue" (verifique se o Xpath está correto)
button_continue = driver.find_element(By.XPATH, '/html/body/form/div[3]/div[2]/div[2]/div/div[2]/div/div[1]/div[1]/div[2]/div[2]/p[2]/input')
button_continue.click()
sleep(5)

# Login no Site usando as credenciais carregadas do arquivo .env
input_usuario = driver.find_element(By.ID, "user")
input_usuario.send_keys(name)
sleep(5)

input_senha = driver.find_element(By.ID, "pwd")
input_senha.send_keys(senha)
sleep(5)

# Clica no botão de login (verifique se o ID está correto)
button_login = driver.find_element(By.ID, "loginBtn")
button_login.click()
sleep(5)

# Resolver o Captcha
link2 = "https://cmegroup-sso.quikstrike.net/User/Disclaimer.aspx?ret=%2fUser%2fQuikStrikeView.aspx%3finit%3d"
WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="pnlControls"]/div/div[2]')))
chave_captcha = driver.find_element(By.XPATH, '//*[@id="pnlControls"]/div/div[2]').get_attribute('data-sitekey')

solver = recaptchaV2Proxyless()
solver.set_verbose(1)
solver.set_key("anticaptcha_api_key")
solver.set_website_url("link2")
solver.set_website_key("chave_captcha")

resposta = solver.solve_and_return_solution()

if resposta != 0:
    print("resposta: " + resposta)
    # preencher o campo do token do captcha
    driver.execute_script(f"document.getElementById('g-recaptcha-response').innerHTML = '{resposta}'")
    driver.find_element(By.ID, 'btnContinue').click()
else:
    print("task finished with error: " + solver.error_code)

# Aguarda a página carregar
sleep(15)

# Acessar Ativo SP
buttonopint = driver.find_element(By.XPATH, '//*[@id="ctl00_ucMenuBar_lvMenuBar_ctrl2_lbMenuItem"]')
buttonopint.click()
sleep(8)

buttonselcativ = driver.find_element(By.XPATH, '//*[@id="ctl08_hlProductArrow"]')
buttonselcativ.click()
sleep(8)

buttoneqind = driver.find_element(By.XPATH, '//*[@id="ctl08_ucProductSelectorPopup_pnlProductSelectorPopup"]/div/div/div[1]/div[2]/a[4]')
buttoneqind.click()
sleep(8)

buttones500 = driver.find_element(By.XPATH, '//*[@id="ctl08_ucProductSelectorPopup_pnlProductSelectorPopup"]/div/div/div[3]/div[2]/a[1]')
buttones500.click()
sleep(8)

buttonOI = driver.find_element(By.XPATH, '//*[@id="MainContent_ucViewControl_OpenInterestV2_lbOIMatrix"]')
buttonOI.click()
sleep(10)

# Baixar a tabela das opções
tbsp500 = driver.find_element(By.XPATH, '//*[@id="MainContent_ucViewControl_OpenInterestV2_ucMatrixVC_ucMatrix_pnlMatrix"]/table')
html_tbsp500 = tbsp500.get_attribute("outerHTML")

def tratar_tabela1(html_tabela1):
    soup1 = BeautifulSoup(html_tabela1, "html.parser")
    tabela1 = soup1.find(name="table")
    df1 = pd.read_html(str(tabela1))[0]

    # Tratamentos da tabela1
    df1.fillna(0, inplace=True)
    df1 = df1.droplevel(0, axis=1)
    df1 = df1.droplevel(0, axis=1)
    df1 = df1.rename(columns={"Unnamed: 0_level_2": "Strike_1", "Unnamed: 1_level_2": "Strike_2"})
    df1.drop('Strike_2', axis=1, inplace=True)
    df1['C_sum'] = df1.filter(like='C').sum(axis=1)
    df1['P_sum'] = df1.filter(like='P').sum(axis=1)
    df1['GEX'] = df1['C_sum'] - df1['P_sum']
    df1 = df1.drop(['C', 'P'], axis=1)
    df1.set_index('Strike_1', inplace=True)

    return df1

# Tratamento da Tabela de Open Interest
df1sp500 = tratar_tabela1(html_tbsp500)

# Acessando a tabela de volatividade do ativo
buttonOInf = driver.find_element(By.XPATH, '//*[@id="ctl00_ucMenuBar_lvMenuBar_ctrl4_lbMenuItem"]')
buttonOInf.click()
sleep(10)

# Importar o HTML com a planilha das informações de vol do ativo
tbsp500i = driver.find_element(By.XPATH, '//*[@id="MainContent_ucViewControl_OptionsInfo_ucSummary_ucSummaryControlV2_ucATMs_ucSheetVC_ucSheet_divSheet"]/table')
html_tbsp500i = tbsp500i.get_attribute("outerHTML")

def tratar_tabela2(html_tabela2):
    soup2 = BeautifulSoup(html_tabela2, "html.parser")
    tabela2 = soup2.find(name="table")
    dfi = pd.read_html(str(tabela2))[0]

    # Preencher os campos vazios do DataFrame com zero
    dfi.fillna(0, inplace=True)

    # Renomear as colunas
    new_columns1 = ['', 'codigo', '', 'Stranddle', 'Dias p Venc', 'Futuro', 'Price', 'Vol', '+/-']
    dfi = dfi.set_axis(new_columns1, axis='columns')

    # Remover a linha 0
    dfi = dfi.drop(0)

    # Copiar o DataFrame
    dfia = dfi.copy()

    # Converter atributos/colunas para 'numeric'
    for atributo in ['Dias p Venc', 'Futuro', 'Price', 'Vol', '+/-']:
        dfia[atributo] = pd.to_numeric(dfia[atributo], errors='coerce')

    # Inserir colunas adicionais
    conversao = 100
    dfia['Volc'] = dfia['Vol'] / conversao
    dfia['Preco'] = dfia['Price'] / conversao
    dfia['dias venc'] = dfia['Dias p Venc'] / conversao
    dfia['Fut'] = dfia['Futuro'] / conversao

    # Remover colunas desnecessárias
    dfia = dfia.drop(['Dias p Venc', 'Futuro', 'Price', 'Vol'], axis=1)

    return dfia

# Tratamento da Tabela de Volatividade
dfiasp500 = tratar_tabela2(html_tbsp500i)

# Acessar Ativo NQ
buttonopint = driver.find_element(By.XPATH, '//*[@id="ctl00_ucMenuBar_lvMenuBar_ctrl2_lbMenuItem"]')
buttonopint.click()
sleep(10)

buttonopint = driver.find_element(By.XPATH, '//*[@id="ctl08_imgArrow"]')
buttonopint.click()
sleep(10)

buttonselcativ1 = driver.find_element(By.XPATH, '//*[@id="ctl08_ucProductSelectorPopup_pnlProductSelectorPopup"]/div/div/div[1]/div[2]/a[4]')
buttonselcativ1.click()
sleep(10)

buttoneqind1 = driver.find_element(By.XPATH, '//*[@id="ctl08_ucProductSelectorPopup_pnlProductSelectorPopup"]/div/div/div[1]/div[2]/a[4]')
buttoneqind1.click()
sleep(10)

buttonnq2 = driver.find_element(By.XPATH, '//*[@id="ctl08_ucProductSelectorPopup_pnlProductSelectorPopup"]/div/div/div[3]/div[2]/a[2]')
buttonnq2.click()
sleep(10)

buttonOI = driver.find_element(By.XPATH, '//*[@id="MainContent_ucViewControl_OpenInterestV2_lbOIMatrix"]')
buttonOI.click()
sleep(10)

# Baixar a tabela das opções
tbnq100 = driver.find_element(By.XPATH, '//*[@id="MainContent_ucViewControl_OpenInterestV2_ucMatrixVC_ucMatrix_pnlMatrix"]/table')
html_tbnq100 = tbnq100.get_attribute("outerHTML")

# Tratamento da Tabela de Open Interest
df1nq100 = tratar_tabela1(html_tbnq100)

# Ir para a pagina com as tabelas de vol do ativo
buttonOInf = driver.find_element(By.XPATH, '//*[@id="ctl00_ucMenuBar_lvMenuBar_ctrl4_lbMenuItem"]')
buttonOInf.click()
sleep(10)

# Importar o HTML com a planilha das informações de vol do ativo
tbnq100i = driver.find_element(By.XPATH, '//*[@id="MainContent_ucViewControl_OptionsInfo_ucSummary_ucSummaryControlV2_ucATMs_ucSheetVC_ucSheet_divSheet"]/table')
html_tbnq100i = tbnq100i.get_attribute("outerHTML")

# Tratamento da Tabela de Volatividade
dfianq100 = tratar_tabela2(html_tbnq100i)

def tratar_tabela3(html_tabela3):
    soup3 = BeautifulSoup(html_tabela3, "html.parser")
    tabela3 = soup3.find(name="table")
    dfaj = pd.read_html(str(tabela3))[0]

    # Remover o indes da tabela
    dfaj = dfaj.reset_index(drop=True)
    
    return dfaj

# Baixar a tabela dos ajustes do NQ
link3 = "https://www.cmegroup.com/markets/equities/nasdaq/e-mini-nasdaq-100.settlements.html"  # Substitua pelo link desejado
driver.get(link3)
sleep(5)

# Importar a tabela dos Ajustes Mini NQ
tbajnq100 = driver.find_element(By.XPATH, '/html/body/main/div/div[3]/div[3]/div/div/div/div/div/div[2]/div/div/div/div/div/div[8]/div/div')
html_tbajnq100 = tbajnq100.get_attribute("outerHTML")

# Tratamento da Tabela de Ajuste
dfajnq100 = tratar_tabela3(html_tbajnq100)

# Baixar a tabela dos ajustes do ES
link4 = "https://www.cmegroup.com/markets/equities/sp/e-mini-sandp500.settlements.html"  # Substitua pelo link desejado
driver.get(link4)
sleep(5)

# Importar a tabela dos Ajustes Mini ESP
tbajsp500 = driver.find_element(By.XPATH, '//*[@id="productTabData"]/div/div/div/div/div/div[2]/div/div/div/div/div/div[8]/div/div')
html_tbajsp500 = tbajsp500.get_attribute("outerHTML")

# Tratamento da Tabela de Ajuste
dfajsp500 = tratar_tabela3(html_tbajsp500)

# Iniciando o Excel
wb = Workbook()
ws = wb.active

import datetime
# Obtém a data atual
data_atual = datetime.datetime.now().strftime("%d%m%y_%H%M%S")

# Nome do arquivo Excel
nome_arquivo_excel = f"Open_Interest_{data_atual}.xlsx"

# Salvar o arquivo Excel
abas = {
    "SP500": df1sp500,
    "sp500vol": dfiasp500,
    "sp500aju": dfajsp500,
    "NQ100": df1nq100,
    "nq100vol": dfianq100,
    "nq100aju": dfajnq100
}

with pd.ExcelWriter(nome_arquivo_excel) as writer:
    for sheet_name, df in abas.items():
        df.to_excel(writer, sheet_name=sheet_name)

# Configuração do corpo do email

# Função para identificar os pontos importantes
def analisar_dataframe(df, nome_da_coluna='GEX', nome_do_dataframe='DataFrame'):
    # Encontre os índices dos três maiores valores
    indices_maiores = df[nome_da_coluna].nlargest(3).index
    # Encontre os índices dos três menores valores
    indices_menores = df[nome_da_coluna].nsmallest(3).index

    # Filtra os valores positivos na coluna
    valores_positivos = df[df[nome_da_coluna] > 0]

    # Encontra o índice do menor valor positivo
    indice_menor_valor_positivo = valores_positivos[nome_da_coluna].idxmin()

    # Exibe os resultados
    maior_indice = indices_maiores[0]
    segundo_maior_indice = indices_maiores[1]
    terceiro_maior_indice = indices_maiores[2]
    menor_indice = indices_menores[0]
    segundo_menor_indice = indices_menores[1]
    terceiro_menor_indice = indices_menores[2]
    menor_valor_positivo_indice = indice_menor_valor_positivo  # Índice do menor valor positivo

    # Retorna os resultados
    return f"{nome_do_dataframe} -> Voll_Trigger, {menor_valor_positivo_indice}, Call_Hall, {maior_indice}, L.Gama+1, {segundo_maior_indice}, L.Gama+2, {terceiro_maior_indice}, Put_Hall, {menor_indice}, L.Gama-1, {segundo_menor_indice}, L.Gama-2, {terceiro_menor_indice}"

# Chamar a função para analisar o DataFrame
resultados1 = analisar_dataframe(df1sp500, nome_do_dataframe='df1sp500')
resultados2 = analisar_dataframe(df1nq100, nome_do_dataframe='df1nq100')

# Configurar o envio de email

server_smtp = "smtp.gmail.com"
port = 587
sender_email = os.getenv("MAIL_USER")
password = os.getenv("MAIL_PASSWORD")

# Configurações do e-mail
receive_email = "josejuniormoreira82@gmail.com"
subject = "Open Interest do dia"
body = (f"Segue a planilha com os dados de Open Interest das opções.\n\n"
        f"Dados SP500:\n{resultados1}\n\n"
        f"Dados NQ100:\n{resultados2}")

# Criando o e-mail
message = MIMEMultipart()
message["From"] = sender_email
message["To"] = receive_email
message["Subject"] = subject
message.attach(MIMEText(body, "plain"))

# Anexando o arquivo Excel
try:
    with open(nome_arquivo_excel, "rb") as attachment:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(nome_arquivo_excel)}")
        message.attach(part)
except Exception as e:
    print(f"Erro ao anexar o arquivo: {e}")


# Função para enviar e-mail com tentativas de repetição
try:
    # Conectar ao servidor SMTP
    smtp = smtplib.SMTP(server_smtp, port)
    smtp.set_debuglevel(1)  # Ativar debug para ver detalhes da conexão
    smtp.ehlo()
    print(f"[*] Echoing the server: OK")

    # Iniciar conexão TLS
    smtp.starttls()
    smtp.ehlo()  # Re-emissão do EHLO depois de iniciar TLS
    print(f"[*] Starting TLS connection: OK")

    # Login no servidor SMTP
    smtp.login(sender_email, password)
    print(f"[*] Logging in: OK")

    # Enviar o e-mail
    smtp.sendmail(sender_email, receive_email, message.as_string())
    print("E-mail enviado com sucesso")

except smtplib.SMTPAuthenticationError as e:
    print(f"Erro de autenticação SMTP: {e.smtp_code}, {e.smtp_error}")
except smtplib.SMTPServerDisconnected as e:
    print(f"Erro de desconexão do servidor SMTP: {e}")
except Exception as e:
    print(f"Houve algum erro: {e}")
finally:
    try:
        smtp.quit()
    except (NameError, smtplib.SMTPServerDisconnected):
        pass


# Fecha o navegador
driver.quit()
