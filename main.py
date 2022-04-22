from distutils.log import error
from email.mime.text import MIMEText
from multiprocessing.connection import wait
from openpyxl import load_workbook
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import *
from time import sleep
import os
import re
import smtplib
import ssl
import email
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart


# Global Variables
user = os.getlogin()


class Charles:
    def begin(self):
        self.__init__()
        self.web_entry_b3()
        self.excel_entry()
        self.fundamentus()
        self.filtros()
        self.emailtask()

    def __init__(self):
        chrome_options = Options()
        chrome_options.add_experimental_option(
            'excludeSwitches', ['enable-logging'])
        chrome_options.add_argument('--lang=pr-BR')
        chrome_options.add_argument('--disable-notifications')
        chrome_options.add_argument('--disable-infobars')
        chrome_options.add_argument('--disable-extensions')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--no-sandbox')
        self.driver = webdriver.Chrome()

    def web_entry_b3(self):
        self.link = 'https://www.b3.com.br/pt_br/produtos-e-servicos/negociacao/renda-variavel/fundos-de-investimentos/fii/fiis-listados/'
        self.driver.get(self.link)
        self.driver.implicitly_wait(10)
        self.driver.maximize_window()
        self.wait = WebDriverWait(self.driver, 10)
        sleep(2)
        ####################### CONDIÇÕES PARA FECHAR POPUPS #######################################
        if self.driver.find_element(by=By.XPATH, value='//*[@id="onetrust-banner-sdk"]/div/div[1]/div').is_displayed():
            self.driver.find_element(
                by=By.XPATH, value='//*[@id="onetrust-close-btn-container"]/button').click()

        self.driver.implicitly_wait(10)

        self.driver.switch_to.frame('bvmf_iframe')

        self.driver.implicitly_wait(10)

        self.driver.find_element(
            By.XPATH, value='//*[@id="divContainerIframeB3"]/div/div/div/div[1]/div[2]/p/a').click()
        sleep(5)

    def excel_entry(self):
        # -------------------  Criação do arquivo excel -------------------
        pathDownload = 'C:/Users/' + user + '/Downloads/fundosListados.csv'
        print()
        if os.path.exists(pathDownload):
            # ------------------- Inserindo dados no excel base -------------------------------------------------
            self.excel_data = pd.read_csv(
                pathDownload, sep=';', encoding='latin1')
            # ---------- tratativa ----------
            self.excel_data.rename(columns={'Código': 'Codigo'}, inplace=True)
            self.excel_data.to_csv(
                pathDownload, sep=';', encoding='latin1')
            # -------
            self.excel_data = pd.read_csv(
                pathDownload, sep=';', encoding='latin1')
            df = pd.DataFrame(self.excel_data, columns=[r'Segmento'])
        df.to_excel('C:/Users/' + user + '/Downloads/fundos_imob.xlsx')
        # remove o arquivo csv já utilizado
        os.remove(pathDownload)

    def fundamentus(self):
        df2 = pd.read_excel('C:/Users/' + user + '/Downloads/fundos_imob.xlsx')
        df2.drop(['Unnamed: 0'], axis=1, inplace=True)
        print(df2)
        baseUrl = 'https://fundamentus.com.br/detalhes.php?papel='

        fii_list = df2.values.tolist()
        rows = []
        # ------------------- Coletando dados no site Fundamentus -------------------------------------------------
        for fii in fii_list:
            fii = fii[0]
            sleep(1)
            self.driver.get(baseUrl + fii + '11')
            sleep(1)
            try:
                msg = self.driver.find_element(
                    by=By.CSS_SELECTOR, value='body > div.center > div.conteudo.clearfix > div > div > h1').text
                print(msg)
            except:
                try:
                    # extrai p/vp do site fundamentus
                    pvp = self.driver.find_element(
                        by=By.CSS_SELECTOR, value='body > div.center > div.conteudo.clearfix > table:nth-child(5) > tbody > tr:nth-child(4) > td:nth-child(4) > span').text
                    # extrai dividend yield do site fundamentus
                    dy = self.driver.find_element(
                        by=By.CSS_SELECTOR, value='body > div.center > div.conteudo.clearfix > table:nth-child(5) > tbody > tr:nth-child(3) > td:nth-child(4) > span').text
                    # extrai cotação do site fundamentus
                    cotacao = self.driver.find_element(
                        by=By.CSS_SELECTOR, value='body > div.center > div.conteudo.clearfix > table:nth-child(3) > tbody > tr:nth-child(1) > td.data.destaque.w3 > span').text
                    # get patrimonio liquido
                    patrimonio_liquido = self.driver.find_element(
                        by=By.CSS_SELECTOR, value='body > div.center > div.conteudo.clearfix > table:nth-child(5) > tbody > tr:nth-child(12) > td:nth-child(6) > span').text
                    # extrai Segmento do site fundamentus
                    segmento = self.driver.find_element(
                        by=By.CSS_SELECTOR, value='body > div.center > div.conteudo.clearfix > table:nth-child(3) > tbody > tr:nth-child(4) > td:nth-child(2) > span > a').text
                    rows.append(
                        [fii, pvp, dy, cotacao, patrimonio_liquido, segmento])
                except Exception as e:
                    print(e)
                    print('erro ao extrair dados do site fundamentus' +
                          '\n' + 'Descricao do Erro:' + str(e))
        sleep(1)
        self.driver.close()
        df = pd.DataFrame(rows, columns=[
                          "Codigo FII", "p/vp", "Dividend Yield", "Cotacao", "Patrimonio Liquido", "Segmento"])
        print(df)
        df.to_excel('C:/Users/' + user + '/Downloads/fundos_imobiliarios.xlsx')
        # remove o arquivo xlsx já utilizado
        os.remove('C:/Users/' + user + '/Downloads/fundos_imob.xlsx')

    def filtros(self):
        df = pd.read_excel('C:/Users/' + user +
                           '/Downloads/fundos_imobiliarios.xlsx')
        df.drop(['Unnamed: 0'], axis=1, inplace=True)
        # ------------------- Tratando Colunas -------------------------------------------------
        df['Dividend Yield'] = df['Dividend Yield'].str.replace(
            r",", ".")
        df['Dividend Yield'] = df['Dividend Yield'].str.replace(r"%", "")
        df['Dividend Yield'] = pd.to_numeric(df['Dividend Yield'])
        df['Patrimonio Liquido'] = df['Patrimonio Liquido'].str.replace(
            r".", "")
        df['Patrimonio Liquido'] = pd.to_numeric(df['Patrimonio Liquido'])
        df['p/vp'] = df['p/vp'].str.replace(r",", ".")
        df['p/vp'] = pd.to_numeric(df['p/vp'])
        # ------------------- Paramêtros -------------------------------------------------
        param_dy = df['Dividend Yield'] >= 9
        param_patrimonio = df['Patrimonio Liquido'] >= 1000000000
        param_pvp = df['p/vp'] <= 0.95
        df2 = df[['Codigo FII', 'Dividend Yield', 'Patrimonio Liquido',
                  'p/vp']][param_dy & param_patrimonio & param_pvp]
        # print(df2)

        table = df2.to_html(classes='mystyle')
        html_string = f'''
        <html>
            <head><title>HTML Pandas Dataframe with CSS</title></head>
            <link rel="stylesheet" type="text/css" href="df_style.css"/>
            <body>
                {table}
            </body>
        </html>
        '''
        return html_string

    #-------------------------- email
    def emailtask(self):
        html_string = self.filtros()
        # --
        msg = MIMEMultipart("alternative")  # Define the main object
        msg["Subject"] = "Fundos Imobiliarios"  # Assunto
        msg["From"] = 'robopython.fii@gmail.com'  # remetente
        msg["To"] = 'vinicius-a-soares@outlook.com'  # 'E-mail do destinatário'
        # ---
        part = MIMEText(html_string, 'html')
        msg.attach(part)
        encoders.encode_base64(part)
        context = ssl.create_default_context()
        sender_email = 'robopython.fii@gmail.com'  # 'E-mail do remetente'
        password = 'pythonfundosimobiliarios'   # senha do email
        # 'E-mail do destinatário'
        receiver_email = 'inserir email destinatario'

        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, msg.as_string())
        print('Email enviado com sucesso!')


start = Charles()
start.begin()
