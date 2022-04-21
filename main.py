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


class Charles:
    def begin(self):
        # self.__init__()
        # self.web_entry_b3()
        # self.excel_entry()
        # self.fundamentus()
        self.filtros()

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
        self.driver.close()

    def excel_entry(self):
        # -------------------  Criação do arquivo excel -------------------
        if os.path.exists('C:/Users/vinic/Downloads/fundosListados.csv'):
            # ------------------- Inserindo dados no excel base -------------------------------------------------
            self.excel_data = pd.read_csv(
                r'C:/Users/vinic/Downloads/fundosListados.csv', sep=';', encoding='latin-1')
            #self.excel_data = self.excel_data.astype(str)
            # ---------- tratativa ----------
            # for column in self.excel_data.columns:
            #    self.excel_data[column] = self.excel_data[column].str.replace(
            #        r'\W', "")
            df = pd.DataFrame(self.excel_data, columns=[r'Codigo'])
            print(df)
        df.to_excel("C:/bots/Charles/fundos_imob.xlsx")
        sleep(2)

    def fundamentus(self):
        df2 = pd.read_excel("C:/bots/Charles/fundos_imob.xlsx")
        df2.drop(['Unnamed: 0'], axis=1, inplace=True)
        baseUrl = 'https://fundamentus.com.br/detalhes.php?papel='

        fii_list = df2.values.tolist()
        rows = []
        # ------------------- Coletando dados no site Fundamentus -------------------------------------------------
        for fii in fii_list:
            fii = fii[0]
            self.driver.get(baseUrl + fii + '11')
            # sleep(0.5)
            try:
                msg = self.driver.find_element(
                    by=By.CSS_SELECTOR, value='body > div.center > div.conteudo.clearfix > div > div > h1').text
                print(msg)
            except:
                #print('Codigo FII: ' + fii)
                # get p/vp
                pvp = self.driver.find_element(
                    by=By.CSS_SELECTOR, value='body > div.center > div.conteudo.clearfix > table:nth-child(5) > tbody > tr:nth-child(4) > td:nth-child(4) > span').text
                #print('PVP: ' + pvp)
                # get dividend yield
                dy = self.driver.find_element(
                    by=By.CSS_SELECTOR, value='body > div.center > div.conteudo.clearfix > table:nth-child(5) > tbody > tr:nth-child(3) > td:nth-child(4) > span').text
                #print('DY: ' + dy)
                # get cotação
                cotacao = self.driver.find_element(
                    by=By.CSS_SELECTOR, value='body > div.center > div.conteudo.clearfix > table:nth-child(3) > tbody > tr:nth-child(1) > td.data.destaque.w3 > span').text
                #print('Cotacao: ' + cotacao)
                # get patrimonio liquido
                patrimonio_liquido = self.driver.find_element(
                    by=By.CSS_SELECTOR, value='body > div.center > div.conteudo.clearfix > table:nth-child(5) > tbody > tr:nth-child(12) > td:nth-child(6) > span').text
                #print('patrimonio liquido: ' + patrimonio_liquido)
                # get Segmento
                segmento = self.driver.find_element(
                    by=By.CSS_SELECTOR, value='body > div.center > div.conteudo.clearfix > table:nth-child(3) > tbody > tr:nth-child(4) > td:nth-child(2) > span > a').text
                #print('Segmento: ' + segmento)
                rows.append(
                    [fii, pvp, dy, cotacao, patrimonio_liquido, segmento])
        df = pd.DataFrame(rows, columns=[
                          "Codigo FII", "p/vp", "Dividend Yield", "Cotacao", "Patrimonio Liquido", "Segmento"])
        # print(df)
        df.to_excel("C:/bots/Charles/fundos_imobiliarios.xlsx")

    def filtros(self):
        df = pd.read_excel("C:/bots/Charles/fundos_imobiliarios.xlsx")
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
        print(df2)


start = Charles()
start.begin()
