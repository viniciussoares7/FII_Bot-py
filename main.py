from argparse import Action
from numpy import inner
import numpy as np
import pandas as pd
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import *
from time import sleep
import smtplib
import os
from email.message import EmailMessage
import re


class Charles:
    def begin(self):
        # self.__init__()
        # self.web_entry()
        self.excel_entry()

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

    def web_entry(self):
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
        if os.path.exists('C:/Users/vinic/Downloads/fundosListados.csv'):
            self.excel_data = pd.read_csv(
                r'C:/Users/vinic/Downloads/fundosListados.csv', sep=';', encoding='latin-1')
            df = pd.DataFrame(self.excel_data, columns=[r'Código'])
            print(df)


start = Charles()
start.begin()
