import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.keys import Keys
import json

EDGE_DRIVER_PATH = r'G:\chromedriver.exe'
URL = 'https://www.vivareal.com.br/venda/'

TIME_SLEEP = 3

def init_driver():
    option = webdriver.ChromeOptions()
    option.add_argument('--start-maximized')
    option.headless = False
    driver = webdriver.Chrome(EDGE_DRIVER_PATH, options=option)
    return driver

def get_pagina_by_bairro(bairro, driver):
    driver.get(URL)
    time.sleep(TIME_SLEEP)
    input_text = driver.find_element_by_xpath('//input[@id="filter-location-search-input"]')
    input_text.send_keys(bairro)
    time.sleep(TIME_SLEEP)
    input_text.send_keys(Keys.RETURN)
    time.sleep(TIME_SLEEP)
    driver.find_element_by_class_name('select-multiple__static').click()
    driver.find_element_by_id('unit-type-HOME|UnitSubType_NONE,SINGLE_STOREY_HOUSE,VILLAGE_HOUSE,KITNET|RESIDENTIAL|HOME').click()
    driver.find_element_by_id('unit-type-HOME|TWO_STORY_HOUSE|RESIDENTIAL|TWO_STORY_HOUSE').click()
    time.sleep(TIME_SLEEP)

def get_imoveis(bairro, driver):
    df_imoveis = pd.DataFrame(columns=['endereco', 'bairro', 'preco', 'area', 'n_quartos', 'n_banheiros', 'n_vagas'])
    pag = 1
    while True:
        imoveis = driver.find_elements_by_xpath('//div[@data-type="property"]')
        for im in imoveis:
            try:
                imovel = im.text.split('\n')
                end = imovel[1].split(bairro + ', São Paulo - SP')[0][:-3]
                preco = float(imovel[-1].split()[1].replace('.',''))
                area = int(imovel[2].split()[0])
                n_quartos = int(imovel[3].split()[0])
                n_banheiros = int(imovel[4].split()[0])
                n_vagas = int(imovel[5].split()[0])
                df_imoveis.loc[len(df_imoveis)+1] = [end,bairro,preco,area,n_quartos,n_banheiros,n_vagas]
            except ValueError:
                continue
        
        pag += 1
        driver.find_element_by_xpath('//li[@class="pagination__item"]//a[@href="?pagina={}"]'.format(str(pag))).send_keys('\n')
        time.sleep(TIME_SLEEP)
        pass

def main():
    list_bairros = ['Alto Da Boa Vista','Chácara Monte Alegre','Granja Julieta','Jardim Cordeiro',
    'Jardim dos Estados','Jardim Marajoara','Jardim Petrópolis','Jardim Prudência','Jardim Santo Amaro São Paulo']

    imoveis = pd.DataFrame(columns=['endereco', 'preco', 'area', 'n_quartos', 'n_banheiros', 'n_vagas'])

    driver = init_driver()

    for bairro in list_bairros:
        get_pagina_by_bairro(bairro,driver)
        df = get_imoveis(bairro, driver)


if __name__=='__main__':
    main()