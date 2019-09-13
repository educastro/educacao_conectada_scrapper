from selenium import webdriver
from bs4 import BeautifulSoup
import csv
import time

firefox = webdriver.Firefox()
firefox.get('http://educacaoconectada.mec.gov.br/consulta-pdde')

element_uf = firefox.find_element_by_id('filter_uf')
for option in element_uf.find_elements_by_tag_name('option'):
    if option.text == 'DF':
        option.click()
        break

element_municipio = firefox.find_element_by_id('filter_municipio')
for option in element_municipio.find_elements_by_tag_name('option'):
    if option.text == 'BRASILIA':
        option.click()
        break

button_consultar = firefox.find_element_by_id('filter')
button_consultar.click()


element_data_table_length = firefox.find_element_by_name('data-table_length')
for option in element_data_table_length.find_elements_by_tag_name('option'):
    if option.text == '100':
        option.click()
        break

time.sleep(10)

table_total = firefox.find_element_by_id('data-table')
table_tr = table_total.find_elements_by_tag_name('tr')
#table_td = table_tr.find_elements_by_tag_name('td')
print("teste1")

for row in table_tr:
    print(row.text)
    print("###")
