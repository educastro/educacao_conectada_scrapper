from selenium import webdriver
from bs4 import BeautifulSoup
import csv
import time

with open('escolas_inscritas_no_educacao_conectada.csv', 'w') as csvFile:
    writer = csv.writer(csvFile)
    csv_header = ['ID', 'Escola', 'Codigo INEP', 'Estado', 'Cidade', 'Tipo', 'Qtd Alunos']
    writer.writerow(csv_header)

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
    table_td = table_total.find_elements_by_tag_name('td')

    #table_td = table_tr.find_elements_by_tag_name('td')
    print("page 1")
    counter = 0
    counterTotal = 0
    citiesList = []

    for row in table_td:
        counter += 1
        citiesList.append(row.text)
        if counter % 6 == 0:
            counterTotal += 1
            citiesList.insert(0, counterTotal)
            writer.writerow(citiesList)
            print(citiesList)
            citiesList = []

    next_button = firefox.find_element_by_link_text('2')
    next_button.click()

    time.sleep(10)

    table_total = firefox.find_element_by_id('data-table')
    table_td = table_total.find_elements_by_tag_name('td')

    #table_td = table_tr.find_elements_by_tag_name('td')
    print("page 2")
    counter = 0
    citiesList = []

    for row in table_td:
        counter += 1
        citiesList.append(row.text)
        if counter % 6 == 0:
            counterTotal += 1
            citiesList.insert(0, counterTotal)
            writer.writerow(citiesList)
            print(citiesList)
            citiesList = []

    next_button = firefox.find_element_by_link_text('3')
    next_button.click()

    time.sleep(10)

    table_total = firefox.find_element_by_id('data-table')
    table_td = table_total.find_elements_by_tag_name('td')

    #table_td = table_tr.find_elements_by_tag_name('td')
    print("page 3")
    counter = 0
    citiesList = []

    for row in table_td:
        counter += 1
        citiesList.append(row.text)
        if counter % 6 == 0:
            counterTotal += 1
            citiesList.insert(0, counterTotal)
            writer.writerow(citiesList)
            print(citiesList)
            citiesList = []


csvFile.close()
