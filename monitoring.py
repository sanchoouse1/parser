import requests
from bs4 import BeautifulSoup
import openpyxl
import time
import urllib3
from urllib.parse import urljoin
import re

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

start_time = time.time()

# ... #

def get_regions():
    url = 'https://monitoring.miccedu.ru/?m=vpo&year=2023'
    response = requests.get(url, verify=False)
    src = response.text
    soup = BeautifulSoup(src, 'lxml')

    regions_elements = soup.select('td > p:not([class]):not([style]) > a')
    regions_list = []
    for element in regions_elements:
        regions_list.append('https://monitoring.miccedu.ru/' + element.get('href'))
    return regions_list


def create_excel():
    # Создаем новый Excel-файл и новый лист (страницу)
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Заголовки для столбцов
    sheet['A1'] = "Наименование образовательной организации"
    sheet['B1'] = "Округ"
    sheet['C1'] = "Город"
    sheet['D1'] = "Ведомственная принадлежность"
    sheet['E1'] = "Профиль организации"
    sheet['F1'] = "Математические и естественные науки"
    sheet['G1'] = "Инженерное дело, технологии и технические науки"
    sheet['H1'] = "Здравоохранение и медицинские науки"
    sheet['I1'] = "Сельское хозяйство и сельскохозяйственные науки"
    sheet['J1'] = "Науки об обществе"
    sheet['K1'] = "Образование и педагогические науки"
    sheet['L1'] = "Гуманитарные науки"
    sheet['M1'] = "Искусство и культура"
    sheet['N1'] = "Оборона и безопасность государства, военные науки"
    sheet['O1'] = "Наименование образовательной организации"
    sheet['P1'] = "E.1. Образовательная деятельность"
    sheet['Q1'] = "РФ 2023 Медиана"
    sheet['R1'] = "Субъект Медиана"
    sheet['S1'] = "E.2. Научно-исследовательская деятельность"
    sheet['T1'] = "РФ 2023 Медиана"
    sheet['U1'] = "Субъект Медиана"
    sheet['V1'] = "E.3. Международная деятельность"
    sheet['W1'] = "РФ 2023 Медиана"
    sheet['X1'] = "Субъект Медиана"
    sheet['Y1'] = "E.4. Финансово-экономическая деятельность"
    sheet['Z1'] = "РФ 2023 Медиана"
    sheet['AA1'] = "Субъект Медиана"
    sheet['AB1'] = "E.5. Заработная плата ППС"
    sheet['AC1'] = "РФ 2023 Медиана"
    sheet['AD1'] = "Субъект Медиана"
    sheet['AE1'] = "E.8. Дополнительный показатель"
    sheet['AF1'] = "РФ 2023 Медиана"
    sheet['AG1'] = "Субъект Медиана"

    return workbook, sheet



def get_common_data(regions_list):
    workbook, sheet = create_excel()
    row = 2  # Начинаем со второй строки (после заголовков)

    parse_tr_to_td = []

    for region_url in regions_list:
        response = requests.get(region_url)
        src = response.text
        soup = BeautifulSoup(src, 'lxml')

        tr_elements = soup.select('table.an > tr') # Получаю список организаций в данной области

        for tr in tr_elements: # получаю ячейки каждой из организаций региона на данной итерации
            parse_tr_to_td.append(tr.select('td'))

    return parse_tr_to_td # Список списков



def write_educ_common_data():
    # ...



def write_characteristics_higher_educ_sys(parse_tr_to_td): #[[данные организации с первой вкладки], [данные организации с первой вкладки], [данные организации с первой вкладки], []],[]
    for educ_org in parse_tr_to_td:
        for data in educ_org:
            # print('https://monitoring.miccedu.ru/iam/2023/_vpo/' + parse_tr_to_td[i][1].select_one('a').get('href'))
            # print(parse_tr_to_td[i][2].text) # математические и естественные науки
            # print(parse_tr_to_td[i][3].text) # инженерное дело, тех...
            # print(parse_tr_to_td[i][4].text) # здравоохранение
            # print(parse_tr_to_td[i][5].text) # сельское хозяйство
            # print(parse_tr_to_td[i][6].text) # науки об обществе
            # print(parse_tr_to_td[i][7].text) # образование и педагогические науки
            # print(parse_tr_to_td[i][8].text) # гуманитарные науки
            # print(parse_tr_to_td[i][9].text) # искусство и культура
            # print(parse_tr_to_td[i][10].text) # оборона и безопасность



def main():
    regions_list = get_regions()
    parse_tr_to_td = get_common_data(regions_list) # плохое имя переменной (исправить)
    write_characteristics_higher_educ_sys(parse_tr_to_td)

if __name__ == "__main__":
    main()


end_time = time.time()
execution_time_seconds = end_time - start_time
execution_time_minutes = execution_time_seconds / 60

print(f"Программа выполнена за {execution_time_minutes} минут")