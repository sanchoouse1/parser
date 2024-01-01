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
    sheet['E1'] = "web-сайт"
    sheet['F1'] = "Профиль организации"
    sheet['G1'] = "Математические и естественные науки"
    sheet['H1'] = "Инженерное дело, технологии и технические науки"
    sheet['I1'] = "Здравоохранение и медицинские науки"
    sheet['J1'] = "Сельское хозяйство и сельскохозяйственные науки"
    sheet['K1'] = "Науки об обществе"
    sheet['L1'] = "Образование и педагогические науки"
    sheet['M1'] = "Гуманитарные науки"
    sheet['N1'] = "Искусство и культура"
    sheet['O1'] = "Оборона и безопасность государства, военные науки"
    sheet['P1'] = "Наименование образовательной организации"
    sheet['Q1'] = "E.1. Образовательная деятельность"
    sheet['R1'] = "РФ 2023 Медиана"
    sheet['S1'] = "Субъект Медиана"
    sheet['T1'] = "E.2. Научно-исследовательская деятельность"
    sheet['U1'] = "РФ 2023 Медиана"
    sheet['V1'] = "Субъект Медиана"
    sheet['W1'] = "E.3. Международная деятельность"
    sheet['X1'] = "РФ 2023 Медиана"
    sheet['Y1'] = "Субъект Медиана"
    sheet['Z1'] = "E.4. Финансово-экономическая деятельность"
    sheet['AA1'] = "РФ 2023 Медиана"
    sheet['AB1'] = "Субъект Медиана"
    sheet['AC1'] = "E.5. Заработная плата ППС"
    sheet['AD1'] = "РФ 2023 Медиана"
    sheet['AE1'] = "Субъект Медиана"
    sheet['AF1'] = "E.8. Дополнительный показатель"
    sheet['AG1'] = "РФ 2023 Медиана"
    sheet['AH1'] = "Субъект Медиана"

    return workbook, sheet



def get_common_data(regions_list):
    parse_tr_to_td = []

    for region_url in regions_list:
        response = requests.get(region_url)
        src = response.text
        soup = BeautifulSoup(src, 'lxml')

        tr_elements = soup.select('table.an > tr') # Получаю список организаций в данной области

        for tr in tr_elements: # получаю ячейки каждой из организаций региона на данной итерации
            parse_tr_to_td.append(tr.select('td'))

        # time.sleep(1)

    return parse_tr_to_td # Список списков



# def get_information_second_tab_list(url):
#     response = requests.get(url)
#     src = response.text
#     soup = BeautifulSoup(src, 'lxml')




def write_characteristics_higher_educ_sys(parse_tr_to_td): #[[данные организации с первой вкладки], [данные организации с первой вкладки], [данные организации с первой вкладки], []],[]
    workbook, sheet = create_excel()
    row = 2  # Начинаем со второй строки (после заголовков)

    for educ_org in parse_tr_to_td:
        # for data in educ_org:

        url = 'https://monitoring.miccedu.ru/iam/2023/_vpo/' + educ_org[1].select_one('a').get('href')
        response = requests.get(url)
        src = response.text
        soup = BeautifulSoup(src, 'lxml')

        general_information_elements = soup.find_all('td', class_='tt')
        general_information_list = []

        for elem in general_information_elements:
            general_information_list.append(elem.find_next_sibling('td').text.strip())

        # town = re.search(r'(?:г\.|город\s)([^,]+)', general_information_list[1]).group(1) if re.search(r'(?:г\.|город\s)([^,]+)', general_information_list[1]) else general_information_list[1]
        town = general_information_list[1]

        indicators_strokes_list = soup.select('#result > tr') if soup.select('#result > tr') else "-"
        cells_strokes = []
        if indicators_strokes_list != "-":
            for stroke in indicators_strokes_list:
                cells_strokes.append(stroke.find_all('td'))

        district = soup.find('div', string=lambda text: text and 'округ' in text.lower()).text.strip() if soup.find('div', string=lambda text: text and 'округ' in text.lower()) else "-"
        #print(district)

        # for cells in cells_strokes:
        #     print('E - ' + cells[5].text)
        #     print('РФ - ' + cells[6].text)
        #     print('Субъект - ' + cells[7].text)

        sheet[f'A{row}'] = general_information_list[0]
        sheet[f'B{row}'] = district
        sheet[f'C{row}'] = town
        sheet[f'D{row}'] = general_information_list[2]
        sheet[f'E{row}'] = general_information_list[3]
        sheet[f'F{row}'] = general_information_list[5]
        sheet[f'G{row}'] = educ_org[2].text
        sheet[f'H{row}'] = educ_org[3].text
        sheet[f'I{row}'] = educ_org[4].text
        sheet[f'J{row}'] = educ_org[5].text
        sheet[f'K{row}'] = educ_org[6].text
        sheet[f'L{row}'] = educ_org[7].text
        sheet[f'M{row}'] = educ_org[8].text
        sheet[f'N{row}'] = educ_org[9].text
        sheet[f'O{row}'] = educ_org[10].text
        sheet[f'P{row}'] = general_information_list[0]

        if indicators_strokes_list == "-":
            row += 1
            # Сохраняем Excel-файл
            workbook.save("Monitoring.xlsx")
            continue

        sheet[f'Q{row}'] = re.search(r'([^|]+)\|', cells_strokes[0][5].text).group(1) if re.search(r'([^|]+)\|', cells_strokes[0][5].text) else cells_strokes[0][5].text
        sheet[f'R{row}'] = cells_strokes[0][6].text
        sheet[f'S{row}'] = cells_strokes[0][7].text
        sheet[f'T{row}'] = re.search(r'([^|]+)\|', cells_strokes[1][5].text).group(1) if re.search(r'([^|]+)\|', cells_strokes[1][5].text) else cells_strokes[1][5].text
        sheet[f'U{row}'] = cells_strokes[1][6].text
        sheet[f'V{row}'] = cells_strokes[1][7].text
        sheet[f'W{row}'] = re.search(r'([^|]+)\|', cells_strokes[2][5].text).group(1) if re.search(r'([^|]+)\|', cells_strokes[2][5].text) else cells_strokes[2][5].text
        sheet[f'X{row}'] = cells_strokes[2][6].text
        sheet[f'Y{row}'] = cells_strokes[2][7].text
        sheet[f'Z{row}'] = re.search(r'([^|]+)\|', cells_strokes[3][5].text).group(1) if re.search(r'([^|]+)\|', cells_strokes[3][5].text) else cells_strokes[3][5].text
        sheet[f'AA{row}'] = cells_strokes[3][6].text
        sheet[f'AB{row}'] = cells_strokes[3][7].text
        sheet[f'AC{row}'] = re.search(r'([^|]+)\|', cells_strokes[4][5].text).group(1) if re.search(r'([^|]+)\|', cells_strokes[4][5].text) else cells_strokes[4][5].text
        sheet[f'AD{row}'] = cells_strokes[4][6].text
        sheet[f'AE{row}'] = cells_strokes[4][7].text
        sheet[f'AF{row}'] = re.search(r'([^|]+)\|', cells_strokes[5][5].text).group(1) if re.search(r'([^|]+)\|', cells_strokes[5][5].text) else cells_strokes[5][5].text
        sheet[f'AG{row}'] = cells_strokes[5][6].text
        sheet[f'AH{row}'] = cells_strokes[5][7].text

        row += 1

        # Сохраняем Excel-файл
        workbook.save("Monitoring.xlsx")

        time.sleep(1)

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