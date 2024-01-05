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
    sheet['B1'] = "Тип"
    sheet['C1'] = "Округ"
    sheet['D1'] = "Город"
    sheet['E1'] = "Ведомственная принадлежность"
    sheet['F1'] = "web-сайт"
    sheet['G1'] = "Профиль организации"
    sheet['H1'] = "Математические и естественные науки"
    sheet['I1'] = "Инженерное дело, технологии и технические науки"
    sheet['J1'] = "Здравоохранение и медицинские науки"
    sheet['K1'] = "Сельское хозяйство и сельскохозяйственные науки"
    sheet['L1'] = "Науки об обществе"
    sheet['M1'] = "Образование и педагогические науки"
    sheet['N1'] = "Гуманитарные науки"
    sheet['O1'] = "Искусство и культура"
    sheet['P1'] = "Оборона и безопасность государства, военные науки"
    sheet['Q1'] = "Наименование образовательной организации"
    sheet['R1'] = "E.1. Образовательная деятельность"
    sheet['S1'] = "РФ 2023 Медиана"
    sheet['T1'] = "Субъект Медиана"
    sheet['U1'] = "E.2. Научно-исследовательская деятельность"
    sheet['V1'] = "РФ 2023 Медиана"
    sheet['W1'] = "Субъект Медиана"
    sheet['X1'] = "E.3. Международная деятельность"
    sheet['Y1'] = "РФ 2023 Медиана"
    sheet['Z1'] = "Субъект Медиана"
    sheet['AA1'] = "E.4. Финансово-экономическая деятельность"
    sheet['AB1'] = "РФ 2023 Медиана"
    sheet['AC1'] = "Субъект Медиана"
    sheet['AD1'] = "E.5. Заработная плата ППС"
    sheet['AE1'] = "РФ 2023 Медиана"
    sheet['AF1'] = "Субъект Медиана"
    sheet['AG1'] = "E.8. Дополнительный показатель"
    sheet['AH1'] = "РФ 2023 Медиана"
    sheet['AI1'] = "Субъект Медиана"
    sheet['AJ1'] = "Наименование образовательной организации"
    sheet['AK1'] = "Средний балл ЕГЭ студентов университета, принятых по результатам ЕГЭ на обучение по очной форме по программам бакалавриата и специалитета за счет средств соответствующих бюджетов бюджетной системы Российской Федерации, за исключением лиц, поступивших с учетом особых прав и в рамках квоты целевого приема"
    sheet['AL1'] = "Удельный вес численности обучающихся (приведенного контингента), по программам магистратуры, подготовки научно-педагогических кадров в аспирантуре (адъюнктуре), ординатуры, ассистентуры-стажировки в общей численности приведенного контингента обучающихся по основным образовательным программам высшего образования"
    sheet['AM1'] = "Удельный вес численности слушателей из сторонних организаций в общей численности слушателей, прошедших обучение в образовательной организации по программам повышения квалификации или профессиональной переподготовки"
    sheet['AN1'] = "Общий объем научно-исследовательских и опытно-конструкторских работ (далее – НИОКР)"
    sheet['AO1'] = "Удельный вес доходов от НИОКР в общих доходах образовательной организации"
    sheet['AP1'] = "Доходы от НИОКР (за исключением средств бюджетов бюджетной системы Российской Федерации, государственных фондов поддержки науки) в расчете на одного НПР[н]"
    sheet['AQ1'] = "Удельный вес численности НПР без ученой степени – до 30 лет, кандидатов наук – до 35 лет, докторов наук – до 40 лет, в общей численности НПР[н]"
    sheet['AR1'] = "Количество полученных грантов за отчетный год в расчете на 100 НПР[н]"
    sheet['AS1'] = "Объем средств, полученных образовательной организацией от выполнения НИОКР от иностранных граждан и иностранных юридических лиц"
    sheet['AT1'] = "Объем средств от образовательной деятельности, полученных образовательной организацией от иностранных граждан и иностранных юридических лиц"
    sheet['AU1'] = "Доходы образовательной организации из средств от приносящей доход деятельности в расчете на одного НПР"
    sheet['AV1'] = "Доля доходов из средств от приносящей доход деятельности в доходах по всем видам финансового обеспечения (деятельности) образовательной организации"
    sheet['AW1'] = "Отношение средней заработной платы НПР в образовательной организации (из всех источников) к средней заработной плате по экономике региона"
    sheet['AX1'] = "Доходы образовательной организации из всех источников в расчете на численность студентов (приведенный контингент)"
    sheet['AY1'] = "Количество персональных компьютеров в расчете на одного студента (приведенного контингента)"
    sheet['AZ1'] = "Удельный вес стоимости машин и оборудования (не старше 5 лет) в общей стоимости машин и оборудования"
    sheet['BA1'] = "Удельный вес НПР, имеющих ученую степень кандидата наук, в общей численности НПР"
    sheet['BB1'] = "Удельный вес НПР имеющих ученую степень доктора наук, в общей численности НПР"
    sheet['BC1'] = "Число НПР, имеющих ученую степень кандидата и доктора наук, в расчете на 100 студентов"

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
        url = 'https://monitoring.miccedu.ru/iam/2023/_vpo/' + educ_org[1].select_one('a').get('href')
        response = requests.get(url, timeout=60)
        src = response.text
        soup = BeautifulSoup(src, 'lxml')

        general_information_elements = soup.find_all('td', class_='tt')
        general_information_list = []

        for elem in general_information_elements:
            general_information_list.append(elem.find_next_sibling('td').text.strip())

        # town = re.search(r'(?:г\.|город\s)([^,]+)', general_information_list[1]).group(1) if re.search(r'(?:г\.|город\s)([^,]+)', general_information_list[1]) else general_information_list[1]
        # town = general_information_list[1]

        indicators_strokes_list = soup.select('#result > tr') if soup.select('#result > tr') else "-"
        cells_strokes = []
        if indicators_strokes_list != "-":
            for stroke in indicators_strokes_list:
                cells_strokes.append(stroke.find_all('td'))

        type_org_dict = {'st__2_0.png': 'образовательные организации высшего образования', 'st__0_0.png':'филиалы', 'st_3_0.png':'вузы (филиалы), находящиеся в стадии реорганизации/ реорганизованные'}
        img_name = educ_org[0].get('style', '').split('url(')[-1].split(')')[0].split('/')[-1]
        if img_name in type_org_dict:
            type_org = type_org_dict[img_name]

        district = soup.find('div', string=lambda text: text and 'округ' in text.lower()).text.strip() if soup.find('div', string=lambda text: text and 'округ' in text.lower()) else "-"
        town = soup.select_one('hr').find_next_sibling('div').text.split(' ')[1].strip() if soup.select_one('hr') else "-"

        sheet[f'A{row}'] = general_information_list[0]
        sheet[f'B{row}'] = type_org
        sheet[f'C{row}'] = district
        sheet[f'D{row}'] = town
        sheet[f'E{row}'] = general_information_list[2]
        sheet[f'F{row}'] = general_information_list[3]
        sheet[f'G{row}'] = general_information_list[5]
        sheet[f'H{row}'] = float(educ_org[2].text.replace(' ', '').replace(',', '.')) if educ_org[2].text else None
        sheet[f'I{row}'] = float(educ_org[3].text.replace(' ', '').replace(',', '.')) if educ_org[3].text else None
        sheet[f'J{row}'] = float(educ_org[4].text.replace(' ', '').replace(',', '.')) if educ_org[4].text else None
        sheet[f'K{row}'] = float(educ_org[5].text.replace(' ', '').replace(',', '.')) if educ_org[5].text else None
        sheet[f'L{row}'] = float(educ_org[6].text.replace(' ', '').replace(',', '.')) if educ_org[6].text else None
        sheet[f'M{row}'] = float(educ_org[7].text.replace(' ', '').replace(',', '.')) if educ_org[7].text else None
        sheet[f'N{row}'] = float(educ_org[8].text.replace(' ', '').replace(',', '.')) if educ_org[8].text else None
        sheet[f'O{row}'] = float(educ_org[9].text.replace(' ', '').replace(',', '.')) if educ_org[9].text else None
        sheet[f'P{row}'] = float(educ_org[10].text.replace(' ', '').replace(',', '.')) if educ_org[10].text else None
        sheet[f'Q{row}'] = general_information_list[0]

        if indicators_strokes_list == "-":
            row += 1
            # Сохраняем Excel-файл
            workbook.save("Monitoring.xlsx")
            continue

        value_mapping = {
            '1.2': None,
            '1.10': None,
            '1.14': None,
            '2.7': None,
            '2.8': None,
            '2.10': None,
            '2.13': None,
            '2.16': None,
            '3.12': None,
            '3.13': None,
            '4.1': None,
            '4.2': None,
            '4.3': None,
            '4.4': None,
            '5.6': None,
            '5.7': None,
            '6.1': None,
            '6.2': None,
            '6.4': None,
        }

        results_monitoring_tr_elems = soup.select('table.napde > tr')

        for elem_tr in results_monitoring_tr_elems:
            td = elem_tr.select('td')
            key = td[0].text.strip()
            if key in value_mapping:
                try:
                    value_mapping[key] = float(td[-1].text.replace(' ', '').replace(',', '.'))
                except ValueError:
                    value_mapping[key] = None

        sheet[f'R{row}'] = float(re.search(r'([^|]+)\|', cells_strokes[0][5].text.replace(' ', '').replace(',', '.')).group(1) if re.search(r'([^|]+)\|', cells_strokes[0][5].text) else cells_strokes[0][5].text.replace(' ', '').replace(',', '.'))
        sheet[f'S{row}'] = float(cells_strokes[0][6].text.replace(' ', '').replace(',', '.'))
        sheet[f'T{row}'] = float(cells_strokes[0][7].text.replace(' ', '').replace(',', '.'))
        sheet[f'U{row}'] = float(re.search(r'([^|]+)\|', cells_strokes[1][5].text.replace(' ', '').replace(',', '.')).group(1) if re.search(r'([^|]+)\|', cells_strokes[1][5].text) else cells_strokes[1][5].text.replace(' ', '').replace(',', '.'))
        sheet[f'V{row}'] = float(cells_strokes[1][6].text.replace(' ', '').replace(',', '.'))
        sheet[f'W{row}'] = float(cells_strokes[1][7].text.replace(' ', '').replace(',', '.'))
        sheet[f'X{row}'] = float(re.search(r'([^|]+)\|', cells_strokes[2][5].text.replace(' ', '').replace(',', '.')).group(1) if re.search(r'([^|]+)\|', cells_strokes[2][5].text) else cells_strokes[2][5].text.replace(' ', '').replace(',', '.'))
        sheet[f'Y{row}'] = float(cells_strokes[2][6].text.replace(' ', '').replace(',', '.'))
        sheet[f'Z{row}'] = float(cells_strokes[2][7].text.replace(' ', '').replace(',', '.'))
        sheet[f'AA{row}'] = float(re.search(r'([^|]+)\|', cells_strokes[3][5].text.replace(' ', '').replace(',', '.')).group(1) if re.search(r'([^|]+)\|', cells_strokes[3][5].text) else cells_strokes[3][5].text.replace(' ', '').replace(',', '.'))
        sheet[f'AB{row}'] = float(cells_strokes[3][6].text.replace(' ', '').replace(',', '.'))
        sheet[f'AC{row}'] = float(cells_strokes[3][7].text.replace(' ', '').replace(',', '.'))
        sheet[f'AD{row}'] = float(re.search(r'([^|]+)\|', cells_strokes[4][5].text.replace(' ', '').replace(',', '.')).group(1) if re.search(r'([^|]+)\|', cells_strokes[4][5].text) else cells_strokes[4][5].text.replace(' ', '').replace(',', '.'))
        sheet[f'AE{row}'] = float(cells_strokes[4][6].text.replace(' ', '').replace(',', '.'))
        sheet[f'AF{row}'] = float(cells_strokes[4][7].text.replace(' ', '').replace(',', '.'))
        sheet[f'AG{row}'] = float(re.search(r'([^|]+)\|', cells_strokes[5][5].text.replace(' ', '').replace(',', '.')).group(1) if re.search(r'([^|]+)\|', cells_strokes[5][5].text) else cells_strokes[5][5].text.replace(' ', '').replace(',', '.'))
        sheet[f'AH{row}'] = float(cells_strokes[5][6].text.replace(' ', '').replace(',', '.'))
        sheet[f'AI{row}'] = float(cells_strokes[5][7].text.replace(' ', '').replace(',', '.'))
        sheet[f'AJ{row}'] = general_information_list[0]
        sheet[f'AK{row}'] = value_mapping['1.2']
        sheet[f'AL{row}'] = value_mapping['1.10']
        sheet[f'AM{row}'] = value_mapping['1.14']
        sheet[f'AN{row}'] = value_mapping['2.7']
        sheet[f'AO{row}'] = value_mapping['2.8']
        sheet[f'AP{row}'] = value_mapping['2.10']
        sheet[f'AQ{row}'] = value_mapping['2.13']
        sheet[f'AR{row}'] = value_mapping['2.16']
        sheet[f'AS{row}'] = value_mapping['3.12']
        sheet[f'AT{row}'] = value_mapping['3.13']
        sheet[f'AU{row}'] = value_mapping['4.1']
        sheet[f'AV{row}'] = value_mapping['4.2']
        sheet[f'AW{row}'] = value_mapping['4.3']
        sheet[f'AX{row}'] = value_mapping['4.4']
        sheet[f'AY{row}'] = value_mapping['5.6']
        sheet[f'AZ{row}'] = value_mapping['5.7']
        sheet[f'BA{row}'] = value_mapping['6.1']
        sheet[f'BB{row}'] = value_mapping['6.2']
        sheet[f'BC{row}'] = value_mapping['6.4']

        row += 1

        # Сохраняем Excel-файл
        workbook.save("Monitoring.xlsx")

        time.sleep(2)

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