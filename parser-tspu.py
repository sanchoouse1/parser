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


def get_employees_links_list():
    url = 'https://www.tspu.edu.ru/all-persons.html'
    response = requests.get(url, verify=False)
    src = response.text
    soup = BeautifulSoup(src, 'lxml') # структурированный объект html

    employees_elements = soup.select('.person-fio > a', href = lambda href_text: href_text and href_text.startswith('/person.html?'))
    employees_links_list = []
    for a in employees_elements:
        employees_links_list.append('https://www.tspu.edu.ru' + a.get('href'))

    return employees_links_list



def create_excel():
    # Создаем новый Excel-файл и новый лист (страницу)
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Заголовки для столбцов
    sheet['A1'] = "Фамилия"
    sheet['B1'] = "Имя"
    sheet['C1'] = "Отчество"
    sheet['D1'] = "Факультет"
    sheet['E1'] = "Должность"
    #sheet['F1'] = "Телефон"
    sheet['F1'] = "Email"
    sheet['G1'] = "Кафедра"
    sheet['H1'] = "Персональная страница"

    return workbook, sheet



def data_parsing(soup, workbook, sheet, row, profile_url):
    fio_element = soup.select_one('.person-fio')
    if fio_element:
        fio_array = fio_element.text.strip().split(' ')
        if len(fio_array) == 3:
            surname = fio_array[0]
            name = fio_array[1]
            patronymic = fio_array[2]
        elif len(fio_array) > 3:
            surname = fio_array[0]
            name = fio_array[1]
            patronymic = ' '.join(fio_array[2:])
        elif len(fio_array) < 3:
            surname = fio_array[0]
            name = fio_array[1]
            patronymic = '-'
    else:
        # return False
        # Тест
        surname = '-'
        name = '-'
        patronymic = '-'


    pattern_division = re.compile('Подразделение')
    division_element = soup.find('strong', string=pattern_division)
    if division_element and division_element.next_sibling:
        division = division_element.next_sibling.strip()
    else:
        division = '-'



    pattern_post = re.compile('Должность')
    post_element = soup.find('strong', string=pattern_post)
    if post_element and post_element.next_sibling:
        post = post_element.next_sibling.strip()
    else:
        post = '-'


    faculties_dict = {
            'БХФ': 'Биолого-химический факультет',
            'ИИЯМС': 'Институт иностранных языков и международного сотрудничества',
            'ИФФ': 'Историко-филологический факультет',
            'ФМФ': 'Физико-математический факультет',
            'ИФКС': 'Институт физической культуры и спорта',
            'ИРПО': 'Институт развития педагогического образования',
            'ИДиА': 'Институт детства и артпедагогики',
            'ФПСО': 'Факультет психологии и специального образования',
            'ТЭФ': 'Технолого-экономический факультет',
            'ЦДОРК': 'Центр дополнительного образования и развития компетенций',
            'ИНИР': 'Институт научных исследований и разработок'
        }

    pattern_faculty = re.compile('Факультет')
    faculty_element = soup.find('strong', string=pattern_faculty)
    if faculty_element and faculty_element.next_sibling:
        faculty_abbreviation = faculty_element.next_sibling.strip()
        if faculty_abbreviation in faculties_dict:
            faculty = faculties_dict[faculty_abbreviation]
        else:
            if any(key in division for key in faculties_dict):
                faculty = next(key for key in faculties_dict if key in division)
            else:
                return False
    else:
        if any(key in division for key in faculties_dict):
            faculty = next(faculties_dict[key] for key in faculties_dict if key in division)
        else:
            return False


    email_element = soup.select_one('.person-email > a')
    if email_element:
        email = email_element.text.strip()
    else:
        return False
    
    data = [
        (surname, name, patronymic, faculty, post, email, division, profile_url)
    ]

    # Начинаем заполнять данные со второй строки (после заголовков)
    for person in data:
        sheet[f'A{row}'] = person[0]
        sheet[f'B{row}'] = person[1]
        sheet[f'C{row}'] = person[2]
        sheet[f'D{row}'] = person[3]
        sheet[f'E{row}'] = person[4]
        sheet[f'F{row}'] = person[5]
        sheet[f'G{row}'] = person[6]
        sheet[f'H{row}'] = person[7]

    # Сохраняем Excel-файл
    workbook.save("employees_TSPU.xlsx")

    return True



def get_data_in_profile(employees_links_list):
    workbook, sheet = create_excel()
    row = 2  # Начинаем со второй строки (после заголовков)
    profile_url = 'https://www.tspu.edu.ru/person.html?person=792'
    for profile_url in employees_links_list:
        response = requests.get(profile_url, verify=False)
        src = response.text
        soup = BeautifulSoup(src, 'lxml')

        # Парсинг данных контакта
        if not data_parsing(soup, workbook, sheet, row, profile_url):
            time.sleep(1)
            continue

        row += 1
        time.sleep(1)

    time.sleep(1)



def main():
    # Сбор сотрудников
    employees_links_list = get_employees_links_list()
    # Получение данных
    get_data_in_profile(employees_links_list)





if __name__ == "__main__":
    main()


end_time = time.time()
execution_time_seconds = end_time - start_time
execution_time_minutes = execution_time_seconds / 60

print(f"Программа выполнена за {execution_time_minutes} минут")