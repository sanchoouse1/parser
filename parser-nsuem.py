import requests
from bs4 import BeautifulSoup
import openpyxl
import time
import urllib3
from urllib.parse import urljoin

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

start_time = time.time()

def get_cathedras_links():
    main_url = 'https://nsuem.ru/education/organization/faculties/'
    response = requests.get(main_url, verify=False)
    src = response.text
    soup = BeautifulSoup(src, "lxml") # структурированный объект html

    cathedras_elements_array = soup.find_all('a', string=lambda text: text and text.startswith('Кафедра'))
    cathedras_links_array = [link.get('href') for link in cathedras_elements_array]

    return cathedras_links_array


def get_employees_division_links(cathedras_links_array):
    employees_divisions_links = []

    for link in cathedras_links_array:
        url = 'https://nsuem.ru' + link
        response = requests.get(url, verify=False)
        src = response.text
        soup = BeautifulSoup(src, "lxml") # структурированный объект html

        employees_division_element = soup.find('a', class_ = 'employees')
        employees_division_link = employees_division_element.get('href')
        employees_divisions_links.append(employees_division_link)
        time.sleep(1)

    return employees_divisions_links



def get_employees_links(employees_divisions_links):
    employees_links_array = []
    
    for link in employees_divisions_links:
        url = 'https://nsuem.ru' + link
        response = requests.get(url, verify=False)
        src = response.text
        soup = BeautifulSoup(src, "lxml") # структурированный объект html

        employees_elements = soup.select('.cathedra__teacher > a')
        employees_links_array.extend([urljoin(url, link.get('href')) for link in employees_elements])
        time.sleep(1)

    return employees_links_array


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
    sheet['F1'] = "Телефон"
    sheet['G1'] = "Email"
    sheet['H1'] = "Кафедра"
    sheet['I1'] = "Персональная страница"

    return workbook, sheet



def data_parsing(soup, workbook, sheet, row, url):
    faculties_dict = {
        'fbp': 'Факультет базовой подготовки',
        'fgs': 'Факультет государственного сектора',
        'fcep': 'Факультет корпоративной экономики и предпринимательства',
        'uf': 'Юридический факультет',
        'itf': 'Факультет цифровых технологий'
    }

    fio = soup.find('h3', class_='npr-title').text.split(' ')
    if len(fio) == 3:
        surname = fio[0]
        name = fio[1]
        patronymic = fio[2]
    elif len(fio) > 3:
        surname = fio[0]
        name = fio[1]
        patronymic = ' '.join(fio[2:])
    elif len(fio) < 3:
        surname = fio[0]
        name = fio[1]
        patronymic = '-'


    url_parts = url.split('/')
    faculty_abbreviation = None
    for part in url_parts:
        if part in faculties_dict:
            faculty_abbreviation = part
            break
    faculty = faculties_dict[faculty_abbreviation]


    post_elements = soup.find('h4', class_='npr-sub-sub-title', string=lambda text: text and 'Должность' in text).text.split()
    post = ''
    for elem in post_elements[1:]:
        post += elem + ' '


    phones = soup.select('a[href^="tel:"]') if soup.select('a[href^="tel:"]') else None
    if phones:
        phone_list = []
        for phone_number in phones:
            phone_list.append(phone_number.text.strip())
    else:
        phone_list = "-"


    email_element = soup.select_one('a[href^="mailto:"]') if soup.select_one('a[href^="mailto:"]') else None
    if email_element:
        email = email_element.text
    else:
        email = "-"

    cathedra = ''
    cathedra_elements = soup.find('h4', class_='npr-sub-sub-title', string=lambda text: text and 'Подразделение' in text).text.split()
    for elem in cathedra_elements[1:]:
        cathedra += elem + ' '

    data = [
        (surname, name, patronymic, faculty, post, ", ".join(phone_list), email, cathedra, url)
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
        sheet[f'I{row}'] = person[8]

    # Сохраняем Excel-файл
    workbook.save("employees_NSUEM.xlsx")


def get_data_in_profile(employees_links_array):
    workbook, sheet = create_excel()
    row = 2  # Начинаем со второй строки (после заголовков)

    for profile in employees_links_array:
        response = requests.get(profile, verify=False)
        src = response.text
        soup = BeautifulSoup(src, "lxml") # структурированный объект html

        # Парсинг данных:
        data_parsing(soup, workbook, sheet, row, profile)
        row += 1
        time.sleep(2)




def main():
    # Сбор всех кафедр
    cathedras_links_array = get_cathedras_links()
    # Сбор ссылок на раздел "Сотрудники" (ссылка "a" с классом "employees")
    employees_divisions_links = get_employees_division_links(cathedras_links_array)
    # Сбор сотрудников (переход в каждого сотрудника элемент "p", класс "cathedra__teacher")
    employees_links_array = get_employees_links(employees_divisions_links)
    # Сбор ФИО, Факультет, Должность, телефон, почта, кафедра
    get_data_in_profile(employees_links_array)
    # Брать паузы


if __name__ == "__main__":
    main()


end_time = time.time()
execution_time_seconds = end_time - start_time
execution_time_minutes = execution_time_seconds / 60

print(f"Программа выполнена за {execution_time_minutes} минут")