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

def get_cath_fac_links():
    url = 'https://ngmu.ru/structure_bosses#depts'
    response = requests.get(url, verify=False)
    src = response.text
    soup = BeautifulSoup(src, 'lxml') # структурированный объект html

    a_elems = soup.find_all('a', string = lambda text: text and text.startswith('Кафедра'))
    links_a_elems = []
    for a in a_elems:
        links_a_elems.append('https://ngmu.ru' + a.get('href'))

    return links_a_elems


def get_employees_links(cath_fac_links):
    employ_and_division_dictionary = {}

    for url in cath_fac_links:
        response = requests.get(url, verify=False)
        src = response.text
        soup = BeautifulSoup(src, 'lxml')

        division = soup.find('div', class_ = 'title_orange no_screen').get_text()
        # ищу массив работников
        employees_elements = soup.find_all('a', href = lambda text: text and text.startswith('/users/'))
        employees_href_list = ['https://ngmu.ru' + element.get('href') for element in employees_elements]
        
        employ_and_division_dictionary[division] = employees_href_list

    return employ_and_division_dictionary


def create_excel():
   # Создаем новый Excel-файл и новый лист (страницу)
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # Заголовки для столбцов
    sheet['A1'] = "Фамилия"
    sheet['B1'] = "Имя"
    sheet['C1'] = "Отчество"
    sheet['D1'] = "Подразделение"
    sheet['E1'] = "Должность"
    #sheet['F1'] = "Телефон"
    sheet['F1'] = "Email"
    sheet['G1'] = "Кафедра"
    sheet['H1'] = "Персональная страница"

    return workbook, sheet



def data_parsing(soup, workbook, sheet, row, cath_fac, profile_url):
    try:
        fio_element = soup.find('div', class_='title_orange')
        if fio_element:
            fio_array = fio_element.text.split(' ')
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

            match = re.search(r'\((.*?)\)', cath_fac)
            if match:
                faculty = match.group(1)


            posts_list = [element.text.strip() for element in soup.select('div.smallgray:last-child div:last-child')]
            posts = ', '.join(posts_list)
            # Использовать метод join для объединения уникальных должностей в строку
            post = ', '.join(posts for posts in set(posts_list))


            email_text = soup.find(string=lambda text: "@" in text) or None
            email = '-'
            if email_text:
                for elem in email_text.split(' '):
                    if "@" in elem:
                        cleaned_email = re.sub('[^a-zA-Z0-9@._]', '', elem)
                        email = cleaned_email.strip()


            cath = re.sub(r'\([^)]*\)', '', cath_fac)

            profile = profile_url

            data = [
                (surname, name, patronymic, faculty, post, email, cath, profile)
            ]

            # Начинаем заполнять данные со второй строки (после заголовков)
            if email and email != '-':
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
            workbook.save("employees_ngmu.xlsx")

    except Exception as e:
        pass

    #fio_array = soup.find('div', class_='title_orange').text.split(' ')
    




def get_data_in_profile(employ_and_division_dictionary):
    workbook, sheet = create_excel()
    row = 2  # Начинаем со второй строки (после заголовков)

    for cath_fac, employees_list in employ_and_division_dictionary.items():
        for profile_url in employees_list:
            response = requests.get(profile_url, verify=False)
            src = response.text
            soup = BeautifulSoup(src, 'lxml')

            # Парсинг данных контакта
            data_parsing(soup, workbook, sheet, row, cath_fac, profile_url)
            row += 1
            time.sleep(1)

        time.sleep(1)


def main():
    # Сбор всех кафедр
    cath_fac_links = get_cath_fac_links()

    # Сбор словаря "Кафедра/факультет: список сотрудников"
    employ_and_division_dictionary = get_employees_links(cath_fac_links)

    # Сбор данных каждого сотрудника
    get_data_in_profile(employ_and_division_dictionary)



if __name__ == "__main__":
    main()


end_time = time.time()
execution_time_seconds = end_time - start_time
execution_time_minutes = execution_time_seconds / 60

print(f"Программа выполнена за {execution_time_minutes} минут")