import requests
from bs4 import BeautifulSoup
import openpyxl
import time
import urllib3
from urllib.parse import urljoin
import re

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

start_time = time.time()

def get_departments_links():
    main_url = 'https://prepod.nspu.ru/course/index.php'
    response = requests.get(main_url, verify=False)
    src = response.text
    soup = BeautifulSoup(src, "lxml") # структурированный объект html

    departments_elements_array = soup.select('h3 > a')
    # Удаляем последний элемент
    departments_elements_array.pop()
    departments_links_array = [link.get('href') for link in departments_elements_array]

    return departments_links_array


def get_cathedras_links(departments_links_array):
    cathedras_links_array = []
    
    for link in departments_links_array:
        response = requests.get(link, verify=False)
        src = response.text
        soup = BeautifulSoup(src, "lxml") # структурированный объект html

        cathedras_elements_array = soup.find_all('a', string=lambda text: text and (text.startswith('Кафедра') or text.startswith('Зачное')))
        for elem in cathedras_elements_array:
            cathedras_links_array.append(elem.get('href'))

    return cathedras_links_array


def get_employees_links(cathedras_links_array):
    employees_links_array = []
    
    for link in cathedras_links_array:
        response = requests.get(link, verify=False)
        src = response.text
        soup = BeautifulSoup(src, "lxml") # структурированный объект html

        employees_links_array.extend([a['href'] for a in soup.select('.coursename > a')])

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
    sheet['D1'] = "Факультет/Институт"
    sheet['E1'] = "Должность"
    sheet['F1'] = "Телефон"
    sheet['G1'] = "Email"
    sheet['H1'] = "Кафедра"
    sheet['I1'] = "Персональная страница"

    return workbook, sheet




def get_dictionary(employees_links_array):
    employees_dictionary = {}

    for profile in employees_links_array:
        response2 = requests.get(profile, verify=False)
        src2 = response2.text
        soup2 = BeautifulSoup(src2, "lxml")

        contact_link = soup2.select_one('a:has(span.instancename:-soup-contains("Мои контакты")), a:has(span.instancename:-soup-contains("Контакты"))')['href'] if soup2.select_one('a:has(span.instancename:-soup-contains("Мои контакты")), a:has(span.instancename:-soup-contains("Контакты"))') else '-'


        post_elements = soup2.select('.content > .no-overflow [style="text-align: center;"]')
        filtered_post_elements = [elem for elem in post_elements if not elem.find('img')]

        post_str = ''

        for post_elem in filtered_post_elements:
            post_str += str(post_elem)

        soup_post = BeautifulSoup(post_str,"lxml")
        post = soup_post.get_text(strip=True)

        if contact_link != '-':
            employees_dictionary[contact_link] = post

        time.sleep(2)

    return employees_dictionary


def data_parsing(soup, workbook, sheet, row, post_not_formatted, profile_url):
    fio_array = soup.select_one('.page-header-headings > h1').text.split(' ')
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

    fio = ' '.join(fio_array).strip() # Для удаление подстроки в должности

    nav_element = [element.text for element in soup.select('span[itemprop="title"]')]

    faculty = nav_element[2]

    post = post_not_formatted.replace(fio, "")
    
    rab_tel_tag = soup.find(lambda tag: tag.name == 'strong' and 'Раб.' in tag.text) or None
    mob_tel_tag = soup.find(lambda tag: tag.name == 'strong' and 'Моб.' in tag.text) or None
    # email_tag = soup.find(lambda tag: tag.name == 'strong' and 'Электронная почта' in tag.text)

    # Если тег найден, получаем следующий текстовый узел (текст номера телефона)
    rab_tel_number = rab_tel_tag.next_sibling.strip() if rab_tel_tag and isinstance(rab_tel_tag.next_sibling, str) else '-'
    mob_tel_number = mob_tel_tag.next_sibling.strip() if mob_tel_tag and isinstance(mob_tel_tag.next_sibling, str) else '-'


    # email = email_tag.next_sibling.strip() if email_tag and isinstance(email_tag.next_sibling, str) else '-'
    # if (email == "-"):
    #     for elem in email_tag.text.split(' '):
    #         if "@" in elem:
    #             cleaned_email = re.sub('[^a-zA-Z0-9@.]', '', elem)
    #             email = cleaned_email

    email_text = soup.find(string=lambda text: "@" in text) or None
    email = '-'
    if email_text:
        for elem in email_text.split(' '):
            if "@" in elem:
                cleaned_email = re.sub('[^a-zA-Z0-9@._]', '', elem)
                email = cleaned_email.strip()



    phone = f"Раб.тел: {rab_tel_number}\nМоб.тел: {mob_tel_number}".strip()

    cathedra = '-'
    for elem in nav_element:
        if 'Кафедра' in elem or 'Заочное' in elem:
            cathedra = elem


    data = [
        (surname, name, patronymic, faculty, post, phone, email, cathedra, profile_url)
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
            sheet[f'I{row}'] = person[8]

    # Сохраняем Excel-файл
    workbook.save("employees_NSPU.xlsx")


def get_data_in_profile(dictionary):
    workbook, sheet = create_excel()
    row = 2  # Начинаем со второй строки (после заголовков)

    for profile_url, post in dictionary.items():
        # Проверка, что profile_url является строкой и не пустой
        try:
            response = requests.get(profile_url, verify=False)
            src = response.text
            soup = BeautifulSoup(src, "lxml") # структурированный объект html

            # Парсинг данных:
            data_parsing(soup, workbook, sheet, row, post, profile_url)
            row += 1
            time.sleep(2)
        except requests.exceptions.RequestException:
            # Просто продолжить выполнение цикла без вывода сообщения об ошибке
            pass


def main():
    # Сбор всех подразделений
    departments_links_array = get_departments_links()
    url_exception1 = 'https://prepod.nspu.ru/course/index.php?categoryid=47'
    exception1_array = []
    new_departments_links_array = [elem for elem in departments_links_array if elem != url_exception1]
    exception1_array.append(url_exception1)
    # Сбор всех кафедр (учесть исключения: институт открытого дистанционного образования, Заочное отделение) - сбор остальных по "Кафедра"
    cathedras_links_array = get_cathedras_links(new_departments_links_array)
    # Сбор сотрудников (переход в каждого сотрудника элемент "p", класс "cathedra__teacher")
    employees_links_array = get_employees_links(cathedras_links_array)
    employees_exception1_array = get_employees_links(exception1_array)
    employees_links_array.extend(employees_exception1_array)
    # Получение словаря {'ссылка на контакт': 'должность'}
    dictionary = get_dictionary(employees_links_array)
    # Сбор ФИО, Факультет, Должность, телефон, почта, кафедра
    get_data_in_profile(dictionary)


if __name__ == "__main__":
    main()


end_time = time.time()
execution_time_seconds = end_time - start_time
execution_time_minutes = execution_time_seconds / 60

print(f"Программа выполнена за {execution_time_minutes} минут")