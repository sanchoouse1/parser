import requests
from bs4 import BeautifulSoup
import openpyxl
import time

start_time = time.time()

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
    sheet['F1'] = "Телефон"
    sheet['G1'] = "Email"
    sheet['H1'] = "Кафедра/Отдел"
    sheet['I1'] = "Персональная страница"

    return workbook, sheet


def find_all_department_url(element, links):
    for a in element.find_all('a'):
        #print('https://sfedu.ru' + a.get('href'))
        links.append('https://sfedu.ru' + a.get('href'))

def data_parsing(soup_link, workbook, sheet, row):
    h2_main = soup_link.select('.wrapper h2') if soup_link.select('.wrapper h2') else None

    h2_fio = soup_link.find('div', class_='card') if soup_link.find('div', class_='card') else None
    if h2_fio:
        h2_fio = h2_fio.find_previous('h2') if h2_fio.find_previous('h2') else None

        if h2_fio:
            fio = h2_fio.text.split(' ')

            # Фамилия
            surname = fio[0]

            # Имя
            name = fio[1]


            # Отчество
            if len(fio) >= 3:
                patronymic = fio[2]
            else:
                patronymic = "-"
        else:
            surname = '-'
            name = '-'
            patronymic = '-'
    else:
        surname = '-'
        name = '-'
        patronymic = '-'



    if h2_main:
        # Подразделение
        university_division = h2_main[0].text.strip() # Учебное подразделение / КАФЕДРА
    else:
        university_division = "-"

    post = "-"
    cathedra_department = "-"

    time.sleep(3)

    # Должность
    text_div = soup_link.find('div', class_='text') if soup_link.find('div', class_='text') else None

    if text_div:
        link_a_department = text_div.find('a', href=lambda href: not href.startswith('tel:')) if text_div.find('a', href=lambda href: not href.startswith('tel:')) else None # КАФЕДРА / ПОДРАЗДЕЛЕНИЕ

        if link_a_department:
            paragraph = link_a_department.find_parent('p') if link_a_department.find_parent('p') else None

            if paragraph:
                post = paragraph.text.split(' -\n')[-1].strip() if paragraph.text.split(' -\n')[-1].strip() else "-" # ДОЛЖНОСТЬ
                # Кафедра / подразделение
                cathedra_department = link_a_department.get_text(strip=True) if link_a_department.get_text(strip=True) else "-"



    # Номера телефонов одного сотрудника (МАССИВ)
    phones = soup_link.select('.phones a[href^="tel:"]') if soup_link.select('.phones a[href^="tel:"]') else None
    if phones:
        phone = []
        for phone_number in phones:
            phone.append(phone_number.text.strip())
    else:
        phone = "-"


    # EMAIL
    email = soup_link.select_one('dd a').text.split('/')[-1] + '@sfedu.ru' if soup_link.select_one('dd a') else "-"

    personal_link_element = soup_link.find('a', href=lambda href: href and href.startswith('https://sfedu.ru/person/')) if soup_link.find('a', href=lambda href: href) else None
    if personal_link_element:
        personal_link = personal_link_element.text
    else:
        personal_link = "-"

    data = [
        (surname, name, patronymic, university_division, post, ", ".join(phone), email, cathedra_department, personal_link)
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
    workbook.save("employees_data.xlsx")



def get_data_in_profile(profiles_url):
    workbook, sheet = create_excel()
    row = 2  # Начинаем со второй строки (после заголовков)

    # Проход по профилям
    for link in profiles_url:
        response_link = requests.get(link)
        src_link = response_link.text
        soup_link = BeautifulSoup(src_link, "lxml")

        # Парсинг данных:
        data_parsing(soup_link, workbook, sheet, row)
        row += 1
        # Подождать 6 секунды перед следующим запросом
        time.sleep(5)



def get_profiles_url(section_employees_of_department):
    profiles_url = []
    # Проход по всем ссылкам во вкладку "Сотрудники"
    for link in section_employees_of_department:
        response_link = requests.get(link)
        src_link = response_link.text
        soup_link = BeautifulSoup(src_link, "lxml")

        # Получение ссылок на абсолютно все профили
        for td in soup_link.select('td'):
            links = td.find_all('a')
            if links:
                first_link = links[0]
                profiles_url.append('https:' + first_link.get('href')) # ПРОВЕРЯТЬ НЕТ ЛИ ОДНИХ И ТЕХ ЖЕ ССЫЛОК

    #print(profiles_url)

    get_data_in_profile(profiles_url)


def transition_to_division(links_department, section_employees_of_department):
    for link in links_department:
        response_link = requests.get(link)
        src_link = response_link.text
        soup_link = BeautifulSoup(src_link, "lxml")
        time.sleep(1)
        section_employees_of_department.append('https://sfedu.ru' + soup_link.select_one('.accardeon_menu > a:-soup-contains("Сотрудники")').get('href'))

    get_profiles_url(section_employees_of_department)



section_employees_of_department = []

url = 'https://sfedu.ru/www/stat_pages22.show?p=UNI/N11900/D'
response = requests.get(url)
src = response.text
soup = BeautifulSoup(src, "lxml") # lxml - парсер



links_department = []

for necessary_h4 in soup.find_all('h4'):
    if necessary_h4.text == 'Учебные и научные подразделения':
        for element in necessary_h4.find_next_siblings():
            if element.name != 'h4':
                find_all_department_url(element, links_department)
            else:
                break


transition_to_division(links_department, section_employees_of_department)

end_time = time.time()
execution_time_seconds = end_time - start_time
execution_time_minutes = execution_time_seconds / 60

print(f"Программа выполнена за {execution_time_minutes} минут")
