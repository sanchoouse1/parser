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

    return workbook, sheet


def find_all_department_url(element, links):
    for a in element.find_all('a'):
        #print('https://sfedu.ru' + a.get('href'))
        links.append('https://sfedu.ru' + a.get('href'))

def data_parsing(soup_link, workbook, sheet, row):
    h2_main = soup_link.select('.wrapper h2')
    h2_fio = soup_link.find('div', class_='card').find_previous('h2')
    fio = h2_fio.text.split(' ')

    # Фамилия
    surname = fio[0]


    # Имя
    name = fio[1]

    # Отчество
    if len(fio) >= 3:
        patronymic = fio[2]
    else:
        patronymic = "Отчество отсутствует"



    # Подразделение
    university_division = h2_main[0].text.strip() # Учебное подразделение / КАФЕДРА


    # Должность
    text_div = soup_link.find('div', class_='text')
    link_a_department = text_div.find('a', href=lambda href: not href.startswith('tel:')) # КАФЕДРА / ПОДРАЗДЕЛЕНИЕ
    paragraph = link_a_department.find_parent('p')
    post = paragraph.text.split(' -\n')[-1].strip() # ДОЛЖНОСТЬ


    # Номера телефонов одного сотрудника (МАССИВ)
    phones = soup_link.select('.phones a[href^="tel:"]')
    if phones:
        phone = []
        for phone_number in phones:
            phone.append(phone_number.text.strip())
    else:
        phone = "-"


    # EMAIL
    email = soup_link.select_one('dd a').text.split('/')[-1] + '@sfedu.ru'


    # Кафедра / подразделение
    cathedra_department = link_a_department.get_text(strip=True)

    data = [
        (surname, name, patronymic, university_division, post, ", ".join(phone), email, cathedra_department)
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
