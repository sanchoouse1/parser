import requests
from bs4 import BeautifulSoup

def find_all_department_url(element, links):
    for a in element.find_all('a'):
        #print('https://sfedu.ru' + a.get('href'))
        links.append('https://sfedu.ru' + a.get('href'))

def get_data_in_profile(profiles_url):
    # Список словарей
    employees_data = []

    # Проход по профилям
    for link in profiles_url:
        response_link = requests.get(link)
        src_link = response_link.text
        soup_link = BeautifulSoup(src_link, "lxml")

        # Заполнение данных:
        = {
            surname:
            name:
            patronymic:
            university_division:
            post:
            email:
            phone:
        }



def get_profile_url(section_employees_of_department):
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

    get_profile_url(section_employees_of_department)



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

#print(links)

transition_to_division(links_department, section_employees_of_department)
