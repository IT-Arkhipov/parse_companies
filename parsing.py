from bs4 import BeautifulSoup
import requests
import openpyxl


main_page = 'https://o-zavodah.ru'
base_url = 'https://o-zavodah.ru/zavody-proizvoditeli-mebeli/'

custom_user_agent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'  # noqa
headers = {'User-Agent': custom_user_agent}
response = requests.get(base_url, headers=headers)
soup = BeautifulSoup(response.text, 'lxml')

target_elements = soup.find_all('ul', class_='fw24i977up')

li_children_with_itemscope = []
for ul in target_elements:
    li_children = ul.find_all('li', attrs={'itemscope': True})
    li_children_with_itemscope.extend(li_children)

href_links = {}
wb = openpyxl.Workbook()
ws = wb.active
ws.append(['Наименование', 'Руководитель', 'ИНН', 'Сайт'])

for li in li_children_with_itemscope:
    company_name = li.find('span', class_='ellipsis name').get_text(strip=True)
    a_tags = li.find_all('a')
    for a in a_tags:
        href_link = a.get('href')
        if href_link:  # Check if href exists before adding it to the list
            href_links.update({company_name: href_link})

for company_name, link in href_links.items():
    response = requests.get(f'{main_page}{link}', headers=headers)
    soup = BeautifulSoup(response.text, 'lxml')
    factory_info = soup.find('ul', class_='stretchFlexBox factoryInfo')

    requisites = {}

    for li in factory_info.find_all('li'):
        # Find the <h4> tag within the current <li>
        h4_tag = li.find('h4')

        # Check if the <h4> tag's text matches either 'ИНН' or 'Юридический адрес'
        if h4_tag and h4_tag.text.strip() in ['ИНН', 'Полное наименование', 'Юридический адрес', 'Сайт', 'Руководитель']:
            # Find the next sibling <p> element
            p_tag = h4_tag.find_next_sibling('p')

            # Extract the text from the <p> element and add it to the results dictionary
            if p_tag:
                requisites[h4_tag.text.strip()] = p_tag.get_text(strip=True)

    company_info = [company_name]
    for requisite, value in requisites.items():
        if requisite in ['ИНН', 'Полное наименование', 'Сайт', 'Руководитель']:
            company_info.append(value)
    ws.append(company_info)
    print(company_info)
    # for key, value in requisites.items():
    #     ws.append(list(value))
    # wb.save(filename='example.xlsx')

wb.save(filename='example.xlsx')
