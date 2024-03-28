from bs4 import BeautifulSoup
import openpyxl
import re
def skippin(your_string):
    words_list = ["Бренд", "Акции", "Наименование", "Вывод данных", "Госреестр СИ", "дюймы", "Исполнение", "Серия"]
    if any(word in your_string for word in words_list):
        return True
    else:
        return 
html_file_path = 'C:/Users/vladk/Desktop/lazy_parse/target.html'
with open(html_file_path, 'r', encoding='utf-8') as file:
    html_content = file.read()
soup = BeautifulSoup(html_content, 'html.parser')
regex_pattern = re.compile(r'\b[A-ZА-Я][A-ZА-Я]+(?:[-][A-ZА-Я0-9]+)?\b')
excel_file_path = "готово2.xlsx"
wb = openpyxl.load_workbook(excel_file_path)
category = soup.find('h1', class_='ty-mainbox-title')
new_sheet = category.text
base_name = "change name"
adder = 0
while True:
    try:
        ws = wb.create_sheet(new_sheet)        
        break
    except ValueError:
        new_sheet = f"{base_name}_{adder}"
        adder += 1
item = 'Мерительный инструмент'
specification = ''
description = ''
number = 1
sub = 0
category = category.text
category= category.strip()
subcategory = ''
skip_next_iteration = 0
tables = soup.find_all('div', class_='long_list')
for table in tables:
    td_cells = table.select('table.featured td')
    specification = ''
    counter = 0
    for td_cell in td_cells:
        td_text = td_cell.text.strip()
        if sub == 1:
            subcategory = td_text
            skip_next_iteration = 0
            sub = 0
            continue
        if skip_next_iteration > 0:
            skip_next_iteration -= 1
            continue
        if skippin(td_text):
            skip_next_iteration = 1
            continue
        if "Вид мерителя:"  in td_text:
            counter += 1
            sub = 1
            continue
        specification += td_text + '    '
    if subcategory == '':
        start_index = specification.find('онструкция:    ')
        if start_index != -1:
            start_index += 15
            end_index = specification.find(' ', start_index)
            subcategory = specification[start_index:end_index]
            print(subcategory)
    subcategory = regex_pattern.sub('', subcategory)
    subcategory = re.sub(r'[([\]),]','', subcategory)
    subcategory = subcategory.strip()
    link = table.find('a', class_='product-title')
    name = link.get_text(strip=True)
    name = regex_pattern.sub('', name)
    name = re.sub(r'[([\])]','', name)
    name = name.strip()
    ws.append([number, name, item, category,subcategory, description,specification])
    number +=1
    wb.save(excel_file_path)