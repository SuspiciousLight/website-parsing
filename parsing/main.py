import requests
from bs4 import BeautifulSoup
import re
from openpyxl import Workbook
from openpyxl.styles import Font
import time
import sys


def print_progress_bar(iteration, total, prefix='', suffix='', length=50, fill='█'):
    percent = ("{0:.1f}").format(100 * (iteration / float(total)))
    filled_length = int(length * iteration // total)
    bar = fill * filled_length + '-' * (length - filled_length)
    sys.stdout.write(f'\r{prefix} |{bar}| {percent}% {suffix}')
    sys.stdout.flush()
    if iteration == total:
        print()


def get_school_links(base_url, max_pages=23):
    headers = {'User-Agent': 'Mozilla/5.0'}
    all_schools = []

    print("\nСбор ссылок на школы со всех страниц:")
    for page in range(0, max_pages + 1):
        url = f"{base_url}&page={page}"
        try:
            print_progress_bar(page, max_pages, prefix='Прогресс страниц:', suffix=f'Страница {page}/{max_pages}')
            response = requests.get(url, headers=headers)
            response.raise_for_status()
            soup = BeautifulSoup(response.text, 'html.parser')

            page_schools = []
            for link in soup.select('a.mainlnk[href^="/school/"]'):
                if re.search(r'СОШ|НОШ|НОШИ|ООШ|ЦО|Гимназия|гимназия|школа|Школа|лицей|сош|Лицей', link.text, re.IGNORECASE):
                    full_name = link.text.strip()
                    page_schools.append({
                        'full_name': full_name,
                        'url': "https://doit-together.ru" + link['href']
                    })

            all_schools.extend(page_schools)
            time.sleep(0.5)

        except Exception as e:
            print(f"\nОшибка при обработке страницы {page}: {e}")
            continue

    print(f"\nВсего найдено школ: {len(all_schools)}")
    return all_schools


def parse_school_page(school_url, full_name):
    headers = {'User-Agent': 'Mozilla/5.0'}
    try:
        response = requests.get(school_url, headers=headers)
        response.raise_for_status()
        soup = BeautifulSoup(response.text, 'html.parser')

        curator_blocks = soup.select('.staff.for_school .contact_block')
        all_curators = []

        for block in curator_blocks:
            curator_info = {
                'ФИО': 'N/A',
                'Должность': 'N/A',
                'Телефон': 'N/A',
                'Email': 'N/A'
            }

            name_tag = block.find('b')
            if name_tag:
                curator_info['ФИО'] = name_tag.get_text(strip=True)

            position_tag = block.find('b', string=lambda t: 'Должность:' in str(t))
            if position_tag:
                curator_info['Должность'] = position_tag.next_sibling.strip()

            phone_tag = block.find('b', string=lambda t: 'Тел:' in str(t))
            if phone_tag:
                curator_info['Телефон'] = phone_tag.next_sibling.strip()

            email_tag = block.select_one('a[href^="mailto:"]')
            if email_tag:
                curator_info['Email'] = email_tag['href'].replace('mailto:', '')

            all_curators.append(curator_info)

        return {
            'Название школы': full_name,
            'Кураторы': all_curators if all_curators else [{
                'ФИО': 'N/A',
                'Должность': 'N/A',
                'Телефон': 'N/A',
                'Email': 'N/A'
            }]
        }

    except Exception as e:
        print(f"\nОшибка при парсинге {school_url}: {e}")
        return None


def save_to_excel(data, filename='schools_data.xlsx'):
    wb = Workbook()
    ws = wb.active
    ws.title = "Школы и кураторы"

    headers = [
        '№',
        'Название школы',
        'ФИО куратора',
        'Должность',
        'Телефон',
        'Email',
        'Ссылка'
    ]
    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True)

    row_num = 1
    for i, school in enumerate(data, start=1):
        for curator in school['Кураторы']:
            row = [
                row_num,
                school['Название школы'],
                curator['ФИО'],
                curator['Должность'],
                curator['Телефон'],
                curator['Email'],
                school['url']
            ]
            ws.append(row)
            row_num += 1

    for column in ws.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(filename)
    print(f"\nДанные сохранены в файл: {filename}")

base_url = "https://doit-together.ru/head/66/?scroll=school_full_list0"
print("Начинаем сбор данных...")

schools = get_school_links(base_url, max_pages=7)

print("\nОбработка информации о школах:")
all_data = []
total_schools = len(schools)

for i, school in enumerate(schools, start=2):
    print_progress_bar(i, total_schools, prefix='Прогресс школ:', suffix=f'{i}/{total_schools}')
    school_details = parse_school_page(school['url'], school['full_name'])
    if school_details:
        school_data = {
            'Название школы': school_details['Название школы'],
            'Кураторы': school_details['Кураторы'],
            'url': school['url']
        }
        all_data.append(school_data)
    time.sleep(0.3)

save_to_excel(all_data)

print("\nОбработка всех данных завершена успешно!")