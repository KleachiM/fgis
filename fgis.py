import requests
from pprint import pprint
import os
import xlwings as xw
import sys
import datetime
import tqdm

from date_trans import get_date
from url_trans import make_url

DAYS_DIFF = 62

proxies = {
  'http': 'http://gate.inet:3128',
  'https': 'http://gate.inet:3128',
}

URL_TEST = 'https://fgis.gost.ru/fundmetrology/cm/results/'
try:
    resp = requests.get(URL_TEST, proxies=proxies)
except requests.exceptions.ConnectionError:
    print('Нет соединения с сервером')
    sys.exit(0)

file_name = 'аршин.xlsm'  # имя файла в текущей папке
# full_path = os.path.join(os.getcwd(), file_name)

dir = r'C:\Users\PycharmProjects\FGIS'  # hardcode для корректного запуска скрипта из экселя

full_path = dir + '\\' + file_name

wb = xw.Book(full_path)
sh1 = wb.sheets[0]
sh2 = wb.sheets[1]

start_row = 2
end_row = sh1.range('A1').current_region.last_cell.row  # столбец для анализа количества строк

for excel_line in tqdm.tqdm(range(start_row, end_row), desc='Выполнение'):  # tqdm для отображения прогрессбара
    mitnumber = sh1.range(f'B{excel_line}').value  # столбец для номера в госреестреq
    number = sh1.range(f'E{excel_line}').value  # столбец для заводского номера
    dates = get_date(DAYS_DIFF)
    if dates == 3:  # если 62 дня назад был этот год
        URL = make_url(mitnumber, number, dates['date_from'], dates['date_to'], dates['verification_year'])
        URLS = [URL]
    else:  # если 62 дня назад был прошлый год
        URL1 = make_url(mitnumber, number, dates['past_year_from'], dates['past_year_to'], dates['past_verification_year'])
        URL2 = make_url(mitnumber, number, dates['this_year_from'], dates['this_year_to'], dates['this_verification_year'])
        URLS = [URL1, URL2]

    for URL in URLS:
        resp = requests.get(URL, proxies=proxies)
        resp_json = resp.json()
        is_response = resp_json.get('response')
        works_count = is_response.get('numFound')
        works = is_response.get('docs')
        if works_count:
            for work in works:
                NEW_URL = 'https://fgis.gost.ru/fundmetrology/cm/iaux/vri/' + work['vri_id']
                work_res = requests.get(NEW_URL, proxies=proxies)
                work_res_json = work_res.json()
                blank_line = sh2.range('G1').current_region.last_cell.row + 1
                owner = work_res_json['result']['vriInfo'].get('miOwner')  # получить владельца если есть
                doc_num = work['result_docnum']  # получить номер документа
                type_si = work['mi.modification']  # получить тип СИ
                name_si = work['mi.mititle']  # получить наименование СИ
                reg_num = work['mi.mitnumber']  # получить номер в госреестре
                si_num = work['mi.number']  # получить заводской номер СИ
                manufact_year = work_res_json['result']['miInfo']
                verif_date = work['verification_date'].split('T')[0]  # получить дату поверки
                valid_date = work.get('valid_date')  # получить дату следующей поверки, если она есть
                if valid_date:
                    valid_date = valid_date.split('T')[0]

                is_etalon = work_res_json['result']['miInfo'].get('etaMI')  # эталон/не эталон

                if number == si_num:  # если полученный заводской номер совпадает с номером из экселя
                    # запись данных в эксель
                    sh2.range(f'B{blank_line}').value = owner
                    sh2.range(f'C{blank_line}').value = reg_num
                    sh2.range(f'D{blank_line}').value = name_si
                    sh2.range(f'E{blank_line}').value = type_si
                    sh2.range(f'F{blank_line}').value = si_num
                    sh2.range(f'G{blank_line}').value = doc_num
                    sh2.range(f'K{blank_line}').value = verif_date
                    sh2.range(f'L{blank_line}').value = valid_date
                    sh2.range(f'M{blank_line}').value = datetime.date.today()

                    if is_etalon:
                        # запись данных относящихся только к эталонам
                        si_category = 'эталон'
                        etalon_reg_num = is_etalon['regNumber']
                        etalon_schema = is_etalon['schemaTitle']
                        etalon_rank = is_etalon['rankCode']

                        sh2.range(f'A{blank_line}').value = 'эталон'
                        sh2.range(f'H{blank_line}').value = etalon_reg_num
                        sh2.range(f'I{blank_line}').value = etalon_schema
                        sh2.range(f'J{blank_line}').value = etalon_rank

