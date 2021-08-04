import requests
import os
import re
import xlwings as xw
import sys
import datetime
import tqdm
import pandas as pd

from date_trans import get_date
from url_trans import make_url

DAYS_DIFF = 62  # количество дней

PATTERN_SI_START = '\(IDSI:\d+\)'
PATTERN_RZ_START = '\(IDRZ:\d+\)'
PATTERN_FINAL = '\d+'

def find_id_si(val):
    id_txt = re.findall(PATTERN_SI_START, val)[0]
    return re.findall(PATTERN_FINAL, id_txt)[0]


def find_id_rz(val):
    rz_txt = re.findall(PATTERN_RZ_START, val)[0]
    return re.findall(PATTERN_FINAL, rz_txt)[0]


def get_works(URL):

    resp = requests.get(URL, proxies=proxies)
    resp_json = resp.json()
    is_response = resp_json.get('response')
    works_count = is_response.get('numFound')
    works = is_response.get('docs')
    if works_count:
        return works
    else:
        return False


def get_data(vri_id):
    NEW_URL = 'https://fgis.gost.ru/fundmetrology/cm/iaux/vri/' + vri_id + '?nonpub=1'
    work_res = requests.get(NEW_URL, proxies=proxies)
    work_res_json = work_res.json()
    is_etalon = work_res_json['result']['miInfo'].get('etaMI')  # эталон/не эталон

    etalon_reg_num = is_etalon['regNumber'] if is_etalon else ''  # получить регистрационный номер эталона
    calibrator = work_res_json['result']['vriInfo'].get('organization')  # получить организацию-поверителя
    reg_num = work['mi.mitnumber']  # получить номер в госреестре
    key = [one for one in work_res_json['result']['miInfo']][0]
    info = work_res_json['result']['miInfo'][f'{key}']
    si_name = info['mitypeTitle']  # получить наименование СИ
    si_type = info['mitypeType']  # получить тип СИ
    si_modification = info['modification']  # получить модификацию
    si_num = str(work['mi.number'])  # получить заводской номер СИ
    si_year = info.get('manufactureYear')  # получить год выпуска если есть
    #TODO: найти путь к гэту
    first_et = ''
    etalon_rank = is_etalon['rankCode'] if is_etalon else ''  #  получить разряд эталона если есть
    verif_date = work['verification_date'].split('T')[0]  # получить дату поверки
    doc_num = work['result_docnum']  # получить номер документа
    applicable = 'Да' if doc_num.startswith('С') else 'Нет'
    valid_date = work.get('valid_date')  # получить дату следующей поверки, если она есть
    owner = work_res_json['result']['vriInfo'].get('miOwner')  # получить владельца если есть
    add_in_url = work.get('vri_id')
    url = f'{URL_}{add_in_url}' if add_in_url else ''  # добавление ссылки на проведенную поверку


    return {
        'etalon_reg_num': etalon_reg_num,
        'calibrator': calibrator,
        'reg_num': reg_num,
        'si_name': si_name,
        'si_type': si_type,
        'si_modification': si_modification,
        'si_num': si_num,
        'si_year': si_year,
        'first_et': first_et,
        'etalon_rank': etalon_rank,
        'verif_date': verif_date,
        'applicable': applicable,
        'valid_date': valid_date,
        'doc_num': doc_num,
        'owner': owner,
        'url': url
    }


proxies = {
  'http': 'http://gate.inet:3128',
  'https': 'http://gate.inet:3128',
}

URL_ = 'https://fgis.gost.ru/fundmetrology/cm/results/'
try:
    resp = requests.get(URL_, proxies=proxies)
except requests.exceptions.ConnectionError:
    print('Нет соединения с сервером')
    sys.exit(0)

file_name = 'аршин.xlsm'  # имя файла в текущей папке
# full_path = os.path.join(os.getcwd(), file_name)

dir = r'C:\Users\PycharmProjects\FGIS'  # hardcode для корректного запуска скрипта из экселя

full_path = dir + '\\' + file_name

# wb = xw.Book(full_path)
# sh1 = wb.sheets['Для_поиска']
# sh2 = wb.sheets['Найденное']
#
# start_row = 2
# end_row = sh1.range('A1').current_region.last_cell.row  # столбец для анализа количества строк
# row = 0

df = pd.read_excel(full_path, usecols='A:G')

columns = ['регистрационный_номер_эталона', 'организация_поверитель', 'госреестр', 'наименование', 'тип', 'модификация', 'заводской_номер',
           'год_выпуска', 'гэт', 'разряд', 'дата_поверки', 'пригодность', 'срок_действия', 'номер_записи', 'владелец', 'url']

res_df = pd.DataFrame(columns=columns)


for line in tqdm.tqdm(df.index, desc='Выполнение'):  # tqdm для отображения прогессбара

    mitnumber = df.loc[line, 'Госреестр']
    number = str(df.loc[line, 'Заводской номер'])

    dates = get_date(DAYS_DIFF)

    if len(dates) == 3:  # если 62 дня назад был этот год
        URL = make_url(mitnumber, number, dates['date_from'], dates['date_to'], dates['verification_year'])
        URLS = [URL]
    else:  # если 62 дня назад был прошлый год
        URL1 = make_url(mitnumber, number, dates['past_year_from'], dates['past_year_to'], dates['past_verification_year'])
        URL2 = make_url(mitnumber, number, dates['this_year_from'], dates['this_year_to'], dates['this_verification_year'])
        URLS = [URL1, URL2]

    for URL in URLS:

        works = get_works(URL)

        if works:
            for work in works:

                result = get_data(work['vri_id'])

                if number == result['si_num']:

                    if res_df.index.empty:
                        res_line = 0
                    else:
                        res_line = res_df.index.max() + 1

                    res_df.loc[res_line, 'регистрационный_номер_эталона'] = result['etalon_reg_num']
                    res_df.loc[res_line, 'организация_поверитель'] = result['calibrator']
                    res_df.loc[res_line, 'госреестр'] = result['reg_num']
                    res_df.loc[res_line, 'наименование'] = result['si_name']
                    res_df.loc[res_line, 'тип'] = result['si_type']
                    res_df.loc[res_line, 'модификация'] = result['si_modification']
                    res_df.loc[res_line, 'заводской_номер'] = result['si_num']
                    res_df.loc[res_line, 'год_выпуска'] = result['si_year']
                    res_df.loc[res_line, 'гэт'] = result['first_et']
                    res_df.loc[res_line, 'разряд'] = result['etalon_rank']
                    res_df.loc[res_line, 'дата_поверки'] = result['verif_date']
                    res_df.loc[res_line, 'пригодность'] = result['applicable']
                    res_df.loc[res_line, 'срок_действия'] = result['valid_date']
                    res_df.loc[res_line, 'номер_записи'] = result['doc_num']
                    res_df.loc[res_line, 'владелец'] = result['owner']
                    res_df.loc[res_line, 'url'] = result['url']

res_df.to_excel(r'C:\Users\BryuhovaME\Desktop\найденные_внешние_поверки.xls', index=False)

# for excel_line in tqdm.tqdm(range(start_row, end_row), desc='Выполнение'):  # tqdm для отображения прогрессбара
#     mitnumber = sh1.range(f'B{excel_line}').value  # столбец для номера в госреестреq
#     number = str(sh1.range(f'E{excel_line}').value)  # столбец для заводского номера
#
#     dates = get_date(DAYS_DIFF)
#
#     if len(dates) == 3:  # если 62 дня назад был этот год
#         URL = make_url(mitnumber, number, dates['date_from'], dates['date_to'], dates['verification_year'])
#         URLS = [URL]
#     else:  # если 62 дня назад был прошлый год
#         URL1 = make_url(mitnumber, number, dates['past_year_from'], dates['past_year_to'], dates['past_verification_year'])
#         URL2 = make_url(mitnumber, number, dates['this_year_from'], dates['this_year_to'], dates['this_verification_year'])
#         URLS = [URL1, URL2]
#
#     for URL in URLS:
#         resp = requests.get(URL, proxies=proxies)
#         resp_json = resp.json()
#         is_response = resp_json.get('response')
#         works_count = is_response.get('numFound')
#         works = is_response.get('docs')
#         if works_count:
#             for work in works:
#                 add_in_url = work.get('vri_id')
#                 NEW_URL = 'https://fgis.gost.ru/fundmetrology/cm/iaux/vri/' + work['vri_id'] + '?nonpub=1'
#                 work_res = requests.get(NEW_URL, proxies=proxies)
#                 work_res_json = work_res.json()
#                 # blank_line = sh2.range('G1').current_region.last_cell.row + 1
#                 owner = work_res_json['result']['vriInfo'].get('miOwner')  # получить владельца если есть
#                 doc_num = work['result_docnum']  # получить номер документа
#                 type_si = work['mi.modification']  # получить тип СИ
#                 name_si = work['mi.mititle']  # получить наименование СИ
#                 reg_num = work['mi.mitnumber']  # получить номер в госреестре
#                 si_num = str(work['mi.number'])  # получить заводской номер СИ
#
#                 # получение IDSI и IDRZ
#                 result_info = work_res_json['result'].get('info')
#                 if result_info:
#                     try:
#                         additional_info = result_info['additional_info']
#                     except KeyError:
#                         additional_info = ''
#                 if additional_info:
#                     pass
#
#                 try:
#                     worker = work_res_json['result']['nonpub'].get('verifiername')
#                 except KeyError:
#                     worker = ''
#                 manufact_year = work_res_json['result']['miInfo']
#                 verif_date = work['verification_date'].split('T')[0]  # получить дату поверки
#                 valid_date = work.get('valid_date')  # получить дату следующей поверки, если она есть
#                 if valid_date:
#                     valid_date = valid_date.split('T')[0]
#
#                 is_etalon = work_res_json['result']['miInfo'].get('etaMI')  # эталон/не эталон
#
#                 if number == si_num:  # если полученный заводской номер совпадает с номером из экселя
#                     # запись данных в эксель
#                     if not row:
#                         blank_line = sh2.range('G1').current_region.last_cell.row + 1
#                         row = blank_line
#                         blank_line = row
#                     else:
#                         row += 1
#                         blank_line = row
#                     sh2.range(f'B{blank_line}').value = owner
#                     sh2.range(f'C{blank_line}').value = reg_num
#                     sh2.range(f'D{blank_line}').value = name_si
#                     sh2.range(f'E{blank_line}').value = type_si
#                     sh2.range(f'F{blank_line}').value = si_num
#                     sh2.range(f'G{blank_line}').value = doc_num
#                     sh2.range(f'K{blank_line}').value = verif_date
#                     sh2.range(f'L{blank_line}').value = valid_date
#                     sh2.range(f'M{blank_line}').value = datetime.date.today()
#                     sh2.range(f'N{blank_line}').value = worker if worker else ''
#                     sh2.range(f'O{blank_line}').value = f'{URL_}{add_in_url}' if add_in_url else ''  # добавление ссылки на проведенную поверку
#
#                     if is_etalon:
#                         # запись данных относящихся только к эталонам
#                         si_category = 'эталон'
#                         etalon_reg_num = is_etalon['regNumber']
#                         etalon_schema = is_etalon['schemaTitle']
#                         etalon_rank = is_etalon['rankCode']
#
#                         sh2.range(f'A{blank_line}').value = 'эталон'
#                         sh2.range(f'H{blank_line}').value = etalon_reg_num
#                         sh2.range(f'I{blank_line}').value = etalon_schema
#                         sh2.range(f'J{blank_line}').value = etalon_rank
