import requests
import os
import re
import xlwings as xw
import sys
import datetime
import tqdm
import pandas as pd
from json.decoder import JSONDecodeError
import time

from date_trans import get_date
from url_trans import make_url

DAYS_DIFF = 62  # количество дней

DELAY_ON = True  # если TRUE, то задержка включена, если FALSE - выключена
# DELAY_ON = False

DELAY = 2  # задержка в секундах

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
    if DELAY_ON:
        time.sleep(DELAY)  # задержка
    resp = requests.get(URL, proxies=proxies)
    try:
        resp_json = resp.json()
    except JSONDecodeError:
        sys.exit(f'status: {resp.status_code}. Неисправность сервера')

    is_response = resp_json.get('response')
    works_count = is_response.get('numFound')
    works = is_response.get('docs')
    if works_count:
        return works
    else:
        return False


def get_data(vri_id):
    NEW_URL = 'https://fgis.gost.ru/fundmetrology/cm/iaux/vri/' + vri_id + '?nonpub=1'
    if DELAY_ON:
        time.sleep(DELAY)  # задержка
    work_res = requests.get(NEW_URL, proxies=proxies)
    try:
        work_res_json = work_res.json()
    except JSONDecodeError:
        sys.exit(f'status: {resp.status_code}. Неисправность сервера')

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
    if is_etalon:
        et_id = str(int(etalon_reg_num.split('.')[-1]))
        url_for_get = f'https://fgis.gost.ru/fundmetrology/cm/icdb/mieta/select?q=rmieta_id:{et_id}&fl=npenumber,schematitle'
        if DELAY_ON:
            time.sleep(DELAY)  # задержка
        res_for_get = requests.get(url_for_get, proxies=proxies)
        res_for_get = res_for_get.json()

        try:
            first_et = res_for_get['response']['docs'][0]['npenumber']  # ГЭТ
        except KeyError:
            first_et = ''
        except IndexError:
            first_et = 'ИСКАТЬ_ВРУЧНУЮ'

        try:
            schematitle = res_for_get['response']['docs'][0]['schematitle']
        except KeyError:
            schematitle = ''
        except IndexError:
            schematitle = 'ИСКАТЬ_ВРУЧНУЮ'

    else:
        first_et = ''
        schematitle = ''
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
        'schematitle': schematitle,
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

dir = r'C:\Users\PycharmProjects\FGIS'  # hardcode для корректного запуска скрипта из экселя

full_path = dir + '\\' + file_name

wb = xw.Book(full_path)
sh1 = wb.sheets['Для_поиска']

df = pd.read_excel(full_path, usecols='A:G')

columns = ['регистрационный_номер_эталона', 'организация_поверитель', 'госреестр', 'наименование', 'тип', 'модификация', 'заводской_номер',
           'год_выпуска', 'гэт', 'разряд', 'дата_поверки', 'пригодность', 'срок_действия', 'номер_записи', 'владелец', 'url', 'ГПС/МП']

res_df = pd.DataFrame(columns=columns)


for line in tqdm.tqdm(df.index, desc='Выполнение'):  # tqdm для отображения прогессбара

    mitnumber = df.loc[line, 'Госреестр']
    number = str(df.loc[line, 'Заводской номер'])

    # для отметки в экселе "найдено/не найдено"
    finded = False
    row = line + 2

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

                # if result == False:
                #     sys.exit('ошибка сервера')

                if number == result['si_num']:

                    finded = True

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
                    res_df.loc[res_line, 'ГПС/МП'] = result['schematitle']
                    # наименование методики добавить сюда

    sh1.range(f'H{row}').value = 'yes' if finded else 'no'

#TODO: передалать формат ячеек для даты в дату
res_df.to_excel(r'C:\Users\BryuhovaME\Desktop\найденные_внешние_поверки.xls', index=False)
