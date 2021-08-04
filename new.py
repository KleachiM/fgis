import requests
import sys
import datetime
import tqdm

from date_trans import get_date
from url_trans import make_url

DAYS_DIFF = 60  # количество дней


URL_TEST = 'https://fgis.gost.ru/fundmetrology/cm/results/'
try:
    resp = requests.get(URL_TEST)
except requests.exceptions.ConnectionError:
    print('Нет соединения с сервером')
    sys.exit(0)

import pandas as pd

PATH = '/Users/mihailkalinin/учеба/FGIS/find.ods'

df = pd.read_excel(PATH, engine='odf')
res_df = pd.DataFrame(columns=[
    "тип", "заводской_номер", "госреестр", "регистрационный_номер_эталона", "поверочная_схема"
])


def get_works(URL):

    resp = requests.get(URL)
    resp_json = resp.json()
    is_response = resp_json.get('response')
    works_count = is_response.get('numFound')
    works = is_response.get('docs')
    if works_count:
        return works
    else:
        return False


def get_data(vri_id):
    NEW_URL = 'https://fgis.gost.ru/fundmetrology/cm/iaux/vri/' + work['vri_id'] + '?nonpub=1'
    work_res = requests.get(NEW_URL)
    work_res_json = work_res.json()
    type_si = work.get('mi.modification')  # получить тип СИ
    si_num = str(work['mi.number'])  # получить заводской номер СИ
    reg_num = work['mi.mitnumber']  # получить номер в госреестре

    is_etalon = work_res_json['result']['miInfo'].get('etaMI')  # эталон/не эталон

    etalon_reg_num = is_etalon['regNumber'] if is_etalon else ''
    etalon_schema = is_etalon['schemaTitle'] if is_etalon else ''
    # if is_etalon:
    #     etalon_reg_num = is_etalon['regNumber']
    #     etalon_schema = is_etalon['schemaTitle']
    # else:
    #     etalon_reg_num = ''
    #     etalon_schema = ''

    return {
        'type_si': type_si,
        'si_num': si_num,
        'reg_num': reg_num,
        'etalon_reg_num': etalon_reg_num,
        'etalon_schema': etalon_schema
    }


for line in tqdm.tqdm(df.index, desc='Выполнение'):  # tqdm для отображения прогрессбара
    mitnumber = df.loc[line, 'госреестр']  # столбец для номера в госреестре
    number = df.loc[line, 'заводской номер']  # столбец для заводского номера
    date = df.loc[line, 'дата поверки']
    type = df.loc[line, 'тип']

    dates = get_date(DAYS_DIFF)

    URL = make_url(mitnumber, str(number), date.date(), date.date(), date.year)

    works = get_works(URL)
    if works:
        for work in works:

            result = get_data(work['vri_id'])
            # NEW_URL = 'https://fgis.gost.ru/fundmetrology/cm/iaux/vri/' + work['vri_id'] + '?nonpub=1'
            # work_res = requests.get(NEW_URL)
            # work_res_json = work_res.json()
            # owner = work_res_json['result']['vriInfo'].get('miOwner')  # получить владельца если есть
            # doc_num = work['result_docnum']  # получить номер документа
            # type_si = work['mi.modification']  # получить тип СИ
            # name_si = work['mi.mititle']  # получить наименование СИ
            # reg_num = work['mi.mitnumber']  # получить номер в госреестре
            # si_num = str(work['mi.number'])  # получить заводской номер СИ
            # try:
            #     worker = work_res_json['result']['nonpub'].get('verifiername')
            # except KeyError:
            #     worker = ''
            # manufact_year = work_res_json['result']['miInfo']
            # verif_date = work['verification_date'].split('T')[0]  # получить дату поверки
            # valid_date = work.get('valid_date')  # получить дату следующей поверки, если она есть
            # if valid_date:
            #     valid_date = valid_date.split('T')[0]
            #
            # is_etalon = work_res_json['result']['miInfo'].get('etaMI')  # эталон/не эталон

            if number == result['si_num']:  # если полученный заводской номер совпадает с номером

                if res_df.index.empty:
                    row = 0
                else:
                    row = res_df.index.max() + 1

                res_df.loc[row, 'тип'] = result['type_si']
                res_df.loc[row, 'госреестр'] = result['reg_num']
                res_df.loc[row, 'заводской_номер'] = result['si_num']
                res_df.loc[row, 'регистрационный_номер_эталона'] = result['etalon_reg_num']
                res_df.loc[row, 'поверочная_схема'] = result['etalon_schema']


res_df.to_excel('/Users/mihailkalinin/учеба/FGIS/finded.ods')