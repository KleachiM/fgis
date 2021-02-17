import requests
from pprint import pprint
import os
import xlwings as xw
import sys

from date_trans import get_date
from url_trans import make_url

DAYS_DIFF = 30

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

dates = get_date(DAYS_DIFF)

if len(dates) > 3:
    pass  # если входит прошлый год


file_name = 'аршин.xlsm' # имя файла в текущей папке
full_path = os.path.join(os.getcwd(), file_name)
wb = xw.Book(full_path)
sh1 = wb.sheets[0]
sh2 = wb.sheets[1]

start_row = 2
end_row = sh1.range('A1').current_region.last_cell.row

for excel_line in range(start_row, end_row):
    mitnumber = sh1.range(f'B{excel_line}').value  # столбец для номера в госреестре
    number = sh1.range(f'F{excel_line}').value  # столбец для заводского номера
    URL = make_url(mitnumber, number, dates['date_from'], dates['date_to'], dates['verification_year'])
    resp = requests.get(URL, proxies=proxies)
    resp_json = resp.json()
    works_count = resp_json['response']['numFound']
    works = resp_json['response']['docs']
    if works_count:
        for work in works:
            NEW_URL = 'https://fgis.gost.ru/fundmetrology/cm/iaux/vri/' + work['vri_id']
            work_res = requests.get(NEW_URL, proxies=proxies)
            work_res_json = work_res.json()
            blank_line = sh2.range('G1').current_region.last_cell.row + 1
            owner = work_res_json['result']['vriInfo']['miOwner']
            doc_num = work['result_docnum']
            type_si = work['mi.modification']
            name_si = work['mi.mititle']
            reg_num = work['mi.mitnumber']
            si_num = work['mi.number']
            is_etalon = work_res_json['result']['miInfo'].get('etaMI')

            sh2.range(f'B{blank_line}').value = owner
            sh2.range(f'C{blank_line}').value = reg_num
            sh2.range(f'D{blank_line}').value = name_si
            sh2.range(f'E{blank_line}').value = type_si
            sh2.range(f'F{blank_line}').value = si_num
            sh2.range(f'G{blank_line}').value = doc_num


            if is_etalon:
                si_category = 'эталон'
                etalon_reg_num = is_etalon['regNumber']
                etalon_schema = is_etalon['schemaTitle']
                etalon_rank = is_etalon['rankCode']

                sh2.range(f'A{blank_line}').value = 'эталон'
                sh2.range(f'H{blank_line}').value = etalon_reg_num
                sh2.range(f'I{blank_line}').value = etalon_schema
                sh2.range(f'J{blank_line}').value = etalon_rank

# wb.save()
# xw.apps.active.quit()
