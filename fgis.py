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

file_name = 'аршин.xlsm'
full_path = os.path.join(os.getcwd(), file_name)
wb = xw.Book(full_path)
sh1 = wb.sheets[0]
sh2 = wb.sheets[1]

start_row = 2
end_row = sh1.range('A1').current_region.last_cell.row

for excel_line in range(start_row, end_row):
    mitnumber = sh1.range(f'B{excel_line}').value
    number = sh1.range(f'F{excel_line}').value
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
            # TODO доделать анализ: эталон или нет?
            is_etalon = work_res_json['result']['miInfo']
            print(is_etalon)



# wb.save()
# xw.apps.active.quit()
