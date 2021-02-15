import requests
from pprint import pprint
import sys
import xlwings
import os

#
# URL = 'https://fgis.gost.ru/fundmetrology/cm/results'
#
# params = {
#     'filter_mi_mitype': '8508',
#     'filter_verification_date_start': '2021-01-20',
#     'filter_verification_date_end': '2021-01-01'
# }
# URL = 'https://fgis.gost.ru/fundmetrology/cm/icdb/vri/select?fq=verification_year:2021&\
# fq=mi.mitype:*8508*&fq=verification_date:[2021-01-10T00:00:00Z%20TO%202021-02-01T23:59:59Z]&\
# q=*&fl=vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,result_docnum&\
# sort=verification_date+desc,org_title+asc&rows=50000'

file_name = 'test.xlsx'
full_path = os.path.join(os.getcwd(), file_name)
wb = xw.Book(full_path)
sh = wb.sheets['Лист1']

try:
    resp = requests.get(URL)
except requests.exceptions.ConnectionError:
    print('Нет соединения с сервером')
    sys.exit(0)

resp_json = resp.json()

headers = resp_json['responseHeader']

print(resp_json['response']['numFound'])

# URL1 = 'https://fgis.gost.ru/fundmetrology/cm/icdb/vri/select'
# resp = requests.get(URL1, headers=headers)
# print(resp)

# wb.save()
# xw.apps.active.quit()
