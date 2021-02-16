import requests
from pprint import pprint
import sys
# import xlwings
import os

from date_trans import get_date

DAYS = 62

dates = get_date(DAYS)

# file_name = 'test.xlsx'
# full_path = os.path.join(os.getcwd(), file_name)
# wb = xw.Book(full_path)
# sh = wb.sheets['Лист1']
start_row = 'find start row '
end_row = 'find row count' # rownum = Range('A1').current_region.last_cell.row

for excel_line in range(start_row, end_row):
    mitnumber = sh.range(f'A{excel_line}').value
    number = sh.range(f'B{excel_line}').value
    if len(dates) > 3:
        pass

URL = 'get from excel'

try:
    resp = requests.get(URL)
except requests.exceptions.ConnectionError:
    print('Нет соединения с сервером')
    # xw.apps.active.quit()
    sys.exit(0)

resp_json = resp.json()

headers = resp_json['responseHeader']

print(resp_json['response']['numFound'])

# wb.save()
# xw.apps.active.quit()
