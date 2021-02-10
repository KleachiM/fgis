import requests
from pprint import pprint
import os
import xlwings as xw

# URL_1 = 'https://fgis.gost.ru/fundmetrology/cm/icdb/vri/select?fq=verification_year:2021&fq=mi.mitype:*8508*&fq=verification_date:[2021-01-20T00:00:00Z%20TO%202021-02-01T23:59:59Z]&q=*&fl=vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,result_docnum&sort=verification_date+desc,org_title+asc&rows=20&start=0'
#
# proxies = {
#   'http': 'http://gate.inet:3128',
#   'https': 'http://gate.inet:3128',
# }
#
# resp = requests.get(URL_1, proxies=proxies)
# resp_json = resp.json()
# pprint(resp_json)
file_name = 'test.xlsx'
full_path = os.path.join(os.getcwd(), file_name)
wb = xw.Book(full_path)
sh = wb.sheets['Лист1']
print(sh.range('A1').value)
sh.range("A2").value = 'new_hello'
wb.save()
xw.apps.active.quit()
