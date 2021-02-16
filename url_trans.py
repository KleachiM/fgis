from typing import Tuple
from urllib.parse import urlparse, urlunparse
from pprint import pprint
import re


# URL = 'https://fgis.gost.ru/fundmetrology/cm/icdb/vri/select?fq=verification_year:2021&\
# fq=mi.mitype:*8508*&fq=verification_date:[2021-01-10T00:00:00Z%20TO%202021-02-01T23:59:59Z]&\
# q=*&fl=vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,result_docnum&\
# sort=verification_date+desc,org_title+asc&rows=50000'
#
# obj = urlparse(URL)
# pprint(obj)
# print(len(obj))
# print('============')
# pprint(obj.query.split('&'))
from date_trans import get_date


def make_url(mitnumber, number, date_from, date_to, verification_year):
    scheme = 'https'
    netloc = 'fgis.gost.ru'
    path = '/fundmetrology/cm/icdb/vri/select'
    params = ''
    list_for_query = [
        f'fq=verification_year:{verification_year}',
        f'fq=mi.mitnumber:*{mitnumber}*',
        f'fq=mi.number:*{add_slashes(number)}*',
        f'fq=verification_date:[{date_from}T00:00:00Z%20TO%20{date_to}T23:59:59Z]',
        'q=*',
        'fl=vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,result_docnum',
        'sort=verification_date+desc,org_title+asc',
        'rows=20',
        'start=0'
    ]
    query = list_for_query[0]
    for line in list_for_query[1:]:
        query += f'&{line}'
    fragments = ''

    created_url = (scheme, netloc, path, params, query, fragments)
    return urlunparse(created_url)


def add_slashes(text):
    pattern = re.compile(r'\W')

    text = text.split()[0]  # убирает пробелы
    new_text = ''
    for char in text:
        match = re.findall(pattern, char)
        if match:
            char = f'\\{match[0]}'
        new_text = new_text + char
    return new_text
dates = get_date(30)
print(make_url('216690-97', '016806', dates['date_from'], dates['date_to'], dates['verification_year']))
# URL_FGIS = 'https://fgis.gost.ru/fundmetrology/cm/icdb/vri/select?fq=verification_year:2021&fq=mi.mitnumber:*25984\-08*&fq=mi.number:*139661207*&fq=verification_date:[2021-01-20T00:00:00Z%20TO%202021-02-14T23:59:59Z]&q=*&fl=vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,result_docnum&sort=verification_date+desc,org_title+asc&rows=20&start=0'
# print(URL_FGIS)

