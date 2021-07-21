from urllib.parse import urlunparse
import re


def make_url(mitnumber, number, date_from, date_to, verification_year):
    scheme = 'https'
    netloc = 'fgis.gost.ru'
    path = '/fundmetrology/cm/icdb/vri/select'
    params = ''
    list_for_query = [
        f'fq=verification_year:{verification_year}',
        f'fq=mi.mitnumber:*{add_slashes(mitnumber)}*',
        f'fq=mi.number:*{add_slashes(number)}*',
        f'fq=verification_date:[{date_from}T00:00:00Z%20TO%20{date_to}T23:59:59Z]',
        'q=*',
        'fl=vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,result_docnum',
        'sort=verification_date+desc,org_title+asc',
        'rows=100',
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

