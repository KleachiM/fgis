# from main import URL
from urllib.parse import urlparse, urlunparse
from pprint import pprint

URL = 'https://fgis.gost.ru/fundmetrology/cm/icdb/vri/select?fq=verification_year:2021&\
fq=mi.mitype:*8508*&fq=verification_date:[2021-01-10T00:00:00Z%20TO%202021-02-01T23:59:59Z]&\
q=*&fl=vri_id,org_title,mi.mitnumber,mi.mititle,mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,result_docnum&\
sort=verification_date+desc,org_title+asc&rows=50000'

obj = urlparse(URL)
pprint(obj)
print(len(obj))
print('============')
pprint(obj.query.split('&'))

scheme = 'https'
netloc = 'fgis.gost.ru'
path = '/fundmetrology/cm/icdb/vri/select'
params = ''
query = 'fq=verification_year:2021&fq=mi.mitype:*8508*&fq=verification_date:[' \
                  '2021-01-10T00:00:00Z%20TO%202021-02-01T23:59:59Z]&q=*&fl=vri_id,org_title,mi.mitnumber,mi.mititle,' \
                  'mi.mitype,mi.modification,mi.number,verification_date,valid_date,applicability,' \
                  'result_docnum&sort=verification_date+desc,org_title+asc&rows=50000 '
fragments = ''
maked_URL = (scheme, netloc, path, params, query, fragments)
maked_URL = urlunparse(maked_URL)
pprint(maked_URL)

