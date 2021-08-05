url = """https://fgis.gost.ru/fundmetrology/cm/icdb/mieta/select?q=rmieta_id:296026&
        fl=npenumber"""

res = requests.get(url)

pprint(res.json()['response']['docs'][0]['npenumber'])
