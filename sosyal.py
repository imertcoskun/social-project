import csv
import json
import xlrd
from collections import OrderedDict
# -*- coding: utf-8 -*-

# print pd.read_json('dosya.json').drop(columns=["Column1","Column2"]).to_excel('Report.xls')

#excel dosyasini oku

# def dosya_oku():

#     xls_dosyasi = "Report.xls"
#     json_dosyasi = "dosya.json"
#     data2 = []
#     with open (xls_dosyasi) as xlsfile:
#         csvReader = csv.DictReader(xlsfile)
#     for csvrow in csvReader:
#         data2.append(csvrow)
#     tiklayanlar = {
#         "isim" : "Column1",
#         "soyisim" : "Column2"
#     }
#     with open (json_dosyasi, "w") as jsonfile:
#         jsonfile.write(json.dumps)
# dosya_oku()c



wb = xlrd.open_workbook('kayitli_olacakr_rapor.xls')
#print (wb.sheet_names())
def kisiler():

    sh = wb.sheet_by_index(2)

    liste = []

    for rownum in range(0,sh.nrows):
        kisiler = OrderedDict()
        row_values = sh.row_values(rownum)
        kisiler['hebele'] = row_values[0]
        kisiler['hubele'] = row_values[1]
        kisiler['uga'] = row_values[2]
        kisiler['buga'] = row_values[3]
        liste.append(kisiler)

    j = json.dumps(liste)
# Write to file
    with open('dosya.json', 'w') as f:
        f.write(j)
    liste = [dict(t) for t in {tuple(d.items()) for d in liste}]
    with open('dosya.json', 'w', encoding='utf-8') as f:
        json.dump(liste, f, ensure_ascii=False) 
kisiler()


def tiklayanlar():

    sh = wb.sheet_by_index(5)

    liste = []

    for rownum in range(0,sh.nrows):
        kisiler = OrderedDict()
        row_values = sh.row_values(rownum)
        kisiler['hebele'] = row_values[0]
        kisiler['hubele'] = row_values[1]
        kisiler['uga'] = row_values[2]
        kisiler['bugad'] = row_values[11]
        liste.append(kisiler)

    j = json.dumps(liste)
# Write to file
    with open('tiklayanlar.json', 'w') as f:
        f.write(j)
    liste = [dict(t) for t in {tuple(d.items()) for d in liste}]
    with open('tiklayanlar.json', 'w', encoding='utf-8') as f:
        json.dump(liste, f, ensure_ascii=False) 

tiklayanlar()


