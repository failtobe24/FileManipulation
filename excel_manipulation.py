
import csv
from unittest import result
import openpyxl




# Apri il file excel in questione: cambi il path in base alle tue esigenze
# excel_path = "D:\\personal\\Esempio_Estrazione_Ldap.xlsx"
excel_path = "/home/alex/Temp_Document.xlsx"
interesting_tags = {
    'cn', 'sn', 'c', 'title', 'postalCode', 'givenName',
    'department', 'streetAddress', 'mail', 'mobile'
}



# Apri il file con il solito pattern di istruzioni.
wb_obj = openpyxl.load_workbook(excel_path)
active_sheet = wb_obj.active


# Raccogli solo le righe che ti interessano (quelle evidenziate in giallo).
# Qui una list comprehension è la cosa più efficiente ed elegante.
# Ci sono molti modi per filtrare, ma (se capisco bene) una regex è la soluzione migliore.
# Contestualmente faccio anche un trim del record pulendolo.


results = []
line_dict = dict()
for row in active_sheet.values:

    if row[0] and ':' in row[0]:
        # possible_key = row[0].partition(':')[0].partition(',')[-1].replace('"', '').strip()
        possible_key = row[0].partition(':')[0].strip()
        possible_value = row[0].rpartition(':')[-1].strip()
    else:
        possible_key = ''
        possible_value = ''


    if row == (None,):
        # riga vuota.
        if line_dict:
            results.append(line_dict)
        line_dict = dict()

    else:
        if possible_key in interesting_tags:
            # breakpoint()
            line_dict[possible_key] = possible_value

if line_dict:
    results.append(line_dict)


with open('final_file.csv', 'w+') as fd:
    dict_writer = csv.DictWriter(fd, interesting_tags, delimiter=';')
    dict_writer.writeheader()
    dict_writer.writerows(results)
