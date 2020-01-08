import xlsxwriter
import json
import time
from datetime import datetime

path_output = './resources/output.xlsx'
workbook = xlsxwriter.Workbook(path_output)
worksheet = workbook.add_worksheet()

working_days = []

print("Czytanie jsons")
path_input = './resources/input_xlsx.json'
with open(path_input) as json_file:
    working_days = json.load(json_file)["days"]
print("Wczytano " + str(len(working_days)) + " dni pracujących")

print("Tworzenie struktury arkusza")

worksheet.merge_range('A1:D1', 'Merged Range', merge_format)

worksheet.write('A1', 'Hello world')


print("Zamykanie arkusza")
workbook.close()