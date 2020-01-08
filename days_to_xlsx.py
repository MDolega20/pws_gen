#-*- coding: utf-8 -*-
import xlsxwriter
import json
import time
from datetime import datetime
import locale
from calendar import monthrange

locale.setlocale(locale.LC_ALL, "Polish_Poland.1250")
path_output = './resources/output.xlsx' # TODO set custom name
workbook = xlsxwriter.Workbook(path_output)
worksheet = workbook.add_worksheet()

working_days = []

def str_e(txt):
    return str(txt.encode("UTF-8"))

print("Czytanie jsons")
path_input = './resources/input_xlsx.json'
with open(path_input) as json_file:
    working_days = json.load(json_file)["days"]
print("Wczytano " + str(len(working_days)) + " dni pracujacych")

print("Wyznaczanie miesiaca")

month = datetime.strptime(working_days[0]["date"], '%d-%m-%Y').strftime("%B")

print("Tworzenie struktury arkusza")

headers = workbook.add_format({
    'bold': 2,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': True})
summary = workbook.add_format({
    'bold': 2,
    'border': 1,
    'align': 'right',
    'text_wrap': True})
day_normal = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter'})
day_weekend = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'fg_color': 'yellow'})

worksheet.merge_range('A1:D1', month.capitalize()+' '+str(working_days[0]["date"][6:10]), headers)
worksheet.merge_range('A2:A4', 'Data', headers)
worksheet.merge_range('B2:C2', 'Mateusz Dołęga', headers)
worksheet.merge_range('D2:D4', 'Podpis', headers)
worksheet.merge_range('B3:B4', 'godziny pracy', headers)
worksheet.merge_range('C3:C4', 'ilość godzin', headers)

worksheet.set_column('A:A', 5)
worksheet.set_column('B:B', 20)
worksheet.set_column('C:C', 9)
worksheet.set_column('D:D', 10)
worksheet.set_row(0, 20)
worksheet.set_row(1, 20)

print("Wypełnianie arkusza danymi")

def find_day(date):
    for index_day, day in enumerate(working_days):
        if day["date"] == date:
            return index_day
    return False
def calc_hours(day):
    t1 = datetime.strptime(day["time"][0], '%H:%M')
    t2 = datetime.strptime(day["time"][1], '%H:%M')
    return (t2 - t1).seconds/3600

for index_day, day_a in enumerate(range(1,int(monthrange(int(working_days[0]["date"][6:10]), int(working_days[0]["date"][4:5]))[1])), start=1):
    day_nr = index_day
    row_add = 4
    date = ("0" if day_nr < 10 else '') + str(day_nr)+"-"+working_days[0]["date"][3:5]+"-"+working_days[0]["date"][6:10]
    ex_day = find_day(date)
    weekend = True if datetime.strptime(date, '%d-%m-%Y').weekday() >= 5 else False
    if ex_day is not False:
        ex_day = working_days[ex_day]
        worksheet.write('B'+str(row_add+index_day),str(ex_day["time"][0] + " - " + ex_day["time"][1]),day_normal if weekend is False else day_weekend)
        worksheet.write('C'+str(row_add+index_day),calc_hours(ex_day),day_normal if weekend is False else day_weekend)
    else:
        worksheet.write('B'+str(row_add+index_day),'',day_normal if weekend is False else day_weekend)
        worksheet.write('C'+str(row_add+index_day),'',day_normal if weekend is False else day_weekend)
    worksheet.write('D'+str(row_add+index_day),'',day_normal if weekend is False else day_weekend)

    worksheet.write('A'+str(row_add+index_day),str(day_nr),day_normal if weekend is False else day_weekend)

print("Podsumowanie")
row = int(monthrange(int(working_days[0]["date"][6:10]), int(working_days[0]["date"][4:5]))[1]) + 4
worksheet.merge_range('A'+str(row)+':B'+str(row), 'RAZEM:', summary)
# hours = 0
# for day in working_days:
#     hours += calc_hours(day)

worksheet.write_formula('C'+str(row), '=SUM(C'+str(5)+':C'+str(row-1)+')',summary)
# worksheet.write('C'+str(row),hours,summary)
worksheet.write('D'+str(row),'',summary)


print("Zamykanie arkusza")

workbook.close()