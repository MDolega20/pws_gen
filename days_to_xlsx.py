import xlsxwriter
import json
import time
from datetime import datetime

summary = xlsxwriter.Workbook('output.xlsx')
worksheet = summary.add_worksheet()

working_days = []

print("Reading jsons")
with open('./resources/input_xlsx.json') as json_file:
    working_days = json.load(json_file)
print("Wczytano " + str(len(working_days)) + " dni pracujących")


