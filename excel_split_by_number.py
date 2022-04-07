import openpyxl
import re

filename = '1.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb.active
f = sheet.max_row
for n in range(sheet.max_row):
    c = sheet[f'B{n + 1}'].value  # Номера домов
    d = sheet[f'A{n + 1}'].value  # Улицы
    a = str(c)
    b = re.split(r',', a)
    if len(b) > 1:
        for k in b:
            sheet.append([d, k])



wb.save("1.xlsx")