import openpyxl
import re

filename = '5.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb['Лист1']
f = sheet.max_row
m = 0
# sheet.insert_cols(2, 1)
for n in range(f):
    c = sheet[f'A{n + 1}'].value
    c = re.split(r',', c)
    sheet[f'A{n + 1}'].value = c[0]
    sheet[f'B{n + 1}'].value = c[1]
    sheet[f'C{n + 1}'].value = c[2]
    print(c)
    m += 1
wb.save("6.xlsx")
print(m)
