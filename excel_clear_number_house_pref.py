import openpyxl
import re

filename = '1.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb.active
f = sheet.max_row
k = 0
for n in range(f):
    k += 1
    a = sheet[f'A{n + 1}'].value
    b = sheet[f'B{n + 1}'].value
    c = sheet[f'C{n + 1}'].value
    e = sheet[f'D{n + 1}'].value
    a = str(a)
    b = str(b)
    c = str(c)
    e = str(e)
    a = re.sub(r'^\s', '', a)
    b = re.sub(r'^\s', '', b)
    c = re.sub(r' ', '', c)
    c = re.sub(r'A', "a", c)
    c = re.sub(r'А', "а", c)
    c = re.sub(r'Б', "б", c)
    c = re.sub(r'В', "в", c)
    c = re.sub(r'Г', "г", c)
    e = re.sub(r' ', '', e)
    sheet[f'A{n + 1}'].value = a
    sheet[f'B{n + 1}'].value = b
    sheet[f'C{n + 1}'].value = c
    sheet[f'D{n + 1}'].value = e
    print(a, b, c, e)
print(k)
wb.save("1.xlsx")
# print(m)
