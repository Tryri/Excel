import openpyxl
import re

filename = '1.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb.active
k = 0
for n in sheet['A']:
    k += 1
    a = sheet[f'A{n.row}'].value
    b = sheet[f'B{n.row}'].value
    c = sheet[f'C{n.row}'].value
    e = sheet[f'D{n.row}'].value
    a = str(a)
    b = str(b)
    c = str(c)
    e = str(e)
    sheet[f'E{n.row}'].value = f'{a}_{b}_{c}_{e}'
    print(a, b, c, e)
print(k)
wb.save("1.xlsx")
# print(m)
