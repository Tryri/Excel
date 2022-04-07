import openpyxl
import re

filename = '1.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb.active
k = 0
for n in sheet['E']:
    sheet[f'F{n.row}'].value = sheet[f'E{n.row}'].value
for n in sheet['E']:
    for k in sheet['E']:
        if n.value == k.value and n.value != None:
            print(n.value, k.value)
            sheet.delete_rows(n.row)
wb.save("1.xlsx")
# print(m)
