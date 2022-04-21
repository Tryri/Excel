import csv
import openpyxl
import re

filename = '111.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb.active
a = []
b = []

for n in sheet["A"]:
    if n.value != None:
        b.append(n.value)

b = set(b)
b = list(b)
k = 1
m = len(b)

print(len(b))
# for i in b:
#     m -= 1
#     print(m)
#     a = []
#     for n in sheet['A']:
#         if n.value == i and n.value != None:
#             a.append(sheet[f"B{n.row}"].value)
#     sheet[f'C{k}'].value = f'{i} - {a}'
#     print(a)
#     k += 1

for i in b:
    m -= 1
    print(m)
    a = []
    for n in sheet['A']:
        if n.value == i and n.value != None:
            a.append(sheet[f"B{n.row}"].value)

    sheet[f'C{k}'].value = f'{i}_{len(a)}'
    # print(a)
    k += 1
#
# print(len(a))
wb.save("111.xlsx")

