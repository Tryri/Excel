import csv
import openpyxl
import re

filename = 'Список для НО РФКР.xlsx'
wb = openpyxl.load_workbook(filename)
sheet = wb['test']
# row = sheet['B2'].value
# a = re.split(r' ул', r' пер',  row)
# print(a)
splits = (' ул', ' пер')
for row in sheet['A']:
    a = row.value
    a = re.sub(' ул', '', a)
    a = re.sub(r' \bул ', '', a)
    # a = ' '.join(a)
    a = re.sub(r' \bпер ', ' ', a)
    # a = ' '.join(a)
    a = re.sub(r' \bшоссе', '', a)
    # a = ' '.join(a)
    a = re.sub(r' \bбульвар', '', a)
    # a = ' '.join(a)
    a = re.sub(r' \bплощадь', '', a)
    # a = ' '.join(a)
    a = re.sub(r' \bпроезд', '', a)
    # a = ' '.join(a)
    a = re.sub(r'\d корп.\d', '' a)
    print(a)


    # print(f'в строке {n} значение {a}')
# for row in sheet['B']:
#     a = re.split(' пер', row.value)
#     print(a)
# print(ulica, dom)
# wb.save('1.xlsx')

