import openpyxl
import os
import requests

from openpyxl import load_workbook

# path = os.chdir("C:\\Users\\AlexMiki\\Desktop\\ПРОЕКТЫ\\Переезд OZON to WB\\OZON товары")
path = os.chdir("D:\\Dropbox\\LeessonS\\WB_Photo\\OZONE")
for name_file in os.listdir(path):
    print (name_file)
    wb = openpyxl.open(name_file)
    # получаем активный лист
    wb.active = 4
    ws = wb.active
    # печатаем значение ячейки
    for row in range (4, ws.max_row):
        if ws[row][1].value :
            id_oscomp = ws[row][1].value
            name_product = ws[row][2].value
            price = ws[row][3].value
            type_product = ws[row][8].value
            bar_ozone = ws[row][9].value
            weight = ws[row][10].value
            width = ws[row][11].value
            height = ws[row][12].value
            depth = ws[row][13].value
            photo_url = ws[row][14].value
            brand = ws[row][19].value
            model = ws[row][20].value
            type_prod_v2 = ws[row][22].value

            print(id_oscomp, '\t', name_product, '\t', price, '\t',
                  type_product, '\t', bar_ozone, '\t', weight, '\t',
                  width, '\t', height, '\t', depth, '\t', photo_url, '\t',
                  brand, '\t', model, '\t', type_prod_v2)


