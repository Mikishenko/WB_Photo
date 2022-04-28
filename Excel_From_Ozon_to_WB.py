import openpyxl
import os
import requests

from openpyxl import load_workbook
# path = os.chdir("C:\\Users\\AlexMiki\\Desktop\\ПРОЕКТЫ\\Переезд OZON to WB\\OZON товары")
path = os.chdir("D:\\Dropbox\\LeessonS\\WB_Photo\\OZONE")
new_wb = openpyxl.Workbook()
new_ws = new_wb.active
i = 1
new_ws.cell(row = i, column = 1).value = "ID"
new_ws.cell(row = i, column = 2).value = "Наименование"
new_ws.cell(row = i, column = 3).value = "ЦенаМаркетПлейс"
new_ws.cell(row = i, column = 4).value = "Тип"
new_ws.cell(row = i, column = 5).value = "ШК_OZONE"
new_ws.cell(row = i, column = 6).value = "Вес"
new_ws.cell(row = i, column = 7).value = "Высота"
new_ws.cell(row = i, column = 8).value = "Ширина"
new_ws.cell(row = i, column = 9).value = "Глубина"
new_ws.cell(row = i, column = 10).value = "URL_фото"
new_ws.cell(row = i, column = 11).value = "Производитель"
new_ws.cell(row = i, column = 12).value = "Модель"
new_ws.cell(row = i, column = 13).value = "Тип_2"
i = 2
file_count = 0
for name_file in os.listdir(path):
    file_count +=1
    print (file_count, '\t', name_file)

    wb = openpyxl.open(name_file)
    # получаем активный лист
    wb.active = 4
    ws = wb.active
    # печатаем значение ячейки
    for row in range (4, ws.max_row+1):
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
            new_ws.cell(row = i, column = 1).value = id_oscomp
            new_ws.cell(row = i, column = 2).value = name_product
            new_ws.cell(row = i, column = 3).value = price
            new_ws.cell(row = i, column = 4).value = type_product
            new_ws.cell(row = i, column = 5).value = bar_ozone
            new_ws.cell(row = i, column = 6).value = weight
            new_ws.cell(row = i, column = 7).value = width
            new_ws.cell(row = i, column = 8).value = height
            new_ws.cell(row = i, column = 9).value = depth
            new_ws.cell(row = i, column = 10).value = photo_url
            new_ws.cell(row = i, column = 11).value = brand
            new_ws.cell(row = i, column = 12).value = model
            new_ws.cell(row = i, column = 13).value = type_prod_v2
        else:
            i -= 1

        i += 1


new_wb.save("to_WB.xlsx")
new_wb.close()