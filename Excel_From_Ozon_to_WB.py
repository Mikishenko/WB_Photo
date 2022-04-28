import openpyxl
import os
import requests

# устанавливаем рабочий каталог
# path = os.chdir("C:\\Users\\AlexMiki\\Desktop\\ПРОЕКТЫ\\Переезд OZON to WB\\OZON товары")
path = os.chdir("D:\\Dropbox\\LeessonS\\WB_Photo\\OZONE")
# инициализация Excel-файла и листа для записи
new_wb = openpyxl.Workbook()
new_ws = new_wb.active

# определяем заголовки в таблице для первой строки ( i = 1 )
i = 1
table_tytles = ('ID','Наименование','ЦенаМаркетПлейс','Тип','ШК_OZONE','Вес',
               'Высота','Ширина','Глубина','URL_фото','Производитель','Модель',
               'Тип_2')

# TODO попробовать сделать в виде цикла со списком названий
for col_num in range (0, 13):
    print(table_tytles[col_num])
    new_ws.cell(row = 1, column=col_num+1).value = table_tytles[col_num]
# new_ws.cell(row = i, column = 1).value = "ID"
# new_ws.cell(row = i, column = 2).value = "Наименование"
# new_ws.cell(row = i, column = 3).value = "ЦенаМаркетПлейс"
# new_ws.cell(row = i, column = 4).value = "Тип"
# new_ws.cell(row = i, column = 5).value = "ШК_OZONE"
# new_ws.cell(row = i, column = 6).value = "Вес"
# new_ws.cell(row = i, column = 7).value = "Высота"
# new_ws.cell(row = i, column = 8).value = "Ширина"
# new_ws.cell(row = i, column = 9).value = "Глубина"
# new_ws.cell(row = i, column = 10).value = "URL_фото"
# new_ws.cell(row = i, column = 11).value = "Производитель"
# new_ws.cell(row = i, column = 12).value = "Модель"
# new_ws.cell(row = i, column = 13).value = "Тип_2"
#
# # присваиваем значение строчки для продолжения работы с 2-й строки
# i = 2
# # счётчик для оценки прогресса количества обработаных файлов
# file_count = 0
# # для каждого имени файла из папки запускаем цикл
# for name_file in os.listdir(path):
#     file_count +=1 #выводим на печать номер обрабатываемого файла
#     print (file_count, '\t', name_file) # выводим имя обрабатываемого файла
# # открываем книгу
#     wb = openpyxl.open(name_file)
#     # устанавливаем активный лист ( известный из строгой структуры файла)
#     wb.active = 4
#     ws = wb.active
#     # считываем список значений ячеек с листа, начиная с 4-й строки
#     # заголовки пропускаются
#     for row in range (4, ws.max_row+1):
#         # выполняем проверку на "пустоту" ячейки
#         # если не пустая, то считывается значение переменной
#         if ws[row][1].value :
#             id_oscomp = ws[row][1].value
#             name_product = ws[row][2].value
#             price = ws[row][3].value
#             type_product = ws[row][8].value
#             bar_ozone = ws[row][9].value
#             weight = ws[row][10].value
#             width = ws[row][11].value
#             height = ws[row][12].value
#             depth = ws[row][13].value
#             photo_url = ws[row][14].value
#             brand = ws[row][19].value
#             model = ws[row][20].value
#             type_prod_v2 = ws[row][22].value
#             # полученные значения  записываются в столбцы текущей строки (i)
#             # у новой рабочей книги 'ws'
#             new_ws.cell(row = i, column = 1).value = id_oscomp
#             new_ws.cell(row = i, column = 2).value = name_product
#             new_ws.cell(row = i, column = 3).value = price
#             new_ws.cell(row = i, column = 4).value = type_product
#             new_ws.cell(row = i, column = 5).value = bar_ozone
#             new_ws.cell(row = i, column = 6).value = weight
#             new_ws.cell(row = i, column = 7).value = width
#             new_ws.cell(row = i, column = 8).value = height
#             new_ws.cell(row = i, column = 9).value = depth
#             new_ws.cell(row = i, column = 10).value = photo_url
#             new_ws.cell(row = i, column = 11).value = brand
#             new_ws.cell(row = i, column = 12).value = model
#             new_ws.cell(row = i, column = 13).value = type_prod_v2
#         # если ячейка пустая - счётчик строки для создаваемого файла уменьшается
#         else:
#             i -= 1
#         # увеличиваем значение строки для записи на следующей строке листа
#         i += 1
# по окончании обработки всех файлов сохраняем изменения в созданном файле
new_wb.save("to_WB.xlsx")
# закрываем книгу для исключения ошибки совместного доступа
new_wb.close()
# конец выполнения программы