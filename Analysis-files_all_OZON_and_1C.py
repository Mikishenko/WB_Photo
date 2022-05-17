import openpyxl
import os

# устанавливаем рабочий каталог, где лежат файлы с OZON
# path = os.chdir("C:\\Users\\AlexMiki\\Desktop\\ПРОЕКТЫ\\Переезд OZON to WB\\OZON товары")
# path = os.chdir("F:\\Dropbox\\LeessonS\\WB_Photo\\OZONE")
path = os.chdir("D:\\Dropbox\\LeessonS\\1С\\")

# РАБОТА С СОЗДАВАЕМЫМ ФАЙЛОМ
# инициализация НОВОГО Excel-файла и листа для записи
# new_wb = openpyxl.Workbook()
# new_ws = new_wb.active

# перечисляем фиксированные имена полей будущей таблицы
table_tytles = ('ID_1C','ID_site_OSCOMP','BAR_OZONE_in_1C','Наименование',
                'Полное Наименование','BAR_OZON_in_OZONE')
# записываем поля в файл
# for col_num in range (0, 5):
#     new_ws.cell(row = 1, column=col_num+1).value = table_tytles[col_num]
#  переменная 'i' - используется в адресации строк (начнём со 2-й строки)
i = 2
# определение рабочих файлов
    # с данными 1С:
wb_1C = openpyxl.open('KATALOG_1C.xlsx')
wb_1C.active = 0 # устанавливаем активный лист
ws_1C = wb_1C.active
    # с данными OZON:
wb_OZONE = openpyxl.open('ALL_catalog_from_OZON_in_one_file.xlsx')
wb_OZONE.active = 0 # устанавливаем активный лист
ws_OZONE = wb_OZONE.active
# считываем список значений ячеек с листа, начиная с 4-й строки
# заголовки пропускаются
ozone_cat = {}
for row_OZONE in range (1, ws_OZONE.max_row+1):
    id_site_in_OZONE = ws_1C[row_OZONE][1].value
    bar_ozone_in_ozone = ws_1C[row_OZONE][5].value
    ozone_cat ['id_site'] = id_site_in_OZONE
    ozone_cat['bar_ozon'] = bar_ozone_in_ozone
print(ozone_cat)

# for row_1C in range (1, ws_1C.max_row+1):
#     id_1C = ws_1C[row_1C][1].value
#     id_site_OSCOMP = ws_1C[row_1C][2].value
#     BAR_OZONE_in_1C = ws_1C[row_1C][3].value
#     name = ws_1C[row_1C][4].value
#     full_name = ws_1C[row_1C][5].value
#     # полученные значения  записываются в столбцы текущей строки (i)
#     # у новой рабочей книги 'ws'
#     new_ws.cell(row=i, column=1).value = id_1C
#     new_ws.cell(row=i, column=2).value = id_site_OSCOMP
#     new_ws.cell(row=i, column=3).value = BAR_OZONE_in_1C
#     new_ws.cell(row=i, column=4).value = name
#     new_ws.cell(row=i, column=5).value = full_name
#     # увеличиваем значение строки для записи на следующей строке листа
#     i += 1
# # по окончании обработки всех файлов сохраняем изменения в созданном файле
# new_wb.save("Analisys_1C_and_OZON.xlsx")
# # закрываем книгу для исключения ошибки совместного доступа
# new_wb.close()
# # конец выполнения программы