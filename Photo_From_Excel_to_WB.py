import xlrd
import os
import requests


os.chdir("D:\Dropbox\LeessonS\WB_Photo")
workbook = xlrd.open_workbook("EXCEL_PICTURE.xls")
worksheet = workbook.sheet_by_index(0)
chars = int(input("Введите количество строк в Excel-файле:"))
chars = chars
for i in range(1, chars):
    for j in range(0, 2):
        num_file = 0
        if j == 0 :
            folder_name = "D:\Dropbox\LeessonS\WB_Photo\Archive"+"\\"+str(worksheet.cell_value(i, j)+"\\Photo")
            os.makedirs (folder_name)
            os.chdir(folder_name)
            my_st = str(worksheet.cell_value(i, j + 1))+";"+str(worksheet.cell_value(i, j + 2)+";")
            list_urls = my_st.split(";")
            for x in list_urls:
                if x != "":
                    name_file = str(worksheet.cell_value(i, j)) + "_" + str(num_file) + ".jpg"
                    print(name_file)
                    print(x)
                    url = str(x)
                    p = requests.get(url)
                    out = open(name_file, "wb")
                    out.write(p.content)
                    out.close()
                    num_file += 1


