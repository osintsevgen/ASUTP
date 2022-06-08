import csv
import os.path
import openpyxl
import glob
from pathlib import Path
from os import path
import shutil
from configparser import ConfigParser
import time

# определяем имя выполняемого скрипта
scr = os.path.basename(__file__)
name_ini = Path(scr).stem
print(name_ini)

try:
    file = name_ini+'.ini'
    config = ConfigParser()
    config.read(file)
    print(file)

    SrcPath = config["General"]["SourcePath"]
    RprtPath = config["General"]["ReportPath"]
    ArchvPath = config["General"]["ArchivePath"]
    ExclTmpNm = config["General"]["ExcelTemplateName"]
    CSVMask = config["General"]["CSVMask"]
    CSVDlmtr = config["General"]["CSVDelimiter"]
    ExclNm = config["General"]["ExcelName"]
    ExclExtnsn = config["General"]["ExcelExtension"]

    CSVStrtRw = config["Mapping"]["CSVStartRow"]
    ExclStrtRw = config["Mapping"]["ExcelStartRow"]
    ClmnCnt = config["Mapping"]["ColumnCount"]
    Clmn1 = config["Mapping"]["Column1"]
    Clmn2 = config["Mapping"]["Column2"]
    Clmn3 = config["Mapping"]["Column3"]
    Clmn4 = config["Mapping"]["Column4"]
    Clmn5 = config["Mapping"]["Column5"]
except KeyError:
    print('Проблемы с .ini файлом')

fl_csv_data = []
fl_csv_data = glob.glob('./*.csv')

for i in range(len(fl_csv_data)):
    path1 = Path(fl_csv_data[i])
    name1 = path1.stem[-11:-9]+"."+path1.stem[22:24]+"."+path1.stem[17:21]  # собираем имя
    path2 = ExclNm + name1 + ExclExtnsn
    csv_data = []
    with open(path1) as file_obj:
        reader = csv.reader(file_obj, delimiter=CSVDlmtr)


        bookExcl = openpyxl.load_workbook(filename=ExclTmpNm)
        sheetExcl = bookExcl.active

        k = int(CSVStrtRw)
        m = int(ExclStrtRw)
        z = int(ClmnCnt)
        l = 0
        print(reader)
        for row in reader:
            csv_data.append(row)
            if (l >= k):
                sheetExcl[m][int(Clmn1)].value = csv_data[l][1]
                sheetExcl[m][int(Clmn2)].value = csv_data[l][2]
                sheetExcl[m][int(Clmn3)].value = csv_data[l][3]
                sheetExcl[m][int(Clmn4)].value = csv_data[l][4]
                sheetExcl[m][int(Clmn5)].value = csv_data[l][5]
                l += 1
                m += 1
            else:
                l +=1
            print(sheetExcl[m][int(Clmn1)].value,sheetExcl[m][int(Clmn2)].value)



    bookExcl.save(path2)

    if path.exists(path1):
        destination_paht1 = ArchvPath
        try:
            new_location = shutil.move(path1, destination_paht1)
            print("% s Csv перемещен в указанное место, % s" % (path1, new_location))
        except shutil.Error:
            print('Проблемы с перемещением файла .csv (уже существует в папке)')

    else:
        print("Файл Csv не существует.")

    if path.exists(path2):
        destination_paht2 = RprtPath
        try:
            new_location = shutil.move(path2, destination_paht2)
            print("% s Excel перемещен в указанное место, % s" % (path2, new_location))
        except shutil.Error:
            print('Проблемы с перемещением файла .xlsx (уже существует в папке)')
    else:
        print("Файл Excel не существует.")

time.sleep(10)