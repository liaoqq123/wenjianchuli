import os
import xlrd
import shutil
import json
from xlutils.copy import copy
import re


class Method:
    def __init__(self):
        self.readMethod()
        self.copyMethod()
        self.moveMethod()
        self.removeMethod()
        self.readData()
        self.storeData()

    def readMethod(self, start_path, excel_path, start):
        file_all = os.listdir(start_path)
        excel = xlrd.open_workbook(str(excel_path))
        line = 0
        new_excel = copy(excel)
        for file in file_all:
            if os.path.isfile(os.path.join(start_path, file)):
                new_sheet = new_excel.get_sheet(0)
                new_sheet.write(line, int(start), file)
                line += 1
        new_excel.save(excel_path)

    #文件复制方法
    def copyMethod(self, start_path, over_path, file_name):
        for root, dirs, files in os.walk(start_path):
            if os.path.isfile(os.path.join(start_path, file_name)) and file_name in files:
                shutil.copy(os.path.join(start_path, file_name), os.path.join(over_path, file_name))  # 复制文件
                print("%s" % (os.path.join(over_path, file_name)))

    #文件移动方法
    def moveMethod(self, start_path, over_path, file_name):
        for root, dirs, files in os.walk(start_path):
            if os.path.isfile(os.path.join(start_path, file_name)) and file_name in files:
                shutil.move(os.path.join(start_path, file_name), os.path.join(over_path, file_name))  # 移动文件
                print("%s" % (os.path.join(over_path, file_name)))

    def renamingMehod(self, start_path, excel_path, start, target):
        pass
        sheet = xlrd.open_workbook(excel_path).sheet_by_index(0)
        list = []
        for start_line in range(0, int(sheet.nrows)):
            list.append(os.path.join(start_path, sheet.cell_value(start_line, start)))
        print(list)
        for file in os.listdir(start_path):
            if file in list:
                start_name = sheet.cell_value(target_line, start)
                target_name = str(int(sheet.cell_value(target_line, target)))
                file_name = os.path.join(start_path, file)
                # 判断是否文件 以及行数据是否小于表格行数
                if os.path.isfile(file_name):
                    # 用正则匹配替换要替换的词
                    newName = re.sub(start_name, target_name, file)
                    # 重设文件名
                    newFileName = file.replace(file, newName)
                    # 重命名
                    os.rename(os.path.join(start_path, file), os.path.join(start_path, newFileName))
                    print(target_name)

    #文件删除方法
    def removeMethod(self, start_path, file_name):
        for root, dirs, files in os.walk(start_path):
            if file_name in files:
                os.remove(os.path.join(start_path, file_name))  # 删除文件
                print("%s" % (os.path.join(start_path, file_name)))


    #读取json文件数据
    def readData(self, start_path, target_path, excel, start, target):
        with open("data.json", mode="r", encoding="utf-8") as json_file:
            text = json_file.read()
        data = json.loads(text)
        return data

    #写入数据至json文件
    def storeData(self, dict):
        dict_json = json.dumps(dict)
        with open('data.json', 'w') as json_file:
            json_file.write(dict_json)


