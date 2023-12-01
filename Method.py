import os

import openpyxl
from openpyxl import Workbook, load_workbook
import shutil
import json
import re


class Method:
    def __init__(self):
        self.readMethod()
        self.copyMethod()
        self.moveMethod()
        self.removeMethod()
        self.readData()
        self.storeData()

    #文件写入方法
    def readMethod(self, start_path, excel_path, start, sheetname):
        file_all = os.listdir(start_path)
        excel = load_workbook(str(excel_path))
        sheet = excel[sheetname]
        line = 1
        for file in file_all:
            if os.path.isfile(os.path.join(start_path, file)):
                sheet.cell(line, int(start)).value = file
                line += 1
                print(file)
                excel.save(excel_path)

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
    #文件改名方法
    def renamingMethod(self, start_path, excel_path, start, target, sheetname):
        sheet = load_workbook(excel_path)[sheetname]
        print(type(sheet.max_row))
        for file in os.listdir(start_path):
            for line in range(1, sheet.max_row + 1):
                # 将单元格内容转换为文本
                start_name = str(sheet.cell(line, start).value)
                target_name = str(sheet.cell(line, target).value)
                if start_name in file:
                    file_name = os.path.join(start_path, file)
                    # 判断是否文件 以及行数据是否小于表格行数
                    if os.path.isfile(file_name):
                        # 用正则匹配替换要替换的词
                        newName = re.sub(start_name, target_name, file)
                        # 重设文件名
                        newFileName = file.replace(file, newName)
                        # 重命名
                        os.rename(os.path.join(start_path, file), os.path.join(start_path, newFileName))
                        print(os.path.join(start_path, newFileName))
                else:
                    pass

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


