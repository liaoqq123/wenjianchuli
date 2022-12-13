import os
import xlrd
import shutil
import json


class Method:
    def __init__(self):
        self.CopyMethod()
        self.MoveMethod()
        self.RemoveMethod()
        self.ReadData()
        self.StoreData()

    #文件复制方法
    def CopyMethod(self, start_path, over_path, file_name):
        for root, dirs, files in os.walk(start_path):
            if os.path.isfile(os.path.join(start_path, file_name)) and file_name in files:
                shutil.copy(os.path.join(start_path, file_name), os.path.join(over_path, file_name))  # 复制文件
                print("%s" % (os.path.join(over_path, file_name)))

    #文件移动方法
    def MoveMethod(self, start_path, over_path, file_name):
        for root, dirs, files in os.walk(start_path):
            if os.path.isfile(os.path.join(start_path, file_name)) and file_name in files:
                shutil.move(os.path.join(start_path, file_name), os.path.join(over_path, file_name))  # 移动文件
                print("%s" % (os.path.join(over_path, file_name)))

    #文件删除方法
    def RemoveMethod(self, start_path, file_name):
        for root, dirs, files in os.walk(start_path):
            if file_name in files:
                os.remove(os.path.join(start_path, file_name))  # 删除文件
                print("%s" % (os.path.join(start_path, file_name)))

    #读取json文件数据
    def ReadData(self, start_path, target_path, excel, start, target):
        with open("data.json", mode="r", encoding="utf-8") as json_file:
            text = json_file.read()
        data = json.loads(text)
        return data

    #写入数据至json文件
    def StoreData(self, dict):
        dict_json = json.dumps(dict)
        with open('data.json', 'w') as json_file:
            json_file.write(dict_json)


