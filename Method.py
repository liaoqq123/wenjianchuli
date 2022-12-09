import os
import xlrd
import shutil


class Method:
    def __init__(self):
        self.CopyMethod()


    def CopyMethod(self, start_path, over_path, file_name):
        for root, dirs, files in os.walk(start_path):
            if os.path.isfile(start_path+file_name) and file_name in files:
                shutil.copy(start_path + file_name, over_path + file_name)  # 复制文件
                print("%s" % (over_path + file_name))