
#检索文件目录下所有表格方法
import os
import re
import shutil, os
import xlrd, xlwt
from xlutils.copy import copy
import glob
import json

# def mycopyfile(srcfile, dstpath, name):  # 复制函数
#     for root, dirs, files in os.walk(dstpath):
#         if os.path.isfile(srcfile) and name in files:
#             fpath, fname = os.path.split(srcfile)  # 分离文件名和路径
#             if not os.path.exists(dstpath):
#                 os.makedirs(dstpath)  # 创建路径
#             shutil.copy(srcfile, dstpath + fname)  # 复制文件
#             print("%s" % (dstpath + fname))
# src_dir = r'F:\格斗火影\100089601\ExportedProject\Assets\Texture2D'
# dst_dir = 'C:/Users/Administrator/Desktop/大蛇丸/'  # 目的路径记得加斜杠
# src_file_list = glob.glob('C:/Users/Administrator/Desktop/大蛇丸特效/*meta*')  # glob获得路径下所有文件，可根据需要修改
# # print(src_file_list)
# for srcfile in src_file_list:
#     mycopyfile(srcfile, dst_dir, 'meta')  # 复制文件


# def search(startplace, overplace, name):
#     filelist = os.listdir(startplace)
#     for file in filelist:
#         file_name = os.path.join(startplace, file)
#         if os.path.isfile(file_name) and name in file and "unity3d" not in file and not os.path.exists(os.path.join(overplace, file)):
#             if not os.path.exists(overplace):
#              os.makedirs(overplace)  # 创建路径
#             shutil.copy(file_name, overplace)
#             # send2trash.send2trash(file_name)
#             print(file_name)
#         elif os.path.isdir(file_name):
#             search(file_name, overplace, name)
#
#
# search(r'F:\火影素材', r'F:\name', 'Name')

# file_address = r'E:\unity工程\call\Assets\ninjaimage_11001101\ExportedProject\Assets\TextAsset'
# for a, aa, line in os.walk(file_address):
#     file_counts = len(line)
# #批量读取文件名字到表格
# def FileName(file_address):
#     file_allname = os.listdir(file_address)
#     line = 0
#     for file_name in file_allname:
#         file_complete_name = os.path.join(file_address, file_name)
#         if os.path.exists(file_complete_name) and line < len(file_allname):
#             file_excel = re.sub('.atlas|.txt|.skel|.meta|.bytes|', "", file_name)
#             print(file_excel)
#             wb = xlrd.open_workbook(r'E:\biaoge\121.xlsx')
#             nwb = copy(wb)
#             nws = nwb.get_sheet(1)
#             nws.write(line, 0, file_excel)
#             nwb.save('121.xlsx')
#             line += 1
# FileName(r'F:\lihui')
#批量修改文件名字
# def NewFileName(file_address, excel):
#     #读取sheet
#     sheet = xlrd.open_workbook(excel).sheet_by_index(0)
#     for file in os.listdir(file_address):
#         for line in range(0, int(sheet.nrows)):
#             a = sheet.cell_value(line, 0)
#             b = str(int(sheet.cell_value(line, 1)))
#             file_name = os.path.join(file_address, file)
#             #判断是否文件夹/文件 以及行数据是否小于表格行数
#             if os.path.isdir(file_name):
#             # if os.path.exists(file_name) and ".spine" in file_name and os.path.isfile(file_name):
#             #     os.remove(file_name)
#                 #用正则匹配替换要替换的词
#                 newName = re.sub(a, b, file)
#                 #重设文件名
#                 newFileName = file.replace(file, newName)
#                 #重命名
#                 os.rename(os.path.join(file_address, file), os.path.join(file_address, newFileName))
#                 line += 1
#                 print(b)
#             # elif os.path.isdir(file_name):
#             #     NewFileName(file_name, excel)
#
# NewFileName(r'F:\lihui\lihui', r'E:\biaoge\121.xlsx')



def check(name,singlePack,id,effects):
    with open("ninjas.json", mode="r", encoding="utf-8") as f:
        text = f.read()
    ninjas = json.loads(text)
    return ninjas
# def NewFileName(file_address, excel):
#     #读取sheet
#     sheet = xlrd.open_workbook(excel).sheet_by_index(0)
#     for file in os.listdir(file_address):
#         for line in range(0, int(sheet.nrows)):

list = check('name', 'singlePack', 'id', 'effects')
a = 0
wb = xlrd.open_workbook(r'C:\Users\Administrator\Desktop\11.xlsx')
for name in list:
    print(name["name"])
    nwb = copy(wb)
    nws = nwb.get_sheet(1)
    nws.write(a, 0, name["name"])
    nwb.save(r'C:\Users\Administrator\Desktop\11.xlsx')
    a += 1