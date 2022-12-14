import tkinter as tk
import ttkbootstrap as ttk
import xlrd
from ttkbootstrap.constants import *
from tkinter import messagebox
import os
from Method import Method


class ProcessFrame(tk.Frame):
    def __init__(self, root):
        super().__init__(root)

        self.start_path = tk.StringVar()
        self.target_path = tk.StringVar()
        self.excel = tk.StringVar()
        self.start = tk.StringVar()
        self.target = tk.StringVar()

        self.pocessInputPage()
        self.processShowPage()

    # 数据及操作页面显示
    def processShowPage(self):
        # 读取json文件数据
        self.list = Method.readData(self, "start_path", "target_path", "excel", "start", "target")
        # 读取字典数据
        self.start_path_data = self.list.get("start_path")
        self.target_path_data = self.list.get("target_path")
        self.excel_data = self.list.get("excel")
        self.start_data = self.list.get("start")
        self.target_data = self.list.get("target")

        self.process_show = ttk.Frame(self)

        # 起始地址显示
        self.target_show = tk.Frame(self.process_show)
        self.target_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.target_show, text="文件起始地址:").pack(side=tk.LEFT)
        ttk.Label(self.target_show, text=self.start_path_data).pack(side=tk.RIGHT)

        # 目标地址显示
        self.target_show = tk.Frame(self.process_show)
        self.target_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.target_show, text="文件目标地址:").pack(side=tk.LEFT)
        ttk.Label(self.target_show, text=self.target_path_data).pack(side=tk.RIGHT)

        # 表格地址显示
        self.address_show = tk.Frame(self.process_show)
        self.address_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.address_show, text="表格文件地址:").pack(side=tk.LEFT)
        ttk.Label(self.address_show, text=self.excel_data).pack(side=tk.RIGHT)

        # 起始数据列显示
        self.start_show = tk.Frame(self.process_show)
        self.start_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.start_show, text="  起始数据列:").pack(side=tk.LEFT)
        ttk.Label(self.start_show, text=self.start_data).pack(side=tk.RIGHT)

        # 目标数据列显示
        self.target_show = tk.Frame(self.process_show)
        self.target_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.target_show, text="  目标数据列:").pack(side=tk.LEFT)
        ttk.Label(self.target_show, text=self.target_data).pack(side=tk.RIGHT)

        # 修改数据显示
        self.modify_show = tk.Frame(self.process_show)
        self.modify_show.pack(side=TOP, pady=10)
        ttk.Button(self.modify_show, text='修改数据', command=self.inputDataButton).pack()

        # 文件操作按钮
        self.button_show = ttk.Frame(self.process_show)
        self.button_show.pack(side=BOTTOM, anchor=tk.S, pady=60)
        ttk.Button(self.button_show, text='文件读取', command=self.readButton).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.button_show, text='文件复制', command=self.cpoyButton).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.button_show, text='文件移动', command=self.moveButton).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.button_show, text='文件改名').pack(side=tk.LEFT, padx=5)
        ttk.Button(self.button_show, text='文件删除', command=self.removeButton).pack(side=tk.LEFT, padx=5)

        self.process_show.pack()

    def inputDataButton(self):
        # 输入框读取数据
        self.start_path.set(self.start_path_data)
        self.target_path.set(self.target_path_data)
        self.excel.set(self.excel_data)
        self.start.set(self.start_data)
        self.target.set(self.target_data)

        self.process_show.pack_forget()
        self.process_input.pack()

    def readButton(self):
        Method.readMethod(self, self.list.get("start_path"), self.list.get("excel"), self.list.get("start"))

    def cpoyButton(self):
        # 打开表格
        excel_data = self.excel_data
        sheet_data = xlrd.open_workbook(excel_data).sheet_by_index(0)
        # 循环遍历表格文件名
        for line in range(0, sheet_data.nrows):
            file_name = sheet_data.cell(line, int(self.start_data)).value
            # 判断起始地址以及目标地址文件是否存在
            if os.path.exists(os.path.join(self.start_path_data, file_name)) and not os.path.exists(
                    os.path.join(self.target_path_data, file_name)):
                Method.copyMethod(self, self.start_path_data, self.target_path_data, str(file_name))
            else:
                pass

    def moveButton(self):
        # 打开表格
        excel_data = self.excel_data
        sheet_data = xlrd.open_workbook(excel_data).sheet_by_index(0)
        # 循环遍历表格文件名
        for line in range(0, sheet_data.nrows):
            file_name = sheet_data.cell(line, int(self.start_data)).value
            # 判断起始地址以及目标地址文件是否存在
            if os.path.exists(os.path.join(self.start_path_data, file_name)) and not os.path.exists(
                    os.path.join(self.target_path_data, file_name)):
                Method.moveMethod(self, self.start_path_data, self.target_path_data, str(file_name))
            else:
                pass

    def removeButton(self):
        # 打开表格
        excel_data = self.excel_data
        sheet_data = xlrd.open_workbook(excel_data).sheet_by_index(0)
        # 循环遍历表格文件名
        for line in range(0, sheet_data.nrows):
            file_name = sheet_data.cell(line, int(self.start_data)).value
            # 判断起始地址以及目标地址文件是否存在
            if os.path.exists(os.path.join(self.start_path_data, file_name)):
                Method.removeMethod(self, self.start_path_data, str(file_name))
            else:
                pass

    def pocessInputPage(self):
        self.process_input = ttk.Frame(self)

        # 文件地址输入
        self.address_show = tk.Frame(self.process_input)
        self.address_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.address_show, text="文件起始地址:").pack(side=tk.LEFT)
        ttk.Entry(self.address_show, text=self.start_path, width=80).pack(side=tk.RIGHT)

        # 文件地址输入
        self.address_show = tk.Frame(self.process_input)
        self.address_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.address_show, text="文件目标地址:").pack(side=tk.LEFT)
        ttk.Entry(self.address_show, text=self.target_path, width=80).pack(side=tk.RIGHT)

        # 表格地址输入
        self.address_show = tk.Frame(self.process_input)
        self.address_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.address_show, text="表格文件地址:").pack(side=tk.LEFT)
        ttk.Entry(self.address_show, text=self.excel, width=80).pack(side=tk.RIGHT)

        # 起始数据列输入
        self.start_show = tk.Frame(self.process_input)
        self.start_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.start_show, text="  起始数据列:").pack(side=tk.LEFT)
        ttk.Entry(self.start_show, text=self.start).pack(side=tk.RIGHT)

        # 目标数据列输入
        self.target_show = tk.Frame(self.process_input)
        self.target_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.target_show, text="  目标数据列:").pack(side=tk.LEFT)
        ttk.Entry(self.target_show, text=self.target).pack(side=tk.RIGHT)

        # 确认修改按钮
        self.button_show = tk.Frame(self.process_input)
        self.button_show.pack(side=TOP, anchor=tk.S, pady=10)
        ttk.Button(self.button_show, text='确认修改', command=self.defineDataButton).pack(side=tk.LEFT, padx=5)

    def defineDataButton(self):
        print(self.excel.get())
        if not os.path.exists(self.start_path.get()) \
                or not os.path.exists(self.target_path.get()):
            messagebox.showwarning(title="警告", message="请输入正确地址")
        elif not os.path.exists(self.excel.get()) \
                or not os.path.splitext(self.excel.get())[1] in [".xlsx", ".xls"]:
            messagebox.showwarning(title="警告", message="没有找到表格文件")
        elif not self.start.get().isdigit() or not self.target.get().isdigit():
            messagebox.showwarning(title="警告", message="请输入正确的行列数据")
        else:
            # 将数据写入字典
            self.list['start_path'] = self.start_path.get()
            self.list['target_path'] = self.target_path.get()
            self.list['excel'] = self.excel.get()
            self.list['start'] = self.start.get()
            self.list['target'] = self.target.get()
            # 将字典写入json
            Method.storeData(self, self.list)

            self.processShowPage()
            self.process_show.pack()
            self.process_input.pack_forget()
