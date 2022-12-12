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
        self.over_path = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.excel = tk.StringVar()
        self.start = tk.StringVar()
        self.target = tk.StringVar()

        self.start_path_data = ""
        self.over_path_data = ""
        self.excel_path_data = ""
        self.excel_data = ""
        self.start_data = ""
        self.target_data = ""

        self.ProcessInputPage()
        self.ProcessShowPage()



    #数据及操作页面显示
    def ProcessShowPage(self):
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
        ttk.Label(self.target_show, text=self.over_path_data).pack(side=tk.RIGHT)

        #表格地址显示
        self.address_show = tk.Frame(self.process_show)
        self.address_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.address_show, text="表格文件地址:").pack(side=tk.LEFT)
        ttk.Label(self.address_show, text=self.excel_path_data).pack(side=tk.RIGHT)

        #表格名称显示
        self.excel_show = tk.Frame(self.process_show)
        self.excel_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.excel_show, text="     表格名称:").pack(side=tk.LEFT)
        ttk.Label(self.excel_show, text=self.excel_data).pack(side=tk.RIGHT)

        #起始数据列显示
        self.start_show = tk.Frame(self.process_show)
        self.start_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.start_show, text="  起始数据列:").pack(side=tk.LEFT)
        ttk.Label(self.start_show, text=self.start_data).pack(side=tk.RIGHT)

        #目标数据列显示
        self.target_show = tk.Frame(self.process_show)
        self.target_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.target_show, text="  目标数据列:").pack(side=tk.LEFT)
        ttk.Label(self.target_show, text=self.target_data).pack(side=tk.RIGHT)

        #修改数据显示
        self.modify_show = tk.Frame(self.process_show)
        self.modify_show.pack(side=TOP, pady=10)
        ttk.Button(self.modify_show, text='修改数据', command=self.InputDataButton).pack()

        #文件操作按钮
        self.button_show = ttk.Frame(self.process_show)
        self.button_show.pack(side=BOTTOM, anchor=tk.S, pady=60)
        ttk.Button(self.button_show, text='文件删除').pack(side=tk.LEFT, padx=5)
        ttk.Button(self.button_show, text='文件复制', command=self.CpoyButton).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.button_show, text='文件移动').pack(side=tk.LEFT, padx=5)
        ttk.Button(self.button_show, text='文件改名').pack(side=tk.LEFT, padx=5)

        self.process_show.pack()


    def InputDataButton(self):

        self.excel_path.set(self.excel_path_data)
        self.excel.set(self.excel_data)
        self.start.set(self.start_data)
        self.target.set(self.target_data)

        self.process_show.pack_forget()
        self.process_input.pack()

    def CpoyButton(self):
        excel = self.excel_data + ".xlsx"
        excel_path = os.path.join(self.excel_path_data, excel)
        excel1 = xlrd.open_workbook(excel_path).sheet_by_index(0)
        for line in range(0, excel1.nrows):
            file_name = excel1.cell(line, int(self.start_data)).value
            Method.CopyMethod(self, self.start_path_data, self.over_path_data, str(file_name))


    def ProcessInputPage(self):
        self.process_input = ttk.Frame(self)

        #文件地址输入
        self.address_show = tk.Frame(self.process_input)
        self.address_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.address_show, text="文件起始地址:").pack(side=tk.LEFT)
        ttk.Entry(self.address_show, text=self.start_path).pack(side=tk.RIGHT)

        #文件地址输入
        self.address_show = tk.Frame(self.process_input)
        self.address_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.address_show, text="文件目标地址:").pack(side=tk.LEFT)
        ttk.Entry(self.address_show, text=self.over_path).pack(side=tk.RIGHT)

        #文件地址输入
        self.address_show = tk.Frame(self.process_input)
        self.address_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.address_show, text="表格文件地址:").pack(side=tk.LEFT)
        ttk.Entry(self.address_show, text=self.excel_path).pack(side=tk.RIGHT)

        #表格名称输入
        self.excel_show = tk.Frame(self.process_input)
        self.excel_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.excel_show, text="   表格名称:").pack(side=tk.LEFT)
        ttk.Entry(self.excel_show, text=self.excel).pack(side=tk.RIGHT)

        #起始数据列输入
        self.start_show = tk.Frame(self.process_input)
        self.start_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.start_show, text="起始数据列:").pack(side=tk.LEFT)
        ttk.Entry(self.start_show, text=self.start).pack(side=tk.RIGHT)

        #目标数据列输入
        self.target_show = tk.Frame(self.process_input)
        self.target_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.target_show, text="目标数据列:").pack(side=tk.LEFT)
        ttk.Entry(self.target_show, text=self.target).pack(side=tk.RIGHT)

        #确认修改按钮
        self.button_show = tk.Frame(self.process_input)
        self.button_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Button(self.button_show, text='确认修改', command=self.DefineDataButton).pack(side=tk.LEFT, padx=5)

    def DefineDataButton(self):
        excel_address = os.path.join(self.excel_path.get(), self.excel.get())
        excel_xlsx = excel_address + ".xlsx"
        excel_xls = excel_address + ".xls"
        if not os.path.exists(self.start_path.get())\
                or not os.path.exists(self.over_path.get())\
                or not os.path.exists(self.excel_path.get()):
            messagebox.showwarning(title="警告", message="请输入正确地址")
        elif not os.path.exists(excel_xlsx) and not os.path.exists(excel_xls):
            messagebox.showwarning(title="警告", message="没有找到表格文件")
        elif not self.start.get().isdigit() or not self.target.get().isdigit():
            messagebox.showwarning(title="警告", message="请输入正确的行列数据")
        else:

            self.start_path_data = self.start_path.get()
            self.over_path_data = self.over_path.get()
            self.excel_path_data = self.excel_path.get()
            self.excel_data = self.excel.get()
            self.start_data = self.start.get()
            self.target_data = self.target.get()

            self.ProcessShowPage()
            self.process_show.pack()
            self.process_input.pack_forget()



