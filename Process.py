import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

class ProcessFrame(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.ProcessInputPage()
        self.ProcessShowPage()
        self.process_show.pack()


    #数据及操作页面显示
    def ProcessShowPage(self):
        self.process_show = ttk.Frame(self)

        #文件地址显示
        self.address_show = tk.Frame(self.process_show)
        self.address_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.address_show, text="   文件地址:").pack(side=tk.LEFT)
        ttk.Label(self.address_show, text="F:\GIT\doc\打包文件").pack(side=tk.RIGHT)

        #表格名称显示
        self.excel_show = tk.Frame(self.process_show)
        self.excel_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.excel_show, text="   表格名称:").pack(side=tk.LEFT)
        ttk.Label(self.excel_show, text="数据表").pack(side=tk.RIGHT)

        #起始数据列显示
        self.start_show = tk.Frame(self.process_show)
        self.start_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.start_show, text="起始数据列:").pack(side=tk.LEFT)
        ttk.Label(self.start_show, text="11").pack(side=tk.RIGHT)

        #目标数据列显示
        self.target_show = tk.Frame(self.process_show)
        self.target_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.target_show, text="目标数据列:").pack(side=tk.LEFT)
        ttk.Label(self.target_show, text="22").pack(side=tk.RIGHT)

        #修改数据显示
        self.modify_show = tk.Frame(self.process_show)
        self.modify_show.pack(side=TOP, pady=10)
        ttk.Button(self.modify_show, text='修改数据', command=self.InputDataButton).pack()

        #文件操作按钮
        self.button_show = ttk.Frame(self.process_show)
        self.button_show.pack(side=BOTTOM, anchor=tk.S, pady=120)
        ttk.Button(self.button_show, text='文件删除', command=self.InputDataButton).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.button_show, text='文件复制', command=self.InputDataButton).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.button_show, text='文件移动', command=self.InputDataButton).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.button_show, text='文件改名', command=self.InputDataButton).pack(side=tk.LEFT, padx=5)


    def InputDataButton(self):
        self.process_show.pack_forget()
        self.process_input.pack()


    def ProcessInputPage(self):
        self.process_input = ttk.Frame(self)

        #文件地址输入
        self.address_show = tk.Frame(self.process_input)
        self.address_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.address_show, text="   文件地址:").pack(side=tk.LEFT)
        ttk.Entry(self.address_show, text="F:\GIT\doc\打包文件").pack(side=tk.RIGHT)

        #表格名称输入
        self.excel_show = tk.Frame(self.process_input)
        self.excel_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.excel_show, text="   表格名称:").pack(side=tk.LEFT)
        ttk.Entry(self.excel_show, text="数据表").pack(side=tk.RIGHT)

        #起始数据列输入
        self.start_show = tk.Frame(self.process_input)
        self.start_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.start_show, text="起始数据列:").pack(side=tk.LEFT)
        ttk.Entry(self.start_show, text="11").pack(side=tk.RIGHT)

        #目标数据列输入
        self.target_show = tk.Frame(self.process_input)
        self.target_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Label(self.target_show, text="目标数据列:").pack(side=tk.LEFT)
        ttk.Entry(self.target_show, text="22").pack(side=tk.RIGHT)

        #确认修改按钮
        self.button_show = tk.Frame(self.process_input)
        self.button_show.pack(side=TOP, anchor=tk.W, pady=10)
        ttk.Button(self.button_show, text='确认修改', command=self.DefineDataButton).pack(side=tk.LEFT, padx=5)

    def DefineDataButton(self):
        self.process_show.pack()
        self.process_input.pack_forget()


