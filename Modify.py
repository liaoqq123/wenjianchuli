import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

#封装页面
class ModifyFrame(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        #数据页面子页面
        self.ModifyPage()


    def ModifyPage(self):
        # 数据显示页面框
        title_data = ('list_name', 'list_sheet_name', 'list_column', 'blank')
        self.show_title = ttk.Treeview(self, show='headings', columns=title_data, bootstyle=INFO)
        self.show_title.column('list_name', width=250, anchor='center')
        self.show_title.column('list_sheet_name', width=150, anchor='center')
        self.show_title.column('list_column', width=100, anchor='center')
        self.show_title.column('blank', width=400, anchor='center')
        self.show_title.heading('list_name', text='表格名称')
        self.show_title.heading('list_sheet_name', text='分表名称')
        self.show_title.heading('list_column', text='表格列数')
        self.show_title.heading('blank', text='')
        self.show_title.pack(fill=tk.BOTH, expand=True, ipady=130)


        #数据显示页面底部按钮
        self.displayButton = ttk.Frame(self)
        ttk.Button(self.displayButton, text='刷新数据').pack(side=tk.LEFT, anchor=tk.N, padx=10)
        ttk.Button(self.displayButton, text='清空数据').pack(side=tk.LEFT, anchor=tk.N, padx=10)
        ttk.Button(self.displayButton, text='检查数据').pack(side=tk.LEFT, anchor=tk.N, padx=10)
        self.displayButton.pack(side=tk.RIGHT, pady=10)
