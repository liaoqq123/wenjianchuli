import tkinter as tk
import ttkbootstrap as ttk
from method import method
import xlrd as xl
from ttkbootstrap.constants import *

class BasicsFrame(tk.Frame):
    def __init__(self, window: tk.Tk):
        super().__init__(window)
        self.address = tk.StringVar()
        self.line = tk.IntVar()
        self.column = tk.IntVar()
        self.int = tk.StringVar()
        self.str = tk.StringVar()
        self.intgroup = tk.StringVar()
        self.strgroup = tk.StringVar()

        self.address_data = method.data_basics(self, 0, 1)
        self.line_data = method.data_basics(self, 1, 1)
        self.column_data = method.data_basics(self, 2, 1)
        self.int_data = method.data_basics(self, 3, 1)
        self.str_data = method.data_basics(self, 4, 1)
        self.intgroup_data = method.data_basics(self, 5, 1)
        self.strgroup_data = method.data_basics(self, 6, 1)


        self.basics_show()
        self.basics_input()


    #基础数据展示页面
    def basics_show(self):
        self.data_show = ttk.Frame(self)
        ttk.Label(self.data_show).grid(row=0, pady=10)

        ttk.Label(self.data_show, text=method.data_basics(self, 0, 0) + ':').grid(row=1, column=1, pady=10)
        ttk.Label(self.data_show, text=self.address_data).grid(row=1, column=2, pady=10)

        ttk.Label(self.data_show, text=method.data_basics(self, 1, 0) + ':').grid(row=2, column=1, pady=10)
        ttk.Label(self.data_show, text=int(self.line_data)).grid(row=2, column=2, pady=10)

        ttk.Label(self.data_show, text=method.data_basics(self, 2, 0) + ':').grid(row=3, column=1, pady=10)
        ttk.Label(self.data_show, text=int(self.column_data)).grid(row=3, column=2, pady=10)

        ttk.Label(self.data_show, text=method.data_basics(self, 3, 0) + ':').grid(row=4, column=1, pady=10)
        ttk.Label(self.data_show, text=self.int_data).grid(row=4, column=2, pady=10)

        ttk.Label(self.data_show, text=method.data_basics(self, 4, 0) + ':').grid(row=5, column=1, pady=10)
        ttk.Label(self.data_show, text=self.str_data).grid(row=5, column=2, pady=10)

        ttk.Label(self.data_show, text=method.data_basics(self, 5, 0) + ':').grid(row=6, column=1, pady=10)
        ttk.Label(self.data_show, text=self.intgroup_data).grid(row=6, column=2, pady=10)

        ttk.Label(self.data_show, text=method.data_basics(self, 6, 0) + ':').grid(row=7, column=1, pady=10)
        ttk.Label(self.data_show, text=self.strgroup_data).grid(row=7, column=2, pady=10)

        ttk.Button(self.data_show, text='修改数据', command=self.revise_data).grid(row=8, column=1, pady=10)

        self.data_show.pack()



    #修改数据按钮方法
    def revise_data(self):
        #获取数据展示在输入界面
        self.address.set(self.address_data)
        self.line.set(self.line_data)
        self.column.set(self.column_data)
        self.int.set(self.int_data)
        self.str.set(self.str_data)
        self.intgroup.set(self.intgroup_data)
        self.strgroup.set(self.strgroup_data)

        self.data_show.pack_forget()
        self.input_show.pack()



    #修改基础数据页面
    def basics_input(self):
        self.input_show = ttk.Frame(self)

        ttk.Label(self.input_show).grid(row=0, pady=10)

        ttk.Label(self.input_show, text=method.data_basics(self, 0, 0) + ':').grid(row=1, column=1, pady=10)
        ttk.Entry(self.input_show, text=self.address).grid(row=1, column=2, pady=10)

        ttk.Label(self.input_show, text=method.data_basics(self, 1, 0) + ':').grid(row=2, column=1, pady=10)
        ttk.Entry(self.input_show, text=self.line).grid(row=2, column=2, pady=10)

        ttk.Label(self.input_show, text=method.data_basics(self, 2, 0) + ':').grid(row=3, column=1, pady=10)
        ttk.Entry(self.input_show, text=self.column).grid(row=3, column=2, pady=10)

        ttk.Label(self.input_show, text=method.data_basics(self, 3, 0) + ':').grid(row=4, column=1, pady=10)
        ttk.Entry(self.input_show, text=self.int).grid(row=4, column=2, pady=10)

        ttk.Label(self.input_show, text=method.data_basics(self, 4, 0) + ':').grid(row=5, column=1, pady=10)
        ttk.Entry(self.input_show, text=self.str).grid(row=5, column=2, pady=10)

        ttk.Label(self.input_show, text=method.data_basics(self, 5, 0) + ':').grid(row=6, column=1, pady=10)
        ttk.Entry(self.input_show, text=self.intgroup).grid(row=6, column=2, pady=10)

        ttk.Label(self.input_show, text=method.data_basics(self, 6, 0) + ':').grid(row=7, column=1, pady=10)
        ttk.Entry(self.input_show, text=self.strgroup).grid(row=7, column=2, pady=10)

        ttk.Button(self.input_show, text='确认修改', command=self.agree_revise).grid(row=8, column=1, pady=10)

    #确认修改按钮方法
    def agree_revise(self):
        input_data = [self.address.get(),
                      self.line.get(),
                      self.column.get(),
                      self.int.get(),
                      self.str.get(),
                      self.intgroup.get(),
                      self.strgroup.get()]
        for line in range(0, len(input_data)):
            method.input_basics(self, line, 1, input_data[line])
            print(input_data[line])



        self.data_show.pack()
        self.input_show.pack_forget()
