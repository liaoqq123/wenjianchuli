import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from data import DataFrame
from search import SearchFrame
from fill import FillFrame
from tips import TipsFrame
from basics import BasicsFrame


class MainPage:
    def __init__(self, master: tk.Tk):
        self.window = master
        self.window.title('表哥模拟器 1.0.0')
        self.window.geometry('960x540')
        self.top()

    def top(self):
        #导入封装页面
        self.basics = BasicsFrame(self.window)
        self.data = DataFrame(self.window)
        self.search = SearchFrame(self.window)
        self.fill = FillFrame(self.window)
        self.tips = TipsFrame(self.window)
        #顶部栏创建
        menubar = ttk.Menu(self.window)
        menubar.add_command(label='基础数据', command=self.show_basics)
        menubar.add_command(label='数据', command=self.show_data)
        menubar.add_command(label='查询', command=self.show_search)
        menubar.add_command(label='填写', command=self.show_fill)
        menubar.add_command(label='说明', command=self.show_tips)
        self.window['menu'] = menubar

    #基础规则按钮方法
    def show_basics(self):
        self.basics.pack()
        self.data.pack_forget()
        self.search.pack_forget()
        self.fill.pack_forget()
        self.tips.pack_forget()

    #数据按钮方法
    def show_data(self):
        self.basics.pack_forget()
        self.data.pack()
        self.search.pack_forget()
        self.fill.pack_forget()
        self.tips.pack_forget()

    #查询按钮方法
    def show_search(self):
        self.basics.pack_forget()
        self.data.pack_forget()
        self.search.pack()
        self.fill.pack_forget()
        self.tips.pack_forget()

    #填写按钮方法
    def show_fill(self):
        self.basics.pack_forget()
        self.data.pack_forget()
        self.search.pack_forget()
        self.fill.pack()
        self.tips.pack_forget()

    def show_tips(self):
        self.basics.pack_forget()
        self.data.pack_forget()
        self.search.pack_forget()
        self.fill.pack_forget()
        self.tips.pack()


if __name__ == '__main__':
    window = tk.Tk()
    MainPage(window)
    window.mainloop()