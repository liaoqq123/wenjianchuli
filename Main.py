import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from Process import DataFrame
from Modify import SearchFrame


class MainPage:
    def __init__(self, master: tk.Tk):
        self.window = master
        self.window.title('文件处理器 1.0.0')
        self.window.geometry('960x540')
        self.top()

    def top(self):
        #顶部栏创建
        menubar = ttk.Menu(self.window)
        menubar.add_command(label='文件处理', command=self.ProcessPage)
        menubar.add_command(label='文件修改', command=self.ModifyPage)
        self.window['menu'] = menubar

    def ProcessPage(self):
        print("niu")


    def ModifyPage(self):
        print("bi")


if __name__ == '__main__':
    window = tk.Tk()
    MainPage(window)
    window.mainloop()