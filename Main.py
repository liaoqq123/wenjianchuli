import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from Process import ProcessFrame
from Modify import ModifyFrame



class MainPage:
    def __init__(self, master: tk.Tk):
        self.window = master
        self.window.title('文件处理器 1.0.0')
        self.window.geometry('960x540')
        self.top()

    def top(self):
        #封装页面导入
        self.Process = ProcessFrame(self.window)
        self.Process.pack()
        self.Modify = ModifyFrame(self.window)

        # #顶部栏创建
        # menubar = ttk.Menu(self.window)
        # menubar.add_command(label='文件处理', command=self.ProcessPage)
        # menubar.add_command(label='文件修改', command=self.ModifyPage)
        # self.window['menu'] = menubar

    def ProcessPage(self):
        self.Process.pack()
        self.Modify.pack_forget()


    def ModifyPage(self):
        self.Process.pack_forget()
        self.Modify.pack()


if __name__ == '__main__':
    window = tk.Tk()
    MainPage(window)
    window.mainloop()