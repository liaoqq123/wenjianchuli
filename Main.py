import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from Process import ProcessFrame



class MainPage:
    def __init__(self, master: tk.Tk):
        self.window = master
        self.window.title('文件处理器 1.0.0')
        self.window.geometry('960x540')
        self.root()

    def root(self):
        #封装页面导入
        self.Process = ProcessFrame(self.window)
        self.Process.pack()



if __name__ == '__main__':
    window = tk.Tk()
    MainPage(window)
    window.mainloop()