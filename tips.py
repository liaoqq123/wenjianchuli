import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

class TipsFrame(tk.Frame):
    def __init__(self, window: tk.Tk):
        super().__init__(window)
        tk.Label(self, text='暂时不支持阴间操作').pack()

