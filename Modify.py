import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

class SearchFrame(tk.Frame):
    def __init__(self, window: tk.Tk):
        super().__init__(window)
        tk.Label(self, text='查询').pack()

