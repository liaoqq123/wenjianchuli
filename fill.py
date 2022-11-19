import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

class FillFrame(tk.Frame):
    def __init__(self, window):
        super().__init__(window)
        tk.Label(self, text='填写').pack()