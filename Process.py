import tkinter as tk
import ttkbootstrap as ttk
from ttkbootstrap.constants import *

class ProcessFrame(tk.Frame):
    def __init__(self, root):
        super().__init__(root)
        self.ProcessInputPage()
        self.ProcessShowPage()
        self.process_show.pack()


    def ProcessShowPage(self):
        self.process_show = ttk.Frame(self)

        ttk.Label(self.process_show).grid(row=0, pady=10)

        ttk.Label(self.process_show, text="1").grid(row=1, column=1, pady=10)
        ttk.Label(self.process_show).grid(row=1, column=2, pady=10)

        ttk.Label(self.process_show, text="2").grid(row=2, column=1, pady=10)
        ttk.Label(self.process_show).grid(row=2, column=2, pady=10)

        ttk.Label(self.process_show, text="3").grid(row=3, column=1, pady=10)
        ttk.Label(self.process_show).grid(row=3, column=2, pady=10)

        ttk.Label(self.process_show, text="4").grid(row=4, column=1, pady=10)
        ttk.Label(self.process_show).grid(row=4, column=2, pady=10)

        ttk.Label(self.process_show, text="5").grid(row=5, column=1, pady=10)
        ttk.Label(self.process_show).grid(row=5, column=2, pady=10)

        ttk.Label(self.process_show, text="6").grid(row=6, column=1, pady=10)
        ttk.Label(self.process_show).grid(row=6, column=2, pady=10)

        ttk.Label(self.process_show, text="7").grid(row=7, column=1, pady=10)
        ttk.Label(self.process_show).grid(row=7, column=2, pady=10)

        ttk.Button(self.process_show, text='修改数据', command=self.InputDataButton).grid(row=8, column=1, pady=10)

    def InputDataButton(self):
        self.process_show.pack_forget()
        self.process_input.pack()


    def ProcessInputPage(self):
        self.process_input = ttk.Frame(self)

        ttk.Label(self.process_input).grid(row=0, pady=10)

        ttk.Label(self.process_input).grid(row=1, column=1, pady=10)
        ttk.Entry(self.process_input).grid(row=1, column=2, pady=10)

        ttk.Label(self.process_input).grid(row=2, column=1, pady=10)
        ttk.Entry(self.process_input).grid(row=2, column=2, pady=10)

        ttk.Label(self.process_input).grid(row=3, column=1, pady=10)
        ttk.Entry(self.process_input).grid(row=3, column=2, pady=10)

        ttk.Label(self.process_input).grid(row=4, column=1, pady=10)
        ttk.Entry(self.process_input).grid(row=4, column=2, pady=10)

        ttk.Label(self.process_input).grid(row=5, column=1, pady=10)
        ttk.Entry(self.process_input).grid(row=5, column=2, pady=10)

        ttk.Label(self.process_input).grid(row=6, column=1, pady=10)
        ttk.Entry(self.process_input).grid(row=6, column=2, pady=10)

        ttk.Label(self.process_input).grid(row=7, column=1, pady=10)
        ttk.Entry(self.process_input).grid(row=7, column=2, pady=10)

        ttk.Button(self.process_input, text='确认修改', command=self.DefineDataButton).grid(row=8, column=1, pady=10)

    def DefineDataButton(self):
        self.process_show.pack()
        self.process_input.pack_forget()


