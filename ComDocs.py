import numpy as np
import pandas as pd
from tkinter import *
from tkinter import filedialog as fd
import tkinter as tk
import tkinter.ttk as ttk
from fase_button import FaceButton as FB
# import xlsxwriter
# import xlrd

#Classes

class Application(Frame):
    def __init__(self, master):
        super(Application, self).__init__(master)
        self.initUI()

    def initUI(self):
        self.canvas = Canvas(root, bg='Black')
        self.canvas.place(relx=0, rely=0, relwidth=1, relheight=1)
        # self.canvas = Canvas(self.canvas1, bg='White')
        # self.canvas.place(x=0, y=180, relwidth=1, relheight=1)
        # self.frame1 = Frame(self.canvas1, bg='White')
        # self.frame1.pack()
        self.notebook = ttk.Notebook(self.canvas)
        # self.notebook = ttk.Notebook(self.frame1, width=550, height=151)

        a_tab = FB(self.notebook)
        self.notebook.add(a_tab, text="Notebook A")
        self.notebook.pack()

        # self.frame1.place(relx=0.0, rely=0.0)

        # c = self.frame1.winfo_height()
        # print(c)

        # self.notebooklow.add(l_tab, text="Notebook A")


        self.notebook.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.place(relx=0, rely=0)

root = Tk()

root['bg'] = '#fafafa'
root.title('Сравнение отчетности')
root.minsize(width=550, height=500)
#root.wm_attributes('-alpha', 0.7)
#root.geometry('550x500') # фиксированный размер
#root.resizable(width=False, height=False) # метод запрета растягивания окон

app = Application(root)
app.mainloop()