import numpy as np
import pandas as pd
from tkinter import *
from tkinter import filedialog as fd
import tkinter as tk
import tkinter.ttk as ttk
from fase_button import FaceButton as FB
from notebook import CustomNotebook as CN




class Application(Frame):
    def __init__(self, master):
        super(Application, self).__init__(master)
        self.initUI()

    def initUI(self):
        self.canvas_head = Canvas(root, bg='AliceBlue', height=35)
        self.canvas_head.place(relx=0, rely=0, relwidth=1)
        self.verLog = Label(self.canvas_head, bg='Azure', text="АО ННВК RSV version 2.3", font=("Verdana", 10))
        self.verLog.place(x=60, y=4)

        self.create_but = Button(self.canvas_head, text='+', bg='CornflowerBlue', font=("Verdana", 8, 'bold'),
                             fg='Black', activebackground='Azure', command=self.create_FB)
        self.create_but.place(x=16, y=4,height=22, width=30)
        self.canvas = Canvas(root, bg='White')
        self.canvas.place(x=0, y=30, relwidth=1, relheight=1)
        self.notebook = CN(self.canvas)
        a_tab = FB(self.notebook)
        self.notebook.add(a_tab, text=" Рабочая область 1   ")
        self.notebook.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.place(relx=0, rely=0)

    c = 1

    def create_FB(self):
        self.c += 1
        c_tab = FB(self.notebook)
        self.notebook.add(c_tab, text=" Рабочая область  " + str(self.c) + "   ")
        self.notebook.place(relx=0, rely=0, relwidth=1, relheight=1)




root = Tk()

root['bg'] = '#fafafa'
root.title('Сравнение отчетности')
root.minsize(width=1800, height=800)
#root.wm_attributes('-alpha', 0.7)
#root.geometry('550x500') # фиксированный размер
#root.resizable(width=False, height=False) # метод запрета растягивания окон

app = Application(root)
app.mainloop()