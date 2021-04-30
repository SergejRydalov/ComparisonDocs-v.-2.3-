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

class CustomNotebook(ttk.Notebook):
    """A ttk Notebook with close buttons on each tab"""

    __initialized = False

    def __init__(self, *args, **kwargs):
        if not self.__initialized:
            self.__initialize_custom_style()
            self.__inititialized = True

        kwargs["style"] = "CustomNotebook"
        ttk.Notebook.__init__(self, *args, **kwargs)

        self._active = None

        self.bind("<ButtonPress-1>", self.on_close_press, True)
        self.bind("<ButtonRelease-1>", self.on_close_release)

    def on_close_press(self, event):
        """Called when the button is pressed over the close button"""

        element = self.identify(event.x, event.y)

        if "close" in element:
            index = self.index("@%d,%d" % (event.x, event.y))
            self.state(['pressed'])
            self._active = index

    def on_close_release(self, event):
        """Called when the button is released over the close button"""
        if not self.instate(['pressed']):
            return

        element =  self.identify(event.x, event.y)
        index = self.index("@%d,%d" % (event.x, event.y))

        if "close" in element and self._active == index:
            self.forget(index)
            self.event_generate("<<NotebookTabClosed>>")

        self.state(["!pressed"])
        self._active = None

    def __initialize_custom_style(self):
        style = ttk.Style()
        self.images = (
            tk.PhotoImage("img_close", data='''
                R0lGODlhCAAIAMIBAAAAADs7O4+Pj9nZ2Ts7Ozs7Ozs7Ozs7OyH+EUNyZWF0ZWQg
                d2l0aCBHSU1QACH5BAEKAAQALAAAAAAIAAgAAAMVGDBEA0qNJyGw7AmxmuaZhWEU
                5kEJADs=
                '''),
            tk.PhotoImage("img_closeactive", data='''
                R0lGODlhCAAIAMIEAAAAAP/SAP/bNNnZ2cbGxsbGxsbGxsbGxiH5BAEKAAQALAAA
                AAAIAAgAAAMVGDBEA0qNJyGw7AmxmuaZhWEU5kEJADs=
                '''),
            tk.PhotoImage("img_closepressed", data='''
                R0lGODlhCAAIAMIEAAAAAOUqKv9mZtnZ2Ts7Ozs7Ozs7Ozs7OyH+EUNyZWF0ZWQg
                d2l0aCBHSU1QACH5BAEKAAQALAAAAAAIAAgAAAMVGDBEA0qNJyGw7AmxmuaZhWEU
                5kEJADs=
            ''')
        )

        style.element_create("close", "image", "img_close",
                            ("active", "pressed", "!disabled", "img_closepressed"),
                            ("active", "!disabled", "img_closeactive"), border=10, sticky='')
        style.layout("CustomNotebook", [("CustomNotebook.client", {"sticky": "nswe"})])
        style.layout("CustomNotebook.Tab", [
            ("CustomNotebook.tab", {
                "sticky": "nswe",
                "children": [
                    ("CustomNotebook.padding", {
                        "side": "top",
                        "sticky": "nswe",
                        "children": [
                            ("CustomNotebook.focus", {
                                "side": "top",
                                "sticky": "nswe",
                                "children": [
                                    ("CustomNotebook.label", {"side": "left", "sticky": ''}),
                                    ("CustomNotebook.close", {"side": "left", "sticky": ''}),
                                ]
                        })
                    ]
                })
            ]
        })
    ])

class Application(Frame):
    def __init__(self, master):
        super(Application, self).__init__(master)
        self.initUI()

    def initUI(self):
        self.canvas_head = Canvas(root, bg='AliceBlue', height=30)
        self.canvas_head.place(relx=0, rely=0, relwidth=1)

        self.create_but = Button(self.canvas_head, text='+', bg='CornflowerBlue', font=("Verdana", 8, 'bold'),
                             fg='Black', activebackground='Azure', command=self.create_FB,
                             height=1, width=5)
        self.create_but.place(x=16, y=4)
        self.canvas = Canvas(root, bg='White')
        self.canvas.place(x=0, y=30, relwidth=1, relheight=1)
        # self.canvas = Canvas(self.canvas1, bg='White')
        # self.canvas.place(x=0, y=180, relwidth=1, relheight=1)
        # self.frame1 = Frame(self.canvas1, bg='White')
        # self.frame1.pack()
        self.notebook = CustomNotebook(self.canvas)
        # self.notebook = ttk.Notebook(self.frame1, width=550, height=151)
        a_tab = FB(self.notebook)
        # b_tab = FB(self.notebook, command=self.create_FB)
        self.notebook.add(a_tab, text="ComDoc 1")
        # self.notebook.pack()
        # self.notebook.add(b_tab, text="+")
        # self.notebook.pack()
        # self.frame1.place(relx=0.0, rely=0.0)
        # c = self.frame1.winfo_height()
        # print(c)
        # self.notebooklow.add(l_tab, text="Notebook A")
        self.notebook.place(relx=0, rely=0, relwidth=1, relheight=1)
        self.place(relx=0, rely=0)

    c = 1

    def create_FB(self):
        self.c += 1
        c_tab = FB(self.notebook)
        self.notebook.add(c_tab, text="ComDocs " + str(self.c))
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