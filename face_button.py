import tkinter as tk
import tkinter.ttk as ttk

import numpy as np
import pandas as pd
import clipboard
import os

from tkinter import *
from tkinter import filedialog as fd
from tkinter.font import BOLD

import datetime


class Table(tk.Frame):
    def __init__(self, parent=None, headings=tuple(), rows=tuple()):
        super().__init__(parent)
        self.table = ttk.Treeview(self, show="headings", selectmode="browse")
        self.table["columns"] = headings
        self.table["displaycolumns"] = headings

        for head in headings:
            self.table.heading(head, text=head, anchor=tk.CENTER)
            if ('Договор' in headings) is True:
                if head == headings[headings.index('Договор')]:
                    self.table.column(head, width=78, stretch=False, anchor=tk.CENTER)
                else:
                    if ('Размер ДЗ в ИС АСУС' in headings) is True:
                        if head == headings[headings.index('Размер ДЗ в ИС АСУС')]:
                            self.table.column(head, width=155, stretch=False, anchor=tk.E)
                        else:
                            if ('Размер ДЗ в ИС ПИР' in headings) is True:
                                if head == headings[headings.index('Размер ДЗ в ИС ПИР')]:
                                    self.table.column(head, width=155, stretch=False, anchor=tk.E)
                                else:
                                    self.table.column(head, anchor=tk.W)
                            else:
                                self.table.column(head, anchor=tk.W)
                    else:
                        if ('Размер ДЗ в ИС ПИР' in headings) is True:
                            if head == headings[headings.index('Размер ДЗ в ИС ПИР')]:
                                self.table.column(head, width=155, stretch=False, anchor=tk.E)
                            else:
                                self.table.column(head, anchor=tk.W)
                        else:
                            self.table.column(head, anchor=tk.W)
            else:

                if ('Размер ДЗ в ИС АСУС' in headings) is True:
                    if head == headings[headings.index('Размер ДЗ в ИС АСУС')]:
                        self.table.column(head, width=155, stretch=False, anchor=tk.E)
                    else:
                        if ('Размер ДЗ в ИС ПИР' in headings) is True:
                            if head == headings[headings.index('Размер ДЗ в ИС ПИР')]:
                                self.table.column(head, width=155, stretch=False, anchor=tk.E)
                            else:
                                self.table.column(head, anchor=tk.W)
                        else:
                            self.table.column(head, anchor=tk.W)
                else:
                    if ('Размер ДЗ в ИС ПИР' in headings) is True:
                        if head == headings[headings.index('Размер ДЗ в ИС ПИР')]:
                            self.table.column(head, width=155, stretch=False, anchor=tk.E)
                        else:
                            self.table.column(head, anchor=tk.W)
                    else:
                        self.table.column(head, anchor=tk.W)

        for row in rows:
            self.table.insert('', tk.END, values=tuple(row), tags='a')
        scrolltable = tk.Scrollbar(self, command=self.table.yview)
        self.table.configure(yscrollcommand=scrolltable.set)
        scrolltable.pack(side=tk.RIGHT, fill=tk.Y)
        self.table.pack(expand=tk.YES, fill=tk.BOTH)

class FaceButton(tk.Canvas):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.init_ui()
        self.col = {}
        self.col2 = {}
        self.activCombo1 = "disabled"
        self.activCombo2 = "disabled"
        self.initComboPir()
        self.initComboAsus()
        self.checkBut = 0
        self.there_is_table = 0
        self.activSelPir = True
        self.activSelAsus = True

    def init_ui(self):
        self.pack(fill=tk.BOTH, expand=1)
        s = ttk.Style()
        s.theme_use('clam')
        s.configure("red.Horizontal.TProgressbar", foreground='red', background='Orange')
        self.progress = ttk.Progressbar(self, style="red.Horizontal.TProgressbar", orient="horizontal",
                                        length=200, mode="determinate")
        self.progress.place(x=0, y=151, relwidth=1)
        self.canvas = Canvas(self, bg='Azure', height=150, width=700)
        self.canvas.place(x=0, y=0)
        self.mylabel1 = Label(self.canvas, bg='Azure', text='', font=("Verdana", 7))
        self.mylabel1.place(x=195, y=15)
        self.mylabel2 = Label(self.canvas, bg='Azure', text='', font=("Verdana", 7))
        self.mylabel2.place(x=195, y=43)
        self.checklabel = Label(self.canvas, bg='Azure', text='', font=("Verdana", 13))
        self.checklabel.place(x=15, y=80)
        self.btnSelPIR = Button(self.canvas, text='Открыть отчет ПИР', bg='#146780', font=("Verdana", 9), fg='#ffffff',
                                highlightcolor='#4a88c7', activebackground='#539ab0', command=self.show_select_pir)
        self.btnSelPIR.place(x=15, y=12, height=27, width=170)
        self.btnSelASUS = Button(self.canvas, text='Открыть отчет АСУС', bg='#146780', font=("Verdana", 9),
                                 fg='#ffffff', activebackground='#539ab0', command=self.show_select_asus)
        self.btnSelASUS.place(x=15, y=42, height=27, width=170)
        self.btnSRV = Button(self.canvas, text='Сравнить отчеты', bg='#b7c3c7', font=("Verdana", 9), fg='#ffffff',
                             state=DISABLED, activebackground='#fac96e', command=lambda:[self.start(), self.compare()])
        self.btnSRV.place(x=15, y=117, height=27, width=170)
        self.btnPRS = Button(self.canvas, text='Просмотреть результат', bg='#b7c3c7', font=("Verdana", 9),
                             fg='#ffffff', state=DISABLED, activebackground='#539ab0', command=self.show_resultat)
        self.btnPRS.place(x=200, y=117, height=27, width=190)
        self.btnSHR = Button(self.canvas, text='Сохранить', bg='#b7c3c7', font=("Verdana", 9), fg='#ffffff',
                             state=DISABLED, activebackground='#24bf1f',
                             command=self.show_save)
        self.btnSHR.place(x=518, y=117, height=27, width=170)
        self.var1 = BooleanVar()
        self.check_win = Checkbutton(self.canvas, text="В окне",
                         font=("Verdana", 9),
                         bg='Azure',
                         state=DISABLED,
                         variable=self.var1,
                         onvalue=1, offvalue=0,
                         command=self.check_show_win)
        self.check_win.place(x=390, y=117)
        self.canvas1 = Canvas(self, bg='White')
        self.canvas1.place(x=0, y=158, relwidth=1, relheight=0.815)
        self.canvas2 = Canvas(self,
                              bg='Honeydew',
                              height=150)
        self.canvas2.place(x=705, y=0, relwidth=1)
        self.mylabel3 = Label(self.canvas2, bg='Honeydew', text='Выполните выбор столбцов для сравнения документов', font=("Verdana", 12))
        self.mylabel3.place(x=15, y=8)
        self.canvas_columnName = Canvas(self.canvas2, bg='Azure')
        self.canvas_columnName.place(x=8, y=35, relheight=0.2, relwidth=0.99)
        self.canvas_column1 = Canvas(self.canvas2, bg='Azure')
        self.canvas_column1.place(x=8, y=70, relheight=0.2, relwidth=0.99)
        self.canvas_column2 = Canvas(self.canvas2, bg='Azure')
        self.canvas_column2.place(x=8, y=105, relheight=0.2, relwidth=0.99)
        self.mylabelPir = Label(self.canvas_column1, bg='Azure', text='ПИР',
                              font=("Verdana", 11, BOLD))
        self.mylabelPir.place(x=6, y=4)
        self.mylabelAsus = Label(self.canvas_column2, bg='Azure', text='АСУС',
                                font=("Verdana", 11, BOLD))
        self.mylabelAsus.place(x=6, y=4)
        self.mylabelPirContract = Label(self.canvas_columnName, bg='Azure', text='Договор',
                                font=("Verdana", 11, BOLD))
        self.mylabelPirContract.place(x=147, y=2)
        self.mylabelPirName = Label(self.canvas_columnName, bg='Azure', text='Наименование контрагента',
                                        font=("Verdana", 11, BOLD))
        self.mylabelPirName.place(x=497, y=2)
        self.mylabelPrZ = Label(self.canvas_columnName, bg='Azure', text='Просроченная задолженность',
                                    font=("Verdana", 11, BOLD))
        self.mylabelPrZ.place(x=847, y=2)

        self.pack()

    def start(self):
        self.bytes = 0
        self.maxBytes = 0
        self.progress["value"] = 0
        self.maxBytes = 2500
        self.progress["maximum"] = 2500
        self.read_bytes()

    def read_bytes(self):
        self.bytes += 100
        self.progress["value"] = self.bytes
        if self.bytes < self.maxBytes:
            self.after(10, self.read_bytes)

    def check_show_win(self):
        self.checkBut = self.var1.get()

    def check_labels(self):
        if self.column1 != '' and self.column2 != '' and self.column3 != '' and self.column4 != '' \
                and self.column5 != '' and self.column6 != '':
            self.btnSRV.config(state=NORMAL, bg='#e8a11c')
        else:
            self.btnSRV.config(state=DISABLED, bg='#b7c3c7')

    def checkTab(self):

        if self.activSelPir != True:
            self.btnSelPIR.config(state=DISABLED, bg='#b7c3c7')
            self.btnDelDocPir.config(state=NORMAL, bg='#ff6e4a')
        else:
            self.btnSelPIR.config(state=NORMAL, bg='#146780')
            self.btnDelDocPir.config(state=DISABLED, bg='#b7c3c7')

        if self.activSelAsus != True:
            self.btnSelASUS.config(state=DISABLED, bg='#b7c3c7')
            self.btnDelDocAsus.config(state=NORMAL, bg='#ff6e4a')
        else:
            self.btnSelASUS.config(state=NORMAL, bg='#146780')
            self.btnDelDocAsus.config(state=DISABLED, bg='#b7c3c7')

    def checktextlabel(self):
        if self.checklabel['text'] == 'Отчеты не совпадают':
            self.btnPRS.config(state=NORMAL, bg='#146780')
            self.btnSHR.config(state=NORMAL, bg='#1b9e16')
            self.check_win.config(state=NORMAL)
            self.var1.set(0)
        else:
            self.btnPRS.config(state=DISABLED, bg='#b7c3c7')
            self.btnSHR.config(state=DISABLED, bg='#b7c3c7')
            self.check_win.config(state=DISABLED)

    def show_select_pir(self):
        self.new_path1 = fd.askopenfilename(
            title="Выбор отчета ПИР",
            defaultextension="*.xlsx",
            filetypes=[
                ("Excel", "*.xlsx"),
                ("Еxcel 93-2003", "*.xls"),
                ("Все файлы", "*.*"),
            ],
        )
        self.fNamePir = os.path.basename(self.new_path1)
        self.mylabel1.config(text=self.new_path1)
        dfDoc = pd.read_excel(self.new_path1)
        dfrezult = dfDoc
        self.data = dfrezult.to_numpy()
        if self.canvas1.winfo_children() == []:
            self.notebookDocs = ttk.Notebook(self.canvas1)

            self.p_tab = Table(self.notebookDocs, headings=dfrezult.columns.tolist(), rows=self.data)
            self.notebookDocs.add(self.p_tab,
                                  text='   ' + str(self.fNamePir) + '                                        ')
            self.notebookDocs.place(relx=0, rely=0, relwidth=1, relheight=1)
        else:
            self.p_tab = Table(self.notebookDocs, headings=dfrezult.columns.tolist(), rows=self.data)
            self.notebookDocs.add(self.p_tab,
                                  text='   ' + str(self.fNamePir) + '                                        ')
            self.notebookDocs.place(relx=0, rely=0, relwidth=1, relheight=1)
            self.notebookDocs.select(self.p_tab)
        self.col = dfDoc.columns.tolist()
        self.col.insert(0, "Выберите столбец таблицы")
        self.activCombo1 = "readonly"
        self.initComboPir()
        self.activSelPir = False
        self.checkTab()

    def show_select_asus(self):
        self.new_path2 = fd.askopenfilename(
            title="Выбор отчета АСУС",
            defaultextension="*.xls",
            filetypes=[
                ("Еxcel 93-2003", "*.xls"),
                ("Еxcel", "*.xlsx"),
                ("Все файлы", "*.*"),
            ],
        )
        self.fNameAsus = os.path.basename(self.new_path2)
        self.mylabel2.config(text=self.new_path2)
        dfDoc = pd.read_excel(self.new_path2)
        dfrezult = dfDoc
        self.data = dfrezult.to_numpy()
        if not self.canvas1.winfo_children():
            self.notebookDocs = ttk.Notebook(self.canvas1)
            self.as_tab = Table(self.notebookDocs, headings=dfrezult.columns.tolist(), rows=self.data)
            self.notebookDocs.add(self.as_tab,
                                  text='   ' + str(self.fNameAsus) + '                                        ')
            self.notebookDocs.place(relx=0, rely=0, relwidth=1, relheight=1)
        else:
            self.as_tab = Table(self.notebookDocs, headings=dfrezult.columns.tolist(), rows=self.data)
            self.notebookDocs.add(self.as_tab,
                                  text='   ' + str(self.fNameAsus) + '                                        ')
            self.notebookDocs.place(relx=0, rely=0, relwidth=1, relheight=1)
            self.notebookDocs.select(self.as_tab)
        self.col2 = dfDoc.columns.tolist()
        self.col2.insert(0, "Выберите столбец таблицы")
        self.activCombo2 = "readonly"
        self.initComboAsus()
        self.activSelAsus = False
        self.checkTab()

    def initComboPir(self):
        self.comboExample1 = ttk.Combobox(self.canvas_column1,
                                         values=self.col, width=50, state=self.activCombo1)
        self.comboExample1.current(0)
        self.comboExample1.place(x=150, y=5)
        self.comboExample1.bind("<<ComboboxSelected>>", self.callbackFunc)
        self.comboExample2 = ttk.Combobox(self.canvas_column1,
                                         values=self.col, width=50, state=self.activCombo1)
        self.comboExample2.current(0)
        self.comboExample2.place(x=500, y=5)
        self.comboExample2.bind("<<ComboboxSelected>>", self.callbackFunc)
        self.comboExample3 = ttk.Combobox(self.canvas_column1,
                                         values=self.col, width=50, state=self.activCombo1)
        self.comboExample3.current(0)
        self.comboExample3.place(x=850, y=5)
        self.comboExample3.bind("<<ComboboxSelected>>", self.callbackFunc)

        self.btnDelDocPir = Button(self.canvas_column1, anchor="w", text=u"\u00D7", bg='#b7c3c7', font=("Arial", 13), fg='#ffffff',
                                 highlightcolor='#4a88c7', activebackground='White', state=DISABLED,
                                 command=self.DelDocPir)
        self.btnDelDocPir.place(x=70, y=5, height=20, width=20)

    def initComboAsus(self):
        self.comboExample4 = ttk.Combobox(self.canvas_column2,
                                          values=self.col2, width=50, state=self.activCombo2)
        self.comboExample4.current(0)
        self.comboExample4.place(x=150, y=5)
        self.comboExample4.bind("<<ComboboxSelected>>", self.callbackFunc)
        self.comboExample5 = ttk.Combobox(self.canvas_column2,
                                          values=self.col2, width=50, state=self.activCombo2)
        self.comboExample5.current(0)
        self.comboExample5.place(x=500, y=5)
        self.comboExample5.bind("<<ComboboxSelected>>", self.callbackFunc)
        self.comboExample6 = ttk.Combobox(self.canvas_column2,
                                          values=self.col2, width=50, state=self.activCombo2)
        self.comboExample6.current(0)
        self.comboExample6.place(x=850, y=5)
        self.comboExample6.bind("<<ComboboxSelected>>", self.callbackFunc)

        self.btnDelDocAsus = Button(self.canvas_column2, anchor="w", text=u"\u00D7", bg='#b7c3c7', font=("Arial", 13),
                                   fg='#ffffff',
                                   highlightcolor='#4a88c7', activebackground='White', state=DISABLED,
                                   command=self.DelDocAsus)
        self.btnDelDocAsus.place(x=70, y=5, height=20, width=20)

    def callbackFunc(self, *args):
        self.column1 = self.comboExample1.get()
        self.column2 = self.comboExample2.get()
        self.column3 = self.comboExample3.get()
        self.column4 = self.comboExample4.get()
        self.column5 = self.comboExample5.get()
        self.column6 = self.comboExample6.get()
        print(self.column1)
        print(self.column2)
        print(self.column3)
        print(self.column4)
        print(self.column5)
        print(self.column6)
        if self.column1 != 'Выберите столбец таблицы' and self.column2 != 'Выберите столбец таблицы' and \
                self.column3 != 'Выберите столбец таблицы' and self.column4 != 'Выберите столбец таблицы' and \
                self.column5 != 'Выберите столбец таблицы' and self.column6 != 'Выберите столбец таблицы':
                self.check_labels()

    def DelDocAsus(self):
        self.notebookDocs.forget(self.as_tab)
        self.activCombo2 = "disabled"
        self.initComboAsus()
        self.activSelAsus = True
        self.col2 = {}
        self.checkTab()
        self.new_path2 = ""
        self.column6 = ''
        self.mylabel2.config(text='')
        self.check_labels()

    def DelDocPir(self):
        self.notebookDocs.forget(self.p_tab)
        self.activCombo1 = "disabled"
        self.initComboPir()
        self.activSelPir = True
        self.col1 = {}
        self.checkTab()
        self.new_path1 = ""
        self.column6 = ''
        self.mylabel1.config(text='')
        self.check_labels()

    def compare(self):
        df = pd.read_excel(self.new_path1
                           #sheet_name='Лист1'
                           )
        self.progress.update_idletasks()
        column_names = {self.column1: 'Договор',
                        self.column2: 'Name_con',
                        #'Unnamed: 4': 'Ob_Z',
                        self.column3: 'Pr_Z',
                        }
        df = df.rename(columns=column_names)

        self.progress.update_idletasks()
        df = df.dropna(thresh=2)[['Договор', 'Name_con', 'Pr_Z']]
        print(df["Pr_Z"].dtype)
        if df["Pr_Z"].dtype == object:
            df = df[~df.Pr_Z.str.contains("Про", na=False)]
        #df.drop([2], inplace=True)
        df["Pr_Z"] = pd.to_numeric(df["Pr_Z"])
        df = df.dropna(thresh=2)[['Договор', 'Name_con', 'Pr_Z']]
        df.index = np.arange(1, len(df) + 1)
        self.progress.update_idletasks()
        df2 = pd.read_excel(self.new_path2)
        self.progress.update_idletasks()
        column_names1 = {self.column4: 'Договор', self.column5: 'Name_con',
                         #df2.columns[5]: 'Ob_Z',
                         self.column6: 'Pr_Z'}
        df2 = df2.rename(columns=column_names1)
        self.progress.update_idletasks()
        df2 = df2.dropna(thresh=1)[['Договор',
                                    'Name_con',
                                    'Pr_Z']]
        df2 = df2.dropna()
        df2 = df2.astype({"Name_con": str})
        df2 = df2[~df2.Name_con.str.contains("Контрагент")]
        df2.index = np.arange(1, len(df2) + 1)
        self.progress.update_idletasks()

        srv = (df.shape == df2.shape)  # первый признак сходимости отчетов по их форме (число строк и столбцов)

        if srv is True:
            message = 'Отчеты совпадают'
            print(message)
            self.checklabel.config(text=message)
        else:
            message = 'Отчеты не совпадают'
            print(message)
            self.checklabel.config(text=message)
            df = df.astype({"Pr_Z": int, "Name_con": str, "Договор": str})
            df2 = df2.astype({"Pr_Z": int, "Name_con": str, "Договор": str})
            self.NWDF = (df.merge(df2, how='outer', on=['Договор'],
                             suffixes=['', '_ASUS'], indicator=True))
            self.NWDF['Pr_Z_ASUS'] = self.NWDF['Pr_Z_ASUS'].fillna(value=0)
            self.NWDF['Name_con_ASUS'] = self.NWDF['Name_con_ASUS'].fillna(value=" ")
            self.NWDF = self.NWDF.query("Pr_Z != Pr_Z_ASUS")
            self.NWDF = self.NWDF.sort_values(by='Договор', ascending=False)
            self.NWDF.index = np.arange(1, len(self.NWDF) + 1)
            del self.NWDF['_merge']
            self.NWDF.groupby(self.NWDF['Договор']).mean()
            self.NWDF['Name_con'].fillna(value=self.NWDF['Name_con_ASUS'], inplace=True)
            del self.NWDF['Name_con_ASUS']
            self.NWDF.fillna('Отсутствует', inplace=True)
            column_names = {'Договор': 'Договор', 'Name_con': 'Наименование контрагента',
                            # 'Name_con_ASUS': 'Наличие контрагента в ИС ПИР',
                            'Pr_Z_ASUS': 'Размер ДЗ в ИС АСУС',
                            'Pr_Z': 'Размер ДЗ в ИС ПИР'}
            dfrez1 = self.NWDF.rename(columns=column_names)
            self.dfrez = dfrez1[['Договор',
                                 'Наименование контрагента',
                                # 'Наличие контрагента в ИС ПИР',
                                 'Размер ДЗ в ИС АСУС',
                                 'Размер ДЗ в ИС ПИР']]
            self.checktextlabel()

    def show_resultat(self):
        dfrezult = self.dfrez
        dateTimeNow = datetime.datetime.now()
        if self.checkBut is True:
            win = Toplevel(self, bd=5, bg="lightblue")
            win.title("Результат сравнения документов " + dateTimeNow.strftime('%Y/%m/%d %H:%M:%S'))
            win.minsize(width=1200, height=800)
            data = dfrezult.to_numpy()
            self.table = Table(win, headings=dfrezult.columns.tolist(), rows=data)
            self.table.pack(expand=tk.YES, fill=tk.BOTH)
        else:
            if self.there_is_table == 0:
                self.there_is_table += 1
                data = dfrezult.to_numpy()
                if not self.canvas1.winfo_children():
                    self.notebookDocs = ttk.Notebook(self.canvas1)
                    r_tab = Table(self.notebookDocs, headings=dfrezult.columns.tolist(), rows=data)
                    self.notebookDocs.add(r_tab, text='Результат сравнения документов ' + dateTimeNow.strftime('%Y/%m/%d %H:%M:%S'))
                    self.notebookDocs.place(relx=0, rely=0, relwidth=1, relheight=1)
                else:
                    r_tab = Table(self.notebookDocs, headings=dfrezult.columns.tolist(), rows=data)
                    self.notebookDocs.add(r_tab, text='Результат сравнения документов ' + dateTimeNow.strftime('%Y/%m/%d %H:%M:%S'))
                    self.notebookDocs.place(relx=0, rely=0, relwidth=1, relheight=1)
                    self.notebookDocs.select(r_tab)

    def show_save(self):
        dateTimeNow = datetime.datetime.now()
        nameDefault = "Результат сравнения документов " + dateTimeNow.strftime('%Y_%m_%d %Hч%Mм%Sс' + ".xlsx")
        filename = fd.asksaveasfilename(initialfile=nameDefault,
            title="Сохранить результат как...",
            filetypes=[("Excel", "*.xlsx")]
            )
        if filename != "":
            path_filename = filename
            writer = pd.ExcelWriter(path_filename + '.xlsx')
            dfrezult = self.dfrez
            dfrezult.to_excel(writer, 'Sheet1')
            worksheet = writer.sheets['Sheet1']
            text_format = writer.book.add_format({'text_wrap': True})
            worksheet.set_column('A:A', 5)
            worksheet.set_column('B:B', 8, text_format)
            worksheet.set_column('C:C', 40, text_format)
            worksheet.set_column('D:D', 20, text_format)
            worksheet.set_column('E:E', 40, text_format)
            worksheet.set_column('F:F', 20, text_format)
            workbook = writer.book
            writer.save()