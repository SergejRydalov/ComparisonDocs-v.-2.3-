import tkinter as tk
import tkinter.ttk as ttk

import numpy as np
import pandas as pd

from tkinter import *
from tkinter import filedialog as fd

class Table(tk.Frame):
    def __init__(self, parent=None, headings=tuple(), rows=tuple()):
        super().__init__(parent)

        table = ttk.Treeview(self, show="headings", selectmode="browse")
        table["columns"] = headings
        table["displaycolumns"] = headings

        for head in headings:
            table.heading(head, text=head, anchor=tk.CENTER)
            if head == headings[0]:
                table.column(head, width=78, stretch=False, anchor=tk.CENTER)
            elif head == headings[2] or head == headings[3]:
                table.column(head, width=155, stretch=False, anchor=tk.CENTER)
            else:
                table.column(head, anchor=tk.CENTER)

        for row in rows:
            table.insert('', tk.END, values=tuple(row))

        scrolltable = tk.Scrollbar(self, command=table.yview)
        table.configure(yscrollcommand=scrolltable.set)
        scrolltable.pack(side=tk.RIGHT, fill=tk.Y)
        table.pack(expand=tk.YES, fill=tk.BOTH)

class FaceButton(tk.Canvas):
    def __init__(self, parent):
        super().__init__(parent)

        self.parent = parent
        self.init_ui()

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

        self.mylabel1 = Label(self.canvas, bg='Azure', text='', font=("Verdana", 9))
        self.mylabel1.place(x=195, y=15)

        self.mylabel2 = Label(self.canvas, bg='Azure', text='', font=("Verdana", 9))
        self.mylabel2.place(x=195, y=43)

        self.checklabel = Label(self.canvas, bg='Azure', text='', font=("Verdana", 13))
        self.checklabel.place(x=15, y=80)

        self.btnSelPIR = Button(self.canvas, text='Выбрать отчет ПИР', bg='#146780', font=("Verdana", 9), fg='#ffffff',
                                highlightcolor='#4a88c7', activebackground='#539ab0', command=self.show_select_pir,
                                height=1, width=18)
        self.btnSelPIR.place(x=15, y=12)
        self.btnSelASUS = Button(self.canvas, text='Выбрать отчет АСУС', bg='#146780', font=("Verdana", 9),
                                 fg='#ffffff', activebackground='#539ab0', command=self.show_select_asus,
                                 height=1, width=18)
        self.btnSelASUS.place(x=15, y=42)

        self.btnSRV = Button(self.canvas, text='Сравнить отчеты', bg='#b7c3c7', font=("Verdana", 9), fg='#ffffff',
                             state=DISABLED, activebackground='#fac96e', command=lambda:[self.start(), self.compare()],
                             height=1, width=18)

        self.btnSRV.place(x=15, y=117)

        self.btnPRS = Button(self.canvas, text='Просмотреть результат', bg='#b7c3c7', font=("Verdana", 9),
                             fg='#ffffff', state=DISABLED, activebackground='#539ab0', command=self.show_resultat,
                             height=1, width=22)

        self.btnPRS.place(x=200, y=117)

        self.btnSHR = Button(self.canvas, text='Сохранить', bg='#b7c3c7', font=("Verdana", 9), fg='#ffffff',
                             state=DISABLED, activebackground='#24bf1f', command=self.show_save,
                             height=1, width=22)
        self.btnSHR.place(x=500, y=117)
        self.btnSHR.place(x=500, y=117)
        self.var1 = BooleanVar()
        self.check_win = Checkbutton(self.canvas, text="В окне",
                         font=("Verdana", 9),
                         bg='Azure',
                         state=DISABLED,
                         variable=self.var1,
                         onvalue=1, offvalue=0,
                         command=self.check_show_win
                         )
        self.check_win.place(x=390, y=117)

        self.canvas1 = Canvas(self, bg='White')
        self.canvas1.place(x=0, y=158, relwidth=0.978, relheight=0.815)
        self.canvas2 = Canvas(self, bg='Honeydew', height=150)
        self.canvas2.place(x=705, y=0, relwidth=0.632)
        self.canvas3 = Canvas(self, bg='AliceBlue', width=40)
        self.canvas3.place(relx=0.977, y=158, relheight=0.815)
        self.checkbut = 0

        # self.canvas1 = Canvas(self, bg='Azure', height=150, width=550)
        # self.canvas1.place(relx=0.0, rely=0.0)

        self.pack()

    def start(self):
        self.bytes = 0
        self.maxbytes = 0
        self.progress["value"] = 0
        self.maxbytes = 2500
        self.progress["maximum"] = 2500
        self.read_bytes()

    def read_bytes(self):
        '''simulate reading 500 bytes; update progress bar'''
        self.bytes += 100
        self.progress["value"] = self.bytes
        if self.bytes < self.maxbytes:
            # read more bytes after 100 ms
            self.after(10, self.read_bytes)
        # else:
        #     self.progress["value"] = 0

    def check_show_win(self):
        self.checkbut = self.var1.get()

    def check_labels(self):
        if self.mylabel2['text'] != '' and self.mylabel1['text'] != '':
            self.btnSRV.config(state=NORMAL, bg='#e8a11c')
        else:
            self.btnSRV.config(state=DISABLED, bg='#b7c3c7')

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
                ("Файлы Python", "*.xlsx"),
                ("Текстовые файлы", "*.xls"),
                ("Все файлы", "*.*"),
            ],
        )
        print(self.new_path1)
        self.mylabel1.config(text=self.new_path1)
        self.check_labels()

    def show_select_asus(self):
        self.new_path2 = fd.askopenfilename(
            title="Выбор отчета АСУС",
            defaultextension="*.xls",
            filetypes=[
                ("Текстовые файлы", "*.xls"),
                ("Файлы Python", "*.xlsx"),
                ("Все файлы", "*.*"),
            ],
        )
        print(self.new_path2)
        self.mylabel2.config(text=self.new_path2)
        self.check_labels()

    def compare(self):
        df = pd.read_excel(self.new_path1, sheet_name='Лист1')
        self.progress.update_idletasks()
        column_names = {'Отчет по дебиторской задолженности': 'Категория', 'Unnamed: 2': 'Договор',
                        'Unnamed: 3': 'Name_con', 'Unnamed: 4': 'Ob_Z',
                        'Unnamed: 5': 'Pr_Z', 'Unnamed: 6': 'KR_Z',
                        'Unnamed: 7': 'Age_start', 'Unnamed: 8': 'Age_end', 'Unnamed: 9': 'quantity_Age'}
        df = df.rename(columns=column_names)
        self.progress.update_idletasks()
        df = df.dropna(thresh=2)[['Категория', 'Договор', 'Name_con', 'Pr_Z']]
        df.drop([2], inplace=True)
        df["Pr_Z"] = pd.to_numeric(df["Pr_Z"])
        df = df.dropna(thresh=2)[['Договор', 'Name_con', 'Pr_Z']]
        df.index = np.arange(1, len(df) + 1)
        self.progress.update_idletasks()
        df2 = pd.read_excel(self.mylabel2['text'])
        self.progress.update_idletasks()
        column_names1 = {df2.columns[0]: 'Договор', df2.columns[1]: 'Name_con',
                         df2.columns[5]: 'Ob_Z', df2.columns[6]: 'Pr_Z'}
        df2 = df2.rename(columns=column_names1)
        self.progress.update_idletasks()
        df2 = df2.dropna(thresh=1)[['Договор', 'Name_con', 'Pr_Z']]
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
            # print(self.NWDF)
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
            # print(dfrez1)
            self.dfrez = dfrez1[['Договор',
                                 'Наименование контрагента',
                                # 'Наличие контрагента в ИС ПИР',
                                 'Размер ДЗ в ИС АСУС',
                                 'Размер ДЗ в ИС ПИР']]
            self.checktextlabel()

    there_is_table = 0

    def show_resultat(self):
        # self.canvas1.destroy()
        # self.canvas1

        dfrezult = self.dfrez
        if self.checkbut is True:
            win = Toplevel(self, bd=5, bg="lightblue")
            win.title("Результат равнения отчетов")
            win.minsize(width=1200, height=800)
            data = dfrezult.to_numpy()
            self.table = Table(win, headings=dfrezult.columns.tolist(), rows=data)
            self.table.pack(expand=tk.YES, fill=tk.BOTH)
        else:
            if self.there_is_table == 0:
                self.there_is_table += 1
                data = dfrezult.to_numpy()
                self.table = Table(self.canvas1, headings=dfrezult.columns.tolist(), rows=data)
        # print(dfrezult.columns.tolist())
                self.table.pack(expand=tk.YES, fill=tk.BOTH)

    def show_save(self):
        filename = fd.asksaveasfilename(
            title="Сохранить результат как..."
            )
        if filename != "":
            path_filename = filename
            writer = pd.ExcelWriter(path_filename + '.xlsx')
            dfrezult=self.dfrez
            dfrezult.to_excel(writer, 'Sheet1')
            worksheet = writer.sheets['Sheet1']
            worksheet.set_column('A:A', 2)
            worksheet.set_column('B:B', 8)
            worksheet.set_column('C:C', 50)
            worksheet.set_column('D:D', 15)
            worksheet.set_column('E:E', 50)
            worksheet.set_column('F:F', 15)
            workbook = writer.book
            writer.save()
