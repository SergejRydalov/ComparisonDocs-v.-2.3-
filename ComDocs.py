import numpy as np
import pandas as pd
from tkinter import *
from tkinter import filedialog as fd
import tkinter as tk
import tkinter.ttk as ttk
# import xlsxwriter
# import xlrd

#Classes


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


class Application(Frame):
    def __init__(self, master):
        super(Application, self).__init__(master)
        self.initUI()

    def initUI(self):
        self.canvas = Canvas(root, height=550, width=550)
        self.canvas.pack()

        self.frame = Frame(root, bg='white')
        self.frame.place(relx=0, rely=0.35, relwidth=1, relheight=1)

        self.mylabel1 = Label(self.canvas, text='', font=("Verdana", 9))
        self.mylabel1.place(relx=0.35, rely=0.03)

        self.mylabel2 = Label(self.canvas, text='', font=("Verdana", 9))
        self.mylabel2.place(relx=0.35, rely=0.10)

        self.checklabel = Label(self.canvas, text='', font=("Verdana", 13))
        self.checklabel.place(relx=0.03, rely=0.18)

        self.btnSelPIR = Button(self.canvas, text='Выбрать отчет ПИР', bg='#146780', font=("Verdana", 10), fg='#ffffff',
                           highlightcolor='#4a88c7', activebackground='#539ab0', command=self.show_select_pir)

        self.btnSelPIR.place(relx=0.03, rely=0.03, relwidth=0.31, relheight=0.05)

        self.btnSelASUS = Button(self.canvas, text='Выбрать отчет АСУС', bg='#146780', font=("Verdana", 10),
                                 fg='#ffffff', activebackground='#539ab0', command=self.show_select_asus)

        self.btnSelASUS.place(relx=0.03, rely=0.10, relwidth=0.31, relheight=0.05)

        self.btnSRV = Button(self.canvas, text='Сравнить отчеты', bg='#b7c3c7', font=("Verdana", 10), fg='#ffffff',
                        state=DISABLED, activebackground='#fac96e', command=self.compare)

        self.btnSRV.place(relx=0.03, rely=0.27, relwidth=0.31, relheight=0.05)

        self.btnPRS = Button(self.canvas, text='Просмотреть результат', bg='#b7c3c7', font=("Verdana", 10),
                             fg='#ffffff', state=DISABLED, activebackground='#539ab0', command=self.show_resultat)

        self.btnPRS.place(relx=0.35, rely=0.27, relwidth=0.31, relheight=0.05)

        self.btnSHR = Button(self.canvas, text='Сохранить', bg='#b7c3c7', font=("Verdana", 10), fg='#ffffff',
                        state=DISABLED, activebackground='#24bf1f', command=self.show_save)

        self.btnSHR.place(relx=0.67, rely=0.27, relwidth=0.31, relheight=0.05)

    def check_labels(self):
        if self.mylabel2['text'] != '' and self.mylabel1['text'] != '':
            self.btnSRV.config(state=NORMAL, bg='#e8a11c')
        else:
            self.btnSRV.config(state=DISABLED, bg='#b7c3c7')

    def checktextlabel(self):
        if self.checklabel['text'] == 'Отчеты не совпадают':
            self.btnPRS.config(state=NORMAL, bg='#146780')
            self.btnSHR.config(state=NORMAL, bg='#1b9e16')
        else:
            self.btnPRS.config(state=DISABLED, bg='#b7c3c7')
            self.btnSHR.config(state=DISABLED, bg='#b7c3c7')

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
        column_names = {'Отчет по дебиторской задолженности': 'Категория', 'Unnamed: 2': 'Договор',
                        'Unnamed: 3': 'Name_con', 'Unnamed: 4': 'Ob_Z',
                        'Unnamed: 5': 'Pr_Z', 'Unnamed: 6': 'KR_Z',
                        'Unnamed: 7': 'Age_start', 'Unnamed: 8': 'Age_end', 'Unnamed: 9': 'quantity_Age'}
        df = df.rename(columns=column_names)
        df = df.dropna(thresh=2)[['Категория', 'Договор', 'Name_con', 'Pr_Z']]
        df.drop([2], inplace=True)
        df["Pr_Z"] = pd.to_numeric(df["Pr_Z"])
        df = df.dropna(thresh=2)[['Договор', 'Name_con', 'Pr_Z']]
        df.index = np.arange(1, len(df) + 1)

        df2 = pd.read_excel(self.mylabel2['text'])
        column_names1 = {df2.columns[0]: 'Договор', df2.columns[1]: 'Name_con',
                         df2.columns[5]: 'Ob_Z', df2.columns[6]: 'Pr_Z'}
        df2 = df2.rename(columns=column_names1)
        df2 = df2.dropna(thresh=1)[['Договор', 'Name_con', 'Pr_Z']]
        df2 = df2.dropna()
        df2 = df2.astype({"Name_con": str})
        df2 = df2[~df2.Name_con.str.contains("Контрагент")]
        df2.index = np.arange(1, len(df2) + 1)

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
            print(self.NWDF)
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
            print(dfrez1)
            self.dfrez = dfrez1[['Договор',
                                 'Наименование контрагента',
                                # 'Наличие контрагента в ИС ПИР',
                                 'Размер ДЗ в ИС АСУС',
                                 'Размер ДЗ в ИС ПИР']]
            self.checktextlabel()

    def show_resultat(self):
        dfrezult=self.dfrez
        win = Toplevel(root, bd=5, bg="lightblue")
        win.title("Результат равнения отчетов")
        win.minsize(width=1200, height=800)
        data = dfrezult.to_numpy()
        print(data)
        table = Table(win, headings=dfrezult.columns.tolist(), rows=data)
        print(dfrezult.columns.tolist())
        table.pack(expand=tk.YES, fill=tk.BOTH)

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


root = Tk()

root['bg'] = '#fafafa'
root.title('Сравнение отчетности')
# root.wm_attributes('-alpha', 0.7)
root.geometry('550x500')
root.resizable(width=False, height=False)

app = Application(root)
app.mainloop()