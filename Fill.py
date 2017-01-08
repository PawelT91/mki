__author__ = 'Тупиков Павел'
__version__ = 1.0
from tkinter.ttk import *
from tkinter import *
import os
import xlrd
import xlwt
import glob
from tkinter.filedialog import askopenfilename
from tkinter.simpledialog import askfloat
import datetime
import re

def full_god():
        listboxDOG.delete(0, END)
        LIST_DOG_OBJ = []
        #os.chdir(r'\\192.168.1.78\database\05-Отчетные\04-03_ГМК МЛ')
        os.chdir(r'\\192.168.1.78\database\05-Отчетные\04-03_ГМК МЛ')
        file = sorted(glob.glob('*'))
        number_dog = ''
        for f in file:
            dog = re.findall(r'^(.{3}\-.{2}\-.{1}).*.xlsm$', f)
            if len(dog) != 0:
                if number_dog != str(dog[0]):
                    number_dog = str(dog[0])
                    listboxDOG.insert(END, number_dog)
                    listboxDOG.itemconfig(listboxDOG.size() - 1 , foreground='blue')

def sel_god(event):
        #window = Toplevel(root)
        w = event.widget
        index = int(w.curselection()[0])
        value = w.get(index)
        print(value)
        f = glob.glob(value + '*')[0]
        nombers = []
        for nom in glob.glob(value + '*'):
            nombers.append(nom.split('-')[3])
        fil = xlrd.open_workbook(f)
        sheet = fil.sheet_by_index(0)
        z = 0
        rez = []
        while True: 
            program = sheet.row_values(20+z)[2]
            date = sheet.row_values(20+z)[43]
            ingerer = sheet.row_values(20+z)[48]     
            if date != '':
                    year, month, day = xlrd.xldate_as_tuple(date,0)[:3]
                    date = datetime.date(year, month, day)
            if program:
                rez.append((program,date,ingerer)) 
                z+=1
                isp = Label(root, text = program)
                
            else:
                break
        global group
        group.destroy()
        group = LabelFrame(root, text="ТАБЛИЧКА", padx=5, pady=5)
        group.place(x=100, y=50)
        b = []
        for t in glob.glob(value + '*'):
            fil = xlrd.open_workbook(t)
            sheet = fil.sheet_by_index(0)
            i = 0
            a = []
            while i< len(rez):
                date = sheet.row_values(20+i)[43]
                if date != '':
                        year, month, day = xlrd.xldate_as_tuple(date,0)[:3]
                        date = datetime.date(year, month, day)
                a.append(date)
                i+=1
            b.append(a)
        rows = []    
        for j in range(len(nombers)):
            label = Label(group,text = nombers[j])
            label.grid(row = j+1, column=0)
        for i in range(len(rez)):
            cols = []
            label = Label(group,text = rez[i][0], wraplength = 150)
            label.grid(row = 0, column=i+1)
            for j in range(len(nombers)):
                ent = Entry(group,relief=RIDGE)
                ent.grid(row=j+1, column=i+1)
                ent.insert(END, b[j][i])
                cols.append(ent)
            rows.append(cols)
    
if  __name__ ==  "__main__" :
    root = Tk()
    root.geometry('1200x800')
    root.title('МКИ')
    listboxDOG = Listbox(root, height=17, width=10, selectmode=EXTENDED)
    listboxDOG.place(x=10, y=30) 
    listboxDOG.bind('<<ListboxSelect>>',sel_god)
    global group
    group = LabelFrame(root, text="ТАБЛИЧКА", padx=5, pady=5)
    group.place(x=100, y=50)
    full_god()
    but = Button(root,text="Это кнопка", width=9,height=1, bg="white",fg="blue",command = press)
    but.place(x=10, y=300)

    # Классы
    class fileXL:
        def __init__(self,name_file):
            self.name_file = name_file
            self.nomber_dog = re.findall('.{3}\-.{2}\-.{1}',name_file)[0]
            self.nomber_position = re.findall('.{4}\.\d+', name_file)[0]

        def copy(self, direct, char = ''):
            shutil.copyfile(self.name_file, os.path.join(direct,self.name_file[:-5] + char + self.name_file[-5:]))

        def move(self, direct, char = ''):
            shutil.move(self.name_file, os.path.join(direct,self.name_file[:-5] + char + self.name_file[-5:]))

    class ML(fileXL):
        def __init__(self,name_file):
            fileXL.__init__(self,name_file)
            fil = xlrd.open_workbook(self.name_file)
            sheet = fil.sheet_by_index(0)
            try:
                self.device = sheet.row_values(6)[41]
                self.firm = sheet.row_values(8)[41]
                self.device_name = sheet.row_values(7)[41]
                self.ingerer = sheet.row_values(16)[48]
                self.quantity = int(sheet.row_values(19)[31])
                var = sheet.row_values(16)[43]
                if var != '':
                    year, month, day = xlrd.xldate_as_tuple(var,0)[:3]
                    self.data_vk_end = datetime.date(year, month, day)
                else:
                    self.data_vk_end = ''
            except:
                from tkinter import messagebox
                messagebox.showinfo('Ошибка', 'С файлом ' + name_file +' что то не так.....')
                self.device = 'ОШИБКА!!!'
                self.firm = ''
                self.device_name = 'ОШИБКА!!!'
                self.ingerer = ''
                self.quantity = ''
                var = ''

    def press():
        global group

        
    

    

