from tkinter import Button, PhotoImage, Label, StringVar, Entry, Tk, Frame, ttk, LEFT
from subprocess import Popen
from extractosMes import run as extractMonth
from extractosMes import runDay as extractDay
from conciliar import run as conciliar
from auxiliar import run as auxiliar
import pandas as pd
import numpy as np

class App:

    def __init__(self, geometry, title):
        self.root = Tk()
        self.code = 1
        self.month = StringVar()
        self.year = StringVar()
        self.codigo = StringVar()
        self.cuenta_contable = StringVar()
        self.nombre = StringVar()
        self.banco = StringVar()
        self.cuenta_bancaria = StringVar()
        self.color_team = ('#E11419')
        self.root.geometry(geometry)
        self.root.resizable(width=0, height=0)
        self.root.title(title)
        self.start = True
        self.bank_list = ['bancolombia', 'bogota', 'davivienda', 'colpatria', 'caja', 'agrario']
        self.frames()
        self.menu()
        if self.start: self.create_home_page()
    
    def openFile(self, doc):
        p = Popen(doc)
        stdout, stderr = p.communicate()
    
    def charge_file(self,file, config):
        self.file = config(file)
        self.file = np.asarray(self.file)
        return self.file
    
    def assign_banks_teams(self, data):
        self.codigo.set(data[int(self.code)-1][0])
        self.data_codigo = self.codigo.get()
        self.data_codigo= self.data_codigo.replace(".0","")
        if int(self.data_codigo) <10: self.codigo.set(f'0{str(self.data_codigo)}')
        self.cuenta_contable.set(data[int(self.code)-1][1])
        self.nombre.set(data[int(self.code)-1][2])
        self.banco.set(data[int(self.code)-1][3])
        self.cuenta_bancaria.set(data[int(self.code)-1][4])
    
    def new_banks_teams(self):
        self.code = len(self.banks_teams)+1
        self.file = pd.read_csv('banks.csv')
        self.file.loc[self.code-1, 'codigo'] = self.code
        self.file.loc[self.code-1, 'cuentaC'] = " "
        self.file.loc[self.code-1, 'nombre'] = " "
        self.file.loc[self.code-1, 'banco'] = " "
        self.file.loc[self.code-1, 'cuentaB'] = " "
        self.file.to_csv('banks.csv', index=False)
        self.indicate_page(self.create_bank_page)
    
    def save_banks_teams(self):
        self.file = pd.read_csv('banks.csv')
        self.file.loc[self.code-1, 'codigo'] = self.codigo.get()
        self.file.loc[self.code-1, 'cuentaC'] = self.cuenta_contable.get()
        self.file.loc[self.code-1, 'nombre'] = self.nombre.get()
        self.file.loc[self.code-1, 'banco'] = self.banco.get()
        self.file.loc[self.code-1, 'cuentaB'] = self.cuenta_bancaria.get()
        self.file.to_csv('banks.csv', index=False)
        


    def consolidate(self, function):
        date_month= self.month.get()
        date_year= self.year.get()
        function(date_year,date_month)

    
    def frames(self,):
        self.option_frame = self.create_frame(self.root, 35, 500, color = '#c3c3c3')
        self.main_frame = self.create_frame(self.root, 565, 500, highlightbackground='black', highlightthickness= 2)
    
    def create_frame(self, master, height, width, color = 'white', highlightbackground=None, highlightthickness=None):
        self.frame = Frame(master, bg= color,highlightbackground= highlightbackground, highlightthickness=highlightthickness)
        self.frame.pack(fill='x')
        self.frame.pack_propagate(False)
        self.frame.configure(width = width, height = height)
        return self.frame
    
    def create_button(self,master, text, x, y, width, height, color, command= None, font = ('Bold', 15), color_text='black'):
        self.button = Button(master, text=text, command= command, bg= color, font=font, fg=color_text)
        self.button.place(relx=x, rely=y, width= width, height= height)
    
    def create_text(self, master, text, bg='white', fg= 'black', font=("Verdana", 16), x=0.01, y=0.01):
        self.text = Label(master, text=text, bg= bg, fg=fg, font=font)
        self.text.pack()
        self.text.place(relx=x,rely=y)
    
    def create_entry(self, master, title, str, x=0.01, y=0.01, font=('Verdana', 12), bg='white', fg='black', widht=50, height=50, move_x=0.1, state='normal', textvariable= False):
        self.entry = Entry(master, textvariable= str)
        self.entry.config(state=state)
        self.entry.place(relx=x+move_x, rely=y, width=widht, height=height)
        self.title = Label(master, text=title)
        self.title.config(font= font, bg=bg, fg=fg)
        self.title.place(relx=x, rely=y)
    
    def create_combobox(self, master,values,str, x=0.01, y=0.01, width=10, height=10):
        self.combobox = ttk.Combobox(master, values= values, width=width, height=height, textvariable=str)
        self.combobox.place(relx=x, rely=y, width=width, height=height)
    
    def move_banks_teams(self, operation):
        if operation == '+': self.code += 1
        if operation == '-': self.code -= 1
        if self.code == 0: self.code = len(self.banks_teams)
        if self.code == len(self.banks_teams)+1 : self.code = 1
        self.indicate_page(self.create_bank_page)

        

    def menu(self):
        self.create_button(self.option_frame, 'Conciliar', 0, 0, 250, 35, self.color_team, command=lambda: self.indicate_page(self.create_home_page), color_text='white')
        self.create_button(self.option_frame, 'Bancos', 0.5, 0, 250, 35, self.color_team, command=lambda: self.indicate_page(self.create_bank_page), color_text='white')
    
    def delete_page(self):
        for frame in self.main_frame.winfo_children():
            frame.destroy()
    
    def indicate_page(self, page):
        self.delete_page()
        page()
    
    def create_home_page(self):
        self.home_page = self.create_frame(self.main_frame, 565, 500)
        self.create_entry(self.home_page, 'Año', self.year, x=0.7, y=0.08, widht=50, height=20, fg=self.color_team)
        self.create_entry(self.home_page, 'Mes', self.month, x=0.7, y=0.15, widht=50, height=20, fg=self.color_team)
        self.create_button(self.home_page, 'LIBROS', 0.7, 0.30, 100, 35, self.color_team, font=('Bold', 10), color_text='white', command= lambda: auxiliar())
        self.create_button(self.home_page, 'CONCILIAR', 0.7, 0.40, 100, 35, self.color_team, font=('Bold', 10), color_text='white', command= lambda: conciliar())
        self.create_button(self.home_page, 'FILTROS', 0.7, 0.50, 100, 35, self.color_team, font=('Bold', 10), color_text='white', command= lambda: self.openFile("openFilters.bat"))

        self.create_text(self.home_page, 'Extractos Mensuales', fg=self.color_team, x=0.05, y=0.08)
        self.create_button(self.home_page, 'ABRIR CARPETA', 0.05, 0.15, 200, 35, self.color_team, font=('Bold', 10), color_text='white', command= lambda: self.openFile("openMonth.bat"))
        self.create_button(self.home_page, 'GENERAR CONSOLIDADO', 0.05, 0.25, 200, 35, self.color_team, font=('Bold', 10), color_text='white', command= lambda: self.consolidate(extractMonth))
        self.create_text(self.home_page, 'Extractos Diarios', fg=self.color_team, x=0.05, y=0.50)
        self.create_button(self.home_page, 'ABRIR CARPETA', 0.05, 0.57, 200, 35, self.color_team, font=('Bold', 10), color_text='white', command= lambda: self.openFile("openDay.bat"))
        self.create_button(self.home_page, 'GENERAR CONSOLIDADO', 0.05, 0.67, 200, 35, self.color_team, font=('Bold', 10), color_text='white', command= lambda: self.consolidate(extractDay))
        

        self.imagen = PhotoImage(file ='logo.png')
        self.lb_imagen =Label(self.home_page, image= self.imagen, bd=0, fg="white").place(relx=0.70,rely=0.7)
        self.create_text(self.home_page, 'Desarrollado por Sebastian Moncada Cel:324-221-0852', font=('Bold', 10), x=0.05, y=0.95)
        self.home_page.pack(fill='both')
    
    def create_bank_page(self):
        self.banks_teams = self.charge_file('banks.csv', pd.read_csv)
        self.assign_banks_teams(self.banks_teams)
        self.bank_page = self.create_frame(self.main_frame, 565, 500)
        self.create_entry(self.bank_page, 'Codigo', self.codigo, font=('Verdana',16), y=0.05, fg=self.color_team, widht=200, height=30, move_x= 0.4, state='disable', textvariable=False)
        self.create_entry(self.bank_page, 'Cuenta Contable', self.cuenta_contable, y=0.15, font=('Verdana',16), fg=self.color_team, widht=200, height=30, move_x= 0.4)
        self.create_entry(self.bank_page, 'Nombre', self.nombre, font=('Verdana',16), y=0.25, fg=self.color_team, widht=200, height=30, move_x= 0.4)
        self.create_text(self.bank_page, 'Banco', font=('Verdana',16), y=0.35, fg=self.color_team,)
        self.create_combobox(self.bank_page, self.bank_list, self.banco, x=0.41, y=0.35, width=200, height=30)
        self.create_entry(self.bank_page, 'Cuenta bancaria', self.cuenta_bancaria, y=0.45, font=('Verdana',16), fg=self.color_team, widht=200, height=30, move_x= 0.4)

        self.create_button(self.bank_page, '<', 0.01, 0.6, 40, 20, self.color_team, font=('bold', 14), color_text='white', command= lambda: self.move_banks_teams('-'))
        self.create_button(self.bank_page, '>', 0.11, 0.6, 40, 20, self.color_team, font=('bold', 14), color_text='white', command= lambda: self.move_banks_teams('+'))
        self.create_button(self.bank_page, 'guardar', 0.31, 0.6, 50, 20, self.color_team, font=('bold', 10), color_text='white', command= lambda: self.save_banks_teams())
        self.create_button(self.bank_page, 'nuevo', 0.45, 0.6, 45, 20, self.color_team, font=('bold', 10), color_text='white', command= lambda: self.new_banks_teams())

        self.imagen = PhotoImage(file ='logo.png')
        self.lb_imagen =Label(self.bank_page, image= self.imagen, bd=0, fg="white").place(relx=0.70,rely=0.7)
        self.create_text(self.bank_page, 'Desarrollado por Sebastian Moncada Cel:324-221-0852', font=('Bold', 10), x=0.05, y=0.95)
        self.bank_page.pack(fill='both')





root = App('500x600', 'Conciliacion Bancaria')
root.root.mainloop()
