import tkinter as tk
from tkinter import Button, PhotoImage, Label, StringVar, Entry
from subprocess import Popen
from extractosMes import run as extractMonth
from extractosMes import runDay as extractDay
from conciliar import run as conciliar
from auxiliar import run as auxiliar

colorTeam = ('#E11419')
global month
global year

def conciliarBank():
    conciliar()

def openFilters():
    p = Popen("openFilters.bat")
    stdout, stderr = p.communicate()

def openMonth():
    # try:
    p = Popen("openMonth.bat")
    stdout, stderr = p.communicate()
    # except Exception as e:
    #     print(e)
    #     while(True):
    #         pass

def librosAuxiliar():
    auxiliar()

def consolidateMonth():
    dataMonth = month.get()
    dataYear = year.get()
    extractMonth(dataYear,dataMonth)

def openDay():
    p = Popen("openDay.bat")
    stdout, stderr = p.communicate()

def consolidateDay():
    dataMonth = month.get()
    dataYear = year.get()
    extractDay(dataYear, dataMonth)

def createbutton(function, nombre, x, y ): 
    button = Button(root, text=nombre, command= function, bg= colorTeam, fg='white')
    button.place(relx=x,rely=y+0.10, relwidth=0.50, relheight=0.08)

def createbutton2(function, nombre, x, y ): 
    button = Button(root, text=nombre, command= function, bg= colorTeam, fg='white')
    button.place(relx=x,rely=y+0.10, relwidth=0.20, relheight=0.08)

def escribirDatos(title, str, x, y):
    data = Entry(root, textvariable=str)
    data2= Label(root, text=title)
    data2.place(rely=y+0.05, relx= x-0.2)
    data2.config(font= ("Verdana", 12),bg='white', fg= colorTeam)
    data.place(relx=x-0.1, rely=y+0.05, relwidth=0.15, relheight=0.05)
    #return data

root = tk.Tk()
root.title('Conciliacion Bancaria')
root.geometry('500x550')
root.config(bg= '#fff')
month = StringVar()
year = StringVar()
texto1 = Label(root, text= "Extractos Mensuales")
texto1.pack()
texto1.config(fg = colorTeam, bg= "white" ,font= ("Verdana", 16))
texto1.place(relx=0.05,rely=0.08)
escribirDatos("Año", year, 0.85, 0.05)
escribirDatos("Mes", month, 0.85, 0.12)
createbutton2(librosAuxiliar, 'LIBROS', 0.7, 0.20)
createbutton2(conciliarBank, 'CONCILIAR', 0.7, 0.35)
createbutton(openMonth, 'ABRIR CARPETA', 0.05, 0.05)
createbutton(consolidateMonth, 'GENERAR CONSOLIDADO', 0.05, 0.17)
texto2 = Label(root, text= "Extractos Diarios")
texto2.pack()
texto2.config(fg = colorTeam, bg= "white" ,font= ("Verdana", 16))
texto2.place(relx=0.05,rely=0.48)
createbutton(openDay, 'ABRIR CARPETA', 0.05, 0.45)
createbutton(consolidateDay, 'GENERAR CONSOLIDADO', 0.05, 0.57)
createbutton(openFilters, 'FILTROS', 0.05, 0.75)
imagen = PhotoImage(file ='logo.png')
lbImagen = Label(root, image= imagen, bd=0, fg="white").place(relx=0.70,rely=0.7)
marca = Label(root, text= "Desarrollado por Sebastian Moncada Cel:324-221-0852 ")
marca.pack()
marca.config(fg = "black", bg= "white" ,font= ("Verdana", 8))
marca.place(relx=0.01,rely=0.95)




root.mainloop()
