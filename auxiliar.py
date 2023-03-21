from tkinter.filedialog import askopenfilename as askf
from openpyxl import Workbook
from subprocess import Popen
from tkinter import Tk
import pandas as pd
import numpy as np

bancos={
    '11100505' :['BANCOLOMBIA', '325-450229-55'],
    '11100510' :['BANCO DE BOGOTA', '538-0393-14'],
    '11100515' :['BANCO DAVIVIENDA CORRIENTE', '3996699969-73'],
    '11100520' :['COLPATRIA MEDELLIN TAT', '6921005278'],
    '11200505' :['BANCOLOMBIA', '325-455868-34'],
    '11200524' :['BOGOTA (CALI POPAYAN)', '538-1765-87'],
    '11200530' :['CONVENIO URABA 11418', '645-626385-88'],
    '11200545' :['CONVENIO CHOCO 11419', '536-629632-21'],
    '11200558' :['CONVENIO CAUCASIA 11423', '371-000005-45'],
    '11200575' :['RECAUDOS TAT', '551-798737-44'],
    '11200586' :['CONVENIO EJE CAFETERO 11422', '859-000119-92'],
    '11200593' :['CONVENIO BOYACA 11421', '358-972792-13'],
    '11200595' :['BANCO AGRARIO', '4-1331-301395-5'],
    '11200598' :['CONVENIO OR OCC 11420', '024-470946-09'],
    '11100540' :['7187 DAVIVIENDA', '399669997187'],
    '11200521' :['CAJA SOCIAL AHORROS', '24115476024'],
    '11100545' :['CAJA SOCIAL CORRIENTE', '21004128844']
}

prefijos = [
    ["1C", "Principal"],
    ["38C", "El tesoro"],
    ["33C", "Pereira Bolivar"],
    ["35C", "Santa Rosa de Osos"],
    ["37C", "Popayan"],
    ["42C", "Sogamoso 2"],
    ["43C", "Zipaquira"],
    ["7C", "Sauces"],
    ["8C", "Istmina"],
    ["18C", "Alpujarra"],
    ["19C", "Caucasia CVS"],
    ["20C", "Sogamoso 1"],
    ["23C", "TAT"],
    ["24C", "Tunja Nieves"],
    ["27C", "La Dorada"],
    ["28C", "Armenia"],
    ["2C", "Apartado 1"],
    ["32C", "Santafe de Antioquia"],
    ["4C", "Turbo"],
    ["5C", "Caucasia"],
    ["6C", "Oviedo"],
    ["T44CS", "La Aurora"],
    ["22C", "Manizales"],
    ["T41CS", "Suba"],
    ["10C", "Andes"],
    ["3C", "Quibdo"],
    ["31C", "Chiquinquira"],
    ["34C", "CALI LA CASONA"], 
    ["36C", "CALI LA 11"],
    ["39C", "MARINILLA PARQUE"], 
    ["46C", "PAMPALINDA"],
    ["17C", "APARTADO2"],
    ["25C", "MARINILLA"],
    ["26C", "TUNJA CVS"],
    ["29C", "PEREIRA ÉXITO"],
    ["35C", "SANTAROSA DE OSOS"],
    ["14C", "RIONEGRO"],
    ["T45CS", "VILLA LUZ"],
    ["0000", "Principal"],
]

document = []

def excel(data): # Generar archivo de excel
    wb = Workbook()
    hoja = wb.active
    hoja.append(('Cuenta', 'CUENTA BANCARIA', 'Documento Ref.', 'Fecha', 'Nro Registro', 'Comprobante', 'Documento', 'SUCURSAL', 'Débitos'))
    celda=2
    for fila in data:
        if (str(fila[2]).replace(" ","") != "nan"):
            if int(fila[8]) > 0:
                hoja.append(fila)
                hoja.cell(row=celda, column=2).number_format="0"
                celda+=1
    wb.save('Consolidado.xlsx')
    p = Popen("openExcel.bat")
    stdout, stderr = p.communicate()

def organizeAuxiliar():
    start = True
    filename = askf()
    if filename != '':
        excel = pd.read_excel(filename)
        data = np.asarray(excel)
        for i in data:
            if start != True:
                cuenta = str(i[0]).replace(" ","")
                try:
                    banco = bancos[cuenta][1]
                except Exception:
                    banco = ""
                docRef = i[4]
                fecha = str(i[5])[3:5]+"/"+str(i[5])[0:2]+str(i[5])[5:10]
                registro = i[6]
                comprobante = i[7]
                documento = i[9]
                sucursal = "N/A"
                debito = i[14]
                for j in prefijos:
                    if j[0] in str(documento):   
                        sucursal = j[1]
                tempData = cuenta, banco, docRef, fecha, registro, comprobante, documento, sucursal, debito
                document.append(tempData)
            if str(i[0]).replace(" ","") == "Cuenta": start = False
def run():
    organizeAuxiliar()
    print(len(document))
    excel(document)

if __name__ == '__main__':
    run()
