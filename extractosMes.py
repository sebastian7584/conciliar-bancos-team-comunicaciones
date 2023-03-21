from openpyxl import Workbook
from subprocess import Popen
import pandas as pd
import numpy as np
import pathlib

global consolidated
consolidated = [] # Inicio variable final
translation={}

banks_team_list = pd.read_csv('banks.csv')
banks_team_list = np.asarray(banks_team_list)
for i in banks_team_list:
    key = f'0{str(i[0])}' if int(i[0]) <10 else str(i[0])
    translation[key] = [i[1],i[2],i[4], i[3]]

# translation = { # informacion de las cuentas
#     "01": ['11100505', 'BANCOLOMBIA', '325-450229-55'],
#     "02": ['11200505', 'BANCOLOMBIA', '325-455868-34'],
#     "03": ['11200530', 'CONVENIO URABA 11418', '645-626385-88'],
#     "04": ['11200545', 'CONVENIO CHOCO 11419', '536-629632-21'],
#     "05": ['11200558', 'CONVENIO CAUCASIA 11423', '371-000005-45'],
#     "06": ['11200586', 'CONVENIO EJE CAFETERO 11422', '859-000119-92'],
#     "07": ['11200593', 'CONVENIO BOYACA 11421', '358-972792-13'],
#     "08": ['11200598', 'CONVENIO OR OCC 11420', '024-470946-09'],
#     "09": ['11200575', 'RECAUDOS TAT', '551-798737-44'],
#     "10": ['11100510', 'BANCO DE BOGOTA', '538-0393-14'],
#     "11": ['11100515', 'BANCO DAVIVIENDA CORRIENTE', '3996699969-73'],
#     "12": ['11100540', '7187 DAVIVIENDA', '399669997187'],
#     "13": ['11100520', 'COLPATRIA MEDELLIN TAT', '6921005278'],
#     "14": ['11200521', 'CAJA SOCIAL AHORROS', '24115476024'],
#     "15": ['11100545', 'CAJA SOCIAL CORRIENTE', '21004128844'],
#     "16": ['11200595', 'BANCO AGRARIO', '4-1331-301395-5'],
#     "17": ['11200524', 'BOGOTA (CALI POPAYAN)', '538-1765-87'],
# }

def orderAccount(file, period ): # Guia de orden datos en cada excel
    code = file[0:2]
    banco = translation[code][3]
    diccionary={
        "bancolombia":{1:["bancolombia", code, True,"",3,7,5],2:["bancolombia", code, True,"",3,7,5]},
        "bogota":{1:["bogota", code, False,"Fecha",0,1,3,4],2:["bogota", code, True,"",0,1,4,5]},
        "davivienda":{1:["davivienda", code, False,"FechadeSistema",0,2,7],2:["davivienda", code, True,"",0,2,7]},
        "colpatria":{1:["colpatria", code, False,"IDMOVIMIENTO",7,9,10,11],2:["colpatria", code, False,"IDMOVIMIENTO",7,9,10,11]},
        "caja":{1:["caja", code, False,"VALOR_TRANSACCION",4,2,0],2:["caja", code, False,"VALOR_TRANSACCION",4,2,0]},
        "agrario":{1:["agrario", code, False,"Fecha",0,2,3,4],2:["agrario", code, False,"Fecha",0,2,3,4]},
    }
    try:
        data = diccionary[banco][period]
        return data
    except KeyError:
        return None
    print(data)

    # if int(code) > 0 and int(code) < 10: # Bancolombia
    #     if period == 1 :
    #         return ["bancolombia", code, True,"",3,7,5]
    #     if period == 2 :
    #         return ["bancolombia", code, True,"",3,7,5]
    # elif code == "10": # Banco de bogota
    #     if period == 1 :
    #         return ["bogota", code, False,"Fecha",0,1,3,4]
    #     if period == 2 :
    #         return ["bogota", code, True,"",0,1,4,5] 
    # elif code == "11" or code == "12": # Davivienda
    #     if period == 1 :
    #         return ["davivienda", code, False,"FechadeSistema",0,2,7]
    #     if period == 2 :
    #         return ["davivienda", code, True,"",0,2,7]
    # elif code == "13": # Colpatria
    #     if period == 1 :
    #         return ["colpatria", code, False,"IDMOVIMIENTO",7,9,10,11]
    #     if period == 2 :
    #         return ["colpatria", code, False,"IDMOVIMIENTO",7,9,10,11]
    # elif code == "14" or code == "15": # Caja social
    #     if period == 1 :
    #         return["caja", code, False,"VALOR_TRANSACCION",4,2,0]
    #     if period == 2 :
    #         return["caja", code, False,"VALOR_TRANSACCION",4,2,0]
    # elif code == "16": # Banco agrario
    #     if period == 1 :
    #         return["agrario", code, False,"Fecha",0,2,3,4]
    #     if period == 2 :
    #         return["agrario", code, False,"Fecha",0,2,3,4]
    # else:
    #     return None
    
def files(folder): # Seleccionar archivos de la carpeta
    directory = pathlib.Path(folder)
    files = [x.name for x in directory.iterdir()]
    return files

def readExcel(file,code): # lee formato de excel y extrae la informacion
    if code == 'bancolombia':
        try:
            fileExcel = pd.read_csv("files\\"+file)
        except:
            fileExcel = pd.read_csv("files\\"+file, encoding='latin-1')
    elif code == 'bogota':
        with open("files\\"+file) as archivo:
            fileExcel=[]
            for linea in archivo:
                linea = linea.replace("\"","")
                fileExcel.append(linea.split(sep=","))
    else:
        fileExcel= pd.read_excel("files\\"+file, header=None)
    if code != 'bogota':
        data = np.asarray(fileExcel)
    else: data = fileExcel
    return data

def readExcelDay(file,code):
    file = "files2\\"+file
    if code == 'davivienda' or code == 'agrario': 
        excelFile= pd.read_excel(file)
    else:
        if code == 'bancolombia': 
            excelFile = pd.read_csv(file)
        elif code == 'caja': 
            with open(file) as archivo:
                excelFile=[]
                for linea in archivo:
                    excelFile.append(linea.split(sep=";"))
        else: 
            excelFile = pd.read_csv(file, encoding='latin-1')
    if code != 'caja': data = np.asarray(excelFile)
    else: data = excelFile
    return data

def removeSpaces(word): # Elimina espacios de un texto
    word = str(word).replace(" ","")
    return word

def organizeDateBancolombia(word): # Cambio formato fecha bancolombia
    word = str(word)
    date = word[6:8]+"/"+word[4:6]+"/"+word[0:4]
    return date

def organizeDateBogota(word, period, year, month):
    if int(month) < 10: month = '0'+ str(int(month))
    if period == 0:
        if (str(type(word)) != "<class 'datetime.datetime'>"):
            if (int(word) < 32 and int(word)> 0): fecha = year+"/"+month+"/"+str(word)
        else: fecha = word
    else:
        fecha = str(word[3:5])+"/"+month+"/"+year
    return fecha

def value(value1, value2): # organizar cuentas con deposito y retiro
    value11 = removeSpaces(value1)
    if removeSpaces(value1) == "nan": value1 = 0
    if removeSpaces(value2) == "nan": value2 = 0
    if int(value1) > int(value2): return value1
    else: return -value2

def organizeValueBogota(indice, data):
    debito = ""
    if len(data)>10:
        indice = 5
    if '$' in data[indice]:
        for i in range(indice, len(data)):
            if data[i] == '': break
            debito += data[i]
    else:
        debito += '-'
        for i in range(indice+1,len(data)):
            if data[i] == '': break
            if data[i][0].isdigit() or '$' in data[i][0]: 
                debito += data[i]
            else: break
    debito = debito.replace("$","")
    debito = debito.replace(".",",")
    debito = debito
    return debito



def organizeData(file,period, year, month): # Organiza los datos
    order = orderAccount(file,period)
    if period == 1: data = readExcel(file, order[0])
    else:
        if order != None: 
            data = readExcelDay(file, order[0])
    if order != None:
        cuenta = translation[order[1]][2]
        start = order[2]
        for i in data:
            if start:
                if order[0] == "bancolombia": fecha = organizeDateBancolombia(i[order[4]])
                elif order[0] == "bogota": fecha = organizeDateBogota(i[order[4]], period, year, month)
                else: fecha = i[order[4]]
                descripcion = i[order[5]]
                if period == 1:
                    if order[0] == 'bogota':
                        debito = organizeValueBogota(order[6],i)
                    elif len(order) == 7:
                        debito = i[order[6]]
                    else:
                        debito = value(i[order[6]],i[order[7]])
                else:
                    code = order[0]
                    if code == 'bancolombia': debito = str(i[order[6]]).replace(".",",")
                    if code == 'bogota':
                        if  str(i[order[6]]).replace(" ","") != 'nan':
                            debito = str(i[order[6]]).replace(",","")
                        else: debito = str(i[order[7]]).replace(",","")
                        debito = debito.replace("$","")
                        debito = debito.replace(".",",")
                    if code == 'davivienda': debito = i[order[6]]
                    if code == 'caja': debito = i[order[6]]
                    if code == 'agrario': 
                        if  str(i[order[6]]).replace(" ","") != 'nan':
                            debito = i[order[6]]
                        else:
                            debito = i[order[7]]
                tempData = cuenta, fecha, descripcion, debito
                consolidated.append(tempData)
            if removeSpaces(i[0]) == order[3]: start =True

def  debugData(): # Filtrar datos 
    tempData = []
    filters= pd.read_excel("filtros.xlsx")
    filters = np.asarray(filters)
    for i in range (0, len(consolidated)):
        x = removeSpaces(consolidated[i][2])
        x = x.lower()
        add = True
        for filter in filters:
            z = str(filter[1])
            y = removeSpaces(filter[0])
            y = y.lower()
            if y in x:
                if z == "DELETE":
                    add= False 
                    break
                if z == "NONE":
                    tempData.append(consolidated[i]) 
                    add=False
                    break
                else:
                    tempData.append(consolidated[i]+(z,))
                    add=False
                    break
        if add: tempData.append(consolidated[i])
    return tempData
     
def excel(data): # Generar archivo de excel
    wb = Workbook()
    hoja = wb.active
    hoja.append(('Cuenta Banco', 'Fecha', 'Descripcion', 'Debitos', 'Notas'))
    celda=2
    for fila in data:
        if (str(fila[2]).replace(" ","") != "nan"):
            hoja.append(fila)
            hoja.cell(row=celda, column=2).number_format="0"
            celda+=1
    wb.save('Consolidado.xlsx')
    p = Popen("openExcel.bat")
    stdout, stderr = p.communicate()




def run(year,month):
    finalconsolidated=[]
    folder = 'files'
    filesName = files(folder)
    for i in filesName:
        organizeData(i, 1, year, month)
    finalconsolidated = debugData()
    excel(finalconsolidated)

def runDay(year, month):
    finalconsolidated=[]
    folder = 'files2'
    filesName = files(folder)
    for i in filesName:
        organizeData(i,2, year, month)
    finalconsolidated = debugData()
    excel(finalconsolidated)



#run('2022','10')