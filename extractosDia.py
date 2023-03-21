import pandas as pd
import numpy as np


# file = 'files\\10CORP_01102022a311020229896827Extractos cuentas Corrientes (1).CSV'

# # with open(file) as archivo:
# #     excelFile=[]
# #     for linea in archivo:
# #         linea = linea.replace("\"","")
# #         excelFile.append(linea.split(sep=","))

# # print(excelFile[5][4])

# archivo = pd.read_csv(file, encoding='latin-1')
# archivo = np.asarray(archivo)
# print(archivo[3][0])


# elif code == 'bogota':
#         with open("files\\"+file) as archivo:
#             fileExcel=[]
#             for linea in archivo:
#                 linea = linea.replace("\"","")
#                 fileExcel.append(linea.split(sep=","))


numero = 'c'
x = numero.isdigit()
print(x)