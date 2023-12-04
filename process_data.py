import openpyxl
import pandas as pd 
import random
import os
import string
# Especifica la ruta del archivo Excel
# Funciones
def muestras_year(year,_data):
    muestras = []
    for month in range(1,13):
        counter = 0
        for dic in _data:
            if month == dic['MES'] and year == dic['YEAR']:
                counter += 1
        muestras.append(counter)
    return muestras

# Lee los archivos que hay esa carpeta
directorio = 'C:/Users/usuar/Documents/Python Scripts/excel_procesador/resultados'
archivos = os.listdir(directorio)

listado_archivos = []
for archivo in archivos:
    if archivo[-5:] == '.xlsx':
        listado_archivos.append(archivo)
def generarTablaMuestras(files_path):
    total_muestras = []
    for path in files_path:
        ruta_excel = 'resultados/'+path
        df = pd.read_excel(ruta_excel)
        diccionario_resultante = df.to_dict(orient='records')
        muestras_2021 = muestras_year(2021,diccionario_resultante)
        muestras_2022 = muestras_year(2022,diccionario_resultante)
        muestras_2021.extend(muestras_2022)
        total_muestras.append(muestras_2021)

    return total_muestras


first_table = generarTablaMuestras(listado_archivos)
second_table = []

workbook = openpyxl.load_workbook('plantilla.xlsx')
hoja_trabajo = workbook.active
letras = ['D','E','F','G','H']
for i in range(5):
    data = first_table[i]
    row = 5
    for j in range(24):
        hoja_trabajo[f'{letras[i]}{row}'] = data[j]
        row += 1

workbook.save("resultados.xlsx")
workbook.close()

