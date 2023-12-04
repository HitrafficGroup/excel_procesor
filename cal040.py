import openpyxl
import pandas as pd
import os


#funcion para ordenar diccionarios
def ordenar_diccionario(total_data):
    for data in total_data:
        for clave, valor in datos_meses.items():
            if data['MES'] == clave:
                data['MES'] = valor
    ordenado_por_dia = sorted(total_data, key=lambda x: x['MES'])
    ordenado_por_year = sorted(ordenado_por_dia, key=lambda x: x['YEAR'])
    return ordenado_por_year
#
datos_meses = {'Enero':1,'Febrero':2,'Marzo':3,'Abril':4,'Mayo':5,'Junio':6,'Julio':7,'Agosto':8,'Septiembre':9,'Octubre':10,'Noviembre':11,'Diciembre':12}
columnas  = ['B','C','D','E','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','W','Z','AG','AH','AI','AJ','AK','AL','AM','AN','AS','AT','AU','AV','AW','AX','AY','AZ','BE','BF','BG','BH','BI','BJ','BK','BL']
encabezados = ['CODIGO_UNICO_NACIONAL','TIPO','GEO-X','GEO-Y','PROVINCIA','CANTON','SUBESTACION','ALIMENTADOR','TRANSFORMADOR','NUM_FASES','F-F','F-N','FECHA_INICIO','HORA_INICIO','FECHA_FINAL','HORA_FINAL','N_REGISTROS','FA_V','FB_V','FC_V','FA8DV9','FA9DV10','FA10DV11','FA11DV12','FA12DV13','FA13DV14','FA14DV15','FA15DV','FB8DV9','FB9DV10','FB10DV11','FB11DV12','FB12DV13','FB13DV14','FB14DV15','FB15DV','FC8DV9','FC9DV10','FC10DV11','FC11DV12','FC12DV13','FC13DV14','FC14DV15','FC15DV']
index_data = {'columnas':columnas,'encabezados':encabezados}
#esta variable contendra la base de datos
data_base = []
def process_sheet(path,file_dir):
    workbook = openpyxl.load_workbook(path+'/'+file_dir)
    lista_de_hojas = workbook.sheetnames
    #esta variable alamacera toda la informacion recopilada de los excels
    data_captured = []
    #abrimos el excel
    #primer for para seleccionar la hoja de calculo que contenga la informacion
    target = ''
    for target_name in lista_de_hojas:
        if target_name[0:3] == 'CAL':
            target = target_name
            break

    # una vez que tenemos el nombre de la hoja que contiene los datos procedemos a abrir esa hoja
    #la variable woorbook contiene el excel con los datos
    sheet_target = workbook[target]
    start_row = 12
    if len(index_data['columnas']) == len(index_data['encabezados']):
        for fila in range(start_row,100):
            if sheet_target[f'{index_data["columnas"][0]}{fila}'].value == None:
                break
            values = []
            for column in index_data['columnas']:
                cell_name = f'{column}{fila}'
                current_cell = sheet_target[cell_name].value
                values.append(current_cell)
            empty_dict = {}
            empty_dict['YEAR'] = sheet_target['D3'].value
            fecha  = sheet_target['D4'].value
            fecha_aux  = fecha.split()
            empty_dict['DIA'] = fecha_aux[0]
            empty_dict['MES'] = fecha_aux[1]
            empty_dict['FILE'] = file_dir
            second_dict = dict(zip(index_data['encabezados'], values))
            ##  empieza formato de la fecha
            fecha_aux = str(second_dict['FECHA_INICIO'])
            fecha_formated = '0-0-0'
            if len(fecha_aux) > 10:
                year = fecha_aux[0:4]
                mes = fecha_aux[5:7]
                dia = fecha_aux[8:10]
                fecha_formated = f'{dia}-{mes}-{year}'
            else:
                dia = fecha_aux[0:2]
                mes = fecha_aux[3:5]
                year = fecha_aux[6:]
                fecha_formated = f'{dia}-{mes}-{year}'
            second_dict['FECHA_INICIO'] = fecha_formated
            fecha_aux = str(second_dict['FECHA_FINAL'])
            fecha_formated = '0-0-0'
            if len(fecha_aux) > 10:
                year = fecha_aux[0:4]
                mes = fecha_aux[5:7]
                dia = fecha_aux[8:10]
                fecha_formated = f'{dia}-{mes}-{year}'
            else:
                dia = fecha_aux[0:2]
                mes = fecha_aux[3:5]
                year = fecha_aux[6:]
                fecha_formated = f'{dia}-{mes}-{year}'
            second_dict['FECHA_FINAL'] = fecha_formated
            ## desde aqui se deja de dar formato a la fecha
            empty_dict.update(second_dict)
            data_captured.append(empty_dict)
        return data_captured

def calcular040():
    # primero revisamos la cantidad de excels que estan en el directorio actual
    # Obtiene el directorio actual
    # Ruta del directorio que quieres listar
    directorio = 'C:/Users/usuar/Documents/Python Scripts/excel_procesador/040'

    # Obtener la lista de archivos en el directorio
    archivos = os.listdir(directorio)

    # Imprime la lista de archivos
    listado_archivos = []
    for archivo in archivos:
        if archivo[-5:] == '.xlsx':
            listado_archivos.append(archivo)

    if len(index_data['columnas']) == len(index_data['encabezados']):
        for path_target in listado_archivos:
            data_base.extend(process_sheet(directorio,path_target))
        data_ordenada = ordenar_diccionario(data_base)
        df = pd.DataFrame(data_ordenada)
        df.to_excel('resultados/cal_040_bd.xlsx', index=False)
