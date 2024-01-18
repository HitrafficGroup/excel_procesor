import openpyxl
import pandas as pd
import os
from datetime import datetime

#funcion para ordenar diccionarios
datos_meses = {'Enero':1,'Febrero':2,'Marzo':3,'Abril':4,'Mayo':5,'Junio':6,'Julio':7,'Agosto':8,'Septiembre':9,'Octubre':10,'Noviembre':11,'Diciembre':12}
def ordenar_diccionario(total_data):
    for data in total_data:
        for clave, valor in datos_meses.items():
            if data['MES'] == clave:
                data['MES'] = valor
    ordenado_por_dia = sorted(total_data, key=lambda x: x['Fila'])         
    ordenado_por_mes = sorted(ordenado_por_dia, key=lambda x: x['MES'])
    ordenado_por_year = sorted(ordenado_por_mes, key=lambda x: x['YEAR'])
    return ordenado_por_year
#
encabezados = {
            'A': 'Fila','B': 'CODIGO_UNICO_NACIONAL', 'D': 'GEO-X', 'E': 'GEO-Y', 'G': 'PROVINCIA', 'H': 'CANTON',
            'I': 'SUBESTACION', 'J': 'ALIMENTADOR', 'M': 'F-F', 'N': 'F-N', 'O': 'FECHA_INICIO', 'P': 'HORA_INICIO',
            'Q': 'FECHA_FINAL', 'R': 'HORA_FINAL', 'S': 'N_REGISTROS', 'T': 'FA_V', 'U': 'FA_PST', 'V': 'FA_VTHD',
            'W': 'FB_V', 'X': 'FB_PST', 'Y': 'FB_VTHD', 'Z': 'FC_V', 'AA': 'FC_PST', 'AB': 'FC_VTHD', 'AC': 'DESEQUILIBRIO',
            'AD': 'FA5DV6', 'AE': 'FA6DV7', 'AF': 'FA7DV8', 'AG': 'FA8DV9', 'AH': 'FA9DV10', 'AI': 'FA10DV11', 'AJ': 'FA11DV12',
            'AK': 'FA12DV13', 'AL': 'FA13DV14', 'AM': 'FA14DV15', 'AN': 'FA15DV', 'AP': 'FB5DV6', 'AQ': 'FB6DV7', 'AR': 'FB7DV8',
            'AS': 'FB8DV9', 'AT': 'FB9DV10', 'AU': 'FB10DV11', 'AV': 'FB11DV12', 'AW': 'FB12DV13', 'AX': 'FB13DV14',
            'AY': 'FB14DV15', 'AZ': 'FB15DV', 'BB': 'FC5DV6', 'BC': 'FC6DV7', 'BD': 'FC7DV8', 'BE': 'FC8DV9', 'BF': 'FC9DV10',
            'BG': 'FC10DV11', 'BH': 'FC11DV12', 'BI': 'FC12DV13', 'BJ': 'FC13DV14', 'BK': 'FC14DV15', 'BL': 'FC15DV',
            'BN': 'OBSERVACIONES'
            }

#esta variable contendra la base de datos
data_base = []
def process_sheet(path,file_dir):
    workbook = openpyxl.load_workbook(path+'/'+file_dir,data_only=True)
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
    for fila in range(start_row,100):
        if sheet_target[f'B{fila}'].value == None:
            break
        dict_aux = {}
        for clave, valor in encabezados.items():
            cell_name = f'{clave}{fila}'
            current_cell = sheet_target[cell_name].value
            dict_aux[valor] = current_cell
        empty_dict = {}
        empty_dict['YEAR'] = sheet_target['D3'].value
        fecha  = sheet_target['D4'].value
        fecha_aux = datetime.strptime(str(fecha), "%Y-%m-%d %H:%M:%S")
        empty_dict['MES'] = fecha_aux.month

    
        ##  empieza formato de la fecha
        fecha_aux = str(dict_aux['FECHA_INICIO'])
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
        dict_aux['FECHA_INICIO'] = fecha_formated
        empty_dict['DIA'] = fecha_formated[0:2]
        empty_dict['FILE'] = file_dir
        fecha_aux = str(dict_aux['FECHA_FINAL'])
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
        dict_aux['FECHA_FINAL'] = fecha_formated
        ## desde aqui se deja de dar formato a la fecha
        empty_dict.update(dict_aux)
        data_captured.append(empty_dict)
    return data_captured

def calcular50(path_source,path_final):
    # primero revisamos la cantidad de excels que estan en el directorio actual
    # Obtiene el directorio actual
    # Ruta del directorio que quieres listar
    directorio = path_source

    # Obtener la lista de archivos en el directorio
    archivos = os.listdir(directorio)

    # Imprime la lista de archivos
    listado_archivos = []
    for archivo in archivos:
        if archivo[-5:] == '.xlsx':
            listado_archivos.append(archivo)
    for path_target in listado_archivos:
        data_base.extend(process_sheet(directorio,path_target))
    data_ordenada = ordenar_diccionario(data_base)
    df = pd.DataFrame(data_ordenada)
    df.to_excel(path_final, index=False)
    return data_ordenada