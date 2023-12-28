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
    ordenado_por_dia = sorted(total_data, key=lambda x: x['DIA'])         
    ordenado_por_mes = sorted(ordenado_por_dia, key=lambda x: x['MES'])
    ordenado_por_year = sorted(ordenado_por_mes, key=lambda x: x['YEAR'])
    return ordenado_por_year
##### 
encabezados = { 
                'B': 'CODIGO', 'C': 'TIPO', 'D': 'GEO-X', 'E': 'GEO-Y', 'G': 'PROVINCIA', 'H': 'CANTON', 'I': 'SUBESTACION',
                'J': 'ALIMENTADOR', 'L': 'NUM DE FASES', 'M': 'F-F', 'N': 'F-N', 'O': 'FECHA_INICIO', 'P': 'HORA_INICIO',
                'Q': 'FECHA_FINAL', 'R': 'HORA_FINAL', 'S': 'N_REGISTROS', 'T': 'FA_V', 'U': 'FA_PST', 'V': 'FA_VTHD',
                'W': 'FB_V', 'X': 'FB_PST','Y': 'FB_VTHD', 'Z': 'FC_V', 'AA': 'FC_PST', 'AB': 'FC_VTHD', 'AC': 'DESEQUILIBRIO',
                'AG': 'FA_8D9','AH': 'FA_9D10', 'AI': 'FA_10D11', 'AJ': 'FA_11D12', 'AK': 'FA_12D13', 'AL': 'FA_13D14',
                'AM': 'FA_14D15','AN': 'FA_15D', 'AS': 'FB_8D9', 'AT': 'FB_9D10', 'AU': 'FB_10D11', 'AV': 'FB_11D12',
                'AW': 'FB_12D13','AX': 'FB_13D14', 'AY': 'FB_14D15', 'AZ': 'FB_15D', 'BE': 'FC_8D9', 'BF': 'FC_9D10',
                'BG': 'FC_10D11', 'BH': 'FC_11D12', 'BI': 'FC_12D13', 'BJ': 'FC_13D14', 'BK': 'FC_14D15', 'BL': 'FC_D15',
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
            if cell_name == f'T{fila}':
                aux_data = round(sheet_target[cell_name].value*100)
                dict_aux[valor] = aux_data
            elif cell_name == f'U{fila}':
                aux_data = round(sheet_target[cell_name].value*100,2)
                dict_aux[valor] = aux_data
            elif cell_name == f'V{fila}':
                aux_data = round(sheet_target[cell_name].value*100,2)
                dict_aux[valor] = aux_data
            elif cell_name == f'W{fila}':
                val = sheet_target[cell_name].value
                if val == None:
                    dict_aux[valor] = val
                else:
                    aux_data = round(val*100,2)
                    dict_aux[valor] = aux_data
            elif cell_name == f'X{fila}':
                val = sheet_target[cell_name].value
                if val == None:
                    dict_aux[valor] = val
                else:
                    aux_data = round(val*100,2)
                    dict_aux[valor] = aux_data
            elif cell_name == f'Y{fila}':
                val = sheet_target[cell_name].value
                if val == None:
                    dict_aux[valor] = val
                else:
                    aux_data = round(val*100,2)
                    dict_aux[valor] = aux_data
            elif cell_name == f'Z{fila}':
                val = sheet_target[cell_name].value
                if val == None:
                    dict_aux[valor] = val
                else:
                    aux_data = round(val*100,2)
                    dict_aux[valor] = aux_data
            elif cell_name == f'AA{fila}':
                val = sheet_target[cell_name].value
                if val == None:
                    dict_aux[valor] = val
                else:
                    aux_data = round(val*100,2)
                    dict_aux[valor] = aux_data
            elif cell_name == f'AC{fila}':
                print(f'fila procesando: {fila}')
                if sheet_target[cell_name].value == None:
                    dict_aux[valor] = ''
                else:
                    aux_data = round(sheet_target[cell_name].value*100,2)
                    dict_aux[valor] = aux_data
            else:
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

def calcular030(path_source,path_final):
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
        print(path_target)
        data_base.extend(process_sheet(directorio,path_target))
    data_ordenada = ordenar_diccionario(data_base)
    df = pd.DataFrame(data_ordenada)
    df.to_excel(path_final, index=False)





