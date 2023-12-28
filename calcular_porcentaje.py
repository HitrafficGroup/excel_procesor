import pandas as pd

def calcularPorcent(ruta_excel,porcent,final_path,condition):
    df = pd.read_excel(ruta_excel)
    # procesamos el dataframe
    valores_unicos_lista = df['YEAR'].unique().tolist()
    meses_unicos = []
    
    for year in valores_unicos_lista:
        c1 = df['YEAR'] == year
        datos_seleccionados = df[c1]
        meses_unicos.append({'meses':datos_seleccionados['MES'].unique().tolist(),'year':year})
    longitud_meses = 0
    
    for lon in meses_unicos:
        longitud_meses += len(lon['meses'])
    
    
    size = df.shape[0]
    total_porcent = 24
    if condition == 1:
        total_porcent = round(int(size)*porcent)
    else:
        N_poblation = size
        error = 0.05 
        confianza = 1.96 
        total_porcent = (N_poblation *(confianza**2)*0.5*0.5)/((error**2)*(N_poblation-1)+(confianza**2)*0.5*0.5)

    ctd_mes = round(total_porcent/longitud_meses)
    muestra_generada = []

    for dic in meses_unicos:
        c2 = df['YEAR'] == dic['year']
        data_frame_year = df[c2]
        for val in dic['meses']:
            c3 = data_frame_year['MES'] == val
            df_mes = data_frame_year[c3]
            elementos_aleatorios = df_mes.sample(n=ctd_mes, replace=False)
            aux = elementos_aleatorios.to_dict(orient='records')
            print(aux)
            
            muestra_generada.extend(aux)
    
        

    # for i in meses_unicos:
    #     for x in i:
    #         condition = df[]
    
    # #generamos un consolidado
    # diccionario_resultante = df.to_dict(orient='records')
    # size = len(diccionario_resultante)
   
    # print(f'el numero total de datos es: {N_poblation}')
    # print(f'el total de la muestra con una confiabilidad del 95 porciento es: {round(muestral_size)}')
    # print(f'el 8 porciento de la cantidad total de datos es: {porcent_total}')
    # aletory_data = []
    # # if round(muestral_size) > porcent:
    # #     for _ in range(round(muestral_size)):
    # #         while True:
    # #             numero_aleatorio = random.randint(0, size-1)
    # #             if numero_aleatorio in aletory_data:
    # #                 pass
    # #             else:
    # #                 aletory_data.append(numero_aleatorio)
    # #                 break
    # # else:
    # #     for _ in range(round(porcent)):
    # #         while True:
    # #             numero_aleatorio = random.randint(0, size-1)
    # #             if numero_aleatorio in aletory_data:
    # #                 pass
    # #             else:
    # #                 aletory_data.append(numero_aleatorio)
    # #                 break
    # # for _ in range(round(porcent_total)):
    # #     while True:
    # #         numero_aleatorio = random.randint(0, size-1)
    # #         if numero_aleatorio in aletory_data:
    # #             pass
    # #         else:
    # #             aletory_data.append(numero_aleatorio)
    # #             break

    # muestra = []
    # for indice in aletory_data:
    #     muestra.append(diccionario_resultante[indice])
  

    # datos_ordenados_month = sorted(muestra, key=lambda x: x['MES'])
    # datos_ordenados_year = sorted(datos_ordenados_month, key=lambda x: x['YEAR'])
    df = pd.DataFrame(muestra_generada)
    df.to_excel(final_path, index=False)