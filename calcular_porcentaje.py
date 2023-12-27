import pandas as pd
import random
import math

def calcularPorcent(ruta_excel,porcent,final_path):
    try:
        aux_porcent = float(porcent)
        if aux_porcent >=1:
            aux_porcent = 0.08
    except:
        aux_porcent = 0.08
    df = pd.read_excel(ruta_excel)
    diccionario_resultante = df.to_dict(orient='records')
    size = len(diccionario_resultante)
    N_poblation = size
    error = 0.05 #error del 5%
    confianza = 1.96 #confianza del 95%
    muestral_size = (N_poblation *(confianza**2)*0.5*0.5)/((error**2)*(N_poblation-1)+(confianza**2)*0.5*0.5)
    porcent_total = round(size*aux_porcent)
    print(f'el numero total de datos es: {N_poblation}')
    print(f'el total de la muestra con una confiabilidad del 95 porciento es: {round(muestral_size)}')
    print(f'el 8 porciento de la cantidad total de datos es: {porcent_total}')
    aletory_data = []
    # if round(muestral_size) > porcent:
    #     for _ in range(round(muestral_size)):
    #         while True:
    #             numero_aleatorio = random.randint(0, size-1)
    #             if numero_aleatorio in aletory_data:
    #                 pass
    #             else:
    #                 aletory_data.append(numero_aleatorio)
    #                 break
    # else:
    #     for _ in range(round(porcent)):
    #         while True:
    #             numero_aleatorio = random.randint(0, size-1)
    #             if numero_aleatorio in aletory_data:
    #                 pass
    #             else:
    #                 aletory_data.append(numero_aleatorio)
    #                 break
    for _ in range(round(porcent_total)):
        while True:
            numero_aleatorio = random.randint(0, size-1)
            if numero_aleatorio in aletory_data:
                pass
            else:
                aletory_data.append(numero_aleatorio)
                break

    muestra = []
    for indice in aletory_data:
        muestra.append(diccionario_resultante[indice])

    datos_ordenados_month = sorted(muestra, key=lambda x: x['MES'])
    datos_ordenados_year = sorted(datos_ordenados_month, key=lambda x: x['YEAR'])
    df = pd.DataFrame(datos_ordenados_year)
    df.to_excel(final_path, index=False)