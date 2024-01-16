import pandas as pd

def calcularPorcent(data,porcent,final_path,condition):
    df = pd.DataFrame(data)
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
            
            muestra_generada.extend(aux)
    print(final_path)
    df = pd.DataFrame(muestra_generada)
    df.to_excel(final_path, index=False)