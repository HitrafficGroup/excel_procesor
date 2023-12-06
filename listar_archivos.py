import os
# Obtiene el directorio actual
# Ruta del directorio que quieres listar
directorio = 'D:/Archivos fuente de mediciones/USUARIO MEDIO VOLTAJE 2021 2022/2021/2 FEBRERO 2021/USUARIOS DE MEDIO VOLTAJE'
archivos = os.listdir(directorio)
listado_archivos = []
for archivo in archivos:
    if archivo[-5:] == '.xlsx':
        pass
    elif archivo[-5:] == '.docx':
        pass
    elif archivo[-5:] == '.jpeg':
        pass
    else:
        listado_archivos.append(archivo)
for i in listado_archivos:
    print(i)
# names_unicos = []
# listado = []
# for i in listado_archivos:
#     if i[-7:] == '.pqm702':
#         pass
#     else:
#         aux = i[:-4]
#         if aux in names_unicos:
#             pass
#         else:
#             num_aux = aux.split()
#             number = num_aux[0]
#             n = int(number[:-1])
#             aux_dict = {'name':aux,'number':n}
#             names_unicos.append(aux)
#             listado.append(aux_dict)


# ordenado = sorted(listado, key=lambda x: x['number'])

# for name in ordenado:
#     print(name['name'])
