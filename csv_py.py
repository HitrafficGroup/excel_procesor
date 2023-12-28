import csv

# Especifica la ruta del archivo CSV
archivo_csv = 'ruta/del/archivo.csv'

# Abre el archivo en modo lectura
with open(archivo_csv, 'r') as archivo:
    # Crea un objeto lector CSV con encabezados
    lector_csv = csv.DictReader(archivo)

    # Itera a través de las filas en el archivo CSV
    for fila in lector_csv:
        # La variable "fila" es un diccionario con los encabezados como claves
        print(fila['nombre_columna1'], fila['nombre_columna2'])
