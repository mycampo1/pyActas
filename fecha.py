import openpyxl
from datetime import datetime
import os

# Ruta de tu archivo de Excel
archivo_excel = "C:/Users/Marlon Campo/Downloads/DRIVE NUEVO/ACTAS 2023-1/ACTAS 2023-1.xlsx"

# Abre el archivo de Excel
workbook = openpyxl.load_workbook(archivo_excel)

# Selecciona una hoja específica (puedes cambiar el nombre de la hoja)
hoja = workbook['Hoja2']

# Obtiene todas las celdas en la columna A
columna_A = hoja['A']

# Directorio actual
directorio_actual = os.getcwd()

# Itera a través de las celdas en la columna A
for celda in columna_A:
    # Verifica si la celda está vacía
    if celda.value is None:
        continue  # Continúa con la siguiente celda si está vacía

    # Agregar una impresión de depuración antes de la conversión de fecha
    print("Valor de la celda antes de la conversión:", celda.value)

    # Intenta convertir el texto a una fecha en el formato deseado "DD/MM/YYYY"
    try:
        fecha_objeto = datetime.strptime(celda.value, "%A, %d de %B de %Y")
        fecha_formateada = fecha_objeto.strftime("%d/%m/%Y")

        # Actualiza el valor de la celda con la fecha formateada
        celda.value = fecha_formateada
    except ValueError:
        # Si la conversión falla, simplemente sigue con la siguiente celda
        pass

    # Agregar una impresión de depuración después de la conversión de fecha
    print("Valor de la celda después de la conversión:", celda.value)

# Guarda los cambios en el archivo de Excel en la carpeta actual
nombre_archivo_modificado = "ACTAS_2023-1_modificado.xlsx"
archivo_modificado = os.path.join(directorio_actual, nombre_archivo_modificado)
workbook.save(archivo_modificado)

# Cierra el archivo de Excel
workbook.close()

