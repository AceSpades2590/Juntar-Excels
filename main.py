import os
import pandas as pd
# Instalar Pandas "pip install pandas"
# Instalar Openpyxl "pip install openpyxl"
from pathlib import Path


def unir_excels(ruta_carpeta, num_columnas, num_filas):
    # Lista para almacenar los DataFrames de cada archivo Excel
    dataframes = []

    # Obtener la lista de archivos en la carpeta y ordenarla alfabéticamente
    archivos = sorted(os.listdir(ruta_carpeta))

    # Iterar sobre los archivos ordenados en la carpeta
    for archivo in archivos:
        # Verificar si el archivo es un Excel
        if archivo.endswith('.xlsx') or archivo.endswith('.xls'):
            # Leer el archivo Excel y seleccionar todas las columnas
            df = pd.read_excel(os.path.join(ruta_carpeta, archivo), nrows=num_filas, header=None)
            # Seleccionar solo las columnas deseadas
            df = df.iloc[:, :num_columnas]
            # Agregar el DataFrame a la lista
            dataframes.append(df)

    # Combinar los DataFrames en uno solo
    df_final = pd.concat(dataframes)

    # Obtener la ruta de la carpeta de descargas del usuario
    carpeta_descargas = str(Path.home() / "Downloads")

    # Guardar el DataFrame combinado en un nuevo archivo Excel en la carpeta de descargas
    ruta_guardado = os.path.join(carpeta_descargas, 'resultado.xlsx')
    df_final.to_excel(ruta_guardado, index=False, header=False)  # No incluir encabezado

    print("¡Excels unidos satisfactoriamente! El archivo se ha guardado en la carpeta de descargas.")


# Pedir al usuario la ruta de la carpeta y el número de columnas y filas
# Reemplazamos barras inclinadas por barras invertidas para Windows
ruta_carpeta = input("Ingrese la ruta de la carpeta donde están los archivos Excel: ").replace('/', '\\')
num_columnas = int(input("Ingrese el número de columnas a unir: "))
num_filas = int(input("Ingrese el número de filas a unir: "))

# Llamar a la función unir_excels con los datos proporcionados por el usuario
unir_excels(ruta_carpeta, num_columnas, num_filas)
