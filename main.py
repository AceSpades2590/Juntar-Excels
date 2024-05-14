import os
import pandas as pd
from pathlib import Path
import time


def get_downloads_folder():
    home = str(Path.home())
    downloads_english = os.path.join(home, "Downloads")
    downloads_spanish = os.path.join(home, "Descargas")

    if os.path.exists(downloads_english):
        return downloads_english
    elif os.path.exists(downloads_spanish):
        return downloads_spanish
    else:
        raise FileNotFoundError("No se encontró la carpeta de descargas predeterminada.")


def get_valid_column_number():
    while True:
        try:
            num_columnas = int(input("Ingrese el número de columnas a unir: "))
            return num_columnas
        except ValueError:
            print("Por favor, ingrese un número entero válido.")


def get_unique_filepath(filepath):
    base, ext = os.path.splitext(filepath)
    counter = 1
    new_filepath = filepath
    while os.path.exists(new_filepath):
        new_filepath = f"{base}({counter}){ext}"
        counter += 1
    return new_filepath


def unir_excels(ruta_carpeta, num_columnas):
    dataframes = []

    archivos = sorted(os.listdir(ruta_carpeta))

    for archivo in archivos:
        if archivo.endswith('.xlsx') or archivo.endswith('.xls'):
            try:
                df = pd.read_excel(os.path.join(ruta_carpeta, archivo), header=None)
                df = df.iloc[:, :num_columnas]
                dataframes.append(df)
            except Exception as e:
                print(f"Error al leer el archivo {archivo}: {e}")

    if not dataframes:
        print("No se encontraron archivos Excel válidos para unir.")
        return

    # Combinar los DataFrames en uno solo
    df_final = pd.concat(dataframes, ignore_index=True)

    carpeta_descargas = get_downloads_folder()
    ruta_guardado = os.path.join(carpeta_descargas, 'resultado.xlsx')

    if os.path.exists(ruta_guardado):
        sobreescribir = input("El archivo 'resultado.xlsx' ya existe. ¿Desea sobrescribirlo? (si/no): ").strip().lower()
        if sobreescribir != 'si':
            ruta_guardado = get_unique_filepath(ruta_guardado)

    # Guardar el DataFrame
    df_final.to_excel(ruta_guardado, index=False, header=False)
    print(f"¡Excels unidos satisfactoriamente! El archivo se ha guardado en {ruta_guardado}.")


if __name__ == "__main__":
    print("Bienvenido a ExcelUnifier")

    while True:
        ruta_carpeta = input("Ingrese la ruta de la carpeta donde están los archivos Excel: ").replace('/', '\\')
        if not os.path.exists(ruta_carpeta):
            print("La ruta especificada no existe.")
            continue

        num_columnas = get_valid_column_number()
        unir_excels(ruta_carpeta, num_columnas)

        continuar = input("¿Desea unir más archivos Excel? (si/no): ").strip().lower()
        if continuar != 'si':
            print("Gracias por usar ExcelUnifier. El programa se cerrará en 5 segundos.")
            time.sleep(5)
            break
