import time
from FileManager import FileManager
from ExcelUnifier import ExcelUnifier
from utils import get_valid_column_number

def main():
    print("Bienvenido a ExcelUnifier")

    while True:
        ruta_carpeta = input("Ingrese la ruta de la carpeta donde están los archivos Excel: ").strip()
        ruta_carpeta = ruta_carpeta.replace("\\ ", " ")

        try:
            FileManager.validate_path(ruta_carpeta)
        except FileNotFoundError as e:
            print(e)
            continue

        num_columnas = get_valid_column_number()
        unifier = ExcelUnifier(ruta_carpeta, num_columnas)
        unifier.load_files()
        unifier.combine_and_save()

        continuar = input("¿Desea unir más archivos Excel? (si/no): ").strip().lower()
        if continuar != 'si':
            print("Gracias por usar ExcelUnifier. El programa se cerrará en 5 segundos.")
            time.sleep(5)
            break

if __name__ == "__main__":
    main()
