import os
import pandas as pd
from FileManager import FileManager
from ODSReader import ODSReader

class ExcelUnifier:
    def __init__(self, ruta_carpeta, num_columnas):
        self.ruta_carpeta = ruta_carpeta
        self.num_columnas = num_columnas
        self.dataframes = []

    def load_files(self):
        archivos = sorted(os.listdir(self.ruta_carpeta))
        for archivo in archivos:
            file_path = os.path.join(self.ruta_carpeta, archivo)
            if archivo.endswith(('.xlsx', '.xls')):
                self.load_excel(file_path, archivo)
            elif archivo.endswith('.ods'):
                self.load_ods(file_path, archivo)

    def load_excel(self, file_path, archivo):
        try:
            df = pd.read_excel(file_path, header=None)
            df = df.iloc[:, :self.num_columnas]
            self.dataframes.append(df)
        except Exception as e:
            print(f"Error al leer el archivo {archivo}: {e}")

    def load_ods(self, file_path, archivo):
        try:
            df = ODSReader.read_ods(file_path, self.num_columnas)
            self.dataframes.append(df)
        except Exception as e:
            print(f"Error al leer el archivo {archivo}: {e}")

    def combine_and_save(self):
        if not self.dataframes:
            print("No se encontraron archivos válidos para unir.")
            return

        df_final = pd.concat(self.dataframes, ignore_index=True)
        carpeta_descargas = FileManager.get_downloads_folder()

        while True:
            print("Seleccione el formato para guardar el archivo:")
            print("1. .xlsx")
            print("2. .ods")
            opcion = input("Ingrese 1 o 2: ").strip()
            if opcion == '1':
                extension = 'xlsx'
                break
            elif opcion == '2':
                extension = 'ods'
                break
            else:
                print("Opción no válida. Por favor, ingrese 1 o 2.")

        ruta_guardado = os.path.join(carpeta_descargas, f'resultado.{extension}')
        if os.path.exists(ruta_guardado):
            sobreescribir = input(f"El archivo 'resultado.{extension}' ya existe. ¿Desea sobrescribirlo? (si/no): ").strip().lower()
            if sobreescribir != 'si':
                ruta_guardado = FileManager.get_unique_filepath(ruta_guardado)

        self.save_dataframe(df_final, ruta_guardado, extension)

    @staticmethod
    def save_dataframe(df_final, ruta_guardado, extension):
        try:
            if extension == 'xlsx':
                df_final.to_excel(ruta_guardado, index=False, header=False)
            elif extension == 'ods':
                from odf.opendocument import OpenDocumentSpreadsheet
                from odf.table import Table, TableRow, TableCell
                from odf.text import P

                doc = OpenDocumentSpreadsheet()
                table = Table(name="Sheet1")

                for row in df_final.values:
                    tr = TableRow()
                    for cell in row:
                        tc = TableCell()
                        p = P(text=str(cell))
                        tc.addElement(p)
                        tr.addElement(tc)
                    table.addElement(tr)

                doc.spreadsheet.addElement(table)
                doc.save(ruta_guardado)

            print(f"¡Excels unidos satisfactoriamente! El archivo se ha guardado en {ruta_guardado}.")
        except Exception as e:
            print(f"Error al guardar el archivo: {e}")
