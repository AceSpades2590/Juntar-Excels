import os
from pathlib import Path

class FileManager:
    @staticmethod
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

    @staticmethod
    def get_unique_filepath(filepath):
        base, ext = os.path.splitext(filepath)
        counter = 1
        new_filepath = filepath
        while os.path.exists(new_filepath):
            new_filepath = f"{base}({counter}){ext}"
            counter += 1
        return new_filepath

    @staticmethod
    def validate_path(ruta_carpeta):
        if not os.path.exists(ruta_carpeta):
            raise FileNotFoundError("La ruta especificada no existe.")
