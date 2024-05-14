# Unión de Archivos Excel y ODS

Este script en Python te permite unir varios archivos Excel (.xlsx, .xls) y archivos ODS en una carpeta en un solo archivo, seleccionando un número específico de columnas de cada uno.

## Requisitos

- Python 3.12
- Pandas
- Openpyxl
- Odfpy

Para instalar las dependencias, puedes utilizar pip: pip install pandas openpyxl odfpy

## Uso

1. Clona o descarga este repositorio en tu máquina local.
2. Abre una terminal o línea de comandos y navega hasta el directorio donde se encuentra el script.
3. Ejecuta el script usando Python: python main.py
4. Se te solicitará ingresar la ruta de la carpeta donde están los archivos Excel y ODS que deseas unir, así como el número de columnas que deseas unir de cada archivo.
5. Una vez proporcionados los datos, el script combinará los archivos y generará un nuevo archivo con el formato elegido en tu carpeta de descargas.

## Formatos de archivo soportados

El script admite los siguientes formatos de archivo:

- Archivos Excel (.xlsx, .xls)
- Archivos ODS

## Notas

- Asegúrate de tener los permisos adecuados para acceder a la carpeta donde se encuentran los archivos Excel y ODS.
- Los archivos Excel y ODS deben tener la misma estructura de columnas para una combinación adecuada.

Si tienes alguna pregunta o encuentras algún problema, no dudes en [abrir un problema](https://github.com/AceSpades2590/Juntar-Excels/issues) en este repositorio.
