import pandas as pd
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P

class ODSReader:
    @staticmethod
    def get_cell_value(cell):
        paragraphs = cell.getElementsByType(P)
        return ''.join([str(p) for p in paragraphs])

    @staticmethod
    def read_ods(file_path, num_columnas):
        doc = load(file_path)
        table = doc.spreadsheet.getElementsByType(Table)[0]
        rows = table.getElementsByType(TableRow)

        data = []
        for row in rows:
            cells = row.getElementsByType(TableCell)
            data.append([ODSReader.get_cell_value(cell) for cell in cells[:num_columnas]])

        return pd.DataFrame(data)
