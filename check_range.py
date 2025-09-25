import win32com.client as win32
from openpyxl.utils import get_column_letter
import time

ruta = r"C:\Inventario2\existenciasvalorizadas_b.xlsx"
nombre_hoja = "EXISTENCIASVALORIZADAS"

excel = win32.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False
time.sleep(1)

wb = excel.Workbooks.Open(ruta)
ws = wb.Worksheets(nombre_hoja)

ultima_fila = ws.UsedRange.Rows.Count
ultima_col  = ws.UsedRange.Columns.Count

col_fin = get_column_letter(ultima_col)
print(f"Última fila: {ultima_fila}")
print(f"Última columna: {ultima_col} ({col_fin})")

wb.Close(False)
excel.Quit()
