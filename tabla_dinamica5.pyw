import time
from openpyxl.utils import get_column_letter
import win32com.client as win32

# --- Configura tu archivo ---
ruta = r"C:\Inventario2\existenciasvalorizadas_d.xlsx"
nombre_hoja = "EXISTENCIASVALORIZADAS"

# --- Abre Excel en segundo plano ---
excel = win32.DispatchEx("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(ruta)
ws = wb.Worksheets(nombre_hoja)

pc = wb.PivotCaches().Create(SourceType=1, SourceData="EXISTENCIASVALORIZADAS!A1:J1081")
pt = pc.CreatePivotTable(TableDestination="EXISTENCIASVALORIZADAS!L1", TableName="TablaPivot")


# ===== 4) Guardar y cerrar =====
wb.Save()
wb.Close()
excel.Quit()
