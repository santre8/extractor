import pdfplumber
import re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

ruta_pdf = "reporte_quantfury.pdf"
ruta_excel = "tickers_quantfury.xlsx"

tickers_open = set()
tickers_closed = set()

# El patrón de fecha DD/MM/YYYY sigue siendo idéntico en ambos idiomas
patron_ticker = re.compile(r"^([A-Z0-9\.\-]+)\s+\d{2}/\d{2}/\d{4}")

seccion_actual = None

with pdfplumber.open(ruta_pdf) as pdf:
    for page in pdf.pages:
        texto = page.extract_text()
        if not texto:
            continue
        
        for linea in texto.split("\n"):
            linea_limpia = linea.strip()
            linea_lower = linea_limpia.lower()
            
            # Detectar cambio de sección en inglés o español
            if "open positions" in linea_lower or "posiciones abiertas" in linea_lower:
                seccion_actual = "OPEN"
                continue
            elif "closed positions" in linea_lower or "posiciones cerradas" in linea_lower:
                seccion_actual = "CLOSED"
                continue
            
            # Extraer ticker
            match = patron_ticker.match(linea_limpia)
            if match and seccion_actual:
                ticker = match.group(1).strip()
                # Descartar palabras de encabezado en ambos idiomas
                if ticker.upper() not in ["SYMBOL", "SÍMBOLO", "SIMBOLO", "TOTAL", "OPEN", "CLOSED"]:
                    if seccion_actual == "OPEN":
                        tickers_open.add(ticker)
                    elif seccion_actual == "CLOSED":
                        tickers_closed.add(ticker)

lista_open = sorted(list(tickers_open))
lista_closed = sorted(list(tickers_closed))

print(f"Tickers en Posiciones Abiertas / Open: {len(lista_open)}")
print(f"Tickers en Posiciones Cerradas / Closed: {len(lista_closed)}")

# Crear libro de Excel
wb = Workbook()

def llenar_hoja(ws, titulo_hoja, lista_tickers, color_hex):
    ws.title = titulo_hoja
    header_cell = ws.cell(row=1, column=1, value="Ticker")
    header_cell.font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_cell.fill = PatternFill(start_color=color_hex, end_color=color_hex, fill_type="solid")
    header_cell.alignment = Alignment(horizontal="center", vertical="center")
    
    for i, ticker in enumerate(lista_tickers, start=2):
        cell = ws.cell(row=i, column=1, value=ticker)
        cell.alignment = Alignment(horizontal="center")
        
    ws.column_dimensions["A"].width = 18

# 1. Pestaña Posiciones Abiertas
ws_open = wb.active
llenar_hoja(ws_open, "Open Positions", lista_open, "1F497D")

# 2. Pestaña Posiciones Cerradas
ws_closed = wb.create_sheet(title="Closed Positions")
llenar_hoja(ws_closed, "Closed Positions", lista_closed, "366092")

wb.save(ruta_excel)
print(f"\nArchivo generado exitosamente: {ruta_excel}")