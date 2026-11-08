from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

AREA = 120        # m2
ESPESOR = 0.10    # m
ADMIN, IMPREV, UTIL = 0.10, 0.05, 0.05   # AIU sobre costo directo

# Precios aproximados en COP, Pasto (Nariño), 2026. Son estimaciones: validar con cotizaciones locales.
items = [
    # (descripción, unidad, cantidad, valor unitario)
    ("Localización, replanteo y descapote", "m2", AREA, 4500),
    ("Conformación y compactación de subrasante (vibrocompactador)", "m2", AREA, 6500),
    ("Recebo compactado e=10 cm", "m3", round(AREA * 0.10, 2), 95000),
    ("Polietileno calibre 6 (barrera de humedad)", "m2", round(AREA * 1.1, 2), 2200),
    ("Malla electrosoldada 15x15 cm, Ø4.0 mm (M-084)", "m2", round(AREA * 1.1, 2), 11500),
    ("Formaleta lateral en madera para fundición", "ml", 44, 16000),
    ("Concreto 3000 psi premezclado (e=10 cm + 5% desperdicio)", "m3", round(AREA * ESPESOR * 1.05, 2), 650000),
    ("Mano de obra: vaciado, vibrado y regleado del concreto", "m2", AREA, 28000),
    ("Afinado y esmaltado de la superficie", "m2", AREA, 15000),
    ("Corte de juntas de control con cortadora", "ml", 60, 7000),
    ("Curado del concreto", "m2", AREA, 1800),
    ("Alquiler de vibrador, regla vibratoria y herramienta menor", "gl", 1, 250000),
]

wb = Workbook()
ws = wb.active
ws.title = "Presupuesto"

ws["A1"] = f"PRESUPUESTO - PLACA DE CONTRAPISO {AREA} m2 - PASTO, NARIÑO (2026)"
ws["A1"].font = Font(bold=True, size=14)
ws.merge_cells("A1:F1")
ws["A2"] = "Valores en pesos colombianos (COP). Precios aproximados, sujetos a cotización."
ws["A2"].font = Font(italic=True, color="666666")

headers = ["Ítem", "Descripción", "Unidad", "Cantidad", "Valor unitario", "Valor total"]
fill = PatternFill("solid", fgColor="1F4E78")
thin = Side(style="thin", color="999999")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
COP = '"$"#,##0'

HEADER_ROW = 4
for c, h in enumerate(headers, 1):
    cell = ws.cell(HEADER_ROW, c, h)
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = fill
    cell.alignment = Alignment(horizontal="center")
    cell.border = border

first = HEADER_ROW + 1
for i, (desc, un, qty, vu) in enumerate(items):
    r = first + i
    ws.cell(r, 1, i + 1)
    ws.cell(r, 2, desc)
    ws.cell(r, 3, un)
    ws.cell(r, 4, qty).number_format = "#,##0.00"
    ws.cell(r, 5, vu).number_format = COP
    ws.cell(r, 6, f"=D{r}*E{r}").number_format = COP
    for c in range(1, 7):
        ws.cell(r, c).border = border

last = first + len(items) - 1
cd = last + 1   # costo directo
rows = {
    "COSTO DIRECTO": cd,
    "Administración": cd + 1,
    "Imprevistos": cd + 2,
    "Utilidad": cd + 3,
    "TOTAL AIU": cd + 4,
    "VALOR TOTAL DEL PRESUPUESTO": cd + 5,
    "Costo por m2": cd + 6,
}
bold = Font(bold=True)

ws.cell(cd, 2, "COSTO DIRECTO").font = bold
ws.cell(cd, 6, f"=SUM(F{first}:F{last})").number_format = COP
for nombre, tasa in (("Administración", ADMIN), ("Imprevistos", IMPREV), ("Utilidad", UTIL)):
    r = rows[nombre]
    ws.cell(r, 2, f"{nombre} (A)" if nombre == "Administración" else f"{nombre} ({nombre[0]})")
    ws.cell(r, 4, tasa).number_format = "0%"
    ws.cell(r, 6, f"=F{cd}*D{r}").number_format = COP
ws.cell(rows["TOTAL AIU"], 2, "TOTAL AIU").font = bold
ws.cell(rows["TOTAL AIU"], 4, f"=SUM(D{rows['Administración']}:D{rows['Utilidad']})").number_format = "0%"
ws.cell(rows["TOTAL AIU"], 6, f"=SUM(F{rows['Administración']}:F{rows['Utilidad']})").number_format = COP
tt = rows["VALOR TOTAL DEL PRESUPUESTO"]
ws.cell(tt, 2, "VALOR TOTAL DEL PRESUPUESTO").font = bold
ws.cell(tt, 6, f"=F{cd}+F{rows['TOTAL AIU']}").number_format = COP
ws.cell(tt, 6).font = bold
ws.cell(rows["Costo por m2"], 2, "Costo por m2")
ws.cell(rows["Costo por m2"], 6, f"=F{tt}/{AREA}").number_format = COP

for col, w in zip("ABCDEF", [6, 62, 9, 12, 16, 18]):
    ws.column_dimensions[col].width = w

wb.save("presupuesto_contrapiso_pasto.xlsx")
print("Creado presupuesto_contrapiso_pasto.xlsx")
