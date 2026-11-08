"""Lee el presupuesto en Excel e indica qué ítem pesa más en el costo total.

Uso: python analizar_presupuesto.py [archivo.xlsx]
"""
import sys
from openpyxl import load_workbook

sys.stdout.reconfigure(encoding="utf-8")

archivo = sys.argv[1] if len(sys.argv) > 1 else "presupuesto_contrapiso_pasto.xlsx"
ws = load_workbook(archivo, data_only=True).active   # valores en caché de Excel

num = lambda v: isinstance(v, (int, float))
header_row = next(r for r in range(1, 20) if ws.cell(r, 2).value == "Descripción")

items, tasas = [], {}
for r in range(header_row + 1, ws.max_row + 1):
    desc, un, qty, vu, total = (ws.cell(r, c).value for c in (2, 3, 4, 5, 6))
    if not desc:
        continue
    if un and num(qty) and num(vu):                 # fila de ítem
        items.append((desc, total if num(total) else qty * vu))   # sin caché: calcula
    elif desc.startswith(("Administración", "Imprevistos", "Utilidad")) and num(qty):
        tasas[desc] = qty                            # tasa AIU (columna D)

if not items:
    sys.exit("No se encontraron ítems en el archivo.")

costo_directo = sum(t for _, t in items)
aiu = {k: costo_directo * v for k, v in tasas.items()}
total = costo_directo + sum(aiu.values())

items.sort(key=lambda x: x[1], reverse=True)
print(f"{'Ítem':<62}{'Valor (COP)':>14}{'% C.D.':>8}{'% Total':>9}")
print("-" * 93)
for desc, t in items:
    print(f"{desc:<62}{t:>14,.0f}{t / costo_directo:>8.1%}{t / total:>9.1%}")
print("-" * 93)
print(f"{'COSTO DIRECTO':<62}{costo_directo:>14,.0f}{'100.0%':>8}{costo_directo / total:>9.1%}")
for k, v in aiu.items():
    print(f"{k:<62}{v:>14,.0f}{'':>8}{v / total:>9.1%}")
print(f"{'VALOR TOTAL DEL PRESUPUESTO':<62}{total:>14,.0f}")

desc, t = items[0]
print(f"\nÍtem que más pesa: '{desc}'")
print(f"  {t:,.0f} COP = {t / costo_directo:.1%} del costo directo "
      f"({t / total:.1%} del total; {t * total / costo_directo:,.0f} COP con AIU incluido).")
