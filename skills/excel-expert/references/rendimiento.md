# Ficheros grandes y rendimiento

## Límites de Excel

1.048.576 filas × 16.384 columnas (XFD) por hoja; 32.767 caracteres por celda; nombres de
hoja ≤31 caracteres; 64.000 formatos de celda distintos por libro (crear un `Font()` nuevo
por celda en un bucle puede agotarlos → reutiliza objetos o usa `NamedStyle`).
`.xls` antiguo: 65.536 filas × 256 columnas.

Si los datos superan el límite: repartir en varias hojas, agregar, o entregar CSV/Parquet
además del Excel.

## Lectura rápida

```python
wb = load_workbook("big.xlsx", read_only=True, data_only=True)
ws = wb["Datos"]
for row in ws.iter_rows(min_row=2, values_only=True): ...
wb.close()
```
- En `read_only` no hay acceso aleatorio eficiente (`ws["A1"]` relee) ni celdas combinadas.
- `pandas.read_excel(..., engine="calamine")` (paquete `python-calamine`) es mucho más rápido
  si está disponible.
- Si solo necesitas unas columnas: `usecols=`.

## Escritura rápida

**openpyxl write_only** (streaming, memoria constante):
```python
from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
wb = Workbook(write_only=True); ws = wb.create_sheet("Datos")
hdr = []
for h in cols:
    c = WriteOnlyCell(ws, value=h); c.font = Font(bold=True); hdr.append(c)
ws.append(hdr)
for rec in registros: ws.append(rec)
wb.save("big.xlsx")
```
Solo se puede añadir filas en orden; no se puede leer ni modificar lo escrito.

**xlsxwriter** (más rápido, mejor soporte de gráficos/formato condicional, pero NO puede
editar ficheros existentes):
```python
import xlsxwriter
wb = xlsxwriter.Workbook("out.xlsx", {"constant_memory": True})
ws = wb.add_worksheet("Datos")
money = wb.add_format({"num_format": '#,##0.00 "€"', "font_name": "Arial"})
ws.write_row(0, 0, cols, wb.add_format({"bold": True, "bg_color": "#1F3864", "font_color": "white"}))
for i, rec in enumerate(registros, 1): ws.write_row(i, 0, rec)
ws.write_formula("D1000", "=SUM(D2:D999)", money)
wb.close()
```
Con pandas: `df.to_excel(path, engine="xlsxwriter")`.
xlsxwriter sí soporta funciones de matriz dinámica correctamente
(`ws.write_dynamic_array_formula`) si el usuario las exige.

## Recalcular libros grandes

`recalc.py libro.xlsx --timeout 300`. Si sigue superando el tiempo: reduce columnas
completas en SUMPRODUCT/matriciales, sustituye `INDIRECT`/`OFFSET` (volátiles) por `INDEX`,
o valida sobre una muestra.

## Buenas prácticas

- Construye las filas en listas de Python y escribe en bloque; evita leer y escribir celda a
  celda alternando.
- Formatea columnas enteras con un bucle sobre el rango usado, no sobre 1.048.576 filas.
- Evita formato condicional sobre columnas completas.
