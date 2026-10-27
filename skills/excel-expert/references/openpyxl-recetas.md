# Recetas openpyxl

## Abrir y guardar

```python
from openpyxl import Workbook, load_workbook
wb = load_workbook("in.xlsx")                       # fórmulas como texto
wb = load_workbook("in.xlsx", data_only=True)       # valores cacheados (¡no guardar!)
wb = load_workbook("in.xlsm", keep_vba=True)        # conservar macros; guardar como .xlsm
wb = load_workbook("big.xlsx", read_only=True)      # streaming de lectura (ver rendimiento.md)
wb.save("out.xlsx")
```
- Plantillas: `.xltx` → `wb.template = False` antes de guardar como `.xlsx`.
- `.xls` (binario antiguo), `.xlsb` y `.ods` no los abre openpyxl → `scripts/convert.py`.
- Libro protegido con contraseña de apertura: openpyxl no puede abrirlo; pide al usuario una
  copia sin contraseña (no intentes romperla).

## Hojas

```python
ws = wb.active; ws = wb["Ventas"]; ws = wb.worksheets[0]
ws = wb.create_sheet("Resumen", 0)          # en posición 0
wb.move_sheet("Resumen", offset=-1)
copia = wb.copy_worksheet(wb["Plantilla"])  # misma hoja del mismo libro (no copia gráficos/imágenes)
copia.title = "Enero"
del wb["Hoja1"]
ws.sheet_state = "hidden"                   # "veryHidden" la oculta del menú de Excel
ws.sheet_properties.tabColor = "1F3864"
ws.sheet_view.showGridLines = False
ws.sheet_view.zoomScale = 90
wb.active = wb.sheetnames.index("Resumen")  # hoja activa al abrir
```
Copiar hojas **entre libros** no está soportado: recrea valores y estilos celda a celda
(`dst.cell(...).value = src.value; dst._style = copy(src._style)` solo vale dentro del mismo
libro; entre libros copia `font`, `fill`, `border`, `alignment`, `number_format` con `copy()`).

## Celdas

```python
ws["A1"] = "Texto"; ws.cell(row=2, column=3, value=12.5)
ws["B2"] = datetime.date(2026, 3, 31); ws["B2"].number_format = "dd/mm/yyyy"
ws["C2"] = "=SUM(C3:C10)"
ws.append([1, 2, 3])                        # añade al final de la hoja
for row in ws.iter_rows(min_row=2, values_only=True): ...
for col in ws.iter_cols(min_col=2, max_col=4): ...
ws.max_row, ws.max_column, ws.dimensions    # ¡max_row cuenta celdas con solo formato!
```
Última fila con datos real:
```python
last = max((c.row for c in ws["A"] if c.value not in (None, "")), default=0)
```
Texto que empieza por `=` pero NO es fórmula: `cell.value = "=hola"; cell.data_type = "s"`.
Números como texto conservando ceros (códigos postales, NIF): escribe la cadena y
`number_format = "@"`.

## Estilos

```python
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment, NamedStyle
c.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
c.fill = PatternFill("solid", start_color="1F3864")
c.border = Border(bottom=Side(style="thin", color="BFBFBF"))
c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True, indent=1)
c.number_format = "#,##0.00"
```
- Los objetos de estilo son inmutables: para cambiar solo un atributo, `c.font = c.font.copy(bold=True)`
  (o crea uno nuevo).
- Estilo con nombre reutilizable (mejor rendimiento y coherencia):
```python
st = NamedStyle(name="moneda"); st.font = Font(name="Arial", size=10); st.number_format = '#,##0.00 "€"'
wb.add_named_style(st); c.style = "moneda"
```
- Copiar estilo de otra celda del mismo libro: `dst._style = copy(src._style)`.

## Filas, columnas, tamaño

```python
ws.column_dimensions["A"].width = 30
ws.row_dimensions[1].height = 28
ws.column_dimensions.group("D", "F", hidden=True)   # agrupar/ocultar
ws.row_dimensions.group(5, 10, outline_level=1)
ws.freeze_panes = "B2"                               # fija fila 1 y columna A
ws.auto_filter.ref = "A1:F200"
```
Inserción/borrado: **usa `xlsx_helpers.insert_rows_safe`** (ver edicion-segura.md).

## Celdas combinadas

```python
ws.merge_cells("A1:F1"); ws["A1"] = "Título"   # escribir solo en la superior izquierda
ws.unmerge_cells("A1:F1")
for rng in ws.merged_cells.ranges: print(rng.coord)
```
Evita combinar dentro de tablas de datos (rompe ordenación y filtros); usa
`Alignment(horizontal="centerContinuous")` para centrar un título sobre varias columnas.

## Comentarios e hipervínculos

```python
from openpyxl.comments import Comment
c.comment = Comment("Fuente: INE, EPA T2-2026", "Analista")
c.hyperlink = "https://www.ine.es"; c.style = "Hyperlink"
c.hyperlink = "#'Resumen'!A1"                       # enlace interno
```

## Nombres definidos

```python
from openpyxl.workbook.defined_name import DefinedName
wb.defined_names["TipoIVA"] = DefinedName("TipoIVA", attr_text="Supuestos!$B$2")
ws["C5"] = "=B5*TipoIVA"
# nombre local a una hoja:
ws.defined_names["Local"] = DefinedName("Local", attr_text="Datos!$A$1:$A$10")
for name, dn in wb.defined_names.items(): print(name, dn.attr_text)
```

## Imágenes

```python
from openpyxl.drawing.image import Image   # requiere Pillow
img = Image("logo.png"); img.width, img.height = 160, 48
ws.add_image(img, "A1")
```

## Protección

```python
ws.protection.sheet = True; ws.protection.password = "clave"
for row in ws["B2:B10"]:
    for c in row: c.protection = c.protection.copy(locked=False)   # celdas editables
wb.security.lockStructure = True
```
La protección de hoja de Excel es disuasoria, no seguridad real; dilo si el usuario la usa
para proteger datos sensibles.

## Impresión

```python
ws.page_setup.orientation = "landscape"; ws.page_setup.paperSize = ws.PAPERSIZE_A4
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
ws.print_title_rows = "1:1"; ws.print_area = "A1:H50"
ws.oddFooter.center.text = "Página &P de &N"
ws.page_margins.left = ws.page_margins.right = 0.5
```

## Recálculo al abrir en Excel

```python
wb.calculation.fullCalcOnLoad = True   # Excel recalcula todo al abrir
```
Útil si entregas sin pasar por `recalc.py`; los lectores que no son Excel seguirán sin valores.

## Propiedades del documento

```python
wb.properties.title = "Informe de ventas"; wb.properties.creator = "Equipo Finanzas"
```
