# Gráficos con openpyxl

Los gráficos de openpyxl son nativos de Excel (se actualizan con los datos). Se definen por
referencias a rangos, no por valores.

## Atajos (`xlsx_helpers`)

```python
add_bar_chart(ws, "B1:D13", "A2:A13", "F2", title="Ventas por región", y_title="€")
add_bar_chart(ws, "B1:D13", "A2:A13", "F20", title="Mix", stacked=True)
add_bar_chart(ws, "B1:B10", "A2:A10", "F2", horizontal=True)      # ranking
add_line_chart(ws, "B1:C25", "A2:A25", "F2", title="Evolución mensual")
add_pie_chart(ws, "B1:B6", "A2:A6", "F2", title="Cuota")
add_scatter_chart(ws, "B2:B50", ["C2:C50"], "F2", x_title="Precio", y_title="Unidades")
# Datos en otra hoja:
add_line_chart(wb["Dashboard"], "B1:B13", "A2:A13", "B2", data_sheet=wb["Datos"])
```
`data_ref` incluye la fila de cabecera (se usa como nombre de serie); `cats_ref` no.

## Elegir el gráfico

| Mensaje | Tipo |
|---|---|
| Evolución en el tiempo | líneas (columnas si son pocos periodos) |
| Comparar categorías | columnas; barras horizontales si las etiquetas son largas o es un ranking |
| Composición (≤6 partes) | tarta/anillo; si hay más o varios periodos → columnas apiladas |
| Relación entre dos variables | dispersión |
| Real vs objetivo | columnas + línea (combinado) |

Ordena las categorías (salvo tiempo), títulos que digan la conclusión si procede
("Las ventas online crecen un 18 %"), sin 3D, sin leyenda si hay una sola serie.

## API completa (cuando necesites más control)

```python
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList

ch = BarChart(); ch.type = "col"; ch.grouping = "clustered"
data = Reference(ws, min_col=2, min_row=1, max_col=3, max_row=13)
cats = Reference(ws, min_col=1, min_row=2, max_row=13)
ch.add_data(data, titles_from_data=True); ch.set_categories(cats)
ch.title = "Ventas"; ch.y_axis.title = "€"; ch.y_axis.numFmt = '#,##0'
ch.y_axis.majorGridlines = None              # sin cuadrícula
ch.x_axis.delete = False; ch.y_axis.delete = False   # ejes visibles (openpyxl ≥3.1)
ch.legend.position = "b"                     # o ch.legend = None
ch.width, ch.height = 18, 9                  # cm
ch.series[0].graphicalProperties.solidFill = "1F3864"
ch.series[0].graphicalProperties.line.solidFill = "1F3864"
ch.dataLabels = DataLabelList(); ch.dataLabels.showVal = True
ws.add_chart(ch, "F2")
```

### Combinado columnas + línea en eje secundario

```python
bar = BarChart(); bar.add_data(Reference(ws, min_col=2, min_row=1, max_row=13), titles_from_data=True)
bar.set_categories(Reference(ws, min_col=1, min_row=2, max_row=13))
line = LineChart(); line.add_data(Reference(ws, min_col=3, min_row=1, max_row=13), titles_from_data=True)
line.y_axis.axId = 200; line.y_axis.title = "Margen %"; line.y_axis.crosses = "max"
bar.y_axis.title = "Ventas €"
bar += line
ws.add_chart(bar, "F2")
```

## Notas

- Los gráficos se dibujan en la hoja donde se añaden; pueden apuntar a datos de otra hoja.
- Si los datos son fórmulas, el gráfico no tendrá valores hasta recalcular (Excel lo hace al
  abrir; `recalc.py` para otros visores).
- Al cargar y volver a guardar un libro con gráficos complejos, openpyxl puede simplificar su
  formato. Si importa conservarlos tal cual → `patch_cells.py`.
- Para una imagen estática (p.ej. para un informe en PDF) usa matplotlib y `ws.add_image`,
  pero en Excel prefiere siempre gráficos nativos.
