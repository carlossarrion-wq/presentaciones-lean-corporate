---
name: excel-expert
description: Crear, leer, editar, analizar, validar y convertir hojas de cálculo Excel (.xlsx, .xlsm, .xls, .csv, .ods). Úsalo siempre que un fichero de hoja de cálculo sea la entrada o la salida principal de la tarea - generar informes o modelos con fórmulas y formato profesional, modificar un libro existente sin romperlo, limpiar datos tabulares desordenados, añadir gráficos, tablas, validaciones o formato condicional, auditar fórmulas o convertir entre formatos. Use for any task whose primary input or output is a spreadsheet file.
---

# Excel Expert

Skill para que un agente trabaje con libros de Excel con la calidad de un analista senior:
fórmulas vivas, cero errores, formato profesional y ediciones que no rompen nada.

> Rutas de scripts relativas a la carpeta de este skill (`scripts/...`).
> Dependencias: `openpyxl`, `pandas`, `lxml` (Python ≥3.9) y LibreOffice para recalcular/convertir.
> No instales nada por adelantado: importa y, solo si falla, `pip install -r requirements.txt`.

## 1. Elige la herramienta

| Tarea | Herramienta |
|---|---|
| Ver la estructura de un libro antes de tocarlo | `python scripts/inspect_xlsx.py libro.xlsx` (añade `--formulas`, `--sheet`, `--json`) |
| Crear un libro con fórmulas y formato | `openpyxl` + `scripts/xlsx_helpers.py` |
| Crear un libro tabular sencillo de forma declarativa | `python scripts/build_from_spec.py spec.json out.xlsx` (ejemplo en `assets/spec_example.json`) |
| Leer/transformar muchos datos | `pandas` (`read_excel`, `to_excel`) → luego estilos con openpyxl |
| Editar un libro existente "normal" | `openpyxl.load_workbook` (+ `keep_vba=True` si es .xlsm) |
| Editar un libro con slicers, formas, controles, sparklines, Power Query… | `python scripts/patch_cells.py` (edición XML quirúrgica; no toca nada más) |
| Insertar filas/columnas sin romper fórmulas | `xlsx_helpers.insert_rows_safe` / `insert_cols_safe` (¡nunca `ws.insert_rows` a secas!) |
| Calcular valores y detectar errores | `python scripts/recalc.py libro.xlsx` |
| Detectar fórmulas incorrectas que no dan error | `python scripts/audit_formulas.py libro.xlsx` |
| Demostrar qué cambió tras una edición | `python scripts/diff_xlsx.py antes.xlsx despues.xlsx [--values]` |
| CSV/XLS/ODS/PDF ↔ XLSX | `python scripts/convert.py origen destino` |
| Ficheros enormes (>100k filas) | ver `references/rendimiento.md` (read_only / write_only / xlsxwriter) |

## 2. Flujo de trabajo obligatorio

1. **Inspecciona** (si hay fichero de entrada): `inspect_xlsx.py`. Fíjate en hojas ocultas, celdas
   combinadas, tablas, validaciones, colores de las celdas de entrada, macros, slicers y
   vínculos externos. Decide con eso si usar openpyxl o `patch_cells.py`
   (árbol de decisión en `references/edicion-segura.md`).
2. **Planifica** dónde va cada cosa: hojas, celdas de supuestos, rangos de datos, totales.
   En libros nuevos: una hoja de supuestos/entradas separada de los cálculos.
3. **Construye** con fórmulas, no con resultados calculados en Python.
   Escribe primero 2-3 fórmulas, recalcula y comprueba que apuntan a lo que crees; después
   replica el patrón.
4. **Recalcula y valida**: `recalc.py` → 0 errores. Luego `audit_formulas.py` y revisa los
   hallazgos `inconsistent` y `empty_ref` (los demás son orientativos).
5. **Verifica**: al editar, `diff_xlsx.py` original vs resultado — solo deben cambiar las
   celdas previstas. Si puedes, renderiza (`convert.py libro.xlsx out.pdf`) y mira el resultado.
6. **Entrega** describiendo en 2-3 líneas qué contiene, supuestos usados y cualquier
   limitación. No digas "listo" si recalc reporta errores.

## 3. Reglas de calidad (siempre)

- **Fórmulas, no valores fijos.** `ws["B10"] = "=SUM(B2:B9)"`, nunca el total calculado en Python.
  Excepción: datos de origen (ventas reales, importaciones) son valores.
- **Cero errores** (`#REF!`, `#DIV/0!`, `#VALUE!`, `#N/A`, `#NAME?`). Si crees que un error ya
  existía, demuéstralo leyendo el ORIGINAL con `data_only=True`.
- **Cada supuesto en su celda**, con etiqueta, referenciado con `$` (`=B5*(1+Supuestos!$B$2)`),
  nunca `=B5*1.05`. Documenta su origen (comentario de celda o columna "Fuente").
- **Respeta la especificación literal del usuario**: nombres exactos de hojas y columnas, la
  fórmula que pidió. No "mejores" lo que no te pidieron.
- **Editando un libro existente, sus convenciones mandan** sobre esta guía: mismo formato
  numérico, fuentes, colores, estructura. Escribe solo en las celdas de entrada designadas y no
  toques fórmulas existentes salvo que te lo pidan.
- **Formato profesional por defecto**: fuente única (Arial/Calibri), cabeceras destacadas,
  formatos numéricos coherentes, anchos ajustados, paneles inmovilizados, sin líneas de
  cuadrícula molestas en informes. Detalle en `references/formato-y-diseno.md`.
- **Plantillas para rellenar**: incluye leyenda de qué celdas editar y una fila de ejemplo
  (solo en libros que creas tú, nunca en uno que te pidieron editar).

## 4. Trampas críticas de openpyxl (memorízalas)

- openpyxl **no calcula**: guarda fórmulas sin valor. `data_only=True` o pandas leerán `None`
  hasta ejecutar `recalc.py`.
- `load_workbook(data_only=True)` **elimina las fórmulas**: si guardas ese objeto, las
  conviertes en valores para siempre. Para leer un modelo carga DOS veces (fórmulas y valores).
- `.xlsm` pierde las macros sin `keep_vba=True` (y debe guardarse con extensión .xlsm).
- `ws.insert_rows/insert_cols/delete_*` **no actualizan** referencias en fórmulas, rangos
  combinados, tablas, validaciones ni formato condicional → usa `insert_rows_safe`.
- Celdas combinadas: escribe solo en la celda superior izquierda.
- Hojas con espacios o caracteres especiales van entre comillas en fórmulas: `='Mis Datos'!A1`.
- Nombres de hoja ≤31 caracteres, sin `[]:*?/\`.
- openpyxl **pierde al guardar**: slicers, formas/cuadros de texto, controles de formulario,
  sparklines, extensiones de formato condicional, conexiones de Power Query; las tablas
  dinámicas se conservan pero no se actualizan. → `patch_cells.py`.
- Fechas: escribe objetos `datetime`/`date`, nunca cadenas, y aplica `number_format`.
- Porcentajes como fracción (`0.15` con formato `0.0%`), no `15`.

## 5. Fórmulas compatibles

Detalle completo y alternativas en `references/formulas.md`.

- Prefiere funciones clásicas: `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `INDEX`+`MATCH`, `IFERROR`,
  `SUMPRODUCT`, `VLOOKUP` (con `FALSE`), `EOMONTH`, `TEXT`, `ROUND`.
- Funciones posteriores a 2007 llevan prefijo `_xlfn.` en el XML: `_xlfn.TEXTJOIN`,
  `_xlfn.CONCAT`, `_xlfn.IFS`, `_xlfn.SWITCH`, `_xlfn.MAXIFS`, `_xlfn.MINIFS`, `_xlfn.IFNA`,
  `_xlfn.STDEV.S`… Sin prefijo → `#NAME?`.
- **Evita funciones de matriz dinámica** (`XLOOKUP`, `FILTER`, `UNIQUE`, `SORT`, `SEQUENCE`,
  `LET`, `LAMBDA`…): sin metadatos de desbordamiento solo calculan la primera celda y no se
  pueden verificar con LibreOffice. Ordena/filtra/deduplica en Python y escribe el resultado,
  o usa `INDEX/MATCH`. Úsalas solo si el usuario lo pide expresamente (y adviértelo).
- Escribe las fórmulas **en inglés y con comas** como separador (`=SUM(A1:A3)`, no
  `=SUMA(A1;A3)`): es el formato interno del fichero; Excel las muestra traducidas al usuario.
- Protege divisiones: `=IFERROR(A2/B2,0)` o `=IF(B2=0,"",A2/B2)`.

## 6. Recalcular: `recalc.py`

```bash
python scripts/recalc.py salida.xlsx                # recalcula y sobrescribe
python scripts/recalc.py entrada.xlsx --check-only  # solo valida (no modifica)
python scripts/recalc.py in.xlsx -o out.xlsx --timeout 120
```

Devuelve JSON: `status` (`success`|`errors_found`), `total_formulas`, `total_errors`,
`error_summary` (celdas por tipo de error), `empty_results` y `warnings` (análisis previo:
funciones sin `_xlfn.`, matrices dinámicas, hojas inexistentes, vínculos externos).
`errors_found` sale con código 0: **lee siempre el JSON**. Una clave `error` significa que no
se recalculó nada. Corrige y repite hasta `success`.

- Recalcular reescribe el fichero con LibreOffice. Para libros con elementos complejos
  (tablas dinámicas, slicers, gráficos muy formateados) usa `--check-only` y entrega el
  fichero sin reescribir; `patch_cells.py` ya marca el libro para que Excel recalcule al abrirlo.
  En libros creados con openpyxl puedes forzar lo mismo con `wb.calculation.fullCalcOnLoad = True`.
- Si hay **vínculos externos** (`='[1]Hoja'!A1`), recalc se niega: sin el fichero externo los
  valores se perderían. Copia antes esos valores o usa `--force` asumiendo la pérdida.
- `success` significa que las fórmulas evalúan, **no que sean correctas** → `audit_formulas.py`.

## 7. Recetas rápidas

```python
import sys; sys.path.insert(0, "scripts")          # ruta del skill
from openpyxl import Workbook, load_workbook
from xlsx_helpers import *

# Crear informe a partir de un DataFrame
wb = Workbook(); apply_default_font(wb)
ws = wb.active; ws.title = "Ventas"
ref = write_dataframe(ws, df, table_name="TVentas",
                      number_formats={"Importe": "eur2", "Fecha": "date", "Margen": "pct"})
n = len(df) + 1
ws[f"A{n+1}"] = "Total"; ws[f"C{n+1}"] = f"=SUM(C2:C{n})"
add_bar_chart(ws, f"C1:C{n}", f"A2:A{n}", "H2", title="Importe por mes")
wb.save("ventas.xlsx")          # → luego: python scripts/recalc.py ventas.xlsx

# Editar un libro existente conservando macros y formato
wb = load_workbook("modelo.xlsm", keep_vba=True)
ws = wb["Supuestos"]; ws["C5"] = 0.035          # solo celdas de entrada
insert_rows_safe(wb, wb["Datos"], idx=20, amount=3, copy_style_from=19)
wb.save("modelo.xlsm")

# Leer un modelo: fórmulas y valores
wf = load_workbook("m.xlsx"); wv = load_workbook("m.xlsx", data_only=True)
```

```bash
# Edición quirúrgica (libro con slicers/formas): solo cambian esas celdas
python scripts/patch_cells.py informe.xlsx -o informe.xlsx \
  --set "Supuestos!B3=0.045" --set "Resumen!D10==SUM(D2:D9)" --set "Resumen!A12=Revisado"
```

## 8. Datos desordenados de entrada

Cabeceras que no están en la fila 1, filas de títulos/notas, totales intercalados, celdas
combinadas, números como texto (`"1.234,56"`), fechas como texto o como serial (`45292`),
columnas vacías. Procedimiento y código en `references/pandas-datos.md`: localizar la
cabecera, normalizar nombres, convertir tipos con formato español, eliminar filas de
totales, validar recuentos y sumas contra el original.

## 9. Referencias (carga solo la que necesites)

| Fichero | Cuándo |
|---|---|
| `references/openpyxl-recetas.md` | API concreta: estilos, combinadas, comentarios, hipervínculos, imágenes, protección, impresión, nombres definidos, hojas ocultas |
| `references/formulas.md` | Compatibilidad de funciones, `_xlfn`, fórmulas matriciales, referencias estructuradas, patrones de búsqueda y agregación |
| `references/formato-y-diseno.md` | Estándares visuales, formatos numéricos (incl. €/España), formato condicional, validación |
| `references/graficos.md` | Gráficos con openpyxl: barras, líneas, tarta, dispersión, combinados, ejes secundarios |
| `references/pandas-datos.md` | Lectura masiva, limpieza de datos sucios, varias hojas, añadir hojas a un libro existente |
| `references/edicion-segura.md` | Qué conserva/pierde cada herramienta, árbol de decisión, inserción de filas, xlsm, plantillas |
| `references/modelos-financieros.md` | Convenciones de color, estructura y formato para modelos financieros |
| `references/rendimiento.md` | Ficheros grandes, modos read_only/write_only, xlsxwriter, límites de Excel |
| `references/problemas-frecuentes.md` | Diagnóstico: "Excel encontró contenido ilegible", `#NAME?`, fechas raras, etc. |
