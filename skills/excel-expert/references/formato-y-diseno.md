# Formato y diseño profesional

## Principios

1. **Una fuente** en todo el libro (Arial 10 o Calibri 11) salvo que el libro existente use otra.
   `apply_default_font(wb)` fija la fuente por defecto.
2. **Jerarquía clara**: título (negrita 14) → cabeceras (negrita, relleno oscuro, texto blanco,
   centradas, con ajuste de texto) → datos → totales (negrita, borde superior fino y doble inferior).
3. **Números alineados a la derecha, texto a la izquierda**, formato coherente por columna.
4. **Anchos ajustados** (`autofit_columns`) y **paneles inmovilizados** bajo la cabecera.
5. **Sin ruido**: sin bordes en todas las celdas, como mucho bandas suaves o líneas finas
   grises; colores con significado (no decorativos). En hojas tipo informe oculta la cuadrícula.
6. **Leyenda** cuando los colores codifican algo (entradas, supuestos, alertas).
7. Unidades en la cabecera: `Importe (€)`, `Ventas (miles €)`, `Margen (%)`.

## Formatos numéricos (`xlsx_helpers.NUMBER_FORMATS`)

| Clave | Código | Muestra |
|---|---|---|
| `int` | `#,##0` | 1.234 |
| `dec2` | `#,##0.00` | 1.234,56 |
| `eur` | `#,##0 "€";-#,##0 "€";"-"` | 1.235 € / – |
| `eur2` | `#,##0.00 "€";-#,##0.00 "€";"-"` | 1.234,56 € |
| `usd` | `$#,##0;($#,##0);"-"` | $1,235 / ($1,235) |
| `pct` | `0.0%` | 15,0 % |
| `multiple` | `0.0"x"` | 8,5x |
| `date` | `dd/mm/yyyy` | 31/03/2026 |
| `month` | `mmm-yy` | mar-26 |
| `thousands` | `#,##0,"k"` | 1.235k |
| `millions` | `#,##0.0,,"M"` | 1,2M |
| `text` | `@` | conserva ceros iniciales |

- En el código se usan siempre `,` para miles y `.` para decimales; **Excel los muestra según
  la configuración regional del usuario** (en España: `1.234,56`). No "traduzcas" el código.
- Tres secciones `positivo;negativo;cero` permiten mostrar ceros como `-` y negativos entre
  paréntesis o en rojo: `#,##0;[Red]-#,##0;"-"`.
- Porcentajes guardados como fracción. Años como texto o número sin separador (`0`) para
  evitar `2.026`.
- Fechas: escribe `datetime.date`/`datetime`, nunca cadenas; aplica formato de fecha.

## Formato condicional (`xlsx_helpers`)

```python
add_color_scale(ws, "E2:E50")                         # rojo→amarillo→verde
add_data_bars(ws, "C2:C50")
highlight_cells(ws, "F2:F50", "lessThan", [0], "bad")   # negativos en rojo
highlight_cells(ws, "F2:F50", "between", [0, 0.05], "neutral")
highlight_formula(ws, "A2:H50", "$H2=\"Vencido\"", "bad")   # fila entera según una columna
```
- En reglas por fórmula, la referencia es relativa a la **celda superior izquierda** del rango
  y se fija la columna con `$` para colorear filas completas.
- Más avanzado (iconos): `from openpyxl.formatting.rule import IconSetRule`;
  `ws.conditional_formatting.add("D2:D50", IconSetRule("3Arrows", "percent", [0, 33, 67]))`.

## Validación de datos

```python
add_list_validation(ws, "D2:D500", ["Pendiente", "En curso", "Cerrado"])
add_list_validation(ws, "E2:E500", source_range="Listas!$A$2:$A$40")   # listas largas
add_number_validation(ws, "F2:F500", minimum=0, maximum=1)             # % entre 0 y 1
```
Para listas dependientes (provincia según comunidad) usa nombres definidos e
`INDIRECT(...)` en `formula1`. Las listas inline tienen un límite de 255 caracteres.

## Tablas de Excel

`add_table(ws, "A1:F200", "TVentas")` añade filtros, bandas y referencia estructurada.
Reglas: nombre único sin espacios; la cabecera no puede tener celdas vacías ni duplicadas;
no combines celdas dentro. Aplica igualmente estilo explícito a la cabecera
(`style_header_row`) porque LibreOffice y Google Sheets no pintan los estilos de tabla.

## Hoja de portada / índice (libros con muchas hojas)

Primera hoja "Índice" con título, fecha, autor, descripción de cada hoja con hipervínculo
(`add_hyperlink_to_sheet`) y leyenda de colores.

## Plantillas para que otros rellenen

- Celdas de entrada en azul (o con relleno suave) y desbloqueadas; el resto protegido si
  procede.
- Validaciones en todas las entradas con valores cerrados.
- Una fila de ejemplo con valores realistas y una nota "Sustituya la fila de ejemplo".
- Instrucciones breves arriba o en una hoja "Instrucciones".
