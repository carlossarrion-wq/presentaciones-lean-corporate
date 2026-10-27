# Editar libros existentes sin romperlos

## Qué conserva cada herramienta

| Elemento | openpyxl (load+save) | pandas `mode="a"` | LibreOffice (recalc/convert) | `patch_cells.py` |
|---|---|---|---|---|
| Valores, fórmulas, estilos, combinadas | ✅ | ✅ (vía openpyxl) | ✅ | ✅ |
| Validación de datos, formato condicional básico | ✅ | ✅ | ✅ | ✅ |
| Formato condicional avanzado (extensiones x14, barras con relleno sólido…) | ❌ se elimina | ❌ | ⚠️ | ✅ |
| Tablas de Excel | ✅ | ✅ | ✅ | ✅ |
| Gráficos | ⚠️ se conservan pero puede simplificarse su formato | ⚠️ | ⚠️ | ✅ |
| Imágenes | ✅ (con Pillow) | ✅ | ✅ | ✅ |
| Formas, cuadros de texto, SmartArt | ❌ | ❌ | ⚠️ | ✅ |
| Slicers / segmentaciones, escalas de tiempo | ❌ | ❌ | ❌ | ✅ |
| Tablas dinámicas | ⚠️ se conservan, no se actualizan | ⚠️ | ⚠️ | ✅ |
| Macros VBA (.xlsm) | ✅ solo con `keep_vba=True` | ❌ | ⚠️ | ✅ |
| Controles de formulario / ActiveX | ❌ | ❌ | ⚠️ | ✅ |
| Sparklines (minigráficos) | ❌ | ❌ | ⚠️ | ✅ |
| Power Query / conexiones / modelo de datos | ❌ | ❌ | ❌ | ✅ |
| Valores cacheados de fórmulas | ❌ se borran | ❌ | recalcula | ✅ (marca recálculo al abrir) |

(⚠️ = suele funcionar, pero verifica.) `inspect_xlsx.py` avisa de macros, tablas dinámicas,
slicers y vínculos externos.

## Árbol de decisión

1. ¿El libro tiene slicers, formas, controles, sparklines, Power Query o gráficos muy
   cuidados, y solo hay que cambiar valores/fórmulas de celdas concretas?
   → **`patch_cells.py`**. Valida con `recalc.py --check-only` (no reescribe el original).
2. ¿Hay que hacer cambios estructurales (añadir hojas, insertar filas, dar formato, crear
   gráficos) en un libro "normal"? → **openpyxl** (+ `keep_vba=True` si .xlsm).
3. ¿Ambas cosas en un libro con elementos frágiles? → informa al usuario del riesgo y
   propón: (a) hacer los cambios estructurales en una hoja/libro nuevo enlazado, o (b) que
   los aplique él en Excel siguiendo tus instrucciones, o (c) aceptar la pérdida documentada.
4. Siempre: **trabaja sobre una copia**, conserva el original y ejecuta
   `diff_xlsx.py original copia` al final.

## Insertar y eliminar filas/columnas

`ws.insert_rows()` de openpyxl desplaza las celdas pero NO actualiza fórmulas (ni de esa hoja
ni de otras), rangos combinados, tablas, validaciones, formato condicional ni nombres
definidos. Resultado: referencias silenciosamente desplazadas.

```python
from xlsx_helpers import insert_rows_safe, insert_cols_safe
insert_rows_safe(wb, ws, idx=15, amount=3, copy_style_from=14)  # amplía una tabla copiando formato y fórmulas de la fila 14
insert_cols_safe(wb, ws, idx=4, amount=1)
```
Limitaciones: no ajusta los rangos de gráficos ni el área de impresión (revísalos), y no hay
`delete_rows_safe`: para eliminar filas, mejor vaciar su contenido o reconstruir el bloque;
si es imprescindible, comprueba después con `recalc.py` que no aparecen `#REF!`.

Para **añadir filas al final de una tabla de datos** basta con escribir debajo y ampliar:
`ws.tables["TVentas"].ref = "A1:F151"` (+ ampliar rangos de fórmulas de totales si no usan la
columna completa).

## Libros con macros (.xlsm)

```python
wb = load_workbook("m.xlsm", keep_vba=True); ...; wb.save("m.xlsm")
```
- Nunca guardes un .xlsm como .xlsx (Excel se niega a abrirlo) ni al revés.
- No puedes ejecutar macros; si el resultado depende de ellas, dilo.
- Nunca escribas ni modifiques código VBA sin petición expresa del usuario.

## Rellenar plantillas corporativas

1. `inspect_xlsx.py --formulas` → identifica celdas de entrada (color de fuente/relleno,
   desbloqueadas, celdas vacías referenciadas por fórmulas) y celdas calculadas.
2. Escribe solo en entradas. Respeta formatos ya aplicados (no cambies `number_format`).
3. Si faltan filas, usa `insert_rows_safe(..., copy_style_from=)`.
4. No borres filas de ejemplo de la plantilla salvo que se indique.
5. `diff_xlsx.py` para demostrar que las fórmulas no han cambiado.

## Encontrar algo en un libro grande

```python
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and "IVA" in c.value: print(ws.title, c.coordinate, c.value)
```
Para localizar la columna por cabecera en lugar de suponer letras:
```python
hdr = {str(c.value).strip(): c.column for c in ws[1] if c.value}
col_importe = hdr["Importe"]
```

## Valores de vínculos externos

Fórmulas como `='[1]Hoja'!B2` dependen de otro fichero. Al guardar con openpyxl se pierde el
valor cacheado; al recalcular sin el fichero se obtiene `#REF!`/`#NAME?`. Antes de editar:
lee esos valores con `data_only=True` del original y, si el usuario lo acepta, sustitúyelos
por valores (documentándolo), o edita con `patch_cells.py`, que no toca esas celdas.
