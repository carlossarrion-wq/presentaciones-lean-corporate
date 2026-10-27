# Problemas frecuentes y diagnóstico

| Síntoma | Causa probable | Solución |
|---|---|---|
| Las celdas con fórmula salen vacías/`None`/NaN al leer | Libro escrito por openpyxl sin valores cacheados | `recalc.py` y volver a leer |
| Todas las fórmulas se convirtieron en números | Se guardó un libro abierto con `data_only=True` | Rehacer desde el original; cargar sin `data_only` para editar |
| "Excel encontró contenido que no se puede leer" | Nombre de tabla duplicado o inválido; cabecera de tabla con celdas vacías/duplicadas; validación >255 caracteres; nombre de hoja inválido; XML manipulado mal; fórmula con sintaxis no válida en el XML | Revisar tablas (`inspect_xlsx.py`), nombres, validaciones; regenerar |
| `#NAME?` | Función en español, sin `_xlfn.`, mal escrita, nombre definido inexistente, matriz dinámica | Ver formulas.md; `recalc.py` muestra avisos |
| `#VALUE!` en referencia a otra hoja | Nombre de hoja con espacios sin comillas | `='Mi Hoja'!A1` |
| Referencias desplazadas tras insertar filas | `ws.insert_rows` no ajusta fórmulas | `insert_rows_safe`; revisar con `diff_xlsx.py` |
| Macros desaparecidas | Faltó `keep_vba=True` o se guardó como .xlsx | Rehacer desde el original |
| Fechas como números (45292) | Falta formato de fecha | `number_format = "dd/mm/yyyy"` |
| Fechas como texto que no ordenan | Se escribieron cadenas | Escribir `datetime`/`date` |
| `1.234,56` tratado como texto | CSV/entrada con formato español leído como anglosajón | `decimal=",", thousands="."` o `num_es()` (pandas-datos.md) |
| Ceros a la izquierda perdidos (CP 08001 → 8001) | Columna leída/escrita como número | `dtype=str` al leer y `number_format="@"` al escribir |
| Tildes rotas al abrir un CSV en Excel | CSV en UTF-8 sin BOM | Guardar con `encoding="utf-8-sig"` (convert.py lo hace) |
| Excel abre el CSV en una sola columna | Excel español espera `;` | `convert.py ... --sep ";"` o entregar .xlsx |
| Porcentajes 1500 % | Se guardó 15 en vez de 0.15 | Guardar fracción |
| Año 2026 mostrado como 2.026 | Formato `#,##0` en años | Formato `0` o texto |
| Gráfico vacío en el visor | Datos son fórmulas sin calcular | `recalc.py` (Excel recalcula al abrir) |
| Ejes del gráfico invisibles | openpyxl ≥3.1 marca ejes como eliminados en algunos casos | `chart.x_axis.delete = False; chart.y_axis.delete = False` (los helpers ya lo hacen) |
| Cabecera de tabla blanca sobre blanco en LibreOffice/Google | No pintan estilos de tabla | Estilo explícito (`style_header_row`) |
| `MergedCell ... read-only` | Escritura en celda no ancla de un rango combinado | Escribir en la superior izquierda |
| `ValueError: Cannot convert ... to Excel` | Tipo no soportado (numpy, Decimal, tz-aware datetime, listas) | `float()`, `int()`, `dt.replace(tzinfo=None)`, `str()` |
| `IllegalCharacterError` | Caracteres de control en texto | `re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", s)` |
| LibreOffice cuelga o no arranca | Instancia previa con el perfil bloqueado; sin entorno gráfico | Los scripts usan perfil temporal y `SAL_USE_VCLPLUGIN=svp`; mata procesos `soffice` colgados; aumenta `--timeout` |
| `LibreOffice no encontrado` | No instalado | `apt-get install -y libreoffice-calc` / `brew install --cask libreoffice`, o `SOFFICE_PATH` |
| Libro muy lento de abrir | Formato aplicado a columnas completas, miles de estilos únicos, volátiles (`INDIRECT`, `OFFSET`, `TODAY`) | Ver rendimiento.md |
| `zipfile.BadZipFile` | El fichero no es .xlsx real (un .xls renombrado, HTML exportado, CSV) | Comprobar con `file f.xlsx`; `convert.py` |
| Contraseña de apertura | openpyxl no descifra | Pedir copia sin contraseña al usuario |

## Verificación final (checklist)

- [ ] `recalc.py` → `success`, `total_errors: 0`, sin `warnings` relevantes
- [ ] `audit_formulas.py` → sin `inconsistent` ni `empty_ref` injustificados
- [ ] Al editar: `diff_xlsx.py` muestra solo los cambios previstos
- [ ] Nombres de hojas/columnas exactamente como pidió el usuario
- [ ] Formatos numéricos coherentes, fechas como fechas, % como fracción
- [ ] Supuestos documentados (fuente) y visibles
- [ ] Extensión correcta (.xlsm si tiene macros)
