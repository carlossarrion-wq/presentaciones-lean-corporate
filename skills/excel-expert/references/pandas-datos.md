# Datos con pandas: lectura, limpieza y escritura

## Leer

```python
import pandas as pd
df = pd.read_excel("f.xlsx", sheet_name="Ventas")            # una hoja
todas = pd.read_excel("f.xlsx", sheet_name=None)              # dict {hoja: df}
df = pd.read_excel("f.xlsx", header=3, usecols="B:H", nrows=500, dtype={"CP": str})
df = pd.read_csv("f.csv", sep=";", decimal=",", thousands=".", encoding="cp1252")  # CSV español
```
- `pd.read_excel` lee **valores cacheados**: libros generados por openpyxl sin recalcular dan NaN
  en las fórmulas → `recalc.py` antes.
- `.xls` necesita `xlrd`; `.xlsb` necesita `pyxlsb`; `.ods` necesita `odfpy`. Alternativa:
  `convert.py` a .xlsx.
- Códigos con ceros a la izquierda (CP, NIF, cuentas): `dtype=str`.

## Datos desordenados: procedimiento

1. **Mira el fichero crudo** (`inspect_xlsx.py --rows 30`) o `pd.read_excel(..., header=None).head(30)`.
2. **Localiza la cabecera**: la primera fila con la mayoría de celdas de texto no vacías y
   que contiene nombres esperados.
```python
raw = pd.read_excel(f, header=None)
hdr = next(i for i, r in raw.iterrows() if r.notna().sum() >= 0.6 * raw.shape[1] and r.astype(str).str.contains("Fecha|Importe", case=False).any())
df = pd.read_excel(f, header=hdr)
```
3. **Normaliza nombres de columna**: `df.columns = df.columns.astype(str).str.strip().str.replace(r"\s+", " ", regex=True)`.
   Elimina columnas `Unnamed:` totalmente vacías: `df = df.dropna(axis=1, how="all")`.
4. **Elimina filas basura**: vacías (`dropna(how="all")`), títulos repetidos (filas iguales
   a la cabecera), totales/subtotales (`~df.iloc[:,0].astype(str).str.contains("total", case=False)`),
   notas al pie.
5. **Rellena celdas combinadas** verticales (categoría escrita solo en la primera fila):
   `df["Región"] = df["Región"].ffill()`.
6. **Convierte tipos**:
```python
def num_es(s):  # "1.234,56 €" → 1234.56 ; "(1.000)" → -1000
    s = s.astype(str).str.strip().str.replace(r"[€$\s%]", "", regex=True)
    neg = s.str.match(r"^\(.*\)$"); s = s.str.strip("()")
    s = s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False)
    v = pd.to_numeric(s, errors="coerce"); return v.where(~neg, -v)
df["Importe"] = num_es(df["Importe"])
df["Fecha"] = pd.to_datetime(df["Fecha"], dayfirst=True, errors="coerce")
# Serial de Excel (45292 → 2024-01-01):
df["F"] = pd.to_datetime(df["F"], unit="D", origin="1899-12-30")
```
   Ojo: si el texto mezcla formato `1,234.56` (anglosajón) y `1.234,56`, detecta por fila.
7. **Valida contra el original**: nº de filas esperado, suma de importes antes/después,
   valores no convertidos (`df["Importe"].isna().sum()` frente a los vacíos originales).
   Informa al usuario de las filas descartadas y por qué.
8. **Deduplica** con criterio explícito (`df.duplicated(subset=[...])`) y reporta cuántas.

## Escribir

```python
with pd.ExcelWriter("out.xlsx", engine="openpyxl") as xw:
    df.to_excel(xw, sheet_name="Datos", index=False)
    resumen.to_excel(xw, sheet_name="Resumen")
```
Después aplica formato con openpyxl (o usa directamente `xlsx_helpers.write_dataframe`, que
escribe y formatea en un paso y permite añadir fórmulas después).

**Añadir/reemplazar una hoja en un libro existente** (conservando las demás):
```python
with pd.ExcelWriter("existente.xlsx", engine="openpyxl", mode="a", if_sheet_exists="replace") as xw:
    df.to_excel(xw, sheet_name="Datos", index=False)
```
`mode="a"` carga el libro con openpyxl (aplican sus limitaciones: ver edicion-segura.md).
`if_sheet_exists="overlay"` escribe encima sin borrar la hoja (útil para pegar datos en una
plantilla a partir de `startrow`/`startcol`).

## Agregaciones típicas antes de escribir

```python
pv = df.pivot_table(index="Región", columns=df["Fecha"].dt.to_period("M"), values="Importe", aggfunc="sum", fill_value=0)
pv.columns = pv.columns.astype(str)
top = df.groupby("Cliente", as_index=False)["Importe"].sum().nlargest(10, "Importe")
```
Si el usuario quiere que el resumen **se actualice** al cambiar los datos, no escribas el
resultado de pandas: escribe fórmulas `SUMIFS` que apunten a la hoja de datos.

## Tablas dinámicas reales

openpyxl no crea tablas dinámicas. Opciones: (a) resumen con fórmulas `SUMIFS` (recomendado,
siempre vivo), (b) resumen estático de pandas indicando que es una foto, (c) si se exige una
pivot real, explica que debe crearse en Excel (Insertar → Tabla dinámica) sobre la Tabla de
datos que entregas.
