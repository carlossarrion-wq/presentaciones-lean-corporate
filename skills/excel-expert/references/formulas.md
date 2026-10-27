# Fórmulas: compatibilidad y patrones

## Cómo se guardan

- El fichero almacena las fórmulas **en inglés, con coma como separador de argumentos y punto
  decimal**: `=SUMIFS(C:C,A:A,"Norte",B:B,">=1000")`. Excel en español las muestra como
  `=SUMAR.SI.CONJUNTO(C:C;A:A;"Norte";...)`. Nunca escribas nombres en español ni `;`.
- openpyxl escribe la cadena tal cual en el XML; no valida nada. Un error tipográfico solo
  aparece al recalcular (`#NAME?`). LibreOffice devuelve en minúsculas las fórmulas que no
  pudo interpretar: pista útil junto a un `#NAME?`.
- Literales de texto con comillas dobles; comillas dentro del texto se duplican: `="Dijo ""hola"""`.

## Prefijo `_xlfn.` (funciones posteriores a Excel 2007)

En el XML, las funciones nuevas llevan prefijo (la interfaz de Excel lo oculta). Sin él → `#NAME?`.

| Escribe | Notas |
|---|---|
| `_xlfn.TEXTJOIN(", ",TRUE,A2:A9)` | |
| `_xlfn.CONCAT(A2,B2)` | o `A2&B2` / `CONCATENATE` (sin prefijo) |
| `_xlfn.IFS(A2>90,"A",A2>70,"B",TRUE,"C")` | alternativa: IF anidados |
| `_xlfn.SWITCH(A2,1,"Uno",2,"Dos","Otro")` | |
| `_xlfn.MAXIFS(C:C,A:A,"Norte")`, `_xlfn.MINIFS(...)` | |
| `_xlfn.IFNA(VLOOKUP(...),"")` | o `IFERROR` (sin prefijo) |
| `_xlfn.STDEV.S`, `_xlfn.VAR.S`, `_xlfn.PERCENTILE.INC`, `_xlfn.RANK.EQ`, `_xlfn.NORM.DIST`… | estadísticas "con punto" |
| `_xlfn.DAYS`, `_xlfn.ISOWEEKNUM`, `_xlfn.CEILING.MATH`, `_xlfn.FLOOR.MATH`, `_xlfn.AGGREGATE` | |

`recalc.py` avisa si detecta alguna de estas funciones sin prefijo.

## Evitar: funciones de matriz dinámica (Excel 365)

`XLOOKUP`, `XMATCH`, `FILTER`, `UNIQUE`, `SORT`, `SORTBY`, `SEQUENCE`, `RANDARRAY`, `LET`,
`LAMBDA`, `VSTACK`, `HSTACK`, `TAKE`, `DROP`, `TEXTSPLIT`, `TEXTBEFORE`, `TEXTAFTER`…

Motivo: necesitan metadatos de "desbordamiento" que openpyxl no escribe; sin ellos solo se
calcula la celda superior izquierda, y LibreOffice no puede validarlas. Alternativas:

| En lugar de | Usa |
|---|---|
| `XLOOKUP(k,A:A,C:C)` | `INDEX(C:C,MATCH(k,A:A,0))` |
| `XLOOKUP` con valor por defecto | `IFERROR(INDEX(C:C,MATCH(k,A:A,0)),"No encontrado")` |
| `XLOOKUP` de dos criterios | `INDEX(C2:C99,MATCH(1,INDEX((A2:A99=k1)*(B2:B99=k2),0),0))` |
| `UNIQUE`, `SORT`, `FILTER` | calcúlalo en pandas y escribe los valores; o tabla dinámica |
| `SEQUENCE(12)` | escribe 1..12 en Python |
| `LET` | celdas auxiliares con nombre |

Si el usuario pide explícitamente una de ellas: escríbela con prefijo (`_xlfn.XLOOKUP`,
`_xlfn._xlws.FILTER`, `_xlfn._xlws.SORT`), avisa de que solo se calculará en Excel 365 y
de que no se ha podido verificar.

## Fórmulas matriciales clásicas (CSE)

```python
from openpyxl.worksheet.formula import ArrayFormula
ws["E2"] = ArrayFormula("E2", "=SUM(B2:B10*C2:C10)")
```
Normalmente `SUMPRODUCT(B2:B10,C2:C10)` evita la necesidad de matriz.

## Referencias

- Relativa `A1`, absoluta `$A$1`, mixta `$A1`/`A$1`. Supuestos siempre absolutos.
- Otra hoja: `Supuestos!B2`; con espacios/acentos/guiones: `'Mis Supuestos'!B2`
  (apóstrofo dentro del nombre se duplica: `'Plan d''acció'!A1`).
- Otro fichero: `'[Datos.xlsx]Hoja1'!A1` — evítalo; los vínculos externos son frágiles.
- Columnas completas (`A:A`) son cómodas en SUMIFS/MATCH; evítalas en SUMPRODUCT y fórmulas
  matriciales (lentas).
- Referencias estructuradas a Tablas (`=SUM(TVentas[Importe])`): Excel las entiende, pero
  openpyxl las escribe como texto y LibreOffice puede no resolverlas al recalcular. Prefiere
  rangos normales salvo petición expresa.
- Copiar una fórmula ajustando referencias:
```python
from openpyxl.formula.translate import Translator
ws["D3"] = Translator(ws["D2"].value, origin="D2").translate_formula("D3")
```

## Patrones habituales

| Necesidad | Fórmula |
|---|---|
| Suma condicionada | `=SUMIFS(Datos!$D:$D,Datos!$B:$B,$A2,Datos!$C:$C,">="&$F$1)` |
| Recuento con varios criterios | `=COUNTIFS(B:B,"Norte",D:D,">0")` |
| Búsqueda exacta | `=IFERROR(INDEX(Tarifas!$C:$C,MATCH($A2,Tarifas!$A:$A,0)),"")` |
| Búsqueda por tramos (aprox.) | `=INDEX($C$2:$C$6,MATCH(B2,$A$2:$A$6,1))` (tabla ordenada asc.) |
| % sobre total | `=IFERROR(B2/B$20,0)` |
| Variación | `=IFERROR(C2/B2-1,"")` |
| CAGR | `=(F2/B2)^(1/4)-1` |
| Fin de mes / meses | `=EOMONTH(A2,0)`, `=EDATE(A2,1)` |
| Días laborables | `=NETWORKDAYS(A2,B2,Festivos!$A$2:$A$20)` |
| Año/mes de una fecha | `=YEAR(A2)`, `=TEXT(A2,"yyyy-mm")` (el código de formato en TEXT depende del idioma de Excel: evita TEXT para lógica, úsalo solo para mostrar) |
| Redondeo financiero | `=ROUND(B2*C2,2)` |
| Promedio ponderado | `=SUMPRODUCT(B2:B10,C2:C10)/SUM(C2:C10)` |
| Total que ignora filas filtradas | `=SUBTOTAL(109,D2:D500)` |
| Última fila no vacía | `=LOOKUP(2,1/(A:A<>""),A:A)` |
| Concatenar con separador (2007-compatible) | `=A2&" - "&B2` |

## Errores y su causa típica

| Error | Causa | Arreglo |
|---|---|---|
| `#NAME?` | función mal escrita, en español, sin `_xlfn.`, nombre definido inexistente, hoja con espacios sin comillas (a veces `#VALUE!`) | revisar la cadena |
| `#REF!` | referencia a fila/columna/hoja eliminada, desplazamiento fuera de la hoja | reconstruir la referencia |
| `#DIV/0!` | divisor 0 o vacío | `IFERROR`/`IF` |
| `#VALUE!` | texto donde se espera número (p.ej. `"1.234,56"` importado como texto) | convertir datos de origen |
| `#N/A` | MATCH/VLOOKUP sin coincidencia; tipos distintos (número vs texto `"123"`) | normalizar tipos, `IFERROR` |
| `#NUM!` | cálculo imposible (IRR sin convergencia, raíz de negativo) | revisar datos |
| Circular | rango que incluye la propia celda | `audit_formulas.py` lo marca como `self_range` |
