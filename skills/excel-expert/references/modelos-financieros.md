# Modelos financieros

Aplica salvo que el usuario o el libro existente usen otra convención.

## Código de colores (estándar de banca de inversión)

| Elemento | Formato |
|---|---|
| Datos introducidos a mano / palancas de escenario | fuente **azul** `0000FF` |
| Fórmulas y cálculos | fuente **negra** `000000` |
| Vínculos a otra hoja del mismo libro | fuente **verde** `008000` |
| Vínculos a otro fichero | fuente **roja** `FF0000` |
| Supuestos clave / celdas a revisar | relleno **amarillo** `FFFF00` |

`xlsx_helpers.mark_input(cell, key_assumption=True)` aplica azul + amarillo.

## Estructura

- Hojas en orden de flujo: `Portada` → `Supuestos` → cálculos (`Ingresos`, `Costes`,
  `Personal`, `Capex`…) → estados (`PyG`, `Balance`, `Flujo de caja`) → `Valoración` /
  `Escenarios` → `Resumen`.
- Periodos en columnas, conceptos en filas; **la misma fórmula en toda la fila** (una celda
  distinta a mitad de fila es el error silencioso más habitual → `audit_formulas.py`).
- Columna de etiquetas, columna de unidades, columna de notas/fuente.
- Cada supuesto en su celda con etiqueta y fuente; nunca constantes dentro de las fórmulas
  (`=C5*(1+$C$3)`, no `=C5*1.05`).
- Escenarios: una celda selectora (1/2/3 con validación) y `CHOOSE($C$2,base,optimista,pesimista)`
  o `INDEX` sobre una tabla de escenarios.
- Controles: el balance cuadra (`=ROUND(Activo-Pasivo-PN,2)=0`), la caja del flujo coincide
  con el balance; una fila "Check" con formato condicional rojo si falla.
- Protege divisiones (`IFERROR`), evita circularidades (intereses sobre caja media → usar
  saldo inicial o declarar explícitamente cálculo iterativo).

## Formato

| Dato | Formato |
|---|---|
| Moneda | `#,##0 "€";(#,##0 "€");"-"` (unidades en la cabecera: `Ingresos (miles €)`) |
| Porcentajes | `0.0%` (guardados como fracción) |
| Múltiplos | `0.0"x"` |
| Años | texto `"2026"` o número con formato `0` (nunca `2.026`) |
| Negativos | entre paréntesis |
| Ceros | `-` |

Históricos vs proyecciones: cabecera `2025A` / `2026E`, o una línea vertical separadora;
proyecciones con relleno suave opcional.

## Funciones útiles

`NPV` (ojo: descuenta desde el periodo 1; la inversión inicial va fuera), `IRR`, `XNPV`/`XIRR`
(fechas irregulares), `PMT`, `IPMT`, `PPMT`, `EOMONTH`, `SUMPRODUCT` para ponderaciones,
`INDEX/MATCH` para tablas de supuestos.

```
VAN  =NPV($C$4,D20:H20)+C20
TIR  =IRR(C20:H20)
Cuota préstamo =PMT($C$6/12,$C$7*12,-$C$5)
```

## Documentar fuentes

Comentario en la celda o columna "Fuente": `Fuente: Cuentas anuales 2025, Nota 18, p. 45`.
Si el dato viene del usuario: `Fuente: dato facilitado por el usuario (27/10/2026)`.
Si es una estimación propia: dilo explícitamente.
