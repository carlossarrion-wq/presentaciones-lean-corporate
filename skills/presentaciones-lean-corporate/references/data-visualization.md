# Visualización de datos en diapositivas

Un gráfico en una diapositiva no es un análisis exploratorio: es **la prueba de
un titular**. Debe entenderse en 5 segundos.

## 1. Elige el gráfico por la pregunta

| Quiero mostrar… | Gráfico (`chart.type`) | Notas |
|---|---|---|
| Ranking / comparación entre categorías | `bar` (horizontal) | Ordena de mayor a menor. Etiquetas largas → siempre horizontal |
| Comparación en pocas categorías o periodos (≤ 7) | `column` | |
| Evolución en el tiempo | `line` | Muchos periodos o varias series. Máx. 4 líneas |
| Evolución de pocos periodos | `column` | 3–6 periodos |
| Composición de un total (2–5 partes) | `doughnut` o `stacked-bar` | Tarta solo si las partes son muy distintas; nunca > 5 porciones |
| Composición que cambia en el tiempo | `stacked-column` / `stacked-column-100` | Máx. 4 segmentos |
| Una o dos cifras | — | No es un gráfico: usa `big-number` o `kpis` |
| Valores exactos para consulta | — | `table` |

Nunca: gráficos 3D, tartas con muchas porciones, doble eje sin necesidad,
radar para audiencias no técnicas.

## 2. Haz que el mensaje salte a la vista

- **Titular = conclusión** del gráfico ("Facturación concentra el 24 % de los
  contactos"), no su descripción ("Contactos por tipología").
- **Resalta lo que importa** con `highlight` (una barra en acento, el resto
  atenuado) o elige el color de la serie relevante en acento.
- Añade `insights` (2–3) o un `takeaway` con el "¿y qué?".
- Indica la unidad en `subtitle` ("Tiempo medio de espera, minutos") o en el formato ("0.0\" M€\"").

## 3. Menos tinta, más dato

El motor aplica por defecto: sin título interno, sin bordes, rejilla tenue o
ninguna, etiquetas de datos en lugar de eje cuando hay una sola serie, leyenda
solo si hay varias series, colores del tema. Además:

- **Ordena** las barras (salvo orden natural: tiempo, escalas).
- Empieza el eje de barras/columnas en 0. En líneas puedes acotar (`y_min`)
  si lo importante es la variación, pero dilo.
- Redondea: 2–3 cifras significativas bastan ("4,2 M", no "4.213.587").
- Máximo ~12 categorías; si hay más, agrupa en "Otros".
- Mismas escalas y colores para la misma métrica en todo el deck.
- Varias series: máximo 4; la importante en acento, el resto en neutros.

## 4. Tablas

- Para datos que se **consultan**, no para tendencias.
- Números alineados a la derecha (`col_align`), texto a la izquierda.
- Misma cantidad de decimales por columna; unidades en la cabecera ("M€").
- Resalta la fila clave (`highlight_row`) y usa `**…**` para totales.
- ≤ 8 filas × 6 columnas en el cuerpo; el resto, en anexo.

## 5. Formatos numéricos (`number_format`)

| Formato | Resultado (locale ES) |
|---|---|
| `0` | 12 |
| `0.0` | 12,4 |
| `#,##0` | 12.400 |
| `0%` (valor 0.24) | 24 % |
| `0.0%` | 24,3 % |
| `0.0" M€"` | 12,4 M€ |
| `+0%;-0%;0%` | +5 % / −3 % |

En el código de formato el separador decimal se escribe con punto; PowerPoint
lo muestra según la configuración regional de quien abre el archivo.

## 6. Fuentes de datos

Toda diapositiva con cifras lleva `source`. Si los datos son estimaciones o
supuestos tuyos, dilo ("Estimación propia a partir de…").

## 7. Gráficos editables

`build_deck.py` crea gráficos **nativos** de PowerPoint con sus datos
incrustados: el usuario puede hacer clic derecho → "Editar datos". No insertes
imágenes de gráficos (matplotlib) salvo para visualizaciones que PowerPoint no
soporta (mapas, redes); en ese caso, exporta a PNG de alta resolución (≥ 200 dpi)
con la paleta y fuente del tema y fondo transparente.
