# Formato de la especificación (deck.json)

```json
{
  "meta": { ... },
  "theme_overrides": { ... },      // opcional
  "slides": [ { "layout": "...", ... }, ... ]
}
```

Ejemplos probados: `assets/examples/oferta_ejemplo.json` (oferta con las
convenciones Lean Corporate) y `assets/examples/ejemplo_completo.json` (los 18 layouts genéricos).

## meta

| Campo | Tipo | Descripción |
|---|---|---|
| `profile` | str | Perfil de plantilla: `lean-corporate` (defecto), otra carpeta de `templates/`, una ruta a `profile.json`, o `"none"` para temas genéricos |
| `title` | str | Título del documento (propiedades del archivo) |
| `author` | str | Autor; también aparece en la portada si no se indica allí |
| `date` | str | Fecha visible en portada ("9 Marzo 2026") |
| `version` | str | Versión visible en portada ("v2.1"); con perfil se muestra "v2.1 / fecha" |
| `client` | str | Línea opcional de cliente en portada |
| `theme` | str | `ejecutivo` (defecto), `consultoria`, `tecnologia`, `calido`, `minimal`, `institucional` |
| `theme_file` | str | Ruta a un JSON de tema propio (ver themes-and-branding.md) |
| `template` | str | Ruta a .pptx/.potx corporativo (se usa su patrón) |
| `use_template_background` | bool | Con plantilla: `true` (defecto) respeta su fondo |
| `footer` | str | Texto del pie (solo temas genéricos; con perfil el pie lo pone la plantilla) |
| `slide_numbers` | bool | Numeración (defecto `true`) |

`theme_overrides`: cambia colores/fuentes sueltos sin crear un tema, p. ej.
`{"accent": "E4572E", "font_heading": "Arial"}`.

## Convenciones de texto

- `**palabras**` → énfasis. En títulos y conclusiones se pinta en color de acento;
  en el cuerpo, en negrita.
- Saltos de línea manuales con `\n` (úsalos poco: el ajuste es automático).
- Listas de viñetas: `["texto", {"text": "texto", "sub": ["detalle"]}]`.
- Números en formato local del público: `"4,2 M€"`, `"61 %"`, `"−4 pp"` (signo menos tipográfico).

## Campos comunes de diapositiva

| Campo | Aplica a | Descripción |
|---|---|---|
| `layout` | todas | Nombre del layout (ver tabla siguiente) |
| `section` | contenido (perfil) | Título grande = sección numerada ("3. Caso de Uso “X”"). Se arrastra a las siguientes diapositivas |
| `title` | casi todas | Con perfil: subtítulo-mensaje (≤ 60 car.). Sin perfil: título-acción (≤ 2 líneas). Admite `**énfasis**` |
| `subtitle` | contenido (perfil) | Subtítulo cuando no hay `section` (entonces `title` va al título grande) |
| `lead` | contenido | Párrafo (o lista de párrafos) introductorio bajo el título |
| `note` | contenido | Nota al pie tipo "(*) …" (con perfil, junto al logotipo) |
| `kicker` | contenido, title, statement | Antetítulo corto en mayúsculas |
| `takeaway` | contenido | Conclusión al pie, 1 línea |
| `source` | contenido | Fuente de datos (se antepone "Fuente:" si falta) |
| `notes` | todas | Notas del orador |
| `title_size` | contenido | Tamaño máximo del título (defecto 28) |

## Campos por layout

### title
`title`, `subtitle`, `kicker`, `author`, `date`, `version`, `client`, `image` (ruta, opcional;
con perfil se ignora: la portada de la plantilla no lleva imagen).
Con perfil Lean: `title` = "Oferta de Servicios Lean IT"; `subtitle` = "CR2026-021 Nombre de la CR"
(negrita turquesa); debajo "versión / fecha".

### section
`number` ("3" → "3. Título"), `title`, `subtitle`. Con perfil usa la portada como
separador y fija la sección vigente para las diapositivas siguientes. Opcional en ofertas.

### agenda
`title`, `items`: lista de str o `{"title", "text"}` (1–10), `current` (índice opcional), `kicker`.

### statement
`text` (o `title`), `kicker`, `subtitle`, `dark` (bool).

### bullets
`bullets` (lista), `size` (máx. pt, defecto 22), `callout`: `{"value", "label"}` o `{"text"}`.

### two-column
`left` y `right`: `{"heading", "bullets"}` o `{"heading", "text"}`; `style`: `"contrast"`.

### cards
`cards`: lista (1–8) de `{"title", "text", "icon"?, "highlight"?}`; `numbered` (bool, defecto true).

### kpis
`kpis`: lista (1–6) de `{"value", "label", "delta"?, "delta_label"?, "trend"?: "up|down|flat", "good"?: "up|down"}`;
`highlight` (índice); `bullets` (opcional, debajo de los indicadores).

### big-number
`value`, `label`, `context` (str o lista de viñetas), `title` (opcional).

### chart
```json
"chart": {
  "type": "column | bar | stacked-column | stacked-bar | stacked-column-100 | line | area | pie | doughnut",
  "categories": ["Ene", "Feb", "Mar"],
  "series": [{"name": "2026", "values": [1.2, 1.5, 1.9], "color": "opcional hex"}],
  "number_format": "0.0 | 0% | #,##0 | 0.0\" M€\"",
  "label_format": "opcional, formato de etiquetas",
  "highlight": 2,              // índice de categoría a resaltar (1 serie)
  "data_labels": true,         // por defecto: true con 1 serie / apilados / tartas
  "show_percent": true,        // tartas/donut: porcentaje (true) o valor (false)
  "y_min": 0, "y_max": 100,    // opcional
  "subtitle": "Unidad o métrica (p. ej. 'Ventas, M€')",
  "gap_width": 60, "font_size": 12, "label_size": 12
}
```
Más `insights` (lista de viñetas para el panel lateral) e `insights_title` (defecto "Claves").
Los valores son números (no strings). Porcentajes como fracción (0.24) con formato `0%`.

### table
`columns` (cabeceras), `rows` (lista de listas de str), `col_align` (lista `left|center|right`),
`col_widths` (relativos), `highlight_row` (índice 0-based dentro de `rows`), `zebra` (bool),
`size` (pt, opcional), `min_size` (defecto 10,5), `numeric: true` (alinea a la derecha todas menos la primera),
`band` (banda turquesa con título sobre la tabla), `header_style` (`dark` | `light`),
`total_row` (última fila como total), `paginate` (defecto `true`).
- Una fila con todas las celdas vacías (`["", "", ""]`) es un **separador** fino entre grupos (p. ej. fases WBS).
- Si la tabla no cabe ni a `min_size`, se **pagina**: diapositivas adicionales con
  "(cont.)" en el subtítulo y la cabecera repetida.

### text
Texto largo estilo documento.
```json
{"layout": "text", "section": "1. Introducción", "title": "Solicitud de Cambio de Alcance",
 "blocks": [
   "Párrafo normal con **negrita**.",
   {"heading": "Encabezado de bloque"},
   {"bullets": ["Viñeta", {"text": "Viñeta con sub", "sub": ["detalle"]}]}
 ],
 "columns": 1,
 "signatures": [{"name": "Nombre Apellido", "role": "Cargo", "company": "Empresa, S.A.", "label": "Firma y fecha"}]}
```
`columns: 2` reparte los bloques en dos columnas (perfiles, listas largas).

### pricing
Modelo económico.
```json
{"layout": "pricing", "section": "5. Modelo Económico",
 "band": "Coste de los Servicios de “CR2026-021 …”",
 "summary": [{"label": "Precio de referencia de la propuesta", "value": "84.276,00 €"},
             {"label": "Descuento comercial aplicado (*)", "value": "–34.276,00 €"},
             {"label": "Importe total final de la oferta", "value": "50.000,00 €", "strong": true}],
 "tables": [{"band": "Desglose del importe final por caso de uso",
             "columns": ["Descripción", "Importe Final"],
             "rows": [["Caso de Uso “A”", "27.500,00 €"], ["Caso de Uso “B”", "22.500,00 €"], ["Total", "50.000,00 €"]],
             "col_widths": [4, 1], "total_row": true},
            {"band": "Plan de facturación", "columns": ["Caso de Uso", "Mes 1", "Mes 2"], "rows": [["…", "6.875 €", "6.875 €"]]}],
 "note": "(*) Nota sobre el descuento o la financiación."}
```

### timeline
`steps`: 2–7 de `{"label" (fecha), "title", "text"}`; `current` (índice).

### process
`steps`: 2–6 de `{"title", "text"}`; `highlight` (índice).

### comparison
`options`: 2–4 de `{"name", "pros": [], "cons": [], "recommended"?: true, "badge"?, "verdict"?}`.

### matrix
`quadrants`: 4 de `{"title", "text"}` (sup-izq, sup-dcha, inf-izq, inf-dcha);
`x_label`, `y_label`, `highlight` (índice), `bullets` (panel lateral opcional).

### quote
`quote`, `author`, `role`, `dark` (bool).

### image
`image` (ruta relativa al JSON o absoluta), `side` (`left`|`right`), `bullets` o `text`, `caption`.

### closing
`title` (defecto "Gracias"), `subtitle`, `next_steps` (lista), `contact` (str o lista).

## Errores comunes del JSON

- Comas finales o comentarios `//` (JSON no los admite; los de arriba son explicativos).
- Valores de gráfico como texto (`"1,5"`) en lugar de número (`1.5`).
- `highlight_row` contando la cabecera (no la cuenta).
- Rutas de imagen relativas a otro directorio: son relativas al archivo JSON.
