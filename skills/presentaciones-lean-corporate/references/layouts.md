# Catálogo de layouts

Cada layout responde a un **tipo de mensaje**. Elige primero el mensaje y
después el layout, nunca al revés. Los campos exactos de cada uno están en
`spec-format.md`.

Campos comunes a casi todos: `title`, `kicker` (antetítulo), `takeaway`
(conclusión al pie), `source` (fuente), `notes` (notas del orador).

| Layout | Úsalo para… | Capacidad cómoda | Evítalo si… |
|---|---|---|---|
| `title` | Portada | Título ≤ 12 palabras, subtítulo 1 línea | — |
| `agenda` | Índice; con `current` se reutiliza como separador | 3–7 puntos (hasta 10 en 2 col.) | El deck tiene < 8 diapositivas |
| `section` | Separar bloques en decks largos | Título 2–5 palabras | Hay < 3 secciones |
| `statement` | Un mensaje clave, transición o la petición final | 1 frase ≤ 20 palabras | Necesitas mostrar pruebas |
| `big-number` | Una cifra que impacta + su desglose | 1 cifra + 1–4 líneas de contexto | Hay varias cifras igual de importantes (→ `kpis`) |
| `kpis` | Cuadro de mando: 2–6 indicadores con variación | 3–4 ideal | Los números no son comparables entre sí |
| `chart` | Evolución, ranking, composición, correlación | 1 gráfico + 2–4 insights | Hay < 3 datos (→ `big-number`/`kpis`) |
| `table` | Datos exactos que se consultan (financieros, especificaciones) | ≤ 8 filas × 6 columnas | Quieres mostrar una tendencia (→ `chart`) |
| `cards` | Pilares, capacidades, beneficios, principios | 3–4 ideal (2–8) | Los elementos tienen orden temporal (→ `process`) |
| `two-column` | Antes/después, problema/solución, dentro/fuera de alcance | 3–5 viñetas por columna | Hay 3+ elementos (→ `cards`/`comparison`) |
| `comparison` | Elegir entre 2–4 opciones con pros/contras | 2–4 pros + 1–3 contras por opción | No hay recomendación (añádela: es el valor) |
| `process` | Pasos de un método o flujo | 3–5 pasos | Los pasos tienen fechas (→ `timeline`) |
| `timeline` | Hoja de ruta, hitos, historia | 4–6 hitos | Más de 7 hitos (agrupa en fases) |
| `matrix` | Priorización 2×2, segmentación | 4 cuadrantes, 1–2 líneas cada uno | Los ejes no son continuos/claros |
| `quote` | Voz del cliente, testimonio, cita de experto | ≤ 35 palabras | La cita es genérica o inventada |
| `image` | Producto, lugar, captura, equipo real | Título + 3–4 viñetas | No tienes una imagen relevante |
| `bullets` | Lista de ideas sin estructura visual mejor | 3–5 viñetas (máx. 6), opcional `callout` | Las ideas son paralelas y cortas (→ `cards`) |
| `closing` | Próximos pasos + contacto | 3–5 pasos | — |
| `text` | Texto de documento: introducción, marco contractual, términos, perfiles | ≤ 170 palabras (oferta) | Se va a presentar en directo (→ `bullets`/`cards`) |
| `pricing` | Modelo económico: importes, desglose, plan de facturación | 3 líneas + 2 tablas cortas | Tablas largas (→ `table` con `band`) |

## Comportamiento con la plantilla Lean Corporate

Con el perfil `lean-corporate` (por defecto):
- `title` → layout *Cover Page*; `agenda` → *Agenda* (lista con numeración automática);
  `section` → *Cover Page* como separador; `closing` → *Logo & Short Text*
  (texto blanco sobre turquesa); todo lo demás → *Title & Bullets 2*.
- En contenido, el título grande es la `section` y el subtítulo es `title`.
- `statement` y `quote` se dibujan sobre fondo claro (la plantilla no tiene un fondo
  oscuro coherente); `"dark": true` se ignora.
- `image` coloca la imagen dentro del área de contenido (no a sangre, para no
  pisar las formas de la plantilla ni el logotipo).
- El pie (logo + número) lo pone la plantilla; `source` y `note` van junto al logo.

## Ritmo del deck

- Alterna densidad: tras una `table` o un `chart` denso, una diapositiva ligera
  (`statement`, `big-number`, `quote`).
- No más de 2 diapositivas seguidas con el mismo layout (excepto series de
  `chart` con la misma estructura, que es legítimo).
- Usa fondos oscuros (`section`, `statement` con `"dark": true`, `quote` con
  `"dark": true`) como respiraciones; no más de 1 de cada 4 diapositivas.
- En decks > 12 diapositivas, reutiliza `agenda` con `current` para indicar en
  qué sección estamos, o usa `section`.

## Detalles útiles por layout

- **title**: `image` opcional ocupa el 42 % derecho a sangre. `author`/`date`
  se toman de `meta` si no se indican.
- **agenda**: `items` como texto o `{title, text}`; `current: n` (0-based)
  atenúa el resto.
- **bullets**: viñetas de 2º nivel con `{"text": "...", "sub": ["..."]}`. El
  `callout` (`{value, label}` o `{text}`) añade un panel a la derecha: úsalo
  para que una diapositiva de viñetas no se vea vacía.
- **two-column**: `"style": "contrast"` pinta la izquierda en gris (antes /
  situación actual) y la derecha tintada (después / propuesta).
- **cards**: `highlight: true` en una tarjeta la rellena con el color primario;
  `icon` sustituye el número (usa 1–3 caracteres: "€", "IA", "24h");
  `"numbered": false` quita los círculos.
- **kpis**: `delta` con signo; `trend` (`up`/`down`/`flat`) y `good` (`up` o
  `down`, qué dirección es buena) deciden flecha y color. Ej.: tiempo de espera
  ↑ es malo → `"good": "down"`.
- **chart**: `highlight: i` resalta una categoría en color de acento y atenúa el
  resto (solo con una serie). `insights` añade panel lateral de claves.
- **table**: `highlight_row`, `col_align`, `col_widths` (relativos),
  `zebra: false`. Negrita con `**…**` dentro de celdas.
- **timeline**: `current: i` marca el hito actual (anteriores en primario,
  siguientes atenuados).
- **process**: `highlight: i` colorea un paso con el acento.
- **comparison**: `recommended: true` + `badge` y `verdict` por opción.
- **matrix**: cuadrantes en orden sup-izq, sup-dcha, inf-izq, inf-dcha;
  `highlight: i`; `bullets` opcional a la derecha.
- **statement / quote**: `"dark": true` para fondo oscuro.
- **image**: `side` `left`/`right`; `caption` sobre la imagen. Si la imagen no
  existe deja un marcador gris y avisa (no entregues con marcadores).

## Diapositivas a medida

Si un mensaje necesita una composición que ningún layout ofrece (un diagrama
de arquitectura, un organigrama, un mapa), constrúyela con las utilidades de
`Deck` para heredar márgenes, colores y fuentes:

```python
import sys; sys.path.insert(0, "scripts")
import build_deck as bd

spec = {...}                                   # tu especificación habitual
deck = bd.build(spec, "tmp.pptx")             # construye el deck estándar
d = deck                                       # objeto Deck con tema y helpers
s = {"title": "La plataforma se integra con **tres sistemas** existentes",
     "source": "Arquitectura objetivo v2"}
sl = d.new_slide()                             # con perfil: layout de contenido de la plantilla
top = d.title(sl, s)                           # sección + subtítulo (+ lead) en su sitio
x, y, w, h = d.body_box(s, top)                # caja útil del cuerpo
box_w = (w - 2 * bd.GAP) / 3
for i, name in enumerate(["CRM", "Facturación", "Telefonía"]):
    bx = x + i * (box_w + bd.GAP)
    d.rect(sl, bx, y + 2.4, box_w, 1.0, d.t["surface"])
    d.text(sl, bx, y + 2.4, box_w, 1.0, name, size=18, bold=True,
           align="center", anchor="middle")
    d.line(sl, bx + box_w / 2, y + 1.85, bx + box_w / 2, y + 2.4, d.t["border"], 1.5)
d.line(sl, x + w / 2, y + 1.3, x + w / 2, y + 1.85, d.t["border"], 1.5)      # tronco
d.line(sl, x + box_w / 2, y + 1.85, x + w - box_w / 2, y + 1.85, d.t["border"], 1.5)  # bus
d.rect(sl, x + w / 2 - 2, y, 4, 1.3, d.t["primary"])
d.text(sl, x + w / 2 - 2, y, 4, 1.3, "Plataforma IA", size=20, bold=True,
       color="FFFFFF", align="center", anchor="middle")
d.footer(sl, s); d.notes(sl, s)
d.prs.save("salida.pptx")
```

Para colocarla en otra posición, reordena después (ver `editing-existing.md`).
Mantén las mismas reglas: un mensaje, título-acción, acento con mesura.
