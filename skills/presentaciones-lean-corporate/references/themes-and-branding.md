# Perfiles de plantilla, temas y marca

Esta versión del skill trabaja por defecto con el **perfil `lean-corporate`**
(`templates/lean-corporate/`). Los temas genéricos de más abajo solo se usan con
`--no-profile` para decks que no son de Lean.

## Perfiles de plantilla (modo nativo)

Un perfil convierte una plantilla corporativa en un "motor" de diapositivas: el
deck se construye **sobre los layouts reales** de la plantilla (fondos, logos,
numeración, estilos de título), y el motor solo dibuja el contenido.

`templates/<nombre>/profile.json`:

| Clave | Contenido |
|---|---|
| `template` | Archivo .pptx de la plantilla (en la misma carpeta), **sin diapositivas** |
| `theme` | Tokens de color y fuentes extraídos de la plantilla (mismas claves que un tema, más `font_emph`, `accent_text`, `bullet_char`, `band`, `table_header`, `table_subheader`) |
| `grid` | Margen, inicio del contenido con/sin subtítulo, límite inferior, posición de la fuente, anchos de texto en portada (pulgadas de diseño 13,333 × 7,5) |
| `sizes` | Tamaños de título, subtítulo, lead y cuerpo (pt de diseño) |
| `roles` | Para cada rol (`content`, `cover`, `section`, `agenda`, `closing`): nombre del layout y `idx` de los marcadores `title` / `subtitle` / `body` |
| `conventions` | Notas para el agente (layouts a evitar y por qué) |

El motor detecta la escala de la plantilla (la de Lean mide 26,66 × 15 in → ×2) y
todo se expresa en unidades de diseño.

### Crear un perfil para otra plantilla

1. Copia la plantilla y elimina sus diapositivas (y limpia `docProps/app.xml`,
   que conserva los títulos de las diapositivas originales):
   ```python
   import sys; sys.path.insert(0, "scripts")
   from build_deck import _remove_all_slides, _load_template
   from pptx import Presentation
   p = Presentation(_load_template("original.pptx")); _remove_all_slides(p); p.save("templates/mi-marca/plantilla.pptx")
   ```
2. Inspecciona layouts y marcadores: `python scripts/inspect_pptx.py templates/mi-marca/plantilla.pptx --template`.
3. Renderiza una diapositiva por layout (rellenando cada marcador con su nombre)
   para ver fondos, logos y numeración; descarta layouts con logos de otra marca.
4. Copia `templates/lean-corporate/profile.json`, cambia `roles` (nombres de layout
   e `idx`), `grid` (posiciones medidas ÷ escala) y `theme` (colores reales usados
   en las diapositivas; `inspect_pptx.py --emit-theme` da un punto de partida).
5. Construye `assets/examples/oferta_ejemplo.json` con `--profile mi-marca`,
   renderiza y ajusta hasta que no haya solapes con logos o formas decorativas.

## Temas incluidos (`themes/themes.json`)

| Tema | Carácter | Fuentes | Para |
|---|---|---|---|
| `ejecutivo` | Azul marino + verde azulado | Calibri | Comités, banca, utilities, sector público (defecto) |
| `consultoria` | Verde profundo + coral | Cambria (títulos, sin negrita) + Calibri | Estrategia, informes tipo consultora |
| `tecnologia` | Modo oscuro, azul eléctrico + turquesa | Arial | Producto, arquitectura, IA, demos |
| `calido` | Terracota + arena | Cambria + Calibri | RR. HH., formación, marketing, sostenibilidad |
| `minimal` | Blanco y negro + rojo | Arial | Keynotes, mensajes de alto impacto |
| `institucional` | Índigo + dorado | Calibri | Propuestas, licitaciones, informes anuales |

Si el usuario no indica nada, elige por el contexto (sector, audiencia, tono)
y menciónalo al entregar.

## Claves de un tema

| Clave | Uso |
|---|---|
| `bg` | Fondo de diapositivas de contenido |
| `bg_dark` | Fondo de portada, secciones, cierre, panel de agenda |
| `surface` | Tarjetas, paneles, bandas de conclusión, filas alternas |
| `border` | Líneas finas, separadores, elementos inactivos |
| `text`, `text_muted` | Texto principal y secundario |
| `text_on_dark`, `text_on_dark_muted` | Texto sobre `bg_dark` |
| `primary` | Color de marca: cabeceras de tabla, cifras, tarjetas destacadas |
| `accent` | Lo que debe mirarse: énfasis, barra resaltada, viñetas |
| `accent_2` | Secundario (poco uso) |
| `positive`, `negative` | Variaciones buenas/malas |
| `chart` | Paleta de series (6 colores, el 1.º y 2.º son los más usados) |
| `font_heading`, `font_body`, `heading_bold` | Tipografía |
| `dark_content` | `true` si el contenido va sobre fondo oscuro |

## Tema propio (marca del cliente)

Crea un JSON con solo lo que cambia; el resto se hereda de `base`:

```json
{
  "base": "ejecutivo",
  "primary": "0530AD",
  "accent": "24A148",
  "bg_dark": "001141",
  "chart": ["0530AD", "24A148", "8A3FFC", "F1C21B", "A2A9B0", "DA1E28"],
  "font_heading": "Arial",
  "font_body": "Arial"
}
```

```bash
python scripts/build_deck.py deck.json -o salida.pptx --theme-file marca.json
```

Comprobaciones al definir una marca:
- Contraste del texto sobre `bg`, `surface` y `bg_dark` (≥ 4,5:1).
- El `accent` debe distinguirse claramente del `primary`.
- Si la marca usa una fuente no estándar, confirma que el público la tiene
  instalada; si no, usa una estándar de métrica similar (Arial, Calibri).

## Plantilla corporativa (.pptx / .potx)

1. **Inspecciona** la plantilla y extrae su tema:
   ```bash
   python scripts/inspect_pptx.py plantilla.potx --template --emit-theme marca.json
   ```
   Muestra colores del tema, fuentes y layouts del patrón, y genera `marca.json`.
   Revisa ese JSON: el orden de los acentos de una plantilla no siempre
   coincide con primario/acento; ajusta `primary`, `accent`, `surface`, `border`.
2. **Construye** sobre la plantilla:
   ```bash
   python scripts/build_deck.py deck.json -o salida.pptx --template plantilla.potx --theme-file marca.json
   ```
   El motor elimina las diapositivas de ejemplo de la plantilla, usa su layout
   en blanco (con logos y elementos del patrón) y respeta su fondo en las
   diapositivas de contenido (`"use_template_background": false` para pintar el
   del tema).
3. **Revisa el render**: los logos o franjas del patrón pueden invadir la zona
   de título o pie. Si ocurre, ajusta `TOP`, `M` o `BOTTOM` al inicio de
   `build_deck.py` (constantes de la rejilla) o elige otro layout del patrón.

Si la plantilla no es 16:9 el motor avisa: los layouts están optimizados para 16:9.

## Fuentes

Las fuentes por defecto (Calibri, Arial, Cambria) están en cualquier Office.
Para previsualizar en Linux, LibreOffice las sustituye por equivalentes
métricos (Carlito, Liberation Sans, Caladea), así que el ajuste de texto
medido coincide con el de PowerPoint.
