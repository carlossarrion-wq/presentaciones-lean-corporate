---
name: presentaciones-lean-corporate
description: Crea presentaciones PowerPoint (.pptx) profesionales con la plantilla corporativa Lean Corporate (Lean Corporate Services, filial de IBM) — ofertas de servicios, solicitudes de cambio de alcance (CR), propuestas, informes de seguimiento, kickoffs y presentaciones a cliente. Úsalo cuando el usuario pida una presentación, un deck, unas diapositivas, un PowerPoint o un .pptx en el contexto de Lean / Lean Corporate / Lean IT, o cuando pida pasar un deck existente a esta plantilla. Incluye la plantilla, su perfil de diseño, método de storytelling, 20 layouts, gráficos editables, tablas que se paginan solas, modelo económico y control de calidad visual.
---

# Presentaciones Lean Corporate

Este skill genera decks con la **plantilla oficial Lean Corporate** (incluida en
`templates/lean-corporate/plantilla.pptx`): mismo patrón, fondos, logotipo,
numeración, tipografía IBM Plex Sans y colores de marca. Encima de la plantilla
aplica un criterio de diseño y narrativa de nivel consultora: un mensaje por
diapositiva, datos protagonistas, jerarquía clara y revisión visual antes de entregar.

La calidad sale de tres cosas, por este orden:
1. **La historia** → `references/storytelling.md`
2. **Las convenciones de la casa** → `references/lean-corporate.md` (léelo siempre)
3. **El control de calidad visual** → `references/qa-checklist.md`

## Estructura

```
presentaciones-lean-corporate/
├── SKILL.md
├── templates/lean-corporate/
│   ├── plantilla.pptx          ← plantilla oficial limpia (16 layouts, sin diapositivas)
│   └── profile.json            ← perfil: layout por rol, marcadores, rejilla, colores, tamaños
├── references/
│   ├── lean-corporate.md       ← CONVENCIONES DE LA CASA: anatomía de diapositiva, estructura de oferta, colores
│   ├── storytelling.md         ← narrativa, titulares, tipos de deck
│   ├── layouts.md              ← catálogo de layouts y cuándo usar cada uno
│   ├── spec-format.md          ← formato JSON completo
│   ├── data-visualization.md   ← gráficos y tablas
│   ├── design-principles.md    ← principios generales de diseño
│   ├── themes-and-branding.md  ← perfiles de plantilla y temas genéricos
│   ├── editing-existing.md     ← editar o migrar decks existentes
│   └── qa-checklist.md         ← revisión antes de entregar
├── scripts/
│   ├── build_deck.py           ← JSON → .pptx sobre la plantilla (perfil lean-corporate por defecto)
│   ├── textfit.py              ← medición de texto con métricas reales de fuente
│   ├── render_preview.py       ← .pptx → PNG + hoja de contactos
│   ├── qa_check.py             ← revisión automática (escala-consciente; modo --documento)
│   ├── inspect_pptx.py         ← radiografía de decks y plantillas
│   ├── pptx_ops.py             ← cambios puntuales sin reconstruir
│   └── install_fonts.sh        ← instala IBM Plex Sans para previsualizar fielmente
├── fonts/                      ← IBM Plex Sans (SIL OFL)
├── themes/themes.json          ← temas genéricos (solo para decks que NO son Lean)
└── assets/examples/
    ├── oferta_ejemplo.json     ← oferta completa con las convenciones de la casa
    └── ejemplo_completo.json   ← los 18 layouts genéricos
```

## Preparación (una vez por sesión)

```bash
pip install python-pptx pillow lxml          # si faltan
bash scripts/install_fonts.sh                # IBM Plex Sans para render y medición fieles
```
Para previsualizar hacen falta LibreOffice (`soffice`) y poppler (`pdftoppm`).

## Flujo de trabajo

### 0. Entender el encargo
Identifica (del mensaje o los adjuntos; pregunta solo lo imprescindible, máx. 1–3 preguntas):
- **Tipo de documento**: oferta / CR (se lee, densidad alta), presentación en
  reunión (se presenta, densidad baja), informe de seguimiento…
- **Cliente, código y versión** (p. ej. "CR2026-021", "v2.1") y fecha.
- **Audiencia y decisión** que se busca.
- **Material fuente**: léelo entero. Nunca inventes importes, plazos, nombres ni firmantes;
  si faltan, déjalos como pendientes visibles y dilo al entregar.

### 1. Diseñar la historia
Lee `references/storytelling.md` y `references/lean-corporate.md`. Define:
1. Las **secciones numeradas** del índice (serán los títulos grandes de las diapositivas).
2. Para cada diapositiva, su **subtítulo-mensaje** (≤ 60 caracteres) y su prueba
   (dato, tabla, gráfico, esquema).
3. El párrafo `lead` cuando haga falta contexto (estilo documento de oferta).

### 2. Elegir layout
`references/layouts.md`. En ofertas: `text` para contexto contractual,
`kpis`/`chart` para la problemática, `process`/`cards` para la solución,
`comparison` para alternativas, `timeline` + `table` (WBS) para el plan,
`pricing` para el modelo económico, `text` con `signatures` para términos.

### 3. Escribir el JSON
Formato en `references/spec-format.md`. Parte de `assets/examples/oferta_ejemplo.json`.
Claves de la casa:
- `section`: "3. Caso de Uso “Gestión del Conocimiento”" (título grande; se arrastra
  a las diapositivas siguientes si no lo repites).
- `title`: subtítulo-mensaje de esa diapositiva.
- `lead`: párrafo introductorio.
- `**texto**` resalta (negrita IBM Plex Sans; en subtítulos, además en turquesa).

### 4. Construir
```bash
python scripts/build_deck.py deck.json -o salida.pptx          # usa la plantilla Lean por defecto
```
Los `AVISO:` indican texto que no cabe o títulos de sección demasiado largos:
corrígelos (acortar o dividir), no los ignores. Las tablas largas se paginan
solas ("(cont.)") repitiendo la cabecera.

### 5. Revisar (obligatorio, al menos un ciclo)
```bash
python scripts/qa_check.py salida.pptx --no-info --documento    # --documento para ofertas/informes
python scripts/render_preview.py salida.pptx --out preview --dpi 40
```
Mira `preview/contact_sheet.png` y, ante cualquier duda, cada diapositiva a tamaño
completo. Repasa `references/qa-checklist.md` y la sección de verificación de
`references/lean-corporate.md`. Corrige el JSON, reconstruye, vuelve a mirar.

### 6. Entregar
El `.pptx` con 1–3 frases: qué contiene, supuestos y pendientes. Nombre de archivo
según el estilo de la casa: `CR2026-021_IAGen_Gestor_Conocimiento_-_SOLPED_v1.0.pptx`.

## Otros escenarios
- **Pasar un deck existente a la plantilla**: `inspect_pptx.py deck.pptx` → reescribir
  en JSON con las convenciones de la casa → construir. Ver `references/editing-existing.md`.
- **Cambios puntuales** en un deck ya hecho: `scripts/pptx_ops.py` (reemplazar texto,
  borrar/mover diapositivas, notas, datos de gráfico).
- **Diagramas a medida** (arquitecturas, flujos con iconos): construye el deck y añade la
  diapositiva con los helpers de `Deck` (ver final de `references/layouts.md`); usa
  `d.new_slide()` para heredar la plantilla.
- **Deck que NO es Lean** (otra marca o uso personal): `--no-profile --theme <tema>`.

## Reglas de oro
1. Usa solo los layouts de la plantilla con el logotipo **lean corporate** (el perfil
   ya lo hace); nunca los que llevan "lean customer".
2. Título grande = sección numerada del índice; subtítulo = mensaje (≤ 60 caracteres).
3. Una idea por diapositiva; si una tabla o un texto no cabe, divide (las tablas lo hacen solas).
4. Colores solo de la paleta de la plantilla; turquesa para resaltar con mesura.
5. Cifras siempre con fuente; importes con formato español (48.000,00 €).
6. Siempre: construir → renderizar → mirar → corregir.
