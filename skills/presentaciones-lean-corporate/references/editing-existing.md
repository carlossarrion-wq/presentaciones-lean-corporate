# Trabajar con presentaciones existentes

Hay tres situaciones distintas; identifica cuál es antes de tocar nada.

| Situación | Enfoque |
|---|---|
| **Cambios puntuales** (erratas, cifras, notas, quitar o mover diapositivas) | Editar el archivo original con `pptx_ops.py` o python-pptx. No reconstruir |
| **Mejorar / rehacer** un deck flojo | Extraer contenido → rehacer la historia → nuevo JSON → `build_deck.py` |
| **Pasar un deck a la plantilla Lean Corporate** | Extraer contenido → JSON con `section`/`title`/`lead` → `build_deck.py` (perfil por defecto) |

Trabaja siempre sobre una copia; nunca sobrescribas el original del usuario.

## 1. Entender el deck

```bash
python scripts/inspect_pptx.py deck.pptx            # texto, gráficos (con datos), tablas y notas
python scripts/inspect_pptx.py deck.pptx --shapes   # + posición y tamaño de cada forma
python scripts/render_preview.py deck.pptx --out antes   # cómo se ve hoy
```

## 2. Cambios puntuales

```bash
python scripts/pptx_ops.py in.pptx out.pptx replace "Q3 2025" "Q3 2026"
python scripts/pptx_ops.py in.pptx out.pptx delete 4,7
python scripts/pptx_ops.py in.pptx out.pptx move 12 3
python scripts/pptx_ops.py in.pptx out.pptx notes 5 "Transición: …"
python scripts/pptx_ops.py in.pptx out.pptx chart 6 nuevos_datos.json
```

Con python-pptx directamente (cuando necesites más control):

```python
from pptx import Presentation
prs = Presentation("in.pptx")
slide = prs.slides[2]
for sh in slide.shapes:
    if sh.has_text_frame and "Título antiguo" in sh.text_frame.text:
        run = sh.text_frame.paragraphs[0].runs[0]
        run.text = "Nuevo título-acción"          # conserva fuente, color y tamaño
        for extra in sh.text_frame.paragraphs[0].runs[1:]:
            extra.text = ""
prs.save("out.pptx")
```

Reglas para no romper el formato:
- Cambia `run.text`, no `shape.text` ni `text_frame.text` (que borran el formato).
- Si el texto nuevo es más largo, comprueba que cabe (render + `qa_check.py`).
- Para gráficos usa `chart.replace_data()`: conserva estilo y colores.
- Para añadir una diapositiva con el estilo del deck, usa uno de sus layouts
  (`prs.slide_layouts[i]`) y rellena los marcadores (`slide.placeholders`).

## 3. Rehacer un deck

1. Extrae el contenido con `inspect_pptx.py --json`.
2. Identifica el mensaje principal y reescribe los títulos como titulares-acción
   (`storytelling.md`). A menudo un deck flojo tiene buen contenido mal ordenado.
3. Para cada diapositiva antigua decide: se mantiene, se fusiona, se divide, va a anexo o sobra.
4. Reasigna layouts (`layouts.md`): los muros de viñetas suelen convertirse en
   `cards`, `process`, `comparison` o `kpis`.
5. Conserva **todos los datos** y sus fuentes; no redondees cifras sin decirlo.
6. Construye, renderiza y compara antes/después.

Al entregar, resume en 2–3 líneas qué cambió (estructura, nº de diapositivas) y
si algo se movió a anexo o se eliminó.

## 4. Validación final

```bash
python scripts/qa_check.py out.pptx --no-info
python scripts/render_preview.py out.pptx --out despues
```

Si el archivo se generó o modificó programáticamente, ábrelo además con
`python -c "from pptx import Presentation; Presentation('out.pptx')"` para
confirmar que es válido.
