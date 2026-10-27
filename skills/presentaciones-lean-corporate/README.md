# presentaciones-lean-corporate — Skill de IA para presentaciones con la plantilla Lean Corporate

Versión 2 del skill de presentaciones profesionales, especializada en la plantilla
oficial **Lean Corporate**. Genera ofertas, CR, propuestas e informes en .pptx
construidos sobre los layouts reales de la plantilla (logotipo, formas, numeración,
IBM Plex Sans), con criterio de narrativa y diseño, gráficos editables, tablas WBS
que se paginan solas, modelo económico y control de calidad visual.

## Instalación

**Claude (claude.ai / app):** Ajustes → Capacidades → Skills → subir el `.zip`.
Para toda la organización, un administrador puede publicarlo.

**Claude Code / Agent SDK:** copiar la carpeta en `~/.claude/skills/presentaciones-lean-corporate/`.

**Dependencias:**
```bash
pip install -r requirements.txt
bash scripts/install_fonts.sh         # IBM Plex Sans (incluida, licencia OFL)
# Previsualización: LibreOffice + poppler
```

## Uso directo

```bash
python scripts/build_deck.py assets/examples/oferta_ejemplo.json -o oferta.pptx
python scripts/qa_check.py oferta.pptx --no-info --documento
python scripts/render_preview.py oferta.pptx --out preview --dpi 40
```

## Novedades respecto a la v1

| | v1 (presentaciones-pro) | v2 (esta) |
|---|---|---|
| Base visual | 6 temas genéricos dibujados por el motor | Plantilla Lean Corporate real (layouts, fondos, logo, numeración) |
| Escala | 13,33 × 7,5 in | Cualquier escala (Lean: 26,66 × 15 in) |
| Títulos | Título-acción | Sección numerada + subtítulo-mensaje + `lead` (estilo de la casa) |
| Layouts | 18 | 20 (+ `text` con firmas, + `pricing`) |
| Tablas | Una diapositiva | Paginación automática, separadores de grupo, bandas, fila total |
| QA | Umbrales 16:9 | Escala-consciente, modo `--documento` |

Si se necesita otra plantilla corporativa, ver "Crear un perfil" en
`references/themes-and-branding.md`.
