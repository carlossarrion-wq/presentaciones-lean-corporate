# excel-expert — Skill para agentes IA

Skill para crear, leer, editar, validar, auditar y convertir ficheros Excel con calidad
profesional: fórmulas vivas, cero errores, formato cuidado y ediciones que no rompen el libro.

## Contenido

```
excel-expert/
├── SKILL.md                     # Instrucciones principales (lo que lee el agente)
├── requirements.txt
├── scripts/
│   ├── inspect_xlsx.py          # Radiografía del libro (estructura, fórmulas, entradas, riesgos)
│   ├── recalc.py                # Recalcula con LibreOffice y detecta errores (#REF!, #NAME?...)
│   ├── audit_formulas.py        # Fórmulas incoherentes, números mágicos, refs vacías, circulares
│   ├── diff_xlsx.py             # Qué celdas cambiaron entre dos versiones
│   ├── patch_cells.py           # Edición XML quirúrgica (conserva slicers, formas, macros...)
│   ├── build_from_spec.py       # Libro completo desde una especificación JSON
│   ├── convert.py               # CSV/XLS/ODS/PDF ↔ XLSX (incl. CSV con formato español)
│   ├── xlsx_helpers.py          # Biblioteca: estilos, tablas, gráficos, validaciones, inserción segura
│   └── _soffice.py              # Ejecución aislada de LibreOffice headless
├── references/                  # Contexto que el agente carga bajo demanda
│   ├── openpyxl-recetas.md      ├── formulas.md          ├── formato-y-diseno.md
│   ├── graficos.md              ├── pandas-datos.md      ├── edicion-segura.md
│   ├── modelos-financieros.md   ├── rendimiento.md       └── problemas-frecuentes.md
├── assets/spec_example.json     # Ejemplo de especificación para build_from_spec.py
└── tests/smoke_test.py          # Verificación de la instalación
```

## Requisitos

- Python ≥ 3.9 → `pip install -r requirements.txt`
- LibreOffice (para recalcular y convertir). Sin él, el skill crea y edita libros, pero no
  puede calcular valores ni verificar errores.
  - Debian/Ubuntu: `apt-get install -y libreoffice-calc`
  - macOS: `brew install --cask libreoffice`
  - Windows: instalador oficial; si no está en el PATH, define `SOFFICE_PATH`.

Verifica la instalación:
```bash
python tests/smoke_test.py
```

## Instalación

### Claude Code
Copia la carpeta en el directorio de skills personal o del proyecto:
```bash
cp -r excel-expert ~/.claude/skills/            # para todos tus proyectos
cp -r excel-expert .claude/skills/              # solo para este repositorio
```

### Claude (app web/escritorio) y Claude API
Sube el `.zip` del skill desde la sección de Skills de la configuración, o regístralo
mediante la API de Skills. Consulta el procedimiento vigente en https://docs.claude.com.
El entorno de ejecución ya incluye openpyxl, pandas y normalmente LibreOffice.

### Otros agentes (LangChain/LangGraph, OpenAI Agents SDK, watsonx Orchestrate, CrewAI…)
El skill es independiente del proveedor. Para integrarlo:
1. Copia la carpeta en el entorno de ejecución del agente.
2. Inyecta el contenido de `SKILL.md` (sin el bloque YAML) en el prompt del sistema, o
   carga solo `name` + `description` y entrega el cuerpo cuando la tarea implique Excel
   (divulgación progresiva).
3. Dale al agente una herramienta para ejecutar Python/shell y otra para leer ficheros,
   para que pueda ejecutar `scripts/` y leer `references/` bajo demanda.
4. Indícale la ruta absoluta del skill (`Las rutas scripts/... son relativas a /opt/skills/excel-expert`).

## Uso rápido (manual)

```bash
python scripts/inspect_xlsx.py libro.xlsx --formulas
python scripts/build_from_spec.py assets/spec_example.json demo.xlsx
python scripts/recalc.py demo.xlsx
python scripts/audit_formulas.py demo.xlsx
python scripts/patch_cells.py demo.xlsx -o demo2.xlsx --set "Supuestos!B2=0.1"
python scripts/diff_xlsx.py demo.xlsx demo2.xlsx
python scripts/convert.py datos.csv datos.xlsx --sep ";" --decimal ","
```

## Personalización para tu organización

- Paleta, fuente y formatos por defecto: diccionarios `PALETTE`, `FONT_NAME` y
  `NUMBER_FORMATS` al inicio de `scripts/xlsx_helpers.py`.
- Convenciones corporativas (plantillas, nombres de hojas, colores de marca): añádelas como
  una sección "Convenciones de <empresa>" en `SKILL.md` o en un nuevo fichero de
  `references/` enlazado desde la tabla de referencias.
