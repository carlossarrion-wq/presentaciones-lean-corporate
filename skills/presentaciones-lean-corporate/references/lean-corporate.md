# Convenciones de la casa: plantilla Lean Corporate

Guía extraída de la plantilla oficial y de ofertas reales de Lean Corporate
Services. El motor aplica automáticamente lo visual; tú decides contenido y estructura.

## 1. La plantilla

- Lienzo **26,66 × 15 in** (16:9 a doble escala). El motor trabaja en "pulgadas
  de diseño" (13,333 × 7,5) y escala ×2: no hace falta tenerlo en cuenta en el JSON.
- Fondo gris muy claro (#F9F9F9) con **formas turquesa** ("swoosh") propias de
  cada layout, logotipo **lean corporate** abajo a la izquierda y **número de
  página** abajo a la derecha (los pone la plantilla).
- Tipografía **IBM Plex Sans**: títulos en Bold, cuerpo en Light. Si el equipo
  que abre el archivo no la tiene, PowerPoint la sustituye; las máquinas
  corporativas IBM/Lean suelen tenerla.

### Layouts que usa el perfil (y por qué)

| Rol | Layout de la plantilla | Marcadores |
|---|---|---|
| Portada (`title`) | Cover Page | título (idx 0), subtítulo (idx 1) |
| Índice (`agenda`) | Agenda | título (0), lista numerada (1) |
| Separador (`section`) | Cover Page | "N. Sección" (0), subtítulo (1) |
| Contenido (resto) | Title & Bullets 2 | sección (0), subtítulo (13) |
| Cierre (`closing`) | Logo & Short Text | texto sobre turquesa (10) |

**No se usan** Divider, Statement/Quote, Thank you, Title, Bullets, Long Text ni
los de imagen: llevan el logotipo "lean customer" o no tienen logo/numeración.
Mezclarlos produce un deck con dos marcas distintas.

## 2. Anatomía de una diapositiva de contenido

```
┌──────────────────────────── swoosh ─────────────────────────────┐
│ 3. Caso de Uso “Gestión del Conocimiento”   ← section (turquesa, Bold, 1 línea)
│ Agentes a desarrollar como parte del Alcance ← title (verde azulado, Light, ≤ 60 car.)
│ A continuación se presentan los cuatro…      ← lead (párrafo, opcional, 1–3 líneas)
│                                                                  │
│   [ contenido: tarjetas, tabla, gráfico, proceso, texto… ]       │
│                                                                  │
│ lean/corporate   Fuente: …                                    9  │
└──────────────────────────────────────────────────────────────────┘
```

- **section**: la sección numerada del índice, idéntica en todas sus diapositivas
  ("1. Introducción", "3. Caso de Uso “…”", "5. Modelo Económico"). Se escribe
  una vez y el motor la arrastra a las siguientes diapositivas.
- **title** (subtítulo): el tema o, mejor, el **mensaje** de la diapositiva.
  Tema: "Problemática actual". Mensaje: "Problemática actual: 420 solicitudes
  manuales al año". Máximo ~60 caracteres para que quepa en una línea.
- **lead**: párrafo que contextualiza (estilo de las ofertas de la casa:
  "A continuación se presentan…", "Se estima que…"). Úsalo cuando el
  contenido visual necesita una frase de lectura; no repitas el subtítulo.
- **source** va junto al logotipo; **note** (notas tipo "(*) …") en la misma línea.

## 3. Estructura de una oferta / CR

Estructura habitual de una oferta de servicios Lean IT (documento para leer):

| # | Sección | Contenido típico | Layouts |
|---|---|---|---|
| – | Portada | "Oferta de Servicios Lean IT" · código y nombre de la CR · versión / fecha | `title` |
| – | Contenido | Índice numerado | `agenda` |
| 1 | Introducción | Solicitud de cambio de alcance, casos de uso, marco contractual (Contrato Master) | `text` |
| 2 | Antecedentes y Objetivos | Contexto, sesiones de trabajo, objetivos | `text`, `cards` |
| 3…n | Caso de Uso "X" | Problemática (cifras) · Solución · Journey · Agentes · Arquitectura · Comparativa de alternativas · Componentes reutilizables · Plan de implantación · Detalle de actividades (WBS) · Entregables | `kpis`, `process`, `cards`, `comparison`, `timeline`, `table`, `text` + diagramas a medida |
| n+1 | Modelo Económico | Precio de referencia, descuento, importe final, desglose por caso de uso, plan de facturación | `pricing` |
| n+2 | Términos y Condiciones | Procedimiento de Gestión de Cambios, aprobación, firmas | `text` + `signatures` |
| n+3 | Perfiles Tipo | Perfiles del equipo: responsabilidades y habilidades | `text` (2 columnas) |
| – | Cierre | Código y nombre de la CR | `closing` |

Para **presentaciones en reunión** (no documento) reduce texto: 8–15
diapositivas, subtítulos-mensaje, `lead` corto o ninguno, más `kpis`,
`statement` y `chart`; detalle (WBS, perfiles) en anexo.

## 4. Redacción

- Español de España, registro formal de propuesta; "Cliente" con mayúscula si se usa genérico.
- Nombres de casos de uso entre comillas tipográficas: Caso de Uso “Gestión del Conocimiento”.
- Códigos de entregables y CR tal cual: E340 – Diseño funcional; CR2026-021.
- Importes: `84.276,00 €`, negativos con signo menos `–34.276,00 €`; porcentajes `60 %`.
- Plazos en semanas ("aproximadamente 13 semanas") y fases con WBS ("WBS 1.x y 2.x").
- Énfasis con `**negrita**` solo en conceptos clave (1–3 por párrafo).

## 5. Colores (tokens del perfil)

| Token | Hex | Uso |
|---|---|---|
| accent | #08BED4 | Turquesa de marca: títulos de sección (plantilla), bandas, barras resaltadas, viñetas |
| accent_text | #0A8FA1 | Turquesa oscurecido para texto pequeño resaltado (contraste) |
| primary | #1F4E5E | Verde azulado: subtítulos, cabeceras de tabla, tarjeta destacada, cifras |
| accent_2 | #156082 | Azul petróleo: énfasis secundario |
| surface | #E5EFEF | Tarjetas y paneles |
| table_subheader | #CAEDFB | Cabeceras claras (tablas del modelo económico) |
| border | #ABD6D5 | Líneas y separadores |
| text / text_muted | #1E2A2F / #4C6573 | Texto principal / secundario |

No introduzcas colores fuera de esta paleta (los morados o verdes de diagramas
antiguos no son de marca). El turquesa #08BED4 no tiene contraste suficiente para
texto pequeño sobre blanco: para texto usa `accent_text` o `primary`.

## 6. Verificación específica antes de entregar

- [ ] Todas las diapositivas usan logotipo **lean corporate** (ninguna "lean customer").
- [ ] Los títulos de sección coinciden exactamente con el índice (numeración y comillas).
- [ ] Ningún subtítulo pasa a 2 líneas sin necesidad; ningún título de sección a 2 líneas.
- [ ] Importes del modelo económico cuadran (referencia – descuento = final; suma del desglose = total; plan de facturación = total).
- [ ] Código de CR, versión y fecha coherentes en portada, cierre y nombre de archivo.
- [ ] Firmantes y cargos reales (o marcados como pendientes) en Términos y Condiciones.
- [ ] Tablas WBS paginadas con "(cont.)" y cabecera repetida.
