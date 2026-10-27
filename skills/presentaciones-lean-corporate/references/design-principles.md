# Principios de diseño

> Con la plantilla Lean Corporate, la rejilla, las fuentes (IBM Plex Sans) y la
> paleta vienen del perfil (`lean-corporate.md`). Los principios de esta guía
> siguen aplicando a todo lo que se dibuja encima.

`build_deck.py` aplica estas reglas automáticamente. Conócelas para (a) elegir
bien el contenido que le das, (b) juzgar el render y (c) diseñar diapositivas a
medida que no desentonen.

## 1. Lienzo y rejilla

- Formato **16:9** (13,333 × 7,5 in). Es el estándar actual de PowerPoint y pantallas.
- Márgenes laterales **0,6 in**; nada de contenido fuera de ellos (excepto
  imágenes y fondos a sangre, que tocan el borde a propósito).
- Zonas fijas en todas las diapositivas de contenido:

| Zona | Posición (in) | Contenido |
|---|---|---|
| Antetítulo (kicker) | y ≈ 0,40 | Opcional, mayúsculas, 11 pt, color acento |
| Título | y ≈ 0,45–1,5 | 20–28 pt, máx. 2 líneas |
| Cuerpo | desde fin del título + 0,3 hasta 6,8 | Contenido |
| Conclusión (takeaway) | banda de 0,66 in sobre el pie | Opcional |
| Fuente | y ≈ 6,72 | 9 pt cursiva |
| Pie | y ≈ 7,02 | Texto de pie + nº de página, 9 pt |

- Separación estándar entre bloques: **0,3 in**. Dentro de una tarjeta: 0,25–0,3 in de relleno.
- Alinea todo a bordes comunes. Si dos elementos están "casi" alineados, el ojo
  lo percibe como error.

## 2. Tipografía

- **Máximo 2 familias** (títulos + cuerpo); a menudo una sola basta.
- Usa fuentes que existan en el equipo de quien abrirá el archivo: Calibri,
  Arial, Cambria, Segoe UI, Georgia. Una fuente bonita que no está instalada se
  sustituye y rompe la maquetación.
- Escala recomendada:

| Elemento | Tamaño | Peso |
|---|---|---|
| Título de portada | 36–44 pt | Negrita |
| Título de diapositiva | 24–28 pt (mín. 20) | Negrita (o regular en serif) |
| Cifra protagonista | 54–150 pt | Negrita |
| Subtítulo / encabezado de bloque | 18–22 pt | Negrita |
| Cuerpo | 14–22 pt (mín. 12) | Regular |
| Etiquetas de gráfico | 11–12 pt | Regular |
| Fuentes, pie | 9 pt | Regular / cursiva |

- Jerarquía por **tamaño y peso**, no por colores distintos ni subrayados.
- Alineación a la izquierda por defecto. Centrado solo en `statement`, portadas
  centradas y etiquetas cortas bajo iconos/hitos. Nunca justificado.
- Interlineado 1,0–1,15; espacio entre viñetas ≈ 40 % del tamaño de letra.
- Mayúsculas solo en antetítulos cortos, con espaciado de letra amplio.

## 3. Color

- Paleta = 1 color primario (marca, fondos oscuros, cabeceras), 1 acento (lo
  que debe mirar la audiencia), neutros (texto, grises, superficies).
- Proporción aproximada **60–30–10**: 60 % fondo/neutros, 30 % primario, 10 % acento.
- El acento significa "mira aquí". Si todo está resaltado, nada lo está:
  **un resalte por diapositiva** (una barra del gráfico, una palabra del título,
  una tarjeta).
- Rojo/verde solo con significado (negativo/positivo) y nunca como única señal
  (acompaña con ▲▼ o signo).
- Contraste mínimo WCAG AA: 4,5:1 para texto normal. Gris claro sobre blanco
  para texto de contenido = ilegible en proyector.
- Fondos oscuros para portada, separadores y cierre (dan ritmo); fondos claros
  para contenido (salvo tema oscuro completo como `tecnologia`).

## 4. Espacio en blanco

- El espacio vacío es un elemento de diseño: separa ideas y dirige la mirada.
- Una diapositiva al 60–70 % de ocupación se percibe "profesional"; al 95 %, "abarrotada".
- Huecos **intencionados** (márgenes, separación entre bloques) sí; huecos
  **accidentales** (media diapositiva vacía bajo tres viñetas cortas) no:
  cambia de layout (p. ej. `bullets` con `callout`, `cards`, `statement`).

## 5. Consistencia

- Mismos márgenes, posición de título, tamaños y colores en todas las diapositivas.
- Mismo formato numérico en todo el deck (separadores, decimales, unidades, "M€" vs "millones").
- Mismo estilo de viñeta, mismos radios y grosores de línea.
- Los elementos que se repiten (agenda con sección activa, etiqueta de sección)
  deben estar en el mismo lugar exacto.

## 6. Imágenes e iconos

- Solo imágenes que aporten información (producto, lugar, persona real del
  proyecto, captura de pantalla). Nada de fotos de stock genéricas de "apretón de manos".
- Imágenes a sangre (tocando el borde) o en un marco claro; nunca flotando con
  bordes irregulares. Recorta, no deformes (el layout `image` recorta solo).
- Iconos: mejor números o iniciales en círculos (como hace `cards`) que iconos
  decorativos sin significado. No uses emojis en decks profesionales.
- No uses fotografías de personas reales sin permiso ni imágenes con copyright.

## 7. Lo que delata un deck amateur (evitar)

- Títulos-etiqueta ("Introducción", "Datos", "Conclusiones") en diapositivas de contenido.
- Muros de viñetas; viñetas de 3+ líneas; párrafos completos.
- Más de 2 fuentes; tamaños aleatorios; texto < 12 pt para "que quepa".
- WordArt, sombras pesadas, degradados chillones, biseles, efectos 3D, gráficos 3D.
- Gráficos con leyenda innecesaria, rejilla densa, ejes decorativos, colores por defecto.
- Logos en todas las diapositivas cuando ya hay plantilla.
- Animaciones y transiciones llamativas.
- Elementos casi alineados; márgenes distintos en cada diapositiva.
- "Gracias" o "¿Preguntas?" como cierre sin próximos pasos.
- Líneas decorativas bajo cada título (recurso manido; el motor no las usa).
