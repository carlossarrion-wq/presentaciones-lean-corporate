# Control de calidad antes de entregar

El primer render casi nunca es el bueno. Revisa con la actitud de buscar
errores, no de confirmar que todo está bien.

## 1. Automático

```bash
python scripts/qa_check.py salida.pptx --no-info               # deck para presentar
python scripts/qa_check.py salida.pptx --no-info --documento   # oferta / informe para leer (umbral 170 palabras)
```
Las comprobaciones de tamaño de letra y posición tienen en cuenta la escala de la
plantilla (Lean = ×2), así que los umbrales equivalen siempre a un 16:9 estándar.

| Mensaje | Qué hacer |
|---|---|
| `ERROR` desbordamiento | Acorta el texto o divide la diapositiva. No bajes la letra por debajo de 12 pt |
| `ERROR` fuera del lienzo | Revisa coordenadas (diapositivas a medida) |
| `ERROR` texto de relleno / marcador | Sustituye "[Imagen…]", "TODO", "Lorem"… o elimina el elemento |
| `AVISO` diapositiva densa / > 6 viñetas | Divide, agrupa o cambia de layout |
| `AVISO` cajas solapadas | Mira el render: si es intencionado (texto sobre forma) ignóralo |
| `AVISO` > 3 fuentes | Unifica |
| `INFO` sin notas | Añádelas si el deck se va a presentar |

Los avisos de `build_deck.py` (`AVISO: … posible desbordamiento`) se tratan igual que los errores.

## 2. Visual (obligatorio)

```bash
python scripts/render_preview.py salida.pptx --out preview --dpi 40          # plantilla Lean (lienzo ×2)
python scripts/render_preview.py salida.pptx --out zoom --dpi 60 --slides 4,7   # detalle
```

Mira `preview/contact_sheet.png` para el conjunto y abre a tamaño completo
cualquier diapositiva con contenido denso (tablas, gráficos, comparativas).

### Diapositiva a diapositiva
- [ ] Ningún texto cortado, desbordado o superpuesto a otra forma.
- [ ] Nada pegado a los bordes; márgenes iguales a los de las demás diapositivas.
- [ ] El título cabe en ≤ 2 líneas y es una afirmación.
- [ ] Hay un único foco visual (el resalte va a lo que dice el título).
- [ ] No hay huecos accidentales grandes (media diapositiva vacía).
- [ ] Gráficos: etiquetas legibles, sin solapes, unidad indicada, fuente presente.
- [ ] Tablas: números alineados a la derecha, mismos decimales, cabecera legible.
- [ ] Contraste suficiente (texto gris claro sobre fondo claro = mal).
- [ ] Imágenes sin deformar ni pixeladas; ningún marcador gris pendiente.

### Específico de la plantilla Lean
Ver la lista de verificación de `lean-corporate.md` (logotipo, secciones, importes, firmantes).

### Deck completo
- [ ] Leyendo solo los títulos se entiende la historia, sin saltos ni repeticiones.
- [ ] Hay variedad de layouts (no 5 diapositivas de viñetas seguidas).
- [ ] Formato numérico y terminología coherentes en todo el deck.
- [ ] Las cifras cuadran entre diapositivas (el total del resumen = el de la tabla).
- [ ] Ortografía y acentos revisados; mayúsculas coherentes en títulos (solo la inicial).
- [ ] El cierre pide algo concreto (decisión, próximos pasos con fecha/responsable).
- [ ] Notas del orador presentes si el deck se presenta en directo.

## 3. Ciclo de corrección

1. Anota todos los problemas de una pasada.
2. Corrige en el JSON (no en el .pptx generado, que se sobrescribe).
3. Reconstruye, vuelve a ejecutar QA y vuelve a renderizar **solo** lo que cambió (`--slides`).
4. Repite hasta que una pasada completa no encuentre problemas nuevos.

## 4. Al entregar

- Nombre de archivo descriptivo (`Plan_Atencion_Cliente_2027_Comite.pptx`).
- Mensaje breve: qué es, supuestos tomados, qué queda pendiente (datos,
  imágenes, validaciones). No describas cada diapositiva.
