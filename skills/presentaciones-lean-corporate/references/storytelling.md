# Storytelling: qué decir y en qué orden

Una presentación profesional se juzga primero por su lógica y después por su
aspecto. Un diseño impecable no rescata una historia confusa; una historia clara
se entiende incluso con un diseño sobrio.

## 1. Empieza por el final: el mensaje principal

Antes de pensar en diapositivas, escribe en **una frase** lo que la audiencia
debe recordar o decidir al salir. Si no cabe en una frase, aún no está claro.

- Malo: "Presentar la situación de atención al cliente".
- Bueno: "Debemos invertir 6,8 M€ en IA para atención al cliente porque el modelo
  actual no escala y el retorno es positivo desde el segundo año".

## 2. Estructuras narrativas

### SCQA (Situación – Complicación – Pregunta – Respuesta)
La más útil para decisiones y propuestas.

| Bloque | Contenido | Diapositivas típicas |
|---|---|---|
| Situación | Lo que la audiencia ya sabe y acepta | 1–2 (`kpis`, `chart`) |
| Complicación | Lo que ha cambiado o falla | 1–3 (`chart`, `big-number`) |
| Pregunta | La pregunta que surge (a menudo implícita) | 0–1 (`statement`) |
| Respuesta | Nuestra propuesta y su justificación | el resto |

### Pirámide (respuesta primero)
Para audiencias ejecutivas con poco tiempo: la recomendación en la diapositiva 2,
luego 3–4 argumentos que la sostienen, cada uno con su evidencia. Permite que
el comité "corte" en cualquier momento y ya tenga lo esencial.

### Problema → Solución → Beneficio → Prueba → Llamada a la acción
Para propuestas comerciales y pitch.

### Cronológica (antes → ahora → después)
Para informes de seguimiento, lecciones aprendidas o hojas de ruta.

## 3. Titulares-acción

Cada título es una **afirmación completa** que resume la conclusión de la
diapositiva. Leyendo solo los títulos, en orden, se entiende toda la historia.

| Título-etiqueta (evitar) | Título-acción (usar) |
|---|---|
| Ventas por canal | El canal online ya supone el 41 % de las ventas y crece 3× más rápido |
| Análisis de costes | El 70 % del sobrecoste se concentra en dos proveedores |
| Encuesta de clientes | Los clientes valoran más resolver a la primera que la rapidez |
| Próximos pasos | Necesitamos aprobar el presupuesto hoy para arrancar en noviembre |

Reglas:
- 1–2 líneas (≤ ~15 palabras). Si necesitas más, la diapositiva tiene dos mensajes.
- Contiene el "¿y qué?": una cifra, una comparación, una implicación.
- Mismo tiempo verbal y tono en todo el deck.
- Excepciones legítimas: portada, agenda, separadores de sección, cierre.

**Prueba del guion**: escribe todos los títulos en una lista. ¿Se sostiene la
historia sin ver el contenido? ¿Hay saltos lógicos? ¿Alguno se repite?

## 4. Una diapositiva = un mensaje + su prueba

El cuerpo de la diapositiva solo existe para **demostrar el título**. Pregúntate
para cada elemento: ¿apoya el titular? Si no, sobra (o va a otra diapositiva o a anexo).

Tipos de prueba, de más a menos persuasiva: dato cuantitativo con fuente →
comparación → ejemplo concreto o caso → cita de cliente/experto → argumento lógico.

## 5. Densidad según el uso

| Uso | Palabras por diapositiva | Notas |
|---|---|---|
| Presentación en directo | 15–50 | El orador cuenta; la diapositiva muestra. Notas del orador completas |
| Documento para leer ("pre-read") | 50–90 | Titulares completos, `takeaway` y fuentes en todas |
| Anexo / backup | hasta 120 | Tablas y detalle; nunca en el cuerpo principal |

## 6. Longitud y estructura por tipo de deck

| Tipo | Diapositivas | Estructura recomendada |
|---|---|---|
| Comité de dirección / decisión | 8–15 | Portada · Resumen ejecutivo (`statement` o `bullets`) · Situación · Complicación · Opciones (`comparison`) · Recomendación · Impacto (`big-number`/`table`) · Hoja de ruta · Riesgos · Decisión solicitada (`statement`) · Cierre |
| Propuesta comercial | 10–18 | Portada · Entendimiento del reto · Nuestro enfoque (`process`) · Solución (`cards`) · Casos de éxito (`quote`, `kpis`) · Equipo · Plan (`timeline`) · Inversión (`table`) · Por qué nosotros · Próximos pasos |
| Resultados periódicos | 8–12 | Portada · Resumen (`kpis`) · Evolución (`chart`) · Desviaciones y causas · Acciones · Previsión · Cierre |
| Kickoff de proyecto | 10–15 | Portada · Objetivos · Alcance (`two-column`: dentro/fuera) · Enfoque (`process`) · Plan (`timeline`) · Gobierno y equipo (`cards`) · Riesgos · Próximos pasos |
| Formación / divulgación | 12–25 | Portada · Agenda · Por módulo: `section` + 3–5 diapositivas variadas · Resumen · Recursos |
| Pitch / keynote | 10–15 | Gancho (`statement`/`big-number`) · Problema · Solución · Por qué ahora · Tracción (`kpis`) · Modelo · Equipo · Petición |

## 7. Apertura y cierre

- **Apertura**: tras la portada, entra en materia en ≤ 2 diapositivas. Para
  ejecutivos, el resumen ejecutivo va en la 2.
- **Cierre**: termina con lo que se pide a la audiencia (decisión, próximos pasos
  con responsable y fecha). Evita cerrar solo con "Gracias" o "¿Preguntas?".

## 8. Notas del orador

Escríbelas siempre en decks para presentar. Contenido útil: la frase de
transición desde la diapositiva anterior, el dato o anécdota que no está escrito,
la pregunta difícil esperable y su respuesta. 2–5 frases por diapositiva.

## 9. Datos que faltan

Nunca inventes cifras, nombres ni citas. Si el usuario no las aporta:
- Usa el dato que sí tengas y explica el supuesto en `notes`.
- O deja un valor claramente pendiente (p. ej. `"value": "[X] M€"`) y dilo al entregar.
  `qa_check.py` no marca corchetes en cifras, así que menciónalo tú.
