# FASE B — Afinar la lógica de clasificación y rotación

> Documento de trabajo acordado con Danié. Recoge TODO lo que hay que
> implementar en la Fase B, para no perder detalle entre fases.
> La Fase A (estructura configurable de la unidad) se hace primero.

## ESTADO
- [HECHO] B1 — Clasificación por turno natural + cómputo de horas
- [HECHO] B2 — Detección y etiquetado del doblaje (badges DBL/+/½ en web y Excel)
- [PENDIENTE] B3 — Anotaciones manuales sobre el cuadrante (COT/COF/franjas)
- [PENDIENTE] B4 — Afinar continuidad/rotación + registro mensual estilo ROTACIONS_INF

---

## B1. Clasificación por TURNO NATURAL (el cambio de fondo)

El problema actual: el motor clasifica cada jornada solo por su forma horaria,
y confunde jornadas normales mal escritas con doblajes.

**La regla correcta:** partir del turno natural de cada profesional (casilla de la
base de datos) y entender lo que aparece en el Excel como modificaciones sobre esa
referencia. El dato decisivo es el **cómputo TOTAL de horas del día**, no la forma.

Procedimiento por profesional y día:
1. Leer su **turno natural** de la base: `TM` (7h mañana), `TT` (7h tarde),
   `TN` (10h noche), `12H Día`, `12H Noche`.
2. Sumar el **total de horas** del día juntando TODOS los tramos en que aparezca
   (puede estar partido en 2+ líneas / secciones).
3. Comparar total vs natural:
   - Total ≈ natural → **jornada normal** (aunque venga partida en tramos).
     Ej.: Irene Moreno es 12H Día; aparece `8-14:20` + `14:20-20` = 12h → 12H Día normal,
     NO es doblaje.
   - Total claramente MAYOR que el natural → algo excepcional:
     - 7h (TM/TT) que hace mañana Y tarde (~14h) → **DOBLAJE**.
     - jornada natural + tramo extra de 7h → concepto **"+"** (turno extra).
   - Total MENOR que el natural → franja parcial pactada (un 12h que solo hace `8-14`).
     En el Excel actual no aparece, pero la lógica debe contemplarlo.

### Rangos de horas acordados (con tolerancia)
- **7h**: entre 6 y 9 horas.
- **10h noche**: alrededor de 10h (rango razonable 9–11; ajustar con casos).
- **12h**: entre 11 y 13 horas.
- **Doblaje** (para un natural de 7h): por encima de ~13–14h.
- Estos umbrales van parametrizados en config para poder ajustarlos.

### Concepto "+" (turno extra)
- Significa **un turno extra** (habitualmente de 7h). Sin más dificultad.
- Puede darse de dos formas:
  - en el **mismo día** que ya trabaja → genera un doblaje (natural + extra).
  - en un **día que no trabajaba** → ese día pasa a trabajar.
- Para detectarlo: se nota en el cómputo de horas del día (sube respecto al natural).

---

## B2. Detección y etiquetado del DOBLAJE
- Un doblaje es un profesional de **7h** (natural TM o TT) que excepcionalmente hace
  mañana + tarde el mismo día (total ~14h, muy por encima de sus 7h).
- En el Excel de origen se manifiesta como la persona apareciendo en la sección de
  MAÑANA **y** en la de TARDE el mismo día, con dos tramos de 7h.
- Hay que **distinguirlo** del 12h natural (que también ocupa mañana+tarde) usando el
  turno natural + total de horas.
- En el cuadrante debe **etiquetarse como "doblaje"** (no como 12H Día).

---

## B3. Anotaciones manuales (lo que NO se puede detectar del Excel)
Verificado: en el Excel de origen NO aparecen estas marcas; las pone la coordinadora
a mano. El dato no existe en el origen, así que no son automatizables:
- **COT**: el profesional no trabajaba ese día y trabaja por un cambio. En el Excel
  simplemente aparece trabajando, sin marca.
- **COF**: el profesional libra (NO debe ponerse en el cronograma). En el Excel no
  aparece ese día.
- **Franjas parciales de 12h** ("8-14h", "14-20h"): un 12h que por un cambio pactado
  solo hace media jornada. En el Excel siempre figura como `8-20`.

**Enfoque acordado a confirmar:** permitir que la coordinadora **añada estas
anotaciones manualmente** sobre el cuadrante ya generado (como hace ahora en su Excel),
en celdas concretas. El sistema genera lo automático; ella apunta lo excepcional.

---

## B4. Continuidad asistencial y rotación mensual (afinado con datos reales)
De los 6 meses de cronogramas reales (`GUARDATS_PLANINGS_INF_2026`) se observó:
- La continuidad se prioriza **por unidad**: cuando un profesional trabaja varios días
  seguidos, se intenta mantenerlo en la **misma unidad**, y a ser posible misma posición.
- La **posición exacta cede** ante la realidad de quién está disponible cada día
  (muchos son 12h o libran días sueltos). No siempre se logra misma posición toda la semana.
- Ajustar la lógica de continuidad/rotación para parecerse a estos patrones reales.

### Registro de rotación mensual (`ROTACIONS_INF`)
- Formato real: una fila por profesional, una columna por mes, celda = unidad de ese mes.
- Separado en **dos hojas: DIA y NIT** (rotación de día y de noche por separado).
- El motor ya genera un "Informe mensual"; revisar para que case con este formato
  (separación día/noche, una unidad por profesional y mes).
- UCCRI esquerra/dreta cuentan como **"UCCRI"** a efectos de rotación (da igual lado).

---

## Notas de contexto (decididas en Fase A)
- La estructura de la unidad es **configurable por semana** y se guarda con casillas
  marcables que persisten para la siguiente vez (la coordinadora marca/desmarca).
- UCCRI puede tener 1 o 2 sub-unidades abiertas (esquerra/dreta) con nº variable de
  enfermeras cada una; a efectos de rotación todo es "UCCRI".
