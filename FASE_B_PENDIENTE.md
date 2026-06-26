# FASE B — Afinar la lógica de clasificación y rotación

> Documento de trabajo acordado con Danié. Recoge TODO lo que hay que
> implementar en la Fase B, para no perder detalle entre fases.
> La Fase A (estructura configurable de la unidad) se hace primero.

## ESTADO
- [HECHO] B1 — Clasificación por turno natural + cómputo de horas
- [HECHO] B2 — Detección y etiquetado del doblaje (badges DBL/+/½ en web y Excel)
- [HECHO] B3 — Anotaciones manuales sobre el cuadrante (COT/COF/franjas)
          Clic en celda → nota; se guarda por semana en config_unidad (Supabase).
          Pendiente opcional: incluir las notas también en el Excel exportado.
- [HECHO] B4 — Informe mensual separado DÍA/NOCHE (estilo ROTACIONS_INF) +
          continuidad intra-semana reforzada (repetir posición los días que trabaja).

FASE B COMPLETA.

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

---

# MEJORAS POSTERIORES (sesiones tras Fase B)

## HECHO en la última sesión
- **Histórico de rotación en apartado propio** (sidebar "Histórico de rotación"),
  vista visual con chips de color por unidad, regla del 66% (≥66% días en una
  unidad → solo esa; si no, las DOS con más días). Quitado de la generación de
  cuadrantes y eliminada la pestaña "Informe mensual".
- **Colores nuevos**: TODOS los de 12h (día y noche) y doblaje → LILA (#D2BEE6).
  Los de 7h, 10h noche y parcial → BLANCO. Sin cubrir → rojo. En web y Excel.
  (El usuario pidió explícitamente que 12h día también sea lila, NO verde.)
- **Filas en blanco entre unidades** y **etiquetas fijas de posición**:
  "+ revisió carro" en UCI_A pos2, SEMIS pos1, REA pos2, UCCRI_ESQ pos2,
  UCCRI_DRE pos2; "ATURADA+ICTUS+CARRO" reemplaza el número en UCI_B pos5.

## PENDIENTE — ROTACIÓN OBLIGATORIA (lo más importante a terminar)

### Reglas acordadas con Danié
1. Solo rotan los profesionales con **Suplente = No**. Los Suplente=Si NO rotan
   (dan flexibilidad para tapar huecos).
2. Cada fijo rota SOLO por las unidades que su **criterio** permite, de entre las
   ABIERTAS en la estructura del mes:
   - SI → UCI A, UCI B, SEMIS, REA, UCCRI
   - NO → SEMIS, UCCRI
   - UCI → UCI A, UCI B
   - UCCRI → UCCRI (única → se queda siempre)
3. **No repite unidad hasta pasar por todas las suyas**; al completar el ciclo, reinicia.
4. La rotación de NOCHE es el mismo concepto que la de día (mismas unidades).
5. El orden del ciclo es libre (el que mejor cubra).
6. Si hay que dejar ubicaciones SIN CUBRIR, el orden de sacrificio es:
   **UCI B pos5 → UCI B pos4** (máx 3 huecos en UCI B) → **UCI A pos3** →
   **SEMIS pos3** → y por último **UCCRI**.
7. El histórico se puede reiniciar las veces necesarias hasta que funcione.

### Estado actual del código (a medio camino)
- Implementadas y EXPORTADAS para test: `unidadRotacionObligatoria(historico,
  nombre, criterio, esSuplente, fechaLunes)`, `unidadesCandidatasProf(criterio)`,
  `historialMensualProf`, `unidadesAbiertasRotacion`. La función AISLADA funciona
  perfecta en todos los casos (probado: tras UCI A va a UCI B; ciclo completo
  reinicia; criterio UCI alterna A↔B; UCCRI se queda; suplente no rota).
- Integrada en `intentarAsignarPos` (asignación estable) y en
  `colocarEnDiaTurno`/`rellenarHuecosYExcedentes` (relleno de huecos), pasando
  historico+fechaLunes. Con esto, los que cambian de unidad pasaron de 19 a 39.

### EL PROBLEMA que falta resolver
- Todavía ~33 de ~72 fijos REPITEN unidad. Causa raíz: **problema de capacidad**.
  Si muchos estaban en UCI A en octubre, en noviembre TODOS calculan la misma
  "siguiente unidad" y saturan unidades pequeñas (SEMIS=3, REA=2). Los que no
  caben se quedan donde estaban.
- Los que repiten se concentran en SEMIS y REA (las pequeñas).

### LA SOLUCIÓN a construir (próxima sesión)
- Hace falta un **reparto inteligente / asignación global por mes**: en vez de que
  cada profesional calcule su unidad siguiente independientemente, repartir las
  PLAZAS de cada unidad entre los candidatos que deben rotar, priorizando a quien
  lleva más meses sin pisar esa unidad. Es un problema de asignación (tipo
  emparejamiento), no un retoque local.
- **PREGUNTA ABIERTA pendiente de responder por Danié**: cuando físicamente no
  caben todos los que deben rotar a una unidad pequeña, ¿Opción A (best-effort:
  repartir lo mejor posible y aceptar que algunos repitan ese mes, garantizando
  que con el tiempo todos roten) u Opción B (más agresivo, dejar a más gente en
  excedente para forzar que quepan los que rotan)? Claude recomienda la A.

### Backup
- Hay backup en src/App.jsx.bak (de antes de tocar la rotación). Quitar antes de
  empaquetar la versión final.

---

# SESIÓN: Mejoras de excedentes, rotación y vinculación (COMPLETADO)

## HECHO Y VERIFICADO
1. **Celdas de excedente vacías → BLANCO** (no rojo). El rojo es solo para
   posiciones reales sin cubrir; los excedentes no son ubicaciones reales.
   Aplicado en web (.exc-cell:empty) y Excel (S.cBlanco).
2. **Orden de llenado de excedentes EQUILIBRADO**: el ciclo
   UCCRI→REA→SEMIS→UCI A→UCI B ahora reparte de forma pareja (una vuelta completa
   antes de poner un segundo en ninguna unidad), respetando criterios. Reescrito
   `repartirExcedentes` con enfoque "menos cargada primero". Antes amontonaba en UCCRI.
3. **Rotación MÁXIMO 2 meses seguidos (regla DURA)**: en `repartirRotacionBloque` y
   en `unidadRotacionObligatoria` se calcula la unidad "prohibida" (la que sería el
   3er mes consecutivo) y nunca se asigna. Si por capacidad alguien no cabe en otra
   unidad, queda sin asignar y fluye a excedente en su unidad de rotación (que sí
   cumple la regla), en vez de forzar el 3er mes. Verificado con los 3 meses reales:
   ningún fijo con varias unidades repite 3 meses. (Suplentes y criterio único pueden
   repetir, es correcto.)
4. **Modal de vinculación de profesionales no reconocidos**: al generar, si hay
   nombres en la planilla que no se vinculan con la BD, aparece la vista "vincular"
   (componente `VincularProfesionales`) ANTES de terminar. Por cada nombre se puede
   "Vincular a existente" (crea alias) o "Añadir como nuevo" (a la BD). Tras aplicar,
   continúa la generación. Funciones motor: `db.estaEnBD()`,
   `extraerNombresSinVincular()`. Alias se guardan en config_unidad
   (clave "alias_profesionales") vía `cargarAlias`/`guardarAliasTodos`; se cargan al
   inicio y se MEZCLAN (no reemplazan) cuando se sube una planilla con alias
   embebidos (`mezclarAlias`).

## Datos de prueba creados
- testdata_octubre.json, testdata_noviembre.json, testdata_diciembre.json
  (convertidos de las 3 planillas reales del proyecto). Permiten probar rotación
  real de 3 meses encadenando el histórico.

## Persistencia
- NO requiere SQL nuevo en Supabase: los alias reutilizan la tabla config_unidad.

## Nota sobre un caso límite (no es bug)
- Profesionales SIN posición estable (p.ej. de noche que tapan huecos) pueden
  aparecer dispersos por varias unidades el mismo mes; su "unidad dominante" del
  66% puede coincidir 2-3 meses por azar aunque realmente rotan. No se fuerza más
  porque distorsionaría el reparto del resto.
