# Cuadrante UCI

Aplicación web para generar cuadrantes de rotación de enfermería en una unidad
de críticos (UCI / SEMIS / REA / UCCRI), a partir de la planilla de turnos del mes.

Reparte al personal entre las 15 posiciones físicas de la unidad aplicando
clasificación de turnos, criterios clínicos, prioridad por antigüedad y suplencia,
cobertura de decalajes y rotación histórica mensual. Exporta el resultado a un
archivo Excel con el mismo código de colores del cuadrante, listo para subir a
Google Sheets.

## Cómo se usa

1. **Subir la planilla del mes** (el `.xlsm` que exporta la coordinadora).
2. **Cargar la base de profesionales** una vez (turno habitual, criterio UCI,
   antigüedad, suplencia y apodos). Se puede editar después desde la propia app.
3. **Elegir el periodo** (lunes de inicio y número de semanas) y generar.
4. **Exportar a Google Sheets**: descarga un `.xlsx` con colores y formato;
   se sube a Google Drive y se abre con Hojas de cálculo de Google.

## Desarrollo local (opcional)

Requiere Node.js instalado.

```bash
npm install      # instala dependencias
npm run dev      # arranca en local (http://localhost:5173)
npm run build    # genera la versión de producción en dist/
```

## Publicación

El proyecto está preparado para publicarse en Vercel conectando este
repositorio de GitHub. Vercel detecta Vite automáticamente; no requiere
configuración adicional.

## Tecnología

- React + Vite (interfaz)
- SheetJS (lectura de la planilla Excel)
- El motor de asignación es JavaScript puro, sin dependencias externas,
  portado fielmente desde la lógica VBA original.

Los datos (base de profesionales e histórico de rotación) se guardan
actualmente en el navegador. La migración a una base de datos compartida
(Supabase) con inicio de sesión está prevista como siguiente fase.
