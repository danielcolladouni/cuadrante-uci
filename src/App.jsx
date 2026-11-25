import React, { useState, useEffect, useRef, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import {
  iniciarSesion, cerrarSesion, sesionActual, onCambioSesion,
  cambiarPassword, recuperarPassword, apodoDeSesion, cambiarApodo, escucharRecuperacion,
  cargarProfesionales, guardarProfesionales,
  cargarHistorico, añadirAlHistorico, vaciarHistorico, borrarHistoricoDeMeses, borrarHistoricoDeRango,
  cargarEstructura, guardarEstructura,
  cargarAnotaciones, guardarAnotaciones,
  cargarAlias, guardarAliasTodos,
  cargarReglas, guardarReglas,
  fijarColectivo,
  cargarEstudiantes, guardarEstudiante, borrarEstudiante, importarEstudiantes, borrarEstudiantesCalendario,
} from "./supabase.js";

/* ============================================================
   MOTOR — embebido (port fiel del VBA). Aislado en un closure;
   sólo se exponen las funciones que la interfaz necesita.
   ============================================================ */
const ENGINE = (function () {
/* ===== MOTOR DE ROTACIÓN (browser bundle) ===== */
/* ============================================================
   MOTOR DE ROTACIÓN DE ENFERMERÍA  —  port fiel del VBA (10 módulos)
   Sin dependencias. Funciona en navegador y en Node.
   La E/S de Excel se hace fuera; aquí sólo lógica pura.
   ============================================================ */

/* ---------- M00_Config: constantes y helpers de texto ---------- */
const TURNO = {
  MANANA: "MAÑANA", TARDE: "TARDE", NOCHE: "NOCHE",
  D12: "12H DÍA", N12: "12H NOCHE", DOBLAJE: "DOBLAJE", PARCIAL: "PARCIAL",
};

const COLOR = {
  MANANA: "#FFE699", TARDE: "#F8CBAD", NOCHE: "#B4A7D6",
  D12: "#A9D08E", N12: "#9DC3E6", DOBLAJE: "#F4B084", PARCIAL: "#D9D9D9",
  BLANCO: "#FFFFFF",
};

function normalizarTexto(s) {
  if (s == null) return "";
  let r = String(s).toUpperCase().trim();
  r = r.replace(/,/g, " ").replace(/\./g, " ");
  const map = { "Á":"A","É":"E","Í":"I","Ó":"O","Ú":"U","À":"A","È":"E","Ì":"I","Ò":"O","Ù":"U","Ü":"U" };
  r = r.replace(/[ÁÉÍÓÚÀÈÌÒÙÜ]/g, c => map[c] || c);
  while (r.indexOf("  ") >= 0) r = r.replace(/  /g, " ");
  return r.trim();
}

function colorPorTurno(t) {
  switch (t) {
    case TURNO.MANANA: return COLOR.MANANA;
    case TURNO.TARDE: return COLOR.TARDE;
    case TURNO.NOCHE: return COLOR.NOCHE;
    case TURNO.D12: return COLOR.D12;
    case TURNO.N12: return COLOR.N12;
    case TURNO.DOBLAJE: return COLOR.DOBLAJE;
    case TURNO.PARCIAL: return COLOR.PARCIAL;
    default: return COLOR.BLANCO;
  }
}

const DIAS_ABBR = ["Dom","Lun","Mar","Mié","Jue","Vie","Sáb"];
const MESES_ABBR = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
function fechaCorta(f) { // f = Date
  return DIAS_ABBR[f.getDay()] + " " + f.getDate() + MESES_ABBR[f.getMonth()];
}
// clave yyyy-mm en local
function mesKey(f) {
  return f.getFullYear() + "-" + String(f.getMonth() + 1).padStart(2, "0");
}
// días entre dos fechas (a nivel de día, local)
function diaSerial(f) {
  return Math.floor((Date.UTC(f.getFullYear(), f.getMonth(), f.getDate())) / 86400000);
}
function addDays(f, n) {
  const d = new Date(f.getFullYear(), f.getMonth(), f.getDate());
  d.setDate(d.getDate() + n);
  return d;
}
// weekday Monday=1..Sunday=7
function weekdayMonday(f) {
  const wd = f.getDay(); // 0=Sun
  return wd === 0 ? 7 : wd;
}

/* ---------- M02_AreasConfig: áreas, mapeo y permisos (CONFIGURABLE) ---------- */
const AREA = { UCI_A:"UCI A", UCI_B:"UCI B", SEMIS:"SEMIS", REA:"REA", UCCRI:"UCCRI" };
const CRIT = { SI:"SI", NO:"NO", UCI:"UCI", UCCRI:"UCCRI" };

// Catálogo de áreas posibles de la unidad real. La coordinadora activa/desactiva
// y ajusta el nº de posiciones cada semana. UCCRI Esquerra y Dreta comparten
// rotación "UCCRI" (a efectos de rotación mensual da igual el lado).
const CATALOGO_AREAS = [
  { id: "UCI_A",     nombre: "CONTROL A",      rotacion: "UCI A",  posDefecto: 4 },
  { id: "UCI_B",     nombre: "CONTROL B",      rotacion: "UCI B",  posDefecto: 4 },
  { id: "SEMIS",     nombre: "SEMIS",          rotacion: "SEMIS",  posDefecto: 3 },
  { id: "REA",       nombre: "REA",            rotacion: "REA",    posDefecto: 2 },
  { id: "UCCRI_ESQ", nombre: "UCCRI ESQUERRA", rotacion: "UCCRI",  posDefecto: 2 },
  { id: "UCCRI_DRE", nombre: "UCCRI DRETA",    rotacion: "UCCRI",  posDefecto: 2 },
];

function estructuraPorDefecto() {
  return CATALOGO_AREAS.map(a => ({
    id: a.id, nombre: a.nombre, rotacion: a.rotacion,
    posiciones: a.posDefecto, activa: true,
  }));
}

// Etiqueta fija POR DEFECTO (convención clásica de la unidad) que sustituye o
// acompaña al número de una posición. Se usa solo si el área no tiene etiquetas
// personalizadas guardadas en la estructura.
function etiquetaFijaPosicion(areaId, posLocal) {
  // "+ revisió carro" en posiciones concretas (acompaña al número)
  const revisioCarro = (
    (areaId === "UCI_A" && posLocal === 2) ||
    (areaId === "SEMIS" && posLocal === 1) ||
    (areaId === "REA" && posLocal === 2) ||
    (areaId === "UCCRI_ESQ" && posLocal === 2) ||
    (areaId === "UCCRI_DRE" && posLocal === 2)
  );
  if (revisioCarro) return { texto: "+ revisió carro", reemplazaNumero: false };
  // Posición 5 de UCI B: cama de parada / ictus / carro (sustituye al número)
  if (areaId === "UCI_B" && posLocal === 5) return { texto: "ATURADA+ICTUS+CARRO", reemplazaNumero: true };
  return null;
}

// Etiqueta EFECTIVA de una posición: si el área trae etiquetas personalizadas
// (campo `etiquetas` en la estructura: {posLocal: texto}), se usan esas; si no,
// las por defecto. Convenio del texto personalizado: si empieza por "=", el
// texto SUSTITUYE al número (p.ej. "=ATURADA+ICTUS+CARRO"); si no, lo acompaña.
// Devuelve { texto, reemplazaNumero } o null.
function etiquetaDePosicion(area, posLocal) {
  const ets = area && area.etiquetas;
  if (ets && Object.prototype.hasOwnProperty.call(ets, String(posLocal))) {
    const raw = String(ets[String(posLocal)] || "").trim();
    if (!raw) return null; // vacío explícito = sin etiqueta
    if (raw.startsWith("=")) return { texto: raw.slice(1).trim(), reemplazaNumero: true };
    return { texto: raw, reemplazaNumero: false };
  }
  return etiquetaFijaPosicion(area ? area.id : "", posLocal);
}

// Etiquetas por defecto de un área en formato editable {posLocal: texto},
// para precargar el editor (usa el convenio "=" para las que reemplazan).
function etiquetasPorDefectoDeArea(areaId, nPos) {
  const out = {};
  for (let p = 1; p <= nPos; p++) {
    const et = etiquetaFijaPosicion(areaId, p);
    out[String(p)] = et ? (et.reemplazaNumero ? "=" + et.texto : et.texto) : "";
  }
  return out;
}

// Compila una config en un mapa de posiciones globales que usa el motor.
function compilarEstructura(config) {
  const activas = (config && config.length ? config : estructuraPorDefecto())
    .filter(a => a.activa && a.posiciones > 0);
  const areas = [];
  let cursor = 1;
  for (const a of activas) {
    const ini = cursor, fin = cursor + a.posiciones - 1;
    areas.push({ id: a.id, nombre: a.nombre, rotacion: a.rotacion,
      posiciones: a.posiciones, posIni: ini, posFin: fin,
      etiquetas: a.etiquetas || null });
    cursor = fin + 1;
  }
  return { areas, nPosTotal: cursor - 1 };
}

// --- Estructura ACTIVA (variable de módulo). El motor la fija antes de generar. ---
let _EST = compilarEstructura(estructuraPorDefecto());
function fijarEstructura(config) { _EST = compilarEstructura(config); return _EST; }
function estructuraActiva() { return _EST; }
function NPOS_TOTAL_actual() { return _EST.nPosTotal; }

// Lista de "áreas" en el sentido antiguo = rotaciones únicas, en orden de recorrido.
function listaAreasOrdenadas() {
  const out = [];
  for (const a of _EST.areas) if (!out.includes(a.rotacion)) out.push(a.rotacion);
  return out;
}

// nº de posiciones de una ROTACIÓN (sumando sus áreas: p.ej. UCCRI esq+dre)
function posicionesDeArea(rotacion) {
  let n = 0;
  for (const a of _EST.areas) if (a.rotacion === rotacion) n += a.posiciones;
  return n;
}

// Dada una posición global, su rotación y posición local DENTRO de la rotación.
function posGlobalAArea(p) {
  // posLocal relativo a la rotación (para que esquerra+dreta numeren correlativo
  // a efectos internos), y también guardamos el área física.
  let acumRot = {};
  for (const a of _EST.areas) {
    if (p >= a.posIni && p <= a.posFin) {
      const base = acumRot[a.rotacion] || 0;
      return { area: a.rotacion, posLocal: base + (p - a.posIni + 1), areaId: a.id, areaNombre: a.nombre, posLocalArea: p - a.posIni + 1 };
    }
    acumRot[a.rotacion] = (acumRot[a.rotacion] || 0) + a.posiciones;
  }
  return { area: "", posLocal: 0, areaId: "", areaNombre: "", posLocalArea: 0 };
}

// Primera posición global de una rotación (su primera área activa).
function posGlobalInicioArea(rotacion) {
  for (const a of _EST.areas) if (a.rotacion === rotacion) return a.posIni;
  return 0;
}
// Última posición global de una rotación (su última área activa).
function posGlobalFinArea(rotacion) {
  let fin = 0;
  for (const a of _EST.areas) if (a.rotacion === rotacion) fin = a.posFin;
  return fin;
}
function areaAPosGlobal(rotacion, posLocal) {
  return posGlobalInicioArea(rotacion) + posLocal - 1;
}

function puedeIrAArea(criterio, rotacion) {
  const c = String(criterio || "").toUpperCase().trim();
  switch (c) {
    case CRIT.SI: return true;
    case CRIT.NO: return rotacion === AREA.SEMIS || rotacion === AREA.UCCRI;
    case CRIT.UCI: return rotacion === AREA.UCI_A || rotacion === AREA.UCI_B;
    case CRIT.UCCRI: return rotacion === AREA.UCCRI;
    default: return true; // conservador
  }
}

/* ---------- REGLAS DE ROTACIÓN configurables ----------
   Cada regla puede activarse o desactivarse desde la web. El motor las consulta
   en los puntos correspondientes. Por defecto, TODAS activas (comportamiento actual).
*/
const REGLAS_DEFECTO = {
  rotacionMensual: true,          // asignar unidad del mes a cada fijo (rotación obligatoria)
  max2Meses: true,                // nunca 3 meses seguidos en la misma unidad
  cicloCompleto: true,            // no repetir unidad hasta pasar por todas las permitidas
  suplentesNoRotan: true,         // los suplentes no entran en la rotación (bolsa flexible)
  continuidadMes: true,           // mantener la misma unidad durante todo el mes
  evitarPosicionSemanaAnterior: true, // dentro de la unidad, no repetir posición de la semana previa
  excedentesEquilibrados: true,   // repartir excedentes equilibrando el ciclo UCCRI→REA→SEMIS→UCI A→UCI B
  excedentesPrioridadSuplentes: true, // a excedente van antes los suplentes y los fijos menos antiguos
  antiguosPosicionFija: true,     // los marcados "Antiguo/a" no se mueven de sitio en toda la semana
};
// Descripciones para la interfaz (título + explicación en lenguaje llano).
const REGLAS_INFO = [
  { clave: "rotacionMensual", titulo: "Rotación mensual obligatoria",
    desc: "Cada mes se asigna a cada profesional fijo una unidad de rotación, repartiendo las plazas para que roten el máximo posible. Si se desactiva, cada uno se coloca donde haya hueco sin obligación de cambiar de unidad.", dependeDe: null },
  { clave: "max2Meses", titulo: "Máximo 2 meses seguidos en la misma unidad",
    desc: "Regla dura: nadie puede estar 3 meses consecutivos en la misma unidad. Antes de permitirlo, se le deja como excedente en otra unidad.", dependeDe: "rotacionMensual" },
  { clave: "cicloCompleto", titulo: "Completar el ciclo antes de repetir",
    desc: "No se repite una unidad hasta haber pasado por todas las que el criterio permite; al completarlas, el ciclo se reinicia.", dependeDe: "rotacionMensual" },
  { clave: "suplentesNoRotan", titulo: "Los suplentes no rotan",
    desc: "El personal suplente queda fuera de la rotación y actúa como bolsa flexible para cubrir huecos. Si se desactiva, los suplentes también reciben unidad de rotación.", dependeDe: "rotacionMensual" },
  { clave: "continuidadMes", titulo: "Misma unidad durante todo el mes",
    desc: "Una vez asignada la unidad del mes, se mantiene en todas las semanas de ese mes (no se cambia de unidad a mitad de mes).", dependeDe: null },
  { clave: "evitarPosicionSemanaAnterior", titulo: "No repetir posición de la semana anterior",
    desc: "Dentro de la misma unidad, se intenta no asignar exactamente la misma posición (cama) que la semana pasada.", dependeDe: null },
  { clave: "excedentesEquilibrados", titulo: "Excedentes repartidos en ciclo equilibrado",
    desc: "Cuando sobra personal, los excedentes se reparten completando una vuelta del ciclo UCCRI → REA → SEMIS → UCI A → UCI B antes de poner un segundo en ninguna unidad. Si se desactiva, se reparten en orden simple.", dependeDe: null },
  { clave: "antiguosPosicionFija", titulo: "Antiguos/as con posición fija",
    desc: "Las personas marcadas como «Antiguo/a» en la plantilla tienen prioridad absoluta: no cambian de sitio en toda la semana, se quedan en la misma unidad todo el mes, y cada semana avanzan una posición dentro de esa unidad (al llegar a la última, vuelven a la primera). Al mes siguiente rotan de unidad como todo el mundo.", dependeDe: null },
  { clave: "excedentesPrioridadSuplentes", titulo: "A excedente: primero suplentes, luego por antigüedad",
    desc: "Cuando sobra gente, van a posiciones de excedente primero los suplentes y después los fijos de menor antigüedad; los más antiguos conservan posición real.", dependeDe: null },
];
// Reglas ACTIVAS (variable de módulo). El motor las fija antes de generar.
let _REGLAS = { ...REGLAS_DEFECTO };
function fijarReglas(r) { _REGLAS = { ...REGLAS_DEFECTO, ...(r || {}) }; return _REGLAS; }
function reglasActivas() { return _REGLAS; }
function normalizarCriterio(valor) {
  let v = String(valor || "").toUpperCase().trim();
  if (v === "SÍ") v = "SI";
  if (v === "SI" || v === "NO" || v === "UCI" || v === "UCCRI") return v;
  return "SI";
}
function areaExcedenteRoundRobin(n) {
  const areas = listaAreasOrdenadas();
  if (areas.length === 0) return "";
  return areas[(n - 1) % areas.length];
}
// Orden de reparto de EXCEDENTES (sobra gente): UCCRI → REA → SEMIS → UCI A → UCI B.
// Solo unidades abiertas. (Distinto del orden de sacrificio de huecos.)
function ordenExcedentes() {
  const pref = ["UCCRI", "REA", "SEMIS", "UCI A", "UCI B"];
  const abiertas = unidadesAbiertasRotacion();
  const clasicas = pref.filter(u => abiertas.includes(u));
  // Rotaciones nuevas (creadas por el usuario) van al final del ciclo.
  const nuevas = abiertas.filter(u => !pref.includes(u));
  return clasicas.concat(nuevas);
}
// Asigna a una lista de excedentes su unidad siguiendo el ciclo
// UCCRI → REA → SEMIS → UCI A → UCI B, repartiendo de forma EQUILIBRADA: cada
// nuevo excedente va a la unidad (en orden de ciclo) que pueda y que menos
// excedentes tenga hasta ahora, de modo que se complete una vuelta del ciclo
// antes de poner un segundo en ninguna unidad. Respeta el criterio de acceso.
// excedentes: [{nombre}]. Devuelve [{nombre, area, ordenEnUnidad}].
function repartirExcedentes(lista, db) {
  const orden = ordenExcedentes();
  const out = [];
  const cont = {};
  for (const u of orden) cont[u] = 0;
  let punteroSimple = 0;
  for (const ex of lista) {
    const crit = db.criterio(ex.nombre);
    const posibles = orden.filter(u => puedeIrAArea(crit, u));
    if (posibles.length === 0) { out.push({ nombre: ex.nombre, area: orden[0] || "", ordenEnUnidad: 1 }); continue; }
    let mejor;
    if (_REGLAS.excedentesEquilibrados) {
      // elegir la unidad posible con menos excedentes; empate -> primera en el ciclo
      mejor = posibles[0];
      for (const u of posibles) if (cont[u] < cont[mejor]) mejor = u;
    } else {
      // reparto simple: siguiente unidad del ciclo que este profesional pueda
      mejor = posibles[0];
      for (let i = 0; i < orden.length; i++) {
        const u = orden[(punteroSimple + i) % orden.length];
        if (posibles.includes(u)) { mejor = u; punteroSimple = (punteroSimple + i + 1) % orden.length; break; }
      }
    }
    cont[mejor]++;
    out.push({ nombre: ex.nombre, area: mejor, ordenEnUnidad: cont[mejor] });
  }
  return out;
}
function numeroLocalDePosGlobal(p) { return String(posGlobalAArea(p).posLocalArea); }
function nombreAreaCortoDePosGlobal(p) {
  const r = posGlobalAArea(p);
  return r.areaId || "";
}

/* ---------- M01_Extractor: parseo y clasificación de horarios ---------- */
function parsearHoraAMinutos(s) {
  let t = String(s || "").trim();
  if (t.length === 0) return -1;
  if (t.indexOf(":") >= 0) {
    const p = t.split(":");
    if (p.length !== 2) return -1;
    if (isNaN(+p[0]) || isNaN(+p[1])) return -1;
    return (+p[0]) * 60 + (+p[1]);
  }
  if (isNaN(+t)) return -1;
  return (+t) * 60;
}

// Parsea "8-14 MMORALES" / "* 8-20 APELLIDO APELLIDO, NOMBRE" -> {horario, nombre}
function parsearLineaPlanilla(linea) {
  let t = String(linea || "").trim();
  let asterisco = "";
  if (t.charAt(0) === "*") { asterisco = "* "; t = t.substring(1).trim(); }
  const pos = t.indexOf("-");
  if (pos < 0) return { horario: "", nombre: t };
  const finHora = t.indexOf(" ", pos);
  if (finHora < 0) return { horario: asterisco + t, nombre: "" };
  return {
    horario: asterisco + t.substring(0, finHora).trim(),
    nombre: t.substring(finHora + 1).trim(),
  };
}

// Clasifica un horario individual -> DOBLE/D12/N12/N10/M/T/N/IGN/?
function clasificarHorarioIndividual(horario, seccion) {
  let h = String(horario || "").trim();
  const tieneAsterisco = h.charAt(0) === "*";
  if (tieneAsterisco) h = h.substring(1).trim();
  const partes = h.split("-");
  if (partes.length !== 2) return "?";
  const ini = parsearHoraAMinutos(partes[0]);
  const fin = parsearHoraAMinutos(partes[1]);
  if (ini < 0 || fin < 0) return "?";
  let dur = fin > ini ? (fin - ini) / 60 : ((24 * 60 - ini) + fin) / 60;

  if (dur < 1.5) return "IGN";
  if (ini <= 7 * 60 + 30 && fin >= 20 * 60 + 45 && fin < 22 * 60 && dur >= 13) return "DOBLE";
  if (tieneAsterisco && ini >= 7 * 60 && fin <= 20 * 60 + 30 && dur >= 11) return "D12";
  if (ini >= 20 * 60 && fin <= 9 * 60 && dur >= 8 && dur < 11) return "N10";
  if (ini >= 19 * 60 && fin <= 9 * 60 && dur >= 11) return "N12";
  if (dur >= 1.5 && dur < 6) return "PARCIAL";
  switch (seccion) {
    case TURNO.MANANA: return "M";
    case TURNO.TARDE: return "T";
    case TURNO.NOCHE: return "N";
    default: return "?";
  }
}

/* ============================================================
   FASE B — Clasificación por TURNO NATURAL + cómputo de horas.
   El turno natural (base de datos) es la verdad de referencia;
   lo del Excel son modificaciones. El dato decisivo es el TOTAL
   de horas del día, no la forma del horario.
   ============================================================ */
const UMBRAL_H = {
  min7: 6, max7: 9,        // 7h: entre 6 y 9 horas
  min10: 9, max10: 11,     // 10h noche: alrededor de 10h
  min12: 11, max12: 13,    // 12h: entre 11 y 13 horas
  dobleMin: 13,            // doblaje (natural 7h): por encima de ~13h
  ignMax: 1.5,             // microcambios: < 1.5h se ignoran
};

// Duración en horas de un tramo "ini-fin" (cruza medianoche si fin<ini)
function durTramoHoras(horario) {
  let h = String(horario || "").trim();
  if (h.charAt(0) === "*") h = h.substring(1).trim();
  const p = h.split("-");
  if (p.length !== 2) return 0;
  const ini = parsearHoraAMinutos(p[0]), fin = parsearHoraAMinutos(p[1]);
  if (ini < 0 || fin < 0) return 0;
  const dur = fin > ini ? (fin - ini) : (24 * 60 - ini) + fin;
  return dur / 60;
}
function iniTramoMin(horario) {
  let h = String(horario || "").trim();
  if (h.charAt(0) === "*") h = h.substring(1).trim();
  const p = h.split("-"); return p.length === 2 ? parsearHoraAMinutos(p[0]) : -1;
}
// Turno natural de la BD -> categoría canónica D12/N12/N10/M/T
function turnoNaturalCanon(turnoBD) {
  const t = String(turnoBD || "").toUpperCase().trim();
  if (t.indexOf("12") >= 0 && (t.indexOf("NOCHE") >= 0 || t.indexOf("NIT") >= 0)) return "N12";
  if (t.indexOf("12") >= 0) return "D12";
  if (t === "TM") return "M";
  if (t === "TT") return "T";
  if (t === "TN") return "N10";
  return "";
}
function horasDeNatural(canon) {
  switch (canon) {
    case "D12": case "N12": return 12;
    case "N10": return 10;
    case "M": case "T": return 7;
    default: return 0;
  }
}
function unirHorariosTramos(validos) {
  const ord = validos.slice().sort((a, b) => iniTramoMin(a.horario) - iniTramoMin(b.horario));
  if (ord.length === 1) return ord[0].horario.replace("* ", "");
  return ord.map(t => t.horario.replace("* ", "")).join(" + ");
}

// Clasifica la jornada de un profesional en un día.
//   natural: turno natural canónico (D12/N12/N10/M/T o "")
//   tramos: [{horario, seccion}]
// -> { turnoFinal, horario, etiqueta, totalHoras }
function clasificarJornadaNatural(natural, tramos) {
  // Sumar horas SIN contar dos veces tramos idénticos (misma jornada 12h
  // repetida en MAÑANA y TARDE como "* 8-20").
  let total = 0;
  const validos = [];
  const vistos = new Set();
  for (const tr of tramos) {
    const d = durTramoHoras(tr.horario);
    if (d < UMBRAL_H.ignMax) continue;
    const clave = String(tr.horario || "").replace(/\s|\*/g, "");
    validos.push({ ...tr, dur: d });
    if (!vistos.has(clave)) { total += d; vistos.add(clave); }
  }
  if (validos.length === 0) return { turnoFinal: "IGN", horario: tramos[0] ? tramos[0].horario : "", etiqueta: "", totalHoras: 0 };

  const secs = new Set(validos.map(t => t.seccion));
  const hayMana = secs.has(TURNO.MANANA), hayTarde = secs.has(TURNO.TARDE), hayNoche = secs.has(TURNO.NOCHE);

  // --- NOCHE ---
  if (natural === "N12" || natural === "N10" || (hayNoche && !hayMana && !hayTarde)) {
    if (total >= UMBRAL_H.min12) return { turnoFinal: TURNO.N12, horario: unirHorariosTramos(validos), etiqueta: "", totalHoras: total };
    return { turnoFinal: TURNO.NOCHE, horario: unirHorariosTramos(validos), etiqueta: "", totalHoras: total };
  }

  // "Doble físico" = mañana Y tarde con tramos DISTINTOS (no la misma jornada repetida)
  const horariosUnicos = new Set(validos.map(t => String(t.horario || "").replace(/\s|\*/g, "")));
  const esDobleFisico = hayMana && hayTarde && horariosUnicos.size >= 2;

  // --- 12h día natural ---
  if (natural === "D12") {
    if (total >= UMBRAL_H.min12 && total <= UMBRAL_H.max12 + 1) {
      return { turnoFinal: TURNO.D12, horario: "8-20", etiqueta: "", totalHoras: total };
    }
    if (total > UMBRAL_H.max12 + 1) {
      return { turnoFinal: TURNO.D12, horario: "8-20", etiqueta: "+", totalHoras: total };
    }
    // menos horas que un 12h => media jornada pactada. Si es un tramo de ~7h
    // ocupa posición normal (mañana/tarde); si es muy corto, tapahuecos PARCIAL.
    if (total >= UMBRAL_H.min7) {
      return { turnoFinal: iniTramoMin(validos[0].horario) < 13 * 60 + 30 ? TURNO.MANANA : TURNO.TARDE,
        horario: unirHorariosTramos(validos), etiqueta: "parcial", totalHoras: total };
    }
    return { turnoFinal: TURNO.PARCIAL, horario: unirHorariosTramos(validos), etiqueta: "parcial", totalHoras: total };
  }

  // --- 7h natural (mañana/tarde) ---
  if (natural === "M" || natural === "T") {
    if (esDobleFisico && total >= UMBRAL_H.dobleMin) {
      return { turnoFinal: TURNO.DOBLAJE, horario: unirHorariosTramos(validos), etiqueta: "doblaje", totalHoras: total };
    }
    if (total > UMBRAL_H.max7 && total < UMBRAL_H.dobleMin) {
      return { turnoFinal: esDobleFisico ? TURNO.DOBLAJE : (natural === "M" ? TURNO.MANANA : TURNO.TARDE),
        horario: unirHorariosTramos(validos), etiqueta: "+", totalHoras: total };
    }
    return { turnoFinal: natural === "M" ? TURNO.MANANA : TURNO.TARDE, horario: validos[0].horario.replace("* ", ""), etiqueta: "", totalHoras: total };
  }

  // --- Natural desconocido: por forma + total ---
  if (esDobleFisico) {
    if (total >= UMBRAL_H.dobleMin) return { turnoFinal: TURNO.DOBLAJE, horario: unirHorariosTramos(validos), etiqueta: "", totalHoras: total };
    return { turnoFinal: TURNO.D12, horario: "8-20", etiqueta: "", totalHoras: total };
  }
  // turno genuinamente corto (1.5–6h) sin natural conocido => tapahuecos PARCIAL
  if (total < UMBRAL_H.min7) {
    return { turnoFinal: TURNO.PARCIAL, horario: unirHorariosTramos(validos), etiqueta: "parcial", totalHoras: total };
  }
  if (hayMana) return { turnoFinal: TURNO.MANANA, horario: validos[0].horario.replace("* ", ""), etiqueta: "", totalHoras: total };
  if (hayTarde) return { turnoFinal: TURNO.TARDE, horario: validos[0].horario.replace("* ", ""), etiqueta: "", totalHoras: total };
  if (hayNoche) return { turnoFinal: TURNO.NOCHE, horario: validos[0].horario.replace("* ", ""), etiqueta: "", totalHoras: total };
  return { turnoFinal: "?", horario: validos[0].horario.replace("* ", ""), etiqueta: "", totalHoras: total };
}


/* ============================================================
   PARTE 2 — Extractor (consolidación), ProfesionalesDB,
   Semana/estadística, Motor de asignación (3 fases + rotación),
   Dobles 12h, Parciales, Decalajes, Histórico/Informe.
   Recibe un "ctx" con la planilla, la BD de profesionales,
   los alias, y el histórico (persistente).
   ============================================================ */



const TURNO_IDX = { MANANA: 0, TARDE: 1, NOCHE: 2 };

/* ---------------------------------------------------------------
   ProfesionalesDB — equivalente a M03 (carga + score + apodo)
   profArray: [{nombre, turno, apodo, criterio, suplente, antig(texto)}]
   --------------------------------------------------------------- */
function parsearAntiguedad(texto) {
  const t = String(texto || "").toUpperCase().trim();
  if (t.indexOf("> 20") >= 0) return 4;
  if (t.indexOf("10 - 20") >= 0 || t.indexOf("10-20") >= 0) return 3;
  if (t.indexOf("> 3") >= 0 && t.indexOf("< 10") >= 0) return 2;
  if (t.indexOf("< 3") >= 0) return 1;
  return 1;
}

/* ---------- FONDO de las celdas (color por profesional) ----------
   Cada profesional tiene un color de fondo fijo en el cuadrante: LILA o BLANCO.
   Por defecto se deduce de su turno habitual (12H Día / 12H Noche -> Lila;
   TM / TT / TN -> Blanco), pero la coordinadora puede cambiarlo a mano desde
   la plantilla, porque hay casos que no siguen la regla.
*/
const FONDO = { LILA: "Lila", BLANCO: "Blanco" };
const COLOR_LILA = "#D2BEE6";
const COLOR_BLANCO = "#FFFFFF";

function normalizarFondo(v) {
  const t = normalizarTexto(v);
  if (!t) return "";
  if (t.indexOf("LILA") >= 0 || t.indexOf("MORAD") >= 0 || t.indexOf("VIOLET") >= 0) return FONDO.LILA;
  if (t.indexOf("BLANC") >= 0) return FONDO.BLANCO;
  return "";
}
// Fondo que corresponde por defecto a un turno habitual.
function fondoPorDefectoDeTurno(turno) {
  const t = normalizarTexto(turno);
  // 12H Día, 12H Noche, o cualquier texto que hable de 12 h -> lila
  if (t.indexOf("12") >= 0) return FONDO.LILA;
  return FONDO.BLANCO;           // TM, TT, TN y el resto -> blanco
}

class ProfesionalesDB {
  constructor(rows, aliasRows) {
    // rows: array de objetos {nombre,turno,apodo,criterio,suplente,antiguedad}
    this.byNorm = new Map();
    this.list = [];
    for (const r of (rows || [])) {
      if (!r || !r.nombre || !String(r.nombre).trim()) continue;
      const p = {
        nombre: r.nombre,
        nombreNorm: normalizarTexto(r.nombre),
        turnoHabitual: String(r.turno || "").trim(),
        apodo: String(r.apodo || "").trim(),
        criterio: normalizarCriterio(r.criterio),
        esSuplente: String(r.suplente || "").toUpperCase().trim() === "SI",
        // "Antiguo/a": posición fija durante toda la semana, misma unidad todo el
        // mes, y una posición más cada semana dentro de esa unidad.
        esAntiguo: String(r.antiguo || "").toUpperCase().trim() === "SI",
        textoAntiguedad: String(r.antiguedad || "").trim(),
        antiguedad: parsearAntiguedad(r.antiguedad),
        // Color de fondo de sus celdas en el cuadrante. Si no viene guardado,
        // se deduce del turno habitual (12h -> lila; 7h/10h -> blanco).
        fondo: normalizarFondo(r.fondo) || fondoPorDefectoDeTurno(r.turno),
      };
      this.list.push(p);
      this.byNorm.set(p.nombreNorm, p);
    }
    // alias: dos índices del diccionario de variantes de nombre.
    //  - this.alias: clave abreviatura en MAYÚSCULAS (compatibilidad)
    //  - this.aliasNorm: clave normalizada (cubre cualquier forma con espacios/acentos)
    this.alias = new Map();
    this.aliasNorm = new Map();
    for (const a of (aliasRows || [])) {
      if (!a || !a.abreviatura) continue;
      const nombreComp = String(a.nombre || "");
      this.alias.set(String(a.abreviatura).toUpperCase().trim(), nombreComp);
      this.aliasNorm.set(normalizarTexto(a.abreviatura), nombreComp);
    }
  }
  get(nombre) { return this.byNorm.get(normalizarTexto(nombre)) || null; }
  criterio(nombre) { const p = this.get(nombre); return p ? p.criterio : CRIT.SI; }
  esSuplente(nombre) { const p = this.get(nombre); return p ? p.esSuplente : false; }
  esAntiguo(nombre) { const p = this.get(nombre); return p ? p.esAntiguo : false; }
  antiguedad(nombre) { const p = this.get(nombre); return p ? p.antiguedad : 1; }
  turnoHabitual(nombre) { const p = this.get(nombre); return p ? p.turnoHabitual : ""; }
  score(nombre) {
    const p = this.get(nombre);
    if (!p) return 100;
    let s = 0;
    if (!p.esSuplente) s += 1000;
    s += p.antiguedad * 100;
    if (p.turnoHabitual.toUpperCase().indexOf("12H") >= 0) s += 10;
    return s;
  }
  apodo(nombre) {
    const p = this.get(nombre);
    if (p && p.apodo) return p.apodo;
    return nombre;
  }
  // Color de fondo de las celdas de este profesional ("Lila" | "Blanco").
  // Si el nombre no está en la base, devuelve "" para que quien llame decida
  // (se usa entonces el tipo de jornada de ese día, como antes).
  fondo(nombre) {
    const p = this.get(nombre);
    return p ? (p.fondo || fondoPorDefectoDeTurno(p.turnoHabitual)) : "";
  }
  // --- Resolución de nombres (M01 ResolverNombreCompleto) ---
  resolverNombre(texto) {
    let t = String(texto || "").trim();
    if (t.length === 0) return "";
    // 1) SIEMPRE consultar primero el diccionario de alias (por la forma exacta
    //    normalizada). Cubre cualquier variante guardada manualmente, tenga coma,
    //    espacio o sea una sola palabra (p.ej. "Eduard Cerezo" -> "Eduardo Cerezo").
    const aliasDirecto = this.aliasNorm.get(normalizarTexto(t));
    if (aliasDirecto) return aliasDirecto;
    // 2) Nombre con coma -> buscar exacto en la base
    if (t.indexOf(",") >= 0) {
      const res = this._buscarExacto(t);
      return res || t;
    }
    // 3) Una sola palabra (abreviatura) -> alias por mayúsculas o heurística
    if (t.indexOf(" ") < 0) {
      const al = this.alias.get(t.toUpperCase().trim());
      if (al) return al;
      const h = this._buscarHeuristica(t);
      return h || t;
    }
    // 4) Varias palabras sin coma y sin alias -> intentar heurística suave por
    //    coincidencia normalizada de "Nombre Apellido" con la base, si no, tal cual
    const porNombreApellido = this._buscarPorNombreApellido(t);
    return porNombreApellido || t;
  }
  _buscarPorNombreApellido(texto) {
    // Compara el texto (normalizado, sin orden) con cada profesional cuyo nombre
    // es "Apellidos, Nombre". Coincide si todas las palabras del texto están en el
    // nombre completo del profesional (tolera orden invertido nombre/apellido).
    const palabras = normalizarTexto(texto).split(" ").filter(Boolean);
    if (palabras.length < 2) return "";
    for (const p of this.list) {
      const completo = normalizarTexto(p.nombre.replace(",", " "));
      const setComp = new Set(completo.split(" ").filter(Boolean));
      if (palabras.every(w => setComp.has(w))) return p.nombre;
    }
    return "";
  }
  _buscarExacto(nombre) {
    const nt = normalizarTexto(nombre);
    const p = this.byNorm.get(nt);
    return p ? p.nombre : "";
  }
  _buscarHeuristica(abbr) {
    const a = String(abbr).toUpperCase().trim();
    if (a.length < 2) return "";
    const inicial = a.charAt(0);
    const apellido = a.substring(1);
    for (const p of this.list) {
      const nb = p.nombre;
      const comaPos = nb.indexOf(",");
      if (comaPos < 0) continue;
      const apellidos = normalizarTexto(nb.substring(0, comaPos));
      const nombres = normalizarTexto(nb.substring(comaPos + 1));
      const primerApellido = apellidos.split(" ")[0];
      if (nombres.charAt(0) === inicial) {
        if (primerApellido === apellido || apellidos.replace(/ /g, "") === apellido) {
          return nb;
        }
      }
    }
    return "";
  }
  // ¿Este nombre (ya resuelto) corresponde a un profesional real de la BD?
  estaEnBD(nombre) {
    return this.byNorm.has(normalizarTexto(nombre));
  }
  // Resuelve un texto que puede ser un APODO (como en la hoja ROTACIONS_INF),
  // un nombre completo, un alias o una abreviatura. Devuelve el nombre completo
  // de la base, o "" si no se puede identificar con seguridad.
  resolverDesdeApodo(texto) {
    const t = String(texto || "").trim();
    if (!t) return "";
    const tn = normalizarTexto(t);
    // 1) coincidencia exacta con un apodo de la base
    for (const p of this.list) {
      if (p.apodo && normalizarTexto(p.apodo) === tn) return p.nombre;
    }
    // 2) alias guardados / nombre completo / heurísticas ya existentes
    const res = this.resolverNombre(t);
    if (this.estaEnBD(res)) return res;
    // 3) todas las palabras del texto están en el nombre completo
    const pal = tn.split(" ").filter(Boolean);
    if (pal.length >= 1) {
      for (const p of this.list) {
        const comp = new Set(normalizarTexto(p.nombre.replace(",", " ")).split(" ").filter(Boolean));
        if (pal.every(w => comp.has(w))) return p.nombre;
      }
    }
    // 4) el apodo de la base empieza por el texto (o al revés), si es inequívoco
    const cand = [];
    for (const p of this.list) {
      const ap = p.apodo ? normalizarTexto(p.apodo) : "";
      if (!ap) continue;
      if (ap.startsWith(tn) || tn.startsWith(ap)) cand.push(p.nombre);
    }
    if (cand.length === 1) return cand[0];
    return "";
  }
}

/* Recorre la planilla en el rango de semanas a generar y devuelve los nombres
   que NO se vinculan con la base de datos (para que el usuario los vincule o
   añada). Devuelve [{ raw, norm, veces, ejemploHorario }] ordenado por frecuencia. */
function extraerNombresSinVincular(planilla, db, fechaLunesInicio, numSemanas) {
  const sinVincular = new Map(); // norm -> {raw, veces, ejemploHorario}
  for (let s = 0; s < numSemanas; s++) {
    const lunes = addDays(fechaLunesInicio, s * 7);
    for (let d = 0; d < 7; d++) {
      const fecha = addDays(lunes, d);
      const colDia = weekdayMonday(fecha) + 2;
      const fechaSerial = diaSerial(fecha);
      let seccion = "";
      for (let i = 0; i < planilla.length; i++) {
        const fila = planilla[i] || [];
        const cel = (c) => fila[c - 1];
        const dval = cel(4);
        if (typeof dval === "string") {
          const u = dval.toUpperCase();
          if (u.indexOf("TURNO DE MAÑANA") >= 0) seccion = "MAÑANA";
          else if (u.indexOf("TURNO DE TARDE") >= 0) seccion = "TARDE";
          else if (u.indexOf("TURNO DE NOCHE") >= 0) seccion = "NOCHE";
        }
        const cat = cel(1) == null ? "" : String(cel(1));
        if (cat.substring(0, 3) !== "INF" || cat === "COOR_INF") continue;
        const kSerial = toSerial(cel(11));
        if (kSerial == null || fechaSerial < kSerial || fechaSerial > kSerial + 6) continue;
        const celda = cel(colDia);
        if (typeof celda !== "string" || celda.trim().length === 0) continue;
        for (const rawLin of celda.split("\n")) {
          const lin = rawLin.trim();
          if (!lin) continue;
          const { horario, nombre: nombreRaw } = parsearLineaPlanilla(lin);
          if (!nombreRaw) continue;
          const resuelto = db.resolverNombre(nombreRaw);
          if (db.estaEnBD(resuelto)) continue; // vinculado, ok
          const norm = normalizarTexto(nombreRaw);
          if (!sinVincular.has(norm)) sinVincular.set(norm, { raw: nombreRaw, veces: 0, ejemploHorario: horario });
          sinVincular.get(norm).veces++;
        }
      }
    }
  }
  const out = [...sinVincular.values()];
  out.sort((a, b) => b.veces - a.veces);
  return out;
}
function obtenerProfesionalesDelDia(planilla, db, fechaObj) {
  const ultimaFila = planilla.length;
  const colDia = weekdayMonday(fechaObj) + 2; // 1-indexed col (C=3..I=9)
  const fechaSerial = diaSerial(fechaObj);
  const dict = new Map(); // claveNorm -> registro crudo
  let seccion = "";

  for (let i = 0; i < ultimaFila; i++) {
    const fila = planilla[i] || [];
    const cell = (c) => fila[c - 1]; // c = 1-indexed col

    const dval = cell(4);
    if (typeof dval === "string") {
      if (dval.toUpperCase().indexOf("TURNO DE MAÑANA") >= 0) seccion = TURNO.MANANA;
      else if (dval.toUpperCase().indexOf("TURNO DE TARDE") >= 0) seccion = TURNO.TARDE;
      else if (dval.toUpperCase().indexOf("TURNO DE NOCHE") >= 0) seccion = TURNO.NOCHE;
    }

    const cat = cell(1) == null ? "" : String(cell(1));
    if (cat.substring(0, 3) !== "INF") continue;
    if (cat === "COOR_INF") continue;

    const fechaCell = cell(11); // col K = lunes de la semana representada
    const fSerial = toSerial(fechaCell);
    if (fSerial == null) continue;
    if (fechaSerial < fSerial || fechaSerial > fSerial + 6) continue;

    const celdaH = cell(colDia);
    if (typeof celdaH !== "string") continue;
    if (celdaH.trim().length === 0) continue;

    const lineas = celdaH.split("\n");
    for (const rawLin of lineas) {
      const lin = rawLin.trim();
      if (lin.length === 0) continue;
      const { horario, nombre: nombreRaw } = parsearLineaPlanilla(lin);
      if (nombreRaw.length === 0) continue;
      const nombreResuelto = db.resolverNombre(nombreRaw);
      const tipo = clasificarHorarioIndividual(horario, seccion);
      acumularEnDict(dict, nombreResuelto, cat, seccion, horario, i + 1, tipo);
    }
  }

  const regs = [];
  for (const v of dict.values()) regs.push(consolidarRegistro(v, db));
  ordenarRegistros(regs);
  return regs;
}

// Convierte una celda de fecha (Date o número de serie Excel) a día-serial unix
function toSerial(v) {
  if (v == null) return null;
  if (v && typeof v === 'object' && v.__date__) { const d=new Date(v.__date__); return diaSerial(new Date(d.getFullYear(),d.getMonth(),d.getDate())); }
  if (v instanceof Date) return diaSerial(v);
  if (typeof v === "number") {
    // serial Excel -> Date (epoch 1899-12-30)
    const ms = (v - 25569) * 86400000;
    const d = new Date(ms);
    return diaSerial(new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()));
  }
  // intento de parsear ISO
  const d = new Date(v);
  if (!isNaN(d.getTime())) return diaSerial(d);
  return null;
}

function acumularEnDict(d, nombre, cat, seccion, horario, fila, tipo) {
  const clave = normalizarTexto(nombre);
  if (d.has(clave)) {
    const v = d.get(clave);
    v.secciones += "|" + seccion;
    v.horarios += "|" + horario;
    v.filas += ", " + fila;
    v.tipos += "|" + tipo;
    v.tramos.push({ horario, seccion });
  } else {
    d.set(clave, { nombre, cat, secciones: seccion, horarios: horario, filas: String(fila), tipos: tipo,
      tramos: [{ horario, seccion }] });
  }
}

function hayTipo(tipos, buscado) {
  return ("|" + tipos + "|").toUpperCase().indexOf("|" + buscado + "|") >= 0;
}
function todosLosTiposSon(tipos, buscado) {
  if (!tipos || tipos.length === 0) return false;
  return tipos.split("|").every(t => t.trim() === buscado);
}
function buscarHorarioPorTipo(horarios, tipos, buscar) {
  const ph = horarios.split("|");
  const pt = tipos.split("|");
  for (let i = 0; i < pt.length; i++) if (pt[i] === buscar) return ph[i];
  return ph[0];
}

function consolidarRegistro(v, db) {
  // FASE B: clasificar por turno natural + cómputo total de horas.
  const natural = db ? turnoNaturalCanon(db.turnoHabitual(v.nombre)) : "";
  const res = clasificarJornadaNatural(natural, v.tramos || []);
  return {
    nombre: v.nombre,
    categoria: v.cat,
    turnoFinal: res.turnoFinal,
    horario: res.horario,
    etiqueta: res.etiqueta || "",        // "doblaje" | "+" | "parcial" | ""
    totalHoras: res.totalHoras || 0,
    reclasificado: res.etiqueta === "doblaje",
    filasOrigen: v.filas,
  };
}

function ordenTurno(t) {
  switch (t) {
    case TURNO.MANANA: return 1; case TURNO.TARDE: return 2; case TURNO.NOCHE: return 3;
    case TURNO.D12: return 4; case TURNO.N12: return 5; case TURNO.DOBLAJE: return 6; default: return 9;
  }
}
function ordenarRegistros(regs) {
  regs.sort((a, b) => {
    const oa = ordenTurno(a.turnoFinal), ob = ordenTurno(b.turnoFinal);
    if (oa !== ob) return oa - ob;
    const na = a.nombre.toUpperCase(), nb = b.nombre.toUpperCase();
    return na < nb ? -1 : na > nb ? 1 : 0;
  });
}


/* ============================================================
   PARTE 3 — Semana/estadística (M04), Motor (M05),
   Dobles (M07), Parciales (M08), Decalajes (M06),
   Histórico/Informe (M09).
   ============================================================ */





/* ---------- M04: extraer semana + estadística ---------- */
function extraerSemana(planilla, db, fechaLunes) {
  const semana = [];
  for (let i = 0; i < 7; i++) {
    const fecha = addDays(fechaLunes, i);
    semana.push({ fecha, diaSemana: i + 1, regs: obtenerProfesionalesDelDia(planilla, db, fecha) });
  }
  return semana;
}

function determinarBloqueDominante(d12, n12, m, t, n10) {
  let mejor = d12, etq = "D12";
  if (n12 > mejor) { mejor = n12; etq = "N12"; }
  if (m > mejor) { mejor = m; etq = "M"; }
  if (t > mejor) { mejor = t; etq = "T"; }
  if (n10 > mejor) { mejor = n10; etq = "N"; }
  return etq;
}

function calcularEstadisticasSemana(semana, db) {
  const dict = new Map();
  for (let d = 0; d < 7; d++) {
    for (const reg of semana[d].regs) {
      const clave = normalizarTexto(reg.nombre);
      if (!dict.has(clave)) {
        dict.set(clave, { nombre: reg.nombre, trabaja: 0, m: 0, t: 0, n: 0, d12: 0, n12: 0, s7m: 0, s7t: 0, n10: 0 });
      }
      const a = dict.get(clave);
      a.trabaja++;
      switch (reg.turnoFinal) {
        case TURNO.MANANA: a.m++; a.s7m++; break;
        case TURNO.TARDE: a.t++; a.s7t++; break;
        case TURNO.NOCHE: a.n++; a.n10++; break;
        case TURNO.D12: a.m++; a.t++; a.d12++; break;
        case TURNO.N12: a.n++; a.n12++; break;
        case TURNO.DOBLAJE: a.m++; a.t++; break;
      }
    }
  }
  const estad = [];
  for (const [k, a] of dict) {
    estad.push({
      nombre: a.nombre, nombreNorm: k,
      diasTrabaja: a.trabaja, diasMana: a.m, diasTarde: a.t, diasNoche: a.n,
      bloqueDominante: determinarBloqueDominante(a.d12, a.n12, a.s7m, a.s7t, a.n10),
      score: db.score(a.nombre),
    });
  }
  estad.sort((x, y) => {
    if (y.score !== x.score) return y.score - x.score;
    return x.nombre.toUpperCase() < y.nombre.toUpperCase() ? -1 : x.nombre.toUpperCase() > y.nombre.toUpperCase() ? 1 : 0;
  });
  return estad;
}

function registroDeProfesionalEnDia(diaSem, nombre) {
  const nt = normalizarTexto(nombre);
  for (const r of diaSem.regs) if (normalizarTexto(r.nombre) === nt) return r;
  return null;
}

/* ---------- M09: histórico + rotación (consulta) ----------
   historico: array de filas {fechaSerial, fecha(Date), mes, profNorm,
   prof, dia, turno(M/T/N), posGlobal, area}
   --------------------------------------------------------------- */
function posicionSemanaAnterior(historico, nombreProf, fechaLunes, area) {
  if (!historico || historico.length === 0) return 0;
  const semAnt = diaSerial(addDays(fechaLunes, -7));
  const lunesSerial = diaSerial(fechaLunes);
  const np = normalizarTexto(nombreProf);
  const conteo = new Map();
  for (const h of historico) {
    if (h.fechaSerial < semAnt || h.fechaSerial >= lunesSerial) continue;
    if (h.profNorm !== np) continue;
    if (h.area !== area) continue;
    conteo.set(h.posGlobal, (conteo.get(h.posGlobal) || 0) + 1);
  }
  let mejorPos = 0, mejorCnt = 0;
  for (const [pos, cnt] of conteo) if (cnt > mejorCnt) { mejorCnt = cnt; mejorPos = pos; }
  return mejorPos;
}

function unidadDelMesActual(historico, nombreProf, fechaLunes) {
  if (!historico || historico.length === 0) return "";
  const lunesSerial = diaSerial(fechaLunes);
  const mes = mesKey(fechaLunes);
  const np = normalizarTexto(nombreProf);
  const conteo = new Map();
  for (const h of historico) {
    if (h.fechaSerial >= lunesSerial) continue;
    if (h.mes !== mes) continue;
    if (h.profNorm !== np) continue;
    conteo.set(h.area, (conteo.get(h.area) || 0) + 1);
  }
  let mejorArea = "", mejorCnt = 0;
  for (const [a, cnt] of conteo) if (cnt > mejorCnt) { mejorCnt = cnt; mejorArea = a; }
  return mejorArea;
}

/* ---------- ROTACIÓN OBLIGATORIA MENSUAL ----------
   Solo rota Suplente=No; cada uno por las unidades que su criterio permite
   (de las abiertas); no repite hasta pasar por todas; al completar, reinicia.
*/
// Unidades de rotación abiertas en la estructura activa (UCCRI cuenta una vez)
function unidadesAbiertasRotacion() {
  const out = [];
  for (const a of _EST.areas) if (!out.includes(a.rotacion)) out.push(a.rotacion);
  return out;
}
function unidadesCandidatasProf(criterio) {
  return unidadesAbiertasRotacion().filter(u => puedeIrAArea(criterio, u));
}
// Historial mensual (mes -> unidad dominante) de un profesional
function historialMensualProf(historico, profNorm, mesActual) {
  const porMes = new Map();
  for (const h of historico) {
    if (h.profNorm !== profNorm) continue;
    if (h.mes >= mesActual) continue; // solo meses anteriores
    if (!porMes.has(h.mes)) porMes.set(h.mes, new Map());
    const m = porMes.get(h.mes);
    m.set(h.area, (m.get(h.area) || 0) + 1);
  }
  const meses = [...porMes.keys()].sort();
  const salida = [];
  for (const mes of meses) {
    const m = porMes.get(mes);
    let mejor = "", cnt = -1;
    for (const [u, c] of m) if (c > cnt) { cnt = c; mejor = u; }
    salida.push({ mes, unidad: mejor });
  }
  return salida;
}
// Unidad que debe ocupar este mes para cumplir la rotación obligatoria.
// Devuelve "" si no se fuerza (suplente o sin candidatas).
function unidadRotacionObligatoria(historico, nombreProf, criterio, esSuplente, fechaLunes) {
  if (!_REGLAS.rotacionMensual) return "";
  if (_REGLAS.suplentesNoRotan && esSuplente) return "";
  const candidatas = unidadesCandidatasProf(criterio);
  if (candidatas.length === 0) return "";
  if (candidatas.length === 1) return candidatas[0];

  const mesActual = mesKey(fechaLunes);
  const hist = historialMensualProf(historico, normalizarTexto(nombreProf), mesActual);

  // ¿Ya hay asignación de unidad ESTE mes en semanas anteriores? Mantenerla
  // (la rotación es mensual; dentro del mes no se cambia de unidad).
  if (_REGLAS.continuidadMes) {
    const yaEsteMes = unidadDelMesActual(historico, nombreProf, fechaLunes);
    if (yaEsteMes && candidatas.includes(yaEsteMes)) return yaEsteMes;
  }

  // Reconstruir el ciclo actual desde el final del historial
  const visitadas = new Set();
  if (_REGLAS.cicloCompleto) {
    for (let i = hist.length - 1; i >= 0; i--) {
      const u = hist[i].unidad;
      if (!candidatas.includes(u)) continue;
      if (visitadas.size >= candidatas.length) break;
      if (visitadas.has(u)) break;
      visitadas.add(u);
    }
    if (visitadas.size >= candidatas.length) visitadas.clear(); // ciclo completo -> reinicia
  }

  const ultima = hist.length ? hist[hist.length - 1].unidad : "";
  const penultima = hist.length >= 2 ? hist[hist.length - 2].unidad : "";
  // Regla dura: no 3 meses seguidos en la misma unidad (si está activa).
  const prohibida = (_REGLAS.max2Meses && ultima && ultima === penultima) ? ultima : "";
  let pendientes = candidatas.filter(u => !visitadas.has(u) && u !== ultima && u !== prohibida);
  if (pendientes.length === 0) pendientes = candidatas.filter(u => !visitadas.has(u) && u !== prohibida);
  if (pendientes.length === 0) pendientes = candidatas.filter(u => u !== ultima && u !== prohibida);
  if (pendientes.length === 0) pendientes = candidatas.filter(u => u !== prohibida);
  if (pendientes.length === 0) pendientes = candidatas.slice();
  return pendientes[0];
}

/* ---------- REPARTO INTELIGENTE de rotación (global por mes y bloque) ----------
   En vez de que cada profesional calcule su unidad por separado (lo que satura
   las unidades pequeñas), se reparten las PLAZAS de cada unidad entre todos los
   que deben rotar, minimizando el nº de los que repiten. Día y noche se reparten
   por separado (ocupan mapas de plazas distintos).
*/
// Cuántos meses hace que el profesional no pisa cada unidad candidata
// (mayor = más urgente). Si nunca la pisó -> Infinity.
function urgenciaPorUnidadRot(hist, candidatas) {
  const urg = {};
  for (const u of candidatas) urg[u] = Infinity;
  for (let i = hist.length - 1, dist = 1; i >= 0; i--, dist++) {
    const u = hist[i].unidad;
    if (u in urg && urg[u] === Infinity) urg[u] = dist;
  }
  return urg;
}
// Reparte un conjunto de profesionales (de un bloque) entre las plazas disponibles.
// capacidad: {unidad: nºplazas}. Devuelve Map(profNorm -> unidad).
function repartirRotacionBloque(profesionales, capacidad, historico, mesActual) {
  const asign = new Map();
  const restante = { ...capacidad };
  const datos = profesionales.map(p => {
    const cand = Object.keys(capacidad).filter(u => puedeIrAArea(p.criterio, u) && capacidad[u] > 0);
    const hist = historialMensualProf(historico, p.profNorm, mesActual);
    const ultima = hist.length ? hist[hist.length - 1].unidad : "";
    const penultima = hist.length >= 2 ? hist[hist.length - 2].unidad : "";
    // REGLA máx. 2 meses seguidos (si activa): si los DOS últimos meses fueron la
    // misma unidad, esa unidad queda PROHIBIDA este mes (sería el 3º consecutivo).
    const prohibida = (_REGLAS.max2Meses && ultima && ultima === penultima) ? ultima : "";
    const visit = new Set();
    if (_REGLAS.cicloCompleto) {
      for (let i = hist.length - 1; i >= 0; i--) {
        const u = hist[i].unidad;
        if (!cand.includes(u)) continue;
        if (visit.size >= cand.length) break;
        if (visit.has(u)) break;
        visit.add(u);
      }
      if (visit.size >= cand.length) visit.clear();
    }
    const urg = urgenciaPorUnidadRot(hist, cand);
    return { ...p, cand, ultima, prohibida, urg, pendientes: cand.filter(u => !visit.has(u)) };
  });
  // Los más restringidos (menos unidades posibles) primero
  datos.sort((a, b) => a.cand.length - b.cand.length);

  // Fase 1: unidad PENDIENTE (no visitada), distinta de la última y NO prohibida
  for (const d of datos) {
    const op = d.pendientes.filter(u => u !== d.ultima && u !== d.prohibida && restante[u] > 0)
      .sort((x, y) => d.urg[y] - d.urg[x]);
    if (op.length) { asign.set(d.profNorm, op[0]); restante[op[0]]--; }
  }
  // Fase 2: cualquier candidata distinta de la última y NO prohibida con hueco
  for (const d of datos) {
    if (asign.has(d.profNorm)) continue;
    const op = d.cand.filter(u => u !== d.ultima && u !== d.prohibida && restante[u] > 0)
      .sort((x, y) => d.urg[y] - d.urg[x]);
    if (op.length) { asign.set(d.profNorm, op[0]); restante[op[0]]--; }
  }
  // Fase 3: permitir repetir la última (pero NO la prohibida) si es la única forma
  for (const d of datos) {
    if (asign.has(d.profNorm)) continue;
    const op = d.cand.filter(u => u !== d.prohibida && restante[u] > 0)
      .sort((x, y) => d.urg[y] - d.urg[x]);
    if (op.length) { asign.set(d.profNorm, op[0]); restante[op[0]]--; }
  }
  // (Sin Fase 4: nunca se asigna la unidad prohibida. Quien quede sin unidad por
  // capacidad fluye al relleno de huecos / excedente en su unidad de rotación,
  // que sí cumple la regla de los 2 meses máximo.)
  return asign;
}
// Capacidad (plazas por unidad) de un bloque. esNoche define qué posiciones cuentan.
function capacidadPorUnidad() {
  const cap = {};
  for (const u of unidadesAbiertasRotacion()) cap[u] = posicionesDeArea(u);
  return cap;
}
// Calcula el reparto del mes para día y noche. Devuelve Map(profNorm -> unidad).
// Solo incluye fijos (Suplente=No) con bloque dominante claro.
function calcularRepartoMes(estad, db, historico, fechaLunes) {
  if (!_REGLAS.rotacionMensual) return new Map();
  const mesActual = mesKey(fechaLunes);
  // Si ya hay unidad asignada este mes (semanas previas), respetarla luego en
  // intentarAsignarPos vía unidadDelMesActual; aquí repartimos a todos los fijos.
  const dia = [], noche = [];
  for (const e of estad) {
    if (_REGLAS.suplentesNoRotan && db.esSuplente(e.nombre)) continue;
    const p = { nombre: e.nombre, profNorm: normalizarTexto(e.nombre), criterio: db.criterio(e.nombre) };
    if (esBloqueNoche(e.bloqueDominante)) noche.push(p); else dia.push(p);
  }
  const capDia = capacidadPorUnidad();
  const capNoche = capacidadPorUnidad();
  const repDia = repartirRotacionBloque(dia, capDia, historico, mesActual);
  const repNoche = repartirRotacionBloque(noche, capNoche, historico, mesActual);
  const total = new Map();
  for (const [k, v] of repDia) total.set("D|" + k, v);
  for (const [k, v] of repNoche) total.set("N|" + k, v);
  return total;
}


/* ---------- M05: MOTOR de asignación ---------- */
function esBloque7h(b) { return b === "M" || b === "T" || b === "N"; }
function esBloqueNoche(b) { return b === "N" || b === "N12"; }
function diasEnBloqueDominante(e) {
  switch (e.bloqueDominante) {
    case "D12": case "M": return e.diasMana;
    case "T": return e.diasTarde;
    case "N": case "N12": return e.diasNoche;
    default: return e.diasTrabaja;
  }
}

function nuevaMatriz() {
  // celdas[d][t][p] (p 1..15) -> nombre
  const celdas = [];
  for (let d = 0; d < 7; d++) {
    celdas.push([]);
    for (let t = 0; t < 3; t++) {
      celdas[d].push(new Array(NPOS_TOTAL_actual() + 1).fill(""));
    }
  }
  return { celdas, excedentes: {} }; // excedentes[d*3+t] = [{nombre, area}]
}

// Intenta asignar posición fija a un profesional (con rotación). Devuelve {pos}
function intentarAsignarPos(e, ocDia, ocNoche, db, historico, fechaLunes, reparto) {
  const nombre = e.nombre;
  const crit = db.criterio(nombre);
  const esSupl = db.esSuplente(nombre);
  const bloque = e.bloqueDominante;
  const areas = listaAreasOrdenadas();
  let posElegida = 0;

  // Unidad asignada por el REPARTO GLOBAL del mes (minimiza repeticiones). Si no
  // está en el reparto (p.ej. suplente), cae a la rotación individual.
  const claveRep = (esBloqueNoche(bloque) ? "N|" : "D|") + normalizarTexto(nombre);
  let unidadRot = reparto && reparto.get(claveRep);
  if (!unidadRot) unidadRot = unidadRotacionObligatoria(historico, nombre, crit, esSupl, fechaLunes);
  // Continuidad dentro del mes (si la regla está activa y ya tenía unidad este mes).
  let unidadMes = unidadRot;
  if (_REGLAS.continuidadMes) {
    const yaMes = unidadDelMesActual(historico, nombre, fechaLunes);
    if (yaMes && puedeIrAArea(crit, yaMes)) unidadMes = yaMes;
  }

  // Pasada 0 y 1: forzar la unidad asignada. Pasada 2: cualquiera compatible
  // (fallback total, solo si no cupo en su unidad — p.ej. unidad llena).
  for (let pasada = 0; pasada <= 2 && posElegida === 0; pasada++) {
    for (let ai = 0; ai < areas.length && posElegida === 0; ai++) {
      const area = areas[ai];
      if (!puedeIrAArea(crit, area)) continue;
      if (pasada <= 1 && unidadMes && area !== unidadMes) continue;
      let posEvitar = 0;
      if (pasada === 0 && _REGLAS.evitarPosicionSemanaAnterior) posEvitar = posicionSemanaAnterior(historico, nombre, fechaLunes, area);
      for (let p = posGlobalInicioArea(area); p <= posGlobalFinArea(area); p++) {
        if (pasada === 0 && p === posEvitar) continue;
        if (esBloqueNoche(bloque)) {
          if (!ocNoche[p]) { ocNoche[p] = true; posElegida = p; break; }
        } else {
          if (!ocDia[p]) { ocDia[p] = true; posElegida = p; break; }
        }
      }
    }
  }
  return posElegida;
}

/* ---------- ANTIGUOS/AS: posición fija y rotación de posición semanal ----------
   Reglas (tienen prioridad absoluta sobre todo lo demás):
   - Durante la semana NO cambian de ubicación: se les da una posición estable y
     la ocupan todos los días que trabajan.
   - Durante el MES no cambian de unidad.
   - Cada semana avanzan UNA posición dentro de su unidad; al pasar la última,
     vuelven a la primera.
   - Al mes siguiente rotan de unidad con las reglas normales.
*/
// Posición GLOBAL que le toca a un antiguo esta semana dentro de `unidad`.
// Se apoya en el histórico: si la semana pasada (del mismo mes) ocupaba una
// posición de esa unidad, le toca la siguiente; si no, la primera libre.
// ocupadas = array de booleanos de posiciones ya cogidas en este bloque.
function posicionAntiguoEstaSemana(historico, nombre, unidad, fechaLunes, ocupadas) {
  const ini = posGlobalInicioArea(unidad);
  const fin = posGlobalFinArea(unidad);
  if (ini <= 0 || fin < ini) return 0;
  const nPos = fin - ini + 1;

  // ¿Es la primera semana del mes? Entonces empieza ciclo nuevo.
  const lunesPrevio = addDays(fechaLunes, -7);
  const mismoMes = mesKey(lunesPrevio) === mesKey(fechaLunes);

  let posPrevia = 0;
  if (mismoMes) posPrevia = posicionSemanaAnterior(historico, nombre, fechaLunes, unidad);

  if (posPrevia >= ini && posPrevia <= fin) {
    // Avanzar una posición dentro de la unidad, ciclando
    const localPrev = posPrevia - ini;            // 0..nPos-1
    for (let salto = 1; salto <= nPos; salto++) {
      const cand = ini + ((localPrev + salto) % nPos);
      if (!ocupadas[cand]) return cand;
    }
  }
  // Sin referencia previa (o todas cogidas): la primera libre de la unidad
  for (let p = ini; p <= fin; p++) if (!ocupadas[p]) return p;
  return 0;
}

function asignarEstables(estad, db, historico, fechaLunes) {
  const ocDia = new Array(NPOS_TOTAL_actual() + 1).fill(false);
  const ocNoche = new Array(NPOS_TOTAL_actual() + 1).fill(false);
  const asignEst = []; // {nombreNorm, posGlobal, bloque}
  const ya = new Array(estad.length).fill(false);

  // Reparto inteligente de rotación del mes (una unidad por fijo, minimizando repeticiones)
  const reparto = calcularRepartoMes(estad, db, historico, fechaLunes);

  const pushAsign = (e, pos) => asignEst.push({ nombreNorm: normalizarTexto(e.nombre), posGlobal: pos, bloque: e.bloqueDominante });

  // PASADA 0 — ANTIGUOS/AS (prioridad absoluta, antes que nadie).
  // Se les fija la unidad del mes y la posición que les toca esta semana.
  if (_REGLAS.antiguosPosicionFija) {
    const antiguos = [];
    for (let i = 0; i < estad.length; i++) {
      if (ya[i] || !db.esAntiguo(estad[i].nombre)) continue;
      antiguos.push(i);
    }
    // Orden estable: más antigüedad primero, y a igualdad por nombre, para que
    // el reparto de posiciones no baile entre generaciones.
    antiguos.sort((a, b) => {
      const d = db.antiguedad(estad[b].nombre) - db.antiguedad(estad[a].nombre);
      if (d !== 0) return d;
      return String(estad[a].nombre).localeCompare(String(estad[b].nombre), "es");
    });
    for (const i of antiguos) {
      const e = estad[i];
      const esNoche = esBloqueNoche(e.bloqueDominante);
      const ocupadas = esNoche ? ocNoche : ocDia;
      const crit = db.criterio(e.nombre);
      // Unidad del mes: la que ya tenga este mes; si no, la del reparto/rotación.
      let unidad = unidadDelMesActual(historico, e.nombre, fechaLunes);
      if (!unidad || !puedeIrAArea(crit, unidad)) {
        unidad = reparto.get((esNoche ? "N|" : "D|") + normalizarTexto(e.nombre)) || "";
      }
      if (!unidad || !puedeIrAArea(crit, unidad)) {
        unidad = unidadRotacionObligatoria(historico, e.nombre, crit, db.esSuplente(e.nombre), fechaLunes);
      }
      if (!unidad || !puedeIrAArea(crit, unidad)) {
        const cand = unidadesCandidatasProf(crit).filter(u => unidadesAbiertasRotacion().includes(u));
        unidad = cand[0] || "";
      }
      let pos = unidad ? posicionAntiguoEstaSemana(historico, e.nombre, unidad, fechaLunes, ocupadas) : 0;
      // Si su unidad está completa, se le busca sitio en otra que su criterio permita
      if (pos === 0) {
        for (const u of listaAreasOrdenadas()) {
          if (!puedeIrAArea(crit, u)) continue;
          pos = posicionAntiguoEstaSemana(historico, e.nombre, u, fechaLunes, ocupadas);
          if (pos > 0) break;
        }
      }
      if (pos > 0) ocupadas[pos] = true;
      pushAsign(e, pos); ya[i] = true;
    }
  }

  // Pasada 1: 12HD y 12HN no-suplentes
  for (let i = 0; i < estad.length; i++) {
    if (!ya[i] && !db.esSuplente(estad[i].nombre)) {
      if (estad[i].bloqueDominante === "D12" || estad[i].bloqueDominante === "N12") {
        const pos = intentarAsignarPos(estad[i], ocDia, ocNoche, db, historico, fechaLunes, reparto);
        pushAsign(estad[i], pos); ya[i] = true;
      }
    }
  }
  // Pasada 1.5: 7h con >=5 días en bloque dominante
  for (let i = 0; i < estad.length; i++) {
    if (!ya[i] && !db.esSuplente(estad[i].nombre)) {
      if (esBloque7h(estad[i].bloqueDominante) && diasEnBloqueDominante(estad[i]) >= 5) {
        const pos = intentarAsignarPos(estad[i], ocDia, ocNoche, db, historico, fechaLunes, reparto);
        pushAsign(estad[i], pos); ya[i] = true;
      }
    }
  }
  // Pasada 2: resto de no-suplentes
  for (let i = 0; i < estad.length; i++) {
    if (!ya[i] && !db.esSuplente(estad[i].nombre)) {
      const pos = intentarAsignarPos(estad[i], ocDia, ocNoche, db, historico, fechaLunes, reparto);
      pushAsign(estad[i], pos); ya[i] = true;
    }
  }
  return { asignEst, reparto };
}

function buscarAsignEstable(asignEst, nombre) {
  const nt = normalizarTexto(nombre);
  for (const a of asignEst) if (a.nombreNorm === nt) return a;
  return null;
}

function aplicarEstablesAdiaPorDia(semana, asignEst, asig) {
  for (let d = 0; d < 7; d++) {
    for (const reg of semana[d].regs) {
      const a = buscarAsignEstable(asignEst, reg.nombre);
      if (!a || a.posGlobal <= 0) continue;
      const pos = a.posGlobal;
      switch (reg.turnoFinal) {
        case TURNO.MANANA: asig.celdas[d][TURNO_IDX.MANANA][pos] = reg.nombre; break;
        case TURNO.TARDE: asig.celdas[d][TURNO_IDX.TARDE][pos] = reg.nombre; break;
        case TURNO.NOCHE: asig.celdas[d][TURNO_IDX.NOCHE][pos] = reg.nombre; break;
        case TURNO.D12:
          asig.celdas[d][TURNO_IDX.MANANA][pos] = reg.nombre;
          asig.celdas[d][TURNO_IDX.TARDE][pos] = reg.nombre; break;
        case TURNO.N12: asig.celdas[d][TURNO_IDX.NOCHE][pos] = reg.nombre; break;
        case TURNO.DOBLAJE:
          asig.celdas[d][TURNO_IDX.MANANA][pos] = reg.nombre;
          asig.celdas[d][TURNO_IDX.TARDE][pos] = reg.nombre; break;
      }
    }
  }
}

function rellenarHuecosYExcedentes(semana, asignEst, asig, db, historico, fechaLunes, reparto) {
  // Memoria de continuidad intra-semana: posición usada por cada profesional sin
  // asignación estable, para intentar repetirla los días siguientes.
  const posPrevia = new Map();
  // Cache de la unidad de rotación por profesional. Usa el REPARTO global si el
  // profesional está en él; si no, la rotación individual.
  const rotCache = new Map();
  const unidadRotDe = (nombre, bloque) => {
    const k = (esBloqueNoche(bloque) ? "N|" : "D|") + normalizarTexto(nombre);
    if (rotCache.has(k)) return rotCache.get(k);
    let u = reparto && reparto.get(k);
    if (!u) u = unidadRotacionObligatoria(historico, nombre, db.criterio(nombre), db.esSuplente(nombre), fechaLunes);
    rotCache.set(k, u);
    return u;
  };
  for (let d = 0; d < 7; d++) {
    // Orden de colocación (si la regla está activa): primero los FIJOS por
    // antigüedad DESC, luego los SUPLENTES. Así los fijos reclaman posición antes
    // y, cuando sobra gente, caen a excedente primero los suplentes y después
    // los fijos de MENOR antigüedad.
    const regsOrden = _REGLAS.excedentesPrioridadSuplentes
      ? semana[d].regs.slice().sort((a, b) => {
          const sa = db.esSuplente(a.nombre) ? 1 : 0, sb = db.esSuplente(b.nombre) ? 1 : 0;
          if (sa !== sb) return sa - sb;                       // fijos primero
          return db.antiguedad(b.nombre) - db.antiguedad(a.nombre); // más antiguos primero
        })
      : semana[d].regs;
    for (const reg of regsOrden) {
      const a = buscarAsignEstable(asignEst, reg.nombre);
      if (a && a.posGlobal > 0) continue;
      const turnos = [-1, -1, -1];
      let bloque = "M";
      switch (reg.turnoFinal) {
        case TURNO.MANANA: turnos[0] = TURNO_IDX.MANANA; bloque = "M"; break;
        case TURNO.TARDE: turnos[0] = TURNO_IDX.TARDE; bloque = "T"; break;
        case TURNO.NOCHE: turnos[0] = TURNO_IDX.NOCHE; bloque = "N"; break;
        case TURNO.D12: turnos[0] = TURNO_IDX.MANANA; turnos[1] = TURNO_IDX.TARDE; bloque = "D12"; break;
        case TURNO.N12: turnos[0] = TURNO_IDX.NOCHE; bloque = "N12"; break;
        case TURNO.DOBLAJE: turnos[0] = TURNO_IDX.MANANA; turnos[1] = TURNO_IDX.TARDE; bloque = "M"; break;
      }
      const usada = colocarEnDiaTurno(asig, d, turnos, reg.nombre, db,
        posPrevia.get(normalizarTexto(reg.nombre)), unidadRotDe(reg.nombre, bloque));
      if (usada > 0) posPrevia.set(normalizarTexto(reg.nombre), usada);
    }
  }
  // Post-pass: asignar a cada excedente su unidad (UCCRI→REA→SEMIS→UCI A→UCI B),
  // respetando criterio. Se hace por cada día/turno con su propia secuencia.
  for (const idx in asig.excedentes) {
    const lista = asig.excedentes[idx];
    if (!lista || lista.length === 0) continue;
    asig.excedentes[idx] = repartirExcedentes(lista, db);
  }
}

// Coloca a un profesional en un día/turno. Prioridades:
//  1) su unidad de ROTACIÓN obligatoria (si la tiene y hay hueco compatible)
//  2) la posición que usó otros días de la semana (continuidad intra-semana)
//  3) primer hueco compatible
// Devuelve la posición usada (0 si fue a excedente).
function colocarEnDiaTurno(asig, d, turnos, nombre, db, posPreferida, unidadRot) {
  if (turnos[0] < 0) return 0;
  const crit = db.criterio(nombre);
  const areas = listaAreasOrdenadas();
  const t = turnos[0];
  let posElegida = 0;

  // 1) Unidad de rotación obligatoria: buscar hueco dentro de esa unidad
  if (unidadRot && puedeIrAArea(crit, unidadRot)) {
    // primero, si la posición preferida cae dentro de la unidad de rotación, reusarla
    if (posPreferida && posPreferida > 0 && asig.celdas[d][t][posPreferida] === "" &&
        posGlobalAArea(posPreferida).area === unidadRot) {
      posElegida = posPreferida;
    }
    for (let p = posGlobalInicioArea(unidadRot); posElegida === 0 && p <= posGlobalFinArea(unidadRot); p++) {
      if (asig.celdas[d][t][p] === "") posElegida = p;
    }
  }

  // 2) Continuidad: repetir la posición preferida si sigue libre y compatible
  if (posElegida === 0 && posPreferida && posPreferida > 0 && asig.celdas[d][t][posPreferida] === "") {
    const rotPref = posGlobalAArea(posPreferida).area;
    if (puedeIrAArea(crit, rotPref)) posElegida = posPreferida;
  }

  // 3) Primer hueco compatible en orden de áreas
  for (let ai = 0; ai < areas.length && posElegida === 0; ai++) {
    const area = areas[ai];
    if (!puedeIrAArea(crit, area)) continue;
    for (let p = posGlobalInicioArea(area); p <= posGlobalFinArea(area); p++) {
      if (asig.celdas[d][t][p] === "") { posElegida = p; break; }
    }
  }
  if (posElegida > 0) {
    for (let k = 0; k < 3; k++) if (turnos[k] >= 0) asig.celdas[d][turnos[k]][posElegida] = nombre;
    return posElegida;
  } else {
    // No cupo: a excedente. El área se asigna después con el orden correcto.
    for (let k = 0; k < 3; k++) {
      if (turnos[k] < 0) continue;
      const idx = d * 3 + turnos[k];
      if (!asig.excedentes[idx]) asig.excedentes[idx] = [];
      asig.excedentes[idx].push({ nombre, area: "" });
    }
    return 0;
  }
}

/* Pasada FINAL para los antiguos/as: después de los ajustes de dobles de 12 h y
   de parciales, alguien puede haberlos desplazado de su posición. Aquí se les
   devuelve a su sitio de forma garantizada (intercambiando con quien lo ocupe,
   para no descolocar al resto). Es lo que hace que la regla sea "sí o sí". */
function reafirmarAntiguos(asig, semana, db, asignEst) {
  if (!_REGLAS.antiguosPosicionFija) return;
  const nPos = NPOS_TOTAL_actual();
  for (const a of asignEst) {
    if (!a || a.posGlobal <= 0) continue;
    const reg0 = { nombreNorm: a.nombreNorm };
    // ¿es antiguo? buscamos su nombre real en los registros de la semana
    let nombreReal = "";
    for (let d = 0; d < 7 && !nombreReal; d++)
      for (const r of semana[d].regs)
        if (normalizarTexto(r.nombre) === a.nombreNorm) { nombreReal = r.nombre; break; }
    if (!nombreReal || !db.esAntiguo(nombreReal)) continue;

    const P = a.posGlobal;
    for (let d = 0; d < 7; d++) {
      // ¿trabaja este día?
      const trabaja = semana[d].regs.some(r => normalizarTexto(r.nombre) === a.nombreNorm);
      if (!trabaja) continue;
      for (let t = 0; t < 3; t++) {
        // ¿dónde está ahora en este turno?
        let posActual = 0;
        for (let p = 1; p <= nPos; p++) {
          const n = asig.celdas[d][t][p];
          if (n && normalizarTexto(n) === a.nombreNorm) { posActual = p; break; }
        }
        if (posActual === 0) continue;   // no está en este turno: nada que hacer
        if (posActual === P) continue;   // ya está en su sitio
        const otro = asig.celdas[d][t][P];
        asig.celdas[d][t][P] = nombreReal;
        asig.celdas[d][t][posActual] = otro || "";   // intercambio limpio
      }
    }
  }
}

/* ============================================================
   REPASO FINAL DE COBERTURA (siempre activo, forma parte del motor)
   Se ejecuta al final de cada semana, después de los ajustes de 12 h,
   parciales y antiguos/as, y antes de los decalajes y del histórico.
   0) Saneado: nadie en posición y excedente a la vez; nadie en una unidad
      que su criterio no permita; los 12 h de día y doblajes, mañana y tarde
      en la misma posición.
   1) Rojos con excedentes: no queda ningún rojo mientras haya un excedente
      compatible en ese turno (suplentes y 12 h incluidos). Primero se cubre
      la unidad con menos gente.
   2) Unidad a cero: si una unidad no tiene a nadie y otra tiene 2 o más, se
      trae a alguien compatible (suplente > sin posición fija > menor
      antigüedad). El resto de la semana esa persona pasa también a esa unidad
      los días en que allí hay hueco y el origen se queda con 2 o más.
   3) Se repite el paso 1.
   4) Una unidad por semana: intercambios (también entre jornadas distintas)
      que solo se aplican si reducen las unidades por persona y no dejan
      ninguna unidad a cero.
   Los antiguos/as con posición fija no se mueven nunca. Cada cambio queda
   registrado en asig.afinado.
   ============================================================ */
function afinarCobertura(asig, semana, db, asignEst) {
  const nPos = NPOS_TOTAL_actual();
  const M = TURNO_IDX.MANANA, T = TURNO_IDX.TARDE, N = TURNO_IDX.NOCHE;
  const C = asig.celdas;
  const log = [];
  asig.afinado = log;
  const nt = (s) => normalizarTexto(s);
  const areaDe = (p) => posGlobalAArea(p).area;
  const unidades = listaAreasOrdenadas();

  // Turno de cada persona cada día
  const turnoDia = [];
  for (let d = 0; d < 7; d++) {
    const m = new Map();
    for (const r of (semana[d] && semana[d].regs) || []) m.set(nt(r.nombre), r.turnoFinal);
    turnoDia.push(m);
  }
  const esDoble = (d, nombre) => {
    const tf = turnoDia[d].get(nt(nombre));
    return tf === TURNO.D12 || tf === TURNO.DOBLAJE;
  };
  // Antiguos/as con posición fija esta semana: intocables
  const fijos = new Set();
  if (_REGLAS.antiguosPosicionFija) {
    for (const a of asignEst || []) {
      if (!a || a.posGlobal <= 0) continue;
      const p = db.list ? db.list.find(x => x.nombreNorm === a.nombreNorm) : null;
      if (p && p.esAntiguo) fijos.add(a.nombreNorm);
    }
  }
  const conPosFija = new Set((asignEst || []).filter(a => a && a.posGlobal > 0).map(a => a.nombreNorm));
  const exc = (d, t) => {
    const k = d * 3 + t;
    if (!asig.excedentes[k]) asig.excedentes[k] = [];
    return asig.excedentes[k];
  };
  const quitarDeExc = (d, t, nombre) => {
    const k = d * 3 + t;
    if (!asig.excedentes[k]) return;
    asig.excedentes[k] = asig.excedentes[k].filter(e => nt(e.nombre) !== nt(nombre));
  };
  const posDe = (d, t, nombre) => {
    const n0 = nt(nombre);
    for (let p = 1; p <= nPos; p++) if (C[d][t][p] && nt(C[d][t][p]) === n0) return p;
    return 0;
  };
  const cuentaUnidad = (d, t, u) => {
    let n = 0;
    for (let p = posGlobalInicioArea(u); p <= posGlobalFinArea(u); p++) if (C[d][t][p]) n++;
    return n;
  };
  const turnosDe = (t) => (t === N ? [N] : [t]);
  const reg = (tipo, d, t, nombre, de, a) => log.push({ tipo, dia: d, turno: ["M", "T", "N"][t], nombre, de: de || "", a: a || "" });

  /* ---------- 0) Saneado ---------- */
  for (let d = 0; d < 7; d++) {
    for (let t = 0; t < 3; t++) {
      // duplicados en posiciones del mismo turno
      const vistos = new Set();
      for (let p = 1; p <= nPos; p++) {
        const n = C[d][t][p];
        if (!n) continue;
        if (vistos.has(nt(n))) { C[d][t][p] = ""; reg("duplicado", d, t, n, areaDe(p), ""); }
        else vistos.add(nt(n));
      }
      // nadie en posición y excedente a la vez
      const k = d * 3 + t;
      if (asig.excedentes[k]) {
        const antes = asig.excedentes[k].length;
        asig.excedentes[k] = asig.excedentes[k].filter(e => e && e.nombre && !vistos.has(nt(e.nombre)));
        // excedentes duplicados
        const ve = new Set();
        asig.excedentes[k] = asig.excedentes[k].filter(e => { const x = nt(e.nombre); if (ve.has(x)) return false; ve.add(x); return true; });
        if (asig.excedentes[k].length !== antes) reg("saneado", d, t, "", "", "");
      }
    }
  }
  // criterio incumplido: se intenta intercambiar con alguien del mismo turno;
  // si no se puede, pasa a excedente (y el paso 1 cubrirá el hueco).
  for (let d = 0; d < 7; d++) {
    for (let t = 0; t < 3; t++) {
      for (let p = 1; p <= nPos; p++) {
        const n = C[d][t][p];
        if (!n || puedeIrAArea(db.criterio(n), areaDe(p))) continue;
        if (fijos.has(nt(n))) continue;
        const doble = t !== N && esDoble(d, n);
        const S = doble ? [M, T] : [t];
        let hecho = false;
        for (let q = 1; q <= nPos && !hecho; q++) {
          if (q === p) continue;
          if (intercambioValido(d, S, p, q, false)) { aplicarIntercambio(d, S, p, q); hecho = true; reg("criterio", d, t, n, areaDe(p), areaDe(q)); }
        }
        if (!hecho) {
          for (const tt of S) { if (C[d][tt][p] && nt(C[d][tt][p]) === nt(n)) { C[d][tt][p] = ""; exc(d, tt).push({ nombre: n, area: "" }); } }
          reg("criterio", d, t, n, areaDe(p), "excedente");
        }
      }
    }
  }
  // 12 h de día / doblajes: mañana y tarde en la misma posición
  for (let d = 0; d < 7; d++) {
    for (const [nn, tf] of turnoDia[d]) {
      if (tf !== TURNO.D12 && tf !== TURNO.DOBLAJE) continue;
      if (fijos.has(nn)) continue;
      let pm = 0, pt = 0, nombre = "";
      for (let p = 1; p <= nPos; p++) {
        if (C[d][M][p] && nt(C[d][M][p]) === nn) { pm = p; nombre = C[d][M][p]; }
        if (C[d][T][p] && nt(C[d][T][p]) === nn) { pt = p; nombre = C[d][T][p]; }
      }
      if (pm && pt && pm === pt) continue;
      if (pm && !pt && !C[d][T][pm]) { C[d][T][pm] = nombre; quitarDeExc(d, T, nombre); reg("12h", d, T, nombre, "", areaDe(pm)); continue; }
      if (pt && !pm && !C[d][M][pt]) { C[d][M][pt] = nombre; quitarDeExc(d, M, nombre); reg("12h", d, M, nombre, "", areaDe(pt)); continue; }
      if (pm && pt && pm !== pt) {
        const movible = (x, destino) => !x || (!fijos.has(nt(x)) && !esDoble(d, x) && puedeIrAArea(db.criterio(x), areaDe(destino)));
        // a) llevar la tarde a la posición de la mañana, intercambiando con quien esté
        const otroT = C[d][T][pm];
        if (movible(otroT, pt) && puedeIrAArea(db.criterio(nombre), areaDe(pm))) {
          C[d][T][pm] = nombre; C[d][T][pt] = otroT || ""; reg("12h", d, T, nombre, areaDe(pt), areaDe(pm));
          continue;
        }
        // b) si no se puede (p.ej. un antiguo/a fijo esa tarde), llevar la mañana a la posición de la tarde
        const otroM = C[d][M][pt];
        if (movible(otroM, pm) && puedeIrAArea(db.criterio(nombre), areaDe(pt))) {
          C[d][M][pt] = nombre; C[d][M][pm] = otroM || ""; reg("12h", d, M, nombre, areaDe(pm), areaDe(pt));
        }
      }
    }
  }

  /* ---------- utilidades de intercambio ----------
     Intercambiar las columnas p y q del día d en los turnos S. Sirve para
     mover a un hueco (q vacía) o intercambiar personas, también entre
     jornadas distintas (un 12 h con un TT + un TM). */
  function intercambioValido(d, S, p, q, exigirNoCero) {
    const ap = areaDe(p), aq = areaDe(q);
    for (const t of S) {
      const a = C[d][t][p], b = C[d][t][q];
      if (a && (fijos.has(nt(a)) || !puedeIrAArea(db.criterio(a), aq))) return false;
      if (b && (fijos.has(nt(b)) || !puedeIrAArea(db.criterio(b), ap))) return false;
    }
    // los 12 h / doblajes no pueden quedar partidos entre mañana y tarde
    if (S.length === 1 && S[0] !== N) {
      const otro = S[0] === M ? T : M;
      for (const n of [C[d][S[0]][p], C[d][S[0]][q]]) if (n && esDoble(d, n) && posDe(d, otro, n)) return false;
    }
    if (exigirNoCero && ap !== aq) {
      for (const t of S) {
        const dp = (C[d][t][p] ? 1 : 0) - (C[d][t][q] ? 1 : 0); // personas que salen de ap
        if (dp > 0 && cuentaUnidad(d, t, ap) - dp <= 0) return false;
        if (dp < 0 && cuentaUnidad(d, t, aq) + dp <= 0) return false;
      }
    }
    return true;
  }
  function aplicarIntercambio(d, S, p, q) {
    for (const t of S) { const x = C[d][t][p]; C[d][t][p] = C[d][t][q]; C[d][t][q] = x; }
  }
  // Unidad "de la semana" de cada persona (la que más días ocupa), por ámbito
  function unidadesSemana(nn, noche) {
    const cnt = {};
    for (let d = 0; d < 7; d++) for (const t of (noche ? [N] : [M, T])) {
      for (let p = 1; p <= nPos; p++) if (C[d][t][p] && nt(C[d][t][p]) === nn) { const u = areaDe(p); cnt[u] = (cnt[u] || 0) + 1; }
    }
    return cnt;
  }
  function unidadPrincipal(nn, noche) {
    const cnt = unidadesSemana(nn, noche);
    let mejor = "", mx = 0;
    for (const u of Object.keys(cnt)) if (cnt[u] > mx) { mx = cnt[u]; mejor = u; }
    return mejor;
  }

  /* ---------- 1) Rojos con excedentes ---------- */
  function cubrirRojos() {
    for (let d = 0; d < 7; d++) {
      for (const t of [M, T, N]) {
        for (;;) {
          const pool = exc(d, t).filter(e => e && e.nombre);
          if (pool.length === 0) break;
          // huecos, empezando por la unidad con menos gente
          const huecos = [];
          for (let p = 1; p <= nPos; p++) if (!C[d][t][p]) huecos.push(p);
          if (huecos.length === 0) break;
          huecos.sort((a, b) => cuentaUnidad(d, t, areaDe(a)) - cuentaUnidad(d, t, areaDe(b)) || a - b);
          let colocado = false;
          for (const p of huecos) {
            const u = areaDe(p);
            let mejor = null, mejorScore = -1;
            for (const e of pool) {
              const n = e.nombre;
              if (!puedeIrAArea(db.criterio(n), u)) continue;
              const doble = t !== N && esDoble(d, n);
              const otroT = t === M ? T : M;
              let ambos = false;
              if (doble) {
                // Un 12 h va mañana y tarde en la misma posición si esa posición está
                // libre en los dos turnos. Si en el otro turno ya está ocupada (p.ej.
                // por un TM), cubre solo este turno y sigue de extra en el otro.
                const ocup = C[d][otroT][p];
                const yaEnOtra = posDe(d, otroT, n);
                if (yaEnOtra && yaEnOtra !== p) continue;   // no partir a nadie entre dos posiciones
                ambos = !ocup || nt(ocup) === nt(n);
              }
              let s = 0;
              if (unidadPrincipal(nt(n), t === N) === u) s += 4;
              if (e.area === u) s += 2;
              if (!doble || ambos) s += 1;
              if (s > mejorScore) { mejorScore = s; mejor = { e, ambos }; }
            }
            if (!mejor) continue;
            const n = mejor.e.nombre;
            const turnos = (t !== N && esDoble(d, n) && mejor.ambos) ? [M, T] : [t];
            for (const tt of turnos) { C[d][tt][p] = n; quitarDeExc(d, tt, n); }
            reg("rojo", d, t, n, "excedente", u);
            colocado = true;
            break;
          }
          if (!colocado) break;
        }
      }
    }
  }
  cubrirRojos();

  /* ---------- 2) Unidad a cero ---------- */
  for (let d = 0; d < 7; d++) {
    for (const t of [M, T, N]) {
      for (const u of unidades) {
        if (cuentaUnidad(d, t, u) > 0) continue;
        // destino: primera posición libre de u (para 12 h, libre mañana y tarde)
        // origen: unidad con más gente (>= 2)
        const origenes = unidades.filter(o => o !== u && cuentaUnidad(d, t, o) >= 2)
          .sort((a, b) => cuentaUnidad(d, t, b) - cuentaUnidad(d, t, a));
        let hecho = null;
        for (const o of origenes) {
          const cands = [];
          for (let p = posGlobalInicioArea(o); p <= posGlobalFinArea(o); p++) {
            const n = C[d][t][p];
            if (!n || fijos.has(nt(n)) || !puedeIrAArea(db.criterio(n), u)) continue;
            cands.push({ n, p });
          }
          cands.sort((a, b) => {
            const sa = db.esSuplente(a.n) ? 0 : 1, sb = db.esSuplente(b.n) ? 0 : 1;
            if (sa !== sb) return sa - sb;
            const fa = conPosFija.has(nt(a.n)) ? 1 : 0, fb = conPosFija.has(nt(b.n)) ? 1 : 0;
            if (fa !== fb) return fa - fb;
            return db.antiguedad(a.n) - db.antiguedad(b.n);
          });
          for (const c of cands) {
            const S = (t !== N && esDoble(d, c.n)) ? [M, T] : [t];
            for (let q = posGlobalInicioArea(u); q <= posGlobalFinArea(u); q++) {
              if (S.some(tt => C[d][tt][q])) continue;
              if (!intercambioValido(d, S, c.p, q, false)) continue;
              aplicarIntercambio(d, S, c.p, q);
              hecho = { n: c.n, o };
              reg("unidad a cero", d, t, c.n, o, u);
              break;
            }
            if (hecho) break;
          }
          if (hecho) break;
        }
        if (!hecho) continue;
        // El resto de la semana, los días que en u hay hueco y el origen se queda con 2 o más
        for (let d2 = 0; d2 < 7; d2++) {
          if (d2 === d) continue;
          for (const t2 of (t === N ? [N] : [M, T])) {
            const p2 = posDe(d2, t2, hecho.n);
            if (!p2 || areaDe(p2) !== hecho.o) continue;
            if (cuentaUnidad(d2, t2, hecho.o) - 1 < 2) continue;
            const S = (t2 !== N && esDoble(d2, hecho.n)) ? [M, T] : [t2];
            if (S.length === 2 && t2 === T) continue; // ya tratado al mirar la mañana
            for (let q = posGlobalInicioArea(u); q <= posGlobalFinArea(u); q++) {
              if (S.some(tt => C[d2][tt][q])) continue;
              if (!intercambioValido(d2, S, p2, q, true)) continue;
              aplicarIntercambio(d2, S, p2, q);
              reg("unidad a cero (resto de semana)", d2, t2, hecho.n, hecho.o, u);
              break;
            }
          }
        }
      }
    }
  }

  /* ---------- 3) Repetir el paso 1 ---------- */
  cubrirRojos();

  /* ---------- 4) Una unidad por semana ---------- */
  // Coste = suma, por persona y ámbito (día/noche), de unidades distintas en la semana.
  function costePersonas(nombres) {
    let c = 0;
    for (const nn of nombres) {
      c += Object.keys(unidadesSemana(nn, false)).length;
      c += Object.keys(unidadesSemana(nn, true)).length;
    }
    return c;
  }
  for (let vuelta = 0; vuelta < 6; vuelta++) {
    let mejoras = 0;
    // personas en más de una unidad
    const todas = new Set();
    for (let d = 0; d < 7; d++) for (let t = 0; t < 3; t++) for (let p = 1; p <= nPos; p++) if (C[d][t][p]) todas.add(nt(C[d][t][p]));
    for (const nn of todas) {
      if (fijos.has(nn)) continue;
      for (const noche of [false, true]) {
        const cnt = unidadesSemana(nn, noche);
        if (Object.keys(cnt).length <= 1) continue;
        const principal = unidadPrincipal(nn, noche);
        for (let d = 0; d < 7; d++) {
          for (const t of (noche ? [N] : [M, T])) {
            const p = posDe(d, t, nn);
            if (!p || areaDe(p) === principal) continue;
            const nombre = C[d][t][p];
            const opciones = (t === N) ? [[N]] : (esDoble(d, nombre) ? [[M, T]] : [[t], [M, T]]);
            let aplicado = false;
            for (const S of opciones) {
              if (aplicado) break;
              for (let q = posGlobalInicioArea(principal); q <= posGlobalFinArea(principal) && !aplicado; q++) {
                if (!intercambioValido(d, S, p, q, true)) continue;
                const afectados = new Set([nn]);
                for (const tt of S) { if (C[d][tt][p]) afectados.add(nt(C[d][tt][p])); if (C[d][tt][q]) afectados.add(nt(C[d][tt][q])); }
                const antes = costePersonas(afectados);
                aplicarIntercambio(d, S, p, q);
                const despues = costePersonas(afectados);
                if (despues < antes) {
                  aplicado = true; mejoras++;
                  const otros = [...afectados].filter(x => x !== nn);
                  reg("una unidad por semana", d, t, nombre, areaDe(p), principal + (otros.length ? " (intercambio)" : ""));
                } else {
                  aplicarIntercambio(d, S, p, q); // deshacer
                }
              }
            }
          }
        }
      }
    }
    if (mejoras === 0) break;
  }
  return log;
}

function construirAsignacionSemanal(planilla, db, historico, fechaLunes) {
  const asig = nuevaMatriz();
  asig.fechaLunes = fechaLunes;
  const semana = extraerSemana(planilla, db, fechaLunes);
  const estad = calcularEstadisticasSemana(semana, db);
  const { asignEst, reparto } = asignarEstables(estad, db, historico, fechaLunes);
  aplicarEstablesAdiaPorDia(semana, asignEst, asig);
  rellenarHuecosYExcedentes(semana, asignEst, asig, db, historico, fechaLunes, reparto);
  return { asig, semana, estad, asignEst };
}


/* ============================================================
   PARTE 4 — Dobles 12h (M07), Parciales (M08),
   Decalajes (M06), Histórico/Informe (M09), Orquestador.
   ============================================================ */







/* ---------- M07: dobles 12h ---------- */
function turnoEfectivoEnDia(planilla, db, fecha, nombre) {
  const regs = obtenerProfesionalesDelDia(planilla, db, fecha);
  const nt = normalizarTexto(nombre);
  for (const r of regs) if (normalizarTexto(r.nombre) === nt) return r.turnoFinal;
  return "";
}

function ajustarTurnosDobles(asig, planilla, db) {
  const semana = extraerSemana(planilla, db, asig.fechaLunes);
  // Paso 1: replicar 12HD de mañana a tarde misma posición
  for (let d = 0; d < 7; d++) {
    for (let p = 1; p <= NPOS_TOTAL_actual(); p++) {
      const nombreM = asig.celdas[d][TURNO_IDX.MANANA][p];
      if (!nombreM) continue;
      const turnoM = turnoEfectivoEnDia(planilla, db, addDays(asig.fechaLunes, d), nombreM);
      if (turnoM !== TURNO.D12 && turnoM !== TURNO.DOBLAJE) continue;
      const nombreT = asig.celdas[d][TURNO_IDX.TARDE][p];
      if (!nombreT) {
        asig.celdas[d][TURNO_IDX.TARDE][p] = nombreM;
      } else if (normalizarTexto(nombreT) === normalizarTexto(nombreM)) {
        // ya está
      } else if (db.esSuplente(nombreT) && !db.esSuplente(nombreM)) {
        asig.celdas[d][TURNO_IDX.TARDE][p] = nombreM;
      }
    }
  }
  // Paso 2 y 3: rellenar huecos de tarde con 7h-tarde (primero no-suplentes, luego suplentes)
  rellenarTarde7h(asig, semana, db, false);
  rellenarTarde7h(asig, semana, db, true);
}

function rellenarTarde7h(asig, semana, db, incluirSuplentes) {
  for (let d = 0; d < 7; d++) {
    const yaUsado = new Set();
    for (let p = 1; p <= NPOS_TOTAL_actual(); p++) {
      const n = asig.celdas[d][TURNO_IDX.TARDE][p];
      if (n) yaUsado.add(normalizarTexto(n));
    }
    let candidatos = [];
    for (const r of semana[d].regs) {
      if (yaUsado.has(normalizarTexto(r.nombre))) continue;
      if (r.turnoFinal !== TURNO.TARDE) continue;
      const sup = db.esSuplente(r.nombre);
      if (!incluirSuplentes && sup) continue;
      if (incluirSuplentes && !sup) continue;
      candidatos.push({ nombre: r.nombre, score: db.score(r.nombre) });
    }
    if (candidatos.length === 0) continue;
    candidatos.sort((a, b) => b.score - a.score);
    const used = new Array(candidatos.length).fill(false);
    for (let p = 1; p <= NPOS_TOTAL_actual(); p++) {
      if (asig.celdas[d][TURNO_IDX.TARDE][p]) continue;
      const { area } = posGlobalAArea(p);
      let mejorIdx = -1;
      for (let k = 0; k < candidatos.length; k++) {
        if (used[k]) continue;
        if (puedeIrAArea(db.criterio(candidatos[k].nombre), area)) { mejorIdx = k; break; }
      }
      if (mejorIdx >= 0) { asig.celdas[d][TURNO_IDX.TARDE][p] = candidatos[mejorIdx].nombre; used[mejorIdx] = true; }
    }
  }
}

/* ---------- M08: parciales ---------- */
function horasDesdeHorario(horario) {
  let h = String(horario || "").trim();
  if (h.charAt(0) === "*") h = h.substring(1).trim();
  const partes = h.split("-");
  if (partes.length !== 2) return { ini: -1, fin: -1 };
  return { ini: parsearHoraAMinutos(partes[0]), fin: parsearHoraAMinutos(partes[1]) };
}

function ajustarParciales(asig, planilla, db) {
  const semana = extraerSemana(planilla, db, asig.fechaLunes);
  for (let d = 0; d < 7; d++) {
    let parciales = [];
    for (const r of semana[d].regs) {
      if (r.turnoFinal !== TURNO.PARCIAL) continue;
      const { ini } = horasDesdeHorario(r.horario);
      if (ini < 0) continue;
      let turnoIdx;
      if (ini >= 21 * 60 || ini < 6 * 60) turnoIdx = TURNO_IDX.NOCHE;
      else if (ini >= 13 * 60 + 30) turnoIdx = TURNO_IDX.TARDE;
      else turnoIdx = TURNO_IDX.MANANA;
      parciales.push({ nombre: r.nombre, score: db.score(r.nombre), turnoIdx });
    }
    if (parciales.length === 0) continue;
    parciales.sort((a, b) => b.score - a.score);
    const yaUsado = new Set();
    for (let p = 1; p <= NPOS_TOTAL_actual(); p++) {
      for (let t = 0; t < 3; t++) { const n = asig.celdas[d][t][p]; if (n) yaUsado.add(normalizarTexto(n)); }
    }
    for (const par of parciales) {
      if (yaUsado.has(normalizarTexto(par.nombre))) continue;
      let mejorIdx = -1;
      for (let p = 1; p <= NPOS_TOTAL_actual(); p++) {
        if (asig.celdas[d][par.turnoIdx][p]) continue;
        const { area } = posGlobalAArea(p);
        if (puedeIrAArea(db.criterio(par.nombre), area)) { mejorIdx = p; break; }
      }
      if (mejorIdx > 0) { asig.celdas[d][par.turnoIdx][mejorIdx] = par.nombre; yaUsado.add(normalizarTexto(par.nombre)); }
    }
  }
}

/* ---------- M06: decalajes ---------- */
function horarioAMinutos(horario) {
  let h = String(horario || "").trim();
  if (h.charAt(0) === "*") h = h.substring(1).trim();
  const partes = h.split("-");
  if (partes.length !== 2) return { ini: -1, fin: -1 };
  return { ini: parsearHoraAMinutos(partes[0]), fin: parsearHoraAMinutos(partes[1]) };
}
function esCoberturaTardeNoche(ini, fin) {
  if (fin < 21 * 60) return false;
  if (fin > 22 * 60 + 30) return false;
  if (ini >= 19 * 60) return false;
  return true;
}
function esCoberturaNocheMana(ini, fin) {
  if (fin < 7 * 60 + 30) return false;
  if (fin > 8 * 60 + 30) return false;
  if (ini < 19 * 60) return false;
  return true;
}
function detectarDecalajesSemana(planilla, db, fechaLunes) {
  const semana = extraerSemana(planilla, db, fechaLunes);
  const dec = [];
  for (let d = 0; d < 7; d++) {
    const dd = { tn: [], nm: [] }; // cobertores TARDE->NOCHE y NOCHE->MAÑANA
    for (const r of semana[d].regs) {
      if (!r.horario || r.horario.indexOf("+") >= 0) continue;
      const { ini, fin } = horarioAMinutos(r.horario);
      if (ini < 0 || fin < 0) continue;
      if (esCoberturaTardeNoche(ini, fin)) dd.tn.push({ nombre: r.nombre, horario: r.horario });
      if (esCoberturaNocheMana(ini, fin)) dd.nm.push({ nombre: r.nombre, horario: r.horario });
    }
    dec.push(dd);
  }
  return dec;
}
function horarioRealDelDia(planilla, db, fecha, nombre) {
  const regs = obtenerProfesionalesDelDia(planilla, db, fecha);
  const nt = normalizarTexto(nombre);
  for (const r of regs) if (normalizarTexto(r.nombre) === nt) return r.horario;
  return "";
}
function horarioEmpiezaTarde(horario) {
  const { ini } = horarioAMinutos(horario);
  if (ini < 0) return false;
  return ini >= 7 * 60 + 30 && ini <= 9 * 60;
}
function horarioEmpiezaNocheTarde(horario) {
  const { ini } = horarioAMinutos(horario);
  if (ini < 0) return false;
  return ini >= 20 * 60 + 30;
}
// Construye lista de avisos de decalaje (lo que el VBA escribe en cols L-P)
function construirAvisosDecalaje(asig, planilla, db) {
  const dec = detectarDecalajesSemana(planilla, db, asig.fechaLunes);
  const avisos = [];
  for (let d = 0; d < 7; d++) {
    const fecha = addDays(asig.fechaLunes, d);
    const usadosTM = new Set(), usadosNT = new Set();
    for (let p = 1; p <= NPOS_TOTAL_actual(); p++) {
      // TM (mañana): ocupante que empieza tarde -> cubierto por alguien que sale de noche (nm)
      const ocM = asig.celdas[d][TURNO_IDX.MANANA][p];
      if (ocM) {
        const hor = horarioRealDelDia(planilla, db, fecha, ocM);
        if (horarioEmpiezaTarde(hor)) {
          const cobertor = pickCobertor(dec[d].nm, ocM, usadosTM);
          if (cobertor) usadosTM.add(normalizarTexto(cobertor));
          avisos.push({ fecha, pos: String(posGlobalAArea(p).posLocal), area: nombreAreaCortoDePosGlobal(p),
            tipo: "TM (07:00-08:00)", cubiertoPor: cobertor, cubierto: !!cobertor });
        }
      }
      // NOCHE: ocupante que empieza noche-tarde -> cubierto por alguien tarde->noche (tn)
      const ocN = asig.celdas[d][TURNO_IDX.NOCHE][p];
      if (ocN) {
        const hor = horarioRealDelDia(planilla, db, fecha, ocN);
        if (horarioEmpiezaNocheTarde(hor)) {
          const cobertor = pickCobertor(dec[d].tn, ocN, usadosNT);
          if (cobertor) usadosNT.add(normalizarTexto(cobertor));
          avisos.push({ fecha, pos: String(posGlobalAArea(p).posLocal), area: nombreAreaCortoDePosGlobal(p),
            tipo: "NOCHE (20:00-21:25)", cubiertoPor: cobertor, cubierto: !!cobertor });
        }
      }
    }
  }
  return avisos;
}
function pickCobertor(lista, evitar, usados) {
  const ne = normalizarTexto(evitar);
  for (const c of lista) {
    const nc = normalizarTexto(c.nombre);
    if (nc === ne) continue;
    if (usados.has(nc)) continue;
    return c.nombre;
  }
  return "";
}

/* ---------- M09: registrar en histórico + informe mensual ---------- */
// Devuelve filas nuevas para el histórico (solo turnos M/T/N ocupados)
/* =============================================================
   IMPORTAR el registro de rotación que lleva la coordinadora a mano
   (libro ROTACIONS_INF: hojas DIA y NIT, una fila por profesional —
   por su APODO —, una columna por mes, celda = unidad de ese mes).
   Se convierte en filas de histórico para que la app continúe la
   rotación a partir de ahí.
   ============================================================= */

// Nombres de mes admitidos (catalán y castellano) -> número de mes 1..12
const MESES_IMPORT = {
  GENER: 1, ENERO: 1,
  FEBRER: 2, FEBRERO: 2,
  MARC: 3, MARZO: 3,
  ABRIL: 4,
  MAIG: 5, MAYO: 5,
  JUNY: 6, JUNIO: 6,
  JULIOL: 7, JULIO: 7,
  AGOST: 8, AGOSTO: 8,
  SETEMBRE: 9, SEPTIEMBRE: 9, SEPTEMBRE: 9,
  OCTUBRE: 10, OCTUBRE_ES: 10,
  NOVEMBRE: 11, NOVIEMBRE: 11,
  DESEMBRE: 12, DICIEMBRE: 12, DECEMBRE: 12,
};
function mesDesdeEncabezado(texto) {
  const t = normalizarTexto(texto).replace(/[^A-ZÑ]/g, "");
  return MESES_IMPORT[t] || 0;
}

// Interpreta el texto de una celda -> lista de unidades de rotación.
// Admite: "SEMIS", "UCI B", "UCI A-B" (las dos), "REA/SEMIS", "semis/" (barra
// suelta), mayúsculas/minúsculas y espacios sobrantes.
// Devuelve { unidades: [...], reconocido: true|false }.
function parsearUnidadesCelda(texto) {
  const bruto = String(texto || "").trim();
  if (!bruto) return { unidades: [], reconocido: true }; // celda vacía: sin dato
  const t = normalizarTexto(bruto);
  // "UCI A-B" (o "UCI A - B") = las dos UCI
  if (/^UCI\s*A\s*-\s*B$/.test(t)) return { unidades: [AREA.UCI_A, AREA.UCI_B], reconocido: true };
  const trozos = t.split("/").map(s => s.trim()).filter(Boolean);
  const unidades = [];
  let reconocido = true;
  for (const tr of trozos) {
    const u = unidadDesdeTexto(tr);
    if (u) { if (!unidades.includes(u)) unidades.push(u); }
    else reconocido = false;
  }
  return { unidades, reconocido: reconocido && unidades.length > 0 };
}
// Un trozo de texto -> nombre de rotación canónico, o "" si no se reconoce.
function unidadDesdeTexto(t) {
  const s = normalizarTexto(t).replace(/\s+/g, " ").trim();
  if (!s) return "";
  if (/^UCI\s*A$/.test(s)) return AREA.UCI_A;
  if (/^UCI\s*B$/.test(s)) return AREA.UCI_B;
  if (/^SEMIS/.test(s)) return AREA.SEMIS;
  if (/^REA/.test(s)) return AREA.REA;
  if (/^UCCRI/.test(s)) return AREA.UCCRI;
  return "";
}

/* Analiza las hojas leídas del libro y prepara la importación.
   hojas: [{ bloque: "DIA"|"NIT", encabezados: [...8 textos...], filas: [{ nombre, celdas: [...] }] }]
   vinculacionesManuales: { apodoExcel: nombreCompletoBD } decididas por el usuario.
   Devuelve un análisis con todo lo necesario para revisar antes de importar. */
function analizarImportacionRotaciones(hojas, db, anio, vinculacionesManuales) {
  const vinc = vinculacionesManuales || {};
  const entradas = [];          // {prof, profNorm, bloque, mes, unidades}
  const sinVincularMap = new Map(); // apodoExcel -> {apodo, bloque, meses}
  const noReconocidas = [];     // {apodo, mesNombre, texto}
  const mesesSet = new Set();

  for (const hoja of (hojas || [])) {
    const bloque = hoja.bloque === "NIT" ? "NIT" : "DIA";
    const mesesCol = (hoja.encabezados || []).map(h => mesDesdeEncabezado(h));
    for (const fila of (hoja.filas || [])) {
      const apodoExcel = String(fila.nombre || "").trim();
      if (!apodoExcel) continue;
      // ¿tiene algún dato?
      const tieneDatos = (fila.celdas || []).some(c => String(c || "").trim());
      if (!tieneDatos) continue;
      // Resolver el profesional
      let prof = vinc[apodoExcel] || "";
      if (!prof) prof = db.resolverDesdeApodo(apodoExcel);
      if (!prof || !db.estaEnBD(prof)) {
        if (!sinVincularMap.has(apodoExcel)) sinVincularMap.set(apodoExcel, { apodo: apodoExcel, bloque, meses: 0 });
        sinVincularMap.get(apodoExcel).meses += (fila.celdas || []).filter(c => String(c || "").trim()).length;
        continue;
      }
      for (let i = 0; i < (fila.celdas || []).length; i++) {
        const txt = String(fila.celdas[i] || "").trim();
        if (!txt) continue;
        const nMes = mesesCol[i] || 0;
        if (!nMes) continue;
        const { unidades, reconocido } = parsearUnidadesCelda(txt);
        if (!reconocido || unidades.length === 0) {
          noReconocidas.push({ apodo: apodoExcel, mesNombre: (hoja.encabezados || [])[i] || "", texto: txt });
          continue;
        }
        const mes = anio + "-" + String(nMes).padStart(2, "0");
        mesesSet.add(mes);
        entradas.push({ prof, profNorm: normalizarTexto(prof), bloque, mes, nMes, unidades });
      }
    }
  }
  const meses = [...mesesSet].sort();
  const profs = new Set(entradas.map(e => e.prof));
  return {
    entradas,
    sinVincular: [...sinVincularMap.values()].sort((a, b) => b.meses - a.meses),
    noReconocidas,
    meses,
    nProfesionales: profs.size,
    nRegistros: entradas.length,
  };
}

/* Convierte el análisis en filas de histórico listas para guardar.
   Por cada (profesional, mes, unidad) se generan varios "días" sintéticos
   dentro de ese mes, para que la unidad quede registrada como la del mes.
   Si la celda traía dos unidades, se reparten los días entre ambas (así el
   histórico refleja que ese mes estuvo repartido, igual que anotó la coordinadora). */
function filasHistoricoDesdeRotaciones(analisis, anio) {
  const filas = [];
  const DIAS_POR_MES = 10; // días sintéticos por profesional y mes
  for (const e of analisis.entradas) {
    const turno = e.bloque === "NIT" ? "N" : "M";
    const nU = e.unidades.length;
    // repartir los días entre las unidades (la primera se lleva el resto)
    const base = Math.floor(DIAS_POR_MES / nU);
    const reparto = e.unidades.map((u, i) => ({ unidad: u, dias: base + (i === 0 ? DIAS_POR_MES - base * nU : 0) }));
    let diaDelMes = 1;
    for (const r of reparto) {
      for (let k = 0; k < r.dias; k++) {
        const fecha = new Date(Number(anio), e.nMes - 1, Math.min(28, diaDelMes));
        diaDelMes++;
        filas.push({
          fechaSerial: diaSerial(fecha),
          fecha,
          mes: e.mes,
          profNorm: e.profNorm,
          prof: e.prof,
          dia: (diaDelMes - 2) % 7,
          turno,
          posGlobal: 0,      // sin posición concreta: solo interesa la unidad
          area: r.unidad,
        });
      }
    }
  }
  return filas;
}

function filasHistoricoDeSemana(asig) {
  const filas = [];
  const letra = ["M", "T", "N"];
  for (let d = 0; d < 7; d++) {
    const fecha = addDays(asig.fechaLunes, d);
    for (let t = 0; t < 3; t++) {
      for (let p = 1; p <= NPOS_TOTAL_actual(); p++) {
        const nombre = asig.celdas[d][t][p];
        if (!nombre) continue;
        const { area } = posGlobalAArea(p);
        filas.push({
          fechaSerial: diaSerial(fecha), fecha, mes: mesKey(fecha),
          profNorm: normalizarTexto(nombre), prof: nombre,
          dia: d, turno: letra[t], posGlobal: p, area,
        });
      }
      // Excedentes de este día/turno: también cuentan para el histórico (su unidad
      // es la que se les asignó como posición extra). posGlobal 0 = sin posición fija.
      const lista = asig.excedentes[d * 3 + t] || [];
      for (const ex of lista) {
        if (!ex.nombre) continue;
        filas.push({
          fechaSerial: diaSerial(fecha), fecha, mes: mesKey(fecha),
          profNorm: normalizarTexto(ex.nombre), prof: ex.nombre,
          dia: d, turno: letra[t], posGlobal: 0, area: ex.area || "",
        });
      }
    }
  }
  return filas;
}

// Informe mensual de rotación, separado en DÍA y NOCHE (como ROTACIONS_INF).
// Para cada (mes, profesional, ámbito día/noche) calcula la unidad principal
// (la de más semanas distintas) y si ha rotado entre unidades ese mes.
// El turno M/T cuenta como DÍA; N cuenta como NOCHE.
function regenerarInformeMensual(historico) {
  // ámbito: "DIA" si turno M/T, "NIT" si turno N
  function ambitoDe(turno) { return turno === "N" ? "NIT" : "DIA"; }
  // (mes||ambito||prof||area) -> set de lunes (semanas distintas en esa unidad)
  const semanasSet = new Map();
  for (const h of historico) {
    const amb = ambitoDe(h.turno);
    const lun = diaSerial(addDays(h.fecha, -((weekdayMonday(h.fecha)) - 1)));
    const claveS = h.mes + "||" + amb + "||" + h.prof + "||" + h.area;
    if (!semanasSet.has(claveS)) semanasSet.set(claveS, new Set());
    semanasSet.get(claveS).add(lun);
  }
  // (mes||ambito||prof) -> {area: nSemanas}
  const acum = new Map();
  for (const [k, set] of semanasSet) {
    const [mes, amb, prof, area] = k.split("||");
    const claveMP = mes + "||" + amb + "||" + prof;
    if (!acum.has(claveMP)) acum.set(claveMP, new Map());
    acum.get(claveMP).set(area, set.size);
  }
  const filas = [];
  for (const [k, areas] of acum) {
    let mejorArea = "", mejorCnt = 0;
    for (const [a, c] of areas) if (c > mejorCnt) { mejorCnt = c; mejorArea = a; }
    let resumen;
    if (areas.size === 1) resumen = mejorArea;
    else {
      resumen = "Unidad principal = " + mejorArea;
      for (const [a, c] of areas) if (a !== mejorArea) resumen += " · " + a + " (" + c + " sem.)";
    }
    const [mes, ambito, prof] = k.split("||");
    filas.push({ mes, ambito, profesional: prof, unidadPrincipal: mejorArea, roto: areas.size > 1, resumen });
  }
  // ordenar por ámbito (DIA antes que NIT), luego mes, luego profesional
  filas.sort((a, b) => {
    if (a.ambito !== b.ambito) return a.ambito === "DIA" ? -1 : 1;
    if (a.mes !== b.mes) return a.mes < b.mes ? -1 : 1;
    return a.profesional.toUpperCase() < b.profesional.toUpperCase() ? -1 : 1;
  });
  return filas;
}

/* ---------- Histórico de rotación mensual (apartado propio) ----------
   Para cada profesional y mes, calcula la unidad (o unidades) donde ha estado,
   según el % de DÍAS trabajados en cada unidad:
   - Si una unidad concentra el 66% o más de los días -> esa unidad sola.
   - Si ninguna llega al 66% -> las DOS unidades con más % (mayor primero).
   Se considera la rotación independientemente del turno (día/noche se mezclan,
   porque lo que importa es la unidad asistencial del mes). UCCRI esq/dre = "UCCRI".
   Devuelve, por profesional, una lista de { mes, unidades:[...], etiqueta }.
*/
function resumenHistoricoRotacion(historico) {
  const UMBRAL = 0.66;
  // (prof||mes) -> Map(area -> nº de días distintos)
  const diasPorArea = new Map();
  // para contar días distintos por (prof,mes,area) usamos un set de fechaSerial
  const setDias = new Map(); // clave prof||mes||area -> Set(fechaSerial)
  for (const h of historico) {
    const claveSet = h.prof + "||" + h.mes + "||" + h.area;
    if (!setDias.has(claveSet)) setDias.set(claveSet, new Set());
    setDias.get(claveSet).add(h.fechaSerial);
  }
  for (const [clave, set] of setDias) {
    const [prof, mes, area] = clave.split("||");
    const claveMP = prof + "||" + mes;
    if (!diasPorArea.has(claveMP)) diasPorArea.set(claveMP, new Map());
    diasPorArea.get(claveMP).set(area, set.size);
  }

  // Agrupar por profesional
  const porProf = new Map();
  for (const [claveMP, areas] of diasPorArea) {
    const [prof, mes] = claveMP.split("||");
    // total de días (suma de días por área; un día en 2 áreas cuenta en cada una,
    // pero para % usamos el total de "presencias-día")
    let total = 0;
    for (const [, n] of areas) total += n;
    // ordenar áreas por nº de días desc
    const orden = [...areas.entries()].sort((a, b) => b[1] - a[1]);
    let unidades;
    const top = orden[0];
    if (top && total > 0 && top[1] / total >= UMBRAL) {
      unidades = [top[0]]; // una sola unidad domina
    } else {
      // las dos con más días (si solo hay una, esa)
      unidades = orden.slice(0, 2).map(e => e[0]);
    }
    if (!porProf.has(prof)) porProf.set(prof, []);
    porProf.get(prof).push({ mes, unidades, etiqueta: unidades.join(" + ") });
  }

  // Convertir a lista ordenada por profesional, con sus meses ordenados
  const salida = [];
  for (const [prof, meses] of porProf) {
    meses.sort((a, b) => a.mes < b.mes ? -1 : a.mes > b.mes ? 1 : 0);
    salida.push({ profesional: prof, meses });
  }
  salida.sort((a, b) => a.profesional.toUpperCase() < b.profesional.toUpperCase() ? -1 : 1);
  return salida;
}
function generarCronograma(planilla, db, historicoInicial, fechaLunesInicio, numSemanas, estructuraConfig, reglasConfig) {
  // Fijar la estructura de la unidad para esta generación (Fase A).
  // Si no se pasa ninguna, usa la estructura por defecto.
  fijarEstructura(estructuraConfig || estructuraPorDefecto());
  // Fijar las reglas de rotación activas (si no se pasan, todas activas).
  fijarReglas(reglasConfig);
  let historico = (historicoInicial || []).slice();
  const semanas = [];
  for (let s = 0; s < numSemanas; s++) {
    const fechaLunes = addDays(fechaLunesInicio, s * 7);
    const { asig, semana, estad, asignEst } = construirAsignacionSemanal(planilla, db, historico, fechaLunes);
    ajustarTurnosDobles(asig, planilla, db);
    ajustarParciales(asig, planilla, db);
    // Los antiguos/as vuelven a su posición si algún ajuste los ha desplazado
    reafirmarAntiguos(asig, semana, db, asignEst);
    // Repaso final: cubrir rojos con extras, rescatar unidades a cero y
    // agrupar a cada persona en una sola unidad durante la semana.
    afinarCobertura(asig, semana, db, asignEst);
    const avisos = construirAvisosDecalaje(asig, planilla, db);
    const filasHist = filasHistoricoDeSemana(asig);
    historico = historico.concat(filasHist); // acumular para que la siguiente rote
    semanas.push({ fechaLunes, asig, semana, estad, avisos });
  }
  const informe = regenerarInformeMensual(historico);
  // devolver también la estructura compilada para que el render dibuje las áreas reales
  return { semanas, historico, informe, estructura: estructuraActiva() };
}


  /* ===== Exportador a XLSX (colores del Excel original) ===== */
/* ============================================================
   EXPORTADOR A XLSX (para subir a Google Sheets)
   Construye el archivo .xlsx a mano (Office Open XML) para poder
   aplicar exactamente el código de colores del Excel original.
   Sin dependencias externas: genera las partes XML y las comprime
   en un ZIP store-only. Google Sheets importa colores y formato.

   API:  exportarCronogramaXLSX(resultado, db, ENGINE) -> Uint8Array
   ============================================================ */

/* ---------- ZIP store-only (sin compresión) ---------- */
function _crc32(bytes) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < bytes.length; i++) {
    let c = (crc ^ bytes[i]) & 0xFF;
    for (let k = 0; k < 8; k++) c = (c & 1) ? (c >>> 1) ^ 0xEDB88320 : c >>> 1;
    crc = (crc >>> 8) ^ c;
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}
function _strBytes(s) {
  // UTF-8 encode
  if (typeof TextEncoder !== "undefined") return new TextEncoder().encode(s);
  const out = [];
  for (let i = 0; i < s.length; i++) {
    let c = s.charCodeAt(i);
    if (c < 0x80) out.push(c);
    else if (c < 0x800) { out.push(0xC0 | (c >> 6), 0x80 | (c & 0x3F)); }
    else { out.push(0xE0 | (c >> 12), 0x80 | ((c >> 6) & 0x3F), 0x80 | (c & 0x3F)); }
  }
  return new Uint8Array(out);
}
function _u16(n) { return [n & 0xFF, (n >> 8) & 0xFF]; }
function _u32(n) { return [n & 0xFF, (n >> 8) & 0xFF, (n >> 16) & 0xFF, (n >> 24) & 0xFF]; }

function _zip(files) {
  // files: [{name, data: Uint8Array}]
  const local = [], central = [];
  let offset = 0;
  for (const f of files) {
    const name = _strBytes(f.name), data = f.data, crc = _crc32(data);
    const lf = [].concat(_u32(0x04034b50), _u16(20), _u16(0), _u16(0), _u16(0), _u16(0),
      _u32(crc), _u32(data.length), _u32(data.length), _u16(name.length), _u16(0));
    local.push(new Uint8Array(lf), name, data);
    const cf = [].concat(_u32(0x02014b50), _u16(20), _u16(20), _u16(0), _u16(0), _u16(0), _u16(0),
      _u32(crc), _u32(data.length), _u32(data.length), _u16(name.length), _u16(0), _u16(0),
      _u16(0), _u16(0), _u32(0), _u32(offset));
    central.push(new Uint8Array(cf), name);
    offset += 30 + name.length + data.length;
  }
  let centralLen = 0; for (const c of central) centralLen += c.length;
  const end = new Uint8Array([].concat(_u32(0x06054b50), _u16(0), _u16(0),
    _u16(files.length), _u16(files.length), _u32(centralLen), _u32(offset), _u16(0)));
  // concat all
  const parts = [...local, ...central, end];
  let total = 0; for (const p of parts) total += p.length;
  const out = new Uint8Array(total);
  let pos = 0; for (const p of parts) { out.set(p, pos); pos += p.length; }
  return out;
}

function _esc(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}
function _colLetter(n) { // 1-based -> A, B, ... AA
  let s = "";
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/* ---------- Paleta (idéntica al VBA M00_Config) ---------- */
// ARGB sin #, prefijo FF (opaco)
const FILL = {
  manana: "FFFFE699", tarde: "FFF8CBAD", noche: "FFB4A7D6",
  d12: "FFA9D08E", n12: "FF9DC3E6", doblaje: "FFF4B084", parcial: "FFD9D9D9",
  blanco: "FFFFFFFF", vacio: "FFFF0000",
  diaBloque: "FFA9D08E",     // verde: cualquier celda de jornada de día
  nitBloque: "FFD2BEE6",     // morado: cualquier celda de jornada de noche
  cabFecha: "FFC6E0B4",     // cabecera de fechas (verde claro)
  matiHdr: "FF8064A2",      // morado bloque mañana/noche (cabecera)
  tardeHdr: "FFED7D31",     // naranja bloque tarde
  areaBg: "FFF2F2F2",       // fondo etiqueta de área
  decHdr: "FFC00000",       // rojo cabecera decalajes
  decSub: "FFDCDCDC",       // gris subcabecera
  decOk: "FFC6E0B4",        // verde cubierto
  decKo: "FFFCCBCB",        // rojo claro sin cubrir
  // Colores de la hoja de la coordinadora "INF_PLANING SETMANAL"
  // (encabezados laterales de unidad y filas de fechas)
  planMati: "FFFFF2CC",     // mañana: amarillo claro
  planTarda: "FFCCFFCC",    // tarde: verde claro
  planNit: "FF99CCFF",      // noche: azul claro
};

/* ---------- Construcción de styles.xml ----------
   Cada combinación (fill, bold, border, align) es un xf en cellXfs.
   Definimos un set fijo de estilos y los referenciamos por índice.
*/
function _buildStyles() {
  // fills: 0=none, 1=gray125 (reservados), luego los nuestros en orden
  const fillList = [
    "none", "gray", FILL.manana, FILL.tarde, FILL.noche, FILL.d12, FILL.n12,
    FILL.doblaje, FILL.parcial, FILL.blanco, FILL.vacio, FILL.cabFecha,
    FILL.matiHdr, FILL.tardeHdr, FILL.areaBg, FILL.decHdr, FILL.decSub, FILL.decOk, FILL.decKo,
    FILL.diaBloque, FILL.nitBloque, FILL.planMati, FILL.planTarda, FILL.planNit,
  ];
  const fillIdx = {};
  fillList.forEach((c, i) => { if (c !== "none" && c !== "gray") fillIdx[c] = i; });

  let fillsXml = '<fill><patternFill patternType="none"/></fill>'
    + '<fill><patternFill patternType="gray125"/></fill>';
  for (let i = 2; i < fillList.length; i++) {
    fillsXml += `<fill><patternFill patternType="solid"><fgColor rgb="${fillList[i]}"/></patternFill></fill>`;
  }

  // fonts: 0 normal, 1 bold, 2 bold white, 3 white,
  //        4 encabezado lateral (Calibri 10 negrita), 5 fecha de día (Arial 12 negrita),
  //        6 fecha de noche (Arial 11 negrita), 7 etiqueta de posición (Calibri 8),
  //        8 etiqueta de posición de noche (Calibri 9)  — como en INF_PLANING SETMANAL
  const fonts =
    '<font><sz val="11"/><name val="Calibri"/></font>'
    + '<font><b/><sz val="11"/><name val="Calibri"/></font>'
    + '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
    + '<font><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
    + '<font><b/><sz val="10"/><name val="Calibri"/></font>'
    + '<font><b/><sz val="12"/><name val="Arial"/></font>'
    + '<font><b/><sz val="11"/><name val="Arial"/></font>'
    + '<font><sz val="8"/><name val="Calibri"/></font>'
    + '<font><sz val="9"/><name val="Calibri"/></font>';

  // borders: 0 none, 1 thin all
  const borders =
    '<border/>'
    + '<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right>'
    + '<top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom></border>';

  // cellXfs: definimos estilos nombrados -> índice
  const xfs = [];
  const styleIdx = {};
  function addXf(name, { font = 0, fill = 0, border = 0, hAlign = null, vAlign = null, wrap = false }) {
    const align = (hAlign || vAlign || wrap)
      ? `<alignment${hAlign ? ` horizontal="${hAlign}"` : ""}${vAlign ? ` vertical="${vAlign}"` : ""}${wrap ? ` wrapText="1"` : ""}/>`
      : "";
    const applyFill = fill > 1 ? ' applyFill="1"' : "";
    const applyBorder = border > 0 ? ' applyBorder="1"' : "";
    const applyAlign = align ? ' applyAlignment="1"' : "";
    xfs.push(`<xf numFmtId="0" fontId="${font}" fillId="${fill}" borderId="${border}" xfId="0"${applyFill}${applyBorder}${applyAlign}>${align}</xf>`);
    styleIdx[name] = xfs.length - 1;
  }

  addXf("normal", {});
  addXf("dateHdr", { font: 1, fill: fillIdx[FILL.cabFecha], border: 1, hAlign: "center" });
  addXf("matiHdr", { font: 2, fill: fillIdx[FILL.matiHdr], border: 1, hAlign: "left" });
  addXf("tardeHdr", { font: 2, fill: fillIdx[FILL.tardeHdr], border: 1, hAlign: "left" });
  addXf("areaName", { font: 1, fill: fillIdx[FILL.areaBg], border: 1, vAlign: "center" });
  addXf("areaNum", { fill: fillIdx[FILL.areaBg], border: 1, hAlign: "center" });
  // Estilo INF_PLANING SETMANAL: encabezado lateral de unidad, fila de fechas y
  // etiqueta de posición, para mañana (amarillo), tarde (verde) y noche (azul)
  addXf("latMati", { font: 4, fill: fillIdx[FILL.planMati], border: 1, hAlign: "left", vAlign: "center", wrap: true });
  addXf("latTarda", { font: 4, fill: fillIdx[FILL.planTarda], border: 1, hAlign: "left", vAlign: "center", wrap: true });
  addXf("latNit", { font: 4, fill: fillIdx[FILL.planNit], border: 1, hAlign: "left", vAlign: "center", wrap: true });
  addXf("fechaMati", { font: 5, fill: fillIdx[FILL.planMati], border: 1, hAlign: "center", vAlign: "center" });
  addXf("fechaTarda", { font: 5, fill: fillIdx[FILL.planTarda], border: 1, hAlign: "center", vAlign: "center" });
  addXf("fechaNit", { font: 6, fill: fillIdx[FILL.planNit], border: 1, hAlign: "center", vAlign: "center" });
  addXf("posDia", { font: 7, fill: fillIdx[FILL.blanco], border: 1, hAlign: "center", vAlign: "center", wrap: true });
  addXf("posNit", { font: 8, fill: fillIdx[FILL.blanco], border: 1, hAlign: "center", vAlign: "center", wrap: true });
  // celdas de datos por color de turno
  addXf("cManana", { fill: fillIdx[FILL.manana], border: 1, hAlign: "center" });
  addXf("cTarde", { fill: fillIdx[FILL.tarde], border: 1, hAlign: "center" });
  addXf("cNoche", { fill: fillIdx[FILL.noche], border: 1, hAlign: "center" });
  addXf("cD12", { fill: fillIdx[FILL.d12], border: 1, hAlign: "center" });
  addXf("cN12", { fill: fillIdx[FILL.n12], border: 1, hAlign: "center" });
  addXf("cDoblaje", { fill: fillIdx[FILL.doblaje], border: 1, hAlign: "center" });
  addXf("cParcial", { fill: fillIdx[FILL.parcial], border: 1, hAlign: "center" });
  addXf("cBlanco", { fill: fillIdx[FILL.blanco], border: 1, hAlign: "center" });
  addXf("cVacio", { fill: fillIdx[FILL.vacio], border: 1, hAlign: "center" });
  addXf("cDia", { fill: fillIdx[FILL.diaBloque], border: 1, hAlign: "center" });
  addXf("cNit", { fill: fillIdx[FILL.nitBloque], border: 1, hAlign: "center" });
  // decalajes
  addXf("decHdr", { font: 2, fill: fillIdx[FILL.decHdr], border: 1, hAlign: "center" });
  addXf("decSub", { font: 1, fill: fillIdx[FILL.decSub], border: 1, hAlign: "center" });
  addXf("decOk", { fill: fillIdx[FILL.decOk], border: 1, hAlign: "left" });
  addXf("decKo", { fill: fillIdx[FILL.decKo], border: 1, hAlign: "left" });
  addXf("decCell", { border: 1, hAlign: "center" });
  // informe / títulos
  addXf("titulo", { font: 1 });
  addXf("thGris", { font: 1, fill: fillIdx[FILL.decSub], border: 1, hAlign: "left" });
  addXf("tdNormal", { border: 1, hAlign: "left" });

  const styleSheet =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + `<fonts count="9">${fonts}</fonts>`
    + `<fills count="${fillList.length}">${fillsXml}</fills>`
    + `<borders count="2">${borders}</borders>`
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + `<cellXfs count="${xfs.length}">${xfs.join("")}</cellXfs>`
    + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    + '</styleSheet>';

  return { styleSheet, styleIdx };
}

/* ---------- Hoja de cálculo: filas/celdas ----------
   Modelo de celda: {v, s, t} -> valor, índice de estilo, tipo ('s' string | 'n' number)
*/
function _cell(ref, cell) {
  if (cell == null || cell.v === "" || cell.v == null) {
    return `<c r="${ref}" s="${cell ? cell.s : 0}"/>`;
  }
  if (cell.t === "n") {
    return `<c r="${ref}" s="${cell.s}"><v>${cell.v}</v></c>`;
  }
  return `<c r="${ref}" s="${cell.s}" t="inlineStr"><is><t xml:space="preserve">${_esc(cell.v)}</t></is></c>`;
}

function _sheetXml(grid, colWidths) {
  // grid: array de filas; cada fila array de celdas (o null)
  let rowsXml = "";
  for (let r = 0; r < grid.length; r++) {
    const row = grid[r] || [];
    let cellsXml = "";
    for (let c = 0; c < row.length; c++) {
      if (row[c] == null) continue;
      cellsXml += _cell(_colLetter(c + 1) + (r + 1), row[c]);
    }
    rowsXml += `<row r="${r + 1}">${cellsXml}</row>`;
  }
  let cols = "";
  if (colWidths && colWidths.length) {
    cols = '<cols>' + colWidths.map((w, i) =>
      `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("") + '</cols>';
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + cols
    + `<sheetData>${rowsXml}</sheetData>`
    + ((grid.merges && grid.merges.length)
        ? `<mergeCells count="${grid.merges.length}">${grid.merges.map(m => `<mergeCell ref="${m}"/>`).join("")}</mergeCells>`
        : "")
    + '</worksheet>';
}

/* ---------- Mapa de estilo de celda de dato según turno ---------- */
function _estiloPorTurno(turnoFinal, S) {
  switch (turnoFinal) {
    case "MAÑANA": return S.cBlanco;     // 7h -> blanco (como en el VBA)
    case "TARDE": return S.cBlanco;
    case "NOCHE": return S.cBlanco;
    case "PARCIAL": return S.cBlanco;
    case "12H DÍA": return S.cD12;
    case "12H NOCHE": return S.cN12;
    case "DOBLAJE": return S.cDoblaje;
    default: return S.cBlanco;
  }
}

/* ---------- Construye la grid de una semana (cuadrante + decalajes) ---------- */
const _DIAS = ["LUN", "MAR", "MIÉ", "JUE", "VIE", "SÁB", "DOM"];
const _BLOQUES = [
  { idx: 0, label: "MAÑANA", suf: "MATI", lat: "latMati", fecha: "fechaMati", pos: "posDia" },
  { idx: 1, label: "TARDE", suf: "TARDA", lat: "latTarda", fecha: "fechaTarda", pos: "posDia" },
  { idx: 2, label: "NOCHE", suf: "NIT", lat: "latNit", fecha: "fechaNit", pos: "posNit" },
];

function _gridSemana(sem, db, ENGINE, S, areasLayout) {
  const grid = [];
  const merges = [];   // celdas combinadas (encabezados laterales de unidad)
  grid.merges = merges;
  const fechas = Array.from({ length: 7 }, (_, i) => ENGINE.addDays(sem.fechaLunes, i));
  // turno y etiqueta por día/nombre
  const turnoMap = [];
  const etiqMap = [];
  for (let d = 0; d < 7; d++) {
    const m = new Map();
    const e = new Map();
    for (const r of sem.semana[d].regs) {
      m.set(ENGINE.normalizarTexto(r.nombre), r.turnoFinal);
      if (r.etiqueta) e.set(ENGINE.normalizarTexto(r.nombre), r.etiqueta);
    }
    turnoMap.push(m); etiqMap.push(e);
  }

  for (const b of _BLOQUES) {
    // fila de cabecera del bloque (col A) + fechas
    const cab = [];
    cab.push({ v: `${b.label} · ${b.suf}`, s: S[b.lat], t: "s" });
    cab.push({ v: "", s: S[b.fecha] }); // col B
    for (let d = 0; d < 7; d++) {
      const f = fechas[d];
      cab.push({ v: `${_DIAS[d]} ${f.getDate()}/${f.getMonth() + 1}`, s: S[b.fecha], t: "s" });
    }
    grid.push(cab);

    // filas por área
    areasLayout.forEach((area, areaIdx) => {
      let local = 1;
      for (let p = area.pIni; p <= area.pFin; p++) {
        const row = [];
        // col A: nombre de área solo en la primera fila
        // (la celda se combina con las de debajo, como en INF_PLANING SETMANAL)
        if (p === area.pIni) {
          row.push({ v: area.nombre + " " + b.suf, s: S[b.lat], t: "s" });
          if (area.pFin > area.pIni) merges.push(`A${grid.length + 1}:A${grid.length + 1 + area.pFin - area.pIni}`);
        }
        else row.push({ v: "", s: S[b.lat] });
        // col B: número local o etiqueta fija de la posición
        const etFija = ENGINE.etiquetaDePosicion(area, local);
        if (etFija && etFija.reemplazaNumero) {
          row.push({ v: etFija.texto, s: S[b.pos], t: "s" });
        } else if (etFija) {
          row.push({ v: local + " " + etFija.texto, s: S[b.pos], t: "s" });
        } else {
          row.push({ v: local, s: S[b.pos], t: "n" });
        }
        // cols C..I: color según el TIPO de jornada del profesional ese día
        // (12h día verde, 12h noche lila, 7h/10h/parcial blanco)
        for (let d = 0; d < 7; d++) {
          const nombre = sem.asig.celdas[d][b.idx][p];
          if (!nombre) {
            row.push({ v: "", s: S.cVacio });
          } else {
            const tf = turnoMap[d].get(ENGINE.normalizarTexto(nombre));
            const et = etiqMap[d].get(ENGINE.normalizarTexto(nombre));
            const suf = et === "doblaje" ? " (DBL)" : et === "+" ? " (+)" : et === "parcial" ? " (½)" : "";
            let estilo;
            const fondoProf = db.fondo(nombre);
            if (fondoProf) {
              // El color lo manda el FONDO del profesional (configurable en la plantilla)
              estilo = fondoProf === ENGINE.FONDO.LILA ? S.cNit : S.cBlanco;
            } else if (tf === ENGINE.TURNO.D12 || tf === ENGINE.TURNO.N12 || tf === ENGINE.TURNO.DOBLAJE) {
              estilo = S.cNit;   // no está en la base: 12h/doblaje -> lila
            } else {
              estilo = S.cBlanco; // no está en la base: 7h, 10h, parcial -> blanco
            }
            row.push({ v: db.apodo(nombre) + suf, s: estilo, t: "s" });
          }
        }
        grid.push(row);
        local++;
      }
      // Filas de EXCEDENTES de esta rotación (solo en la última área de la rotación,
      // para no duplicar UCCRI esq/dre). Posiciones extra sin número.
      const esUltimaDeRot = !areasLayout.slice(areaIdx + 1).some(a => a.rotacion === area.rotacion);
      if (esUltimaDeRot) {
        let maxExced = 0;
        for (let d = 0; d < 7; d++) {
          const lista = (sem.asig.excedentes[d * 3 + b.idx] || []).filter(e => e.area === area.rotacion);
          if (lista.length > maxExced) maxExced = lista.length;
        }
        for (let e = 0; e < maxExced; e++) {
          const row = [{ v: "", s: S.normal }, { v: "+", s: S[b.pos], t: "s" }];
          for (let d = 0; d < 7; d++) {
            const lista = (sem.asig.excedentes[d * 3 + b.idx] || []).filter(x => x.area === area.rotacion);
            const ex = lista[e];
            if (ex) {
              // El excedente lleva el MISMO fondo que tendría en una posición real
              const fEx = db.fondo(ex.nombre);
              row.push({ v: db.apodo(ex.nombre), s: fEx === ENGINE.FONDO.LILA ? S.cNit : S.cBlanco, t: "s" });
            }
            else row.push({ v: "", s: S.cBlanco }); // excedente vacío -> blanco (no es ubicación real)
          }
          grid.push(row);
        }
      }
      // fila en blanco de separación entre unidades (no tras la última)
      if (areaIdx < areasLayout.length - 1) grid.push([]);
    });
    // fila en blanco entre bloques
    grid.push([]);
  }

  // --- Panel de decalajes debajo ---
  if (sem.avisos && sem.avisos.length) {
    grid.push([]);
    grid.push([{ v: "AVISOS DE DECALAJE", s: S.decHdr, t: "s" }]);
    grid.push([
      { v: "FECHA", s: S.decSub, t: "s" }, { v: "POS", s: S.decSub, t: "s" },
      { v: "ÁREA", s: S.decSub, t: "s" }, { v: "DECALAJE", s: S.decSub, t: "s" },
      { v: "CUBIERTO POR", s: S.decSub, t: "s" },
    ]);
    for (const a of sem.avisos) {
      const est = a.cubierto ? S.decOk : S.decKo;
      grid.push([
        { v: ENGINE.fechaCorta(a.fecha), s: S.decCell, t: "s" },
        { v: a.pos, s: S.decCell, t: "s" },
        { v: a.area.replace("_", " "), s: S.decCell, t: "s" },
        { v: a.tipo, s: S.decCell, t: "s" },
        { v: a.cubierto ? db.apodo(a.cubiertoPor) : "— sin cubrir —", s: est, t: "s" },
      ]);
    }
  }

  return grid;
}

/* ---------- Hoja de informe mensual ---------- */
function _gridInforme(informe, S, ambito) {
  const grid = [];
  const filas = informe.filter(f => f.ambito === ambito);
  grid.push([
    { v: "Mes", s: S.thGris, t: "s" },
    { v: "Profesional", s: S.thGris, t: "s" },
    { v: "Unidad / rotación", s: S.thGris, t: "s" },
  ]);
  for (const f of filas) {
    grid.push([
      { v: f.mes, s: S.tdNormal, t: "s" },
      { v: f.profesional, s: S.tdNormal, t: "s" },
      { v: f.resumen + (f.roto ? "  [rota]" : ""), s: S.tdNormal, t: "s" },
    ]);
  }
  if (filas.length === 0) {
    grid.push([{ v: "(sin registros)", s: S.tdNormal, t: "s" }]);
  }
  return grid;
}

/* ---------- API principal ---------- */
function exportarCronogramaXLSX(resultado, db, ENGINE) {
  const { styleSheet, styleIdx: S } = _buildStyles();
  // Áreas reales del cuadrante (estructura dinámica)
  const est = resultado.estructura || ENGINE.estructuraActiva();
  const areasLayout = (est.areas || []).map(a => ({ id: a.id, nombre: a.nombre, rotacion: a.rotacion, pIni: a.posIni, pFin: a.posFin, etiquetas: a.etiquetas }));

  // Una hoja por semana + una hoja de informe
  const hojas = [];
  resultado.semanas.forEach((sem, i) => {
    const nombre = "Sem " + ENGINE.fechaCorta(sem.fechaLunes).replace(/[^\w]/g, "");
    hojas.push({ name: nombre.slice(0, 31), grid: _gridSemana(sem, db, ENGINE, S, areasLayout),
      widths: [20, 14, 20, 20, 20, 20, 20, 20, 20] });
  });
  hojas.push({ name: "Rotació DIA", grid: _gridInforme(resultado.informe, S, "DIA"),
    widths: [10, 34, 60] });
  hojas.push({ name: "Rotació NIT", grid: _gridInforme(resultado.informe, S, "NIT"),
    widths: [10, 34, 60] });

  // Construir partes del paquete
  const sheetOverrides = hojas.map((_, i) =>
    `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("");
  const contentTypes = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    + sheetOverrides
    + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    + '</Types>';

  const rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    + '</Relationships>';

  const wbSheets = hojas.map((h, i) =>
    `<sheet name="${_esc(h.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("");
  const workbook = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + `<sheets>${wbSheets}</sheets></workbook>`;

  const wbRelsList = hojas.map((_, i) =>
    `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("");
  const stylesRelId = hojas.length + 1;
  const wbRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + wbRelsList
    + `<Relationship Id="rId${stylesRelId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
    + '</Relationships>';

  const files = [
    { name: "[Content_Types].xml", data: _strBytes(contentTypes) },
    { name: "_rels/.rels", data: _strBytes(rels) },
    { name: "xl/workbook.xml", data: _strBytes(workbook) },
    { name: "xl/_rels/workbook.xml.rels", data: _strBytes(wbRels) },
    { name: "xl/styles.xml", data: _strBytes(styleSheet) },
  ];
  hojas.forEach((h, i) => {
    files.push({ name: `xl/worksheets/sheet${i + 1}.xml`, data: _strBytes(_sheetXml(h.grid, h.widths)) });
  });

  return _zip(files);
}

  return {
    TURNO, COLOR, AREA, CRIT, NPOS_TOTAL_actual,
    CATALOGO_AREAS, estructuraPorDefecto, fijarEstructura, estructuraActiva,
    etiquetaFijaPosicion, etiquetaDePosicion, etiquetasPorDefectoDeArea,
    FONDO, COLOR_LILA, COLOR_BLANCO, fondoPorDefectoDeTurno, normalizarFondo,
    REGLAS_DEFECTO, REGLAS_INFO, fijarReglas, reglasActivas,
    normalizarTexto, colorPorTurno, fechaCorta, mesKey, addDays, weekdayMonday,
    posGlobalAArea, nombreAreaCortoDePosGlobal,
    ProfesionalesDB, obtenerProfesionalesDelDia,
    extraerSemana, calcularEstadisticasSemana,
    construirAsignacionSemanal, generarCronograma, afinarCobertura,
    ajustarTurnosDobles, ajustarParciales, construirAvisosDecalaje,
    filasHistoricoDeSemana, regenerarInformeMensual, parsearAntiguedad,
    analizarImportacionRotaciones, filasHistoricoDesdeRotaciones,
    parsearUnidadesCelda, mesDesdeEncabezado,
    resumenHistoricoRotacion, unidadRotacionObligatoria, unidadesCandidatasProf,
    extraerNombresSinVincular,
    turnoNaturalCanon, clasificarJornadaNatural, UMBRAL_H,
    exportarCronogramaXLSX,
  };

})();

/* ============================================================
   Lectura de Excel
   ============================================================ */
// Planilla -> matriz de filas [ [colA, colB, ...colK], ... ]
function leerPlanilla(workbook) {
  // Busca la hoja de planilla: la que tenga "TURNO DE MAÑANA" en col D, o se llame "Planilla"
  let sheetName = workbook.SheetNames.find(n => /planilla/i.test(n));
  if (!sheetName) {
    // heurística: hoja con marcador de sección
    for (const n of workbook.SheetNames) {
      const ws = workbook.Sheets[n];
      const ref = ws["!ref"]; if (!ref) continue;
      const txt = XLSX.utils.sheet_to_csv(ws).slice(0, 5000);
      if (/TURNO DE MA/i.test(txt)) { sheetName = n; break; }
    }
  }
  if (!sheetName) sheetName = workbook.SheetNames[0];
  const ws = workbook.Sheets[sheetName];
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const rows = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row = [];
    for (let c = 0; c <= 10; c++) { // cols A..K
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell) { row.push(null); continue; }
      if (cell.t === "d") row.push({ __date__: cell.v instanceof Date ? cell.v.toISOString() : cell.v });
      else if (cell.t === "n" && c === 10) row.push(cell.v); // K como número serie posible
      else row.push(cell.v);
    }
    rows.push(row);
  }
  return { rows, sheetName };
}

// DATOS_PROFESIONALES -> array de objetos
function leerProfesionales(workbook) {
  const name = workbook.SheetNames.find(n => /datos_profesionales|profesionales/i.test(n));
  if (!name) return null;
  const ws = workbook.Sheets[name];
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const out = [];
  for (let r = 4; r <= range.e.r; r++) { // fila 5 (0-indexed 4) en adelante
    const g = (c) => { const cell = ws[XLSX.utils.encode_cell({ r, c })]; return cell ? cell.v : ""; };
    const nombre = g(0);
    if (!nombre || !String(nombre).trim()) continue;
    // Cols G y H (índices 6 y 7) opcionales: color de fondo y marca de Antiguo/a.
    // Si no vienen, el fondo se deduce del turno y Antiguo/a queda en "No".
    out.push({ nombre, turno: g(1), apodo: g(2), criterio: g(3), suplente: g(4),
               antiguedad: g(5), fondo: g(6), antiguo: g(7) });
  }
  return out;
}
/* Lee el libro de rotaciones de la coordinadora (hojas DIA y NIT).
   Fila 1 = encabezados de mes (col B en adelante). Col A = apodo del profesional.
   Devuelve [{ bloque, encabezados, filas: [{nombre, celdas}] }]. */
function leerRotacionesInf(workbook) {
  const hojas = [];
  for (const nombreHoja of workbook.SheetNames) {
    const esNit = /^\s*(NIT|NOCHE|NOCHES)\s*$/i.test(nombreHoja);
    const esDia = /^\s*(DIA|DÍA|DIAS|DÍAS)\s*$/i.test(nombreHoja);
    if (!esNit && !esDia) continue;
    const ws = workbook.Sheets[nombreHoja];
    if (!ws || !ws["!ref"]) continue;
    const range = XLSX.utils.decode_range(ws["!ref"]);
    const val = (r, c) => {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      return cell == null ? "" : (cell.w != null ? cell.w : cell.v);
    };
    // encabezados: fila 0, columnas 1..8 (B..I)
    const encabezados = [];
    const colsMes = [];
    for (let c = 1; c <= range.e.c && c <= 12; c++) {
      const h = String(val(0, c) || "").trim();
      encabezados.push(h);
      colsMes.push(c);
    }
    const filas = [];
    for (let r = 1; r <= range.e.r; r++) {
      const nombre = String(val(r, 0) || "").trim();
      if (!nombre) continue;
      const celdas = colsMes.map(c => String(val(r, c) || "").trim());
      filas.push({ nombre, celdas });
    }
    hojas.push({ bloque: esNit ? "NIT" : "DIA", nombreHoja, encabezados, filas });
  }
  return hojas;
}

function leerAlias(workbook) {
  const name = workbook.SheetNames.find(n => /alias/i.test(n));
  if (!name) return null;
  const ws = workbook.Sheets[name];
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const out = [];
  for (let r = 1; r <= range.e.r; r++) {
    const g = (c) => { const cell = ws[XLSX.utils.encode_cell({ r, c })]; return cell ? cell.v : ""; };
    const a = g(0);
    if (!a || !String(a).trim()) continue;
    out.push({ abreviatura: a, nombre: g(1) });
  }
  return out;
}

/* ============================================================
   Utilidades de fecha para inputs
   ============================================================ */
function lunesDeLaSemana(d) {
  const x = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const wd = x.getDay() === 0 ? 7 : x.getDay();
  x.setDate(x.getDate() - (wd - 1));
  return x;
}
function isoDate(d) {
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
}
function parseISO(s) {
  const [y, m, d] = s.split("-").map(Number);
  return new Date(y, m - 1, d);
}
function fechaLarga(d) {
  const dias = ["domingo","lunes","martes","miércoles","jueves","viernes","sábado"];
  const meses = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
  return dias[d.getDay()] + " " + d.getDate() + " " + meses[d.getMonth()] + " " + d.getFullYear();
}

/* ============================================================
   COMPONENTE PRINCIPAL
   ============================================================ */
const DIAS_CORTOS = ["LUN", "MAR", "MIÉ", "JUE", "VIE", "SÁB", "DOM"];

// Mezcla dos listas de alias [{abreviatura, nombre}] sin duplicar abreviaturas.
// Los del segundo argumento (nuevos) prevalecen sobre los del primero.
function mezclarAlias(previos, nuevos) {
  const map = new Map();
  for (const a of (previos || [])) if (a && a.abreviatura) map.set(String(a.abreviatura).toUpperCase().trim(), a.nombre);
  for (const a of (nuevos || [])) if (a && a.abreviatura) map.set(String(a.abreviatura).toUpperCase().trim(), a.nombre);
  return [...map.entries()].map(([abreviatura, nombre]) => ({ abreviatura, nombre }));
}

export default function App() {
  const [planilla, setPlanilla] = useState(null);
  const [planillaInfo, setPlanillaInfo] = useState(null);
  const [profs, setProfs] = useState(null);
  const [alias, setAlias] = useState(null);
  const [staffSource, setStaffSource] = useState(""); // de dónde vino la BD
  const [fechaLunes, setFechaLunes] = useState(isoDate(lunesDeLaSemana(new Date())));
  const [numSemanas, setNumSemanas] = useState(2);
  const [historico, setHistorico] = useState([]);
  const [estructura, setEstructura] = useState(null); // config de la unidad (Fase A)
  const [reglas, setReglas] = useState(null); // reglas de rotación activas
  const [resultado, setResultado] = useState(null);
  const [nombresSinVincular, setNombresSinVincular] = useState(null);
  const [importRot, setImportRot] = useState(null); // {hojas, fileName} del libro de rotaciones
  const [histGuardado, setHistGuardado] = useState(null); // aviso "guardado en el histórico"
  const [semanaActiva, setSemanaActiva] = useState(0);
  const [vista, setVista] = useState("inicio"); // inicio | cronograma | decalajes | informe | profesionales | editor
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  // Sesión / login
  const [session, setSession] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [cargandoDatos, setCargandoDatos] = useState(false);
  const [menuPerfil, setMenuPerfil] = useState(false);
  // Grupo profesional con el que se trabaja: null (aún sin elegir) | "ENF" | "TCAE".
  // Cada grupo tiene su plantilla, histórico, estructura, reglas y alias por separado.
  const [colectivo, setColectivo] = useState(null);
  const planillaInputRef = useRef(null);
  const staffInputRef = useRef(null);
  const rotacionesInputRef = useRef(null);

  // Vuelve a leer la sesión desde Supabase (p.ej. tras cambiar el apodo)
  async function refrescarSesion() {
    const s = await sesionActual();
    // clonar para garantizar nueva referencia y forzar el re-render
    setSession(s ? { ...s } : null);
  }

  // Comprobar sesión al arrancar y suscribirse a cambios (login/logout)
  useEffect(() => {
    let sub;
    (async () => {
      const s = await sesionActual();
      setSession(s);
      setAuthLoading(false);
      sub = onCambioSesion((nueva) => setSession(nueva));
    })();
    return () => { if (sub) sub.unsubscribe(); };
  }, []);

  // Cuando hay sesión y se ha elegido grupo profesional, cargar sus datos de Supabase.
  // Depende del id de usuario (no del objeto sesión) para que refrescos del token
  // o cambios de apodo no borren el trabajo en curso.
  const usuarioId = session?.user?.id || null;
  useEffect(() => {
    // Al cerrar sesión o cambiar de grupo, limpiar todo lo que hay en memoria
    setProfs(null); setAlias(null); setHistorico([]); setStaffSource("");
    setResultado(null); setVista("inicio"); setPlanilla(null); setPlanillaInfo(null);
    setEstructura(null); setReglas(null); setNombresSinVincular(null); setImportRot(null);
    setHistGuardado(null); setSemanaActiva(0); setError("");
    if (!usuarioId) { setColectivo(null); return; }
    if (!colectivo) return;
    fijarColectivo(colectivo);
    if (colectivo === "EST") return;   // Estudiantes no usa plantilla/histórico de cronograma
    (async () => {
      setCargandoDatos(true);
      setError("");
      try {
        const p = await cargarProfesionales();
        if (p && p.length) {
          setProfs(p);
          setStaffSource("base compartida · " + p.length + " profesionales");
        }
        const h = await cargarHistorico();
        if (h && h.length) setHistorico(h);
        // Alias de profesionales guardados (vínculos abreviatura → nombre)
        try {
          const al = await cargarAlias();
          if (al && al.length) setAlias(al);
        } catch { /* opcional */ }
        // Estructura de la unidad: la guardada, o la por defecto
        const est = await cargarEstructura();
        setEstructura(est && est.length ? est : ENGINE.estructuraPorDefecto());
        // Reglas de rotación: las guardadas, o todas activas por defecto
        try {
          const rg = await cargarReglas();
          setReglas(rg ? { ...ENGINE.REGLAS_DEFECTO, ...rg } : { ...ENGINE.REGLAS_DEFECTO });
        } catch { setReglas({ ...ENGINE.REGLAS_DEFECTO }); }
      } catch (err) {
        setError("No se pudieron cargar los datos: " + err.message);
      } finally {
        setCargandoDatos(false);
      }
    })();
  }, [usuarioId, colectivo]);

  // Detectar si el usuario ha llegado desde un enlace de recuperación de contraseña
  const [modoRecuperacion, setModoRecuperacion] = useState(false);
  useEffect(() => {
    const sub = escucharRecuperacion(() => setModoRecuperacion(true));
    return () => { if (sub) sub.unsubscribe(); };
  }, []);

  const db = useMemo(() => {
    if (!profs) return null;
    return new ENGINE.ProfesionalesDB(profs, alias || []);
  }, [profs, alias]);

  async function onPlanillaFile(e) {
    const file = e.target.files[0]; if (!file) return;
    setError(""); setBusy(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const { rows, sheetName } = leerPlanilla(wb);
      setPlanilla(rows);
      setPlanillaInfo({ name: file.name, sheet: sheetName, filas: rows.length });
      // ¿Trae también profesionales/alias embebidos? (p.ej. el libro v3)
      const p = leerProfesionales(wb);
      const al = leerAlias(wb);
      if (al && al.length) setAlias(prev => mezclarAlias(prev, al));
      // Solo poblar la base compartida si todavía está vacía
      if (p && p.length && (!profs || profs.length === 0)) {
        const guardados = await guardarProfesionales(p);
        setProfs(guardados.length ? guardados : p);
        setStaffSource("base compartida · cargada de " + file.name);
      }
    } catch (err) {
      setError("No se pudo leer la planilla: " + err.message);
    } finally { setBusy(false); }
  }

  async function onStaffFile(e) {
    const file = e.target.files[0]; if (!file) return;
    setError(""); setBusy(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const p = leerProfesionales(wb);
      const al = leerAlias(wb);
      if (!p || !p.length) { setError("Ese archivo no contiene la hoja DATOS_PROFESIONALES."); return; }
      if (al && al.length) setAlias(prev => mezclarAlias(prev, al));
      const guardados = await guardarProfesionales(p);
      setProfs(guardados.length ? guardados : p);
      setStaffSource("base compartida · cargada de " + file.name);
    } catch (err) {
      setError("No se pudo guardar la base de profesionales: " + err.message);
    } finally { setBusy(false); }
  }

  const puedeGenerar = planilla && db;

  // Generación: primero comprueba nombres sin vincular. Si los hay, abre el modal
  // de vinculación; si no, genera directamente.
  function generar() {
    if (!puedeGenerar) return;
    setError("");
    let sinVincular = [];
    try {
      const lunes = parseISO(fechaLunes);
      const ns = Math.max(1, Math.min(12, +numSemanas));
      sinVincular = ENGINE.extraerNombresSinVincular(planilla, db, lunes, ns);
    } catch (err) {
      setError("Error al analizar la planilla: " + err.message);
      return;
    }
    if (sinVincular.length > 0) {
      setNombresSinVincular(sinVincular);
      setVista("vincular");
    } else {
      ejecutarGeneracion();
    }
  }

  function ejecutarGeneracion() {
    setError(""); setBusy(true); setNombresSinVincular(null);
    setTimeout(() => {
      try {
        const lunes = parseISO(fechaLunes);
        const ns = Math.max(1, Math.min(12, +numSemanas));
        // Histórico previo (de Supabase) para que la rotación vea lo ya ocurrido.
        // Se excluye SOLO el rango de días que se va a (re)generar ahora; así las
        // semanas anteriores del mismo mes (p.ej. las sincronizadas con la hoja
        // real INF_PLANING SETMANAL) sí cuentan para continuidad y recuento mensual.
        const desdeGen = diaSerial(lunes);
        const hastaGen = desdeGen + ns * 7 - 1;
        const histPrevio = (Array.isArray(historico) ? historico : [])
          .filter(h => h.fechaSerial < desdeGen || h.fechaSerial > hastaGen);
        const out = ENGINE.generarCronograma(planilla, db, histPrevio, lunes, ns, estructura, reglas || ENGINE.REGLAS_DEFECTO);
        setResultado(out);
        setSemanaActiva(0);
        setVista("cronograma");
        // Volcado AUTOMÁTICO al histórico: el cuadrante generado queda registrado
        // sin depender de que nadie pulse un botón. Reemplaza exactamente las
        // semanas generadas, así que regenerar no duplica ni pierde nada.
        guardarEnHistorico(out, true);
      } catch (err) {
        setError("Error al generar: " + err.message);
        console.error(err);
      } finally { setBusy(false); }
    }, 30);
  }

  // Aplica las vinculaciones decididas por el usuario en el modal:
  //  - aliasNuevos: { rawNombre: nombreCompletoExistente }  -> añade alias
  //  - profsNuevos: [ {nombre, turno, apodo, criterio, suplente, antiguedad} ] -> añade a BD
  async function aplicarVinculaciones(aliasNuevos, profsNuevos) {
    setBusy(true); setError("");
    try {
      // 1) Añadir nuevos profesionales a la base (si hay)
      if (profsNuevos && profsNuevos.length > 0) {
        const nuevosTodos = [...(profs || []), ...profsNuevos];
        const guardados = await guardarProfesionales(nuevosTodos);
        setProfs(guardados && guardados.length ? guardados : nuevosTodos);
      }
      // 2) Añadir alias (vínculos a profesionales existentes)
      if (aliasNuevos && Object.keys(aliasNuevos).length > 0) {
        const nuevosAlias = Object.entries(aliasNuevos).map(([abreviatura, nombre]) => ({ abreviatura, nombre }));
        const todosAlias = [...(alias || []), ...nuevosAlias];
        setAlias(todosAlias);
        try { await guardarAliasTodos(todosAlias); } catch (e) { /* persistencia opcional */ }
      }
      setNombresSinVincular(null);
      // Tras vincular, generar directamente (ya están en profs/alias).
      setTimeout(() => ejecutarGeneracion(), 80);
    } catch (err) {
      setError("No se pudieron guardar las vinculaciones: " + err.message);
      setBusy(false);
    }
  }

  /* Guarda en el histórico las semanas del cuadrante que hay en pantalla.
     Reemplaza EXACTAMENTE el rango de días generado (no el mes entero), para
     no borrar semanas del mismo mes que no se estén regenerando.
     `silencioso` = se llama solo tras generar; no molesta con mensajes de éxito. */
  async function guardarEnHistorico(resultadoOpc, silencioso) {
    const res = resultadoOpc || resultado;
    if (!res) return false;
    if (!silencioso) setError("");
    const nuevasFilas = res.semanas.flatMap(s => ENGINE.filasHistoricoDeSemana(s.asig));
    if (nuevasFilas.length === 0) return false;
    // Rango de días que abarca lo generado (primer día ... último día)
    let minS = Infinity, maxS = -Infinity;
    for (const f of nuevasFilas) {
      if (f.fechaSerial < minS) minS = f.fechaSerial;
      if (f.fechaSerial > maxS) maxS = f.fechaSerial;
    }
    try {
      await borrarHistoricoDeRango(minS, maxS);
      await añadirAlHistorico(nuevasFilas);
      // Se recarga de la base: los días ya sincronizados con la hoja real están
      // protegidos y no se sobrescriben, así que lo que queda guardado puede no
      // coincidir con lo generado.
      try { setHistorico(await cargarHistorico()); }
      catch (e) {
        const resto = (historico || []).filter(h => h.fechaSerial < minS || h.fechaSerial > maxS);
        setHistorico(resto.concat(nuevasFilas));
      }
      setHistGuardado({ n: nuevasFilas.length, cuando: Date.now() });
      return true;
    } catch (err) {
      setError("No se pudo guardar el histórico: " + err.message);
      return false;
    }
  }
  async function limpiarHistorico() {
    setError("");
    try {
      await vaciarHistorico();
      setHistorico([]);
    } catch (err) {
      setError("No se pudo vaciar el histórico: " + err.message);
    }
  }

  // --- Importar el registro de rotación de la coordinadora (libro ROTACIONS_INF) ---
  async function onRotacionesFile(e) {
    const file = e.target.files[0];
    e.target.value = ""; // permitir volver a elegir el mismo archivo
    if (!file) return;
    setError(""); setBusy(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array", cellDates: true });
      const hojas = leerRotacionesInf(wb);
      if (!hojas.length) {
        setError("Ese archivo no tiene hojas llamadas DIA o NIT. Revisa que sea el libro de rotaciones.");
        return;
      }
      setImportRot({ hojas, fileName: file.name });
      setVista("importar");
    } catch (err) {
      setError("No se pudo leer el archivo de rotaciones: " + err.message);
    } finally { setBusy(false); }
  }

  // Guarda las filas importadas como histórico, reemplazando esos meses.
  async function aplicarImportacionRotaciones(filas, meses) {
    setBusy(true); setError("");
    try {
      if (meses && meses.length) await borrarHistoricoDeMeses(meses);
      await añadirAlHistorico(filas);
      const setMeses = new Set(meses || []);
      const resto = (historico || []).filter(h => !setMeses.has(h.mes));
      setHistorico(resto.concat(filas));
      setImportRot(null);
      setVista("historico");
    } catch (err) {
      setError("No se pudo guardar la importación: " + err.message);
    } finally { setBusy(false); }
  }

  // Descarga el cronograma como .xlsx con los colores del Excel original,
  // listo para subir a Google Sheets (Abrir con > Hojas de cálculo de Google).
  function exportarExcel() {
    if (!resultado) return;
    try {
      const bytes = ENGINE.exportarCronogramaXLSX(resultado, db, ENGINE);
      const blob = new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
      const url = URL.createObjectURL(blob);
      const lunes = parseISO(fechaLunes);
      const nombre = "Cuadrante_" + isoDate(lunes) + ".xlsx";
      const a = document.createElement("a");
      a.href = url; a.download = nombre;
      document.body.appendChild(a); a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 1500);
    } catch (err) {
      setError("No se pudo exportar: " + err.message);
      console.error(err);
    }
  }

  // Guarda la plantilla editada en la base compartida (Supabase).
  // Como `db` depende de [profs, alias] vía useMemo, la próxima generación usa
  // automáticamente los datos modificados.
  async function saveProfs(nuevosProfs, nuevoSource) {
    const guardados = await guardarProfesionales(nuevosProfs);
    const finales = guardados.length ? guardados : nuevosProfs.filter(p => p.nombre && p.nombre.trim());
    setProfs(finales);
    const src = nuevoSource || "base compartida · " + finales.length + " profesionales";
    setStaffSource(src);
  }

  const verEditor = () => setVista("editor");
  const verEstructura = () => setVista("estructura");
  const verHistorico = () => setVista("historico");

  // Guarda la estructura de la unidad en la base compartida.
  async function saveEstructura(nuevaConfig) {
    await guardarEstructura(nuevaConfig);
    setEstructura(nuevaConfig);
  }

  async function handleLogout() {
    setMenuPerfil(false);
    await cerrarSesion();
    // el efecto de [session] limpia los datos en memoria
  }

  // --- Puerta de acceso: mientras carga la sesión, o si no hay sesión ---
  if (authLoading) {
    return (
      <div className="app">
        <style>{CSS}</style>
        <div className="bootscreen">Cargando…</div>
      </div>
    );
  }
  // Si viene de un enlace de recuperación, pedir nueva contraseña antes de nada
  if (modoRecuperacion) {
    return (
      <div className="app">
        <style>{CSS}</style>
        <NuevaPasswordScreen
          onGuardar={cambiarPassword}
          onListo={() => setModoRecuperacion(false)}
        />
      </div>
    );
  }
  if (!session) {
    return (
      <div className="app">
        <style>{CSS}</style>
        <LoginScreen onLogin={iniciarSesion} onRecuperar={recuperarPassword} />
      </div>
    );
  }

  const apodoUsuario = apodoDeSesion(session);

  const cabecera = (
      <Header
        colectivo={colectivo}
        onCambiarColectivo={() => { setMenuPerfil(false); setColectivo(null); }}
        email={session?.user?.email}
        apodo={apodoUsuario}
        abierto={menuPerfil}
        onToggle={() => setMenuPerfil(v => !v)}
        onCerrarMenu={() => setMenuPerfil(false)}
        onLogout={handleLogout}
        onCambiarApodo={cambiarApodo}
        onCambiarPassword={cambiarPassword}
        onSesionActualizada={refrescarSesion}
      />
  );

  // Pantalla de inicio: elegir el grupo profesional con el que se va a trabajar
  if (colectivo === "EST") {
    return (
      <div className="app">
        <style>{CSS}</style>
        {cabecera}
        <EstudiantesPanel />
      </div>
    );
  }

  if (!colectivo) {
    return (
      <div className="app">
        <style>{CSS}</style>
        {cabecera}
        <SelectorColectivo nombre={apodoUsuario || (session?.user?.email || "").split("@")[0]} onElegir={setColectivo} />
      </div>
    );
  }

  return (
    <div className="app">
      <style>{CSS}</style>
      {cabecera}
      <div className="layout">
        <Sidebar
          planillaInfo={planillaInfo} staffSource={staffSource} profs={profs}
          fechaLunes={fechaLunes} setFechaLunes={setFechaLunes}
          numSemanas={numSemanas} setNumSemanas={setNumSemanas}
          onPlanillaClick={() => planillaInputRef.current.click()}
          onStaffClick={() => staffInputRef.current.click()}
          generar={generar} puedeGenerar={puedeGenerar} busy={busy}
          historicoLen={historico.length}
          resultado={resultado}
          guardarEnHistorico={guardarEnHistorico}
          limpiarHistorico={limpiarHistorico}
          verEditor={verEditor} vistaActiva={vista}
          verEstructura={verEstructura}
          verReglas={() => setVista("reglas")} reglas={reglas}
          estructura={estructura}
          verHistorico={verHistorico}
        />
        <input ref={planillaInputRef} type="file" accept=".xlsx,.xlsm,.xls" hidden onChange={onPlanillaFile} />
        <input ref={staffInputRef} type="file" accept=".xlsx,.xlsm,.xls" hidden onChange={onStaffFile} />
        <input ref={rotacionesInputRef} type="file" accept=".xlsx,.xlsm,.xls" hidden onChange={onRotacionesFile} />

        <main className="main">
          {error && <div className="error">{error}</div>}
          {histGuardado && vista === "cronograma" && (
            <div className="saved-note" style={{ marginBottom: 14 }}>
              Cuadrante registrado en el histórico de rotación automáticamente
              ({histGuardado.n} registros). Lo ven todas las personas que entran en la web.
            </div>
          )}

          {vista === "estructura" && (
            <EditorEstructura
              estructura={estructura || ENGINE.estructuraPorDefecto()}
              catalogo={ENGINE.CATALOGO_AREAS}
              porDefecto={ENGINE.estructuraPorDefecto()}
              saveEstructura={saveEstructura}
              ENGINE={ENGINE}
              onCerrar={() => setVista(resultado ? "cronograma" : "inicio")}
            />
          )}

          {vista === "reglas" && (
            <EditorReglas
              reglas={reglas} ENGINE={ENGINE}
              saveReglas={async (r) => { await guardarReglas(r); setReglas(r); }}
              onCerrar={() => setVista(resultado ? "cronograma" : "inicio")}
            />
          )}

          {vista === "historico" && (
            <HistoricoRotacion
              historico={historico} ENGINE={ENGINE}
              onImportar={() => rotacionesInputRef.current.click()}
              onCerrar={() => setVista(resultado ? "cronograma" : "inicio")}
            />
          )}

          {vista === "importar" && importRot && (
            <ImportarRotaciones
              datos={importRot} db={db} profs={profs} ENGINE={ENGINE} busy={busy}
              onAplicar={aplicarImportacionRotaciones}
              onCancelar={() => { setImportRot(null); setVista("historico"); }}
            />
          )}

          {vista === "vincular" && nombresSinVincular && (
            <VincularProfesionales
              nombres={nombresSinVincular} profs={profs} ENGINE={ENGINE}
              busy={busy}
              onAplicar={aplicarVinculaciones}
              onCancelar={() => { setNombresSinVincular(null); setVista(resultado ? "cronograma" : "inicio"); }}
            />
          )}

          {vista === "editor" && (
            <EditorPlantilla
              profs={profs} saveProfs={saveProfs}
              onCerrar={() => setVista(resultado ? "cronograma" : "inicio")}
              hayResultado={!!resultado}
              onCargarExcel={() => staffInputRef.current.click()}
            />
          )}

          {vista !== "editor" && vista !== "estructura" && vista !== "historico" && vista !== "vincular" && vista !== "reglas" && vista !== "importar" && !resultado && <Welcome planilla={planilla} db={db} />}

          {vista !== "editor" && vista !== "estructura" && vista !== "historico" && vista !== "vincular" && vista !== "reglas" && vista !== "importar" && resultado && (
            <>
              <Tabs vista={vista} setVista={setVista}
                semanas={resultado.semanas} semanaActiva={semanaActiva} setSemanaActiva={setSemanaActiva}
                exportarExcel={exportarExcel} />
              {vista === "cronograma" && (
                <Cronograma semana={resultado.semanas[semanaActiva]} db={db} ENGINE={ENGINE} planilla={planilla} estructura={resultado.estructura}
                  cargarAnotaciones={cargarAnotaciones} guardarAnotaciones={guardarAnotaciones} />
              )}
              {vista === "decalajes" && (
                <Decalajes avisos={resultado.semanas[semanaActiva].avisos} db={db} ENGINE={ENGINE} />
              )}
              {vista === "profesionales" && <Profesionales db={db} />}
            </>
          )}
        </main>
      </div>
    </div>
  );
}

/* ---------- Cabecera con menú de perfil ---------- */
function iniciales(texto) {
  const t = String(texto || "").trim();
  if (!t) return "·";
  const partes = t.split(/[\s@.]+/).filter(Boolean);
  if (partes.length >= 2) return (partes[0][0] + partes[1][0]).toUpperCase();
  return t.slice(0, 2).toUpperCase();
}

/* ---------- Pantalla de inicio: elegir grupo profesional ---------- */
function SelectorColectivo({ nombre, onElegir }) {
  const opciones = [
    { id: "ENF", titulo: "Enfermería", desc: "Cuadrante y rotación de las enfermeras de la UCI.",
      icono: <path d="M16 7v18M7 16h18" stroke="currentColor" strokeWidth="2.6" strokeLinecap="round" /> },
    { id: "TCAE", titulo: "TCAE", desc: "Cuadrante y rotación de las técnicas auxiliares en cuidados de enfermería.",
      icono: <g fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><circle cx="16" cy="11" r="4.5" /><path d="M7.5 25c1.2-4.6 4.4-7 8.5-7s7.3 2.4 8.5 7" /></g> },
    { id: "EST", titulo: "Estudiantes", desc: "Asignación de estudiantes de enfermería de grado y de máster universitario.",
      icono: <g fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><path d="M3.5 12.5 16 7l12.5 5.5L16 18z" /><path d="M9 15.5v5.2c0 1.6 3.1 3.3 7 3.3s7-1.7 7-3.3v-5.2" /><path d="M28.5 12.5v7" /></g> },
  ];
  return (
    <div className="selector-wrap">
      <div className="selector-head">
        <h1>{nombre ? "Hola, " + nombre : "Hola"}</h1>
        <p>¿Con qué apartado vas a trabajar?</p>
      </div>
      <div className="selector-grid">
        {opciones.map(o => (
          <button key={o.id} className={"selector-card " + (o.id === "TCAE" ? "is-tcae" : "is-enf")} onClick={() => onElegir(o.id)}>
            <span className="selector-icon"><svg viewBox="0 0 32 32" width="30" height="30">{o.icono}</svg></span>
            <span className="selector-title">{o.titulo}</span>
            <span className="selector-desc">{o.desc}</span>
            <span className="selector-go">Entrar →</span>
          </button>
        ))}
      </div>
      <p className="selector-nota">Cada apartado guarda sus datos por separado. Puedes cambiar de apartado en cualquier momento desde la barra superior.</p>
    </div>
  );
}

/* ============================================================
   APARTADO ESTUDIANTES
   Sustituye el Excel «Infermeras pràctiques» que la coordinadora llevaba a
   mano: periodo de prácticas, enfermera tutora, turno, estudiante, escuela,
   curso y contacto. Muestra cuántos estudiantes hay cada mes (grado / máster,
   curso y turno) y los periodos con sus parejas tutora–estudiante.
   ============================================================ */
const EST_MODALIDADES = [
  { id: "DIA_12H", label: "Día 12 h", color: "#f0b44c" },
  { id: "NIT_12H", label: "Noche 12 h", color: "#5b6fb3" },
  { id: "NIT_10H", label: "Noche 10 h", color: "#8fa6dd" },
  { id: "MANANA", label: "Mañana", color: "#f5d77a" },
  { id: "TARDE", label: "Tarde", color: "#86c08a" },
  { id: "", label: "Sin indicar", color: "#cfd5dd" },
];
const EST_CURSOS = ["1º", "2º", "3º", "4º", "Máster"];
const EST_COLOR = { GRADO: "#2f7d6b", MASTER: "#c27c2c" };
const EST_MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"];
const EST_MESES_LARGO = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
const estMod = (id) => EST_MODALIDADES.find(m => m.id === (id || "")) || EST_MODALIDADES[EST_MODALIDADES.length - 1];
const estTipoDeCurso = (c) => (c === "Máster" ? "MASTER" : "GRADO");
const estNombre = (e) => [e.nombre, e.apellido1, e.apellido2].filter(Boolean).join(" ");
const estFechaCorta = (iso) => { const [y, m, d] = iso.split("-"); return `${+d} ${EST_MESES[+m - 1].toLowerCase()} ${y}`; };

const estPref = (anio, mes) => `${anio}-${String(mes + 1).padStart(2, "0")}-`;
// Días de un estudiante del calendario del máster dentro de un mes
const estDiasMes = (e, anio, mes) => (Array.isArray(e.dias) ? e.dias.filter(d => d.f.startsWith(estPref(anio, mes))) : []);
// Turno de un estudiante en un mes: el suyo, o el más frecuente de sus días ese mes
function estModMes(e, anio, mes) {
  const ds = estDiasMes(e, anio, mes);
  if (!ds.length) return e.modalidad || "";
  const c = {}; for (const d of ds) c[d.t] = (c[d.t] || 0) + 1;
  return Object.keys(c).sort((a, b) => c[b] - c[a])[0];
}
function estEnMes(e, anio, mes) {
  if (e.origen === "CAL_MASTER" && Array.isArray(e.dias) && e.dias.length) return estDiasMes(e, anio, mes).length > 0;
  const ini = `${anio}-${String(mes + 1).padStart(2, "0")}-01`;
  const fin = `${anio}-${String(mes + 1).padStart(2, "0")}-${String(new Date(anio, mes + 1, 0).getDate()).padStart(2, "0")}`;
  return e.inicio <= fin && e.fin >= ini;
}

/* ---------- Lectura de Excel (formato de la coordinadora o plantilla) ---------- */
function estNormTxt(v) { return String(v == null ? "" : v).normalize("NFD").replace(/[̀-ͯ]/g, "").toUpperCase().replace(/\s+/g, " ").trim(); }
function estModalidad(v) {
  const t = estNormTxt(v);
  if (!t) return "";
  const noche = /NIT|NOCHE|NOCT|\bTN\b/.test(t);
  if (t.includes("12")) return noche ? "NIT_12H" : "DIA_12H";
  if (t.includes("10")) return noche ? "NIT_10H" : "DIA_12H";
  if (noche || t === "TN" || t === "N") return "NIT_10H";
  if (/MAN|MATI/.test(t) || t === "TM" || t === "M") return "MANANA";
  if (/TARD/.test(t) || t === "TT" || t === "T") return "TARDE";
  if (/DIA/.test(t)) return "DIA_12H";
  return "";
}
function estCurso(v) {
  const t = estNormTxt(v);
  if (!t) return "";
  if (t.startsWith("MAST")) return "Máster";
  if (/QUART|CUART|4/.test(t)) return "4º";
  if (/TERCER|3/.test(t)) return "3º";
  if (/SEGON|SEGUN|2/.test(t)) return "2º";
  if (/PRIMER|1/.test(t)) return "1º";
  return "";
}
function estFecha(v) {
  if (v instanceof Date && !isNaN(v)) return `${v.getFullYear()}-${String(v.getMonth() + 1).padStart(2, "0")}-${String(v.getDate()).padStart(2, "0")}`;
  if (typeof v === "number" && v > 20000) { const d = new Date(Math.round((v - 25569) * 86400000)); return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`; }
  const s = String(v || "").trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return `${m[1]}-${m[2].padStart(2, "0")}-${m[3].padStart(2, "0")}`;
  m = s.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/); if (m) { const y = m[3].length === 2 ? "20" + m[3] : m[3]; return `${y}-${m[2].padStart(2, "0")}-${m[1].padStart(2, "0")}`; }
  return "";
}
const estTexto = (v) => (v == null ? "" : String(typeof v === "number" ? Math.round(v) : v).replace(/\s+/g, " ").trim());
/* ---------- Excel de estudiantes ----------
   Entiende dos formatos:
   1) El Excel que llevaba la coordinadora a mano (PERIODE, INFERMERA, TORN,
      NOM, COGNOM 1/2, ESCOLA, CURS, MAIL).
   2) Las ofertas de plazas que envían las universidades (UDPT, UIC…):
      ÀREA, UNITAT, TORN, horario, TUTORA CLÍNICA (1)/(2), CURS, PRACTICUM,
      DATA INICI/FI, CENTRE, TUTORA ACADÈMICA, NOM E, COGNOM 1/2 E, correo y
      teléfono. De estas SOLO se cogen los estudiantes del área de CRÍTICS.
      Si el archivo trae una hoja de horarios de críticos («horaris 12h
      critics»), se añaden los días concretos de cada estudiante. */
function estLeerExcel(wb) {
  const filas = [], errores = [];
  let excluidos = 0;
  for (const hoja of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[hoja], { header: 1, raw: true, defval: "" });
    let h = -1, col = {};
    for (let r = 0; r < Math.min(rows.length, 15) && h < 0; r++) {
      const c = {};
      rows[r].forEach((v, i) => {
        const t = estNormTxt(v);
        if (!t) return;
        if (t.startsWith("PERIODE") || t.startsWith("PERIODO")) { if (c.inicio == null) c.inicio = i; else if (c.fin == null) c.fin = i; }
        else if (t.startsWith("DATA INICI") || t.startsWith("FECHA INICIO") || t.startsWith("INICI")) c.inicio = i;
        else if (t.startsWith("DATA FI") || t.startsWith("FECHA FIN") || t === "FIN" || t.startsWith("FIN ") || t.startsWith("FI ")) c.fin = i;
        else if (t.includes("ACADEM")) c.tutoraAcad = i;
        else if (t.startsWith("TUTORA CLINICA (2)") || t.startsWith("TUTOR CLINICO (2)")) c.tutora2 = i;
        else if (t.startsWith("INFERMER") || t.startsWith("ENFERMER") || t.startsWith("TUTOR")) { if (c.tutora == null) c.tutora = i; }
        else if (t.startsWith("HORARI INICI") || t.startsWith("HORARIO INICIO")) c.horaIni = i;
        else if (t.startsWith("TORN") || t.startsWith("TURNO") || t.startsWith("MODALIDAD")) c.modalidad = i;
        else if (t.startsWith("COGNOM 1") || t.startsWith("APELLIDO 1") || t === "COGNOM" || t === "APELLIDO") c.apellido1 = i;
        else if (t.startsWith("COGNOM 2") || t.startsWith("APELLIDO 2")) c.apellido2 = i;
        else if (t === "NOM" || t.startsWith("NOM ") || t.startsWith("NOMBRE")) c.nombre = i;
        else if (t.startsWith("ESCOLA") || t.startsWith("ESCUELA") || t.startsWith("UNIVERS") || t === "CENTRE" || t === "CENTRO") c.escuela = i;
        else if (t.startsWith("CURS")) c.curso = i;
        else if (t.startsWith("PRACTICUM")) c.practicum = i;
        else if (t.startsWith("MAIL") || t.startsWith("CORREU") || t.startsWith("CORREO") || t.startsWith("CONTACT")) c.contacto = i;
        else if (t.startsWith("TELEF")) c.telefono = i;
        else if (t === "AREA" || t.startsWith("AREA ")) c.area = i;
        else if (t.startsWith("UNITAT") || t.startsWith("UNIDAD") || t.startsWith("SERVEI")) c.unidad = i;
        else if (t.startsWith("NOTA") || t.startsWith("OBSERV")) c.notas = i;
      });
      if (c.inicio != null && c.fin != null && c.nombre != null) { h = r; col = c; }
    }
    if (h < 0) continue;   // hoja sin listado de estudiantes (p.ej. hojas de horarios)
    for (let r = h + 1; r < rows.length; r++) {
      const row = rows[r] || [];
      const g = (k) => (col[k] != null ? row[col[k]] : "");
      const nombre = estTexto(g("nombre"));
      const inicio = estFecha(g("inicio")), fin = estFecha(g("fin"));
      if (!nombre) continue;                         // filas vacías o títulos («PRACTIQUES 2026»)
      // Solo el área de Crítics cuando el archivo trae área o unidad
      if (col.area != null || col.unidad != null) {
        const zona = estNormTxt(g("area") + " " + g("unidad"));
        if (!zona.includes("CRITIC")) { excluidos++; continue; }
      }
      const fallo = !inicio ? "falta la fecha de inicio" : !fin ? "falta la fecha de fin" : fin < inicio ? "el fin es anterior al inicio" : "";
      if (fallo) { errores.push(`${nombre} ${estTexto(g("apellido1"))} (fila ${r + 1}): ${fallo}.`); continue; }
      const curso = estCurso(g("curso"));
      const horaIni = g("horaIni");
      const empiezaNoche = (horaIni instanceof Date ? horaIni.getHours() : typeof horaIni === "number" ? Math.round(horaIni * 24) : parseInt(String(horaIni), 10)) >= 19;
      let modalidad = estModalidad(g("modalidad"));
      if (modalidad === "DIA_12H" && empiezaNoche) modalidad = "NIT_12H";
      const tutoras = [estTexto(g("tutora")), estTexto(g("tutora2"))].filter(Boolean).join(" / ");
      const notas = [estTexto(g("notas")), g("practicum") ? "Pràcticum " + estTexto(g("practicum")) : "", g("tutoraAcad") ? "Tutora académica: " + estTexto(g("tutoraAcad")) : ""].filter(Boolean).join(" · ");
      filas.push({
        nombre, apellido1: estTexto(g("apellido1")), apellido2: estTexto(g("apellido2")),
        curso, tipo: estTipoDeCurso(curso), escuela: estTexto(g("escuela")).toUpperCase(),
        tutora: tutoras, modalidad, inicio, fin,
        contacto: [estTexto(g("contacto")), estTexto(g("telefono"))].filter(Boolean).join(" · "), notas,
      });
    }
  }
  // Horarios día a día (hojas «horaris … critics»)
  const asignados = estLeerHorarios(wb, filas);
  return { filas, errores, excluidos, conHorario: asignados };
}

// Hoja(s) de horarios de críticos: una fila por estudiante (HORARI, ÀREA,
// ESTUDIANT, TUTORA CLÍNICA) y una columna por día con «12» los días que va.
// Las fechas salen de una fila de fechas encima o, si la cabecera trae solo
// números de día, se deducen a partir de la fecha de inicio del estudiante.
function estLeerHorarios(wb, filas) {
  let asignados = 0;
  const clave = (e) => estNormTxt(e.tutora.split("/")[0]) + "|" + estNormTxt(e.nombre).split(" ")[0];
  const porTutora = new Map(filas.map(e => [clave(e), e]));
  const porNombre = new Map(filas.map(e => [estNormTxt(estNombre(e)), e]));
  for (const hoja of wb.SheetNames) {
    if (!estNormTxt(hoja).includes("CRITIC")) continue;
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[hoja], { header: 1, raw: true, defval: "" });
    let fechasCol = null, numsCol = null, cEst = -1, cTut = -1, cHor = -1, cArea = -1;
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r] || [];
      const iEst = row.findIndex(v => estNormTxt(v).startsWith("ESTUDIANT"));
      if (iEst >= 0) {   // cabecera de un bloque
        cEst = iEst;
        cTut = row.findIndex(v => estNormTxt(v).startsWith("TUTOR"));
        cHor = row.findIndex(v => estNormTxt(v).startsWith("HORARI"));
        cArea = row.findIndex(v => estNormTxt(v).startsWith("AREA"));
        numsCol = {}; fechasCol = {};
        row.forEach((v, i) => { if (i > cEst && typeof v === "number" && v >= 1 && v <= 31) numsCol[i] = v; });
        for (let k = r - 1; k >= Math.max(0, r - 3); k--) (rows[k] || []).forEach((v, i) => {
          if (v instanceof Date && !isNaN(v) && fechasCol[i] == null) fechasCol[i] = estFecha(v);
        });
        continue;
      }
      if (cEst < 0 || !numsCol) continue;
      const nombre = estTexto(row[cEst]);
      if (!nombre) continue;
      if (cArea >= 0 && row[cArea] && !estNormTxt(row[cArea]).includes("CRITIC")) continue;
      const tut = cTut >= 0 ? estNormTxt(row[cTut]) : "";
      const est = porTutora.get(tut + "|" + estNormTxt(nombre).split(" ")[0]) || porNombre.get(estNormTxt(nombre));
      if (!est) continue;
      const hor = estNormTxt(cHor >= 0 ? row[cHor] : "");
      const noche = /DE (19|20|21|22)/.test(hor) || /NIT|NOCHE/.test(hor);
      // fechas de este bloque
      let fechas = { ...fechasCol };
      if (!Object.keys(fechas).length && Object.keys(numsCol).length) {
        let [y, m] = est.inicio.split("-").map(Number); let prev = 0;
        for (const i of Object.keys(numsCol).map(Number).sort((a, b) => a - b)) {
          const d = numsCol[i];
          if (d < prev) { m++; if (m > 12) { m = 1; y++; } }
          prev = d;
          fechas[i] = `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
        }
      }
      const dias = [];
      for (const [i, f] of Object.entries(fechas)) {
        const v = row[i];
        const n = typeof v === "number" ? v : parseFloat(String(v).replace(",", "."));
        if (!n || isNaN(n)) continue;
        const t = n >= 11 ? (noche ? "NIT_12H" : "DIA_12H") : n >= 9 ? "NIT_10H" : (noche ? "NIT_10H" : est.modalidad === "TARDE" ? "TARDE" : "MANANA");
        dias.push({ f, t });
      }
      if (dias.length) {
        est.dias = dias.sort((a, b) => a.f.localeCompare(b.f));
        asignados++;
      }
    }
  }
  return asignados;
}

/* ---------- Calendario del máster («CALENDARIO UCI 26-27») ----------
   Una hoja por mes. Título en A1 («GENER 2027»). Por cada semana: una fila con
   los números de día (columnas B..H = lunes..domingo) y debajo filas
   «TORN MATÍ / TORN TARDA / TORN NIT» (las filas sin etiqueta siguen al turno de
   arriba) con el nombre del estudiante en la columna del día. El COLOR de la
   celda dice la jornada, según la leyenda de la propia hoja:
   naranja = 12 h · verde = 7,33 h (TM o TT) · amarillo = 10 h (TN). */
const EST_MESES_CAL = { GENER: 0, ENERO: 0, JANUARY: 0, FEBRER: 1, FEBRERO: 1, FEBRUARY: 1, MARC: 2, MARZO: 2, MARCH: 2,
  ABRIL: 3, APRIL: 3, MAIG: 4, MAYO: 4, MAY: 4, JUNY: 5, JUNIO: 5, JUNE: 5, JULIOL: 6, JULIO: 6, JULY: 6,
  AGOST: 7, AGOSTO: 7, AUGUST: 7, SETEMBRE: 8, SEPTIEMBRE: 8, SEPTEMBER: 8, OCTUBRE: 9, OCTOBER: 9,
  NOVEMBRE: 10, NOVIEMBRE: 10, NOVEMBER: 10, DESEMBRE: 11, DICIEMBRE: 11, DECEMBER: 11 };
function estEsCalendario(wb) {
  return wb.SheetNames.some(n => {
    const ws = wb.Sheets[n]; let c = 0;
    for (let r = 1; r <= 40; r++) { const v = ws["A" + r]; if (v && /^TORN\b/.test(estNormTxt(v.v))) c++; }
    return c >= 2;
  });
}
function estMesDeHoja(ws, nombreHoja) {
  for (const txt of [ws.A1 && ws.A1.v, nombreHoja]) {
    const t = estNormTxt(txt).replace(/[^A-Z0-9 ]/g, " ");
    const y = (t.match(/(20\d\d)/) || [])[1];
    const m = t.split(" ").map(w => EST_MESES_CAL[w]).find(x => x != null);
    if (y && m != null) return { anio: +y, mes: m };
  }
  return null;
}
function estLeerCalendarioMaster(wb) {
  const COLS = ["B", "C", "D", "E", "F", "G", "H"];
  const porNombre = new Map();
  const errores = [];
  let desde = "", hasta = "";
  for (const nombreHoja of wb.SheetNames) {
    const ws = wb.Sheets[nombreHoja];
    const mm = estMesDeHoja(ws, nombreHoja);
    if (!mm) continue;
    const rango = XLSX.utils.decode_range(ws["!ref"] || "A1:A1");
    // Leyenda de colores de la hoja (p.ej. K10 «12h» en naranja)
    const leyenda = {};
    for (let r = 0; r <= rango.e.r; r++) for (let c = 8; c <= rango.e.c; c++) {
      const cel = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cel || cel.v == null) continue;
      const rgb = cel.s && cel.s.fgColor && cel.s.fgColor.rgb;
      const t = estNormTxt(cel.v);
      if (rgb && /12 ?H/.test(t)) leyenda[rgb.slice(-6).toUpperCase()] = 12;
      else if (rgb && /7[,.]33|TM|TT/.test(t)) leyenda[rgb.slice(-6).toUpperCase()] = 7;
      else if (rgb && /10 ?H|TN/.test(t)) leyenda[rgb.slice(-6).toUpperCase()] = 10;
    }
    if (!Object.keys(leyenda).length) Object.assign(leyenda, { FF9900: 12, "00FF00": 7, FFFF00: 10 });
    let dias = null, seccion = "";
    for (let r = 1; r <= rango.e.r + 1; r++) {
      // ¿fila de números de día?
      const nums = COLS.map(c => { const v = ws[c + r] && ws[c + r].v; return typeof v === "number" && v >= 1 && v <= 31 ? v : null; });
      if (nums.some(n => n != null)) { dias = nums; seccion = ""; continue; }
      const a = estNormTxt(ws["A" + r] && ws["A" + r].v);
      if (a.startsWith("TORN") || a.startsWith("TURNO")) seccion = /NIT|NOCHE/.test(a) ? "NIT" : /TARD/.test(a) ? "TARDA" : /MATI|MANANA/.test(a) ? "MATI" : "DIA";
      if (!dias || !seccion) continue;
      COLS.forEach((c, i) => {
        const cel = ws[c + r];
        if (!cel || typeof cel.v !== "string" || !cel.v.trim() || dias[i] == null) return;
        const nombre = cel.v.replace(/\s+/g, " ").trim();
        const rgb = ((cel.s && cel.s.fgColor && cel.s.fgColor.rgb) || "").slice(-6).toUpperCase();
        const h = leyenda[rgb];
        const t = h === 12 ? (seccion === "NIT" ? "NIT_12H" : "DIA_12H")
          : h === 10 ? "NIT_10H"
          : h === 7 ? (seccion === "TARDA" ? "TARDE" : seccion === "NIT" ? "NIT_10H" : "MANANA")
          : seccion === "NIT" ? "NIT_10H" : seccion === "TARDA" ? "TARDE" : seccion === "MATI" ? "MANANA" : "DIA_12H";
        const f = `${mm.anio}-${String(mm.mes + 1).padStart(2, "0")}-${String(dias[i]).padStart(2, "0")}`;
        const k = estNormTxt(nombre);
        if (!porNombre.has(k)) porNombre.set(k, { nombre, dias: [] });
        const est = porNombre.get(k);
        if (!est.dias.some(d => d.f === f && d.t === t)) est.dias.push({ f, t });
      });
    }
    const ini = `${mm.anio}-${String(mm.mes + 1).padStart(2, "0")}-01`;
    const fin = `${mm.anio}-${String(mm.mes + 1).padStart(2, "0")}-${String(new Date(mm.anio, mm.mes + 1, 0).getDate()).padStart(2, "0")}`;
    if (!desde || ini < desde) desde = ini;
    if (!hasta || fin > hasta) hasta = fin;
  }
  const filas = [...porNombre.values()].map(x => {
    x.dias.sort((a, b) => a.f.localeCompare(b.f));
    const cuenta = {}; for (const d of x.dias) cuenta[d.t] = (cuenta[d.t] || 0) + 1;
    const partes = x.nombre.split(" ");
    return {
      nombre: partes[0], apellido1: partes.slice(1).join(" "), apellido2: "",
      tipo: "MASTER", curso: "Máster", escuela: "MÁSTER UCI", tutora: "",
      modalidad: Object.keys(cuenta).sort((a, b) => cuenta[b] - cuenta[a])[0] || "",
      inicio: x.dias[0].f, fin: x.dias[x.dias.length - 1].f, contacto: "", notas: "",
      dias: x.dias, origen: "CAL_MASTER",
    };
  });
  if (!filas.length) errores.push("No se ha encontrado ningún estudiante en el calendario.");
  return { filas, errores, desde, hasta };
}

function estExportarExcel(lista) {
  const cab = ["INICIO", "FIN", "ENFERMERA TUTORA", "TURNO", "NOMBRE", "APELLIDO 1", "APELLIDO 2", "ESCUELA", "CURSO", "CONTACTO", "NOTAS"];
  const filas = lista.slice().sort((a, b) => (a.inicio + a.fin).localeCompare(b.inicio + b.fin) || estNombre(a).localeCompare(estNombre(b)))
    .map(e => [e.inicio.split("-").reverse().join("/"), e.fin.split("-").reverse().join("/"), e.tutora, e.modalidad ? estMod(e.modalidad).label : "",
      e.nombre, e.apellido1, e.apellido2, e.escuela, e.curso, e.contacto, e.notas]);
  const ws = XLSX.utils.aoa_to_sheet([cab, ...filas]);
  ws["!cols"] = [10, 10, 34, 11, 16, 14, 14, 9, 8, 32, 24].map(w => ({ wch: w }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Estudiantes");
  XLSX.writeFile(wb, "Estudiantes_practicas.xlsx");
}

const EST_VACIO = { nombre: "", apellido1: "", apellido2: "", curso: "4º", escuela: "", tutora: "", modalidad: "DIA_12H", inicio: "", fin: "", contacto: "", notas: "" };

function EstCalendarioMaster({ lista, anio, mes, onEditar }) {
  const est = lista.filter(e => estDiasMes(e, anio, mes).length > 0);
  if (!est.length) return null;
  const nDias = new Date(anio, mes + 1, 0).getDate();
  const offset = (new Date(anio, mes, 1).getDay() + 6) % 7;   // lunes = 0
  const porDia = {};
  for (const e of est) for (const d of estDiasMes(e, anio, mes)) {
    (porDia[d.f] = porDia[d.f] || []).push({ e, t: d.t });
  }
  const orden = { MANANA: 0, DIA_12H: 1, TARDE: 2, NIT_12H: 3, NIT_10H: 4 };
  const celdas = [];
  for (let i = 0; i < offset; i++) celdas.push(null);
  for (let d = 1; d <= nDias; d++) celdas.push(d);
  const resumen = est.map(e => {
    const ds = estDiasMes(e, anio, mes); const c = {};
    for (const d of ds) c[d.t] = (c[d.t] || 0) + 1;
    return { e, n: ds.length, c };
  }).sort((a, b) => b.n - a.n || estNombre(a.e).localeCompare(estNombre(b.e)));
  return (
    <div className="est-card">
      <div className="est-card-head">
        <h2>Calendario de turnos · {EST_MESES_LARGO[mes]}</h2>
        <div className="est-leyenda">{EST_MODALIDADES.filter(m => m.id && est.some(e => estDiasMes(e, anio, mes).some(d => d.t === m.id))).map(m => <span key={m.id}><i style={{ background: m.color }} />{m.label}</span>)}</div>
      </div>
      <div className="est-scroll">
        <div className="est-cal">
          {["Lun", "Mar", "Mié", "Jue", "Vie", "Sáb", "Dom"].map(d => <div key={d} className="est-cal-dow">{d}</div>)}
          {celdas.map((d, i) => {
            if (d == null) return <div key={"v" + i} className="est-cal-dia is-vacio" />;
            const f = `${anio}-${String(mes + 1).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
            const items = (porDia[f] || []).slice().sort((a, b) => (orden[a.t] ?? 9) - (orden[b.t] ?? 9));
            return (
              <div key={f} className={"est-cal-dia" + (i % 7 >= 5 ? " is-finde" : "")}>
                <div className="est-cal-num">{d}{items.length > 0 && <span>{items.length}</span>}</div>
                {items.map((x, j) => { const m = estMod(x.t); return (
                  <div key={j} className="est-cal-chip" style={{ background: m.color, color: x.t === "NIT_12H" ? "#fff" : "#1d2530" }} title={`${estNombre(x.e)} · ${m.label}`}>{estNombre(x.e)}</div>
                ); })}
              </div>
            );
          })}
        </div>
      </div>
      <table className="est-tabla est-resumen-master">
        <thead><tr><th>Estudiante</th><th className="est-num">Turnos en {EST_MESES_LARGO[mes]}</th><th>Reparto</th><th></th></tr></thead>
        <tbody>
          {resumen.map(({ e, n, c }) => (
            <tr key={e.id}>
              <td><b>{estNombre(e)}</b> <span className="est-tag" style={{ background: EST_COLOR[e.tipo] }}>{e.tipo === "MASTER" ? "Máster" : "Grado " + (e.curso || "")}</span></td>
              <td className="est-num"><b>{n}</b></td>
              <td>{Object.keys(c).sort((a, b) => (orden[a] ?? 9) - (orden[b] ?? 9)).map(t => { const m = estMod(t); return (
                <span key={t} className="est-turno" style={{ background: m.color, color: t === "NIT_12H" ? "#fff" : "#1d2530", marginRight: 4 }}>{m.label} · {c[t]}</span>
              ); })}</td>
              <td className="est-nowrap"><button className="est-link" onClick={() => onEditar(e)}>Editar</button></td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function EstudiantesPanel() {
  const hoy = new Date();
  const [lista, setLista] = useState(null);
  const [error, setError] = useState("");
  const [aviso, setAviso] = useState("");
  const [anio, setAnio] = useState(hoy.getFullYear());
  const [mesSel, setMesSel] = useState(hoy.getMonth());
  const [verTodos, setVerTodos] = useState(false);
  const [busca, setBusca] = useState("");
  const [form, setForm] = useState(null);
  const [guardando, setGuardando] = useState(false);
  const fileRef = useRef(null);

  const recargar = useCallback(async () => {
    try { setLista(await cargarEstudiantes()); setError(""); }
    catch (e) { setError(e.message); setLista([]); }
  }, []);
  useEffect(() => { recargar(); }, [recargar]);

  const porMes = useMemo(() => Array.from({ length: 12 }, (_, m) => {
    const est = (lista || []).filter(e => estEnMes(e, anio, m));
    return { m, est, total: est.length, grado: est.filter(e => e.tipo === "GRADO").length, master: est.filter(e => e.tipo === "MASTER").length };
  }), [lista, anio]);
  const maxMes = Math.max(1, ...porMes.map(x => x.total));
  const mesActual = porMes[mesSel];

  // Desglose del mes: curso x turno
  const desglose = useMemo(() => {
    const g = new Map();
    for (const e of mesActual.est) {
      const k = e.curso || "Sin curso";
      if (!g.has(k)) g.set(k, { curso: k, tipo: e.tipo, total: 0, mods: {} });
      const md = estModMes(e, anio, mesSel);
      const x = g.get(k); x.total++; x.mods[md] = (x.mods[md] || 0) + 1;
    }
    const orden = ["4º", "3º", "2º", "1º", "Máster", "Sin curso"];
    return [...g.values()].sort((a, b) => orden.indexOf(a.curso) - orden.indexOf(b.curso));
  }, [mesActual, anio, mesSel]);
  const modsUsadas = EST_MODALIDADES.filter(m => mesActual.est.some(e => estModMes(e, anio, mesSel) === m.id));

  // Periodos (mismas fechas de inicio y fin) visibles: los del mes elegido o todos
  const periodos = useMemo(() => {
    const q = estNormTxt(busca);
    const base = (lista || []).filter(e => e.origen !== "CAL_MASTER")
      .filter(e => verTodos || estEnMes(e, anio, mesSel))
      .filter(e => !q || estNormTxt([estNombre(e), e.tutora, e.escuela, e.contacto].join(" ")).includes(q));
    const g = new Map();
    for (const e of base) {
      const k = e.inicio + "|" + e.fin;
      if (!g.has(k)) g.set(k, { inicio: e.inicio, fin: e.fin, est: [] });
      g.get(k).est.push(e);
    }
    return [...g.values()].sort((a, b) => (verTodos ? b.inicio.localeCompare(a.inicio) : a.inicio.localeCompare(b.inicio)))
      .map(p => ({ ...p, est: p.est.sort((a, b) => (a.modalidad || "").localeCompare(b.modalidad || "") || a.tutora.localeCompare(b.tutora)),
        escuelas: [...new Set(p.est.map(e => e.escuela).filter(Boolean))], cursos: [...new Set(p.est.map(e => e.curso).filter(Boolean))] }));
  }, [lista, anio, mesSel, verTodos, busca]);

  async function guardar() {
    if (!form.nombre.trim()) { setAviso("Escribe el nombre del estudiante."); return; }
    if (!form.inicio || !form.fin) { setAviso("Indica la fecha de inicio y de fin."); return; }
    if (form.fin < form.inicio) { setAviso("La fecha de fin es anterior a la de inicio."); return; }
    setGuardando(true);
    try { await guardarEstudiante({ ...form, tipo: estTipoDeCurso(form.curso) }); setForm(null); setAviso(""); await recargar(); }
    catch (e) { setAviso(e.message); }
    finally { setGuardando(false); }
  }
  async function borrar(e) {
    if (!window.confirm(`¿Borrar a ${estNombre(e)}?`)) return;
    try { await borrarEstudiante(e.id); setForm(null); await recargar(); } catch (err) { setError(err.message); }
  }
  async function importar(ev) {
    const f = ev.target.files && ev.target.files[0];
    ev.target.value = "";
    if (!f) return;
    try {
      const buf = await f.arrayBuffer();
      const wbCal = XLSX.read(buf, { cellStyles: true });
      if (estEsCalendario(wbCal)) {
        // Calendario del máster: sustituye lo importado antes de ese calendario
        const { filas, errores, desde, hasta } = estLeerCalendarioMaster(wbCal);
        if (!filas.length) { setAviso(errores.join(" ")); return; }
        await borrarEstudiantesCalendario(desde, hasta);
        await importarEstudiantes(filas);
        await recargar();
        const turnos = filas.reduce((n, e) => n + e.dias.length, 0);
        setAviso(`Calendario del máster «${f.name}»: ${filas.length} estudiantes y ${turnos} turnos entre ${estFechaCorta(desde)} y ${estFechaCorta(hasta)}. Si ya se había importado, se ha actualizado.`);
        return;
      }
      const wb = XLSX.read(buf, { cellDates: true });
      const { filas, errores, excluidos, conHorario } = estLeerExcel(wb);
      if (!filas.length && !errores.length) { setAviso(`«${f.name}»: no se ha encontrado ningún estudiante de Crítics.`); return; }
      // No duplicar: si el estudiante ya está (mismo correo, o mismo nombre y
      // primer apellido con fechas que se solapan) se actualiza con lo nuevo.
      const correo = (e) => ((e.contacto || "").match(/[^\s·]+@[^\s·]+/) || [""])[0].toLowerCase();
      const mismo = (a, b) => (correo(a) && correo(a) === correo(b)) ||
        (estNormTxt(a.nombre) === estNormTxt(b.nombre) && estNormTxt(a.apellido1) === estNormTxt(b.apellido1) && a.inicio <= b.fin && b.inicio <= a.fin);
      const existentes = (lista || []).filter(e => e.origen !== "CAL_MASTER");
      const nuevas = [], actualizar = [];
      for (const e of filas) {
        const prev = existentes.find(x => mismo(x, e));
        if (!prev) { nuevas.push(e); continue; }
        const merged = { ...prev };
        for (const k of ["nombre", "apellido1", "apellido2", "curso", "tipo", "escuela", "tutora", "modalidad", "inicio", "fin", "contacto", "notas"]) if (e[k]) merged[k] = e[k];
        if (e.dias && e.dias.length) merged.dias = e.dias;
        actualizar.push(merged);
      }
      if (nuevas.length) await importarEstudiantes(nuevas);
      for (const e of actualizar) await guardarEstudiante(e);
      await recargar();
      let msg = `«${f.name}»: ${nuevas.length} estudiantes nuevos`;
      if (actualizar.length) msg += `, ${actualizar.length} ya estaban y se han actualizado`;
      if (excluidos) msg += `, ${excluidos} de otras áreas ignorados (solo se cogen los de Crítics)`;
      if (conHorario) msg += `, ${conHorario} con horario día a día`;
      msg += ".";
      if (errores.length) msg += ` ${errores.length} filas sin importar — ${errores.slice(0, 4).join(" ")}${errores.length > 4 ? " …" : ""}`;
      setAviso(msg);
    } catch (e) { setAviso(e.message); }
  }

  if (lista === null) return <div className="est-wrap"><p className="est-muted">Cargando estudiantes…</p></div>;
  const titMes = EST_MESES_LARGO[mesSel].replace(/^./, c => c.toUpperCase()) + " " + anio;

  return (
    <div className="est-wrap">
      {error && <div className="est-alerta">{error}</div>}

      <div className="est-top">
        <div className="est-anio">
          <button onClick={() => setAnio(anio - 1)} aria-label="Año anterior">‹</button>
          <span>{anio}</span>
          <button onClick={() => setAnio(anio + 1)} aria-label="Año siguiente">›</button>
        </div>
        <div className="est-acciones">
          <button className="est-btn" onClick={() => fileRef.current && fileRef.current.click()}>Importar Excel</button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls,.xlsm,.csv" style={{ display: "none" }} onChange={importar} />
          <button className="est-btn" onClick={() => estExportarExcel(lista)} disabled={!lista.length}>Exportar Excel</button>
          <button className="est-btn est-btn-primary" onClick={() => { setForm({ ...EST_VACIO }); setAviso(""); }}>+ Añadir estudiante</button>
        </div>
      </div>
      {aviso && !form && <div className="est-aviso">{aviso}</div>}

      {lista.length === 0 && (
        <div className="est-card est-vacio">
          <h2>Aún no hay estudiantes</h2>
          <p>Importa el Excel de prácticas (el de periodos, enfermera, torn, nombre, escuela y curso) o añade estudiantes uno a uno.</p>
        </div>
      )}

      {/* Estudiantes por mes */}
      <div className="est-card">
        <div className="est-card-head">
          <h2>Estudiantes por mes · {anio}</h2>
          <div className="est-leyenda"><span><i style={{ background: EST_COLOR.GRADO }} />Grado</span><span><i style={{ background: EST_COLOR.MASTER }} />Máster</span></div>
        </div>
        <div className="est-barras">
          {porMes.map(x => (
            <button key={x.m} className={"est-barra" + (x.m === mesSel ? " is-sel" : "")} onClick={() => setMesSel(x.m)}
              title={`${EST_MESES_LARGO[x.m]}: ${x.total} (${x.grado} de grado, ${x.master} de máster)`}>
              <span className="est-barra-n">{x.total || ""}</span>
              <span className="est-barra-col">
                {x.master > 0 && <span style={{ height: `${(x.master / maxMes) * 100}%`, background: EST_COLOR.MASTER }} />}
                {x.grado > 0 && <span style={{ height: `${(x.grado / maxMes) * 100}%`, background: EST_COLOR.GRADO }} />}
              </span>
              <span className="est-barra-m">{EST_MESES[x.m]}</span>
            </button>
          ))}
        </div>
      </div>

      {/* Resumen del mes */}
      <div className="est-card">
        <div className="est-card-head"><h2>{titMes}</h2><span className="est-muted">Pulsa otro mes en el gráfico para cambiar</span></div>
        <div className="est-kpis">
          <div className="est-kpi"><b>{mesActual.total}</b><span>estudiantes</span></div>
          <div className="est-kpi"><b style={{ color: EST_COLOR.GRADO }}>{mesActual.grado}</b><span>de grado</span></div>
          <div className="est-kpi"><b style={{ color: EST_COLOR.MASTER }}>{mesActual.master}</b><span>de máster</span></div>
        </div>
        {mesActual.total > 0 && (<>
          <div className="est-mods">
            {modsUsadas.map(m => { const n = mesActual.est.filter(e => estModMes(e, anio, mesSel) === m.id).length; return (
              <div key={m.id} className="est-mod" style={{ flex: n, background: m.color }} title={`${m.label}: ${n}`}><span>{m.label} · {n}</span></div>
            ); })}
          </div>
          <div className="est-scroll">
            <table className="est-tabla">
              <thead><tr><th>Curso</th><th className="est-num">Total</th>{modsUsadas.map(m => <th key={m.id} className="est-num"><i style={{ background: m.color }} />{m.label}</th>)}</tr></thead>
              <tbody>
                {desglose.map(g => (
                  <tr key={g.curso}>
                    <td><span className="est-tag" style={{ background: EST_COLOR[g.tipo] }}>{g.tipo === "MASTER" ? "Máster" : "Grado"}</span> {g.tipo === "MASTER" ? "" : g.curso === "Sin curso" ? "sin curso" : `${g.curso} curso`}</td>
                    <td className="est-num"><b>{g.total}</b></td>
                    {modsUsadas.map(m => <td key={m.id} className="est-num">{g.mods[m.id] || ""}</td>)}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </>)}
      </div>

      {/* Calendario del máster (estudiantes que vienen por días sueltos) */}
      <EstCalendarioMaster lista={lista} anio={anio} mes={mesSel} onEditar={e => { setForm({ ...EST_VACIO, ...e }); setAviso(""); }} />

      {/* Periodos con parejas tutora - estudiante */}
      <div className="est-card">
        <div className="est-card-head">
          <h2>{verTodos ? "Todos los periodos de prácticas" : `Periodos de prácticas en ${EST_MESES_LARGO[mesSel]}`}</h2>
          <div className="est-filtros">
            <input className="est-busca" placeholder="Buscar estudiante, enfermera o escuela…" value={busca} onChange={ev => setBusca(ev.target.value)} />
            <label className="est-check"><input type="checkbox" checked={verTodos} onChange={ev => setVerTodos(ev.target.checked)} />Ver todos</label>
          </div>
        </div>
        {periodos.length === 0 ? <p className="est-muted">No hay estudiantes {verTodos ? "que coincidan con la búsqueda" : "en este mes"}.</p> : (
          <div className="est-periodos">
            {periodos.map(p => (
              <section key={p.inicio + p.fin} className="est-periodo">
                <header>
                  <div>
                    <div className="est-periodo-fechas">{estFechaCorta(p.inicio)} <span>→</span> {estFechaCorta(p.fin)}</div>
                    <div className="est-periodo-meta">{p.est.length} estudiante{p.est.length === 1 ? "" : "s"}{p.escuelas.length ? " · " + p.escuelas.join(", ") : ""}{p.cursos.length ? " · " + p.cursos.join(", ") : ""}</div>
                  </div>
                </header>
                <div className="est-scroll">
                  <table className="est-tabla est-lista">
                    <colgroup><col style={{ width: "23%" }} /><col style={{ width: "11%" }} /><col style={{ width: "25%" }} /><col style={{ width: "8%" }} /><col style={{ width: "7%" }} /><col /><col style={{ width: "64px" }} /></colgroup>
                    <thead><tr><th>Enfermera tutora</th><th>Turno</th><th>Estudiante</th><th>Escuela</th><th>Curso</th><th>Contacto</th><th></th></tr></thead>
                    <tbody>
                      {p.est.map(e => { const m = estMod(e.modalidad); return (
                        <tr key={e.id}>
                          <td className="est-tutora est-c-tut">{e.tutora || <span className="est-falta">sin tutora</span>}</td>
                          <td className="est-c-tur"><span className="est-turno" style={{ background: m.color, color: m.id === "NIT_12H" ? "#fff" : "#1d2530" }}>{m.label}</span></td>
                          <td className="est-c-est"><b>{estNombre(e)}</b></td>
                          <td className="est-c-esc">{e.escuela}</td>
                          <td className="est-c-cur">{e.curso}</td>
                          <td className="est-contacto est-c-con">{e.contacto}</td>
                          <td className="est-nowrap est-c-edi"><button className="est-link" onClick={() => { setForm({ ...EST_VACIO, ...e }); setAviso(""); }}>Editar</button></td>
                        </tr>
                      ); })}
                    </tbody>
                  </table>
                </div>
              </section>
            ))}
          </div>
        )}
      </div>

      {form && (
        <div className="est-modal" onClick={() => !guardando && setForm(null)}>
          <div className="est-modal-card" onClick={ev => ev.stopPropagation()}>
            <h2>{form.id ? "Editar estudiante" : "Añadir estudiante"}</h2>
            <div className="est-form">
              <label>Nombre<input value={form.nombre} onChange={ev => setForm({ ...form, nombre: ev.target.value })} /></label>
              <label>Primer apellido<input value={form.apellido1} onChange={ev => setForm({ ...form, apellido1: ev.target.value })} /></label>
              <label>Segundo apellido<input value={form.apellido2} onChange={ev => setForm({ ...form, apellido2: ev.target.value })} /></label>
              <label>Curso<select value={form.curso} onChange={ev => setForm({ ...form, curso: ev.target.value })}>{EST_CURSOS.map(c => <option key={c}>{c}</option>)}</select></label>
              <label>Escuela<input value={form.escuela} onChange={ev => setForm({ ...form, escuela: ev.target.value.toUpperCase() })} placeholder="UDPT, UIC, EUIT…" /></label>
              <label>Turno<select value={form.modalidad || ""} onChange={ev => setForm({ ...form, modalidad: ev.target.value })}>{EST_MODALIDADES.map(m => <option key={m.id} value={m.id}>{m.label}</option>)}</select></label>
              <label className="est-form-full">Enfermera tutora<input value={form.tutora} onChange={ev => setForm({ ...form, tutora: ev.target.value })} placeholder="Si son dos, sepáralas con /" /></label>
              <label>Inicio<input type="date" value={form.inicio} onChange={ev => setForm({ ...form, inicio: ev.target.value })} /></label>
              <label>Fin<input type="date" value={form.fin} onChange={ev => setForm({ ...form, fin: ev.target.value })} /></label>
              <label className="est-form-full">Contacto (correo o teléfono)<input value={form.contacto} onChange={ev => setForm({ ...form, contacto: ev.target.value })} /></label>
              <label className="est-form-full">Notas<input value={form.notas} onChange={ev => setForm({ ...form, notas: ev.target.value })} /></label>
            </div>
            {aviso && <div className="est-aviso">{aviso}</div>}
            <div className="est-modal-btns">
              {form.id && <button className="est-btn est-btn-del" disabled={guardando} onClick={() => borrar(form)}>Borrar</button>}
              <span style={{ flex: 1 }} />
              <button className="est-btn" disabled={guardando} onClick={() => setForm(null)}>Cancelar</button>
              <button className="est-btn est-btn-primary" disabled={guardando} onClick={guardar}>{guardando ? "Guardando…" : "Guardar"}</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function Header({ colectivo, onCambiarColectivo, email, apodo, abierto, onToggle, onCerrarMenu, onLogout, onCambiarApodo, onCambiarPassword, onSesionActualizada }) {
  const nombreMostrado = apodo || (email ? email.split("@")[0] : "");
  return (
    <header className="header">
      <div className="brand">
        <div className="brand-mark" aria-hidden="true">
          <svg viewBox="0 0 32 32" width="26" height="26">
            <rect x="3" y="3" width="26" height="26" rx="6" fill="none" stroke="currentColor" strokeWidth="2" />
            <path d="M16 9v14M9 16h14" stroke="currentColor" strokeWidth="2.4" strokeLinecap="round" />
          </svg>
        </div>
        <div>
          <div className="brand-title">Cuadrante UCI</div>
          <div className="brand-sub">
            {colectivo === "TCAE" ? "Rotación de TCAE · planificación semanal"
              : colectivo === "ENF" ? "Rotación de enfermería · planificación semanal"
              : colectivo === "EST" ? "Asignación de estudiantes de grado y máster"
              : "Planificación semanal de la unidad"}
          </div>
        </div>
        {colectivo && (
          <button className="colectivo-chip" onClick={onCambiarColectivo} title="Cambiar de grupo profesional">
            <span className={"colectivo-dot " + (colectivo === "TCAE" ? "is-tcae" : colectivo === "EST" ? "is-est" : "is-enf")} />
            {colectivo === "TCAE" ? "TCAE" : colectivo === "EST" ? "Estudiantes" : "Enfermería"}
            <span className="colectivo-cambiar">Cambiar</span>
          </button>
        )}
      </div>
      {email && (
        <div className="perfil">
          <button className="perfil-btn" onClick={onToggle} aria-label="Tu perfil">
            <span className="avatar">{iniciales(apodo || email)}</span>
            <span className="perfil-nombre">{nombreMostrado}</span>
            <svg viewBox="0 0 12 12" width="11" height="11" className="perfil-caret"><path d="M2 4l4 4 4-4" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" strokeLinejoin="round"/></svg>
          </button>
          {abierto && (
            <ProfileMenu
              email={email} apodo={apodo}
              onCerrar={onCerrarMenu} onLogout={onLogout}
              onCambiarApodo={onCambiarApodo} onCambiarPassword={onCambiarPassword}
              onSesionActualizada={onSesionActualizada}
            />
          )}
        </div>
      )}
    </header>
  );
}

/* ---------- Menú/panel de perfil ---------- */
function ProfileMenu({ email, apodo, onCerrar, onLogout, onCambiarApodo, onCambiarPassword, onSesionActualizada }) {
  const [vista, setVista] = useState("menu"); // menu | apodo | password
  const ref = useRef(null);

  // Cerrar al hacer clic fuera
  useEffect(() => {
    function fuera(e) { if (ref.current && !ref.current.contains(e.target)) onCerrar(); }
    document.addEventListener("mousedown", fuera);
    return () => document.removeEventListener("mousedown", fuera);
  }, [onCerrar]);

  return (
    <>
      <div className="perfil-backdrop" />
      <div className="perfil-menu" ref={ref}>
        {vista === "menu" && (
          <>
            <div className="perfil-head">
              <span className="avatar grande">{iniciales(apodo || email)}</span>
              <div className="perfil-info">
                <strong>{apodo || "Sin apodo"}</strong>
                <span title={email}>{email}</span>
              </div>
            </div>
            <button className="perfil-item" onClick={() => setVista("apodo")}>
              <IcoUser /> Cambiar apodo
            </button>
            <button className="perfil-item" onClick={() => setVista("password")}>
              <IcoLock /> Cambiar contraseña
            </button>
            <div className="perfil-sep" />
            <button className="perfil-item salir" onClick={onLogout}>
              <IcoSalir /> Cerrar sesión
            </button>
          </>
        )}
        {vista === "apodo" && (
          <FormApodo apodoActual={apodo} onVolver={() => setVista("menu")}
            onCambiarApodo={onCambiarApodo} onSesionActualizada={onSesionActualizada} />
        )}
        {vista === "password" && (
          <FormPassword onVolver={() => setVista("menu")} onCambiarPassword={onCambiarPassword} />
        )}
      </div>
    </>
  );
}

function FormApodo({ apodoActual, onVolver, onCambiarApodo, onSesionActualizada }) {
  const [valor, setValor] = useState(apodoActual || "");
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState("");
  async function guardar() {
    setBusy(true); setMsg("");
    try {
      await onCambiarApodo(valor);
      await onSesionActualizada();
      setMsg("ok");
      setTimeout(onVolver, 700);
    } catch (err) { setMsg(err.message || "No se pudo guardar."); setBusy(false); }
  }
  return (
    <div className="perfil-form">
      <button className="perfil-volver" onClick={onVolver}>← Volver</button>
      <h4>Cambiar apodo</h4>
      <p className="perfil-hint">El apodo se muestra arriba en lugar de tu email.</p>
      <input className="perfil-input" value={valor} maxLength={24}
        placeholder="Tu apodo" autoFocus
        onChange={e => setValor(e.target.value)}
        onKeyDown={e => e.key === "Enter" && guardar()} />
      {msg && msg !== "ok" && <div className="perfil-error">{msg}</div>}
      {msg === "ok" && <div className="perfil-ok">Apodo guardado ✓</div>}
      <button className="perfil-guardar" disabled={busy} onClick={guardar}>
        {busy ? "Guardando…" : "Guardar apodo"}
      </button>
    </div>
  );
}

function FormPassword({ onVolver, onCambiarPassword }) {
  const [p1, setP1] = useState("");
  const [p2, setP2] = useState("");
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState("");
  async function guardar() {
    if (p1.length < 6) { setMsg("La contraseña debe tener al menos 6 caracteres."); return; }
    if (p1 !== p2) { setMsg("Las dos contraseñas no coinciden."); return; }
    setBusy(true); setMsg("");
    try {
      await onCambiarPassword(p1);
      setMsg("ok");
      setTimeout(onVolver, 900);
    } catch (err) { setMsg(err.message || "No se pudo cambiar."); setBusy(false); }
  }
  return (
    <div className="perfil-form">
      <button className="perfil-volver" onClick={onVolver}>← Volver</button>
      <h4>Cambiar contraseña</h4>
      <input className="perfil-input" type="password" value={p1} placeholder="Nueva contraseña" autoFocus
        autoComplete="new-password" onChange={e => setP1(e.target.value)} />
      <input className="perfil-input" type="password" value={p2} placeholder="Repite la contraseña"
        autoComplete="new-password" onChange={e => setP2(e.target.value)}
        onKeyDown={e => e.key === "Enter" && guardar()} />
      {msg && msg !== "ok" && <div className="perfil-error">{msg}</div>}
      {msg === "ok" && <div className="perfil-ok">Contraseña actualizada ✓</div>}
      <button className="perfil-guardar" disabled={busy} onClick={guardar}>
        {busy ? "Guardando…" : "Cambiar contraseña"}
      </button>
    </div>
  );
}

function IcoUser() { return <svg viewBox="0 0 16 16" width="15" height="15"><circle cx="8" cy="5" r="2.6" fill="none" stroke="currentColor" strokeWidth="1.3"/><path d="M3 13c0-2.5 2.2-4 5-4s5 1.5 5 4" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round"/></svg>; }
function IcoLock() { return <svg viewBox="0 0 16 16" width="15" height="15"><rect x="3.5" y="7" width="9" height="6.5" rx="1.2" fill="none" stroke="currentColor" strokeWidth="1.3"/><path d="M5.5 7V5a2.5 2.5 0 015 0v2" fill="none" stroke="currentColor" strokeWidth="1.3"/></svg>; }
function IcoSalir() { return <svg viewBox="0 0 16 16" width="15" height="15"><path d="M6 3H3.5v10H6M10 5l3 3-3 3M13 8H6.5" fill="none" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round"/></svg>; }

/* ---------- Pantalla de inicio de sesión ---------- */
function LoginScreen({ onLogin, onRecuperar }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [busy, setBusy] = useState(false);
  const [modo, setModo] = useState("login"); // login | recuperar
  const [aviso, setAviso] = useState("");

  async function submit() {
    if (!email.trim() || !password) { setError("Introduce email y contraseña."); return; }
    setError(""); setBusy(true);
    try {
      await onLogin(email.trim(), password);
    } catch (err) {
      const msg = /invalid login credentials/i.test(err.message || "")
        ? "Email o contraseña incorrectos."
        : (err.message || "No se pudo iniciar sesión.");
      setError(msg);
      setBusy(false);
    }
  }

  async function enviarRecuperacion() {
    if (!email.trim()) { setError("Escribe tu email para enviarte el enlace."); return; }
    setError(""); setBusy(true);
    try {
      await onRecuperar(email.trim());
      setAviso("Si ese email tiene cuenta, te hemos enviado un enlace para restablecer la contraseña. Revisa tu correo (y la carpeta de spam).");
      setBusy(false);
    } catch (err) {
      setError(err.message || "No se pudo enviar el correo.");
      setBusy(false);
    }
  }

  return (
    <div className="login-wrap">
      <div className="login-card">
        <div className="login-mark" aria-hidden="true">
          <svg viewBox="0 0 32 32" width="30" height="30">
            <rect x="3" y="3" width="26" height="26" rx="7" fill="none" stroke="currentColor" strokeWidth="2.2" />
            <path d="M16 9v14M9 16h14" stroke="currentColor" strokeWidth="2.6" strokeLinecap="round" />
          </svg>
        </div>
        <h1>Cuadrante UCI</h1>
        <p className="login-sub">
          {modo === "login" ? "Acceso para coordinadoras de la unidad" : "Recuperar contraseña"}
        </p>

        {aviso ? (
          <>
            <div className="login-ok">{aviso}</div>
            <button className="login-btn" onClick={() => { setModo("login"); setAviso(""); }}>
              Volver al inicio
            </button>
          </>
        ) : (
          <>
            <label className="login-field">
              <span>Email</span>
              <input type="email" autoComplete="username" value={email}
                onChange={e => setEmail(e.target.value)}
                onKeyDown={e => e.key === "Enter" && (modo === "login" ? submit() : enviarRecuperacion())} />
            </label>
            {modo === "login" && (
              <label className="login-field">
                <span>Contraseña</span>
                <input type="password" autoComplete="current-password" value={password}
                  onChange={e => setPassword(e.target.value)}
                  onKeyDown={e => e.key === "Enter" && submit()} />
              </label>
            )}

            {error && <div className="login-error">{error}</div>}

            {modo === "login" ? (
              <>
                <button className="login-btn" disabled={busy} onClick={submit}>
                  {busy ? "Entrando…" : "Entrar"}
                </button>
                <button className="login-link" onClick={() => { setModo("recuperar"); setError(""); }}>
                  ¿Has olvidado la contraseña?
                </button>
              </>
            ) : (
              <>
                <button className="login-btn" disabled={busy} onClick={enviarRecuperacion}>
                  {busy ? "Enviando…" : "Enviarme el enlace"}
                </button>
                <button className="login-link" onClick={() => { setModo("login"); setError(""); }}>
                  ← Volver
                </button>
              </>
            )}
          </>
        )}

        {modo === "login" && !aviso && (
          <p className="login-foot">
            El acceso está restringido a las cuentas autorizadas. Si no puedes entrar,
            contacta con la persona responsable de la herramienta.
          </p>
        )}
      </div>
    </div>
  );
}

/* ---------- Pantalla de nueva contraseña (tras enlace de recuperación) ---------- */
function NuevaPasswordScreen({ onGuardar, onListo }) {
  const [p1, setP1] = useState("");
  const [p2, setP2] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [hecho, setHecho] = useState(false);

  async function guardar() {
    if (p1.length < 6) { setError("La contraseña debe tener al menos 6 caracteres."); return; }
    if (p1 !== p2) { setError("Las dos contraseñas no coinciden."); return; }
    setError(""); setBusy(true);
    try {
      await onGuardar(p1);
      setHecho(true);
    } catch (err) {
      setError(err.message || "No se pudo guardar la contraseña.");
      setBusy(false);
    }
  }

  return (
    <div className="login-wrap">
      <div className="login-card">
        <div className="login-mark" aria-hidden="true">
          <svg viewBox="0 0 32 32" width="30" height="30">
            <rect x="3" y="3" width="26" height="26" rx="7" fill="none" stroke="currentColor" strokeWidth="2.2" />
            <path d="M16 9v14M9 16h14" stroke="currentColor" strokeWidth="2.6" strokeLinecap="round" />
          </svg>
        </div>
        <h1>Nueva contraseña</h1>
        {hecho ? (
          <>
            <div className="login-ok">Contraseña actualizada correctamente.</div>
            <button className="login-btn" onClick={onListo}>Entrar</button>
          </>
        ) : (
          <>
            <p className="login-sub">Introduce tu nueva contraseña.</p>
            <label className="login-field">
              <span>Nueva contraseña</span>
              <input type="password" value={p1} autoComplete="new-password" autoFocus
                onChange={e => setP1(e.target.value)} />
            </label>
            <label className="login-field">
              <span>Repite la contraseña</span>
              <input type="password" value={p2} autoComplete="new-password"
                onChange={e => setP2(e.target.value)}
                onKeyDown={e => e.key === "Enter" && guardar()} />
            </label>
            {error && <div className="login-error">{error}</div>}
            <button className="login-btn" disabled={busy} onClick={guardar}>
              {busy ? "Guardando…" : "Guardar contraseña"}
            </button>
          </>
        )}
      </div>
    </div>
  );
}

/* ---------- Panel lateral ---------- */
function Sidebar(props) {
  const { planillaInfo, staffSource, profs, fechaLunes, setFechaLunes, numSemanas, setNumSemanas,
    onPlanillaClick, onStaffClick, generar, puedeGenerar, busy, historicoLen,
    resultado, guardarEnHistorico, limpiarHistorico, verEditor, vistaActiva,
    verEstructura, estructura, verHistorico, verReglas, reglas } = props;
  const nAreas = estructura ? estructura.filter(a => a.activa && a.posiciones > 0).length : 0;
  const nPos = estructura ? estructura.filter(a => a.activa).reduce((s, a) => s + (a.posiciones || 0), 0) : 0;
  return (
    <aside className="sidebar">
      <Section step="1" title="Planilla del mes">
        <button className="filebtn" onClick={onPlanillaClick}>
          {planillaInfo ? "Cambiar planilla" : "Subir planilla (.xlsm)"}
        </button>
        {planillaInfo
          ? <div className="filemeta">
              <strong>{planillaInfo.name}</strong>
              <span>Hoja «{planillaInfo.sheet}» · {planillaInfo.filas} filas</span>
            </div>
          : <p className="hint">El archivo que exporta la coordinadora con los turnos del personal.</p>}
      </Section>

      <Section step="2" title="Base de profesionales">
        <button className="filebtn ghost" onClick={onStaffClick}>
          {profs ? "Sustituir desde Excel" : "Cargar desde Excel"}
        </button>
        {profs
          ? <div className="filemeta">
              <strong>{profs.length} profesionales</strong>
              <span>{staffSource}</span>
            </div>
          : <p className="hint">Base compartida entre coordinadoras. Punto de partida: el libro con la hoja DATOS_PROFESIONALES. Luego se edita en el apartado 4 sin volver a Excel.</p>}
      </Section>

      <Section step="3" title="Estructura de la unidad">
        <button className={"filebtn ghost" + (vistaActiva === "estructura" ? " is-active" : "")}
          onClick={verEstructura}>
          {vistaActiva === "estructura" ? "Configurando unidad" : "Configurar unidad"}
        </button>
        {estructura
          ? <div className="filemeta">
              <strong>{nPos} posiciones · {nAreas} áreas</strong>
              <span>Camas operativas de esta semana</span>
            </div>
          : <p className="hint">Define qué unidades están abiertas y con cuántas posiciones.</p>}
      </Section>

      <Section step="R" title="Reglas de rotación">
        <button className={"filebtn ghost" + (vistaActiva === "reglas" ? " is-active" : "")}
          onClick={verReglas}>
          {vistaActiva === "reglas" ? "Configurando reglas" : "Configurar reglas"}
        </button>
        {reglas
          ? <div className="filemeta">
              <strong>{Object.values(reglas).filter(Boolean).length} de {Object.keys(reglas).length} reglas activas</strong>
              <span>Normas aplicadas al generar</span>
            </div>
          : <p className="hint">Elige qué normas de rotación se aplican al generar el cuadrante.</p>}
      </Section>

      <Section step="4" title="Periodo">
        <label className="field">
          <span>Lunes de inicio</span>
          <input type="date" value={fechaLunes}
            onChange={e => setFechaLunes(isoDate(lunesDeLaSemana(parseISO(e.target.value))))} />
        </label>
        <label className="field">
          <span>Semanas a generar</span>
          <input type="number" min="1" max="12" value={numSemanas}
            onChange={e => setNumSemanas(e.target.value)} />
        </label>
      </Section>

      <button className="generate" disabled={!puedeGenerar || busy} onClick={generar}>
        {busy ? "Generando…" : "Generar cuadrante"}
      </button>
      {!puedeGenerar && (
        <p className="hint center">Sube la planilla y la base de profesionales para empezar.</p>
      )}

      <Section step="5" title="Profesionales en plantilla">
        <button className={"filebtn ghost" + (vistaActiva === "editor" ? " is-active" : "")}
          onClick={verEditor} disabled={!profs}>
          {vistaActiva === "editor" ? "Editando plantilla" : "Ver y editar plantilla"}
        </button>
        <p className="hint">
          {profs
            ? "Modifica turno, criterio, antigüedad, suplencia y apodo de cada profesional. Los cambios se guardan y se aplican a las siguientes planillas."
            : "Disponible cuando haya una base cargada."}
        </p>
      </Section>

      <Section step="" title="Histórico de rotación">
        <button className={"filebtn ghost" + (vistaActiva === "historico" ? " is-active" : "")}
          onClick={verHistorico}>
          {vistaActiva === "historico" ? "Viendo histórico" : "Ver histórico de rotación"}
        </button>
        <p className="hint">
          {historicoLen > 0
            ? "Unidad de cada profesional mes a mes. Cada cuadrante que generas se registra aquí automáticamente."
            : "Aún no hay histórico. En cuanto generes un cuadrante, la rotación se registra aquí automáticamente."}
        </p>
        <div className="row-btns">
          <button className="minibtn" disabled={!resultado} onClick={() => guardarEnHistorico()}>Volver a guardar en histórico</button>
          <button className="minibtn danger" disabled={historicoLen === 0} onClick={limpiarHistorico}>Vaciar</button>
        </div>
      </Section>
    </aside>
  );
}
function Section({ step, title, children }) {
  return (
    <section className="block">
      <div className="block-head">
        {step && <span className="step">{step}</span>}
        <h3>{title}</h3>
      </div>
      {children}
    </section>
  );
}

/* ---------- Bienvenida / estado vacío ---------- */
function Welcome({ planilla, db }) {
  return (
    <div className="welcome">
      <h2>Listo para planificar</h2>
      <p>
        Esta herramienta reparte al personal de enfermería entre las quince posiciones físicas
        de la unidad — UCI A, UCI B, SEMIS, REA y UCCRI — partiendo de la planilla de turnos.
        Clasifica cada turno, respeta los criterios clínicos de cada profesional, prioriza por
        antigüedad y plantilla fija, cubre los huecos de transición entre turnos y rota las
        posiciones a lo largo del mes.
      </p>
      <div className="checklist">
        <Check ok={!!planilla} label="Planilla cargada" />
        <Check ok={!!db} label="Base de profesionales disponible" />
      </div>
      <p className="hint">Cuando ambos estén listos, pulsa «Generar cuadrante».</p>
    </div>
  );
}
function Check({ ok, label }) {
  return (
    <div className={"check " + (ok ? "on" : "")}>
      <span className="dot">{ok ? "✓" : ""}</span>{label}
    </div>
  );
}

/* ---------- Pestañas ---------- */
function Tabs({ vista, setVista, semanas, semanaActiva, setSemanaActiva, exportarExcel }) {
  const [hint, setHint] = useState(false);
  return (
    <div className="tabsbar">
      <div className="tabs">
        {[["cronograma","Cuadrante"],["decalajes","Decalajes"],["profesionales","Plantilla"]]
          .map(([k, l]) => (
          <button key={k} className={"tab " + (vista === k ? "active" : "")} onClick={() => setVista(k)}>{l}</button>
        ))}
      </div>
      <div className="tabsbar-right">
        {semanas.length > 1 && (vista === "cronograma" || vista === "decalajes") && (
          <div className="weekpick">
            {semanas.map((s, i) => (
              <button key={i} className={"weekbtn " + (semanaActiva === i ? "active" : "")}
                onClick={() => setSemanaActiva(i)}>
                Sem. {i + 1}<small>{ENGINE.fechaCorta(s.fechaLunes)}</small>
              </button>
            ))}
          </div>
        )}
        <div className="export-wrap">
          <button className="exportbtn" onClick={exportarExcel}
            onMouseEnter={() => setHint(true)} onMouseLeave={() => setHint(false)}>
            <svg viewBox="0 0 16 16" width="15" height="15" aria-hidden="true">
              <path d="M8 1v8m0 0L5 6.5M8 9l3-2.5M2.5 11v2.5h11V11" fill="none" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
            Exportar a Google Sheets
          </button>
          {hint && (
            <div className="export-hint">
              Descarga un <strong>.xlsx</strong> con los colores del cuadrante. Súbelo a Google Drive y elige
              «Abrir con → Hojas de cálculo de Google»: conserva colores y formato.
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

/* ---------- Cronograma (rejilla principal) ---------- */
const BLOQUES = [
  { idx: 0, label: "MAÑANA", suf: "MATI" },
  { idx: 1, label: "TARDE", suf: "TARDE" },
  { idx: 2, label: "NOCHE", suf: "NIT" },
];

function Cronograma({ semana, db, ENGINE, planilla, estructura, cargarAnotaciones, guardarAnotaciones }) {
  const fechas = useMemo(() => Array.from({ length: 7 }, (_, i) => ENGINE.addDays(semana.fechaLunes, i)), [semana, ENGINE]);
  const lunesISO = useMemo(() => isoDate(semana.fechaLunes), [semana]);

  // --- Anotaciones manuales (Fase B3) ---
  const [anotaciones, setAnotaciones] = useState({});
  const [editando, setEditando] = useState(null); // {d, t, p} o null
  const [modoAnotar, setModoAnotar] = useState(false);
  useEffect(() => {
    let vivo = true;
    (async () => {
      if (!cargarAnotaciones) return;
      try {
        const a = await cargarAnotaciones(lunesISO);
        if (vivo) setAnotaciones(a || {});
      } catch { if (vivo) setAnotaciones({}); }
    })();
    return () => { vivo = false; };
  }, [lunesISO, cargarAnotaciones]);

  const claveCelda = (d, t, p) => d + "-" + t + "-" + p;
  async function aplicarAnotacion(d, t, p, texto) {
    const k = claveCelda(d, t, p);
    const nuevo = { ...anotaciones };
    if (texto && texto.trim()) nuevo[k] = texto.trim();
    else delete nuevo[k];
    setAnotaciones(nuevo);
    setEditando(null);
    if (guardarAnotaciones) {
      try { await guardarAnotaciones(lunesISO, nuevo); } catch { /* se mantiene en memoria */ }
    }
  }

  // Layout de áreas a partir de la estructura real del cuadrante generado.
  const areasLayout = useMemo(() => {
    const est = estructura || ENGINE.estructuraActiva();
    return (est.areas || []).map(a => ({ id: a.id, nombre: a.nombre, rotacion: a.rotacion, pIni: a.posIni, pFin: a.posFin, etiquetas: a.etiquetas }));
  }, [estructura, ENGINE]);
  // Precalcular el turnoFinal y la etiqueta de cada profesional por día
  const turnoPorDiaNombre = useMemo(() => {
    const m = [];
    for (let d = 0; d < 7; d++) {
      const map = new Map();
      for (const r of semana.semana[d].regs) map.set(ENGINE.normalizarTexto(r.nombre), r.turnoFinal);
      m.push(map);
    }
    return m;
  }, [semana, ENGINE]);
  const etiquetaPorDiaNombre = useMemo(() => {
    const m = [];
    for (let d = 0; d < 7; d++) {
      const map = new Map();
      for (const r of semana.semana[d].regs) if (r.etiqueta) map.set(ENGINE.normalizarTexto(r.nombre), r.etiqueta);
      m.push(map);
    }
    return m;
  }, [semana, ENGINE]);

  function etiquetaCorta(et) {
    if (et === "doblaje") return "DBL";
    if (et === "+") return "+";
    if (et === "parcial") return "½";
    return "";
  }

  // Color según el TIPO de jornada del profesional ese día:
  //  - 12h (día o noche) y doblaje -> lila
  //  - 7h (mañana/tarde), 10h noche y parcial -> BLANCO
  //  - sin cubrir -> rojo
  function cellColor(nombre, d, bloqueIdx) {
    if (!nombre) return { background: "var(--empty)", color: "transparent" };
    // 1) El color lo manda el FONDO del profesional (configurable en la plantilla).
    const f = db.fondo(nombre);
    if (f) return { background: f === ENGINE.FONDO.LILA ? ENGINE.COLOR_LILA : ENGINE.COLOR_BLANCO };
    // 2) Si no está en la base, se deduce del tipo de jornada de ese día.
    const tf = turnoPorDiaNombre[d].get(ENGINE.normalizarTexto(nombre));
    const esDoce = tf === ENGINE.TURNO.D12 || tf === ENGINE.TURNO.N12 || tf === ENGINE.TURNO.DOBLAJE;
    return { background: esDoce ? ENGINE.COLOR_LILA : ENGINE.COLOR_BLANCO };
  }

  return (
    <div className="crono-wrap">
      <div className="crono-toolbar">
        <div className="crono-legend">
          <LegendItem color="#D2BEE6" label="Fondo lila (12 h)" />
          <LegendItem color="#fff" label="Fondo blanco (7 h / 10 h)" border />
          <LegendItem color="var(--empty)" label="Sin cubrir" />
          <span className="legend-item"><span className="badge badge-dbl">DBL</span> Doblaje</span>
          <span className="legend-item"><span className="badge badge-plus">+</span> Extra</span>
          <span className="legend-item"><span className="badge badge-half">½</span> Media jornada</span>
        </div>
        <button className={"anotar-btn" + (modoAnotar ? " on" : "")}
          onClick={() => { setModoAnotar(v => !v); setEditando(null); }}>
          {modoAnotar ? "✓ Anotando (clic en una celda)" : "✎ Añadir anotaciones"}
        </button>
      </div>
      {modoAnotar && (
        <div className="anotar-help">
          Haz clic en cualquier celda para añadir una nota (por ejemplo <strong>COT</strong>,
          <strong> COF</strong>, o una franja como <strong>8-14h</strong>). Estas anotaciones
          las pones tú a mano; el sistema no puede deducirlas del Excel. Se guardan y las ven
          el resto de coordinadoras.
        </div>
      )}
      {BLOQUES.map(b => (
        <div key={b.idx} className="bloque">
          <div className="bloque-title" data-suf={b.suf}>
            <span>{b.label}</span><small>{b.suf}</small>
          </div>
          <div className="grid">
            <table className="cuadro">
              <thead>
                <tr>
                  <th className="area-h"></th>
                  <th className="num-h"></th>
                  {fechas.map((f, i) => (
                    <th key={i} className="day-h">
                      <span className="dow">{DIAS_CORTOS[i]}</span>
                      <span className="dnum">{f.getDate()}/{f.getMonth() + 1}</span>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {areasLayout.map((area, areaIdx) => {
                  const rows = [];
                  let local = 1;
                  // ¿Es esta la última área de su rotación? (para no duplicar excedentes
                  // de UCCRI cuando hay esquerra y dreta, ambas rotación "UCCRI").
                  const esUltimaDeRotacion = !areasLayout.slice(areaIdx + 1).some(a => a.rotacion === area.rotacion);
                  // Excedentes de esta ROTACIÓN: máximo en los 7 días (para el rowSpan).
                  let maxExced = 0;
                  if (esUltimaDeRotacion) {
                    for (let d = 0; d < 7; d++) {
                      const lista = (semana.asig.excedentes[d * 3 + b.idx] || []).filter(e => e.area === area.rotacion);
                      if (lista.length > maxExced) maxExced = lista.length;
                    }
                  }
                  const filasArea = (area.pFin - area.pIni + 1) + maxExced;
                  for (let p = area.pIni; p <= area.pFin; p++) {
                    const isFirst = p === area.pIni;
                    const etFija = ENGINE.etiquetaDePosicion(area, local);
                    rows.push(
                      <tr key={p} className={isFirst ? "area-start" : ""}>
                        {isFirst && (
                          <td className="area-cell" rowSpan={filasArea}>
                            <span className="area-name">{area.nombre}</span>
                          </td>
                        )}
                        <td className={"num-cell" + (etFija ? " num-label" : "")}>
                          {etFija && etFija.reemplazaNumero
                            ? <span className="pos-fija">{etFija.texto}</span>
                            : <>{local}{etFija && <span className="pos-fija-sub">{etFija.texto}</span>}</>}
                        </td>
                        {fechas.map((f, d) => {
                          const nombre = semana.asig.celdas[d][b.idx][p];
                          const st = cellColor(nombre, d, b.idx);
                          const et = nombre ? etiquetaPorDiaNombre[d].get(ENGINE.normalizarTexto(nombre)) : "";
                          const etTxt = etiquetaCorta(et);
                          const k = claveCelda(d, b.idx, p);
                          const nota = anotaciones[k];
                          const estaEditando = editando && editando.d === d && editando.t === b.idx && editando.p === p;
                          return (
                            <td key={d} className={"data-cell" + (modoAnotar ? " anotable" : "") + (nota ? " con-nota" : "")} style={st}
                              title={nombre ? (nombre + (et ? " · " + et : "") + (nota ? " · " + nota : "")) : (nota || "Sin cubrir")}
                              onClick={modoAnotar && !estaEditando ? () => setEditando({ d, t: b.idx, p }) : undefined}>
                              {nombre ? db.apodo(nombre) : ""}
                              {etTxt && <span className={"badge badge-" + (et === "doblaje" ? "dbl" : et === "+" ? "plus" : "half")}>{etTxt}</span>}
                              {nota && !estaEditando && <span className="nota-cell">{nota}</span>}
                              {estaEditando && (
                                <CeldaAnotacion inicial={nota || ""}
                                  onGuardar={(txt) => aplicarAnotacion(d, b.idx, p, txt)}
                                  onCancelar={() => setEditando(null)} />
                              )}
                            </td>
                          );
                        })}
                      </tr>
                    );
                    local++;
                  }
                  // Filas de EXCEDENTES de esta unidad (posiciones extra sin número).
                  for (let e = 0; e < maxExced; e++) {
                    rows.push(
                      <tr key={"exc-" + area.id + "-" + e} className="exc-row">
                        <td className="num-cell exc-num" title="Excedente">+</td>
                        {fechas.map((f, d) => {
                          const lista = (semana.asig.excedentes[d * 3 + b.idx] || []).filter(x => x.area === area.rotacion);
                          const ex = lista[e];
                          return (
                            <td key={d} className="data-cell exc-cell"
                              style={ex ? cellColor(ex.nombre, d, b.idx) : undefined}
                              title={ex ? ex.nombre + " · excedente" : ""}>
                              {ex ? db.apodo(ex.nombre) : ""}
                            </td>
                          );
                        })}
                      </tr>
                    );
                  }
                  // Fila en blanco de separación entre unidades (no tras la última)
                  if (areaIdx < areasLayout.length - 1) {
                    rows.push(
                      <tr key={"sep-" + area.id} className="sep-row" aria-hidden="true">
                        <td className="sep-cell" colSpan={2 + 7}></td>
                      </tr>
                    );
                  }
                  return rows;
                })}
              </tbody>
            </table>
          </div>
        </div>
      ))}
    </div>
  );
}
/* Edición en línea de una anotación de celda (Fase B3) */
function CeldaAnotacion({ inicial, onGuardar, onCancelar }) {
  const [txt, setTxt] = useState(inicial);
  const ref = useRef(null);
  useEffect(() => { if (ref.current) { ref.current.focus(); ref.current.select(); } }, []);
  return (
    <div className="nota-edit" onClick={e => e.stopPropagation()}>
      <input ref={ref} value={txt} maxLength={14} placeholder="COT, 8-14h…"
        onChange={e => setTxt(e.target.value)}
        onKeyDown={e => {
          if (e.key === "Enter") onGuardar(txt);
          if (e.key === "Escape") onCancelar();
        }} />
      <div className="nota-edit-btns">
        <button className="nota-ok" onClick={() => onGuardar(txt)}>Guardar</button>
        <button className="nota-x" onClick={onCancelar}>✕</button>
      </div>
    </div>
  );
}

function LegendItem({ color, label, border }) {
  return (
    <span className="legend-item">
      <span className="swatch" style={{ background: color, border: border ? "1px solid var(--line)" : "none" }} />
      {label}
    </span>
  );
}

/* ---------- Decalajes ---------- */
function Decalajes({ avisos, db, ENGINE }) {
  if (!avisos.length) return <div className="empty-note">No se detectaron decalajes esta semana.</div>;
  // agrupar por día
  const porDia = new Map();
  for (const a of avisos) {
    const k = ENGINE.fechaCorta(a.fecha);
    if (!porDia.has(k)) porDia.set(k, []);
    porDia.get(k).push(a);
  }
  return (
    <div className="decalajes">
      <p className="section-intro">
        Huecos de transición entre turnos y quién los cubre. En verde, cubierto; en rojo, sin cobertura.
      </p>
      {[...porDia.entries()].map(([dia, items]) => (
        <div key={dia} className="dec-day">
          <h4>{dia}</h4>
          <table className="dec-table">
            <thead>
              <tr><th>Posición</th><th>Área</th><th>Decalaje</th><th>Cubierto por</th></tr>
            </thead>
            <tbody>
              {items.map((a, i) => (
                <tr key={i} className={a.cubierto ? "ok" : "ko"}>
                  <td>{a.pos}</td>
                  <td>{a.area.replace("_", " ")}</td>
                  <td>{a.tipo}</td>
                  <td>{a.cubierto ? db.apodo(a.cubiertoPor) : "— sin cubrir —"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ))}
    </div>
  );
}

/* ---------- Informe mensual ---------- */
/* ---------- Histórico de rotación (apartado propio, vista visual) ---------- */
const MESES_ES = ["Gener", "Febrer", "Març", "Abril", "Maig", "Juny", "Juliol", "Agost", "Setembre", "Octubre", "Novembre", "Desembre"];
function nombreMes(mesKey) {
  const partes = String(mesKey || "").split("-");
  if (partes.length !== 2) return mesKey;
  const m = parseInt(partes[1], 10) - 1;
  return (MESES_ES[m] || mesKey) + " " + partes[0];
}
function colorUnidad(u) {
  switch (u) {
    case "UCI A": return { bg: "#dce9f7", fg: "#1c4e8a", bd: "#b4cdec" };
    case "UCI B": return { bg: "#e2dcf3", fg: "#4a328a", bd: "#cabdee" };
    case "SEMIS": return { bg: "#dcefe0", fg: "#1f6a3a", bd: "#bfe0c8" };
    case "REA": return { bg: "#fdecd2", fg: "#9a6111", bd: "#f3d9a9" };
    case "UCCRI": return { bg: "#f7dce6", fg: "#8a2c50", bd: "#eebdce" };
    default: return { bg: "#e9e9e9", fg: "#555", bd: "#d4d4d4" };
  }
}
function ChipUnidad({ u }) {
  const c = colorUnidad(u);
  return <span className="uni-chip" style={{ background: c.bg, color: c.fg, borderColor: c.bd }}>{u}</span>;
}

/* Buscador con autocompletado para elegir un profesional de la base. */
function BuscadorProfesional({ profesionales, valor, onElegir }) {
  const [texto, setTexto] = useState(valor || "");
  const [abierto, setAbierto] = useState(false);
  const [idx, setIdx] = useState(0);
  const ref = useRef(null);

  useEffect(() => { setTexto(valor || ""); }, [valor]);

  const norm = (s) => String(s || "").toUpperCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const filtrados = useMemo(() => {
    const q = norm(texto).trim();
    if (!q) return profesionales.slice(0, 30);
    return profesionales.filter(p => norm(p.nombre).includes(q) || norm(p.apodo).includes(q)).slice(0, 30);
  }, [texto, profesionales]);

  useEffect(() => {
    function fuera(e) { if (ref.current && !ref.current.contains(e.target)) setAbierto(false); }
    document.addEventListener("mousedown", fuera);
    return () => document.removeEventListener("mousedown", fuera);
  }, []);

  function elegir(p) {
    onElegir(p.nombre);
    setTexto(p.nombre);
    setAbierto(false);
  }

  const elegido = valor && texto === valor;

  return (
    <div className="buscador" ref={ref}>
      <input
        className={"buscador-input" + (elegido ? " elegido" : "")}
        value={texto}
        placeholder="Escribe para buscar un profesional…"
        onChange={e => { setTexto(e.target.value); setAbierto(true); setIdx(0); if (valor) onElegir(""); }}
        onFocus={() => setAbierto(true)}
        onKeyDown={e => {
          if (!abierto && (e.key === "ArrowDown" || e.key === "ArrowUp")) { setAbierto(true); return; }
          if (e.key === "ArrowDown") { e.preventDefault(); setIdx(i => Math.min(i + 1, filtrados.length - 1)); }
          else if (e.key === "ArrowUp") { e.preventDefault(); setIdx(i => Math.max(i - 1, 0)); }
          else if (e.key === "Enter") { e.preventDefault(); if (filtrados[idx]) elegir(filtrados[idx]); }
          else if (e.key === "Escape") setAbierto(false);
        }}
      />
      {elegido && <span className="buscador-ok">✓</span>}
      {abierto && filtrados.length > 0 && (
        <ul className="buscador-lista">
          {filtrados.map((p, i) => (
            <li key={i} className={i === idx ? "activo" : ""}
              onMouseEnter={() => setIdx(i)} onMouseDown={() => elegir(p)}>
              <span className="buscador-nombre">{p.nombre}</span>
              {p.apodo && <span className="buscador-apodo">{p.apodo}</span>}
            </li>
          ))}
        </ul>
      )}
      {abierto && filtrados.length === 0 && texto.trim() && (
        <ul className="buscador-lista"><li className="buscador-vacio">Sin coincidencias</li></ul>
      )}
    </div>
  );
}

/* ---------- Vinculación de profesionales sin reconocer (al generar) ---------- */
function VincularProfesionales({ nombres, profs, ENGINE, onAplicar, onCancelar, busy }) {
  // Por cada nombre sin vincular: decisión "vincular" (a un profesional existente)
  // o "nuevo" (añadir a la base). Estado por fila.
  const profsOrden = useMemo(
    () => (profs || []).slice().sort((a, b) => String(a.nombre).localeCompare(String(b.nombre))),
    [profs]
  );
  // estado: { [norm]: { modo:'vincular'|'nuevo'|'', vinculadoA:'', nuevo:{...} } }
  const [dec, setDec] = useState(() => {
    const init = {};
    for (const n of nombres) {
      init[ENGINE.normalizarTexto(n.raw)] = {
        modo: "", vinculadoA: "",
        nuevo: { nombre: n.raw.indexOf(",") >= 0 ? n.raw : "", turno: "TM", apodo: n.raw, criterio: "SI", suplente: "No", antiguedad: "< 3 ANY" },
      };
    }
    return init;
  });

  const setFila = (norm, patch) => setDec(prev => ({ ...prev, [norm]: { ...prev[norm], ...patch } }));
  const setNuevoCampo = (norm, campo, valor) =>
    setDec(prev => ({ ...prev, [norm]: { ...prev[norm], nuevo: { ...prev[norm].nuevo, [campo]: valor } } }));

  // ¿Está todo decidido? (cada fila con modo y datos válidos)
  const pendientes = nombres.filter(n => {
    const d = dec[ENGINE.normalizarTexto(n.raw)];
    if (!d || !d.modo) return true;
    if (d.modo === "vincular") return !d.vinculadoA;
    if (d.modo === "nuevo") return !d.nuevo.nombre.trim();
    return true;
  });

  function aplicar() {
    const aliasNuevos = {};
    const profsNuevos = [];
    for (const n of nombres) {
      const d = dec[ENGINE.normalizarTexto(n.raw)];
      if (!d || !d.modo) continue;
      if (d.modo === "vincular" && d.vinculadoA) {
        aliasNuevos[n.raw] = d.vinculadoA;
      } else if (d.modo === "nuevo" && d.nuevo.nombre.trim()) {
        // el apodo por defecto es la abreviatura de la planilla, para que case
        profsNuevos.push({ ...d.nuevo, apodo: d.nuevo.apodo || n.raw });
        // además, si el texto de la planilla es una abreviatura (sin coma) y el
        // nombre nuevo tiene coma, creamos alias abreviatura→nombre para vincularlos
        if (n.raw.indexOf(",") < 0 && d.nuevo.nombre.indexOf(",") >= 0) {
          aliasNuevos[n.raw] = d.nuevo.nombre;
        }
      }
    }
    onAplicar(aliasNuevos, profsNuevos);
  }

  return (
    <div className="vincular">
      <div className="editor-head">
        <div>
          <h2>Profesionales sin reconocer</h2>
          <p className="section-intro">
            Estos nombres aparecen en la planilla pero no se han podido vincular con la base
            de profesionales. Para cada uno, vincúlalo con un profesional existente (si es una
            abreviatura) o añádelo como nuevo. Así el cronograma quedará completo.
          </p>
        </div>
        <button className="linkbtn" onClick={onCancelar} disabled={busy}>← Cancelar</button>
      </div>

      <div className="vinc-list">
        {nombres.map(n => {
          const norm = ENGINE.normalizarTexto(n.raw);
          const d = dec[norm];
          return (
            <div key={norm} className="vinc-card">
              <div className="vinc-head">
                <div className="vinc-name">{n.raw}</div>
                <div className="vinc-meta">aparece {n.veces}× · p.ej. «{n.ejemploHorario}»</div>
              </div>
              <div className="vinc-modos">
                <button className={"vinc-tab" + (d.modo === "vincular" ? " on" : "")}
                  onClick={() => setFila(norm, { modo: "vincular" })}>Vincular a existente</button>
                <button className={"vinc-tab" + (d.modo === "nuevo" ? " on" : "")}
                  onClick={() => setFila(norm, { modo: "nuevo" })}>Añadir como nuevo</button>
              </div>

              {d.modo === "vincular" && (
                <div className="vinc-body">
                  <label className="vinc-lbl">Es la misma persona que:</label>
                  <BuscadorProfesional
                    profesionales={profsOrden}
                    valor={d.vinculadoA}
                    onElegir={(nombre) => setFila(norm, { vinculadoA: nombre })}
                  />
                </div>
              )}

              {d.modo === "nuevo" && (
                <div className="vinc-body vinc-form">
                  <div className="vinc-row">
                    <label>Nombre completo<span className="req">*</span></label>
                    <input value={d.nuevo.nombre} placeholder="Apellidos, Nombre"
                      onChange={e => setNuevoCampo(norm, "nombre", e.target.value)} />
                  </div>
                  <div className="vinc-row">
                    <label>Apodo (cómo se muestra)</label>
                    <input value={d.nuevo.apodo}
                      onChange={e => setNuevoCampo(norm, "apodo", e.target.value)} />
                  </div>
                  <div className="vinc-grid">
                    <div className="vinc-row">
                      <label>Turno</label>
                      <select value={d.nuevo.turno} onChange={e => setNuevoCampo(norm, "turno", e.target.value)}>
                        <option>TM</option><option>TT</option><option>TN</option>
                        <option>12H Día</option><option>12H Noche</option>
                      </select>
                    </div>
                    <div className="vinc-row">
                      <label>Criterio</label>
                      <select value={d.nuevo.criterio} onChange={e => setNuevoCampo(norm, "criterio", e.target.value)}>
                        <option>SI</option><option>NO</option><option>UCI</option><option>UCCRI</option>
                      </select>
                    </div>
                    <div className="vinc-row">
                      <label>Suplente</label>
                      <select value={d.nuevo.suplente} onChange={e => setNuevoCampo(norm, "suplente", e.target.value)}>
                        <option>No</option><option>Si</option>
                      </select>
                    </div>
                    <div className="vinc-row">
                      <label>Antigüedad</label>
                      <select value={d.nuevo.antiguedad} onChange={e => setNuevoCampo(norm, "antiguedad", e.target.value)}>
                        <option>{"< 3 ANY"}</option>
                        <option>{"> 3 ANY < 10 ANY"}</option>
                        <option>{"10 - 20 ANY"}</option>
                        <option>{"> 20 ANY"}</option>
                      </select>
                    </div>
                  </div>
                </div>
              )}
            </div>
          );
        })}
      </div>

      <div className="vinc-footer">
        <div className="vinc-status">
          {pendientes.length > 0
            ? pendientes.length + " de " + nombres.length + " sin decidir"
            : "Todo listo para generar"}
        </div>
        <div className="vinc-actions">
          <button className="btn-secondary" onClick={onCancelar} disabled={busy}>Cancelar</button>
          <button className="btn-primary" onClick={aplicar} disabled={busy || pendientes.length > 0}>
            {busy ? "Guardando…" : "Vincular y generar"}
          </button>
        </div>
      </div>
    </div>
  );
}

function HistoricoRotacion({ historico, ENGINE, onCerrar, onImportar }) {
  const datos = useMemo(() => ENGINE.resumenHistoricoRotacion(historico || []), [historico, ENGINE]);
  const [busca, setBusca] = useState("");
  const meses = useMemo(() => {
    const set = new Set();
    for (const p of datos) for (const m of p.meses) set.add(m.mes);
    return [...set].sort();
  }, [datos]);
  const filtrados = useMemo(() => {
    const q = busca.trim().toUpperCase();
    if (!q) return datos;
    return datos.filter(p => p.profesional.toUpperCase().includes(q));
  }, [datos, busca]);

  return (
    <div className="historico">
      <div className="editor-head">
        <div>
          <h2>Histórico de rotación</h2>
          <p className="section-intro">
            Unidad donde ha estado cada profesional mes a mes. Cuando ha pasado la mayor parte
            del mes (66% o más de sus días) en una unidad, aparece solo esa; si ha estado más
            repartido, aparecen las dos unidades donde más tiempo ha pasado.
          </p>
        </div>
        <div className="hist-head-acciones">
          {onImportar && (
            <button className="minibtn" onClick={onImportar} title="Cargar el registro de rotación llevado a mano">
              Importar desde Excel
            </button>
          )}
          <button className="linkbtn" onClick={onCerrar}>← Volver</button>
        </div>
      </div>

      {datos.length === 0 ? (
        <div className="empty-note">
          Todavía no hay histórico. En cuanto generes un cuadrante se registra aquí automáticamente
          en el panel lateral para empezar a registrar la rotación mensual.
        </div>
      ) : (
        <>
          <input className="search" placeholder="Buscar profesional…" value={busca}
            onChange={e => setBusca(e.target.value)} style={{ marginBottom: 16 }} />
          <div className="hist-grid-wrap">
            <table className="hist-grid">
              <thead>
                <tr>
                  <th className="hist-prof-h">Profesional</th>
                  {meses.map(m => <th key={m} className="hist-mes-h">{nombreMes(m)}</th>)}
                </tr>
              </thead>
              <tbody>
                {filtrados.map(p => {
                  const porMes = new Map(p.meses.map(m => [m.mes, m.unidades]));
                  return (
                    <tr key={p.profesional}>
                      <td className="hist-prof">{p.profesional}</td>
                      {meses.map(m => {
                        const unidades = porMes.get(m);
                        return (
                          <td key={m} className="hist-cell">
                            {unidades
                              ? <div className="hist-chips">{unidades.map((u, i) => <ChipUnidad key={i} u={u} />)}</div>
                              : <span className="hist-vacio">·</span>}
                          </td>
                        );
                      })}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
          <p className="foot-hint" style={{ marginTop: 16 }}>
            UCCRI agrupa esquerra y dreta: a efectos de rotación cuentan como una sola unidad.
          </p>
        </>
      )}
    </div>
  );
}

function Informe({ informe }) {
  if (!informe.length) return <div className="empty-note">Sin datos de informe. Genera cuadrantes y guárdalos en el histórico para ver la rotación mensual.</div>;
  const dia = informe.filter(f => f.ambito === "DIA");
  const nit = informe.filter(f => f.ambito === "NIT");
  const Tabla = ({ filas }) => (
    <table className="inf-table">
      <thead><tr><th>Mes</th><th>Profesional</th><th>Unidad / rotación</th></tr></thead>
      <tbody>
        {filas.map((f, i) => (
          <tr key={i}>
            <td>{f.mes}</td>
            <td>{f.profesional}</td>
            <td>{f.resumen}{f.roto && <span className="roto-tag">rota</span>}</td>
          </tr>
        ))}
      </tbody>
    </table>
  );
  return (
    <div className="informe">
      <p className="section-intro">
        Rotación mensual de cada profesional: la unidad principal donde ha estado cada mes,
        separada en jornada de día y de noche. Si ha pasado por más de una unidad, se marca «rota».
      </p>
      <div className="inf-bloque">
        <h3 className="inf-titulo"><span className="inf-pill dia">DÍA</span> Rotación de jornada de día</h3>
        {dia.length ? <Tabla filas={dia} /> : <div className="empty-note small">Sin registros de día.</div>}
      </div>
      <div className="inf-bloque">
        <h3 className="inf-titulo"><span className="inf-pill nit">NOCHE</span> Rotación de jornada de noche</h3>
        {nit.length ? <Tabla filas={nit} /> : <div className="empty-note small">Sin registros de noche.</div>}
      </div>
    </div>
  );
}

/* ---------- Editor de estructura de la unidad (apartado 3) ----------
   La coordinadora marca/desmarca qué unidades están abiertas y ajusta
   el nº de posiciones de cada una. Se guarda en la base compartida.
*/
/* ---------- Importar el registro de rotación de la coordinadora ---------- */
function ImportarRotaciones({ datos, db, profs, ENGINE, onAplicar, onCancelar, busy }) {
  // Año por defecto: el Excel solo trae el nombre del mes. Si el último mes con
  // datos aún no ha llegado este año, lo más probable es que sea del año pasado.
  const anioSugerido = useMemo(() => {
    const hoy = new Date();
    let maxMes = 0;
    for (const h of (datos.hojas || [])) {
      (h.encabezados || []).forEach((enc, i) => {
        const m = ENGINE.mesDesdeEncabezado(enc);
        if (!m) return;
        const hayDatos = (h.filas || []).some(f => String((f.celdas || [])[i] || "").trim());
        if (hayDatos && m > maxMes) maxMes = m;
      });
    }
    const anioHoy = hoy.getFullYear();
    return String(maxMes > hoy.getMonth() + 1 ? anioHoy - 1 : anioHoy);
  }, [datos, ENGINE]);
  const [anio, setAnio] = useState(anioSugerido);
  const [vinc, setVinc] = useState({});   // apodoExcel -> nombre completo BD
  const [omitidos, setOmitidos] = useState({}); // apodoExcel -> true (no importar)

  const profsOrden = useMemo(
    () => (profs || []).slice().sort((a, b) => String(a.nombre).localeCompare(String(b.nombre))),
    [profs]
  );

  const analisis = useMemo(
    () => ENGINE.analizarImportacionRotaciones(datos.hojas, db, anio, vinc),
    [datos, db, anio, vinc, ENGINE]
  );

  const pendientes = analisis.sinVincular.filter(s => !omitidos[s.apodo] && !vinc[s.apodo]);
  const mesesTexto = analisis.meses.map(m => {
    const [y, mm] = m.split("-");
    return MESES_ES[parseInt(mm, 10) - 1] + " " + y;
  });

  function aplicar() {
    const filas = ENGINE.filasHistoricoDesdeRotaciones(analisis, anio);
    onAplicar(filas, analisis.meses);
  }

  return (
    <div className="importar">
      <div className="editor-head">
        <div>
          <h2>Importar rotación desde Excel</h2>
          <p className="section-intro">
            Carga el registro que lleva la coordinadora a mano (hojas DIA y NIT) para que la
            app parta de esa rotación como histórico. A partir de ahí, los cuadrantes que
            generes continuarán la rotación desde donde lo dejó.
          </p>
        </div>
        <button className="linkbtn" onClick={onCancelar} disabled={busy}>← Cancelar</button>
      </div>

      <div className="imp-file">
        <strong>{datos.fileName}</strong>
        <span>{datos.hojas.map(h => h.nombreHoja + " (" + h.filas.length + " filas)").join(" · ")}</span>
      </div>

      <div className="imp-anio">
        <label>Año de estos meses</label>
        <input type="number" min="2000" max="2100" value={anio}
          onChange={e => setAnio(e.target.value)} />
        <span className="imp-hint">
          El Excel indica el mes pero no el año. Comprueba que sea el correcto.
        </span>
      </div>

      <div className="imp-resumen">
        <div className="imp-dato">
          <strong>{analisis.nProfesionales}</strong>
          <span>profesionales reconocidos</span>
        </div>
        <div className="imp-dato">
          <strong>{analisis.nRegistros}</strong>
          <span>registros mes/unidad</span>
        </div>
        <div className="imp-dato">
          <strong>{analisis.meses.length}</strong>
          <span>meses detectados</span>
        </div>
      </div>

      {mesesTexto.length > 0 && (
        <p className="imp-meses">Meses que se importarán: <strong>{mesesTexto.join(", ")}</strong>.
          Si ya tenías histórico de esos meses, se reemplazará.</p>
      )}

      {analisis.sinVincular.length > 0 && (
        <div className="imp-bloque">
          <h3>Profesionales sin reconocer ({analisis.sinVincular.length})</h3>
          <p className="imp-sub">
            Estos nombres del Excel no se han podido emparejar con la base. Vincúlalos o
            márcalos para omitirlos (no se importarán).
          </p>
          {analisis.sinVincular.map(s => (
            <div key={s.apodo} className={"imp-vinc" + (omitidos[s.apodo] ? " omitido" : "")}>
              <div className="imp-vinc-nom">
                <strong>{s.apodo}</strong>
                <span className="imp-vinc-meta">{s.bloque} · {s.meses} mes(es) con dato</span>
              </div>
              <div className="imp-vinc-sel">
                {!omitidos[s.apodo] ? (
                  <BuscadorProfesional
                    profesionales={profsOrden}
                    valor={vinc[s.apodo] || ""}
                    onElegir={(nombre) => setVinc(prev => ({ ...prev, [s.apodo]: nombre }))}
                  />
                ) : <span className="imp-omitido-txt">Se omitirá</span>}
              </div>
              <button className="minibtn"
                onClick={() => setOmitidos(prev => ({ ...prev, [s.apodo]: !prev[s.apodo] }))}>
                {omitidos[s.apodo] ? "Recuperar" : "Omitir"}
              </button>
            </div>
          ))}
        </div>
      )}

      {analisis.noReconocidas.length > 0 && (
        <div className="imp-bloque">
          <h3>Celdas que no se han entendido ({analisis.noReconocidas.length})</h3>
          <p className="imp-sub">
            Estos textos no corresponden a ninguna unidad conocida, así que se ignorarán.
            El resto de meses de esos profesionales sí se importa.
          </p>
          <ul className="imp-raras">
            {analisis.noReconocidas.slice(0, 12).map((n, i) => (
              <li key={i}><strong>{n.apodo}</strong> · {n.mesNombre}: «{n.texto}»</li>
            ))}
            {analisis.noReconocidas.length > 12 && <li>… y {analisis.noReconocidas.length - 12} más</li>}
          </ul>
        </div>
      )}

      <div className="vinc-footer">
        <div className="vinc-status">
          {pendientes.length > 0
            ? pendientes.length + " sin vincular (se omitirán si no los vinculas)"
            : "Todo listo para importar"}
        </div>
        <div className="vinc-actions">
          <button className="btn-secondary" onClick={onCancelar} disabled={busy}>Cancelar</button>
          <button className="btn-primary" onClick={aplicar}
            disabled={busy || analisis.nRegistros === 0}>
            {busy ? "Importando…" : "Importar como histórico"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ---------- Editor de REGLAS DE ROTACIÓN (ítems marcables) ---------- */
function EditorReglas({ reglas, ENGINE, saveReglas, onCerrar }) {
  const [estado, setEstado] = useState(() => ({ ...ENGINE.REGLAS_DEFECTO, ...(reglas || {}) }));
  const [busy, setBusy] = useState(false);
  const [dirty, setDirty] = useState(false);
  const [msg, setMsg] = useState("");

  function toggle(clave) {
    setEstado(prev => ({ ...prev, [clave]: !prev[clave] }));
    setDirty(true); setMsg("");
  }
  async function guardar() {
    setBusy(true); setMsg("");
    try {
      await saveReglas(estado);
      setDirty(false); setMsg("ok");
      setTimeout(() => setMsg(""), 2600);
    } catch (err) { setMsg(err.message || "No se pudo guardar."); }
    finally { setBusy(false); }
  }
  function restaurar() {
    setEstado({ ...ENGINE.REGLAS_DEFECTO });
    setDirty(true); setMsg("");
  }

  const activasN = Object.values(estado).filter(Boolean).length;

  return (
    <div className="reglas">
      <div className="editor-head">
        <div>
          <h2>Reglas de rotación</h2>
          <p className="section-intro">
            Marca qué normas quieres que se apliquen al generar el cuadrante. Los cambios
            se guardan y afectan a las próximas generaciones. El criterio de acceso de cada
            profesional (SI / NO / UCI / UCCRI) se respeta siempre, no es desactivable.
          </p>
        </div>
        <button className="linkbtn" onClick={onCerrar}>← Volver</button>
      </div>

      <div className="reglas-list">
        {ENGINE.REGLAS_INFO.map(r => {
          const apagadaPorDependencia = r.dependeDe && !estado[r.dependeDe];
          return (
            <div key={r.clave}
              className={"regla-card" + (estado[r.clave] ? " on" : "") + (apagadaPorDependencia ? " dep-off" : "")}>
              <button className={"chk grande " + (estado[r.clave] ? "on" : "")}
                onClick={() => toggle(r.clave)}
                disabled={busy}
                aria-label={estado[r.clave] ? "Desactivar" : "Activar"}>
                {estado[r.clave] ? "✓" : ""}
              </button>
              <div className="regla-txt">
                <div className="regla-titulo">{r.titulo}</div>
                <div className="regla-desc">{r.desc}</div>
                {apagadaPorDependencia && (
                  <div className="regla-aviso">Sin efecto mientras «Rotación mensual obligatoria» esté desactivada.</div>
                )}
              </div>
            </div>
          );
        })}
      </div>

      <div className="vinc-footer">
        <div className="vinc-status">{activasN} de {ENGINE.REGLAS_INFO.length} reglas activas{msg === "ok" ? " · Guardado ✓" : (msg ? " · " + msg : "")}</div>
        <div className="vinc-actions">
          <button className="btn-secondary" onClick={restaurar} disabled={busy}>Restaurar todas</button>
          <button className="btn-primary" onClick={guardar} disabled={busy || !dirty}>
            {busy ? "Guardando…" : "Guardar reglas"}
          </button>
        </div>
      </div>
    </div>
  );
}

function EditorEstructura({ estructura, catalogo, porDefecto, saveEstructura, onCerrar, ENGINE }) {
  // Copia local: unidades del catálogo (completadas con lo guardado) + unidades
  // personalizadas que estén en la estructura guardada pero no en el catálogo.
  const inicial = () => {
    const filasCat = catalogo.map(cat => {
      const guardada = (estructura || []).find(a => a.id === cat.id);
      return guardada
        ? { ...cat, nombre: guardada.nombre || cat.nombre, posiciones: guardada.posiciones,
            activa: guardada.activa, etiquetas: guardada.etiquetas || null, esCatalogo: true }
        : { ...cat, posiciones: cat.posDefecto, activa: false, etiquetas: null, esCatalogo: true };
    });
    const personalizadas = (estructura || [])
      .filter(a => !catalogo.some(c => c.id === a.id))
      .map(a => ({ ...a, esCatalogo: false }));
    return [...filasCat, ...personalizadas];
  };
  const [filas, setFilas] = useState(inicial);
  const [busy, setBusy] = useState(false);
  const [dirty, setDirty] = useState(false);
  const [msg, setMsg] = useState("");
  const [abiertaEtiquetas, setAbiertaEtiquetas] = useState(null); // id del área con etiquetas desplegadas

  const marca = () => { setDirty(true); setMsg(""); };
  function setCampo(id, campo, val) {
    setFilas(prev => prev.map(f => f.id === id ? { ...f, [campo]: val } : f));
    marca();
  }
  function setActiva(id, val) { setCampo(id, "activa", val); }
  function setPos(id, val) {
    const n = Math.max(0, Math.min(12, parseInt(val, 10) || 0));
    setFilas(prev => prev.map(f => f.id === id ? { ...f, posiciones: n } : f));
    marca();
  }
  function setEtiqueta(id, posLocal, texto) {
    setFilas(prev => prev.map(f => {
      if (f.id !== id) return f;
      const base = f.etiquetas || ENGINE.etiquetasPorDefectoDeArea(f.id, f.posiciones);
      return { ...f, etiquetas: { ...base, [String(posLocal)]: texto } };
    }));
    marca();
  }
  function anadirUnidad() {
    const n = filas.filter(f => !f.esCatalogo).length + 1;
    const id = "CUSTOM_" + Date.now();
    setFilas(prev => [...prev, {
      id, nombre: "Nueva unidad " + n, rotacion: "UCI A", posiciones: 2,
      activa: true, etiquetas: null, esCatalogo: false,
    }]);
    marca();
  }
  function borrarUnidad(id) {
    setFilas(prev => prev.filter(f => f.id !== id));
    if (abiertaEtiquetas === id) setAbiertaEtiquetas(null);
    marca();
  }
  async function guardar() {
    setBusy(true); setMsg("");
    try {
      const config = filas.map(f => ({
        id: f.id, nombre: (f.nombre || "").trim() || f.id, rotacion: f.rotacion,
        posiciones: f.posiciones, activa: f.activa,
        etiquetas: f.etiquetas || null,
      }));
      await saveEstructura(config);
      setDirty(false); setMsg("ok");
      setTimeout(() => setMsg(""), 2600);
    } catch (err) {
      setMsg(err.message || "No se pudo guardar.");
    } finally { setBusy(false); }
  }
  function restaurar() {
    const def = porDefecto;
    setFilas(catalogo.map(cat => {
      const d = def.find(a => a.id === cat.id);
      return { ...cat, posiciones: d ? d.posiciones : cat.posDefecto, activa: d ? d.activa : false, etiquetas: null, esCatalogo: true };
    }));
    setAbiertaEtiquetas(null);
    setDirty(true); setMsg("");
  }

  const totalPos = filas.filter(f => f.activa).reduce((s, f) => s + (f.posiciones || 0), 0);
  const rotacionesDisponibles = ["UCI A", "UCI B", "SEMIS", "REA", "UCCRI"];

  return (
    <div className="estructura">
      <div className="editor-head">
        <div>
          <h2>Estructura de la unidad</h2>
          <p className="section-intro">
            Marca qué unidades están abiertas, ajusta sus posiciones, cambia su nombre si lo
            necesitas, y edita la etiqueta de cada posición (por ejemplo «+ revisió carro»).
            También puedes crear unidades nuevas. Todo se guarda para las próximas generaciones.
          </p>
        </div>
        <button className="linkbtn" onClick={onCerrar}>← Volver</button>
      </div>

      <div className="est-card">
        <div className="est-row est-head-row">
          <span className="est-c-check">Abierta</span>
          <span className="est-c-nombre">Unidad (nombre editable)</span>
          <span className="est-c-rot">Cuenta como</span>
          <span className="est-c-pos">Posiciones</span>
          <span className="est-c-acciones"></span>
        </div>
        {filas.map(f => (
          <React.Fragment key={f.id}>
            <div className={"est-row" + (f.activa ? "" : " inactiva")}>
              <span className="est-c-check">
                <button className={"chk " + (f.activa ? "on" : "")}
                  onClick={() => setActiva(f.id, !f.activa)}
                  aria-label={f.activa ? "Cerrar unidad" : "Abrir unidad"}>
                  {f.activa ? "✓" : ""}
                </button>
              </span>
              <span className="est-c-nombre">
                <input className="est-nombre-input" value={f.nombre}
                  onChange={e => setCampo(f.id, "nombre", e.target.value)} />
              </span>
              <span className="est-c-rot">
                {f.esCatalogo
                  ? <span className="rot-tag">{f.rotacion}</span>
                  : <select className="rot-select" value={f.rotacion}
                      onChange={e => setCampo(f.id, "rotacion", e.target.value)}>
                      {rotacionesDisponibles.map(r => <option key={r}>{r}</option>)}
                    </select>}
              </span>
              <span className="est-c-pos">
                <input type="number" min="0" max="12" value={f.posiciones}
                  disabled={!f.activa}
                  onChange={e => setPos(f.id, e.target.value)} />
              </span>
              <span className="est-c-acciones">
                <button className="minibtn" disabled={!f.activa || f.posiciones === 0}
                  onClick={() => setAbiertaEtiquetas(abiertaEtiquetas === f.id ? null : f.id)}>
                  {abiertaEtiquetas === f.id ? "Ocultar etiquetas" : "Etiquetas…"}
                </button>
                {!f.esCatalogo && (
                  <button className="minibtn peligro" onClick={() => borrarUnidad(f.id)} title="Eliminar esta unidad">✕</button>
                )}
              </span>
            </div>
            {abiertaEtiquetas === f.id && f.activa && (
              <div className="est-etiquetas">
                <div className="est-et-hint">
                  Texto que acompaña al número de cada posición (deja vacío si no lleva).
                  Si quieres que el texto <strong>sustituya</strong> al número, empieza por «=»
                  (p. ej. <code>=ATURADA+ICTUS+CARRO</code>).
                </div>
                {Array.from({ length: f.posiciones }, (_, i) => i + 1).map(pl => {
                  const base = f.etiquetas || ENGINE.etiquetasPorDefectoDeArea(f.id, f.posiciones);
                  return (
                    <div key={pl} className="est-et-row">
                      <span className="est-et-num">Posición {pl}</span>
                      <input className="est-et-input" value={base[String(pl)] || ""}
                        placeholder="(sin etiqueta)"
                        onChange={e => setEtiqueta(f.id, pl, e.target.value)} />
                    </div>
                  );
                })}
              </div>
            )}
          </React.Fragment>
        ))}
        <div className="est-total">
          <span>Total de posiciones abiertas</span>
          <strong>{totalPos}</strong>
        </div>
      </div>

      <div className="est-add">
        <button className="minibtn" onClick={anadirUnidad}>+ Añadir unidad nueva</button>
      </div>

      {msg && msg !== "ok" && <div className="login-error" style={{ marginTop: 14 }}>{msg}</div>}
      {msg === "ok" && <div className="saved-note" style={{ marginTop: 14 }}>Estructura guardada. Vuelve a generar el cuadrante para aplicarla.</div>}
      {dirty && msg !== "ok" && <div className="dirty-note" style={{ marginTop: 14 }}>Tienes cambios sin guardar.</div>}

      <div className="est-actions">
        <button className="minibtn" onClick={restaurar}>Restaurar por defecto</button>
        <button className="savebtn" disabled={busy || !dirty} onClick={guardar}>
          {busy ? "Guardando…" : "Guardar estructura"}
        </button>
      </div>

      <p className="foot-hint" style={{ marginTop: 18 }}>
        «Cuenta como» define el grupo de rotación y qué criterios pueden acceder: por ejemplo,
        UCCRI Esquerra y Dreta cuentan ambas como «UCCRI». Las unidades nuevas usan el grupo
        que elijas; su personal se rige por las mismas normas de acceso de ese grupo.
      </p>
    </div>
  );
}

/* ---------- Editor de plantilla (apartado 5) ----------
   Edita la base de profesionales que alimenta el motor. Trabaja sobre
   una copia local; al guardar, persiste y queda disponible para las
   siguientes planillas (el motor relee estos datos en cada generación).
   El registro tiene la misma forma que produce el lector de Excel:
   { nombre, turno, apodo, criterio, suplente, antiguedad }
*/
const OPC_TURNO = ["TM", "TT", "TN", "12H Día", "12H Noche"];
const OPC_CRITERIO = ["SI", "NO", "UCI", "UCCRI"];
const OPC_ANTIG = ["< 3 ANY", "> 3 ANY < 10 ANY", "10 - 20 ANY", "> 20 ANY"];

function filaVacia() {
  return { nombre: "", turno: "TM", apodo: "", criterio: "SI", suplente: "No",
           antiguedad: "< 3 ANY", fondo: ENGINE.fondoPorDefectoDeTurno("TM"), antiguo: "No" };
}
// Normaliza un registro venido de Excel a los valores de los desplegables
function normalizarFila(p) {
  const turno = OPC_TURNO.find(t => t.toLowerCase() === String(p.turno || "").toLowerCase().trim()) || String(p.turno || "TM");
  let crit = String(p.criterio || "").toUpperCase().trim();
  if (crit === "SÍ") crit = "SI";
  if (!OPC_CRITERIO.includes(crit)) crit = "SI";
  const sup = String(p.suplente || "").toUpperCase().trim() === "SI" ? "Si" : "No";
  const antig = OPC_ANTIG.find(a => a.toUpperCase() === String(p.antiguedad || "").toUpperCase().trim()) || String(p.antiguedad || "< 3 ANY");
  // Fondo: el guardado si es válido; si no, el que toca por el turno.
  const fondo = ENGINE.normalizarFondo(p.fondo) || ENGINE.fondoPorDefectoDeTurno(turno);
  const ant = String(p.antiguo || "").toUpperCase().trim() === "SI" ? "Si" : "No";
  return { nombre: p.nombre || "", turno, apodo: p.apodo || "", criterio: crit,
           suplente: sup, antiguedad: antig, fondo, antiguo: ant };
}

function EditorPlantilla({ profs, saveProfs, onCerrar, hayResultado, onCargarExcel }) {
  const [rows, setRows] = useState(() => (profs || []).map(normalizarFila));
  const [q, setQ] = useState("");
  const [dirty, setDirty] = useState(false);
  const [savedFlash, setSavedFlash] = useState(false);
  const [orden, setOrden] = useState("nombre"); // nombre | turno | criterio

  // Si cambia la base externamente (p.ej. nueva carga de Excel) y no hay
  // cambios sin guardar, reflejarla.
  useEffect(() => {
    if (!dirty) setRows((profs || []).map(normalizarFila));
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [profs]);

  const total = rows.length;

  function setCampo(globalIdx, campo, valor) {
    setRows(prev => {
      const next = prev.slice();
      const fila = { ...next[globalIdx], [campo]: valor };
      // Al cambiar el turno, el fondo se preconfigura solo:
      // TM/TT/TN -> Blanco ; 12H Día / 12H Noche -> Lila.
      if (campo === "turno") fila.fondo = ENGINE.fondoPorDefectoDeTurno(valor);
      next[globalIdx] = fila;
      return next;
    });
    setDirty(true); setSavedFlash(false);
  }
  function eliminar(globalIdx) {
    setRows(prev => prev.filter((_, i) => i !== globalIdx));
    setDirty(true); setSavedFlash(false);
  }
  function añadir() {
    setRows(prev => [filaVacia(), ...prev]);
    setDirty(true); setSavedFlash(false);
    setQ("");
  }
  async function guardar() {
    // Descarta filas totalmente vacías (sin nombre)
    const limpio = rows.filter(r => r.nombre && r.nombre.trim());
    await saveProfs(limpio, "editada en la app · " + limpio.length + " profesionales");
    setRows(limpio.map(normalizarFila));
    setDirty(false); setSavedFlash(true);
    setTimeout(() => setSavedFlash(false), 2600);
  }
  function descartar() {
    setRows((profs || []).map(normalizarFila));
    setDirty(false);
  }

  // Índices visibles según búsqueda + orden, conservando el índice real
  const visibles = useMemo(() => {
    let idx = rows.map((r, i) => ({ r, i }));
    if (q.trim()) {
      const t = q.toLowerCase();
      idx = idx.filter(({ r }) =>
        (r.nombre || "").toLowerCase().includes(t) || (r.apodo || "").toLowerCase().includes(t));
    }
    const cmp = {
      nombre: (a, b) => (a.r.nombre || "").localeCompare(b.r.nombre || "", "es"),
      turno: (a, b) => (a.r.turno || "").localeCompare(b.r.turno || "") || (a.r.nombre || "").localeCompare(b.r.nombre || "", "es"),
      criterio: (a, b) => (a.r.criterio || "").localeCompare(b.r.criterio || "") || (a.r.nombre || "").localeCompare(b.r.nombre || "", "es"),
    }[orden];
    return idx.slice().sort(cmp);
  }, [rows, q, orden]);

  return (
    <div className="editor">
      <div className="editor-head">
        <div>
          <h2>Profesionales en plantilla</h2>
          <p className="section-intro">
            Esta es la base que usa el motor para asignar posiciones. Lo que cambies aquí se guarda
            en la base compartida y se aplica a las próximas planillas que generes.
          </p>
          <p className="section-intro">
            La columna <strong>Fondo</strong> es el color de las celdas de esa persona en el
            cuadrante. Se rellena sola según el turno (TM, TT y TN → blanco; 12 H Día y
            12 H Noche → lila), y puedes cambiarla a mano cuando haga falta. El punto ·
            indica que está cambiada respecto a lo que le tocaría por su turno.
          </p>
        </div>
        {hayResultado && (
          <button className="linkbtn" onClick={onCerrar}>← Volver al cuadrante</button>
        )}
      </div>

      <div className="editor-bar">
        <input className="search" placeholder="Buscar por nombre o apodo…"
          value={q} onChange={e => setQ(e.target.value)} />
        <label className="sortpick">
          <span>Ordenar</span>
          <select value={orden} onChange={e => setOrden(e.target.value)}>
            <option value="nombre">Nombre</option>
            <option value="turno">Turno</option>
            <option value="criterio">Criterio</option>
          </select>
        </label>
        <span className="count">{visibles.length} de {total}</span>
        <div className="editor-actions">
          <button className="minibtn" onClick={añadir}>+ Añadir profesional</button>
          {dirty && <button className="minibtn" onClick={descartar}>Descartar cambios</button>}
          <button className="savebtn" disabled={!dirty} onClick={guardar}>
            {savedFlash ? "✓ Guardado" : dirty ? "Guardar cambios" : "Sin cambios"}
          </button>
        </div>
      </div>

      {dirty && (
        <div className="dirty-note">Tienes cambios sin guardar. Pulsa «Guardar cambios» para aplicarlos a las siguientes planillas.</div>
      )}
      {savedFlash && !dirty && (
        <div className="saved-note">Plantilla guardada. Vuelve a generar el cuadrante para ver los cambios aplicados.</div>
      )}

      <p className="scroll-hint">Desliza la tabla hacia los lados para ver todas las columnas →</p>
      <div className="editor-tablewrap">
        <table className="edit-table">
          <thead>
            <tr>
              <th className="c-nombre">Nombre completo</th>
              <th className="c-apodo">Apodo</th>
              <th className="c-turno">Turno habitual</th>
              <th className="c-fondo">Fondo</th>
              <th className="c-crit">Criterio UCI</th>
              <th className="c-antig">Antigüedad</th>
              <th className="c-sup">Suplente</th>
              <th className="c-ant">Antiguo/a</th>
              <th className="c-del"></th>
            </tr>
          </thead>
          <tbody>
            {visibles.length === 0 && (
              <tr><td colSpan="9" className="edit-empty">
                {q ? "Ningún profesional coincide con la búsqueda." : "No hay profesionales. Añade uno o carga un Excel."}
              </td></tr>
            )}
            {visibles.map(({ r, i }) => (
              <tr key={i}>
                <td>
                  <input className="cell-input" value={r.nombre}
                    placeholder="Apellidos, Nombre"
                    onChange={e => setCampo(i, "nombre", e.target.value)} />
                </td>
                <td>
                  <input className="cell-input" value={r.apodo}
                    placeholder="—"
                    onChange={e => setCampo(i, "apodo", e.target.value)} />
                </td>
                <td>
                  <select className="cell-select" value={r.turno}
                    onChange={e => setCampo(i, "turno", e.target.value)}>
                    {OPC_TURNO.map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                </td>
                <td>
                  <span className="fondo-cell">
                    <span className={"fondo-muestra " + (r.fondo === ENGINE.FONDO.LILA ? "lila" : "blanco")}
                      aria-hidden="true" />
                    <select className="cell-select fondo-sel" value={r.fondo}
                      title={r.fondo === ENGINE.fondoPorDefectoDeTurno(r.turno)
                        ? "Color por defecto según el turno"
                        : "Color cambiado a mano (no es el que toca por el turno)"}
                      onChange={e => setCampo(i, "fondo", e.target.value)}>
                      <option value={ENGINE.FONDO.BLANCO}>Blanco</option>
                      <option value={ENGINE.FONDO.LILA}>Lila</option>
                    </select>
                    {r.fondo !== ENGINE.fondoPorDefectoDeTurno(r.turno) && (
                      <span className="fondo-manual" title="Cambiado a mano">•</span>
                    )}
                  </span>
                </td>
                <td>
                  <select className={"cell-select crit-sel crit-" + r.criterio} value={r.criterio}
                    onChange={e => setCampo(i, "criterio", e.target.value)}>
                    {OPC_CRITERIO.map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                </td>
                <td>
                  <select className="cell-select" value={r.antiguedad}
                    onChange={e => setCampo(i, "antiguedad", e.target.value)}>
                    {OPC_ANTIG.map(o => <option key={o} value={o}>{o}</option>)}
                  </select>
                </td>
                <td className="td-center">
                  <button className={"toggle " + (r.suplente === "Si" ? "on" : "")}
                    onClick={() => setCampo(i, "suplente", r.suplente === "Si" ? "No" : "Si")}
                    title="Marcar como suplente">
                    {r.suplente === "Si" ? "Sí" : "No"}
                  </button>
                </td>
                <td className="td-center">
                  <button className={"toggle ant " + (r.antiguo === "Si" ? "on" : "")}
                    onClick={() => setCampo(i, "antiguo", r.antiguo === "Si" ? "No" : "Si")}
                    title="Antiguo/a: posición fija toda la semana y misma unidad todo el mes">
                    {r.antiguo === "Si" ? "Sí" : "No"}
                  </button>
                </td>
                <td className="td-center">
                  <button className="delbtn" onClick={() => eliminar(i)} title="Eliminar profesional">×</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      <div className="editor-foot">
        <button className="minibtn ghost" onClick={onCargarExcel}>Sustituir todo desde un Excel</button>
        <span className="foot-hint">El criterio decide a qué áreas puede ir cada profesional; la antigüedad y la suplencia, su prioridad para ocupar posición fija.</span>
      </div>
    </div>
  );
}

/* ---------- Plantilla ---------- */
function Profesionales({ db }) {
  const [q, setQ] = useState("");
  const list = db.list.filter(p => !q || p.nombre.toLowerCase().includes(q.toLowerCase()) || (p.apodo || "").toLowerCase().includes(q.toLowerCase()));
  return (
    <div className="plantilla">
      <div className="plantilla-bar">
        <input className="search" placeholder="Buscar por nombre o apodo…" value={q} onChange={e => setQ(e.target.value)} />
        <span className="count">{list.length} de {db.list.length}</span>
      </div>
      <table className="inf-table">
        <thead><tr><th>Nombre</th><th>Apodo</th><th>Turno</th><th>Criterio</th><th>Antig.</th><th>Suplente</th></tr></thead>
        <tbody>
          {list.map((p, i) => (
            <tr key={i}>
              <td>{p.nombre}</td><td>{p.apodo}</td><td>{p.turnoHabitual}</td>
              <td><span className={"crit crit-" + p.criterio}>{p.criterio}</span></td>
              <td>{p.textoAntiguedad}</td>
              <td>{p.esSuplente ? "Sí" : "—"}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/* ============================================================ */
const CSS = `
:root {
  --ink: #1a2230;
  --ink-soft: #404a5c;
  --muted: #6b7686;
  --line: #e3e7ee;
  --line-strong: #cfd5df;
  --paper: #fbfcfd;
  --surface: #ffffff;
  --panel: #f5f7fa;
  --accent: #0f7c74;
  --accent-deep: #0a5d57;
  --accent-wash: #e6f3f1;
  --empty: #f0c2c2;
  --empty-strong: #d96b6b;
  --danger: #b3433a;
  --shadow: 0 1px 2px rgba(20,30,48,.06), 0 4px 16px rgba(20,30,48,.05);
  --radius: 10px;
  --header-h: 60px;
  --font: "Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
}

* { box-sizing: border-box; }
.app {
  font-family: var(--font);
  color: var(--ink);
  background: var(--paper);
  min-height: 100vh;
  font-size: 14px;
  -webkit-font-smoothing: antialiased;
}

/* ---------- Header ---------- */
.header {
  height: var(--header-h);
  display: flex; align-items: center;
  padding: 0 22px;
  background: var(--surface);
  border-bottom: 1px solid var(--line);
  position: sticky; top: 0; z-index: 20;
}
.brand { display: flex; align-items: center; gap: 13px; }
.brand-mark {
  width: 40px; height: 40px; border-radius: 10px;
  display: grid; place-items: center;
  background: var(--accent-wash); color: var(--accent);
}
.brand-title { font-weight: 680; font-size: 16px; letter-spacing: -.01em; line-height: 1.1; }
.brand-sub { font-size: 11.5px; color: var(--muted); margin-top: 2px; }

/* ---------- Layout ---------- */
.layout { display: grid; grid-template-columns: 308px 1fr; align-items: start; }
.sidebar {
  position: sticky; top: var(--header-h);
  height: calc(100vh - var(--header-h));
  overflow-y: auto;
  background: var(--surface);
  border-right: 1px solid var(--line);
  padding: 20px 18px 40px;
  display: flex; flex-direction: column; gap: 20px;
}
.main { padding: 24px 28px 60px; min-width: 0; }

/* ---------- Sidebar blocks ---------- */
.block { display: flex; flex-direction: column; gap: 10px; }
.block-head { display: flex; align-items: center; gap: 9px; }
.block-head h3 { margin: 0; font-size: 12.5px; font-weight: 650; letter-spacing: .02em; text-transform: uppercase; color: var(--ink-soft); }
.step {
  width: 19px; height: 19px; flex: none; border-radius: 50%;
  background: var(--ink); color: #fff;
  font-size: 11px; font-weight: 650;
  display: grid; place-items: center;
}
.hint { margin: 0; font-size: 12px; line-height: 1.5; color: var(--muted); }
.hint.center { text-align: center; }

.filebtn {
  width: 100%; padding: 10px 12px;
  background: var(--ink); color: #fff;
  border: none; border-radius: 8px;
  font: inherit; font-weight: 580; font-size: 13px;
  cursor: pointer; transition: background .15s;
}
.filebtn:hover { background: #2a3548; }
.filebtn.ghost { background: var(--surface); color: var(--ink); border: 1px solid var(--line-strong); }
.filebtn.ghost:hover { background: var(--panel); }

.filemeta { display: flex; flex-direction: column; gap: 2px; padding: 9px 11px; background: var(--panel); border-radius: 7px; border: 1px solid var(--line); }
.filemeta strong { font-size: 12.5px; font-weight: 620; }
.filemeta span { font-size: 11px; color: var(--muted); }

.field { display: flex; flex-direction: column; gap: 5px; }
.field span { font-size: 11.5px; color: var(--ink-soft); font-weight: 540; }
.field input {
  padding: 8px 10px; font: inherit; font-size: 13px;
  border: 1px solid var(--line-strong); border-radius: 7px; background: var(--surface);
  color: var(--ink);
}
.field input:focus { outline: 2px solid var(--accent-wash); border-color: var(--accent); }

.generate {
  width: 100%; padding: 13px; margin-top: 2px;
  background: var(--accent); color: #fff;
  border: none; border-radius: 9px;
  font: inherit; font-weight: 650; font-size: 14.5px;
  cursor: pointer; transition: background .15s, transform .05s;
  box-shadow: 0 1px 0 var(--accent-deep);
}
.generate:hover:not(:disabled) { background: var(--accent-deep); }
.generate:active:not(:disabled) { transform: translateY(1px); }
.generate:disabled { background: var(--line-strong); color: #fff; cursor: not-allowed; box-shadow: none; }

.row-btns { display: flex; gap: 7px; }
.minibtn {
  flex: 1; padding: 7px 8px; font: inherit; font-size: 11.5px; font-weight: 560;
  background: var(--surface); border: 1px solid var(--line-strong); border-radius: 7px; cursor: pointer; color: var(--ink-soft);
}
.minibtn:hover:not(:disabled) { background: var(--panel); }
.minibtn:disabled { opacity: .45; cursor: not-allowed; }
.minibtn.danger { color: var(--danger); border-color: #e8cdca; }
.minibtn.danger:hover:not(:disabled) { background: #fbf1f0; }

/* ---------- Welcome ---------- */
.welcome { max-width: 560px; padding: 28px 4px; }
.welcome h2 { margin: 0 0 14px; font-size: 24px; font-weight: 680; letter-spacing: -.02em; }
.welcome p { margin: 0 0 18px; line-height: 1.62; color: var(--ink-soft); font-size: 14.5px; }
.checklist { display: flex; flex-direction: column; gap: 9px; margin: 22px 0; }
.check { display: flex; align-items: center; gap: 11px; font-size: 14px; color: var(--muted); }
.check.on { color: var(--ink); font-weight: 540; }
.check .dot {
  width: 22px; height: 22px; border-radius: 50%; flex: none;
  border: 1.5px solid var(--line-strong); display: grid; place-items: center;
  font-size: 12px; color: #fff;
}
.check.on .dot { background: var(--accent); border-color: var(--accent); }

.error { padding: 11px 14px; background: #fbf0ef; border: 1px solid #e8cdca; color: var(--danger); border-radius: 8px; margin-bottom: 16px; font-size: 13px; }

/* ---------- Tabs ---------- */
.tabsbar { display: flex; flex-wrap: wrap; align-items: center; justify-content: space-between; gap: 14px; margin-bottom: 20px; }
.tabs { display: flex; gap: 2px; background: var(--panel); padding: 4px; border-radius: 10px; }
.tab {
  padding: 8px 15px; border: none; background: transparent; cursor: pointer;
  font: inherit; font-size: 13px; font-weight: 560; color: var(--muted); border-radius: 7px; transition: all .12s;
}
.tab:hover { color: var(--ink); }
.tab.active { background: var(--surface); color: var(--ink); box-shadow: var(--shadow); font-weight: 620; }

.weekpick { display: flex; gap: 6px; }
.weekbtn {
  display: flex; flex-direction: column; align-items: center; gap: 1px;
  padding: 6px 13px; border: 1px solid var(--line-strong); background: var(--surface);
  border-radius: 8px; cursor: pointer; font: inherit; font-size: 12.5px; font-weight: 600; color: var(--ink-soft);
  line-height: 1.2;
}
.weekbtn small { font-size: 10px; font-weight: 500; color: var(--muted); }
.weekbtn.active { border-color: var(--accent); background: var(--accent-wash); color: var(--accent-deep); }
.weekbtn.active small { color: var(--accent); }

/* ---------- Cronograma ---------- */
.crono-wrap { display: flex; flex-direction: column; gap: 26px; }
.crono-legend { display: flex; flex-wrap: wrap; gap: 16px; padding: 12px 16px; background: var(--surface); border: 1px solid var(--line); border-radius: var(--radius); }
.legend-item { display: flex; align-items: center; gap: 7px; font-size: 12px; color: var(--ink-soft); }
.swatch { width: 14px; height: 14px; border-radius: 4px; flex: none; }

.bloque { background: var(--surface); border: 1px solid var(--line); border-radius: var(--radius); overflow: hidden; box-shadow: var(--shadow); }
.bloque-title {
  display: flex; align-items: baseline; gap: 9px;
  padding: 11px 16px; font-weight: 660; font-size: 14px; letter-spacing: .02em;
  color: #fff;
}
.bloque-title small { font-size: 11px; font-weight: 500; opacity: .85; letter-spacing: .06em; }
.bloque-title[data-suf="MATI"]  { background: linear-gradient(90deg, #6a5aa8, #7d6cba); }
.bloque-title[data-suf="TARDE"] { background: linear-gradient(90deg, #c4764a, #d4895a); }
.bloque-title[data-suf="NIT"]   { background: linear-gradient(90deg, #3a3463, #4a4477); }

.grid { overflow-x: auto; }
.cuadro { border-collapse: collapse; width: 100%; min-width: 720px; }
.cuadro th, .cuadro td { border: 1px solid var(--line); }
.cuadro thead th { background: var(--panel); padding: 7px 6px; font-size: 11px; }
.area-h, .num-h { background: var(--panel); }
.day-h { display: table-cell; }
.day-h .dow { display: block; font-weight: 680; font-size: 11px; color: var(--ink); letter-spacing: .03em; }
.day-h .dnum { display: block; font-size: 10.5px; font-weight: 500; color: var(--muted); margin-top: 1px; font-variant-numeric: tabular-nums; }

.area-cell { background: var(--panel); padding: 6px 9px; vertical-align: middle; width: 86px; }
.area-name { display: block; font-weight: 660; font-size: 12px; color: var(--ink); }
.area-sub { display: block; font-size: 10px; color: var(--muted); margin-top: 1px; }
.num-cell { background: var(--panel); text-align: center; width: 26px; font-size: 11px; font-weight: 600; color: var(--muted); font-variant-numeric: tabular-nums; }
.num-cell.num-label { width: 96px; padding: 2px 5px; line-height: 1.15; }
.pos-fija { font-size: 8.5px; font-weight: 700; color: var(--ink-soft); text-transform: uppercase; letter-spacing: .01em; display: block; }
.pos-fija-sub { display: block; font-size: 8px; font-weight: 600; color: #b5853a; margin-top: 1px; }
.sep-row td { height: 10px; background: var(--paper); border: none; padding: 0; }
.exc-row .exc-num { background: #f3eef8; color: #8a6db8; font-weight: 700; }
.exc-cell { background: #faf7fd; font-style: italic; color: var(--ink-soft); }
.exc-cell:empty { background: var(--paper); }
.sep-cell { border: none !important; }

.data-cell {
  padding: 5px 8px; font-size: 12px; font-weight: 530;
  text-align: center; white-space: nowrap; color: var(--ink);
  max-width: 120px; overflow: hidden; text-overflow: ellipsis;
  position: relative;
}
.badge {
  display: inline-block; margin-left: 4px; padding: 0 4px;
  font-size: 8.5px; font-weight: 700; line-height: 1.5; border-radius: 4px;
  vertical-align: middle; letter-spacing: .02em;
}
.badge-dbl { background: #d4763a; color: #fff; }
.badge-plus { background: #3a7a6a; color: #fff; }
.badge-half { background: #888; color: #fff; }

/* ---------- Anotaciones manuales (B3) ---------- */
.crono-toolbar { display: flex; align-items: flex-start; justify-content: space-between; gap: 16px; flex-wrap: wrap; margin-bottom: 12px; }
.anotar-btn { flex: none; padding: 8px 14px; border: 1px solid var(--line-strong); border-radius: 8px; background: var(--surface); color: var(--ink-soft); font: inherit; font-size: 12.5px; font-weight: 560; cursor: pointer; white-space: nowrap; }
.anotar-btn:hover { background: var(--panel); }
.anotar-btn.on { background: var(--accent); border-color: var(--accent); color: #fff; }
.anotar-help { background: var(--accent-wash); border: 1px solid #bfe0db; border-radius: 8px; padding: 10px 13px; font-size: 12.5px; line-height: 1.55; color: var(--ink-soft); margin-bottom: 14px; }
.anotar-help strong { color: var(--accent-deep); }
.data-cell.anotable { cursor: pointer; }
.data-cell.anotable:hover { outline: 2px solid var(--accent); outline-offset: -2px; }
.data-cell { /* permitir que el editor de nota se vea fuera de la celda */ }
td.data-cell:has(.nota-edit) { overflow: visible; }
.nota-cell { display: block; font-size: 9px; font-weight: 700; color: #b5532a; margin-top: 1px; text-transform: uppercase; letter-spacing: .02em; }
.data-cell.con-nota { box-shadow: inset 0 -2px 0 #e0a890; }
.nota-edit { position: absolute; z-index: 30; left: 50%; top: 100%; transform: translateX(-50%); margin-top: 2px; background: var(--surface); border: 1px solid var(--accent); border-radius: 8px; box-shadow: 0 6px 20px rgba(20,30,48,.2); padding: 7px; display: flex; flex-direction: column; gap: 6px; width: 130px; }
.nota-edit input { width: 100%; padding: 6px 8px; border: 1px solid var(--line-strong); border-radius: 6px; font: inherit; font-size: 12px; }
.nota-edit input:focus { outline: 2px solid var(--accent-wash); border-color: var(--accent); }
.nota-edit-btns { display: flex; gap: 5px; }
.nota-ok { flex: 1; padding: 5px; border: none; border-radius: 6px; background: var(--accent); color: #fff; font: inherit; font-size: 11.5px; font-weight: 600; cursor: pointer; }
.nota-ok:hover { background: var(--accent-deep); }
.nota-x { padding: 5px 8px; border: 1px solid var(--line-strong); border-radius: 6px; background: var(--surface); color: var(--muted); font: inherit; font-size: 11.5px; cursor: pointer; }
.nota-x:hover { background: var(--panel); color: var(--ink); }
.area-start td { border-top: 2px solid var(--line-strong); }

/* ---------- Decalajes ---------- */
.section-intro { color: var(--muted); font-size: 13px; margin: 0 0 18px; line-height: 1.5; }
.decalajes { display: flex; flex-direction: column; gap: 20px; }
.dec-day h4 { margin: 0 0 8px; font-size: 13px; font-weight: 650; color: var(--ink-soft); }
.dec-table, .inf-table {
  width: 100%; border-collapse: collapse; background: var(--surface);
  border: 1px solid var(--line); border-radius: var(--radius); overflow: hidden;
}
.dec-table th, .inf-table th { text-align: left; padding: 9px 13px; background: var(--panel); font-size: 11.5px; font-weight: 620; color: var(--ink-soft); letter-spacing: .02em; border-bottom: 1px solid var(--line); }
.dec-table td, .inf-table td { padding: 9px 13px; font-size: 13px; border-bottom: 1px solid var(--line); }
.dec-table tr:last-child td, .inf-table tr:last-child td { border-bottom: none; }
.dec-table tr.ok td:last-child { color: var(--accent-deep); font-weight: 560; }
.dec-table tr.ko { background: #fcf3f2; }
.dec-table tr.ko td:last-child { color: var(--danger); font-weight: 560; }

/* ---------- Vinculación de profesionales ---------- */
.vincular { max-width: 760px; }
.vinc-list { display: flex; flex-direction: column; gap: 14px; margin-bottom: 20px; }
.vinc-card { border: 1px solid var(--line); border-radius: var(--radius); background: var(--surface); padding: 16px; box-shadow: var(--shadow); }
.vinc-head { display: flex; align-items: baseline; justify-content: space-between; gap: 12px; margin-bottom: 12px; flex-wrap: wrap; }
.vinc-name { font-size: 16px; font-weight: 640; color: var(--ink); }
.vinc-meta { font-size: 12px; color: var(--muted); }
.vinc-modos { display: flex; gap: 8px; margin-bottom: 12px; }
.vinc-tab { flex: 1; padding: 8px 12px; border: 1px solid var(--line-strong); border-radius: 8px; background: var(--surface); color: var(--ink-soft); font: inherit; font-size: 13px; font-weight: 540; cursor: pointer; }
.vinc-tab:hover { background: var(--panel); }
.vinc-tab.on { background: var(--accent); border-color: var(--accent); color: #fff; }
.vinc-body { padding-top: 4px; }
.vinc-lbl { display: block; font-size: 12.5px; color: var(--ink-soft); margin-bottom: 6px; font-weight: 540; }
.vinc-select, .vinc-form input, .vinc-form select { width: 100%; padding: 8px 10px; border: 1px solid var(--line-strong); border-radius: 8px; font: inherit; font-size: 13px; background: var(--surface); color: var(--ink); }
.vinc-select:focus, .vinc-form input:focus, .vinc-form select:focus { outline: 2px solid var(--accent-wash); border-color: var(--accent); }
.vinc-form { display: flex; flex-direction: column; gap: 10px; }
.vinc-row { display: flex; flex-direction: column; gap: 4px; }
.vinc-row label { font-size: 12px; color: var(--ink-soft); font-weight: 540; }
.vinc-row .req { color: #c0392b; margin-left: 2px; }
.vinc-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.vinc-footer { position: sticky; bottom: 0; display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 14px 0; background: linear-gradient(transparent, var(--paper) 30%); }
.vinc-status { font-size: 13px; color: var(--ink-soft); font-weight: 540; }
.vinc-actions { display: flex; gap: 10px; }

/* ---------- Importar rotación desde Excel ---------- */
.importar { max-width: 820px; }
.hist-head-acciones { display: flex; align-items: center; gap: 12px; }
.imp-file { display: flex; flex-direction: column; gap: 3px; background: var(--panel); border: 1px solid var(--line); border-radius: 10px; padding: 12px 14px; margin-bottom: 14px; }
.imp-file strong { font-size: 14px; color: var(--ink); }
.imp-file span { font-size: 12px; color: var(--muted); }
.imp-anio { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; margin-bottom: 16px; }
.imp-anio label { font-size: 13px; font-weight: 560; color: var(--ink-soft); }
.imp-anio input { width: 100px; padding: 7px 10px; border: 1px solid var(--line-strong); border-radius: 8px; font: inherit; font-size: 13px; text-align: center; background: var(--surface); color: var(--ink); }
.imp-anio input:focus { outline: 2px solid var(--accent-wash); border-color: var(--accent); }
.imp-hint { font-size: 12px; color: var(--muted); }
.imp-resumen { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 12px; }
.imp-dato { flex: 1; min-width: 150px; background: var(--surface); border: 1px solid var(--line); border-radius: var(--radius); padding: 14px 16px; box-shadow: var(--shadow); }
.imp-dato strong { display: block; font-size: 22px; color: var(--accent-deep); line-height: 1.2; }
.imp-dato span { font-size: 12px; color: var(--ink-soft); }
.imp-meses { font-size: 13px; color: var(--ink-soft); background: var(--accent-wash); border-radius: 8px; padding: 10px 13px; margin-bottom: 18px; line-height: 1.5; }
.imp-bloque { border: 1px solid var(--line); border-radius: var(--radius); background: var(--surface); padding: 16px; margin-bottom: 16px; box-shadow: var(--shadow); }
.imp-bloque h3 { font-size: 14.5px; margin: 0 0 4px; color: var(--ink); }
.imp-sub { font-size: 12.5px; color: var(--ink-soft); margin: 0 0 12px; line-height: 1.45; }
.imp-vinc { display: grid; grid-template-columns: 200px 1fr 90px; gap: 10px; align-items: center; padding: 9px 0; border-top: 1px solid var(--line); }
.imp-vinc.omitido { opacity: .55; }
.imp-vinc-nom strong { font-size: 13.5px; color: var(--ink); display: block; }
.imp-vinc-meta { font-size: 11.5px; color: var(--muted); }
.imp-omitido-txt { font-size: 12.5px; color: var(--muted); font-style: italic; }
.imp-raras { margin: 0; padding-left: 18px; font-size: 12.5px; color: var(--ink-soft); line-height: 1.6; }
@media (max-width: 700px) { .imp-vinc { grid-template-columns: 1fr; } }

/* ---------- Editor de reglas de rotación ---------- */
.reglas { max-width: 760px; }
.reglas-list { display: flex; flex-direction: column; gap: 10px; margin-bottom: 18px; }
.regla-card { display: flex; gap: 14px; align-items: flex-start; border: 1px solid var(--line); border-radius: var(--radius); background: var(--surface); padding: 14px 16px; box-shadow: var(--shadow); transition: border-color .15s; }
.regla-card.on { border-color: var(--accent); }
.regla-card.dep-off { opacity: .55; }
.regla-txt { flex: 1; }
.regla-titulo { font-size: 14.5px; font-weight: 640; color: var(--ink); margin-bottom: 3px; }
.regla-desc { font-size: 12.8px; color: var(--ink-soft); line-height: 1.45; }
.regla-aviso { font-size: 12px; color: #a2652a; margin-top: 5px; font-style: italic; }
.chk.grande { width: 26px; height: 26px; min-width: 26px; font-size: 15px; margin-top: 2px; }

/* ---------- Editor de estructura ampliado ---------- */
.est-nombre-input { width: 100%; padding: 6px 9px; border: 1px solid var(--line-strong); border-radius: 7px; font: inherit; font-size: 13px; background: var(--surface); color: var(--ink); }
.est-nombre-input:focus { outline: 2px solid var(--accent-wash); border-color: var(--accent); }
.rot-select { padding: 5px 8px; border: 1px solid var(--line-strong); border-radius: 7px; font: inherit; font-size: 12.5px; background: var(--surface); }
.est-c-acciones { display: flex; gap: 6px; align-items: center; justify-content: flex-end; min-width: 150px; }
.minibtn.peligro { color: #b3402f; border-color: #e0b7af; }
.minibtn.peligro:hover { background: #fbeeeb; }
.est-etiquetas { grid-column: 1 / -1; background: var(--panel); border: 1px solid var(--line); border-radius: 10px; padding: 12px 14px; margin: 2px 0 10px; }
.est-et-hint { font-size: 12px; color: var(--ink-soft); margin-bottom: 10px; line-height: 1.5; }
.est-et-hint code { background: var(--surface); padding: 1px 5px; border-radius: 4px; border: 1px solid var(--line); }
.est-et-row { display: flex; align-items: center; gap: 10px; margin-bottom: 6px; }
.est-et-num { min-width: 88px; font-size: 12.5px; color: var(--ink-soft); font-weight: 540; }
.est-et-input { flex: 1; padding: 6px 9px; border: 1px solid var(--line-strong); border-radius: 7px; font: inherit; font-size: 12.5px; background: var(--surface); }
.est-add { margin-top: 12px; }

/* Buscador con autocompletado */
.buscador { position: relative; }
.buscador-input { width: 100%; padding: 9px 32px 9px 12px; border: 1px solid var(--line-strong); border-radius: 8px; font: inherit; font-size: 13px; background: var(--surface); color: var(--ink); }
.buscador-input:focus { outline: 2px solid var(--accent-wash); border-color: var(--accent); }
.buscador-input.elegido { border-color: var(--accent); background: var(--accent-wash); font-weight: 540; }
.buscador-ok { position: absolute; right: 11px; top: 50%; transform: translateY(-50%); color: var(--accent-deep); font-weight: 700; pointer-events: none; }
.buscador-lista { position: absolute; z-index: 40; top: calc(100% + 3px); left: 0; right: 0; max-height: 240px; overflow-y: auto; margin: 0; padding: 4px; list-style: none; background: var(--surface); border: 1px solid var(--line-strong); border-radius: 8px; box-shadow: 0 8px 24px rgba(20,30,48,.18); }
.buscador-lista li { display: flex; align-items: baseline; justify-content: space-between; gap: 10px; padding: 8px 10px; border-radius: 6px; cursor: pointer; font-size: 13px; }
.buscador-lista li.activo { background: var(--accent-wash); }
.buscador-nombre { color: var(--ink); font-weight: 520; }
.buscador-apodo { color: var(--muted); font-size: 11.5px; white-space: nowrap; }
.buscador-vacio { color: var(--muted); font-style: italic; cursor: default; padding: 8px 10px; }
.historico { max-width: 1100px; }
.hist-grid-wrap { overflow-x: auto; border: 1px solid var(--line); border-radius: var(--radius); background: var(--surface); box-shadow: var(--shadow); }
.hist-grid { border-collapse: collapse; width: 100%; }
.hist-grid th, .hist-grid td { border-bottom: 1px solid var(--line); border-right: 1px solid var(--line); }
.hist-grid th:last-child, .hist-grid td:last-child { border-right: none; }
.hist-prof-h { position: sticky; left: 0; z-index: 2; background: var(--panel); text-align: left; padding: 11px 14px; font-size: 11.5px; font-weight: 620; color: var(--ink-soft); text-transform: uppercase; letter-spacing: .02em; min-width: 180px; }
.hist-mes-h { background: var(--panel); padding: 11px 12px; font-size: 12px; font-weight: 620; color: var(--ink); white-space: nowrap; min-width: 130px; }
.hist-prof { position: sticky; left: 0; z-index: 1; background: var(--surface); padding: 10px 14px; font-size: 13px; font-weight: 540; white-space: nowrap; }
.hist-cell { padding: 8px 12px; vertical-align: middle; }
.hist-chips { display: flex; flex-wrap: wrap; gap: 4px; }
.uni-chip { display: inline-block; padding: 3px 10px; border-radius: 20px; font-size: 11.5px; font-weight: 620; border: 1px solid; white-space: nowrap; }
.hist-vacio { color: var(--line-strong); font-size: 16px; }
.hist-grid tbody tr:hover .hist-prof { background: var(--panel); }
.hist-grid tbody tr:last-child td { border-bottom: none; }

/* ---------- Informe / Plantilla ---------- */
.informe, .plantilla { max-width: 940px; }
.inf-bloque { margin-bottom: 26px; }
.inf-titulo { display: flex; align-items: center; gap: 10px; font-size: 14px; font-weight: 640; margin: 0 0 12px; color: var(--ink); }
.inf-pill { display: inline-block; padding: 3px 11px; border-radius: 20px; font-size: 11px; font-weight: 700; letter-spacing: .03em; color: #fff; }
.inf-pill.dia { background: #d9a441; }
.inf-pill.nit { background: #5b6bbf; }
.roto-tag { display: inline-block; margin-left: 8px; padding: 1px 7px; border-radius: 5px; font-size: 10px; font-weight: 650; background: var(--accent-wash); color: var(--accent-deep); vertical-align: middle; }
.empty-note.small { font-size: 12.5px; padding: 12px; }
.plantilla-bar { display: flex; align-items: center; gap: 14px; margin-bottom: 14px; }
.search { flex: 1; max-width: 360px; padding: 9px 13px; border: 1px solid var(--line-strong); border-radius: 8px; font: inherit; font-size: 13px; }
.search:focus { outline: 2px solid var(--accent-wash); border-color: var(--accent); }
.count { font-size: 12px; color: var(--muted); }
.crit { display: inline-block; padding: 2px 8px; border-radius: 20px; font-size: 11px; font-weight: 600; }
.crit-SI { background: #e6f3f1; color: var(--accent-deep); }
.crit-NO { background: #f0ede6; color: #7a6a3a; }
.crit-UCI { background: #e7ecf7; color: #3a4a8e; }
.crit-UCCRI { background: #f3e7f0; color: #8e3a75; }

.empty-note { padding: 40px; text-align: center; color: var(--muted); background: var(--surface); border: 1px dashed var(--line-strong); border-radius: var(--radius); }

/* ---------- Sidebar active state for editor button ---------- */
.filebtn.ghost.is-active { background: var(--accent-wash); border-color: var(--accent); color: var(--accent-deep); }

/* ---------- Editor de plantilla ---------- */
.editor { max-width: 1140px; }
.editor-head { display: flex; align-items: flex-start; justify-content: space-between; gap: 20px; margin-bottom: 18px; }
.editor-head h2 { margin: 0 0 6px; font-size: 22px; font-weight: 680; letter-spacing: -.02em; }
.editor-head .section-intro { margin: 0; max-width: 640px; }
.linkbtn { background: none; border: none; color: var(--accent); font: inherit; font-size: 13px; font-weight: 580; cursor: pointer; padding: 6px 2px; white-space: nowrap; }
.linkbtn:hover { color: var(--accent-deep); text-decoration: underline; }

.editor-bar { display: flex; align-items: center; gap: 14px; flex-wrap: wrap; margin-bottom: 14px; }
.sortpick { display: flex; align-items: center; gap: 7px; font-size: 12px; color: var(--ink-soft); }
.sortpick select { padding: 7px 9px; border: 1px solid var(--line-strong); border-radius: 7px; font: inherit; font-size: 12.5px; background: var(--surface); color: var(--ink); }
.editor-actions { display: flex; align-items: center; gap: 8px; margin-left: auto; flex-wrap: wrap; }
.minibtn.ghost { background: var(--surface); }
.savebtn {
  padding: 9px 16px; border: none; border-radius: 8px;
  background: var(--accent); color: #fff; font: inherit; font-weight: 620; font-size: 13px; cursor: pointer;
  box-shadow: 0 1px 0 var(--accent-deep); transition: background .15s;
}
.savebtn:hover:not(:disabled) { background: var(--accent-deep); }
.savebtn:disabled { background: var(--line-strong); box-shadow: none; cursor: default; }

.dirty-note { padding: 9px 14px; background: #fff7ec; border: 1px solid #f0dcb6; color: #8a6a2a; border-radius: 8px; font-size: 12.5px; margin-bottom: 14px; }
.saved-note { padding: 9px 14px; background: var(--accent-wash); border: 1px solid #bfe0db; color: var(--accent-deep); border-radius: 8px; font-size: 12.5px; margin-bottom: 14px; }

/* overflow-x: auto para poder desplazar la tabla de lado en móvil.
   -webkit-overflow-scrolling da inercia en iPhone. */
.editor-tablewrap { background: var(--surface); border: 1px solid var(--line); border-radius: var(--radius); overflow-x: auto; overflow-y: hidden; box-shadow: var(--shadow); -webkit-overflow-scrolling: touch; }
.edit-table { width: 100%; border-collapse: collapse; min-width: 900px; }
.edit-table thead th {
  text-align: left; padding: 10px 12px; background: var(--panel);
  font-size: 11.5px; font-weight: 620; color: var(--ink-soft); letter-spacing: .02em;
  border-bottom: 1px solid var(--line); position: sticky; top: 0; z-index: 1;
}
.edit-table td { padding: 5px 8px; border-bottom: 1px solid var(--line); vertical-align: middle; }
.edit-table tr:last-child td { border-bottom: none; }
.edit-table tbody tr:hover { background: #fafbfc; }
.c-nombre { width: 26%; } .c-apodo { width: 14%; } .c-turno { width: 12%; }
.c-fondo { width: 11%; }
.c-crit { width: 11%; } .c-antig { width: 14%; } .c-sup { width: 7%; } .c-ant { width: 8%; } .c-del { width: 36px; }
.toggle.ant.on { background: #efe6fa; border-color: #b79bd8; color: #5c3d84; }

/* Casilla de color de fondo del profesional */
.fondo-cell { display: flex; align-items: center; gap: 6px; }
.fondo-muestra { width: 14px; height: 14px; border-radius: 4px; border: 1px solid var(--line-strong); flex: none; }
.fondo-muestra.lila { background: #D2BEE6; }
.fondo-muestra.blanco { background: #FFFFFF; }
.fondo-sel { flex: 1; min-width: 0; }
.fondo-manual { color: var(--accent-deep); font-size: 16px; line-height: 1; cursor: default; }
.td-center { text-align: center; }
.edit-empty { text-align: center; color: var(--muted); padding: 32px; font-size: 13px; }

.cell-input {
  width: 100%; padding: 7px 9px; border: 1px solid transparent; border-radius: 6px;
  font: inherit; font-size: 13px; background: transparent; color: var(--ink);
}
.cell-input:hover { border-color: var(--line); }
.cell-input:focus { outline: none; border-color: var(--accent); background: #fff; box-shadow: 0 0 0 2px var(--accent-wash); }
.cell-input::placeholder { color: var(--line-strong); }

.cell-select {
  width: 100%; padding: 7px 8px; border: 1px solid var(--line); border-radius: 6px;
  font: inherit; font-size: 12.5px; background: var(--surface); color: var(--ink); cursor: pointer;
}
.cell-select:focus { outline: none; border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-wash); }
.crit-sel { font-weight: 600; }
.crit-sel.crit-SI { color: var(--accent-deep); }
.crit-sel.crit-NO { color: #7a6a3a; }
.crit-sel.crit-UCI { color: #3a4a8e; }
.crit-sel.crit-UCCRI { color: #8e3a75; }

.toggle {
  min-width: 42px; padding: 6px 10px; border-radius: 20px; cursor: pointer;
  font: inherit; font-size: 12px; font-weight: 600;
  background: var(--panel); border: 1px solid var(--line-strong); color: var(--muted);
}
.toggle.on { background: var(--accent-wash); border-color: var(--accent); color: var(--accent-deep); }

.delbtn {
  width: 26px; height: 26px; border-radius: 6px; cursor: pointer;
  background: transparent; border: 1px solid transparent; color: var(--muted);
  font-size: 18px; line-height: 1; display: grid; place-items: center;
}
.delbtn:hover { background: #fbf1f0; border-color: #e8cdca; color: var(--danger); }

.editor-foot { display: flex; align-items: center; gap: 16px; margin-top: 16px; flex-wrap: wrap; }
.foot-hint { font-size: 12px; color: var(--muted); max-width: 620px; line-height: 1.5; }

/* ---------- Botón exportar ---------- */
.tabsbar-right { display: flex; align-items: center; gap: 14px; flex-wrap: wrap; }
.export-wrap { position: relative; }
.exportbtn {
  display: inline-flex; align-items: center; gap: 7px;
  padding: 9px 15px; border: none; border-radius: 8px;
  background: var(--accent); color: #fff; font: inherit; font-weight: 600; font-size: 13px;
  cursor: pointer; box-shadow: 0 1px 0 var(--accent-deep); transition: background .15s;
}
.exportbtn:hover { background: var(--accent-deep); }
.exportbtn svg { flex: none; }
.export-hint {
  position: absolute; right: 0; top: calc(100% + 8px); z-index: 30;
  width: 280px; padding: 11px 13px; font-size: 12px; line-height: 1.5;
  background: var(--ink); color: #eef1f6; border-radius: 9px;
  box-shadow: 0 8px 28px rgba(20,30,48,.22);
}
.export-hint strong { color: #fff; }

/* ---------- Cabecera: menú de perfil ---------- */
.header { justify-content: space-between; }
.perfil { position: relative; }
.perfil-btn {
  display: flex; align-items: center; gap: 9px;
  padding: 6px 10px 6px 6px; border: 1px solid var(--line); border-radius: 24px;
  background: var(--surface); cursor: pointer; font: inherit; transition: background .12s, border-color .12s;
}
.perfil-btn:hover { background: var(--panel); border-color: var(--line-strong); }
.avatar {
  width: 30px; height: 30px; border-radius: 50%; flex: none;
  display: grid; place-items: center; font-size: 12px; font-weight: 680;
  background: var(--accent); color: #fff; letter-spacing: .02em;
}
.avatar.grande { width: 40px; height: 40px; font-size: 15px; }
.perfil-nombre { font-size: 13px; font-weight: 560; color: var(--ink); max-width: 140px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.perfil-caret { color: var(--muted); flex: none; }

.perfil-backdrop { position: fixed; inset: 0; z-index: 40; }
.perfil-menu {
  position: absolute; right: 0; top: calc(100% + 10px); z-index: 50;
  width: 270px; background: var(--surface); border: 1px solid var(--line);
  border-radius: 14px; box-shadow: 0 10px 34px rgba(20,30,48,.18); padding: 8px;
}
.perfil-head { display: flex; align-items: center; gap: 12px; padding: 10px 10px 12px; }
.perfil-info { min-width: 0; display: flex; flex-direction: column; gap: 2px; }
.perfil-info strong { font-size: 14px; font-weight: 640; }
.perfil-info span { font-size: 11.5px; color: var(--muted); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.perfil-item {
  display: flex; align-items: center; gap: 11px; width: 100%;
  padding: 10px 11px; border: none; background: transparent; border-radius: 9px;
  font: inherit; font-size: 13.5px; color: var(--ink-soft); cursor: pointer; text-align: left;
}
.perfil-item svg { color: var(--muted); flex: none; }
.perfil-item:hover { background: var(--panel); color: var(--ink); }
.perfil-item:hover svg { color: var(--accent); }
.perfil-item.salir { color: var(--danger); }
.perfil-item.salir svg { color: var(--danger); }
.perfil-item.salir:hover { background: #fbf1f0; }
.perfil-sep { height: 1px; background: var(--line); margin: 6px 4px; }

.perfil-form { padding: 8px 10px 12px; }
.perfil-volver { background: none; border: none; color: var(--accent); font: inherit; font-size: 12.5px; font-weight: 560; cursor: pointer; padding: 4px 0; margin-bottom: 6px; }
.perfil-form h4 { margin: 2px 0 4px; font-size: 14px; font-weight: 650; }
.perfil-hint { margin: 0 0 11px; font-size: 11.5px; color: var(--muted); line-height: 1.5; }
.perfil-input {
  width: 100%; padding: 9px 11px; margin-bottom: 9px;
  border: 1px solid var(--line-strong); border-radius: 8px; font: inherit; font-size: 13px;
  background: var(--surface); color: var(--ink);
}
.perfil-input:focus { outline: 2px solid var(--accent-wash); border-color: var(--accent); }
.perfil-error { background: #fbf0ef; border: 1px solid #e8cdca; color: var(--danger); padding: 7px 10px; border-radius: 7px; font-size: 12px; margin-bottom: 9px; }
.perfil-ok { background: var(--accent-wash); border: 1px solid #bfe0db; color: var(--accent-deep); padding: 7px 10px; border-radius: 7px; font-size: 12px; margin-bottom: 9px; }
.perfil-guardar {
  width: 100%; padding: 10px; border: none; border-radius: 8px;
  background: var(--accent); color: #fff; font: inherit; font-weight: 620; font-size: 13px; cursor: pointer;
}
.perfil-guardar:hover:not(:disabled) { background: var(--accent-deep); }
.perfil-guardar:disabled { background: var(--line-strong); cursor: default; }

/* login extras */
.login-link { background: none; border: none; color: var(--accent); font: inherit; font-size: 13px; font-weight: 540; cursor: pointer; margin-top: 14px; padding: 4px; }
.login-link:hover { color: var(--accent-deep); text-decoration: underline; }
.login-ok { background: var(--accent-wash); border: 1px solid #bfe0db; color: var(--accent-deep); padding: 12px 14px; border-radius: 9px; font-size: 13px; line-height: 1.5; margin-bottom: 16px; text-align: left; }

/* ---------- Pantalla de arranque ---------- */
/* ---------- Selector de grupo profesional ---------- */
.selector-wrap { max-width: 980px; margin: 0 auto; padding: 56px 22px 40px; }
.selector-head h1 { margin: 0 0 6px; font-size: 24px; font-weight: 680; letter-spacing: -.02em; }
.selector-head p { margin: 0 0 26px; color: var(--muted); font-size: 15px; }
.selector-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
.selector-card {
  display: flex; flex-direction: column; align-items: flex-start; gap: 8px; text-align: left;
  background: var(--surface); border: 1px solid var(--line); border-radius: 14px; padding: 22px 22px 18px;
  box-shadow: var(--shadow); cursor: pointer; font: inherit; color: var(--ink);
  transition: border-color .15s, transform .15s, box-shadow .15s;
}
.selector-card:hover { border-color: var(--accent); transform: translateY(-2px); box-shadow: 0 2px 4px rgba(20,30,48,.06), 0 10px 26px rgba(20,30,48,.08); }
.selector-card:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
.selector-icon { width: 48px; height: 48px; border-radius: 12px; display: grid; place-items: center; margin-bottom: 4px; }
.selector-card.is-enf .selector-icon { background: var(--accent-wash); color: var(--accent-deep); }
.selector-card.is-tcae .selector-icon { background: #efeaf8; color: #5b4a8a; }
.selector-title { font-size: 18px; font-weight: 660; letter-spacing: -.01em; }
.selector-desc { color: var(--ink-soft); font-size: 13.5px; line-height: 1.45; flex: 1; }
.selector-go { margin-top: 8px; font-size: 13px; font-weight: 600; color: var(--accent); }
.selector-card.is-tcae .selector-go { color: #5b4a8a; }
.selector-nota { margin: 22px 2px 0; color: var(--muted); font-size: 12.5px; line-height: 1.5; }
@media (max-width: 560px) {
  .selector-wrap { padding-top: 32px; }
  .selector-grid { grid-template-columns: 1fr; }
}
.colectivo-chip {
  margin-left: 16px; display: inline-flex; align-items: center; gap: 7px;
  background: var(--panel); border: 1px solid var(--line); border-radius: 999px;
  padding: 5px 6px 5px 11px; font: inherit; font-size: 12.5px; font-weight: 600; color: var(--ink); cursor: pointer;
}
.colectivo-chip:hover { border-color: var(--line-strong); }
.colectivo-dot { width: 8px; height: 8px; border-radius: 50%; }
.colectivo-dot.is-enf { background: var(--accent); }
.colectivo-dot.is-tcae { background: #7a66b3; }
.colectivo-dot.is-est { background: #c27c2c; }
.selector-card.is-est .selector-icon { background: #fbefe0; color: #8a5518; }
.selector-card.is-est .selector-go { color: #8a5518; }
.selector-grid { grid-template-columns: repeat(3, 1fr); }
@media (max-width: 760px) { .selector-grid { grid-template-columns: 1fr; } }
.est-wrap { max-width: 1080px; margin: 0 auto; padding: 24px 22px 48px; display: flex; flex-direction: column; gap: 16px; }
.est-top { display: flex; justify-content: space-between; align-items: center; gap: 12px; flex-wrap: wrap; }
.est-anio { display: inline-flex; align-items: center; gap: 4px; background: var(--surface); border: 1px solid var(--line); border-radius: 10px; padding: 3px; }
.est-anio span { font-weight: 680; font-size: 16px; padding: 0 10px; font-variant-numeric: tabular-nums; }
.est-anio button { border: 0; background: transparent; font-size: 18px; width: 30px; height: 30px; border-radius: 7px; cursor: pointer; color: var(--ink-soft); }
.est-anio button:hover { background: var(--panel); }
.est-acciones { display: flex; gap: 8px; flex-wrap: wrap; }
.est-card { background: var(--surface); border: 1px solid var(--line); border-radius: 14px; padding: 20px 22px; box-shadow: var(--shadow); }
.est-card-head { display: flex; justify-content: space-between; align-items: baseline; gap: 12px; margin-bottom: 14px; flex-wrap: wrap; }
.est-card h2, .est-modal-card h2 { margin: 0; font-size: 16.5px; font-weight: 660; letter-spacing: -.01em; }
.est-muted { color: var(--muted); font-size: 13px; margin: 0; }
.est-pista { margin-top: 8px; }
.est-alerta { background: #fdecec; border: 1px solid #f3c2c2; color: #8a1f1f; border-radius: 10px; padding: 10px 14px; font-size: 13.5px; }
.est-aviso { background: var(--panel); border: 1px solid var(--line); border-radius: 10px; padding: 10px 14px; font-size: 13px; color: var(--ink-soft); }
.est-leyenda { display: flex; gap: 14px; font-size: 12.5px; color: var(--ink-soft); }
.est-leyenda i, .est-tabla th i { display: inline-block; width: 10px; height: 10px; border-radius: 3px; margin-right: 6px; vertical-align: -1px; }
.est-barras { display: grid; grid-template-columns: repeat(12, 1fr); gap: 8px; height: 200px; align-items: end; }
.est-barra { display: flex; flex-direction: column; align-items: center; justify-content: flex-end; height: 100%; border: 0; background: transparent; cursor: pointer; padding: 4px 0; border-radius: 8px; font: inherit; }
.est-barra:hover { background: var(--panel); }
.est-barra.is-sel { background: var(--accent-wash); }
.est-barra-n { font-size: 12.5px; font-weight: 650; height: 18px; font-variant-numeric: tabular-nums; }
.est-barra-col { width: 62%; max-width: 34px; flex: 1; display: flex; flex-direction: column; justify-content: flex-end; gap: 2px; }
.est-barra-col span { display: block; width: 100%; border-radius: 4px; }
.est-barra-m { font-size: 12px; color: var(--ink-soft); margin-top: 6px; }
.est-barra.is-sel .est-barra-m { color: var(--ink); font-weight: 650; }
.est-kpis { display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-bottom: 16px; }
.est-kpi { background: var(--panel); border-radius: 12px; padding: 12px 14px; display: flex; flex-direction: column; }
.est-kpi b { font-size: 28px; font-weight: 700; line-height: 1.1; font-variant-numeric: tabular-nums; }
.est-kpi span { font-size: 12.5px; color: var(--muted); }
.est-mods { display: flex; gap: 3px; height: 30px; border-radius: 8px; overflow: hidden; margin-bottom: 16px; }
.est-mod { display: flex; align-items: center; justify-content: center; min-width: 0; }
.est-mod span { font-size: 12px; font-weight: 600; color: #1d2530; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; padding: 0 6px; }
.est-tabla { width: 100%; border-collapse: collapse; font-size: 13.5px; }
.est-tabla th { text-align: left; font-size: 12px; font-weight: 600; color: var(--muted); padding: 8px 10px; border-bottom: 1px solid var(--line); white-space: nowrap; }
.est-tabla td { padding: 9px 10px; border-bottom: 1px solid var(--line); }
.est-tabla tr:last-child td { border-bottom: 0; }
.est-num { text-align: center; font-variant-numeric: tabular-nums; }
.est-tag { display: inline-block; color: #fff; font-size: 11px; font-weight: 650; border-radius: 999px; padding: 2px 8px; margin-right: 4px; }
.est-scroll { overflow-x: auto; }
.est-nowrap { white-space: nowrap; }
.est-link { border: 0; background: transparent; color: var(--accent); font: inherit; font-size: 12.5px; font-weight: 600; cursor: pointer; padding: 2px 6px; }
.est-del { color: #b33; }
.est-modal { position: fixed; inset: 0; background: rgba(20,30,48,.35); display: grid; place-items: center; z-index: 50; padding: 16px; }
.est-modal-card { background: var(--surface); border-radius: 14px; padding: 22px; width: min(560px, 100%); box-shadow: 0 10px 40px rgba(20,30,48,.25); display: flex; flex-direction: column; gap: 14px; }
.est-form { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
.est-form label { display: flex; flex-direction: column; gap: 4px; font-size: 12.5px; color: var(--ink-soft); font-weight: 600; }
.est-form input, .est-form select { font: inherit; font-size: 14px; font-weight: 400; color: var(--ink); border: 1px solid var(--line); border-radius: 8px; padding: 8px 10px; background: var(--surface); }
.est-form-full { grid-column: 1 / -1; }
.est-modal-btns { display: flex; justify-content: flex-end; gap: 8px; }
.est-btn { font: inherit; font-size: 13px; font-weight: 580; padding: 8px 14px; border-radius: 8px; border: 1px solid var(--line-strong); background: var(--surface); color: var(--ink); cursor: pointer; white-space: nowrap; }
.est-btn:hover { background: var(--panel); }
.est-btn-primary { background: var(--accent); border-color: var(--accent); color: #fff; }
.est-btn-primary:hover { background: var(--accent-deep); }
.est-btn:disabled { opacity: .6; cursor: default; }
.est-btn-del { color: #b33; border-color: #e6c3c3; }
.est-vacio p { margin: 6px 0 0; color: var(--ink-soft); font-size: 14px; }
.est-filtros { display: flex; gap: 12px; align-items: center; flex-wrap: wrap; }
.est-busca { font: inherit; font-size: 13px; border: 1px solid var(--line); border-radius: 8px; padding: 7px 10px; width: 260px; max-width: 100%; }
.est-check { display: inline-flex; align-items: center; gap: 6px; font-size: 13px; color: var(--ink-soft); cursor: pointer; }
.est-periodos { display: flex; flex-direction: column; gap: 14px; }
.est-periodo { border: 1px solid var(--line); border-radius: 12px; overflow: hidden; }
.est-periodo header { background: var(--panel); padding: 10px 14px; border-bottom: 1px solid var(--line); }
.est-periodo-fechas { font-weight: 660; font-size: 14.5px; }
.est-periodo-fechas span { color: var(--muted); margin: 0 4px; }
.est-periodo-meta { font-size: 12.5px; color: var(--muted); margin-top: 2px; }
.est-periodo .est-tabla th, .est-periodo .est-tabla td { padding-left: 14px; }
.est-tutora { font-size: 12.5px; text-transform: uppercase; letter-spacing: .01em; color: var(--ink-soft); }
.est-falta { color: #b33; text-transform: none; font-style: italic; }
.est-turno { display: inline-block; font-size: 11.5px; font-weight: 650; color: #1d2530; border-radius: 999px; padding: 2px 9px; white-space: nowrap; }
.est-contacto { font-size: 12.5px; color: var(--ink-soft); word-break: break-all; }
.est-lista { table-layout: fixed; }
.est-cal { display: grid; grid-template-columns: repeat(7, minmax(110px, 1fr)); gap: 4px; min-width: 780px; margin-bottom: 16px; }
.est-cal-dow { font-size: 11.5px; font-weight: 650; color: var(--muted); text-transform: uppercase; letter-spacing: .04em; padding: 2px 6px; }
.est-cal-dia { border: 1px solid var(--line); border-radius: 8px; padding: 5px; min-height: 74px; display: flex; flex-direction: column; gap: 3px; background: var(--surface); }
.est-cal-dia.is-vacio { border: 0; background: transparent; }
.est-cal-dia.is-finde { background: var(--panel); }
.est-cal-num { font-size: 12px; font-weight: 650; color: var(--ink-soft); display: flex; justify-content: space-between; }
.est-cal-num span { font-weight: 600; color: var(--muted); font-size: 11px; }
.est-cal-chip { font-size: 11px; font-weight: 600; border-radius: 5px; padding: 2px 5px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
@media (max-width: 640px) {
  .est-periodo .est-lista, .est-periodo .est-lista tbody { display: block; }
  .est-periodo .est-lista thead, .est-periodo .est-lista colgroup { display: none; }
  .est-periodo .est-lista tr { display: grid; grid-template-columns: auto auto 1fr auto; grid-template-areas: "est est est tur" "tut tut tut edi" "esc cur . ." "con con con con"; gap: 3px 10px; padding: 11px 14px; border-bottom: 1px solid var(--line); align-items: center; }
  .est-periodo .est-lista tr:last-child { border-bottom: 0; }
  .est-periodo .est-lista td { padding: 0 !important; border: 0; }
  .est-c-est { grid-area: est; } .est-c-tur { grid-area: tur; justify-self: end; } .est-c-tut { grid-area: tut; }
  .est-c-edi { grid-area: edi; justify-self: end; } .est-c-esc { grid-area: esc; font-size: 12.5px; color: var(--muted); }
  .est-c-cur { grid-area: cur; font-size: 12.5px; color: var(--muted); } .est-c-con { grid-area: con; }
  .est-c-tut::before { content: "Tutora: "; text-transform: none; color: var(--muted); }
  .est-busca { width: 100%; }
}
@media (max-width: 640px) {
  .est-barras { gap: 3px; height: 170px; }
  .est-barra-m { font-size: 10.5px; }
  .est-kpi b { font-size: 22px; }
  .est-form { grid-template-columns: 1fr; }
  .est-tabla { font-size: 12.5px; }
}
.colectivo-cambiar { font-weight: 520; color: var(--muted); background: var(--surface); border: 1px solid var(--line); border-radius: 999px; padding: 2px 8px; font-size: 11.5px; }
@media (max-width: 560px) { .colectivo-cambiar { display: none; } .colectivo-chip { margin-left: 8px; } }
.bootscreen { display: grid; place-items: center; height: 100vh; color: var(--muted); font-size: 14px; }

/* ---------- Pantalla de login ---------- */
.login-wrap { min-height: 100vh; display: grid; place-items: center; padding: 24px; background: var(--paper); }
.login-card {
  width: 100%; max-width: 380px; background: var(--surface);
  border: 1px solid var(--line); border-radius: 16px; padding: 36px 32px;
  box-shadow: 0 6px 30px rgba(20,30,48,.08); text-align: center;
}
.login-mark {
  width: 52px; height: 52px; margin: 0 auto 16px; border-radius: 13px;
  display: grid; place-items: center; background: var(--accent-wash); color: var(--accent);
}
.login-card h1 { margin: 0 0 4px; font-size: 21px; font-weight: 680; letter-spacing: -.02em; }
.login-sub { margin: 0 0 26px; font-size: 13px; color: var(--muted); }
.login-field { display: block; text-align: left; margin-bottom: 14px; }
.login-field span { display: block; font-size: 12px; font-weight: 540; color: var(--ink-soft); margin-bottom: 5px; }
.login-field input {
  width: 100%; padding: 11px 13px; font: inherit; font-size: 14px;
  border: 1px solid var(--line-strong); border-radius: 9px; background: var(--surface); color: var(--ink);
}
.login-field input:focus { outline: 2px solid var(--accent-wash); border-color: var(--accent); }
.login-error { background: #fbf0ef; border: 1px solid #e8cdca; color: var(--danger); padding: 9px 12px; border-radius: 8px; font-size: 12.5px; margin-bottom: 14px; text-align: left; }
.login-btn {
  width: 100%; padding: 12px; margin-top: 4px;
  background: var(--accent); color: #fff; border: none; border-radius: 9px;
  font: inherit; font-weight: 650; font-size: 14.5px; cursor: pointer;
  box-shadow: 0 1px 0 var(--accent-deep); transition: background .15s;
}
.login-btn:hover:not(:disabled) { background: var(--accent-deep); }
.login-btn:disabled { background: var(--line-strong); box-shadow: none; cursor: default; }
.login-foot { margin: 22px 0 0; font-size: 11.5px; line-height: 1.5; color: var(--muted); }

/* ---------- Editor de estructura ---------- */
.estructura { max-width: 720px; }
.est-card { background: var(--surface); border: 1px solid var(--line); border-radius: var(--radius); overflow: hidden; box-shadow: var(--shadow); }
.est-row { display: grid; grid-template-columns: 78px 1fr 120px 90px 160px; align-items: center; gap: 10px; padding: 11px 16px; border-bottom: 1px solid var(--line); }
.est-row:last-of-type { border-bottom: none; }
.est-head-row { background: var(--panel); font-size: 11.5px; font-weight: 620; color: var(--ink-soft); letter-spacing: .02em; text-transform: uppercase; }
.est-head-row span { font-size: 11px; }
.est-row.inactiva { opacity: .5; }
.est-c-nombre { font-size: 14px; font-weight: 560; }
.rot-tag { display: inline-block; padding: 2px 9px; border-radius: 20px; font-size: 11px; font-weight: 600; background: var(--accent-wash); color: var(--accent-deep); }
.est-c-pos input { width: 70px; padding: 7px 9px; border: 1px solid var(--line-strong); border-radius: 7px; font: inherit; font-size: 13px; text-align: center; background: var(--surface); color: var(--ink); }
.est-c-pos input:disabled { background: var(--panel); color: var(--line-strong); }
.est-c-pos input:focus { outline: 2px solid var(--accent-wash); border-color: var(--accent); }
.chk { width: 28px; height: 28px; border-radius: 7px; border: 1.5px solid var(--line-strong); background: var(--surface); cursor: pointer; font-size: 14px; font-weight: 700; color: #fff; display: grid; place-items: center; transition: all .12s; }
.chk.on { background: var(--accent); border-color: var(--accent); }
.est-total { display: flex; align-items: center; justify-content: space-between; padding: 13px 16px; background: var(--panel); font-size: 13px; color: var(--ink-soft); }
.est-total strong { font-size: 18px; color: var(--ink); }
.est-actions { display: flex; justify-content: flex-end; gap: 10px; margin-top: 16px; }

/* ---------- Responsive ---------- */
@media (max-width: 860px) {
  .layout { grid-template-columns: 1fr; }
  .sidebar { position: static; height: auto; border-right: none; border-bottom: 1px solid var(--line); }
  .main { padding: 18px 14px 50px; }
}

@media (max-width: 560px) {
  .header { padding: 0 12px; gap: 8px; }
  .brand-sub { display: none; }              /* el subtítulo largo desaparece en móvil */
  .brand-mark { width: 34px; height: 34px; }
  .brand-title { font-size: 15px; }
  .perfil-nombre { max-width: 90px; }
  .perfil-menu { width: calc(100vw - 24px); right: -6px; }
  .tabsbar-right { width: 100%; }
  .export-hint { width: 240px; }
  /* Plantilla en móvil: la barra de herramientas se apila y la tabla
     se desplaza de lado (la tabla lleva min-width y overflow-x: auto). */
  .editor-bar { flex-wrap: wrap; gap: 8px; }
  .editor-bar .search { flex: 1 1 100%; }
  .editor-actions { flex: 1 1 100%; }
  .editor-actions button { flex: 1; }
  .scroll-hint { display: block; }
}
/* Aviso de "desliza de lado", solo visible en pantallas estrechas */
.scroll-hint { display: none; font-size: 12px; color: var(--muted); margin: 0 0 8px; }

@media (prefers-reduced-motion: reduce) {
  * { transition: none !important; }
}

`;
