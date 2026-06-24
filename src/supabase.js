/* ============================================================
   Conexión con Supabase (base de datos en la nube + login)
   ------------------------------------------------------------
   Estos dos valores son PÚBLICOS y seguros para estar aquí:
   la protección real de los datos la dan las reglas de seguridad
   (Row Level Security) configuradas en la base de datos, que solo
   permiten acceso a usuarios con la sesión iniciada.
   ============================================================ */
import { createClient } from "@supabase/supabase-js";

const SUPABASE_URL = "https://hsosqcylazbxhiljiiag.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imhzb3NxY3lsYXpieGhpbGppaWFnIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODIyNzg0OTQsImV4cCI6MjA5Nzg1NDQ5NH0.PRWFYP2uCzt0fl-aIeGSyHKaC_oIv1VMl0XLoGC7d84";

export const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

/* ---------- Autenticación ---------- */
export async function iniciarSesion(email, password) {
  const { data, error } = await supabase.auth.signInWithPassword({ email, password });
  if (error) throw error;
  return data;
}
export async function cerrarSesion() {
  await supabase.auth.signOut();
}
export async function sesionActual() {
  const { data } = await supabase.auth.getSession();
  return data.session;
}
export function onCambioSesion(callback) {
  const { data } = supabase.auth.onAuthStateChange((_event, session) => callback(session));
  return data.subscription;
}

/* ---------- Profesionales (plantilla compartida) ----------
   La app usa registros con forma { nombre, turno, apodo, criterio,
   suplente, antiguedad }. La tabla usa esas mismas columnas + id.
*/
export async function cargarProfesionales() {
  const { data, error } = await supabase
    .from("profesionales")
    .select("nombre, apodo, turno, criterio, suplente, antiguedad")
    .order("nombre", { ascending: true });
  if (error) throw error;
  return data || [];
}

// Reemplaza por completo la plantilla: borra todo y vuelve a insertar.
// Es la operación más simple y robusta para "guardar la plantilla editada".
export async function guardarProfesionales(profs) {
  // 1. Borrar todos los registros existentes
  const del = await supabase.from("profesionales").delete().neq("id", -1);
  if (del.error) throw del.error;
  // 2. Insertar los nuevos (si hay)
  const limpio = (profs || [])
    .filter(p => p && p.nombre && String(p.nombre).trim())
    .map(p => ({
      nombre: p.nombre,
      apodo: p.apodo || "",
      turno: p.turno || "TM",
      criterio: p.criterio || "SI",
      suplente: p.suplente || "No",
      antiguedad: p.antiguedad || "< 3 ANY",
    }));
  if (limpio.length === 0) return [];
  const { data, error } = await supabase.from("profesionales").insert(limpio).select();
  if (error) throw error;
  return data || [];
}

/* ---------- Histórico de rotación (compartido) ----------
   La app usa filas con { fechaSerial, fecha(Date), mes, profNorm,
   prof, dia, turno, posGlobal, area }. La tabla usa nombres con
   guion_bajo; convertimos en ambos sentidos.
*/
function filaAppARegistro(f) {
  return {
    fecha_serial: f.fechaSerial,
    fecha: (f.fecha instanceof Date ? f.fecha : new Date(f.fecha)).toISOString().slice(0, 10),
    mes: f.mes,
    prof_norm: f.profNorm,
    prof: f.prof,
    dia: f.dia,
    turno: f.turno,
    pos_global: f.posGlobal,
    area: f.area,
  };
}
function registroAFilaApp(r) {
  return {
    fechaSerial: r.fecha_serial,
    fecha: new Date(r.fecha),
    mes: r.mes,
    profNorm: r.prof_norm,
    prof: r.prof,
    dia: r.dia,
    turno: r.turno,
    posGlobal: r.pos_global,
    area: r.area,
  };
}

export async function cargarHistorico() {
  // Puede haber muchas filas; paginamos para traerlas todas.
  const PAGE = 1000;
  let desde = 0;
  let todo = [];
  for (;;) {
    const { data, error } = await supabase
      .from("historico")
      .select("fecha_serial, fecha, mes, prof_norm, prof, dia, turno, pos_global, area")
      .order("fecha_serial", { ascending: true })
      .range(desde, desde + PAGE - 1);
    if (error) throw error;
    todo = todo.concat((data || []).map(registroAFilaApp));
    if (!data || data.length < PAGE) break;
    desde += PAGE;
  }
  return todo;
}

// Añade filas nuevas al histórico (no borra lo existente).
export async function añadirAlHistorico(filasApp) {
  const registros = (filasApp || []).map(filaAppARegistro);
  if (registros.length === 0) return;
  // Insertar en lotes para no exceder límites de la petición
  const LOTE = 500;
  for (let i = 0; i < registros.length; i += LOTE) {
    const { error } = await supabase.from("historico").insert(registros.slice(i, i + LOTE));
    if (error) throw error;
  }
}

export async function vaciarHistorico() {
  const { error } = await supabase.from("historico").delete().neq("id", -1);
  if (error) throw error;
}
