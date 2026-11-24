/**
 * Coevaluación — el "cerebro" que vive en Google (Apps Script).
 * Guarda las calificaciones en una hoja privada de Google. La lista de alumnos
 * (matrículas) y la contraseña del profe viven aquí, NO en la página pública.
 *
 * Instalación (una sola vez):
 *  1. script.google.com → Proyecto nuevo → pegar este archivo.
 *  2. Llenar ALUMNOS y ADMIN_INICIAL (solo en el editor; en GitHub van vacíos).
 *  3. Ejecutar setup (crea la hoja «Coevaluación Proyecto N - respuestas»).
 *  4. Implementar → Aplicación web · Ejecutar como: yo · Acceso: cualquier persona.
 *  5. Pegar la URL /exec en CONFIG.API de index.html.
 */

const EQUIPOS = ["Fashion B", "Bags Pro", "Prime Bags", "Superbag", "Mon Sac a Dos", "Saga Sacs", "Urban Sacs"];
const DIMS = ["Conocimientos", "Decisiones", "Narrativa", "Visuales", "Organización", "Nos representa"];
const ADMIN_INICIAL = "CAMBIAR";
/* [matrícula, nombre, equipo] — solo en el editor de Apps Script */
const ALUMNOS = [];

const COLS = ["matricula", "nombre", "equipo_propio", "equipo_evaluado"].concat(DIMS, ["promedio", "estado", "actualizado"]);

function setup() {
  const p = props();
  const id = p.getProperty("SHEET_ID");
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create("Coevaluación Proyecto N - respuestas");
  p.setProperty("SHEET_ID", ss.getId());

  const a = ss.getSheetByName("Alumnos") || ss.insertSheet("Alumnos");
  if (ALUMNOS.length) {
    a.clear();
    a.getRange(1, 1, 1, 3).setValues([["matricula", "nombre", "equipo"]]).setFontWeight("bold");
    a.getRange(2, 1, ALUMNOS.length, 1).setNumberFormat("@");
    a.getRange(2, 1, ALUMNOS.length, 3).setValues(ALUMNOS);
    a.setFrozenRows(1);
    a.autoResizeColumns(1, 3);
  }

  const e = ss.getSheetByName("Evaluaciones") || ss.insertSheet("Evaluaciones", 0);
  if (e.getLastRow() < 1) {
    e.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight("bold");
    e.setFrozenRows(1);
    e.getRange("A:A").setNumberFormat("@");
  }

  const r = ss.getSheetByName("Resumen") || ss.insertSheet("Resumen", 0);
  r.clear();
  const head = ["Equipo"].concat(DIMS, ["Promedio", "Evaluaciones enviadas"]);
  r.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
  const letras = "EFGHIJK"; // columnas de Evaluaciones: 6 dimensiones + promedio
  EQUIPOS.forEach(function (eq, i) {
    const row = i + 2;
    r.getRange(row, 1).setValue(eq);
    for (let c = 0; c < 7; c++) {
      const L = letras[c];
      r.getRange(row, c + 2).setFormula('=IFERROR(ROUND(AVERAGEIFS(Evaluaciones!$' + L + ':$' + L + ',Evaluaciones!$D:$D,$A' + row + ',Evaluaciones!$L:$L,"enviado"),1),"")');
    }
    r.getRange(row, 9).setFormula('=COUNTIFS(Evaluaciones!$D:$D,$A' + row + ',Evaluaciones!$L:$L,"enviado")');
  });
  r.setFrozenRows(1);
  r.getRange(2, 8, EQUIPOS.length, 1).setFontWeight("bold");

  ["Sheet1", "Hoja 1", "Hoja1"].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  if (!p.getProperty("ADMIN_PASS")) p.setProperty("ADMIN_PASS", ADMIN_INICIAL);
  CacheService.getScriptCache().remove("alumnos");
  Logger.log(ss.getUrl());
}

function props() { return PropertiesService.getScriptProperties(); }
function libro() { return SpreadsheetApp.openById(props().getProperty("SHEET_ID")); }

function doGet() { return salida({ ok: true, msg: "Coevaluación activa" }); }

function doPost(e) {
  let out;
  try {
    const d = JSON.parse(e.postData.contents);
    switch (d.action) {
      case "login": out = login(d); break;
      case "draft": out = guardar(d, false); break;
      case "submit": out = guardar(d, true); break;
      case "profe": admin(d); out = tablero(); break;
      default: throw new Error("Acción desconocida.");
    }
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return salida(out);
}

function salida(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function admin(d) {
  if (String(d.pass || "") !== props().getProperty("ADMIN_PASS")) {
    Utilities.sleep(1500);
    throw new Error("Contraseña incorrecta.");
  }
}

function norm(m) { return String(m || "").toUpperCase().replace(/[^A-Z0-9]/g, ""); }

function alumnos() {
  const c = CacheService.getScriptCache();
  const hit = c.get("alumnos");
  if (hit) return JSON.parse(hit);
  const sh = libro().getSheetByName("Alumnos");
  const v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues() : [];
  const out = v.filter(function (r) { return r[0]; }).map(function (r) { return { m: norm(r[0]), n: String(r[1]).trim(), e: String(r[2]).trim() }; });
  c.put("alumnos", JSON.stringify(out), 600);
  return out;
}

function alumno(m) {
  const a = alumnos().filter(function (x) { return x.m === norm(m); })[0];
  if (!a) {
    Utilities.sleep(700);
    throw new Error("No encontramos esa matrícula. Revísala: empieza con A0…");
  }
  return a;
}

function pila(nombre) { // "Amaya Pineda, María Fernanda" → "María Fernanda"
  const i = nombre.indexOf(",");
  return i >= 0 ? nombre.slice(i + 1).trim() : nombre;
}

function filas() {
  const sh = libro().getSheetByName("Evaluaciones");
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues().filter(function (r) { return r[0]; });
}

function deAlumno(m, rows) {
  const ev = {};
  let enviado = false;
  rows.forEach(function (r) {
    if (norm(r[0]) !== m) return;
    const t = r[12] instanceof Date ? r[12].getTime() : (Date.parse(r[12]) || 0);
    ev[r[3]] = { d: r.slice(4, 10).map(Number), t: t };
    if (r[11] === "enviado") enviado = true;
  });
  return { evals: ev, enviado: enviado };
}

function login(d) {
  const a = alumno(d.matricula);
  const yo = deAlumno(a.m, filas());
  return { ok: true, matricula: a.m, nombre: pila(a.n), equipo: a.e, evals: yo.evals, enviado: yo.enviado };
}

function dims(x) {
  if (!x || x.length !== DIMS.length) throw new Error("Faltan calificaciones.");
  return x.map(function (v) {
    v = Math.round(Number(v));
    if (!(v >= 0 && v <= 100)) throw new Error("Calificación fuera de rango.");
    return v;
  });
}

/* draft: un equipo, estado "borrador".  submit: todos los equipos (menos el propio), estado "enviado". */
function guardar(d, final) {
  const a = alumno(d.matricula);
  const otros = EQUIPOS.filter(function (x) { return x !== a.e; });
  const evals = {};
  if (final) {
    otros.forEach(function (eq) { evals[eq] = dims(d.evals && d.evals[eq]); });
  } else {
    if (otros.indexOf(d.equipo) < 0) throw new Error("Ese equipo no lo puedes evaluar.");
    evals[d.equipo] = dims(d.d);
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = libro().getSheetByName("Evaluaciones");
    const rows = filas();
    if (deAlumno(a.m, rows).enviado) {
      if (final) return { ok: true, ya: true };
      throw new Error("Ya enviaste tus calificaciones.");
    }
    const idx = {};
    rows.forEach(function (r, i) { idx[norm(r[0]) + "|" + r[3]] = i + 2; });
    const ahora = new Date();
    const nuevas = [];
    Object.keys(evals).forEach(function (eq) {
      const v = evals[eq];
      const prom = Math.round(v.reduce(function (s, x) { return s + x; }, 0) / v.length * 10) / 10;
      const fila = [a.m, a.n, a.e, eq].concat(v, [prom, final ? "enviado" : "borrador", ahora]);
      const at = idx[a.m + "|" + eq];
      if (at) sh.getRange(at, 1, 1, fila.length).setValues([fila]);
      else nuevas.push(fila);
    });
    if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, COLS.length).setValues(nuevas);
    return { ok: true, t: ahora.getTime() };
  } finally {
    lock.releaseLock();
  }
}

function tablero() {
  const rows = filas();
  const al = alumnos();
  const eq = {};
  EQUIPOS.forEach(function (x) { eq[x] = { n: 0, s: [0, 0, 0, 0, 0, 0], b: 0 }; });
  const est = {};
  rows.forEach(function (r) {
    const m = norm(r[0]);
    est[m] = est[m] || { k: 0, env: false };
    est[m].k++;
    if (r[11] === "enviado") est[m].env = true;
    const t = eq[r[3]];
    if (!t) return;
    if (r[11] === "enviado") {
      t.n++;
      for (let i = 0; i < 6; i++) t.s[i] += Number(r[4 + i]) || 0;
    } else t.b++;
  });
  return {
    ok: true,
    dims: DIMS,
    equipos: EQUIPOS.map(function (x) {
      const t = eq[x];
      return { nombre: x, n: t.n, borradores: t.b, prom: t.n ? t.s.map(function (s) { return Math.round(s / t.n * 10) / 10; }) : null };
    }),
    alumnos: al.map(function (x) {
      const s = est[x.m] || { k: 0, env: false };
      return { nombre: x.n, equipo: x.e, estado: s.env ? "enviado" : (s.k ? "en curso" : "sin empezar"), k: s.env ? 6 : s.k };
    })
  };
}
