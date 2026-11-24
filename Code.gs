/**
 * La quincena — el "cerebro" que vive en Google Apps Script.
 * Guarda los gastos en una hoja de Google privada y lee el presupuesto de la pestaña «Presupuesto».
 * La página pública no trae ningún número: todo pasa por aquí, y solo con la clave correcta.
 *
 * Instalación (una sola vez):
 *  1. script.google.com → Proyecto nuevo → pegar este archivo.
 *  2. Poner los montos en PRESUPUESTO y ejecutar «setup»: crea la hoja, genera una clave al azar
 *     y deja tu enlace personal (con la clave) en la pestaña «Tu enlace».
 *  3. Implementar → Nueva implementación → Aplicación web
 *     · Ejecutar como: yo   · Quién tiene acceso: cualquier persona
 *  4. Pegar la URL que termina en /exec en API_URL de index.html.
 * Después, el presupuesto se cambia directo en la pestaña «Presupuesto» de la hoja.
 * Para cambiar la clave: borrar la propiedad CLAVE del script y volver a ejecutar «setup».
 */

const APP_URL = "https://pablosalazarrojo-stack.github.io/la-quincena/";

// tipo: ingreso | rubro | pension | caja. fijo: «sí» si es un pago fijo. subcategorias: separadas por coma.
const PRESUPUESTO = [
  ["tipo", "clave", "nombre", "q1", "q2", "fijo", "subcategorias"],
  ["ingreso", "sueldo", "Sueldo", 0, 0, "", ""],
  ["ingreso", "vales", "Vales de despensa", 0, 0, "", ""],
  ["rubro", "casa", "Casa", 0, 0, "sí", ""],
  ["rubro", "super", "Súper", 0, 0, "", "Costco, Soriana, Snack"],
  ["rubro", "suplementos", "Suplementos", 0, 0, "", ""],
  ["rubro", "celulares", "Celulares", 0, 0, "sí", ""],
  ["rubro", "transporte", "Transporte", 0, 0, "", ""],
  ["rubro", "pelu", "Pelu", 0, 0, "", ""],
  ["rubro", "lavanderia", "Lavandería", 0, 0, "", ""],
  ["rubro", "social", "Social", 0, 0, "", ""],
  ["rubro", "plataformas", "Plataformas", 0, 0, "sí", "Apple, Claude, Netflix, Google"],
  ["rubro", "otros", "Otros", 0, 0, "", ""],
  ["pension", "pension", "Pensión", 0, 0, "", ""],
  ["caja", "caja", "Caja de ahorro", 0, 0, "", ""]
];

const HOJA_G = "Gastos";
const HOJA_P = "Presupuesto";
const HOJA_C = "Cierre de año";
const COLS_C = ["tabla", "seccion", "clave", "concepto", "presupuesto", "real"];
const HOJA_CM = "Cierre movimientos";
const HOJA_R = "Reportes";
const COLS_R = ["id", "fecha", "rubro", "sub", "monto", "nota", "reportado", "creado"];
const COLS_CM = ["id", "fecha", "clave", "monto", "nota", "creado"];
const COLS = ["id", "fecha", "mes", "quincena", "rubro", "sub", "monto", "nota", "creado", "editado"];

function props() { return PropertiesService.getScriptProperties(); }
function libro() {
  const id = props().getProperty("SHEET_ID");
  if (!id) throw new Error("Falta correr setup.");
  return SpreadsheetApp.openById(id);
}

function setup() {
  const p = props();
  let ss = p.getProperty("SHEET_ID") ? SpreadsheetApp.openById(p.getProperty("SHEET_ID")) : null;
  if (!ss) { ss = SpreadsheetApp.create("La quincena - gastos"); p.setProperty("SHEET_ID", ss.getId()); }
  const g = ss.getSheetByName(HOJA_G) || ss.insertSheet(HOJA_G);
  g.getRange("A:C").setNumberFormat("@");
  g.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight("bold");
  g.setFrozenRows(1);
  let pr = ss.getSheetByName(HOJA_P);
  if (!pr) {
    pr = ss.insertSheet(HOJA_P);
    pr.getRange(1, 1, PRESUPUESTO.length, PRESUPUESTO[0].length).setValues(PRESUPUESTO);
    pr.getRange(1, 1, 1, PRESUPUESTO[0].length).setFontWeight("bold");
    pr.setFrozenRows(1);
  }
  ["Sheet1", "Hoja 1", "Hoja1"].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  let c = p.getProperty("CLAVE");
  if (!c) {
    c = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "").slice(0, 40);
    p.setProperty("CLAVE", c);
  }
  const t = ss.getSheetByName("Tu enlace") || ss.insertSheet("Tu enlace");
  t.clear();
  t.getRange("A1").setValue("Tu enlace personal a la app. Ábrelo en tu celular y guárdalo en la pantalla de inicio. No lo compartas: trae tu clave.");
  t.getRange("A3").setRichTextValue(SpreadsheetApp.newRichTextValue().setText("Abrir la quincena").setLinkUrl(APP_URL + "#k=" + c).build());
  t.getRange("A3").setFontSize(16).setFontWeight("bold");
  t.setColumnWidth(1, 520);
  Logger.log("Hoja lista. Tu enlace está en la pestaña «Tu enlace».");
}

function doGet() { return salida({ ok: true, msg: "La quincena activa" }); }

function doPost(e) {
  let out;
  try {
    const d = JSON.parse(e.postData.contents);
    checar(d.k);
    switch (d.action) {
      case "load": { const f = filasP(); out = { ok: true, cfg: cfgDe(f), filas: f, pagos: pagos(), cierre: filasCierre(), cierreMovs: movsCierre(), reportes: reportes(), gastos: leer(String(d.mes || "")) }; break; }
      case "savePresupuesto": { guardarPresupuesto(d.filas); const f = filasP(); out = { ok: true, cfg: cfgDe(f), filas: f }; break; }
      case "cierreReal": guardarRealCierre(String(d.clave || ""), d.real); out = { ok: true }; break;
      case "repAdd": out = { ok: true, id: agregarRep(d.doc) }; break;
      case "repSave": guardarRep(String(d.id || ""), d.doc); out = { ok: true }; break;
      case "repMarcar": out = { ok: true, reportado: marcarRep(String(d.id || ""), !!d.si) }; break;
      case "repDel": borrarRep(String(d.id || "")); out = { ok: true }; break;
      case "cierrePresupuesto": guardarPresupuestoCierre(d.valores); out = { ok: true, cierre: filasCierre() }; break;
      case "cierreMov": out = Object.assign({ ok: true }, agregarMovCierre(d)); break;
      case "cierreMovSave": out = Object.assign({ ok: true }, editarMovCierre(String(d.id || ""), d)); break;
      case "cierreMovDel": out = Object.assign({ ok: true }, borrarMovCierre(String(d.id || ""))); break;
      case "pago": guardarPago(String(d.clave || ""), String(d.fecha || "")); out = { ok: true, pagos: pagos() }; break;
      case "add": out = { ok: true, id: agregar(d.doc) }; break;
      case "save": guardar(String(d.id || ""), d.doc); out = { ok: true }; break;
      case "del": borrar(d.ids); out = { ok: true }; break;
      default: throw new Error("Acción desconocida.");
    }
  } catch (err) {
    out = { ok: false, error: String(err.message || err), code: err.code || "" };
  }
  return salida(out);
}

function salida(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function checar(k) {
  const c = props().getProperty("CLAVE");
  if (!c || c === "CAMBIAR" || String(k || "") !== c) {
    Utilities.sleep(1500); // frena a quien intente adivinar
    const e = new Error("Clave incorrecta.");
    e.code = "clave";
    throw e;
  }
}

function limpio(s, n) { return String(s == null ? "" : s).trim().replace(/\s+/g, " ").slice(0, n || 60); }
function sinAcentos(s) { return String(s || "").normalize("NFD").replace(/[̀-ͯ]/g, ""); }
function conCandado(fn) {
  const l = LockService.getScriptLock();
  l.waitLock(10000);
  try { return fn(); } finally { l.releaseLock(); }
}

function leer(mes) {
  const ss = libro(), sh = ss.getSheetByName(HOJA_G), tz = ss.getSpreadsheetTimeZone();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      const f = r[1] instanceof Date ? Utilities.formatDate(r[1], tz, "yyyy-MM-dd") : String(r[1]);
      const m = r[2] instanceof Date ? Utilities.formatDate(r[2], tz, "yyyy-MM") : String(r[2] || f.slice(0, 7));
      return {
        id: String(r[0]), fecha: f, mes: m, q: Number(r[3]) === 1 ? 1 : 2, rubro: String(r[4]),
        sub: String(r[5] || ""), monto: Number(r[6]) || 0, nota: String(r[7] || ""),
        creado: r[8] instanceof Date ? r[8].toISOString() : String(r[8] || "")
      };
    })
    .filter(function (x) { return !mes || x.mes === mes; });
}

/* Solo deja pasar una fecha real, un rubro con nombre corto y un monto positivo. */
function validar(doc) {
  doc = doc || {};
  const fecha = String(doc.fecha || "");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) throw new Error("Fecha inválida.");
  const rubro = String(doc.rubro || "");
  if (!/^[a-z0-9_-]{2,24}$/.test(rubro)) throw new Error("Rubro inválido.");
  const monto = Math.round(Number(doc.monto) * 100) / 100;
  if (!(monto > 0 && monto < 10000000)) throw new Error("Monto inválido.");
  return [fecha, fecha.slice(0, 7), Number(doc.q) === 1 ? 1 : 2, rubro, limpio(doc.sub, 30), monto, limpio(doc.nota, 500)];
}

function agregar(doc) {
  const v = validar(doc);
  const id = Utilities.getUuid().replace(/-/g, "").slice(0, 16);
  const fila = [id].concat(v).concat([new Date().toISOString(), ""]);
  conCandado(function () { libro().getSheetByName(HOJA_G).appendRow(fila); });
  return id;
}

function filaDe(sh, id) {
  if (sh.getLastRow() < 2) return -1;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) return i + 2;
  return -1;
}

function guardar(id, doc) {
  const v = validar(doc);
  conCandado(function () {
    const sh = libro().getSheetByName(HOJA_G), r = filaDe(sh, id);
    if (r < 0) throw new Error("Ese gasto ya no existe.");
    sh.getRange(r, 2, 1, v.length).setValues([v]);
    sh.getRange(r, COLS.length).setValue(new Date().toISOString());
  });
}

function borrar(ids) {
  ids = (ids || []).map(String);
  conCandado(function () {
    const sh = libro().getSheetByName(HOJA_G);
    if (sh.getLastRow() < 2) return;
    const col = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let r = col.length - 1; r >= 0; r--) if (ids.indexOf(String(col[r][0])) >= 0) sh.deleteRow(r + 2);
  });
}

/* Las filas de la pestaña «Presupuesto», limpias: [tipo, clave, nombre, q1, q2, fijo, subcategorias]. */
function filasP() {
  const sh = libro().getSheetByName(HOJA_P);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues()
    .filter(function (r) { return String(r[0]).trim() && String(r[1]).trim(); })
    .map(function (r) {
      return [sinAcentos(r[0]).trim().toLowerCase(), sinAcentos(r[1]).trim().toLowerCase(), String(r[2]).trim(),
        Number(r[3]) || 0, Number(r[4]) || 0, /^(si|x|1|true)$/i.test(sinAcentos(r[5]).trim()) ? "sí" : "", String(r[6] || "").trim()];
    });
}

function cfgDe(filas) {
  const cfg = { ingreso: { q1: [], q2: [] }, pension: { q1: 0, q2: 0 }, caja: { q1: 0, q2: 0 }, rubros: [] };
  filas.forEach(function (r) {
    const t = r[0], k = r[1], n = r[2], a = r[3], b = r[4];
    if (t === "ingreso") { if (a) cfg.ingreso.q1.push([n, a]); if (b) cfg.ingreso.q2.push([n, b]); }
    else if (t === "pension") cfg.pension = { q1: a, q2: b };
    else if (t === "caja") cfg.caja = { q1: a, q2: b };
    else if (t === "rubro" && /^[a-z0-9_-]{2,24}$/.test(k)) cfg.rubros.push({
      k: k, n: n || k, q1: a, q2: b, fijo: r[5] === "sí",
      subs: String(r[6] || "").split(",").map(function (x) { return x.trim(); }).filter(String)
    });
  });
  return cfg;
}

function presupuesto() { return cfgDe(filasP()); }

/* Reescribe la pestaña «Presupuesto» con lo que se editó en la app. */
function guardarPresupuesto(filas) {
  if (!Array.isArray(filas) || !filas.length || filas.length > 80) throw new Error("Presupuesto inválido.");
  const vistos = {};
  let pen = 0, caja = 0, otros = false;
  const out = filas.map(function (r) {
    r = r || [];
    const t = String(r[0] || ""), k = String(r[1] || "");
    if (["ingreso", "rubro", "pension", "caja"].indexOf(t) < 0) throw new Error("Tipo inválido.");
    if (!/^[a-z0-9_-]{2,24}$/.test(k)) throw new Error("Clave inválida: " + k);
    if (vistos[t + "/" + k]) throw new Error("Rubro repetido: " + k);
    vistos[t + "/" + k] = 1;
    if (t === "pension") pen++;
    if (t === "caja") caja++;
    if (t === "rubro" && k === "otros") otros = true;
    const n = limpio(r[2], 40) || k;
    const a = Math.round(Number(r[3]) || 0), b = Math.round(Number(r[4]) || 0);
    if (a < 0 || b < 0 || a > 9999999 || b > 9999999) throw new Error("Monto inválido en " + n + ".");
    return [t, k, n, a, b, r[5] ? "sí" : "", limpio(r[6], 200)];
  });
  if (pen > 1 || caja > 1) throw new Error("Solo puede haber una pensión y una caja de ahorro.");
  if (!otros) throw new Error("El rubro Otros no se puede quitar.");
  conCandado(function () {
    const sh = libro().getSheetByName(HOJA_P);
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 7).clearContent();
    sh.getRange(2, 1, out.length, 7).setValues(out);
  });
}

/* Fecha del siguiente pago de cada quincena, por ejemplo {"2026-11-2": "2026-11-28"}. */
function pagos() {
  try { return JSON.parse(props().getProperty("PAGOS") || "{}"); } catch (e) { return {}; }
}
function guardarPago(clave, fecha) {
  if (!/^\d{4}-\d{2}-[12]$/.test(clave)) throw new Error("Quincena inválida.");
  if (fecha && !/^\d{4}-\d{2}-\d{2}$/.test(fecha)) throw new Error("Fecha inválida.");
  conCandado(function () {
    const p = pagos();
    if (fecha) p[clave] = fecha; else delete p[clave];
    props().setProperty("PAGOS", JSON.stringify(p));
  });
}

/* Cierre de año: pestaña con [tabla, seccion, clave, concepto, presupuesto, real].
   Tabla 1 = gastos de fin de año; tabla 2 = ingresos. El «real» se escribe desde la app. */
function filasCierre() {
  const sh = libro().getSheetByName(HOJA_C);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, COLS_C.length).getValues()
    .filter(function (r) { return String(r[2]).trim(); })
    .map(function (r) {
      return {
        tabla: Number(r[0]) === 2 ? 2 : 1, seccion: String(r[1]).trim(), clave: String(r[2]).trim(),
        concepto: String(r[3]).trim(), presupuesto: Number(r[4]) || 0,
        real: (r[5] === "" || r[5] === null) ? null : Number(r[5])
      };
    });
}

function guardarRealCierre(clave, real) {
  if (!/^[a-z0-9_-]{2,30}$/.test(clave)) throw new Error("Concepto inválido.");
  let v = "";
  if (real !== "" && real !== null && real !== undefined) {
    v = Math.round(Number(real) * 100) / 100;
    if (!(v >= 0 && v < 100000000)) throw new Error("Monto inválido.");
  }
  conCandado(function () {
    const sh = libro().getSheetByName(HOJA_C);
    if (!sh || sh.getLastRow() < 2) throw new Error("Falta la pestaña «Cierre de año».");
    const col = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues();
    for (let i = 0; i < col.length; i++) {
      if (String(col[i][0]).trim() === clave) { sh.getRange(i + 2, 6).setValue(v); return; }
    }
    throw new Error("Ese concepto ya no existe.");
  });
}

/* Gastos del cierre de año: cada gasto queda en la pestaña «Cierre movimientos»
   y su monto se suma al «real» del concepto (al borrarlo o cambiarlo, se ajusta). */
function hojaMovCierre() {
  const ss = libro();
  let sh = ss.getSheetByName(HOJA_CM);
  if (!sh) {
    sh = ss.insertSheet(HOJA_CM);
    sh.getRange("A:C").setNumberFormat("@");
    sh.getRange(1, 1, 1, COLS_CM.length).setValues([COLS_CM]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function movsCierre() {
  const ss = libro(), sh = ss.getSheetByName(HOJA_CM), tz = ss.getSpreadsheetTimeZone();
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, COLS_CM.length).getValues()
    .filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) {
      return {
        id: String(r[0]), fecha: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, "yyyy-MM-dd") : String(r[1]).slice(0, 10),
        clave: String(r[2]).trim(), monto: Number(r[3]) || 0, nota: String(r[4] || ""),
        creado: r[5] instanceof Date ? r[5].toISOString() : String(r[5] || "")
      };
    });
}

function montoMov(m) {
  const v = Math.round(Number(m) * 100) / 100;
  if (!(v > 0 && v < 100000000)) throw new Error("Monto inválido.");
  return v;
}

/* suma delta al «real» del concepto (dentro del candado); devuelve el real nuevo ("" si queda en 0) */
function sumarReal(clave, delta) {
  const sh = libro().getSheetByName(HOJA_C);
  if (!sh || sh.getLastRow() < 2) throw new Error("Falta la pestaña «Cierre de año».");
  const col = sh.getRange(2, 3, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < col.length; i++) {
    if (String(col[i][0]).trim() !== clave) continue;
    const c = sh.getRange(i + 2, 6), cur = c.getValue();
    let v = Math.round(((cur === "" || cur === null ? 0 : Number(cur) || 0) + delta) * 100) / 100;
    if (!(v > 0.004)) v = "";
    c.setValue(v);
    return v === "" ? null : v;
  }
  throw new Error("Ese concepto ya no existe.");
}

function agregarMovCierre(d) {
  const clave = String(d.clave || "");
  if (!/^[a-z0-9_-]{2,30}$/.test(clave)) throw new Error("Concepto inválido.");
  const monto = montoMov(d.monto);
  const fecha = /^\d{4}-\d{2}-\d{2}$/.test(String(d.fecha || "")) ? String(d.fecha) : Utilities.formatDate(new Date(), libro().getSpreadsheetTimeZone(), "yyyy-MM-dd");
  const id = Utilities.getUuid().replace(/-/g, "").slice(0, 16);
  return conCandado(function () {
    const real = sumarReal(clave, monto);
    hojaMovCierre().appendRow([id, fecha, clave, monto, limpio(d.nota, 500), new Date().toISOString()]);
    return { id: id, clave: clave, real: real };
  });
}

function editarMovCierre(id, d) {
  const monto = montoMov(d.monto);
  return conCandado(function () {
    const sh = hojaMovCierre(), r = filaDe(sh, id);
    if (r < 0) throw new Error("Ese gasto ya no existe.");
    const clave = String(sh.getRange(r, 3).getValue()).trim(), antes = Number(sh.getRange(r, 4).getValue()) || 0;
    const real = sumarReal(clave, monto - antes);
    sh.getRange(r, 4, 1, 2).setValues([[monto, limpio(d.nota, 500)]]);
    return { clave: clave, real: real };
  });
}

function borrarMovCierre(id) {
  return conCandado(function () {
    const sh = hojaMovCierre(), r = filaDe(sh, id);
    if (r < 0) throw new Error("Ese gasto ya no existe.");
    const clave = String(sh.getRange(r, 3).getValue()).trim(), antes = Number(sh.getRange(r, 4).getValue()) || 0;
    let real = null;
    try { real = sumarReal(clave, -antes); } catch (e) { /* el concepto ya no existe: solo se borra el gasto */ }
    sh.deleteRow(r);
    return { clave: clave, real: real };
  });
}

/* «Ajustar presupuesto» del cierre: { clave: monto } → columna «presupuesto» de la pestaña «Cierre de año» */
function guardarPresupuestoCierre(valores) {
  valores = valores || {};
  const vals = {};
  Object.keys(valores).forEach(function (k) {
    if (!/^[a-z0-9_-]{2,30}$/.test(k)) throw new Error("Concepto inválido.");
    const v = Math.round(Number(valores[k]) * 100) / 100;
    if (!(v >= 0 && v < 100000000)) throw new Error("Monto inválido.");
    vals[k] = v;
  });
  conCandado(function () {
    const sh = libro().getSheetByName(HOJA_C);
    if (!sh || sh.getLastRow() < 2) throw new Error("Falta la pestaña «Cierre de año».");
    const n = sh.getLastRow() - 1, claves = sh.getRange(2, 3, n, 1).getValues(), pres = sh.getRange(2, 5, n, 1).getValues();
    for (let i = 0; i < n; i++) { const k = String(claves[i][0]).trim(); if (Object.prototype.hasOwnProperty.call(vals, k)) pres[i][0] = vals[k]; }
    sh.getRange(2, 5, n, 1).setValues(pres);
  });
}

/* Reportes pendientes: gastos que hay que reportar (no cuentan en la quincena).
   «reportado» queda vacío mientras está pendiente; al marcarlo guarda la fecha. */
function hojaRep() {
  const ss = libro();
  let sh = ss.getSheetByName(HOJA_R);
  if (!sh) {
    sh = ss.insertSheet(HOJA_R);
    sh.getRange("A:B").setNumberFormat("@");
    sh.getRange("G:G").setNumberFormat("@");
    sh.getRange(1, 1, 1, COLS_R.length).setValues([COLS_R]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function reportes() {
  const ss = libro(), sh = ss.getSheetByName(HOJA_R), tz = ss.getSpreadsheetTimeZone();
  if (!sh || sh.getLastRow() < 2) return [];
  const f = function (v) { return v instanceof Date ? Utilities.formatDate(v, tz, "yyyy-MM-dd") : String(v || "").slice(0, 10); };
  return sh.getRange(2, 1, sh.getLastRow() - 1, COLS_R.length).getValues()
    .filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) {
      return {
        id: String(r[0]), fecha: f(r[1]), rubro: String(r[2]), sub: String(r[3] || ""), monto: Number(r[4]) || 0,
        nota: String(r[5] || ""), reportado: f(r[6]), creado: r[7] instanceof Date ? r[7].toISOString() : String(r[7] || "")
      };
    });
}

function validarRep(doc) {
  doc = doc || {};
  const rubro = String(doc.rubro || "");
  if (!/^[a-z0-9_-]{2,24}$/.test(rubro)) throw new Error("Rubro inválido.");
  const fecha = /^\d{4}-\d{2}-\d{2}$/.test(String(doc.fecha || "")) ? String(doc.fecha) : Utilities.formatDate(new Date(), libro().getSpreadsheetTimeZone(), "yyyy-MM-dd");
  return [fecha, rubro, limpio(doc.sub, 30), montoMov(doc.monto), limpio(doc.nota, 500)];
}

function agregarRep(doc) {
  const v = validarRep(doc), id = Utilities.getUuid().replace(/-/g, "").slice(0, 16);
  conCandado(function () { hojaRep().appendRow([id].concat(v).concat(["", new Date().toISOString()])); });
  return id;
}

function guardarRep(id, doc) {
  const v = validarRep(doc);
  conCandado(function () {
    const sh = hojaRep(), r = filaDe(sh, id);
    if (r < 0) throw new Error("Ese gasto ya no existe.");
    sh.getRange(r, 2, 1, v.length).setValues([v]);
  });
}

function marcarRep(id, si) {
  const v = si ? Utilities.formatDate(new Date(), libro().getSpreadsheetTimeZone(), "yyyy-MM-dd") : "";
  conCandado(function () {
    const sh = hojaRep(), r = filaDe(sh, id);
    if (r < 0) throw new Error("Ese gasto ya no existe.");
    sh.getRange(r, 7).setValue(v);
  });
  return v;
}

function borrarRep(id) {
  conCandado(function () {
    const sh = hojaRep(), r = filaDe(sh, id);
    if (r >= 0) sh.deleteRow(r);
  });
}
