/**
 * La quincena — el "cerebro" que vive en Google Apps Script.
 * Guarda los gastos en una hoja de Google privada y lee el presupuesto de la pestaña «Presupuesto».
 * La página pública no trae ningún número: todo pasa por aquí, y solo con la clave correcta.
 *
 * Instalación (una sola vez):
 *  1. script.google.com → Proyecto nuevo → pegar este archivo.
 *  2. Poner los montos en PRESUPUESTO y ejecutar `setup`: crea la hoja, genera una clave al azar
 *     y deja tu enlace personal (con la clave) en la pestaña «Tu enlace».
 *  3. Implementar → Nueva implementación → Aplicación web
 *     · Ejecutar como: yo   · Quién tiene acceso: cualquier persona
 *  4. Pegar la URL que termina en /exec en API_URL de index.html.
 * Después, el presupuesto se cambia directo en la pestaña «Presupuesto» de la hoja.
 * Para cambiar la clave: borrar la propiedad CLAVE del script y volver a ejecutar `setup`.
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
      case "load": out = { ok: true, cfg: presupuesto(), gastos: leer(String(d.mes || "")) }; break;
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

function presupuesto() {
  const sh = libro().getSheetByName(HOJA_P);
  const cfg = { ingreso: { q1: [], q2: [] }, pension: { q1: 0, q2: 0 }, caja: { q1: 0, q2: 0 }, rubros: [] };
  if (!sh || sh.getLastRow() < 2) return cfg;
  sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues().forEach(function (r) {
    const t = sinAcentos(r[0]).trim().toLowerCase(), k = sinAcentos(r[1]).trim().toLowerCase(), n = String(r[2]).trim();
    const a = Number(r[3]) || 0, b = Number(r[4]) || 0;
    if (t === "ingreso") { if (a) cfg.ingreso.q1.push([n, a]); if (b) cfg.ingreso.q2.push([n, b]); }
    else if (t === "pension") cfg.pension = { q1: a, q2: b };
    else if (t === "caja") cfg.caja = { q1: a, q2: b };
    else if (t === "rubro" && /^[a-z0-9_-]{2,24}$/.test(k)) cfg.rubros.push({
      k: k, n: n || k, q1: a, q2: b,
      fijo: /^(si|x|1|true)$/i.test(sinAcentos(r[5]).trim()),
      subs: String(r[6] || "").split(",").map(function (s) { return s.trim(); }).filter(String)
    });
  });
  return cfg;
}
