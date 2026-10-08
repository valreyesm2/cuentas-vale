/**
 * Cuentas de Vale — conecta la app con la hoja «Presupuesto Mensual Vale 2026».
 *
 * La app lee y escribe en la pestaña «Transactions Log» con las mismas columnas de siempre
 * (FECHA, TIPO, CATEGORIA, MONTO, IMPUESTO, MEDIO DE PAGO, MONEDA, DETALLES), así que tus
 * pestañas Mensual, Anual y los desplegables siguen funcionando igual. Las categorías salen
 * de la pestaña «Setup».
 *
 * Solo agrega tres columnas a la derecha de «Transactions Log», que ninguna fórmula usa:
 *   I  ID              (oculta; identifica cada fila para que la app edite la correcta)
 *   J  REEMBOLSO DE    (si un ingreso es el reembolso de un gasto, qué categoría)
 *   K  EXTRAORDINARIO  («Si» si ese gasto es extraordinario)
 * Los ajustes de la app (cupo de la TC, topes, saldos reales) se guardan dentro de este
 * programa, no en la hoja.
 *
 * Instalación (una sola vez), desde la hoja: Extensiones → Apps Script.
 *  1. Borra lo que haya, pega este archivo y guarda.
 *  2. Elige la función «configurar» y toca Ejecutar. Acepta los permisos.
 *     En «Registro de ejecución» aparece tu CLAVE: cópiala.
 *  3. Implementar → Nueva implementación → Aplicación web.
 *     Ejecutar como: Yo. Quién tiene acceso: Cualquier persona. → Implementar. Copia la URL /exec.
 *
 * Sin la clave nadie puede leer ni escribir. Para cambiarla, ejecuta «nuevaClave».
 */

const HOJA_LOG = 'Transactions Log';
const HOJA_SETUP = 'Setup';
const FILA_TITULOS = 4;           // fila con FECHA, TIPO, CATEGORIA…
const PRIMERA_FILA = 5;           // primera fila de movimientos
const COL = { fecha: 1, tipo: 2, categoria: 3, monto: 4, impuesto: 5, medio: 6, moneda: 7, detalle: 8, id: 9, reembolsaA: 10, extra: 11 };
const N_COLS = 11;
const TIPOS = ['Ingreso', 'Gastos Fijos', 'Gastos Variables', 'Ahorros', 'Deudas'];

/* ---------- instalación ---------- */
function configurar() {
  const h = hoja_(HOJA_LOG);
  if (h.getMaxColumns() < N_COLS) h.insertColumnsAfter(h.getMaxColumns(), N_COLS - h.getMaxColumns());
  h.getRange(FILA_TITULOS, COL.id, 1, 3).setValues([['ID', 'REEMBOLSO DE', 'EXTRAORDINARIO']]).setFontWeight('bold');
  asignarIds_(h);
  h.hideColumns(COL.id);
  const props = PropertiesService.getScriptProperties();
  let clave = props.getProperty('CLAVE');
  if (!clave) { clave = generarClave_(); props.setProperty('CLAVE', clave); }
  Logger.log('Tu CLAVE es: ' + clave);
}

function nuevaClave() {
  const clave = generarClave_();
  PropertiesService.getScriptProperties().setProperty('CLAVE', clave);
  Logger.log('Tu nueva CLAVE es: ' + clave);
}

function generarClave_() {
  return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 40);
}

/* ---------- API ---------- */
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return salida_({ ok: false, error: 'formato' }); }
  const clave = PropertiesService.getScriptProperties().getProperty('CLAVE');
  if (!clave || req.clave !== clave) return salida_({ ok: false, error: 'clave' });
  try {
    switch (req.accion) {
      case 'cargar': return conBloqueo_(() => salida_(Object.assign({ ok: true }, cargar_())));
      case 'guardarMovs': return conBloqueo_(() => { const ids = guardarMovs_(req.movs || []); SpreadsheetApp.flush(); return salida_({ ok: true, ids }); });
      case 'borrarMov': return conBloqueo_(() => { const borrado = borrarMov_(req.id); SpreadsheetApp.flush(); return salida_({ ok: true, borrado }); });
      case 'guardarAjustes': return conBloqueo_(() => { guardarJson_('AJUSTES', req.ajustes || {}); return salida_({ ok: true }); });
      case 'guardarSaldo': return conBloqueo_(() => { guardarSaldo_(req.mes, req.COP, req.USD); return salida_({ ok: true }); });
      case 'diagnostico': return salida_({ ok: true, validacion: diagnostico_() });
      default: return salida_({ ok: false, error: 'accion' });
    }
  } catch (err) {
    return salida_({ ok: false, error: String(err && err.message || err) });
  }
}

function doGet() {
  return salida_({ ok: true, app: 'Cuentas de Vale' });
}

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function conBloqueo_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/* ---------- utilidades ---------- */
function hoja_(nombre) {
  const h = SpreadsheetApp.getActive().getSheetByName(nombre);
  if (!h) throw new Error('No encuentro la pestaña «' + nombre + '».');
  return h;
}
function zona_() { return SpreadsheetApp.getActive().getSpreadsheetTimeZone(); }
function fechaTexto_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, zona_(), 'yyyy-MM-dd');
  const s = String(v || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}
function nuevoId_() { return 'r' + Utilities.getUuid().replace(/-/g, '').slice(0, 10); }
function leerJson_(clave, defecto) {
  const v = PropertiesService.getScriptProperties().getProperty(clave);
  if (!v) return defecto;
  try { return JSON.parse(v); } catch (err) { return defecto; }
}
function guardarJson_(clave, valor) { PropertiesService.getScriptProperties().setProperty(clave, JSON.stringify(valor)); }

// Filas de movimientos: [{fila, valores[N_COLS]}], solo las que tienen tipo, categoría o monto.
function filas_(h) {
  const ultima = h.getLastRow();
  if (ultima < PRIMERA_FILA) return [];
  const vals = h.getRange(PRIMERA_FILA, 1, ultima - PRIMERA_FILA + 1, N_COLS).getValues();
  const out = [];
  vals.forEach((v, i) => { if (v[COL.tipo - 1] !== '' || v[COL.categoria - 1] !== '' || v[COL.monto - 1] !== '') out.push({ fila: PRIMERA_FILA + i, v }); });
  return out;
}

function asignarIds_(h) {
  filas_(h).forEach(r => { if (!r.v[COL.id - 1]) h.getRange(r.fila, COL.id).setValue(nuevoId_()); });
}

function categorias_() {
  const h = hoja_(HOJA_SETUP);
  const n = h.getLastRow();
  const titulos = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0].map(String);
  const vals = n > 1 ? h.getRange(2, 1, n - 1, titulos.length).getValues() : [];
  const out = {};
  TIPOS.forEach(t => {
    const c = titulos.indexOf(t);
    out[t] = c < 0 ? [] : vals.map(r => String(r[c]).trim()).filter(Boolean);
  });
  return out;
}

/* ---------- acciones ---------- */
function cargar_() {
  const h = hoja_(HOJA_LOG);
  if (PropertiesService.getScriptProperties().getProperty('VALIDACION_REPARADA') !== '2') repararValidacionCategoria();
  asignarIds_(h);
  const movimientos = filas_(h).map(r => {
    const v = r.v;
    const t = {
      id: String(v[COL.id - 1]),
      fecha: fechaTexto_(v[COL.fecha - 1]),
      tipo: String(v[COL.tipo - 1]).trim(),
      categoria: String(v[COL.categoria - 1]).trim(),
      monto: Number(v[COL.monto - 1]) || 0,
      impuesto: String(v[COL.impuesto - 1]).trim() || 'No Aplica',
      medio: String(v[COL.medio - 1]).trim(),
      moneda: String(v[COL.moneda - 1]).trim(),
      detalle: String(v[COL.detalle - 1]).trim(),
      creado: String(r.fila).padStart(6, '0'),
    };
    const re = String(v[COL.reembolsaA - 1]).trim();
    if (re) t.reembolsaA = re;
    if (String(v[COL.extra - 1]).trim().toLowerCase() === 'si') t.extra = true;
    return t;
  });
  return { movimientos, categorias: categorias_(), ajustes: leerJson_('AJUSTES', {}), saldos: leerJson_('SALDOS', {}) };
}

function guardarMovs_(movs) {
  const h = hoja_(HOJA_LOG);
  const filas = filas_(h);
  const porId = {};
  filas.forEach(r => { porId[String(r.v[COL.id - 1])] = r; });
  const ids = [];
  movs.forEach(t => {
    if (!t || !/^\d{4}-\d{2}-\d{2}$/.test(String(t.fecha))) throw new Error('Falta la fecha');
    const fecha = Utilities.parseDate(t.fecha, zona_(), 'yyyy-MM-dd');
    const valores = [fecha, t.tipo, t.categoria, Number(t.monto) || 0, t.impuesto || 'No Aplica', t.medio || 'NA', t.moneda, t.detalle || ''];
    const extras = [t.reembolsaA || '', t.extra === true ? 'Si' : ''];
    const actual = porId[String(t.id)];
    const conservarMonto = actual && Math.abs((Number(actual.v[COL.monto - 1]) || 0) - valores[3]) < 0.005;
    const fila = actual ? actual.fila : Math.max(h.getLastRow() + 1, PRIMERA_FILA);
    const id = actual ? String(t.id) : (t.id && !porId[String(t.id)] ? String(t.id) : nuevoId_());
    escribirFila_(h, fila, valores, id, extras, conservarMonto);
    porId[id] = { fila, v: valores.concat([id]).concat(extras) };
    ids.push(id);
  });
  return ids;
}

// Escribe una fila respetando la hoja: el TIPO va primero para que la lista desplegable de
// CATEGORIA (pestaña DROPDOWN) se calcule antes de poner la categoría.
function escribirFila_(h, fila, valores, id, extras, conservarMonto) {
  h.getRange(fila, COL.fecha, 1, 2).setValues([valores.slice(0, 2)]);
  if (!conservarMonto) h.getRange(fila, COL.monto).setValue(valores[3]); // si no cambió, se conserva la fórmula
  h.getRange(fila, COL.impuesto, 1, 4).setValues([valores.slice(4, 8)]);
  h.getRange(fila, COL.id, 1, 3).setValues([[id].concat(extras)]);
  SpreadsheetApp.flush();
  ponerCategoria_(h, fila, valores[2]);
}

// La lista desplegable de CATEGORIA es dependiente: cada fila mira su propia fila de la pestaña
// DROPDOWN, que Google recalcula con retraso, y la regla está en «rechazar». Por eso la categoría
// se escribe sin regla y luego se le copia la validación de otra fila, como al copiar y pegar a
// mano: Google ajusta la referencia a la fila correcta y el desplegable queda igual que en el resto.
function ponerCategoria_(h, fila, valor) {
  const celda = h.getRange(fila, COL.categoria);
  if (String(celda.getValue()) === String(valor)) return;
  celda.clearDataValidations();
  celda.setValue(valor);
  copiarValidacion_(h, fila, 1);
}

function copiarValidacion_(h, fila, n) {
  const origen = fila === PRIMERA_FILA ? PRIMERA_FILA + 1 : PRIMERA_FILA;
  h.getRange(origen, COL.categoria).copyTo(h.getRange(fila, COL.categoria, n, 1), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
}

// Deja toda la columna CATEGORIA con la misma validación de la primera fila (se ejecuta sola una vez;
// también la puedes ejecutar a mano si alguna celda queda con el desplegable raro).
function repararValidacionCategoria() {
  const h = hoja_(HOJA_LOG);
  const n = h.getMaxRows() - PRIMERA_FILA;
  if (n > 0) copiarValidacion_(h, PRIMERA_FILA + 1, n);
  PropertiesService.getScriptProperties().setProperty('VALIDACION_REPARADA', '2');
}

function borrarMov_(id) {
  const h = hoja_(HOJA_LOG);
  const r = filas_(h).find(x => String(x.v[COL.id - 1]) === String(id));
  if (!r) return false;
  h.deleteRow(r.fila);
  return true;
}

// Solo lectura: qué lista desplegable tiene la celda CATEGORIA de la primera fila y de las últimas.
function diagnostico_() {
  const h = hoja_(HOJA_LOG);
  const ult = h.getLastRow();
  return [PRIMERA_FILA, PRIMERA_FILA + 1, ult - 2, ult - 1, ult].filter((f, i, a) => f >= PRIMERA_FILA && a.indexOf(f) === i).map(f => {
    const v = h.getRange(f, COL.categoria).getDataValidation();
    if (!v) return { fila: f, regla: null };
    const args = v.getCriteriaValues().map(a => a && a.getA1Notation ? a.getSheet().getName() + '!' + a.getA1Notation() : String(a));
    return { fila: f, valor: String(h.getRange(f, COL.categoria).getValue()), tipo: String(v.getCriteriaType()), lista: args, rechaza: !v.getAllowInvalid() };
  });
}

function guardarSaldo_(mes, cop, usd) {
  if (!/^\d{4}-\d{2}$/.test(String(mes))) throw new Error('mes inválido');
  const saldos = leerJson_('SALDOS', {});
  const s = {};
  if (cop !== null && cop !== undefined && cop !== '') s.COP = Number(cop);
  if (usd !== null && usd !== undefined && usd !== '') s.USD = Number(usd);
  if (Object.keys(s).length) saldos[mes] = s; else delete saldos[mes];
  guardarJson_('SALDOS', saldos);
}
