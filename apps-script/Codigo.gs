/**
 * Cuentas de Vale — API sobre Google Sheets.
 *
 * Instalación (una sola vez):
 *  1. En tu hoja de Google: Extensiones → Apps Script. Borra lo que haya y pega este archivo.
 *  2. Elige la función «configurar» arriba y toca Ejecutar. Acepta los permisos.
 *     En «Registro de ejecución» aparece tu CLAVE: cópiala.
 *  3. Implementar → Nueva implementación → Tipo: Aplicación web.
 *     Ejecutar como: Yo. Quién tiene acceso: Cualquier persona. → Implementar.
 *     Copia la URL que termina en /exec.
 *  4. En la app escribe esa URL y la clave.
 *
 * Sin la clave nadie puede leer ni escribir. Si alguna vez la quieres cambiar,
 * ejecuta «nuevaClave» y vuelve a conectar tus dispositivos.
 */

const HOJA_MOV = 'Movimientos';
const HOJA_AJUSTES = 'Ajustes';
const HOJA_SALDOS = 'Saldos';
const CAMPOS = ['id', 'fecha', 'tipo', 'categoria', 'monto', 'impuesto', 'medio', 'moneda', 'detalle', 'reembolsaA', 'reembolsaTipo', 'extra', 'creado'];
const TITULOS = ['ID', 'FECHA', 'TIPO', 'CATEGORIA', 'MONTO', 'IMPUESTO', 'MEDIO DE PAGO', 'MONEDA', 'DETALLES', 'REEMBOLSO DE', 'REEMBOLSO TIPO', 'EXTRAORDINARIO', 'CREADO'];

/* ---------- instalación ---------- */
function configurar() {
  const ss = SpreadsheetApp.getActive();
  const mov = prepararHoja_(ss, HOJA_MOV, TITULOS);
  mov.getRange('A:D').setNumberFormat('@');
  mov.getRange('F:M').setNumberFormat('@');
  mov.getRange('E:E').setNumberFormat('#,##0.00');
  const aj = prepararHoja_(ss, HOJA_AJUSTES, ['CLAVE', 'VALOR']);
  aj.getRange('A:B').setNumberFormat('@');
  const sal = prepararHoja_(ss, HOJA_SALDOS, ['MES', 'COP', 'USD']);
  sal.getRange('A:A').setNumberFormat('@');
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

function prepararHoja_(ss, nombre, titulos) {
  const h = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  h.getRange(1, 1, 1, titulos.length).setValues([titulos]).setFontWeight('bold');
  h.setFrozenRows(1);
  return h;
}

/* ---------- API ---------- */
function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return salida_({ ok: false, error: 'formato' }); }
  const clave = PropertiesService.getScriptProperties().getProperty('CLAVE');
  if (!clave || req.clave !== clave) return salida_({ ok: false, error: 'clave' });
  try {
    switch (req.accion) {
      case 'cargar': return salida_(Object.assign({ ok: true }, cargar_()));
      case 'guardarMovs': return conBloqueo_(() => salida_({ ok: true, guardados: guardarMovs_(req.movs || []) }));
      case 'borrarMov': return conBloqueo_(() => salida_({ ok: true, borrado: borrarMov_(req.id) }));
      case 'guardarAjustes': return conBloqueo_(() => { guardarAjustes_(req.ajustes || {}); return salida_({ ok: true }); });
      case 'guardarSaldo': return conBloqueo_(() => { guardarSaldo_(req.mes, req.COP, req.USD); return salida_({ ok: true }); });
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

function hoja_(nombre) {
  const h = SpreadsheetApp.getActive().getSheetByName(nombre);
  if (!h) throw new Error('Falta la hoja «' + nombre + '». Ejecuta «configurar».');
  return h;
}

function texto_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  return v === null || v === undefined ? '' : String(v);
}

function cargar_() {
  const filas = filas_(hoja_(HOJA_MOV));
  const movimientos = filas.filter(f => f[0] !== '').map(f => {
    const t = {};
    CAMPOS.forEach((c, i) => {
      const v = f[i];
      if (c === 'monto') t.monto = Number(v) || 0;
      else if (c === 'extra') { if (String(v).toLowerCase() === 'si') t.extra = true; }
      else if (v !== '' && v !== null) t[c] = texto_(v);
    });
    return t;
  });
  let ajustes = {};
  filas_(hoja_(HOJA_AJUSTES)).forEach(f => { if (f[0] === 'config' && f[1]) { try { ajustes = JSON.parse(f[1]); } catch (err) {} } });
  const saldos = {};
  filas_(hoja_(HOJA_SALDOS)).forEach(f => {
    const mes = texto_(f[0]).slice(0, 7);
    if (!mes) return;
    const s = {};
    if (f[1] !== '' && f[1] !== null) s.COP = Number(f[1]);
    if (f[2] !== '' && f[2] !== null) s.USD = Number(f[2]);
    saldos[mes] = s;
  });
  return { movimientos, ajustes, saldos };
}

function filas_(h) {
  const n = h.getLastRow() - 1;
  if (n < 1) return [];
  return h.getRange(2, 1, n, h.getLastColumn()).getValues();
}

function fila_(t) {
  return CAMPOS.map(c => {
    if (c === 'monto') return Number(t.monto) || 0;
    if (c === 'extra') return t.extra === true ? 'Si' : '';
    return t[c] === undefined || t[c] === null ? '' : String(t[c]);
  });
}

function guardarMovs_(movs) {
  const h = hoja_(HOJA_MOV);
  const ids = filas_(h).map(f => String(f[0]));
  const nuevos = [];
  movs.forEach(t => {
    if (!t || !t.id || !/^\d{4}-\d{2}-\d{2}$/.test(String(t.fecha))) return;
    const i = ids.indexOf(String(t.id));
    if (i >= 0) h.getRange(i + 2, 1, 1, CAMPOS.length).setValues([fila_(t)]);
    else { nuevos.push(fila_(t)); ids.push(String(t.id)); }
  });
  if (nuevos.length) h.getRange(h.getLastRow() + 1, 1, nuevos.length, CAMPOS.length).setValues(nuevos);
  return movs.length;
}

function borrarMov_(id) {
  const h = hoja_(HOJA_MOV);
  const i = filas_(h).findIndex(f => String(f[0]) === String(id));
  if (i < 0) return false;
  h.deleteRow(i + 2);
  return true;
}

function guardarAjustes_(ajustes) {
  const h = hoja_(HOJA_AJUSTES);
  const i = filas_(h).findIndex(f => f[0] === 'config');
  const valor = JSON.stringify(ajustes);
  if (i >= 0) h.getRange(i + 2, 2).setValue(valor);
  else h.appendRow(['config', valor]);
}

function guardarSaldo_(mes, cop, usd) {
  if (!/^\d{4}-\d{2}$/.test(String(mes))) throw new Error('mes inválido');
  const h = hoja_(HOJA_SALDOS);
  const i = filas_(h).findIndex(f => texto_(f[0]).slice(0, 7) === mes);
  const vacio = (cop === null || cop === undefined) && (usd === null || usd === undefined);
  if (vacio) { if (i >= 0) h.deleteRow(i + 2); return; }
  const fila = [mes, cop === null || cop === undefined ? '' : Number(cop), usd === null || usd === undefined ? '' : Number(usd)];
  if (i >= 0) h.getRange(i + 2, 1, 1, 3).setValues([fila]);
  else h.appendRow(fila);
}
