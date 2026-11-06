/**
 * Finanzas · Vicente — backend en Google Apps Script.
 *
 * Guarda los movimientos de los bancos (Santander, Banco Falabella) en la planilla
 * "Finanzas" y los entrega a la app (index.html). Los movimientos llegan como CSV
 * a la carpeta de Drive "Finanzas – importar": Claude convierte las cartolas a ese
 * formato y las deja ahí; este script las carga sin duplicar (cada hora, o con el
 * botón "Importar ahora" de la app).
 *
 * Instalación: ver SETUP.md.
 */

// Se muestra al abrir la URL /exec: sirve para comprobar que Google tiene el código nuevo.
var VERSION = '2026-11-06 · primera versión';

var HOJA_MOV = 'Movimientos';
var HOJA_REGLAS = 'Reglas';
var HOJA_PRESUPUESTO = 'Presupuesto';
var HOJA_TARJETAS = 'Tarjetas';
var HOJA_CENTROS = 'Centros';
var HOJA_CUENTAS = 'Cuentas';
var HOJA_IMPORTACIONES = 'Importaciones';

var COLUMNAS_MOV = ['ID', 'Fecha', 'Banco', 'Producto', 'Descripción', 'Monto', 'Tipo',
  'Centro de costo', 'Cuenta', 'Cuotas', 'Nota', 'Clave', 'Origen', 'Creado'];
var COLUMNAS_REGLAS = ['Contiene', 'Tipo', 'Centro de costo', 'Cuenta', 'Creado'];
var COLUMNAS_PRESUPUESTO = ['Centro de costo', 'Cuenta', 'Tope mensual'];
var COLUMNAS_TARJETAS = ['Producto', 'Fecha estado', 'Deuda total', 'Monto a pagar', 'Fecha pago', 'Cupo total', 'Origen', 'Creado'];
var COLUMNAS_IMPORTACIONES = ['Archivo', 'Nuevos', 'Repetidos', 'Con error', 'Detalle', 'Creado'];

// Tipo de cada movimiento. Solo "Gasto" e "Ingreso" cuentan en los resúmenes:
// el pago de la tarjeta y los traspasos entre tus cuentas no son gasto (si no, se contaría doble).
var TIPOS = ['Gasto', 'Ingreso', 'Pago tarjeta', 'Traspaso', 'Ignorar'];

// Centros y cuentas: se copian de la planilla de Casa App (propiedad CASA_SHEET_ID).
// Si no está configurada, se usa esta lista base (la misma de Casa App).
var CENTROS_INICIALES = ['Auto', 'Familia', 'Salud', 'Vicky', 'Fran', 'Vicente', 'Casa'];
var CUENTAS_INICIALES = {
  'Auto': ['Seguro', 'Bencina', 'TAG'],
  'Familia': ['Regalos', 'Eventos / Salidas', 'Supermercado', 'Luna'],
  'Salud': ['Isapre', 'Farmacia'],
  'Vicky': ['Sueldo', 'Traslado', 'Imposiciones'],
  'Casa': ['Fumigación', 'Conectividad', 'Gas', 'Agua', 'Tintorería']
};
// Los ingresos tienen su propio centro (no viene de Casa App).
var CENTRO_INGRESOS = 'Ingresos';
var CUENTAS_INGRESOS = ['Sueldo', 'Bonos', 'Reembolsos', 'Intereses', 'Otros ingresos'];

var CARPETA_IMPORTAR = 'Finanzas – importar';
var CARPETA_PROCESADAS = 'Procesadas';
var MESES_EN_APP = 24;   // la app recibe los últimos 24 meses

// ---------------------------------------------------------------------------
// Instalación: ejecutar una vez desde el editor (botón "Ejecutar" con setup).
// ---------------------------------------------------------------------------

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty('SHEET_ID');
    ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('Finanzas');
    props.setProperty('SHEET_ID', ss.getId());
  }
  ss.setSpreadsheetTimeZone('America/Santiago');

  prepararHoja_(ss, HOJA_MOV, COLUMNAS_MOV);
  prepararHoja_(ss, HOJA_REGLAS, COLUMNAS_REGLAS);
  prepararHoja_(ss, HOJA_PRESUPUESTO, COLUMNAS_PRESUPUESTO);
  prepararHoja_(ss, HOJA_TARJETAS, COLUMNAS_TARJETAS);
  prepararHoja_(ss, HOJA_IMPORTACIONES, COLUMNAS_IMPORTACIONES);

  var centros = prepararHoja_(ss, HOJA_CENTROS, ['Centro de costo']);
  var cuentas = prepararHoja_(ss, HOJA_CUENTAS, ['Centro de costo', 'Cuenta']);
  if (centros.getLastRow() < 2) {
    centros.getRange(2, 1, CENTROS_INICIALES.length, 1).setValues(CENTROS_INICIALES.map(function (c) { return [c]; }));
  }
  if (cuentas.getLastRow() < 2) {
    var filas = [];
    Object.keys(CUENTAS_INICIALES).forEach(function (c) {
      CUENTAS_INICIALES[c].forEach(function (x) { filas.push([c, x]); });
    });
    cuentas.getRange(2, 1, filas.length, 2).setValues(filas);
  }

  var hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && hoja1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(hoja1);

  // Pide (una vez) los permisos de Drive y crea la carpeta de importación.
  var carpeta = carpetaImportar_();
  activarImportacionAutomatica();
  sincronizarCentros_(true);
  darFormato();

  Logger.log('Listo.\nPlanilla: ' + ss.getUrl() + '\nCarpeta para importar: ' + carpeta.getUrl());
}

// Revisa la carpeta de importación cada hora (se instala con setup; se puede volver a ejecutar).
function activarImportacionAutomatica() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'importarPendientes') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('importarPendientes').timeBased().everyHours(1).create();
  Logger.log('Importación automática activada: cada hora.');
}

// ---------------------------------------------------------------------------
// Hojas
// ---------------------------------------------------------------------------

function libro_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Falta ejecutar setup() en el editor de Apps Script');
  return SpreadsheetApp.openById(id);
}

// Crea la hoja con sus encabezados si no existe, y agrega al final las columnas que falten.
function prepararHoja_(ss, nombre, columnas) {
  var hoja = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  if (hoja.getLastRow() === 0) {
    hoja.getRange(1, 1, 1, columnas.length).setValues([columnas]).setFontWeight('bold');
    hoja.setFrozenRows(1);
    return hoja;
  }
  var enc = encabezados_(hoja);
  columnas.forEach(function (c) {
    if (enc.indexOf(c) === -1) {
      hoja.getRange(1, hoja.getLastColumn() + 1).setValue(c).setFontWeight('bold');
      enc.push(c);
    }
  });
  return hoja;
}

function hoja_(nombre, columnas) {
  var ss = libro_();
  var hoja = ss.getSheetByName(nombre);
  if (!hoja) return prepararHoja_(ss, nombre, columnas);
  return hoja;
}

function encabezados_(hoja) {
  if (hoja.getLastColumn() === 0) return [];
  return hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0].map(function (v) { return String(v).trim(); });
}

// Lee la hoja como objetos { 'Columna': valor, _fila: n }. Las columnas se buscan por su
// encabezado, así que se pueden mover en la planilla sin romper nada.
function leerTabla_(hoja) {
  var enc = encabezados_(hoja);
  if (hoja.getLastRow() < 2) return { enc: enc, filas: [] };
  var valores = hoja.getRange(2, 1, hoja.getLastRow() - 1, enc.length).getValues();
  var filas = valores.map(function (v, i) {
    var o = { _fila: i + 2 };
    enc.forEach(function (c, j) { if (c) o[c] = v[j]; });
    return o;
  });
  return { enc: enc, filas: filas };
}

// Convierte objetos a filas en el orden de las columnas de la hoja.
function aFilas_(enc, objetos) {
  return objetos.map(function (o) {
    return enc.map(function (c) { return o[c] === undefined ? '' : o[c]; });
  });
}

function agregarFilas_(hoja, objetos) {
  if (!objetos.length) return;
  var enc = encabezados_(hoja);
  hoja.getRange(hoja.getLastRow() + 1, 1, objetos.length, enc.length).setValues(aFilas_(enc, objetos));
}

// Escribe solo algunas columnas de una fila.
function escribirCampos_(hoja, enc, fila, campos) {
  Object.keys(campos).forEach(function (c) {
    var col = enc.indexOf(c);
    if (col !== -1) hoja.getRange(fila, col + 1).setValue(campos[c]);
  });
}

function leerLista_(nombreHoja) {
  var hoja = libro_().getSheetByName(nombreHoja);
  if (!hoja || hoja.getLastRow() < 2) return [];
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues()
    .map(function (f) { return String(f[0]).trim(); })
    .filter(function (v) { return v; });
}

function alfabetico_(lista) {
  return lista.slice().sort(function (a, b) { return String(a).localeCompare(String(b), 'es', { sensitivity: 'base' }); });
}

function sinTildes_(t) {
  return String(t || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/\s+/g, ' ').trim();
}

function zona_() { return libro_().getSpreadsheetTimeZone() || 'America/Santiago'; }

function isoFecha_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz || zona_(), 'yyyy-MM-dd');
  return String(v || '').slice(0, 10);
}

// 'yyyy-MM-dd' o 'dd/MM/yyyy' (o 'dd-MM-yyyy') → Date a mediodía (sin líos de zona horaria).
function parsearFecha_(texto) {
  var t = String(texto || '').trim();
  var m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12);
  m = t.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
  if (m) {
    var anio = Number(m[3]) < 100 ? 2000 + Number(m[3]) : Number(m[3]);
    return new Date(anio, Number(m[2]) - 1, Number(m[1]), 12);
  }
  return null;
}

// "$ -12.345" / "-12345" / "12.345,00" → número entero de pesos.
function parsearMonto_(v) {
  if (typeof v === 'number') return Math.round(v);
  var t = String(v || '').replace(/\s|\$/g, '');
  if (!t) return NaN;
  var negativo = /^-|^\(.*\)$|-$/.test(t);
  t = t.replace(/[()\-+]/g, '').replace(/,\d{1,2}$/, '').replace(/[.,]/g, '');
  var n = Number(t);
  return isNaN(n) ? NaN : (negativo ? -n : n);
}

// ---------------------------------------------------------------------------
// Centros de costo y cuentas (copiados desde Casa App)
// ---------------------------------------------------------------------------

// Si existe la propiedad CASA_SHEET_ID, copia las hojas Centros y Cuentas de Casa App
// a esta planilla (como mucho cada 10 minutos). Si falla, se sigue con lo que hay.
function sincronizarCentros_(forzar) {
  var casaId = propiedad_('CASA_SHEET_ID');
  if (!casaId) return '';
  var cache = CacheService.getScriptCache();
  if (!forzar && cache.get('centrosSync')) return '';
  try {
    var casa = SpreadsheetApp.openById(casaId);
    var hc = casa.getSheetByName('Centros');
    var hq = casa.getSheetByName('Cuentas');
    if (!hc) throw new Error('la planilla de Casa App no tiene la hoja Centros');
    var centros = hc.getLastRow() < 2 ? [] : hc.getRange(2, 1, hc.getLastRow() - 1, 1).getValues()
      .map(function (f) { return [String(f[0]).trim()]; }).filter(function (f) { return f[0]; });
    var cuentas = (!hq || hq.getLastRow() < 2) ? [] : hq.getRange(2, 1, hq.getLastRow() - 1, 2).getValues()
      .map(function (f) { return [String(f[0]).trim(), String(f[1]).trim()]; }).filter(function (f) { return f[0] && f[1]; });
    if (!centros.length) throw new Error('la hoja Centros de Casa App está vacía');
    var ss = libro_();
    var c = prepararHoja_(ss, HOJA_CENTROS, ['Centro de costo']);
    var q = prepararHoja_(ss, HOJA_CUENTAS, ['Centro de costo', 'Cuenta']);
    if (c.getLastRow() > 1) c.getRange(2, 1, c.getLastRow() - 1, c.getLastColumn()).clearContent();
    if (q.getLastRow() > 1) q.getRange(2, 1, q.getLastRow() - 1, q.getLastColumn()).clearContent();
    c.getRange(2, 1, centros.length, 1).setValues(centros);
    if (cuentas.length) q.getRange(2, 1, cuentas.length, 2).setValues(cuentas);
    cache.put('centrosSync', '1', 600);
    return '';
  } catch (err) {
    cache.put('centrosSync', '1', 600);
    Logger.log('No se pudieron copiar los centros de Casa App: ' + err.message);
    return 'No se pudieron leer los centros de Casa App (' + err.message + '). Se usa la última copia.';
  }
}

// Ejecutar a mano después de poner CASA_SHEET_ID, para copiar los centros al tiro.
function copiarCentrosDeCasaApp() {
  var aviso = sincronizarCentros_(true);
  Logger.log(aviso || 'Listo: centros y cuentas copiados desde Casa App.');
}

function categorias_() {
  var centros = leerLista_(HOJA_CENTROS).filter(function (c) { return c !== CENTRO_INGRESOS; });
  var cuentas = {};
  var hoja = libro_().getSheetByName(HOJA_CUENTAS);
  if (hoja && hoja.getLastRow() > 1) {
    hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues().forEach(function (f) {
      var c = String(f[0]).trim(), x = String(f[1]).trim();
      if (!c || !x) return;
      (cuentas[c] = cuentas[c] || []).push(x);
    });
  }
  Object.keys(cuentas).forEach(function (c) { cuentas[c] = alfabetico_(cuentas[c]); });
  cuentas[CENTRO_INGRESOS] = CUENTAS_INGRESOS.slice();
  return { centros: alfabetico_(centros), cuentas: cuentas };
}

function validarCategoria_(tipo, centro, cuenta) {
  if (TIPOS.indexOf(tipo) === -1) throw new Error('Tipo no válido: ' + tipo);
  if (tipo !== 'Gasto' && tipo !== 'Ingreso') return { centro: '', cuenta: '' };
  var cat = categorias_();
  if (tipo === 'Ingreso') {
    if (cuenta && CUENTAS_INGRESOS.indexOf(cuenta) === -1) throw new Error('Esa categoría de ingreso no existe');
    return { centro: CENTRO_INGRESOS, cuenta: cuenta || '' };
  }
  if (!centro) return { centro: '', cuenta: '' };
  if (cat.centros.indexOf(centro) === -1) throw new Error('El centro "' + centro + '" no existe. Revisa los centros en Casa App.');
  if (cuenta && (cat.cuentas[centro] || []).indexOf(cuenta) === -1) throw new Error('La cuenta "' + cuenta + '" no existe en ' + centro);
  return { centro: centro, cuenta: cuenta || '' };
}

// ---------------------------------------------------------------------------
// Reglas: "si la descripción contiene X → tipo / centro / cuenta"
// ---------------------------------------------------------------------------

function leerReglas_() {
  var t = leerTabla_(hoja_(HOJA_REGLAS, COLUMNAS_REGLAS));
  return t.filas.map(function (r) {
    return { contiene: sinTildes_(r['Contiene']), tipo: String(r['Tipo'] || '').trim(),
      centro: String(r['Centro de costo'] || '').trim(), cuenta: String(r['Cuenta'] || '').trim() };
  }).filter(function (r) { return r.contiene; })
    // La regla más específica (texto más largo) gana.
    .sort(function (a, b) { return b.contiene.length - a.contiene.length; });
}

function reglaPara_(descripcion, reglas) {
  var d = sinTildes_(descripcion);
  for (var i = 0; i < reglas.length; i++) if (d.indexOf(reglas[i].contiene) !== -1) return reglas[i];
  return null;
}

function esTarjeta_(producto) { return /tarjeta|cmr|visa|master|amex|\btc\b|credito/i.test(sinTildes_(producto)); }

// Tipo automático cuando la cartola no lo trae y ninguna regla calza.
function tipoAutomatico_(producto, descripcion, monto) {
  var d = sinTildes_(descripcion);
  if (/PAGO.*(TARJETA|TARJ|\bTC\b|CMR|VISA|MASTER|MONTO FACTURADO)|PAGO TOTAL|PAGO MINIMO|PAGO DE TARJETA|PAGO EN LINEA TC|ABONO TARJETA/.test(d)) return 'Pago tarjeta';
  if (/TRASPASO|TRANSFERENCIA ENTRE MIS CUENTAS|TRANSF\. ENTRE CTAS|ENTRE CUENTAS PROPIAS/.test(d)) return 'Traspaso';
  if (esTarjeta_(producto)) {
    // En la tarjeta un abono que no es pago es una devolución: resta al gasto.
    return 'Gasto';
  }
  return monto > 0 ? 'Ingreso' : 'Gasto';
}

function guardarRegla_(contiene, tipo, centro, cuenta) {
  contiene = sinTildes_(contiene).slice(0, 60);
  if (contiene.length < 3) throw new Error('El texto de la regla es muy corto (mínimo 3 letras)');
  var hoja = hoja_(HOJA_REGLAS, COLUMNAS_REGLAS);
  var t = leerTabla_(hoja);
  var existente = t.filas.filter(function (r) { return sinTildes_(r['Contiene']) === contiene; })[0];
  var datos = { 'Contiene': contiene, 'Tipo': tipo, 'Centro de costo': centro, 'Cuenta': cuenta, 'Creado': new Date() };
  if (existente) escribirCampos_(hoja, t.enc, existente._fila, datos);
  else agregarFilas_(hoja, [datos]);
}

// Aplica las reglas a los movimientos sin categoría (nunca cambia lo que ya categorizaste).
function aplicarReglasPendientes_() {
  var reglas = leerReglas_();
  if (!reglas.length) return 0;
  var hoja = hoja_(HOJA_MOV, COLUMNAS_MOV);
  var t = leerTabla_(hoja);
  var cTipo = t.enc.indexOf('Tipo'), cCentro = t.enc.indexOf('Centro de costo'), cCuenta = t.enc.indexOf('Cuenta');
  var n = 0;
  t.filas.forEach(function (m) {
    if (m['Centro de costo'] || m['Tipo'] === 'Ignorar' || m['Tipo'] === 'Pago tarjeta' || m['Tipo'] === 'Traspaso') return;
    var r = reglaPara_(m['Descripción'], reglas);
    if (!r) return;
    var tipo = r.tipo || m['Tipo'] || 'Gasto';
    hoja.getRange(m._fila, cTipo + 1).setValue(tipo);
    if (r.centro) hoja.getRange(m._fila, cCentro + 1).setValue(r.centro);
    if (r.cuenta) hoja.getRange(m._fila, cCuenta + 1).setValue(r.cuenta);
    n++;
  });
  return n;
}

// Ejecutar a mano si agregaste reglas directo en la planilla.
function aplicarReglas() {
  Logger.log('Movimientos categorizados: ' + aplicarReglasPendientes_());
}

// ---------------------------------------------------------------------------
// Importar: CSV en la carpeta "Finanzas – importar" de Drive.
//
// Movimientos (una fila por movimiento, la primera fila con estos títulos):
//   fecha, banco, producto, descripcion, monto [, cuotas, tipo, centro, cuenta, nota]
//   - fecha: yyyy-mm-dd o dd/mm/yyyy
//   - monto: negativo = sale plata (compra, cargo); positivo = entra (sueldo, abono)
//   - cuotas: "3/12" si es una cuota de una compra en cuotas
// Estado de cuenta de tarjeta (una fila por tarjeta):
//   producto, fecha_estado, deuda_total, monto_a_pagar, fecha_pago [, cupo_total]
// ---------------------------------------------------------------------------

function carpetaImportar_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('CARPETA_IMPORTAR_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) {}
  }
  var it = DriveApp.getFoldersByName(CARPETA_IMPORTAR);
  var carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA_IMPORTAR);
  props.setProperty('CARPETA_IMPORTAR_ID', carpeta.getId());
  return carpeta;
}

function carpetaProcesadas_(carpeta) {
  var it = carpeta.getFoldersByName(CARPETA_PROCESADAS);
  return it.hasNext() ? it.next() : carpeta.createFolder(CARPETA_PROCESADAS);
}

// La corre el activador cada hora y el botón "Importar ahora".
function importarPendientes() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return [];
  try {
    var carpeta = carpetaImportar_();
    var procesadas = carpetaProcesadas_(carpeta);
    var archivos = [];
    var it = carpeta.getFiles();
    while (it.hasNext()) {
      var f = it.next();
      if (/\.csv$/i.test(f.getName()) || f.getMimeType() === 'text/csv') archivos.push(f);
    }
    archivos.sort(function (a, b) { return a.getName() < b.getName() ? -1 : 1; });
    var resultados = archivos.map(function (f) {
      var r;
      try {
        r = importarArchivo_(f);
      } catch (err) {
        r = { archivo: f.getName(), nuevos: 0, repetidos: 0, errores: 1, detalle: err.message, fatal: true };
      }
      agregarFilas_(hoja_(HOJA_IMPORTACIONES, COLUMNAS_IMPORTACIONES), [{
        'Archivo': r.archivo, 'Nuevos': r.nuevos, 'Repetidos': r.repetidos, 'Con error': r.errores,
        'Detalle': r.detalle || '', 'Creado': new Date()
      }]);
      // Solo se mueve a "Procesadas" si se pudo leer; si no, queda para revisarlo.
      if (!r.fatal) f.moveTo(procesadas);
      return r;
    });
    if (resultados.length) aplicarReglasPendientes_();
    return resultados;
  } finally {
    lock.releaseLock();
  }
}

function importarArchivo_(archivo) {
  var texto = archivo.getBlob().getDataAsString('UTF-8').replace(/^﻿/, '');
  var sep = (texto.split('\n')[0].match(/;/g) || []).length > (texto.split('\n')[0].match(/,/g) || []).length ? ';' : ',';
  var filas = Utilities.parseCsv(texto, sep);
  if (filas.length < 2) return { archivo: archivo.getName(), nuevos: 0, repetidos: 0, errores: 0, detalle: 'Archivo vacío' };
  var enc = filas[0].map(function (c) { return sinTildes_(c).toLowerCase().replace(/ /g, '_'); });
  var datos = filas.slice(1).filter(function (f) { return f.join('').trim(); }).map(function (f) {
    var o = {};
    enc.forEach(function (c, i) { o[c] = String(f[i] === undefined ? '' : f[i]).trim(); });
    return o;
  });
  if (enc.indexOf('deuda_total') !== -1) return importarTarjetas_(archivo.getName(), datos);
  ['fecha', 'producto', 'descripcion', 'monto'].forEach(function (c) {
    if (enc.indexOf(c) === -1) throw new Error('Al archivo le falta la columna "' + c + '"');
  });
  return importarMovimientos_(archivo.getName(), datos);
}

function claveMovimiento_(fechaIso, producto, descripcion, monto, n) {
  var base = [fechaIso, sinTildes_(producto), sinTildes_(descripcion), monto, n].join('|');
  var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, base, Utilities.Charset.UTF_8);
  return Utilities.base64EncodeWebSafe(hash).slice(0, 22);
}

function importarMovimientos_(nombre, datos) {
  var hoja = hoja_(HOJA_MOV, COLUMNAS_MOV);
  var tz = zona_();
  var existentes = {};
  leerTabla_(hoja).filas.forEach(function (m) { if (m['Clave']) existentes[m['Clave']] = true; });
  var reglas = leerReglas_();
  var cat = categorias_();
  var vistos = {};   // dos compras iguales el mismo día son dos movimientos distintos
  var nuevos = [], repetidos = 0, errores = [];
  datos.forEach(function (d, i) {
    var fecha = parsearFecha_(d.fecha);
    var monto = parsearMonto_(d.monto);
    if (!fecha || isNaN(monto) || !d.descripcion) {
      errores.push('fila ' + (i + 2));
      return;
    }
    var iso = isoFecha_(fecha, tz);
    var base = [iso, sinTildes_(d.producto), sinTildes_(d.descripcion), monto].join('|');
    vistos[base] = (vistos[base] || 0) + 1;
    var clave = claveMovimiento_(iso, d.producto, d.descripcion, monto, vistos[base]);
    if (existentes[clave]) { repetidos++; return; }
    existentes[clave] = true;

    var tipo = TIPOS.indexOf(d.tipo) !== -1 ? d.tipo : '';
    var centro = '', cuenta = '';
    var regla = reglaPara_(d.descripcion, reglas);
    if (!tipo && regla && regla.tipo) tipo = regla.tipo;
    if (!tipo) tipo = tipoAutomatico_(d.producto, d.descripcion, monto);
    if (d.centro && (cat.centros.indexOf(d.centro) !== -1 || d.centro === CENTRO_INGRESOS)) {
      centro = d.centro;
      cuenta = (cat.cuentas[centro] || []).indexOf(d.cuenta) !== -1 ? d.cuenta : '';
    } else if (regla && (tipo === 'Gasto' || tipo === 'Ingreso')) {
      centro = regla.centro;
      cuenta = regla.cuenta;
    }
    nuevos.push({
      'ID': Utilities.getUuid(), 'Fecha': fecha, 'Banco': d.banco || '', 'Producto': d.producto,
      'Descripción': d.descripcion, 'Monto': monto, 'Tipo': tipo, 'Centro de costo': centro, 'Cuenta': cuenta,
      'Cuotas': d.cuotas || '', 'Nota': d.nota || '', 'Clave': clave, 'Origen': nombre, 'Creado': new Date()
    });
  });
  nuevos.sort(function (a, b) { return a['Fecha'] - b['Fecha']; });
  agregarFilas_(hoja, nuevos);
  return {
    archivo: nombre, nuevos: nuevos.length, repetidos: repetidos, errores: errores.length,
    detalle: errores.length ? 'No se pudieron leer: ' + errores.slice(0, 10).join(', ') : ''
  };
}

function importarTarjetas_(nombre, datos) {
  var hoja = hoja_(HOJA_TARJETAS, COLUMNAS_TARJETAS);
  var t = leerTabla_(hoja);
  var tz = zona_();
  var existentes = {};
  t.filas.forEach(function (r) { existentes[sinTildes_(r['Producto']) + '|' + isoFecha_(r['Fecha estado'], tz)] = true; });
  var nuevos = [], repetidos = 0, errores = [];
  datos.forEach(function (d, i) {
    var fecha = parsearFecha_(d.fecha_estado);
    var deuda = parsearMonto_(d.deuda_total);
    if (!d.producto || !fecha || isNaN(deuda)) { errores.push('fila ' + (i + 2)); return; }
    var k = sinTildes_(d.producto) + '|' + isoFecha_(fecha, tz);
    if (existentes[k]) { repetidos++; return; }
    existentes[k] = true;
    var pagar = parsearMonto_(d.monto_a_pagar);
    var cupo = parsearMonto_(d.cupo_total);
    nuevos.push({
      'Producto': d.producto, 'Fecha estado': fecha, 'Deuda total': Math.abs(deuda),
      'Monto a pagar': isNaN(pagar) ? '' : Math.abs(pagar), 'Fecha pago': parsearFecha_(d.fecha_pago) || '',
      'Cupo total': isNaN(cupo) ? '' : Math.abs(cupo), 'Origen': nombre, 'Creado': new Date()
    });
  });
  agregarFilas_(hoja, nuevos);
  return {
    archivo: nombre, nuevos: nuevos.length, repetidos: repetidos, errores: errores.length,
    detalle: 'Estados de cuenta de tarjeta' + (errores.length ? ' · no se pudieron leer: ' + errores.join(', ') : '')
  };
}

// ---------------------------------------------------------------------------
// Datos para la app
// ---------------------------------------------------------------------------

function datos_() {
  var aviso = sincronizarCentros_(false);
  formatoSiFalta_();
  var tz = zona_();
  var desde = new Date();
  desde.setMonth(desde.getMonth() - MESES_EN_APP);
  desde.setDate(1);
  var desdeIso = isoFecha_(desde, tz);

  // Movimientos compactos: [id, fecha, producto, descripción, monto, tipo, centro, cuenta, cuotas, banco, nota]
  var movs = leerTabla_(hoja_(HOJA_MOV, COLUMNAS_MOV)).filas.map(function (m) {
    return [String(m['ID']), isoFecha_(m['Fecha'], tz), String(m['Producto'] || ''), String(m['Descripción'] || ''),
      Number(m['Monto']) || 0, String(m['Tipo'] || 'Gasto'), String(m['Centro de costo'] || ''), String(m['Cuenta'] || ''),
      String(m['Cuotas'] || ''), String(m['Banco'] || ''), String(m['Nota'] || '')];
  }).filter(function (m) { return m[0] && m[1] >= desdeIso; })
    .sort(function (a, b) { return a[1] < b[1] ? 1 : a[1] > b[1] ? -1 : 0; });

  // Último estado de cuenta de cada tarjeta.
  var tarjetas = {};
  leerTabla_(hoja_(HOJA_TARJETAS, COLUMNAS_TARJETAS)).filas.forEach(function (r) {
    var p = String(r['Producto'] || '').trim();
    var f = isoFecha_(r['Fecha estado'], tz);
    if (!p || (tarjetas[p] && tarjetas[p].fecha >= f)) return;
    tarjetas[p] = {
      producto: p, fecha: f, deuda: Number(r['Deuda total']) || 0, pagar: Number(r['Monto a pagar']) || 0,
      pago: r['Fecha pago'] ? isoFecha_(r['Fecha pago'], tz) : '', cupo: Number(r['Cupo total']) || 0
    };
  });

  var presupuesto = leerTabla_(hoja_(HOJA_PRESUPUESTO, COLUMNAS_PRESUPUESTO)).filas.map(function (r) {
    return { centro: String(r['Centro de costo'] || '').trim(), cuenta: String(r['Cuenta'] || '').trim(), tope: Number(r['Tope mensual']) || 0 };
  }).filter(function (p) { return p.centro && p.tope > 0; });

  var reglas = leerReglas_().length;
  var imp = leerTabla_(hoja_(HOJA_IMPORTACIONES, COLUMNAS_IMPORTACIONES)).filas;
  var ultima = imp.length ? imp[imp.length - 1] : null;

  var cat = categorias_();
  return {
    ok: true, version: VERSION, aviso: aviso,
    centros: cat.centros, cuentas: cat.cuentas, tipos: TIPOS,
    movimientos: movs,
    tarjetas: Object.keys(tarjetas).map(function (k) { return tarjetas[k]; }),
    presupuesto: presupuesto,
    reglas: reglas,
    ultimaImportacion: ultima ? { archivo: String(ultima['Archivo']), nuevos: Number(ultima['Nuevos']) || 0,
      fecha: Utilities.formatDate(new Date(ultima['Creado']), tz, "yyyy-MM-dd'T'HH:mm") } : null
  };
}

// Cambia tipo / centro / cuenta / nota de un movimiento y, si se pide, guarda una regla.
function categorizar_(body) {
  var tipo = String(body.tipo || '').trim();
  var cat = validarCategoria_(tipo, String(body.centro || '').trim(), String(body.cuenta || '').trim());
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var hoja = hoja_(HOJA_MOV, COLUMNAS_MOV);
    var t = leerTabla_(hoja);
    var m = t.filas.filter(function (x) { return String(x['ID']) === String(body.id); })[0];
    if (!m) throw new Error('No encontré ese movimiento. Actualiza la app.');
    var campos = { 'Tipo': tipo, 'Centro de costo': cat.centro, 'Cuenta': cat.cuenta };
    if (body.nota !== undefined) campos['Nota'] = String(body.nota || '').slice(0, 200);
    escribirCampos_(hoja, t.enc, m._fila, campos);
    if (body.recordar) {
      guardarRegla_(body.recordar, tipo, cat.centro, cat.cuenta);
      aplicarReglasPendientes_();
    }
  } finally {
    lock.releaseLock();
  }
}

function guardarPresupuesto_(body) {
  var centro = String(body.centro || '').trim();
  var cuenta = String(body.cuenta || '').trim();
  var tope = Math.round(Number(body.tope) || 0);
  if (tope < 0) throw new Error('El tope no puede ser negativo');
  validarCategoria_('Gasto', centro, cuenta);
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var hoja = hoja_(HOJA_PRESUPUESTO, COLUMNAS_PRESUPUESTO);
    var t = leerTabla_(hoja);
    var fila = t.filas.filter(function (r) {
      return String(r['Centro de costo']).trim() === centro && String(r['Cuenta'] || '').trim() === cuenta;
    })[0];
    if (fila && !tope) hoja.deleteRow(fila._fila);
    else if (fila) escribirCampos_(hoja, t.enc, fila._fila, { 'Tope mensual': tope });
    else if (tope) agregarFilas_(hoja, [{ 'Centro de costo': centro, 'Cuenta': cuenta, 'Tope mensual': tope }]);
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------------------------
// API web. Todas las acciones van por POST (text/plain para evitar CORS).
// ---------------------------------------------------------------------------

function doGet() {
  return json_({ ok: true, mensaje: 'Finanzas: backend funcionando.', version: VERSION });
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'modo') return json_(modoAcceso_());
    if (body.action === 'loginGoogle') return json_(loginGoogle_(body));
    autorizar_(body);

    switch (body.action) {
      case 'datos':
        return json_(datos_());
      case 'categorizar':
        categorizar_(body);
        return json_(datos_());
      case 'presupuesto':
        guardarPresupuesto_(body);
        return json_(datos_());
      case 'importar':
        var r = importarPendientes();
        return json_(Object.assign(datos_(), { importados: r }));
      case 'planilla':
        return json_({ ok: true, url: libro_().getUrl(), carpeta: carpetaImportar_().getUrl() });
      default:
        throw new Error('Acción desconocida');
    }
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ---------------------------------------------------------------------------
// Acceso (igual que Casa App)
// - Con las propiedades GOOGLE_CLIENT_ID y CORREOS_PERMITIDOS: solo entran esos
//   correos, con su cuenta de Google.
// - Sin ellas: PIN (sirve mientras se configura Google).
// 5 intentos fallidos seguidos bloquean el ingreso por 15 minutos.
// ---------------------------------------------------------------------------
var MAX_FALLOS = 5;
var BLOQUEO_SEGUNDOS = 15 * 60;
var DURACION_SESION_DIAS = 365;

function propiedad_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }

function correosPermitidos_() {
  return propiedad_('CORREOS_PERMITIDOS').split(/[\s,;]+/)
    .map(function (c) { return c.trim().toLowerCase(); })
    .filter(function (c) { return c; });
}

function modoGoogle_() { return !!(propiedad_('GOOGLE_CLIENT_ID') && correosPermitidos_().length); }

function modoAcceso_() {
  return { ok: true, google: modoGoogle_(), clientId: modoGoogle_() ? propiedad_('GOOGLE_CLIENT_ID') : '' };
}

function revisarBloqueo_() {
  var n = Number(CacheService.getScriptCache().get('fallos') || 0);
  if (n >= MAX_FALLOS) throw new Error('Demasiados intentos fallidos. Espera 15 minutos e intenta de nuevo.');
}

function sumarFallo_() {
  var cache = CacheService.getScriptCache();
  cache.put('fallos', String(Number(cache.get('fallos') || 0) + 1), BLOQUEO_SEGUNDOS);
}

function claveSesion_(token) {
  var hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token));
  return 'SES_' + Utilities.base64EncodeWebSafe(hash).replace(/=+$/, '');
}

function autorizar_(body) {
  if (modoGoogle_()) {
    var token = String(body.sesion || '');
    var guardada = token ? propiedad_(claveSesion_(token)) : '';
    var s = null;
    try { s = guardada ? JSON.parse(guardada) : null; } catch (e) {}
    var vencida = s && (Date.now() - s.creado) > DURACION_SESION_DIAS * 864e5;
    if (!s || vencida || correosPermitidos_().indexOf(s.email) === -1) {
      if (s) PropertiesService.getScriptProperties().deleteProperty(claveSesion_(token));
      throw new Error('Sesión no válida: entra con tu cuenta de Google');
    }
    return s.email;
  }
  revisarBloqueo_();
  var esperado = propiedad_('PIN');
  if (!esperado) throw new Error('El PIN no está configurado en el script');
  if (String(body.pin || '') !== String(esperado)) {
    sumarFallo_();
    throw new Error('PIN incorrecto');
  }
  return '';
}

function loginGoogle_(body) {
  if (!modoGoogle_()) throw new Error('El ingreso con Google no está activado');
  revisarBloqueo_();
  var idToken = String(body.idToken || '');
  var nonce = String(body.nonce || '');
  if (!idToken || !nonce) throw new Error('Faltan datos de Google');
  var r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('Google no validó el ingreso. Intenta de nuevo.');
  var t = JSON.parse(r.getContentText());
  var valido = t.aud === propiedad_('GOOGLE_CLIENT_ID') &&
    (t.iss === 'accounts.google.com' || t.iss === 'https://accounts.google.com') &&
    String(t.email_verified) === 'true' &&
    Number(t.exp) * 1000 > Date.now() &&
    t.nonce === nonce;
  if (!valido) throw new Error('Google no validó el ingreso. Intenta de nuevo.');
  var email = String(t.email || '').toLowerCase();
  if (correosPermitidos_().indexOf(email) === -1) {
    sumarFallo_();
    throw new Error('La cuenta ' + email + ' no tiene acceso a esta app.');
  }
  var token = Utilities.getUuid().replace(/-/g, '');
  PropertiesService.getScriptProperties().setProperty(claveSesion_(token), JSON.stringify({ email: email, creado: Date.now() }));
  return { ok: true, sesion: token, email: email };
}

// Ejecutar UNA vez desde el editor después de poner GOOGLE_CLIENT_ID y CORREOS_PERMITIDOS.
function activarAccesoGoogle() {
  UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=x', { muteHttpExceptions: true });
  var lineas = [];
  lineas.push(propiedad_('GOOGLE_CLIENT_ID') ? '✓ GOOGLE_CLIENT_ID configurado.' : '✗ Falta la propiedad GOOGLE_CLIENT_ID.');
  var correos = correosPermitidos_();
  lineas.push(correos.length ? '✓ Correos con acceso: ' + correos.join(', ') : '✗ Falta la propiedad CORREOS_PERMITIDOS.');
  lineas.push(modoGoogle_() ? 'Listo: desde ahora solo se entra con esas cuentas de Google (publica una nueva versión).' : 'Todavía se entra con PIN.');
  Logger.log(lineas.join('\n'));
}

function cerrarTodasLasSesiones() {
  var props = PropertiesService.getScriptProperties();
  var n = 0;
  Object.keys(props.getProperties()).forEach(function (k) {
    if (k.indexOf('SES_') === 0) { props.deleteProperty(k); n++; }
  });
  Logger.log('Sesiones cerradas: ' + n);
}

function verSesiones() {
  var props = PropertiesService.getScriptProperties().getProperties();
  var lineas = Object.keys(props).filter(function (k) { return k.indexOf('SES_') === 0; }).map(function (k) {
    var s = JSON.parse(props[k]);
    return s.email + ' · desde ' + new Date(s.creado).toLocaleString();
  });
  Logger.log(lineas.length ? lineas.join('\n') : 'No hay sesiones abiertas.');
}

// ===========================================================================
// FORMATO VISUAL DE LA PLANILLA (solo apariencia, nunca datos)
// ===========================================================================
var FORMATO_VERSION = '2026-11-06 v1';

var ESTILO = {
  fuente: 'Arial', tamano: 10, texto: '#1F2937', textoSuave: '#6B7280',
  encabezado: '#1E3A5F', encabezadoTexto: '#FFFFFF', banda: '#F3F6FA', blanco: '#FFFFFF'
};

var COLOR_PESTANA = {
  'Movimientos': '#1E3A5F', 'Tarjetas': '#1E3A5F',
  'Presupuesto': '#2E7D5B', 'Reglas': '#2E7D5B',
  'Centros': '#9AA5B1', 'Cuentas': '#9AA5B1', 'Importaciones': '#9AA5B1'
};
var ORDEN_PESTANAS = ['Movimientos', 'Tarjetas', 'Presupuesto', 'Reglas', 'Centros', 'Cuentas', 'Importaciones'];

var FMT_PESOS = '$#,##0" ";-$#,##0" "';

var COLUMNAS_ESTILO = {
  'ID': { ancho: 80, suave: true },
  'Fecha': { ancho: 95, formato: 'dd/mm/yyyy', alinear: 'center' },
  'Fecha estado': { ancho: 100, formato: 'dd/mm/yyyy', alinear: 'center' },
  'Fecha pago': { ancho: 100, formato: 'dd/mm/yyyy', alinear: 'center' },
  'Creado': { ancho: 125, formato: 'dd/mm/yyyy hh:mm', alinear: 'center', suave: true },
  'Banco': { ancho: 110 },
  'Producto': { ancho: 170 },
  'Descripción': { ancho: 300, ajustar: true },
  'Monto': { ancho: 110, formato: FMT_PESOS, alinear: 'right' },
  'Deuda total': { ancho: 115, formato: FMT_PESOS, alinear: 'right' },
  'Monto a pagar': { ancho: 115, formato: FMT_PESOS, alinear: 'right' },
  'Cupo total': { ancho: 115, formato: FMT_PESOS, alinear: 'right' },
  'Tope mensual': { ancho: 115, formato: FMT_PESOS, alinear: 'right' },
  'Tipo': { ancho: 105, alinear: 'center' },
  'Centro de costo': { ancho: 135 },
  'Cuenta': { ancho: 145 },
  'Cuotas': { ancho: 70, formato: '@', alinear: 'center' },
  'Nota': { ancho: 200, ajustar: true },
  'Clave': { ancho: 90, suave: true },
  'Origen': { ancho: 180, suave: true },
  'Contiene': { ancho: 220 },
  'Archivo': { ancho: 240 },
  'Nuevos': { ancho: 80, formato: '0', alinear: 'right' },
  'Repetidos': { ancho: 80, formato: '0', alinear: 'right' },
  'Con error': { ancho: 80, formato: '0', alinear: 'right' },
  'Detalle': { ancho: 280, ajustar: true }
};
var ANCHO_POR_DEFECTO = 120;

function darFormato() {
  var ss = libro_();
  ss.getSheets().forEach(function (hoja) {
    try { estiloTabla_(hoja); } catch (err) { Logger.log('Formato "' + hoja.getName() + '": ' + err); }
  });
  var pos = 1;
  ORDEN_PESTANAS.forEach(function (nombre) {
    var hoja = ss.getSheetByName(nombre);
    if (!hoja) return;
    ss.setActiveSheet(hoja);
    ss.moveActiveSheet(pos++);
  });
  PropertiesService.getScriptProperties().setProperty('FORMATO_VERSION', FORMATO_VERSION);
}

function formatoSiFalta_() {
  try {
    if (propiedad_('FORMATO_VERSION') !== FORMATO_VERSION) return darFormato();
    var mov = libro_().getSheetByName(HOJA_MOV);
    var b = mov && mov.getBandings()[0];
    if (mov && (!b || b.getRange().getLastRow() < mov.getMaxRows())) estiloTabla_(mov);
  } catch (err) {
    Logger.log('formatoSiFalta_: ' + err);
  }
}

function estiloTabla_(hoja) {
  if (hoja.isSheetHidden()) return;
  var nombre = hoja.getName();
  if (COLOR_PESTANA[nombre]) hoja.setTabColor(COLOR_PESTANA[nombre]);
  if (hoja.getLastRow() === 0 || hoja.getLastColumn() === 0) return;
  var nCols = hoja.getLastColumn();
  var nDatos = Math.max(hoja.getMaxRows() - 1, 1);
  var enc = encabezados_(hoja);
  hoja.getRange(1, 1, nDatos + 1, nCols).setFontFamily(ESTILO.fuente).setFontSize(ESTILO.tamano)
    .setFontColor(ESTILO.texto).setVerticalAlignment('middle');
  hoja.setHiddenGridlines(true);
  hoja.setFrozenRows(1);
  hoja.getRange(1, 1, 1, nCols).setBackground(ESTILO.encabezado).setFontColor(ESTILO.encabezadoTexto)
    .setFontWeight('bold').setWrap(true);
  hoja.setRowHeight(1, 30);
  enc.forEach(function (c, i) {
    var cfg = COLUMNAS_ESTILO[c] || {};
    var col = i + 1;
    hoja.setColumnWidth(col, cfg.ancho || ANCHO_POR_DEFECTO);
    var datos = hoja.getRange(2, col, nDatos, 1);
    if (cfg.formato) datos.setNumberFormat(cfg.formato);
    datos.setHorizontalAlignment(cfg.alinear || 'left');
    hoja.getRange(1, col).setHorizontalAlignment(cfg.alinear || 'left');
    datos.setWrapStrategy(cfg.ajustar ? SpreadsheetApp.WrapStrategy.WRAP : SpreadsheetApp.WrapStrategy.CLIP);
    if (cfg.suave) datos.setFontColor(ESTILO.textoSuave);
  });
  hoja.getBandings().forEach(function (b) { b.remove(); });
  hoja.getRange(1, 1, nDatos + 1, nCols).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false)
    .setHeaderRowColor(ESTILO.encabezado).setFirstRowColor(ESTILO.blanco).setSecondRowColor(ESTILO.banda);
  if (nombre === HOJA_MOV && !hoja.getFilter()) {
    hoja.getRange(1, 1, Math.max(hoja.getLastRow(), 2), nCols).createFilter();
  }
  if (nombre === HOJA_MOV) {
    // Montos que entran en verde.
    var col = enc.indexOf('Monto') + 1;
    if (col > 0) {
      var rango = hoja.getRange(2, col, nDatos, 1);
      hoja.setConditionalFormatRules([SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0)
        .setFontColor('#1E6B3F').setRanges([rango]).build()]);
    }
  }
}
