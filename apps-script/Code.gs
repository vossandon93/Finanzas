/**
 * Backend de Gastos Familia (Google Apps Script).
 *
 * Se pega en Extensiones > Apps Script de la planilla de Google.
 * Ver SETUP.md para las instrucciones completas.
 *
 * El PIN se guarda en Configuración del proyecto > Propiedades del script,
 * con la clave PIN. Nunca se escribe en este archivo.
 */

var HOJA_GASTOS = 'Gastos';
var HOJA_CENTROS = 'Centros';
var HOJA_PERSONAS = 'Personas';
var HOJA_RESUMEN_MENSUAL = 'Resumen mensual';
var HOJA_RESUMEN_ANUAL = 'Resumen anual';

var COLUMNAS_GASTOS = ['ID', 'Fecha', 'Centro de costo', 'Monto', 'Descripción', 'Registró', 'Creado'];
var CENTROS_INICIALES = ['Farmacia', 'Luna', 'Supermercado', 'Fran'];
var PERSONAS_INICIALES = ['Vicente', 'Fran'];
var ULTIMOS_GASTOS = 20;

// ---------------------------------------------------------------------------
// Instalación: ejecutar una vez desde el editor (botón "Ejecutar" con setup).
// ---------------------------------------------------------------------------

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    // Proyecto creado en script.google.com (no dentro de una planilla):
    // se usa la planilla creada antes o se crea una nueva en tu Drive.
    var props = PropertiesService.getScriptProperties();
    var id = props.getProperty('SHEET_ID');
    ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create('Gastos Familia');
    props.setProperty('SHEET_ID', ss.getId());
    Logger.log('Planilla: ' + ss.getUrl());
  }
  ss.setSpreadsheetTimeZone('America/Santiago');

  var gastos = obtenerOCrearHoja_(ss, HOJA_GASTOS);
  if (gastos.getLastRow() === 0) {
    gastos.appendRow(COLUMNAS_GASTOS);
  }
  gastos.setFrozenRows(1);
  gastos.getRange(1, 1, 1, COLUMNAS_GASTOS.length).setFontWeight('bold');
  gastos.getRange('B:B').setNumberFormat('dd/mm/yyyy');
  gastos.getRange('D:D').setNumberFormat('$#,##0');
  gastos.getRange('G:G').setNumberFormat('dd/mm/yyyy hh:mm');
  gastos.hideColumns(1);

  var centros = obtenerOCrearHoja_(ss, HOJA_CENTROS);
  if (centros.getLastRow() === 0) {
    centros.appendRow(['Centro de costo']);
    CENTROS_INICIALES.forEach(function (c) { centros.appendRow([c]); });
    centros.getRange(1, 1).setFontWeight('bold');
    centros.setFrozenRows(1);
  }

  var personas = obtenerOCrearHoja_(ss, HOJA_PERSONAS);
  if (personas.getLastRow() === 0) {
    personas.appendRow(['Persona']);
    PERSONAS_INICIALES.forEach(function (p) { personas.appendRow([p]); });
    personas.getRange(1, 1).setFontWeight('bold');
    personas.setFrozenRows(1);
  }

  crearResumenes_(ss);

  var hoja1 = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (hoja1 && hoja1.getLastRow() === 0 && ss.getSheets().length > 1) {
    ss.deleteSheet(hoja1);
  }

  if (!PropertiesService.getScriptProperties().getProperty('PIN')) {
    Logger.log('Falta definir el PIN en Configuración del proyecto > Propiedades del script.');
  }
}

function crearResumenes_(ss) {
  var rango = HOJA_GASTOS + '!B:F';
  // Las letras dentro de QUERY son relativas al rango B:F: B=Fecha, C=Centro, D=Monto.
  var mensual = obtenerOCrearHoja_(ss, HOJA_RESUMEN_MENSUAL);
  mensual.clear();
  mensual.getRange('A1').setValue('Total por mes').setFontWeight('bold');
  mensual.getRange('A2').setFormula(
    '=IFERROR(QUERY(' + rango + ', "select year(B), month(B)+1, sum(D) where B is not null ' +
    'group by year(B), month(B)+1 order by year(B) desc, month(B)+1 desc ' +
    'label year(B) \'Año\', month(B)+1 \'Mes\', sum(D) \'Total\'", 1), "Sin gastos todavía")'
  );
  mensual.getRange('E1').setValue('Por centro de costo').setFontWeight('bold');
  mensual.getRange('E2').setFormula(
    '=IFERROR(QUERY(' + rango + ', "select year(B), month(B)+1, sum(D) where B is not null ' +
    'group by year(B), month(B)+1 pivot C order by year(B) desc, month(B)+1 desc ' +
    'label year(B) \'Año\', month(B)+1 \'Mes\'", 1), "")'
  );
  mensual.getRange('C:C').setNumberFormat('$#,##0');
  mensual.getRange('G:Z').setNumberFormat('$#,##0');
  mensual.setFrozenRows(2);

  var anual = obtenerOCrearHoja_(ss, HOJA_RESUMEN_ANUAL);
  anual.clear();
  anual.getRange('A1').setValue('Total por año').setFontWeight('bold');
  anual.getRange('A2').setFormula(
    '=IFERROR(QUERY(' + rango + ', "select year(B), sum(D) where B is not null ' +
    'group by year(B) order by year(B) desc label year(B) \'Año\', sum(D) \'Total\'", 1), "Sin gastos todavía")'
  );
  anual.getRange('D1').setValue('Por centro de costo').setFontWeight('bold');
  anual.getRange('D2').setFormula(
    '=IFERROR(QUERY(' + rango + ', "select year(B), sum(D) where B is not null ' +
    'group by year(B) pivot C order by year(B) desc label year(B) \'Año\'", 1), "")'
  );
  anual.getRange('B:B').setNumberFormat('$#,##0');
  anual.getRange('E:Z').setNumberFormat('$#,##0');
  anual.setFrozenRows(2);
}

function libro_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  var id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if (!id) throw new Error('Falta la planilla: ejecuta setup() primero');
  return SpreadsheetApp.openById(id);
}

function obtenerOCrearHoja_(ss, nombre) {
  return ss.getSheetByName(nombre) || ss.insertSheet(nombre);
}

// ---------------------------------------------------------------------------
// API web. Todas las acciones van por POST (text/plain para evitar CORS).
// ---------------------------------------------------------------------------

function doGet() {
  return json_({ ok: true, mensaje: 'Gastos Familia: backend funcionando.' });
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    validarPin_(body.pin);

    switch (body.action) {
      case 'config':
        return json_({ ok: true, centros: leerLista_(HOJA_CENTROS), personas: leerLista_(HOJA_PERSONAS) });
      case 'list':
        return json_({ ok: true, gastos: ultimosGastos_(), totalMes: totalMesActual_() });
      case 'add':
        agregarGasto_(body);
        return json_({ ok: true, gastos: ultimosGastos_(), totalMes: totalMesActual_() });
      case 'delete':
        borrarGasto_(body.id);
        return json_({ ok: true, gastos: ultimosGastos_(), totalMes: totalMesActual_() });
      default:
        throw new Error('Acción desconocida');
    }
  } catch (err) {
    return json_({ ok: false, error: err.message });
  }
}

function validarPin_(pin) {
  var esperado = PropertiesService.getScriptProperties().getProperty('PIN');
  if (!esperado) throw new Error('El PIN no está configurado en el script');
  if (String(pin || '') !== String(esperado)) throw new Error('PIN incorrecto');
}

function leerLista_(nombreHoja) {
  var hoja = libro_().getSheetByName(nombreHoja);
  if (!hoja || hoja.getLastRow() < 2) return [];
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues()
    .map(function (fila) { return String(fila[0]).trim(); })
    .filter(function (v) { return v; });
}

function agregarGasto_(body) {
  var centro = String(body.centro || '').trim();
  var persona = String(body.persona || '').trim();
  var descripcion = String(body.descripcion || '').trim().slice(0, 200);
  var monto = Math.round(Number(body.monto));
  var fecha = parsearFecha_(body.fecha);

  if (leerLista_(HOJA_CENTROS).indexOf(centro) === -1) throw new Error('Centro de costo inválido');
  if (leerLista_(HOJA_PERSONAS).indexOf(persona) === -1) throw new Error('Persona inválida');
  if (!(monto > 0)) throw new Error('Monto inválido');

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    hojaGastos_().appendRow([Utilities.getUuid(), fecha, centro, monto, descripcion, persona, new Date()]);
  } finally {
    lock.releaseLock();
  }
}

function borrarGasto_(id) {
  if (!id) throw new Error('Falta el ID');
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var hoja = hojaGastos_();
    var ultima = hoja.getLastRow();
    if (ultima < 2) throw new Error('Gasto no encontrado');
    var ids = hoja.getRange(2, 1, ultima - 1, 1).getValues();
    for (var i = ids.length - 1; i >= 0; i--) {
      if (ids[i][0] === id) {
        hoja.deleteRow(i + 2);
        return;
      }
    }
    throw new Error('Gasto no encontrado');
  } finally {
    lock.releaseLock();
  }
}

function ultimosGastos_() {
  var hoja = hojaGastos_();
  var ultima = hoja.getLastRow();
  if (ultima < 2) return [];
  var n = Math.min(ULTIMOS_GASTOS, ultima - 1);
  var tz = libro_().getSpreadsheetTimeZone();
  return hoja.getRange(ultima - n + 1, 1, n, COLUMNAS_GASTOS.length).getValues()
    .reverse()
    .map(function (f) {
      return {
        id: f[0],
        fecha: f[1] instanceof Date ? Utilities.formatDate(f[1], tz, 'yyyy-MM-dd') : String(f[1]),
        centro: f[2],
        monto: Number(f[3]) || 0,
        descripcion: f[4],
        persona: f[5]
      };
    });
}

function totalMesActual_() {
  var hoja = hojaGastos_();
  var ultima = hoja.getLastRow();
  if (ultima < 2) return 0;
  var tz = libro_().getSpreadsheetTimeZone();
  var mesActual = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  return hoja.getRange(2, 2, ultima - 1, 3).getValues().reduce(function (suma, f) {
    if (f[0] instanceof Date && Utilities.formatDate(f[0], tz, 'yyyy-MM') === mesActual) {
      return suma + (Number(f[2]) || 0);
    }
    return suma;
  }, 0);
}

function hojaGastos_() {
  var hoja = libro_().getSheetByName(HOJA_GASTOS);
  if (!hoja) throw new Error('Falta la hoja Gastos: ejecuta setup() primero');
  return hoja;
}

function parsearFecha_(texto) {
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(texto || ''));
  if (!m) return new Date();
  // Mediodía para que ningún cambio de zona horaria mueva el día.
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12, 0, 0);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
