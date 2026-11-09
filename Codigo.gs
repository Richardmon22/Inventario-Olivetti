/**
 * ========================================================
 *  CONTROL DE IMPRESORAS OLIVETTI - BACKEND
 * ========================================================
 */

const SPREADSHEET_ID = '1pAifMKJJ5NL63EOrrnWA2WaER8G85sgWO7KeqE4AWjA';
const SHEET_NAME = 'Hoja 2';
const SHEET_INDEX_FALLBACK = 1;
const CACHE_KEY = 'olivetti_data_v4';
const CACHE_TTL = 30;

const SHEET_PROVEDURIA = 'Proveduria';
const SHEET_PROVEDURIA_FALLBACK = 2;
const CACHE_KEY_PROVEDURIA = 'olivetti_proveduria_v2'; 

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Control de Impresoras Olivetti')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSpreadsheet_() {
  let id = '';
  try {
    id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || '';
  } catch (e) {}
  if (!id) id = SPREADSHEET_ID;
  if (!id || id === 'PEGA_AQUI_TU_ID_DE_LA_HOJA') {
    throw new Error('Debes configurar SPREADSHEET_ID con el ID de tu hoja de cálculo.');
  }
  try {
    return SpreadsheetApp.openById(id);
  } catch (e) {
    throw new Error('No se pudo abrir la hoja de cálculo. Verifica el ID y los permisos: ' + e.message);
  }
}

function getSheet_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.getSheets()[SHEET_INDEX_FALLBACK];
  if (!sheet) throw new Error('No se encontró la hoja "' + SHEET_NAME + '"');
  return sheet;
}

function getSheetProveduria_() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_PROVEDURIA);
  if (!sheet) sheet = ss.getSheets()[SHEET_PROVEDURIA_FALLBACK];
  return sheet || null;
}

function invalidarCache_() {
  try { CacheService.getScriptCache().remove(CACHE_KEY); } catch(e) {}
}
function invalidarCacheProveduria_() {
  try { CacheService.getScriptCache().remove(CACHE_KEY_PROVEDURIA); } catch(e) {}
}

function obtenerDatosIniciales() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(CACHE_KEY);
    if (cached) return JSON.parse(cached);
  } catch(e) {}

  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return { impresoras: [], catalogos: { agencias: [], modelos: [] } };
  }

  const impresoras = [];
  const agenciasSet = new Set();
  const modelosSet = new Set();

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const agencia   = limpiar_(row[0]);
    const serial    = limpiarSerial_(row[1]);
    const puertoRaw = limpiar_(row[2]);
    const modeloD   = limpiar_(row[3]);
    const modeloE   = limpiar_(row[4]);
    const obsF      = limpiar_(row[5]);
    const obsG      = limpiar_(row[6]);
    const obsH      = limpiar_(row[7]);
    const estadoRaw = limpiar_(row[8]);
    const activoFijoRaw = limpiar_(row[9]);

    if (!agencia) continue;

    const sinInformacion = !serial && !puertoRaw && !modeloD && !modeloE;
    const modelo = determinarModelo_(modeloD, modeloE);
    const puertoUsb = normalizarPuerto_(puertoRaw);
    const observaciones = combinarObservaciones_(obsF, obsG, obsH);
    const estado = normalizarEstado_(estadoRaw) === 'No operativa' ? 'No operativa' : 'Operativa';
    const activoFijo = activoFijoRaw || 'No especificado';

    if (agencia) agenciasSet.add(agencia);
    if (!sinInformacion && modelo && modelo !== 'Sin especificar') modelosSet.add(modelo);

    impresoras.push({
      rowNum: i + 1,
      agencia: agencia,
      serial: serial,
      activoFijo: activoFijo,
      puertoUsb: puertoUsb,
      modelo: modelo,
      modeloD: modeloD,
      modeloE: modeloE,
      observaciones: observaciones,
      puertoRaw: puertoRaw,
      estado: estado,
      sinInformacion: sinInformacion
    });
  }

  const resultado = {
    impresoras: impresoras,
    catalogos: {
      agencias: Array.from(agenciasSet).sort((a,b) => a.localeCompare(b, 'es')),
      modelos: Array.from(modelosSet).sort((a,b) => a.localeCompare(b, 'es'))
    }
  };

  try {
    CacheService.getScriptCache().put(CACHE_KEY, JSON.stringify(resultado), CACHE_TTL);
  } catch(e) {}

  return resultado;
}

function obtenerDatosProveduria() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(CACHE_KEY_PROVEDURIA);
    if (cached) return JSON.parse(cached);
  } catch(e) {}

  const sheet = getSheetProveduria_();
  if (!sheet) return { proveduria: [], correspondencia: [] };

  const data = sheet.getDataRange().getValues();
  const proveduria = [];
  const correspondencia = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const codA = limpiar_(row[0]);
    if (codA) {
      proveduria.push({
        rowNum: i + 1, grupo: 'Proveduria', codigo: codA, activoFijo: codA,
        serial: limpiar_(row[1]), puertoUsb: normalizarPuerto_(row[2]),
        modelo: determinarModelo_(row[3], row[4]), estado: normalizarEstado_(row[5])
      });
    }

    const codCorrespondencia = limpiar_(row[8]); 
    if (codCorrespondencia) {
      correspondencia.push({
        rowNum: i + 1, grupo: 'Correspondencia', codigo: codCorrespondencia, activoFijo: codCorrespondencia,
        serial: limpiar_(row[9]), puertoUsb: normalizarPuerto_(row[10]),
        modelo: determinarModelo_(row[11], row[12]), estado: normalizarEstado_(row[13])
      });
    }
  }

  const resultado = { proveduria, correspondencia };

  try {
    CacheService.getScriptCache().put(CACHE_KEY_PROVEDURIA, JSON.stringify(resultado), CACHE_TTL);
  } catch(e) {}

  return resultado;
}

/* ===================== MOVER IMPRESORA ===================== */
function encontrarFilaVacia_(sheet, startCol) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 2;
  const data = sheet.getRange(2, startCol, lastRow - 1, 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (!limpiar_(data[i][0])) {
      return i + 2;
    }
  }
  return lastRow + 1;
}

function moverImpresora(origen, rowNum, destino, agenciaDestino) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let datos = {};
    // 1. Obtener datos y BORRAR del origen
    if (origen === 'Agencia') {
      const sheet = getSheet_();
      const rowValues = sheet.getRange(rowNum, 1, 1, 10).getValues()[0];
      datos = {
        serial: limpiar_(rowValues[1]), puertoUsb: limpiar_(rowValues[2]),
        modeloD: limpiar_(rowValues[3]), modeloE: limpiar_(rowValues[4]),
        observaciones: limpiar_(rowValues[5]), estado: limpiar_(rowValues[8]),
        activoFijo: limpiar_(rowValues[9])
      };
      sheet.deleteRow(rowNum);
    } else if (origen === 'Proveduria') {
      const sheet = getSheetProveduria_();
      const rowValues = sheet.getRange(rowNum, 1, 1, 7).getValues()[0];
      datos = {
        activoFijo: limpiar_(rowValues[0]), serial: limpiar_(rowValues[1]),
        puertoUsb: limpiar_(rowValues[2]), modeloD: limpiar_(rowValues[3]),
        modeloE: limpiar_(rowValues[4]), estado: limpiar_(rowValues[5]),
        observaciones: limpiar_(rowValues[6])
      };
      sheet.getRange(rowNum, 1, 1, 7).clearContent();
    } else if (origen === 'Correspondencia') {
      const sheet = getSheetProveduria_();
      const rowValues = sheet.getRange(rowNum, 9, 1, 7).getValues()[0];
      datos = {
        activoFijo: limpiar_(rowValues[0]), serial: limpiar_(rowValues[1]),
        puertoUsb: limpiar_(rowValues[2]), modeloD: limpiar_(rowValues[3]),
        modeloE: limpiar_(rowValues[4]), estado: limpiar_(rowValues[5]),
        observaciones: limpiar_(rowValues[6])
      };
      sheet.getRange(rowNum, 9, 1, 7).clearContent();
    }

    // 2. INSERTAR en el destino
    if (destino === 'Agencia') {
      const sheet = getSheet_();
      const modeloTxt = determinarModelo_(datos.modeloD, datos.modeloE);
      const [modD, modE] = modeloValores_(modeloTxt);
      const puerto = normalizarPuerto_(datos.puertoUsb) === 'Sí' ? 'SI' : 'NO';
      const estado = normalizarEstado_(datos.estado) === 'No operativa' ? 'No operativa' : 'Operativa';

      const fila = [
        agenciaDestino, datos.serial || '', puerto, modD, modE, datos.observaciones || '', '', '', estado, datos.activoFijo || ''
      ];
      
      // En lugar de appendRow, buscar la primera fila vacía real
      const targetRow = encontrarFilaVacia_(sheet, 1);
      sheet.getRange(targetRow, 1, 1, 10).setValues([fila]);

    } else if (destino === 'Proveduria' || destino === 'Correspondencia') {
      const sheet = getSheetProveduria_();
      const isProv = destino === 'Proveduria';
      const startCol = isProv ? 1 : 9;
      
      // En lugar de buscar a mano, usar la misma función
      const targetRow = encontrarFilaVacia_(sheet, startCol);

      const modeloTxt = determinarModelo_(datos.modeloD, datos.modeloE);
      const [modD, modE] = modeloValores_(modeloTxt);
      const puerto = normalizarPuerto_(datos.puertoUsb) === 'Sí' ? 'si' : 'no';
      const estado = normalizarEstado_(datos.estado) === 'No operativa' ? 'No Operativo' : 'Operativa'; 

      const fila = [
        datos.activoFijo || '', datos.serial || '', puerto, modD, modE, estado, datos.observaciones || ''
      ];
      sheet.getRange(targetRow, startCol, 1, 7).setValues([fila]);
    }

    invalidarCache_();
    invalidarCacheProveduria_();
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ===================== HELPERS ===================== */
function limpiar_(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number' && Number.isInteger(v)) return String(v);
  return String(v).trim();
}
function limpiarSerial_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (typeof v === 'number' && Number.isInteger(v)) return String(v);
  return String(v).trim();
}
function determinarModelo_(d, e) {
  const valD = String(d || '').trim().toUpperCase();
  const valE = String(e || '').trim().toUpperCase();
  const tieneD = valD && !esNegativo_(valD);
  const tieneE = valE && !esNegativo_(valE);
  if (tieneD && tieneE) return 'PR2E / PR2E Plus';
  if (tieneD) return 'PR2E';
  if (tieneE) return 'PR2E Plus';
  return 'Sin especificar';
}
function esNegativo_(v) {
  return v === 'NO' || v === 'NO TIENE' || v === 'NO POSEE' || v === 'N/A' || v === '-';
}
function normalizarPuerto_(raw) {
  if (raw === null || raw === undefined || String(raw).trim() === '') return 'No especificado';
  const v = String(raw).trim().toUpperCase();
  if (v.includes('NO POSEE') || v.includes('NO TIENE') || v === 'NO' || v === 'N') return 'No';
  if (v === 'X' || v === 'SÍ' || v === 'SI' || v === 'YES' || v === 'Y' ||
      v === '1' || v === 'TRUE' || v.includes('SI POSEE') || v.includes('SI TIENE') ||
      v.includes('USB') || v === 'S') {
    return 'Sí';
  }
  return 'No';
}
function normalizarEstado_(raw) {
  if (!raw) return 'Sin especificar';
  const v = String(raw).trim().toUpperCase();
  if (v.includes('NO OPERATIVA') || v.includes('NO OPERATIVO') ||
      v.includes('INOPERATIVA') || v.includes('INOPERATIVO')) {
    return 'No operativa';
  }
  if (v.includes('OPERATIVA') || v.includes('OPERATIVO')) return 'Operativa';
  return 'Sin especificar';
}
function combinarObservaciones_(f, g, h) {
  return [f, g, h].map(x => String(x || '').trim()).filter(Boolean).join(' · ');
}
function modeloValores_(modelo) {
  const m = String(modelo || '').toUpperCase().trim();
  if (m === 'PR2E') return ['PR2E', ''];
  if (m === 'PR2E PLUS' || m === 'PR2PLUS' || m === 'PR2E_PLUS') return ['', 'PR2E Plus'];
  if (m === 'AMBOS' || m === 'PR2E / PR2E PLUS' || m === 'PR2E/PR2E PLUS') return ['PR2E', 'PR2E Plus'];
  return ['', ''];
}

function guardarImpresora(datos) {
  if (!datos || !datos.agencia) throw new Error('La agencia es obligatoria');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const [modD, modE] = modeloValores_(datos.modelo);
    const puerto = datos.puertoUsb === 'Sí' ? 'SI' : 'NO';
    const estado = datos.estado === 'No operativa' ? 'No operativa' : 'Operativa';
    const activoFijo = String(datos.activoFijo || '').trim();

    const fila = [
      String(datos.agencia).trim(), String(datos.serial || '').trim(), puerto, modD, modE,
      String(datos.observaciones || '').trim(), '', '', estado, activoFijo
    ];

    const lastRow = sheet.getLastRow();
    if (datos.rowNum && datos.rowNum > 1 && datos.rowNum <= lastRow) {
      sheet.getRange(datos.rowNum, 1, 1, fila.length).setValues([fila]);
    } else {
      sheet.appendRow(fila);
    }
    invalidarCache_();
    return obtenerDatosIniciales();
  } finally {
    lock.releaseLock();
  }
}

function actualizarEstadoImpresora(rowNum, nuevoEstado) {
  if (!rowNum) throw new Error('Fila inválida');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (rowNum <= 1 || rowNum > lastRow) throw new Error('Fila fuera de rango');
    const estadoNorm = normalizarEstado_(nuevoEstado) === 'No operativa' ? 'No operativa' : 'Operativa';
    sheet.getRange(rowNum, 9).setValue(estadoNorm);
    invalidarCache_();
    return obtenerDatosIniciales();
  } finally {
    lock.releaseLock();
  }
}

function actualizarEstadoProveduria(grupo, rowNum, nuevoEstado) {
  if (!rowNum) throw new Error('Fila inválida');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheetProveduria_();
    if (!sheet) throw new Error('No se encontró la hoja "Proveduria"');
    const lastRow = sheet.getLastRow();
    if (rowNum <= 1 || rowNum > lastRow) throw new Error('Fila fuera de rango');
    const col = String(grupo).toLowerCase() === 'correspondencia' ? 14 : 6; 
    const estadoNorm = normalizarEstado_(nuevoEstado) === 'No operativa' ? 'No operativa' : 'Operativa';
    sheet.getRange(rowNum, col).setValue(estadoNorm);
    invalidarCacheProveduria_();
    return obtenerDatosProveduria();
  } finally {
    lock.releaseLock();
  }
}

function eliminarImpresora(rowNum) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();
    if (rowNum && rowNum > 1 && rowNum <= lastRow) sheet.deleteRow(rowNum);
    invalidarCache_();
    return obtenerDatosIniciales();
  } finally {
    lock.releaseLock();
  }
}

function formatearHoja() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet.getRange(1, 1, 1, 10)
    .setFontWeight('bold')
    .setBackground('#3b82f6')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(1);
}

/* SDK Y CONFIG */
const SDK_SPREADSHEET_ID = '1qMTkBZoXnYLIq4yAqiRG5cxnlhtn5ur0DY0bhC6ZVjI';
const SDK_MESES_VALIDOS = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre', '2026'];

function _esHojaMensualSDK_(nombreHoja) {
  const n = String(nombreHoja || '').trim().toLowerCase();
  if (n.startsWith('manual')) return false;
  if (n === 'valores' || n === 'bs' || n === 'control_migraciones' || n === 'observaciones') return false;
  return SDK_MESES_VALIDOS.some(m => n.includes(m));
}

function obtenerIncidenciasOlivettiSDK() {
  try {
    const ss = SpreadsheetApp.openById(SDK_SPREADSHEET_ID);
    const sheets = ss.getSheets();
    const incidencias = [];

    for (let s = 0; s < sheets.length; s++) {
      const sheet = sheets[s];
      const sheetName = sheet.getName().trim();
      if (!_esHojaMensualSDK_(sheetName)) continue;
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) continue;

      const data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
      for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const rawFecha = row[0];
        if (!rawFecha) continue;
        const fechaStr = String(rawFecha).toLowerCase();
        if (fechaStr.includes('fecha') || fechaStr.includes('llamads')) continue;
        if (!row[1] && !row[2]) continue;

        const incidenteVal = String(row[5] || '').trim();
        const filaTexto = row.map(c => String(c || '').toLowerCase()).join(' ');
        const esOlivetti = filaTexto.includes('olivett');
        const normalizadoFila = filaTexto.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
        const esLaser = /\blasser\b|\blazers?\b/.test(normalizadoFila);
        if (!esOlivetti || esLaser) continue;

        const analistaVal = String(row[7] || '').trim();
        if (['colaboardor', 'colaborador'].includes(analistaVal.toLowerCase())) continue;

        let fechaFormatted = '';
        if (rawFecha instanceof Date) {
          const y = rawFecha.getFullYear();
          const m = String(rawFecha.getMonth() + 1).padStart(2, '0');
          const d = String(rawFecha.getDate()).padStart(2, '0');
          fechaFormatted = `${y}-${m}-${d}`;
        } else {
          fechaFormatted = String(rawFecha).substring(0, 10);
        }

        incidencias.push({
          hoja: sheetName, fila: i + 2, fecha: fechaFormatted,
          bs: String(row[1] || '').trim(), nombre: String(row[2] || '').trim(),
          area: String(row[3] || '').trim(), cargo: String(row[4] || '').trim(),
          incidente: incidenteVal, canal: String(row[6] || '').trim(), colaborador: analistaVal
        });
      }
    }
    incidencias.sort((a, b) => b.fecha.localeCompare(a.fecha));
    return { incidencias: incidencias };
  } catch (e) {
    return { incidencias: [], error: 'No se pudo conectar al ServiceDesk: ' + e.message };
  }
}
