/**
 * VEVA Enterprise — Svc_Admin.js  v1.0.0
 * CRUD para: MAPEO_FLUJO, CAT_CUENTAS_MAPEO
 *
 * Funciones públicas:
 *   getMapeoFlujo, saveMapeoFlujoItem, deleteMapeoFlujoItem
 *   getCatCuentasAdmin, saveCatCuenta, deleteCatCuenta
 */

// ─── MAPEO_FLUJO ─────────────────────────────────────────────────────────────

function getMapeoFlujo() {
  try {
    Session.getActiveUser().getEmail();
    var ss    = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'admin-flujo')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var sheet = _ensureMapeoFlujoSheet(ss);
    if (sheet.getLastRow() <= 1) return { status: 'success', data: [], message: '' };

    var data   = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      result.push({
        tipo:          String(row[0] || '').trim(),
        clasificacion: String(row[1] || '').trim(),
        clasificacion2:String(row[2] || '').trim(),
        orden:         parseInt(row[3]) || 0,
        activo:        row[4] === true || String(row[4]).toUpperCase() === 'TRUE'
      });
    }
    result.sort(function(a, b) {
      var ta = a.tipo + '|' + a.clasificacion + '|' + ('0000' + a.orden).slice(-4);
      var tb = b.tipo + '|' + b.clasificacion + '|' + ('0000' + b.orden).slice(-4);
      return ta < tb ? -1 : 1;
    });
    return { status: 'success', data: result, message: '' };
  } catch(e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

/**
 * Agrega o actualiza una fila en MAPEO_FLUJO.
 * payload: { tipo, clasificacion, clasificacion2, orden, activo }
 * La clave de identidad es tipo+clasificacion+clasificacion2 (los tres juntos).
 */
function saveMapeoFlujoItem(payload) {
  try {
    Session.getActiveUser().getEmail();
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var ss    = SpreadsheetApp.openById(SALDOS_SHEET_ID);
      if (!_tieneAccesoAVista(ss, 'admin-flujo')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
      var sheet = _ensureMapeoFlujoSheet(ss);
      var tipo  = String(payload.tipo           || '').trim();
      var clasif = String(payload.clasificacion  || '').trim();
      var c2    = String(payload.clasificacion2 || '').trim();
      if (!tipo || !clasif || !c2) return { status: 'error', data: {}, message: 'tipo, clasificacion y clasificacion2 requeridos' };

      var data  = sheet.getLastRow() > 1
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues()
        : [];
      var found = -1;
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0]||'').trim() === tipo &&
            String(data[i][1]||'').trim() === clasif &&
            String(data[i][2]||'').trim() === c2) { found = i; break; }
      }
      var fila = [tipo, clasif, c2, parseInt(payload.orden) || 99, payload.activo !== false];
      if (found >= 0) { data[found] = fila; }
      else            { data.push(fila); }

      if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).clearContent();
      if (data.length) sheet.getRange(2, 1, data.length, 5).setValues(data);
      return { status: 'success', data: {}, message: 'Guardado' };
    } finally {
      lock.releaseLock();
    }
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  }
}

/**
 * Elimina una fila de MAPEO_FLUJO por clave tipo+clasificacion+clasificacion2.
 */
function deleteMapeoFlujoItem(payload) {
  try {
    Session.getActiveUser().getEmail();
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var ss    = SpreadsheetApp.openById(SALDOS_SHEET_ID);
      if (!_tieneAccesoAVista(ss, 'admin-flujo')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
      var sheet = _ensureMapeoFlujoSheet(ss);
      if (sheet.getLastRow() <= 1) return { status: 'success', data: {}, message: 'Sin datos' };

      var tipo  = String(payload.tipo           || '').trim();
      var clasif = String(payload.clasificacion  || '').trim();
      var c2    = String(payload.clasificacion2 || '').trim();
      var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
      var nueva = [];
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0]||'').trim() === tipo &&
            String(data[i][1]||'').trim() === clasif &&
            String(data[i][2]||'').trim() === c2) continue;
        nueva.push(data[i]);
      }
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).clearContent();
      if (nueva.length) sheet.getRange(2, 1, nueva.length, 5).setValues(nueva);
      return { status: 'success', data: {}, message: 'Eliminado' };
    } finally {
      lock.releaseLock();
    }
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  }
}

// ─── CAT_CUENTAS_MAPEO ────────────────────────────────────────────────────────

/**
 * Devuelve todas las cuentas del catálogo para el UI de administración.
 */
function getCatCuentasAdmin() {
  try {
    Session.getActiveUser().getEmail();
    var ss    = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'admin-flujo')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var sheet = ss.getSheetByName('CAT_CUENTAS_MAPEO');
    if (!sheet || sheet.getLastRow() <= 1) return { status: 'success', data: [], message: '' };

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var hi      = {};
    headers.forEach(function(h, i) { if (h) hi[String(h).trim().toUpperCase()] = i; });
    function col(name, fb) { return hi[name] !== undefined ? hi[name] : fb; }

    var iCuenta = col('NUMERO_CUENTA', 1);
    var iBanco  = col('BANCO', 2);
    var iIdSoc  = col('ID_SOCIEDAD', 3);
    var iNomSoc = col('NOMBRE_SOCIEDAD', 4);
    var iMon    = col('MONEDA', 5);
    var iTipo   = col('TIPO_CUENTA', 6);
    var iTipo2  = col('TIPO_CUENTA2', 7);
    var iNCorto = col('NOMBRE_CORTO', 8);
    var iAbr    = col('ABR_COBRANZA', 11);

    var data   = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
    var result = [];
    for (var i = 0; i < data.length; i++) {
      var row    = data[i];
      var cuenta = String(row[iCuenta] || '').trim();
      if (!cuenta) continue;
      result.push({
        cuenta:         cuenta,
        banco:          String(row[iBanco]  || '').trim(),
        idSociedad:     String(row[iIdSoc]  || '').trim(),
        nombreSociedad: String(row[iNomSoc] || '').trim(),
        moneda:         String(row[iMon]    || 'MXN').trim(),
        tipoCuenta:     String(row[iTipo]   || '').trim(),
        tipoCuenta2:    String(row[iTipo2]  || '').trim(),
        nombreCorto:    String(row[iNCorto] || '').trim(),
        abrCobranza:    String(row[iAbr]    || '').trim()
      });
    }
    return { status: 'success', data: result, message: '' };
  } catch(e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

/**
 * Agrega o actualiza una cuenta en CAT_CUENTAS_MAPEO.
 * La clave es NUMERO_CUENTA.
 * payload: { cuenta, banco, idSociedad, nombreSociedad, moneda, tipoCuenta, tipoCuenta2, nombreCorto, abrCobranza }
 */
function saveCatCuenta(payload) {
  try {
    Session.getActiveUser().getEmail();
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var ss    = SpreadsheetApp.openById(SALDOS_SHEET_ID);
      if (!_tieneAccesoAVista(ss, 'admin-flujo')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
      var sheet = ss.getSheetByName('CAT_CUENTAS_MAPEO');
      if (!sheet) return { status: 'error', data: {}, message: 'CAT_CUENTAS_MAPEO no existe' };

      var cuenta = String(payload.cuenta || '').trim();
      if (!cuenta) return { status: 'error', data: {}, message: 'cuenta requerida' };

      var lastCol = sheet.getLastColumn();
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var hi      = {};
      headers.forEach(function(h, i) { if (h) hi[String(h).trim().toUpperCase()] = i; });
      function col(name, fb) { return hi[name] !== undefined ? hi[name] : fb; }

      var iCuenta = col('NUMERO_CUENTA', 1);
      var iBanco  = col('BANCO', 2);
      var iIdSoc  = col('ID_SOCIEDAD', 3);
      var iNomSoc = col('NOMBRE_SOCIEDAD', 4);
      var iMon    = col('MONEDA', 5);
      var iTipo   = col('TIPO_CUENTA', 6);
      var iTipo2  = col('TIPO_CUENTA2', 7);
      var iNCorto = col('NOMBRE_CORTO', 8);
      var iAbr    = col('ABR_COBRANZA', 11);

      var data  = sheet.getLastRow() > 1
        ? sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues()
        : [];
      var found = -1;
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][iCuenta] || '').trim() === cuenta) { found = i; break; }
      }

      var fila = found >= 0 ? data[found].slice() : (function(n){ var a=[]; for(var x=0;x<n;x++) a.push(''); return a; })(lastCol);
      fila[iCuenta] = cuenta;
      fila[iBanco]  = String(payload.banco          || '').trim();
      fila[iIdSoc]  = String(payload.idSociedad     || '').trim();
      fila[iNomSoc] = String(payload.nombreSociedad || '').trim();
      fila[iMon]    = String(payload.moneda         || 'MXN').trim();
      fila[iTipo]   = String(payload.tipoCuenta     || '').trim();
      fila[iTipo2]  = String(payload.tipoCuenta2    || '').trim();
      fila[iNCorto] = String(payload.nombreCorto    || '').trim();
      if (iAbr >= 0) fila[iAbr] = String(payload.abrCobranza || '').trim();

      if (found >= 0) { data[found] = fila; }
      else            { data.push(fila); }

      if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).clearContent();
      if (data.length) sheet.getRange(2, 1, data.length, lastCol).setValues(data);
      return { status: 'success', data: {}, message: 'Cuenta guardada' };
    } finally {
      lock.releaseLock();
    }
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  }
}

/**
 * Elimina una cuenta de CAT_CUENTAS_MAPEO por NUMERO_CUENTA.
 * payload: { cuenta }
 */
function deleteCatCuenta(payload) {
  try {
    Session.getActiveUser().getEmail();
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var ss    = SpreadsheetApp.openById(SALDOS_SHEET_ID);
      if (!_tieneAccesoAVista(ss, 'admin-flujo')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
      var sheet = ss.getSheetByName('CAT_CUENTAS_MAPEO');
      if (!sheet || sheet.getLastRow() <= 1) return { status: 'success', data: {}, message: 'Sin datos' };

      var cuenta  = String(payload.cuenta || '').trim();
      var lastCol = sheet.getLastColumn();
      var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      var hi      = {};
      headers.forEach(function(h, i) { if (h) hi[String(h).trim().toUpperCase()] = i; });
      var iCuenta = hi['NUMERO_CUENTA'] !== undefined ? hi['NUMERO_CUENTA'] : 1;

      var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
      var nueva = data.filter(function(row) { return String(row[iCuenta] || '').trim() !== cuenta; });

      sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).clearContent();
      if (nueva.length) sheet.getRange(2, 1, nueva.length, lastCol).setValues(nueva);
      return { status: 'success', data: {}, message: 'Cuenta eliminada' };
    } finally {
      lock.releaseLock();
    }
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  }
}
