/**
 * VEVA Enterprise — Svc_Flujo.js  v1.0.0
 * Backend para módulo Flujo de Efectivo (método directo).
 * Funciones públicas: getFlujoEfectivo, getSaldoInicialAdmin, setSaldoInicial, bloquearPeriodoSaldo
 *
 * Hoja SALDO_INICIAL — 7 cols (A–G):
 *   A=CUENTA  B=BANCO  C=SOCIEDAD  D=PERIODO  E=SALDO  F=BLOQUEADO  G=FECHA_ACTUALIZACION
 *
 * Hoja MAPEO_FLUJO — 5 cols (A–E):
 *   A=TIPO  B=CLASIFICACION  C=CLASIFICACION2  D=ORDEN  E=ACTIVO
 */

// Derivacion fallback cuando SUBCATEGORIA de HIST está vacía
var _FLUJO_FALLBACK = {
  'COBRANZA':           { clasif: 'OPERATIVO',      clasif2: 'Cobranza' },
  'OPERATIVOS':         { clasif: 'OPERATIVO',      clasif2: 'Cobranza' },
  'DEPOSITOS':          { clasif: 'OPERATIVO',      clasif2: 'Cobranza' },
  'PROVEEDORES':        { clasif: 'OPERATIVO',      clasif2: 'Proveedores' },
  'NOMINA':             { clasif: 'OPERATIVO',      clasif2: 'Nominas' },
  'COMISIONES':         { clasif: 'OPERATIVO',      clasif2: 'Comisiones Bancarias' },
  'IMPUESTOS':          { clasif: 'IMPUESTOS FEDE', clasif2: 'IVA' },
  'INVERSION':          { clasif: 'FINANCIEROS',    clasif2: 'Inversiones' },
  'INVERSIONES':        { clasif: 'FINANCIEROS',    clasif2: 'Inversiones' },
  'PRESTAMOS':          { clasif: 'FINANCIEROS',    clasif2: 'Prestamos bancarios' },
  'PRESTAMOS_INTERNOS': { clasif: 'FINANCIEROS',    clasif2: 'Prestamos bancarios' },
  'TRASPASOS_INTERNOS': { clasif: 'NO OPERATIVO',   clasif2: 'Traspasos entre cuentas' },
  'INTERCIAS':          { clasif: 'INTERCIAS',      clasif2: 'Prestamos Intercias' },
  'CAPEX':              { clasif: 'CAPEX',           clasif2: 'Capex' }
};

// Secciones en orden de presentacion
var _FLUJO_SECCIONES_IN  = ['OPERATIVO', 'INTERCIAS', 'NO OPERATIVO', 'FINANCIEROS'];
var _FLUJO_SECCIONES_EG  = ['OPERATIVO', 'NO OPERATIVO', 'CAPEX', 'FINANCIEROS', 'INTERCIAS', 'IMPUESTOS FEDE', 'IMPUESTOS LOCA'];

function _leerMapeoFlujo(ss) {
  var sheet = ss.getSheetByName('MAPEO_FLUJO');
  if (!sheet || sheet.getLastRow() <= 1) return {};
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  // map: clasif2.toUpperCase() → { tipo, clasificacion, orden }
  var m = {};
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    if (!row[4]) continue; // ACTIVO=false
    var c2 = String(row[2] || '').trim().toUpperCase();
    if (!c2) continue;
    m[c2] = { tipo: String(row[0] || '').trim(), clasificacion: String(row[1] || '').trim(), orden: parseInt(row[3]) || 99 };
  }
  return m;
}

function _resolverClasif(clasificacion, subcategoria, tipo, mapeoFlujo) {
  // 1. SUBCATEGORIA del HIST tiene valor: buscarlo en MAPEO_FLUJO
  var sub = (subcategoria || '').trim();
  if (sub) {
    var found = mapeoFlujo[sub.toUpperCase()];
    if (found) return { clasif: found.clasificacion, clasif2: sub };
  }
  // 2. Fallback desde CLASIFICACION actual
  var cl = (clasificacion || '').trim().toUpperCase().replace(/\s+/g, '_');
  var fb = _FLUJO_FALLBACK[cl];
  if (fb) return { clasif: fb.clasif, clasif2: fb.clasif2 };
  // 3. Default: sin mapeo conocido → sección visible para revisión
  return {
    clasif:  'NO OPERATIVO',
    clasif2: 'Sin Clasificar'
  };
}

function _nodoVacio(sociedades) {
  var n = { _total: 0 };
  for (var i = 0; i < sociedades.length; i++) n[sociedades[i]] = 0;
  return n;
}

function _sumar(nodo, soc, monto) {
  nodo._total += monto;
  if (nodo[soc] === undefined) nodo[soc] = 0;
  nodo[soc] += monto;
}

/**
 * Flujo de efectivo — método directo.
 * Lee SUM_EFE (pre-agregado) en lugar de HIST — respuesta <1s.
 * params: { desde, hasta }  (YYYY-MM-DD)
 */
function getFlujoEfectivo(params) {
  try {
    Session.getActiveUser().getEmail();
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);

    var desde = params && params.desde ? String(params.desde).trim() : '';
    var hasta = params && params.hasta ? String(params.hasta).trim() : '';
    if (!desde || !hasta) return { status: 'error', data: {}, message: 'Se requiere desde y hasta' };

    var desdePerio = desde.slice(0, 7);
    var hastaPerio = hasta.slice(0, 7);

    var efeSheet = ss.getSheetByName('SUM_EFE');
    if (!efeSheet || efeSheet.getLastRow() <= 1) {
      return { status: 'success', data: _flujoVacio(), message: 'Sin datos — recarga archivos para generar SUM_EFE' };
    }

    var mapeoFlujo = _leerMapeoFlujo(ss);

    // Lectura única de SUM_EFE (~4k filas max — siempre rápido)
    var efeRaw = efeSheet.getRange(2, 1, efeSheet.getLastRow() - 1, _SUM_EFE_COLS).getValues();

    var PART_C2S = {
      'Traspasos entre cuentas':         true,
      'Aportaciones capital Intercia':   true,
      'Inversiones en subsidiaria':      true,
      'Prestamos Intercias':             true,
      'Intereses partes relacionadas':   true,
      'Venta de Activo Fijo Intercias':  true,
      'Venta de Activo Fijo':            true,
      'Inversiones':                     true
    };

    // Pasada única: sociedades + cuentas + agregados EFE + partidas
    var socSet       = {};
    var cuentaPorSoc = {};
    var agIN = {}, agEG = {};
    var partIN = {}, partEG = {};
    var ctasIN = {}, ctasEG = {};

    for (var i = 0; i < efeRaw.length; i++) {
      var row  = efeRaw[i];
      var rPer = _periodoStr(row[3]);
      if (rPer < desdePerio || rPer > hastaPerio) continue;

      var cuenta = String(row[0] || '').trim();
      var banco  = String(row[1] || '').trim();
      var soc    = String(row[2] || '').trim();
      var tipo    = String(row[4] || '').trim();
      var sec     = String(row[5] || '').trim();
      var c2      = String(row[6] || '').trim();
      var cliente = String(row[7] || '').trim();
      var monto   = parseFloat(row[8]) || 0;
      if (!soc || (tipo !== 'INGRESO' && tipo !== 'EGRESO') || monto <= 0) continue;

      socSet[soc] = true;
      if (cuenta) {
        if (!cuentaPorSoc[soc]) cuentaPorSoc[soc] = {};
        if (!cuentaPorSoc[soc][cuenta]) {
          var ult4 = cuenta.length > 4 ? cuenta.slice(-4) : cuenta;
          cuentaPorSoc[soc][cuenta] = { cuenta: cuenta, label: (banco || '?') + ' *' + ult4 };
        }
      }

      var agBag  = tipo === 'INGRESO' ? agIN : agEG;
      var ctaBag = tipo === 'INGRESO' ? ctasIN : ctasEG;
      if (!agBag[sec])     agBag[sec]     = {};
      if (!agBag[sec][c2]) agBag[sec][c2] = _nodoVacio([]);
      _sumar(agBag[sec][c2], soc, monto);

      if (!ctaBag[c2])          ctaBag[c2] = {};
      if (!ctaBag[c2][soc])     ctaBag[c2][soc] = {};
      if (!ctaBag[c2][soc][cuenta]) ctaBag[c2][soc][cuenta] = 0;
      ctaBag[c2][soc][cuenta] += monto;

      if (PART_C2S[c2]) {
        var partBag = tipo === 'INGRESO' ? partIN : partEG;
        if (!partBag[c2]) partBag[c2] = _nodoVacio([]);
        _sumar(partBag[c2], soc, monto);
      }
    }

    var sociedades = Object.keys(socSet).sort();

    // Ajustar _total de nodos a la suma de sociedades
    function _ajustarNodo(nodo) {
      nodo._total = 0;
      for (var s in nodo) {
        if (s !== '_total') nodo._total += (nodo[s] || 0);
      }
    }
    function _ajustarBag(bag) {
      for (var sec in bag) {
        for (var c2 in bag[sec]) _ajustarNodo(bag[sec][c2]);
      }
    }
    function _ajustarPartBag(bag) {
      for (var c2 in bag) _ajustarNodo(bag[c2]);
    }
    _ajustarBag(agIN); _ajustarBag(agEG);
    _ajustarPartBag(partIN); _ajustarPartBag(partEG);

    var totalIN = _nodoVacio(sociedades);
    var totalEG = _nodoVacio(sociedades);

    function _serializarSecciones(agBag, secciones, totalNodo, ctaBag) {
      var result = [];
      for (var si = 0; si < secciones.length; si++) {
        var sec   = secciones[si];
        var items = agBag[sec];
        if (!items) continue;
        var secNodo  = _nodoVacio(sociedades);
        var itemsArr = [];
        var c2keys   = Object.keys(items).sort(function(a, b) {
          var oa = (mapeoFlujo[a.toUpperCase()] || {}).orden || 99;
          var ob = (mapeoFlujo[b.toUpperCase()] || {}).orden || 99;
          return oa - ob;
        });
        for (var ci = 0; ci < c2keys.length; ci++) {
          var c2   = c2keys[ci];
          var nodo = items[c2];
          secNodo._total += (nodo._total || 0);
          for (var si2 = 0; si2 < sociedades.length; si2++) {
            var s2 = sociedades[si2];
            if (secNodo[s2] === undefined) secNodo[s2] = 0;
            secNodo[s2] += (nodo[s2] || 0);
          }
          var cuentasArr = [];
          if (ctaBag[c2]) {
            var csocs = Object.keys(ctaBag[c2]);
            for (var csoci = 0; csoci < csocs.length; csoci++) {
              var csoc   = csocs[csoci];
              var cctMap = ctaBag[c2][csoc];
              var ccts   = Object.keys(cctMap);
              for (var ccti = 0; ccti < ccts.length; ccti++) {
                var cct   = ccts[ccti];
                var u4    = cct.length > 4 ? cct.slice(-4) : cct;
                var lbl   = (cuentaPorSoc[csoc] && cuentaPorSoc[csoc][cct])
                              ? cuentaPorSoc[csoc][cct].label : '? *' + u4;
                cuentasArr.push({ cuenta: cct, label: lbl,
                                  sociedad: csoc, monto: cctMap[cct] });
              }
            }
          }
          itemsArr.push({ clasificacion2: c2, nodo: nodo, cuentas: cuentasArr });
        }
        totalNodo._total += secNodo._total;
        for (var ts = 0; ts < sociedades.length; ts++) {
          if (totalNodo[sociedades[ts]] === undefined) totalNodo[sociedades[ts]] = 0;
          totalNodo[sociedades[ts]] += (secNodo[sociedades[ts]] || 0);
        }
        result.push({ clasificacion: sec, nodo: secNodo, items: itemsArr });
      }
      return result;
    }

    var seccionesIN = _serializarSecciones(agIN, _FLUJO_SECCIONES_IN, totalIN, ctasIN);
    var seccionesEG = _serializarSecciones(agEG, _FLUJO_SECCIONES_EG, totalEG, ctasEG);

    // Saldo inicial: SUM_MOV col T (SALDO_BANCO_FINAL) del periodo inmediato anterior
    var efectivoInicial = _nodoVacio(sociedades);
    var mesAnterior     = _periodoAnterior(desdePerio);
    var sumSheet        = ss.getSheetByName('SUM_MOV');
    if (sumSheet && sumSheet.getLastRow() > 1) {
      var sumRaw = sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, _SUM_COLS).getValues();
      for (var smi = 0; smi < sumRaw.length; smi++) {
        var smPer   = _periodoStr(sumRaw[smi][3]);
        if (smPer !== mesAnterior) continue;
        var smSoc   = String(sumRaw[smi][2] || '').trim();
        var smSaldo = parseFloat(sumRaw[smi][19]); // col T idx=19: SALDO_BANCO_FINAL
        if (smSoc && !isNaN(smSaldo) && smSaldo !== 0) _sumar(efectivoInicial, smSoc, smSaldo);
      }
    }
    // Fallback a SALDO_INICIAL manual para primer periodo (sin SUM_MOV previo)
    if (efectivoInicial._total === 0) {
      var siSheet = ss.getSheetByName('SALDO_INICIAL');
      if (siSheet && siSheet.getLastRow() > 1) {
        var siData = siSheet.getRange(2, 1, siSheet.getLastRow() - 1, 7).getValues();
        for (var sii = 0; sii < siData.length; sii++) {
          var siPer   = _periodoStr(siData[sii][3]);
          var siSaldo = parseFloat(siData[sii][4]) || 0;
          var siSoc   = String(siData[sii][2] || '').trim();
          if (siPer === desdePerio && siSoc) _sumar(efectivoInicial, siSoc, siSaldo);
        }
      }
    }

    var flujoEfectivo = _nodoVacio(sociedades);
    var efectivoFinal = _nodoVacio(sociedades);
    flujoEfectivo._total = totalIN._total - totalEG._total;
    efectivoFinal._total = efectivoInicial._total + flujoEfectivo._total;
    for (var fs2 = 0; fs2 < sociedades.length; fs2++) {
      var s2 = sociedades[fs2];
      flujoEfectivo[s2] = (totalIN[s2] || 0) - (totalEG[s2] || 0);
      efectivoFinal[s2] = (efectivoInicial[s2] || 0) + flujoEfectivo[s2];
    }

    var partidas = _construirPartidas(partIN, partEG, sociedades);

    return {
      status: 'success',
      data: {
        filtro:          { desde: desde, hasta: hasta },
        sociedades:      sociedades,
        cuentasPorSoc:   cuentaPorSoc,
        efectivoInicial: efectivoInicial,
        ingresos:        { nodo: totalIN, secciones: seccionesIN },
        egresos:         { nodo: totalEG, secciones: seccionesEG },
        flujoEfectivo:   flujoEfectivo,
        efectivoFinal:   efectivoFinal,
        partidas:        partidas
      },
      message: ''
    };
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  }
}

/**
 * Devuelve movimientos individuales de HIST para drill-down.
 * params: { desde, hasta, clasif2?, sociedad?, cuenta? }
 */
function getMovimientosDetalle(params) {
  try {
    Session.getActiveUser().getEmail();
    var ss       = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    var desde    = params && params.desde    ? String(params.desde).trim()    : '';
    var hasta    = params && params.hasta    ? String(params.hasta).trim()    : '';
    var clasif2  = params && params.clasif2  ? String(params.clasif2).trim()  : '';
    var sociedad = params && params.sociedad ? String(params.sociedad).trim() : '';
    var cuenta   = params && params.cuenta   ? String(params.cuenta).trim()   : '';
    if (!desde || !hasta) return { status: 'error', data: [], message: 'Se requiere desde y hasta' };

    var histSheet = ss.getSheetByName('HIST_MOVIMIENTOS');
    if (!histSheet || histSheet.getLastRow() <= 1) return { status: 'success', data: [], message: '' };

    var catMap     = _buildCatalogMap(ss);
    var mapeoFlujo = _leerMapeoFlujo(ss);
    var raw        = histSheet.getRange(2, 1, histSheet.getLastRow() - 1, _HIST_COLS).getValues();
    var result     = [];

    for (var i = 0; i < raw.length; i++) {
      var row  = raw[i];
      var fs   = _histFechaStr(row[1]);
      if (fs < desde || fs > hasta) continue;
      var rtipo = String(row[12] || '').trim();
      if (rtipo !== 'INGRESO' && rtipo !== 'EGRESO') continue;
      var rsoc = String(row[2] || '').trim();
      var rcta = String(row[4] || '').trim();
      if (sociedad && rsoc !== sociedad) continue;
      if (cuenta   && rcta !== cuenta)   continue;

      var rclasif  = String(row[13] || '').trim();
      var rsub     = String(row[14] || '').trim();
      var resolved = _resolverClasif(rclasif, rsub, rtipo, mapeoFlujo);
      if (clasif2 && resolved.clasif2 !== clasif2) continue;

      var rMoneda = (String(row[22] || 'MXN').trim().toUpperCase() || 'MXN');
      var rTc     = parseFloat(row[23]) || 1;
      var rCargo  = parseFloat(row[6]) || 0;
      var rAbono  = parseFloat(row[7]) || 0;
      result.push({
        fecha:       fs,
        sociedad:    rsoc,
        banco:       String(row[3] || '').trim(),
        cuenta:      rcta,
        descripcion: String(row[5] || '').trim(),
        cargo:       rCargo,
        abono:       rAbono,
        cargoMxn:    Math.round(rCargo * rTc * 100) / 100,
        abonoMxn:    Math.round(rAbono * rTc * 100) / 100,
        moneda:      rMoneda,
        tc:          rTc,
        tipo:        rtipo,
        clasif:      resolved.clasif,
        clasif2:     resolved.clasif2,
        cliente:     String(row[21] || '').trim(),
        archivo:     String(row[17] || '').trim()
      });
    }
    result.sort(function(a, b) { return a.fecha < b.fecha ? -1 : a.fecha > b.fecha ? 1 : 0; });
    return { status: 'success', data: result, message: result.length + ' movimientos' };
  } catch(e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

function _periodoAnterior(periodoYYYYMM) {
  var y = parseInt(periodoYYYYMM.slice(0, 4));
  var m = parseInt(periodoYYYYMM.slice(5, 7)) - 1;
  if (m === 0) { m = 12; y -= 1; }
  return y + '-' + (m < 10 ? '0' + m : String(m));
}

function _histFechaStr(raw) {
  return raw instanceof Date
    ? Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
    : String(raw || '').slice(0, 10);
}

function _construirPartidas(partIN, partEG, sociedades) {
  var grupos = [
    { id: 'traspasos',   keys: ['Traspasos entre cuentas'] },
    { id: 'intercias',   keys: ['Aportaciones capital Intercia', 'Inversiones en subsidiaria', 'Prestamos Intercias', 'Intereses partes relacionadas', 'Venta de Activo Fijo Intercias', 'Venta de Activo Fijo'] },
    { id: 'inversiones', keys: ['Inversiones'] }
  ];
  var result = [];
  for (var gi = 0; gi < grupos.length; gi++) {
    var g    = grupos[gi];
    var gIN  = _nodoVacio(sociedades);
    var gEG  = _nodoVacio(sociedades);
    for (var ki = 0; ki < g.keys.length; ki++) {
      var k = g.keys[ki];
      if (partIN[k]) { gIN._total += partIN[k]._total; for (var s2i = 0; s2i < sociedades.length; s2i++) { gIN[sociedades[s2i]] = (gIN[sociedades[s2i]] || 0) + (partIN[k][sociedades[s2i]] || 0); } }
      if (partEG[k]) { gEG._total += partEG[k]._total; for (var s3i = 0; s3i < sociedades.length; s3i++) { gEG[sociedades[s3i]] = (gEG[sociedades[s3i]] || 0) + (partEG[k][sociedades[s3i]] || 0); } }
    }
    var neto = _nodoVacio(sociedades);
    neto._total = gIN._total - gEG._total;
    for (var s4i = 0; s4i < sociedades.length; s4i++) { neto[sociedades[s4i]] = (gIN[sociedades[s4i]] || 0) - (gEG[sociedades[s4i]] || 0); }
    result.push({ id: g.id, ingresos: gIN, egresos: gEG, neto: neto });
  }
  return result;
}

function _flujoVacio() {
  return { filtro: {}, sociedades: [], cuentasPorSoc: {}, efectivoInicial: { _total: 0 },
           ingresos: { nodo: { _total: 0 }, secciones: [] }, egresos: { nodo: { _total: 0 }, secciones: [] },
           flujoEfectivo: { _total: 0 }, efectivoFinal: { _total: 0 }, partidas: [] };
}

// ─── Admin: SALDO_INICIAL ────────────────────────────────────────────────────

/**
 * Devuelve entradas de SALDO_INICIAL para el UI de administración.
 * params: { periodo } filtra por periodo YYYY-MM (opcional)
 */
function getSaldoInicialAdmin(params) {
  try {
    Session.getActiveUser().getEmail();
    var ss      = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    var siSheet = _ensureSaldoInicialSheet(ss);
    if (siSheet.getLastRow() <= 1) return { status: 'success', data: [], message: '' };

    var data    = siSheet.getRange(2, 1, siSheet.getLastRow() - 1, 7).getValues();
    var filtPer = params && params.periodo ? String(params.periodo).trim().slice(0, 7) : '';
    var result  = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var per = _periodoStr(row[3]);
      if (filtPer && per !== filtPer) continue;
      result.push({
        cuenta:    String(row[0] || '').trim(),
        banco:     String(row[1] || '').trim(),
        sociedad:  String(row[2] || '').trim(),
        periodo:   per,
        saldo:     parseFloat(row[4]) || 0,
        bloqueado: row[5] === true || String(row[5]).toUpperCase() === 'SI'
      });
    }
    result.sort(function(a, b) { return (a.sociedad + a.periodo + a.cuenta) < (b.sociedad + b.periodo + b.cuenta) ? -1 : 1; });
    return { status: 'success', data: result, message: '' };
  } catch(e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

/**
 * Guarda o actualiza un saldo inicial. Si BLOQUEADO=SI no sobreescribe.
 * payload: { cuenta, periodo, saldo }  — puede desbloquear con  { cuenta, periodo, desbloquear: true }
 */
function setSaldoInicial(payload) {
  try {
    Session.getActiveUser().getEmail();
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var ss      = SpreadsheetApp.openById(SALDOS_SHEET_ID);
      var siSheet = _ensureSaldoInicialSheet(ss);
      var cuenta  = String(payload.cuenta  || '').trim();
      var periodo = String(payload.periodo || '').trim().slice(0, 7);
      if (!cuenta || !periodo) return { status: 'error', data: {}, message: 'cuenta y periodo requeridos' };

      var data = siSheet.getLastRow() > 1
        ? siSheet.getRange(2, 1, siSheet.getLastRow() - 1, 7).getValues()
        : [];

      var found    = -1;
      var bloqueado = false;
      for (var i = 0; i < data.length; i++) {
        if (String(data[i][0] || '').trim() === cuenta && _periodoStr(data[i][3]) === periodo) {
          found    = i;
          bloqueado = data[i][5] === true || String(data[i][5]).toUpperCase() === 'SI';
          break;
        }
      }

      if (payload.desbloquear) {
        if (found >= 0) { data[found][5] = 'NO'; data[found][6] = new Date(); }
      } else {
        if (bloqueado) return { status: 'error', data: {}, message: 'Periodo bloqueado. Desbloquea antes de editar.' };
        var catMap = _buildCatalogMap(ss);
        var cat    = catMap[cuenta] || {};
        var fila   = [cuenta, cat.banco || '', cat.nombreCorto || cat.nombreSociedad || '', periodo,
                      parseFloat(payload.saldo) || 0, 'NO', new Date()];
        if (found >= 0) { data[found] = fila; }
        else            { data.push(fila); }
      }

      if (siSheet.getLastRow() > 1) siSheet.getRange(2, 1, siSheet.getLastRow() - 1, 7).clearContent();
      if (data.length) siSheet.getRange(2, 1, data.length, 7).setValues(data);
      return { status: 'success', data: {}, message: 'Guardado' };
    } finally {
      lock.releaseLock();
    }
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  }
}

/**
 * Bloquea todos los saldos del periodo indicado (BLOQUEADO=SI).
 * payload: { periodo }
 */
function bloquearPeriodoSaldo(payload) {
  try {
    Session.getActiveUser().getEmail();
    var lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      var ss      = SpreadsheetApp.openById(SALDOS_SHEET_ID);
      var siSheet = _ensureSaldoInicialSheet(ss);
      var periodo = String(payload && payload.periodo || '').trim().slice(0, 7);
      if (!periodo) return { status: 'error', data: {}, message: 'periodo requerido' };
      if (siSheet.getLastRow() <= 1) return { status: 'success', data: { bloqueados: 0 }, message: '0 filas' };

      var data      = siSheet.getRange(2, 1, siSheet.getLastRow() - 1, 7).getValues();
      var ahora     = new Date();
      var bloqueados = 0;
      for (var i = 0; i < data.length; i++) {
        if (_periodoStr(data[i][3]) === periodo) {
          data[i][5] = 'SI';
          data[i][6] = ahora;
          bloqueados++;
        }
      }
      if (data.length) siSheet.getRange(2, 1, data.length, 7).setValues(data);
      return { status: 'success', data: { bloqueados: bloqueados }, message: bloqueados + ' saldos bloqueados' };
    } finally {
      lock.releaseLock();
    }
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  }
}
