/**
 * VEVA Enterprise — Svc_Auditoria.js  v1.0.0
 * Validacion de cadena de saldos entre periodos.
 * Funciones publicas: getCadenaSaldos, validarCadenaSaldos
 *
 * getCadenaSaldos: tabla mes a mes por cuenta — saldo_banco_final[M-1] vs saldo_inicial[M]
 * validarCadenaSaldos: retorna alertas de diferencias entre periodos consecutivos
 */

/**
 * Devuelve la cadena de saldos para auditar continuidad entre meses.
 * params: { desde: 'YYYY-MM', hasta: 'YYYY-MM', sociedad? }
 * Retorna: { filas: [{ cuenta, banco, sociedad, periodos: [{ periodo, saldoFinal, saldoInicial, diferencia, ok }] }] }
 */
function getCadenaSaldos(params) {
  try {
    Session.getActiveUser().getEmail();
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'admin-flujo')) return { status: 'error', data: {}, message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };

    var desde    = params && params.desde    ? String(params.desde).trim().slice(0, 7)    : '';
    var hasta    = params && params.hasta    ? String(params.hasta).trim().slice(0, 7)    : '';
    var filtSoc  = params && params.sociedad ? String(params.sociedad).trim()             : '';
    if (!desde || !hasta) return { status: 'error', data: {}, message: 'Se requiere desde y hasta (YYYY-MM)' };

    // Construir lista de periodos en el rango
    var periodos = _rangoMeses(desde, hasta);
    if (!periodos.length) return { status: 'error', data: {}, message: 'Rango invalido' };

    var sumSheet = ss.getSheetByName('SUM_MOV');
    if (!sumSheet || sumSheet.getLastRow() <= 1) {
      return { status: 'success', data: { filas: [] }, message: 'Sin datos SUM_MOV' };
    }

    // Leer SUM_MOV: necesitamos CUENTA(0), BANCO(1), SOCIEDAD(2), PERIODO(3), SALDO_BANCO_FINAL(19)
    var sumRaw = sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, _SUM_COLS).getValues();

    // Construir mapa: cuentaKey → periodoKey → { saldoBancoFinal }
    // cuentaKey = cuenta + '|' + sociedad
    var sumMap = {};
    for (var i = 0; i < sumRaw.length; i++) {
      var row    = sumRaw[i];
      var cuenta = String(row[0] || '').trim();
      var soc    = String(row[2] || '').trim();
      var banco  = String(row[1] || '').trim();
      var per    = _periodoStr(row[3]);
      if (!cuenta || !soc || !per) continue;
      if (filtSoc && soc !== filtSoc) continue;
      var cKey = cuenta + '|' + soc;
      if (!sumMap[cKey]) sumMap[cKey] = { cuenta: cuenta, banco: banco, sociedad: soc, periodos: {} };
      var sbf = parseFloat(row[19]);
      sumMap[cKey].periodos[per] = { saldoBancoFinal: isNaN(sbf) ? null : sbf };
    }

    // Leer SALDO_INICIAL para primer periodo (fallback)
    var siSheet = ss.getSheetByName('SALDO_INICIAL');
    var siMap = {};
    if (siSheet && siSheet.getLastRow() > 1) {
      var siRaw = siSheet.getRange(2, 1, siSheet.getLastRow() - 1, 7).getValues();
      for (var si = 0; si < siRaw.length; si++) {
        var siCta = String(siRaw[si][0] || '').trim();
        var siSoc = String(siRaw[si][2] || '').trim();
        var siPer = _periodoStr(siRaw[si][3]);
        var siVal = parseFloat(siRaw[si][4]);
        if (siCta && siSoc && siPer && !isNaN(siVal)) {
          siMap[siCta + '|' + siSoc + '|' + siPer] = siVal;
        }
      }
    }

    // Construir filas de cadena
    var filas = [];
    var ctaKeys = Object.keys(sumMap);
    for (var ci = 0; ci < ctaKeys.length; ci++) {
      var cItem = sumMap[ctaKeys[ci]];
      var cadena = [];
      for (var pi = 0; pi < periodos.length; pi++) {
        var per = periodos[pi];
        var perAnterior = _periodoAnterior(per);
        // Saldo final del mes anterior = SALDO_BANCO_FINAL[M-1]
        var datoAnt  = cItem.periodos[perAnterior];
        var saldoFinal = datoAnt ? datoAnt.saldoBancoFinal : null;
        // Saldo inicial del mes actual = SALDO_INICIAL manual (si existe) o SUM_MOV anterior
        var siKey = cItem.cuenta + '|' + cItem.sociedad + '|' + per;
        var saldoInicial = (siMap[siKey] !== undefined) ? siMap[siKey] : saldoFinal;
        // Diferencia
        var diferencia = (saldoFinal !== null && saldoInicial !== null)
          ? Math.round((saldoFinal - saldoInicial) * 100) / 100
          : null;
        var ok = diferencia !== null ? Math.abs(diferencia) <= 1 : null;
        cadena.push({ periodo: per, saldoFinal: saldoFinal, saldoInicial: saldoInicial,
                      diferencia: diferencia, ok: ok });
      }
      var tieneDatos = cadena.some(function(c) { return c.saldoFinal !== null || c.saldoInicial !== null; });
      if (tieneDatos) filas.push({ cuenta: cItem.cuenta, banco: cItem.banco, sociedad: cItem.sociedad, cadena: cadena });
    }

    filas.sort(function(a, b) {
      var ka = a.sociedad + a.cuenta;
      var kb = b.sociedad + b.cuenta;
      return ka < kb ? -1 : ka > kb ? 1 : 0;
    });

    var alertas = 0;
    for (var fi = 0; fi < filas.length; fi++) {
      for (var fci = 0; fci < filas[fi].cadena.length; fci++) {
        if (filas[fi].cadena[fci].ok === false) alertas++;
      }
    }

    return {
      status:  'success',
      data:    { periodos: periodos, filas: filas, alertas: alertas },
      message: filas.length + ' cuentas, ' + alertas + ' alertas de continuidad'
    };
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  }
}

/**
 * Valida si el saldo_banco_final de cada cuenta en `periodo` coincide con
 * el saldo_inicial del periodo siguiente. Para alertas rápidas post-carga.
 * params: { periodo: 'YYYY-MM' }
 */
function validarCadenaSaldos(params) {
  try {
    Session.getActiveUser().getEmail();
    var ss  = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'admin-flujo')) return { status: 'error', data: {}, message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var per = params && params.periodo ? String(params.periodo).trim().slice(0, 7) : '';
    if (!per) return { status: 'error', data: [], message: 'Se requiere periodo YYYY-MM' };

    var perSig = _periodoSiguiente(per);
    var result = getCadenaSaldos({ desde: perSig, hasta: perSig });
    if (result.status !== 'success') return result;

    var alertas = [];
    var filas   = result.data.filas || [];
    for (var fi = 0; fi < filas.length; fi++) {
      var fila = filas[fi];
      for (var ci = 0; ci < fila.cadena.length; ci++) {
        var c = fila.cadena[ci];
        if (c.ok === false) {
          alertas.push({
            cuenta:    fila.cuenta,
            banco:     fila.banco,
            sociedad:  fila.sociedad,
            periodo:   c.periodo,
            saldoFinal:    c.saldoFinal,
            saldoInicial:  c.saldoInicial,
            diferencia:    c.diferencia
          });
        }
      }
    }
    return {
      status:  'success',
      data:    alertas,
      message: alertas.length + ' diferencias de continuidad detectadas'
    };
  } catch(e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

// ─── Helpers locales ─────────────────────────────────────────────────────────

function _rangoMeses(desde, hasta) {
  var result = [];
  var y = parseInt(desde.slice(0, 4));
  var m = parseInt(desde.slice(5, 7));
  var yH = parseInt(hasta.slice(0, 4));
  var mH = parseInt(hasta.slice(5, 7));
  var max = 60; // maximos 5 anios para evitar loops infinitos
  while ((y < yH || (y === yH && m <= mH)) && max-- > 0) {
    result.push(y + '-' + (m < 10 ? '0' + m : String(m)));
    m++;
    if (m > 12) { m = 1; y++; }
  }
  return result;
}

function _periodoSiguiente(periodoYYYYMM) {
  var y = parseInt(periodoYYYYMM.slice(0, 4));
  var m = parseInt(periodoYYYYMM.slice(5, 7)) + 1;
  if (m > 12) { m = 1; y++; }
  return y + '-' + (m < 10 ? '0' + m : String(m));
}
