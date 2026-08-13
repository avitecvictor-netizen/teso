/**
 * VEVA Enterprise — Backend Google Apps Script
 * v0.3.0
 */

var SALDOS_SHEET_ID = '1-DzPpYYViWVG0RHuGGldkDo4LUyYcCLAjQOIlp2DdS8'; // BD_TESO (nueva base) -- antes apuntaba a la BD vieja de VEVA

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    // Corregido 2026-08-12 (pendiente desde el rebranding del 2026-08-10):
    // este setTitle pisaba el <title> ya rebrandeado dentro de
    // Index_1.html -- el titulo de la pestana del navegador seguia
    // diciendo "VEVA" pase lo que pase.
    .setTitle('Tesorería VLMM')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =============================================
// CONTROL DE ACCESO POR VISTA (2026-08-10)
// =============================================
// Bloque 1 de 3 (plan aprobado por el usuario, ver PROJECT_STATUS.md):
// mecanismo central + Pagos. Los Bloques 2 (Lista de Distribucion /
// Notificaciones de Pagos, solo administradores) y 3 (las demas vistas)
// se implementan en rondas separadas, cada una con su propio
// revisor/qa/deploy -- el volumen total (~55-60 funciones) es demasiado
// grande para un solo diff seguro de revisar.
//
// Reusa las funciones de rol/admin YA existentes en Svc_Pagos.js
// (_tieneAlgunRol, _esAdminTesoreria, ROL_TESORERIA, ROL_CONTADOR) --
// Apps Script comparte namespace global entre archivos del mismo
// proyecto, no hace falta mover nada (mismo patron ya usado: este mismo
// archivo define SALDOS_SHEET_ID y Svc_Pagos.js ya lo consume sin
// import explicito).
//
// A diferencia de _tieneAccesoValido (falla ABIERTO si CAT_USUARIOS esta
// vacio, para no bloquear el primer despliegue del modulo Pagos original),
// este control de VISTAS es explicitamente FALL-CLOSED (decision directa
// del usuario, 2026-08-10): si el catalogo esta vacio o el usuario no
// tiene ninguno de los roles esperados, no ve la vista. _tieneAlgunRol ya
// se comporta asi de forma nativa (sin ningun wrapper que "abra" el
// acceso), asi que basta con llamarla directo, sin replicar el criterio
// fail-open de otras partes del modulo.
var VISTA_ROLES_PERMITIDOS = {
  // 'Auditor' (2026-08-13, pedido del usuario): rol ya existente en
  // ROLES_VALIDOS (Svc_Pagos.js) pero sin ningun proposito real hasta
  // ahora ("fantasma", identificado en la auditoria de roles del
  // 2026-08-11) -- se reutiliza para el caso de acceso de solo lectura a
  // Posicion Bancaria (ej. usuarios de otra organizacion, como
  // venturacp.com, que solo necesitan VER la posicion, nunca actuar).
  // getBankingPosition (unica funcion gateada a esta vista) es de solo
  // lectura -- no expone ninguna accion de escritura a este rol nuevo.
  'posicion': ['Tesorería', 'Auditor'],
  'cobranza': ['Tesorería'],
  'reporte-fechas': ['Tesorería'],
  'carga-mov': ['Tesorería'],
  'clasificador': ['Tesorería'],
  'conciliacion': ['Tesorería'],
  'validacion': ['Tesorería'],
  'flujo-efectivo': ['Tesorería'],
  'historico-pagos': ['Tesorería'],
  'admin-flujo': ['Tesorería'],
  'workflow-pagos': ['Tesorería', 'Contador', 'Solicitante'],
  'solicitud': ['Tesorería', 'Solicitante']
};

// Lista-dist (Lista de Distribucion, incluye "Reporte de Posicion" y
// "CC de Pagos") y notificaciones-pagos: NO se gatean por rol -- quedan
// restringidas a la MISMA allowlist de administradores ya construida
// para CONFIG_TESORERIA (_esAdminTesoreria, Svc_Pagos.js), reusada tal
// cual por decision explicita del usuario en vez de crear una lista
// nueva. Esto es una reduccion real de acceso respecto a hoy: Tesoreria
// Y Contador dejan de poder editar CC de Pagos/Notificaciones, salvo que
// ademas esten en ADMINS_TESORERIA.
var VISTAS_SOLO_ADMIN_TESORERIA = ['lista-dist', 'notificaciones-pagos'];

/** Unico punto de verdad para "¿puede este usuario ver esta vista?" --
 * usado tanto para filtrar el menu (getUsuarioActual) como para el gate
 * real server-side de cada funcion de backend que sirve esa vista. Una
 * vista NO mapeada en ninguna de las 2 listas de arriba se deniega
 * (fail-closed tambien para vistas desconocidas -- evita que una vista
 * nueva quede sin control por simple olvido de agregarla al mapa). */
function _tieneAccesoAVista(ss, vistaId) {
  if (VISTAS_SOLO_ADMIN_TESORERIA.indexOf(vistaId) >= 0) return _esAdminTesoreria(ss);
  var roles = VISTA_ROLES_PERMITIDOS[vistaId];
  if (!roles) return false;
  return _tieneAlgunRol(ss, roles);
}

// =============================================
// POSICION BANCARIA
// =============================================

function getBankingPosition(params) {
  try {
    Session.getActiveUser().getEmail();

    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'posicion')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };

    // Fecha seleccionada como Date a medianoche local (sin fill-forward)
    var fechaSeleccionada = (params && params.fecha) ? _validateFecha(params.fecha) : null;

    // --- 1. HIST_POSICION_BANCARIA (fuente unica) ---
    var txSheet = ss.getSheetByName('HIST_POSICION_BANCARIA');
    if (!txSheet || txSheet.getLastRow() <= 1) {
      return { status: 'success', data: { fecha: '', tc_usd: 1, societies: [], _dbg: null }, message: 'Sin datos' };
    }

    var lastCol = txSheet.getLastColumn();
    var lastRow = txSheet.getLastRow();
    var headers = txSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var colIdx  = {};
    headers.forEach(function(h, i) { if (h) colIdx[String(h).trim()] = i; });
    // Columna C puede llamarse SOCIEDAD (legacy) o NOMBRE_CORTO (actual)
    if (colIdx['NOMBRE_CORTO'] !== undefined && colIdx['SOCIEDAD'] === undefined) {
      colIdx['SOCIEDAD'] = colIdx['NOMBRE_CORTO'];
    }

    var txData = txSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // --- 2a. Pasada 1: computar ultimaDisponible y penultimaDisponible ---
    var ultimaDisponible    = null;
    var penultimaDisponible = null; // fecha anterior a ultimaDisponible (para scope 'anterior')
    var fechasValidas       = {};   // set de timestamps con al menos un registro valido

    txData.forEach(function(row) {
      var soc = String(row[colIdx['SOCIEDAD']] || '').trim();
      var cnt = _normalizarCuenta(row[colIdx['CUENTA']]);
      var fd  = _validateFecha(row[colIdx['FECHA_POSICION']]);
      if (!fd || !soc || !cnt) return;
      var yr = fd.getFullYear();
      if (yr < 2000 || yr > 2100) return;
      fechasValidas[fd.getTime()] = fd;
      if (!ultimaDisponible || fd > ultimaDisponible) ultimaDisponible = fd;
    });

    // penultimaDisponible: la fecha inmediatamente anterior a ultimaDisponible
    if (ultimaDisponible) {
      var ultimaTs = ultimaDisponible.getTime();
      Object.keys(fechasValidas).forEach(function(ts) {
        var fd = fechasValidas[ts];
        if (fd.getTime() < ultimaTs) {
          if (!penultimaDisponible || fd > penultimaDisponible) penultimaDisponible = fd;
        }
      });
    }

    // Si no se especifico fecha, usar ultima disponible
    if (!fechaSeleccionada && ultimaDisponible) fechaSeleccionada = ultimaDisponible;

    // --- 2b. Pasada 2: construir saldosMap para la fecha seleccionada ---
    var saldosMap   = {};
    var tcUsd       = 1;
    var tcUsdDate   = -1;

    // Diagnostico (visible en UI cuando societies=0)
    var _dbg = { total: txData.length, sinFecha: 0, fueraRango: 0, fueraLimite: 0, sinSociedad: 0, ok: 0 };
    var _headersEncontrados = Object.keys(colIdx).sort().join(',');

    txData.forEach(function(row) {
      // --- Lectura y normalizacion ---
      var sociedad  = String(row[colIdx['SOCIEDAD']]        || '').trim();
      var numCuenta = _normalizarCuenta(row[colIdx['CUENTA']]);
      var moneda    = String(row[colIdx['MONEDA_ORIGINAL']] || 'MXN').trim().toUpperCase();
      if (moneda === 'MXP') moneda = 'MXN';
      var saldo     = parseFloat(row[colIdx['SALDO_ORIGINAL']])  || 0;
      var banco     = String(row[colIdx['BANCO']]           || '').trim();
      var tc        = parseFloat(row[colIdx['TC_APLICADO']]) || 1;
      var tipo      = String(row[colIdx['TIPO_CUENTA']]     || 'OTRAS').trim();
      var idCarga   = String(row[colIdx['ID_CARGA']]        || '');
      var cargaTs   = parseInt((idCarga.match(/CARGA_(\d+)_/) || [])[1] || '0', 10);
      var fechaDate = _validateFecha(row[colIdx['FECHA_POSICION']]);

      if (!fechaDate) { _dbg.sinFecha++; return; }
      var yr = fechaDate.getFullYear();
      if (yr < 2000 || yr > 2100) { _dbg.fueraRango++; return; }

      if (!fechaSeleccionada) return;
      if (fechaDate.getTime() !== fechaSeleccionada.getTime()) { _dbg.fueraLimite++; return; }

      if (!sociedad || !numCuenta) { _dbg.sinSociedad++; return; }
      _dbg.ok++;

      // FIX: tiebreaker determinista — misma clave mismo dia: gana el cargaTs mas alto
      var key = sociedad + '|' + numCuenta + '|' + moneda;
      if (!saldosMap[key] || cargaTs > saldosMap[key].cargaTs) {
        saldosMap[key] = {
          sociedad: sociedad, numCuenta: numCuenta, moneda: moneda,
          saldo: saldo, banco: banco, tc: tc, tipo: tipo,
          fechaDate: fechaDate, cargaTs: cargaTs
        };
      }

      // TC USD: registro USD con cargaTs mas reciente
      if (moneda === 'USD' && tc > 1 && cargaTs > tcUsdDate) {
        tcUsd = tc; tcUsdDate = cargaTs;
      }
    });

    // --- 3. Agrupar por sociedad ---
    var societiesMap = {};

    Object.keys(saldosMap).forEach(function(key) {
      var item      = saldosMap[key];
      var tipo      = (item.tipo || '').trim().toUpperCase();
      var tipoLabel = tipo === 'JV'                  ? 'JV'
                    : tipo.includes('NO OPERAT')      ? 'No Operativas'
                    : tipo.includes('RESTRI')         ? 'Restringidas'
                    : tipo.includes('OPERAT')         ? 'Operativas'
                    : 'Otras';
      var nombre    = item.sociedad;
      var saldoMxn  = item.saldo * item.tc;

      var socKey = item.sociedad + '||' + tipoLabel;

      if (!societiesMap[socKey]) {
        societiesMap[socKey] = {
          id:        socKey,
          name:      nombre,
          shortName: item.sociedad,
          type:      tipoLabel,
          totalMxn:  0,
          currencies: {}
        };
      }

      var soc = societiesMap[socKey];
      soc.totalMxn += saldoMxn;

      if (!soc.currencies[item.moneda]) {
        soc.currencies[item.moneda] = {
          currencyCode: item.moneda, tc: item.tc, totalMxn: 0, entries: []
        };
      }

      soc.currencies[item.moneda].totalMxn += saldoMxn;
      soc.currencies[item.moneda].entries.push({
        bankName:       item.banco,
        cuenta:         '*' + item.numCuenta.slice(-4),
        originalAmount: item.saldo,
        mxnAmount:      saldoMxn,
        moneda:         item.moneda,
        tc:             item.tc
      });
    });

    var societies = Object.keys(societiesMap).map(function(k) {
      var soc = societiesMap[k];
      return {
        id:        soc.id,
        name:      soc.name,
        shortName: soc.shortName,
        type:      soc.type,
        totalMxn:  soc.totalMxn,
        currencies: Object.keys(soc.currencies).map(function(m) { return soc.currencies[m]; })
      };
    });

    // fechaStr = fecha seleccionada si hay datos; '' si no hay datos o no se selecciono fecha
    var fechaStr = (fechaSeleccionada && Object.keys(saldosMap).length > 0)
      ? Utilities.formatDate(fechaSeleccionada, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : '';
    var ultimaDisponibleStr    = ultimaDisponible
      ? Utilities.formatDate(ultimaDisponible,    Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';
    var penultimaDisponibleStr = penultimaDisponible
      ? Utilities.formatDate(penultimaDisponible, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '';

    // --- 4. Cobranza diaria: dia(s) previo(s) habiles desde HIST_MOVIMIENTOS ---
    var cobranzaPorSociedad = [];
    var fechasCobranza      = [];
    if (fechaSeleccionada) {
      var catMapPos   = _buildCatalogMap(ss);
      var diasInh     = _getDiasInhabiles(ss);
      fechasCobranza  = _prevBusinessDayDates(fechaSeleccionada, diasInh);
      var fechasSetPos = {};
      for (var fpi = 0; fpi < fechasCobranza.length; fpi++) fechasSetPos[fechasCobranza[fpi]] = true;

      var histCobSheet = ss.getSheetByName('HIST_MOVIMIENTOS');
      if (histCobSheet && histCobSheet.getLastRow() > 1) {
        var histCobRaw = histCobSheet.getRange(2, 1, histCobSheet.getLastRow() - 1, _HIST_COLS).getValues();
        var cobMap  = {};
        for (var hci = 0; hci < histCobRaw.length; hci++) {
          var hcr = histCobRaw[hci];
          if (String(hcr[16] || '') !== 'SI') continue;
          var hcCuenta = String(hcr[4] || '').trim();
          var hcCat    = catMapPos[hcCuenta] || {};
          if (!hcCat.tipoCuenta || !hcCat.tipoCuenta.toUpperCase().includes('OPERAT')) continue;
          var hcFRaw = hcr[1];
          var hcFStr = hcFRaw instanceof Date
            ? Utilities.formatDate(hcFRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
            : String(hcFRaw || '').slice(0, 10);
          if (!fechasSetPos[hcFStr]) continue;
          var hcAbono = parseFloat(hcr[7]) || 0;
          if (hcAbono === 0) continue;
          var hcAbr = hcCat.abrCobranza || String(hcCat.nombreCorto || '').substring(0, 3).toUpperCase();
          var hcMon = (hcCat.moneda || 'MXN').toUpperCase();
          if (!cobMap[hcAbr]) cobMap[hcAbr] = { abr: hcAbr, mxn: 0, usd: 0 };
          if (hcMon === 'USD') cobMap[hcAbr].usd += hcAbono;
          else                 cobMap[hcAbr].mxn += hcAbono;
        }
        var cobKeys = Object.keys(cobMap);
        for (var cki = 0; cki < cobKeys.length; cki++) cobranzaPorSociedad.push(cobMap[cobKeys[cki]]);
        cobranzaPorSociedad.sort(function(a, b) { return (b.mxn + b.usd) - (a.mxn + a.usd); });
      }
    }

    return {
      status:  'success',
      data:    {
        fecha: fechaStr, ultimaDisponible: ultimaDisponibleStr,
        penultimaDisponible: penultimaDisponibleStr,
        tc_usd: tcUsd, societies: societies,
        cobranzaPorSociedad: cobranzaPorSociedad,
        fechasCobranza: fechasCobranza,
        _dbg: { headers: _headersEncontrados, conteos: _dbg }
      },
      message: ''
    };

  } catch (e) {
    return { status: 'error', data: {}, message: e.toString() };
  }
}

// =============================================
// HISTORIAL DE POSICION BANCARIA
// =============================================

function _buildCatalogMap(ss) {
  var catSheet = ss.getSheetByName('CAT_CUENTAS_MAPEO');
  if (!catSheet || catSheet.getLastRow() <= 1) return {};
  var lastCol = catSheet.getLastColumn();

  // Leer encabezados para ubicar columnas por nombre (tolerante a reordenamiento)
  var headers = catSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var hi = {};
  headers.forEach(function (h, i) { if (h) hi[String(h).trim().toUpperCase()] = i; });
  function col(name, fallback) { return (hi[name] !== undefined) ? hi[name] : fallback; }

  var iCuenta      = col('NUMERO_CUENTA',  1);  // col B
  var iBanco       = col('BANCO',          2);  // col C
  var iSoc         = col('NOMBRE_SOCIEDAD',4);  // col E
  var iMoneda      = col('MONEDA',         5);  // col F
  var iTipo        = col('TIPO_CUENTA',    6);  // col G
  var iTipo2       = col('TIPO_CUENTA2',   7);  // col H
  var iNombreCorto = col('NOMBRE_CORTO',   8);  // col I
  var iAbr         = col('ABR_COBRANZA',  11);  // col L

  var data = catSheet.getRange(2, 1, catSheet.getLastRow() - 1, lastCol).getValues();
  var map  = {};
  data.forEach(function (row) {
    var numCuenta = _normalizarCuenta(row[iCuenta]);
    if (!numCuenta) return;
    map[numCuenta] = {
      nombreCorto:    String(row[iNombreCorto] || '').trim(),
      nombreSociedad: String(row[iSoc]         || '').trim(),
      banco:          String(row[iBanco]       || '').trim(),
      moneda:         String(row[iMoneda]      || 'MXN').trim(),
      tipoCuenta:     String(row[iTipo]        || '').trim(),
      tipoCuenta2:    String(row[iTipo2]       || '').trim(),
      abrCobranza:    String(row[iAbr]         || '').trim()
    };
  });
  return map;
}

/**
 * Construye un Date en medianoche local a partir de:
 *   - string "YYYY-MM-DD"  (input[type=date] del frontend)
 *   - objeto Date          (normaliza a medianoche, elimina componente hora)
 * Devuelve null si el valor es invalido.
 * Uso: evitar que new Date("YYYY-MM-DD") parsee como UTC midnight y produzca
 *      desfase de un dia en timezone Mexico (UTC-6).
 */
function _parseFechaLocal(s) {
  if (!s) return null;
  if (s instanceof Date) {
    if (isNaN(s.getTime())) return null;
    return new Date(s.getFullYear(), s.getMonth(), s.getDate(), 0, 0, 0, 0);
  }
  var p = String(s).split('-');
  if (p.length !== 3) return null;
  var y = +p[0], mo = +p[1] - 1, d = +p[2];
  if (isNaN(y) || isNaN(mo) || isNaN(d)) return null;
  return new Date(y, mo, d, 0, 0, 0, 0);
}

/**
 * Normaliza numero de cuenta a string canonico sin ceros iniciales.
 * Maneja el caso donde Sheets almacena el numero como Number (pierde cero inicial)
 * y el parser lo devuelve como String con cero inicial — ambos resultan en la misma clave.
 * Ejemplos: "0406605300" -> "406605300", 406605300 -> "406605300"
 */
function _normalizarCuenta(valor) {
  if (valor === null || valor === undefined || valor === '') return '';
  var s = (typeof valor === 'number')
    ? String(Math.round(valor))
    : String(valor).trim().replace(/\D/g, '');
  // Eliminar ceros iniciales: "0406605300" == "406605300" (numero sin cero en Sheets)
  return s.replace(/^0+(\d)/, '$1') || s;
}

/**
 * Parsea y valida una fecha: retorna Date local a medianoche o null si invalida.
 * Rechaza años fuera del rango 2000-2100 (filtra fechas corruptas en HIST).
 */
function _validateFecha(valor) {
  var d = _parseFechaLocal(valor);
  if (!d) return null;
  var y = d.getFullYear();
  if (y < 2000 || y > 2100) return null;
  return d;
}

/**
 * Enriquece un registro crudo con datos del catalogo. Funcion pura — sin I/O.
 * Retorna el objeto enriquecido, o null si la cuenta no tiene mapeo completo.
 * @param {Object} reg     — registro del parser: { cuenta, saldo, moneda, banco, fecha }
 * @param {Object} catMap  — resultado de _buildCatalogMap
 * @param {number} tcDia   — TC del dia a aplicar a moneda != MXN
 */
function _enrichirRegistro(reg, catMap, tcDia) {
  var cuentaNorm = _normalizarCuenta(reg.cuenta);
  var cat = catMap[cuentaNorm];
  if (!cat || (!cat.nombreCorto && !cat.nombreSociedad)) return null;
  var moneda = (cat.moneda || reg.moneda || 'MXN').trim().toUpperCase();
  if (moneda === 'MXP') moneda = 'MXN';
  var tc = (moneda === 'MXN') ? 1 : (tcDia > 0 ? tcDia : 1);
  return {
    sociedad:    cat.nombreCorto || cat.nombreSociedad,
    banco:       cat.banco || reg.banco || '',
    cuenta:      cuentaNorm,
    saldo:       reg.saldo,
    moneda:      moneda,
    tc:          tc,
    saldoMxn:    reg.saldo * tc,
    tipo:        cat.tipoCuenta  || 'OTRAS',
    tipo2:       cat.tipoCuenta2 || '',
    abrCobranza: cat.abrCobranza || ''
  };
}

function _buildTcMap(ss) {
  var tcSheet = ss.getSheetByName('TC');
  if (!tcSheet || tcSheet.getLastRow() <= 1) return {};
  var data = tcSheet.getRange(2, 1, tcSheet.getLastRow() - 1, 2).getValues();
  var m = {};
  for (var i = 0; i < data.length; i++) {
    var d = data[i][0]; var v = parseFloat(data[i][1]);
    if (!d || isNaN(v) || v <= 0) continue;
    var key = d instanceof Date
      ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(d).trim().slice(0, 10);
    if (key.length === 10 && (!m[key] || v > m[key])) m[key] = v;
  }
  return m;
}

function _tcFromMap(tcMap, fechaStr) {
  if (!fechaStr || !Object.keys(tcMap).length) return 1;
  if (tcMap[fechaStr]) return tcMap[fechaStr];
  var keys = Object.keys(tcMap).filter(function(k) { return k <= fechaStr; }).sort();
  return keys.length ? tcMap[keys[keys.length - 1]] : 1;
}

function _getTcForDate(ss, fechaDate) {
  var tcSheet = ss.getSheetByName('TC');
  if (!tcSheet || tcSheet.getLastRow() <= 1) return 1;
  var data       = tcSheet.getRange(2, 1, tcSheet.getLastRow() - 1, 2).getValues();
  var tc         = 1;
  var bestTime   = -1;
  var targetTime = (fechaDate instanceof Date ? fechaDate : new Date(fechaDate)).getTime();
  data.forEach(function (row) {
    var d = row[0] instanceof Date ? row[0] : new Date(row[0]);
    if (isNaN(d.getTime())) return;
    var t   = d.getTime();
    var val = parseFloat(row[1]);
    if (t <= targetTime && t > bestTime && val > 0) {
      tc = val; bestTime = t;
    }
  });
  return tc;
}

function _buildExistingKeys(sheet) {
  var keys = {};
  if (sheet.getLastRow() <= 1) return keys;
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
  data.forEach(function (row) {
    var fechaStr = '';
    try {
      fechaStr = row[1] instanceof Date
        ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(row[1]).substring(0, 10);
    } catch (e) { /* skip */ }
    var cuenta = _normalizarCuenta(row[4]); // normalizar para coincidir con clave de insercion
    if (fechaStr && cuenta) keys[fechaStr + '|' + cuenta] = true;
  });
  return keys;
}

/**
 * Previsualiza archivos sin guardar — devuelve registros parseados y enriquecidos.
 * payload = { archivos: [{ nombre, contenido, banco, fecha, tc, tipo }] }
 */
function previsualizarSaldos(payload) {
  try {
    Session.getActiveUser().getEmail();
    var ss     = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'carga-mov')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var catMap = _buildCatalogMap(ss);
    var preview = [];

    payload.archivos.forEach(function (archivo) {
      var result = archivo.tipo === 'movimientos'
        ? ParserBancario.procesarArchivoMovimientos(archivo.contenido, archivo.nombre, archivo.banco)
        : ParserBancario.procesarArchivoSaldos(archivo.contenido, archivo.nombre, archivo.banco, archivo.fecha, archivo.tc);

      if (!result.success) {
        preview.push({ nombre: archivo.nombre, ok: false, error: result.error, registros: [] });
        return;
      }

      var fechaDate = _validateFecha(archivo.fecha);
      var tcDia     = fechaDate ? _getTcForDate(ss, fechaDate) : 1;

      var registros = result.datos.map(function (reg) {
        var enriquecido = _enrichirRegistro(reg, catMap, tcDia);
        if (enriquecido) {
          return {
            cuenta:      enriquecido.cuenta,
            banco:       enriquecido.banco,
            sociedad:    enriquecido.sociedad,
            moneda:      enriquecido.moneda,
            tipo:        enriquecido.tipo,
            tipo2:       enriquecido.tipo2,
            abrCobranza: enriquecido.abrCobranza,
            saldo:       enriquecido.saldo,
            tc:          enriquecido.tc,
            saldoMxn:    enriquecido.saldoMxn
          };
        }
        // Sin mapeo — mostrar en preview con flag para que el usuario lo vea antes de guardar
        return {
          cuenta:      _normalizarCuenta(reg.cuenta) || reg.cuenta,
          banco:       reg.banco || '',
          sociedad:    '(sin mapeo en catalogo)',
          moneda:      reg.moneda || 'MXN',
          tipo:        '',
          tipo2:       '',
          abrCobranza: '',
          saldo:       reg.saldo,
          tc:          1,
          saldoMxn:    reg.saldo
        };
      });

      preview.push({
        nombre:       archivo.nombre,
        ok:           true,
        registros:    registros,
        advertencias: result.advertencias || []
      });
    });

    return { status: 'success', data: preview, message: '' };
  } catch (e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

/**
 * Parsea y guarda en HIST_POSICION_BANCARIA — con deduplicacion y escritura batch.
 * payload = { archivos: [{ nombre, contenido, banco, fecha, tc, tipo }] }
 */
function guardarSaldosEnHistorial(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    Session.getActiveUser().getEmail();

    var ss        = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'carga-mov')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var histSheet = ss.getSheetByName('HIST_POSICION_BANCARIA');
    if (!histSheet) {
      return { status: 'error', data: {}, message: 'Hoja HIST_POSICION_BANCARIA no encontrada' };
    }

    var catMap    = _buildCatalogMap(ss);
    var existKeys = _buildExistingKeys(histSheet);
    var allRows   = [];
    var errores   = [];
    var sinMapeo  = [];
    var duplicados = 0;

    payload.archivos.forEach(function (archivo) {
      var result = archivo.tipo === 'movimientos'
        ? ParserBancario.procesarArchivoMovimientos(archivo.contenido, archivo.nombre, archivo.banco)
        : ParserBancario.procesarArchivoSaldos(archivo.contenido, archivo.nombre, archivo.banco, archivo.fecha, archivo.tc);

      if (!result.success) {
        errores.push({ archivo: archivo.nombre, error: result.error });
        return;
      }

      var archFechaDate = _validateFecha(archivo.fecha);
      var tcDia         = archFechaDate ? _getTcForDate(ss, archFechaDate) : 1;

      result.datos.forEach(function (reg) {
        // Fecha: prioridad reg.fecha (parser), fallback archivo.fecha
        var fechaDate = _validateFecha(reg.fecha) || _validateFecha(archivo.fecha);
        if (!fechaDate) {
          errores.push({ archivo: archivo.nombre, cuenta: reg.cuenta, error: 'Fecha no determinada' });
          return;
        }
        var fechaStr = Utilities.formatDate(fechaDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

        // Enriquecer con catalogo — rechazar cuentas sin mapeo (no escribir filas huerfanas)
        var enriquecido = _enrichirRegistro(reg, catMap, tcDia);
        if (!enriquecido) {
          sinMapeo.push({ archivo: archivo.nombre, cuenta: _normalizarCuenta(reg.cuenta) || reg.cuenta });
          return;
        }

        // Dedup por fecha|cuenta — en modo reemplazar solo dedup dentro del batch actual
        var key = fechaStr + '|' + enriquecido.cuenta;
        if (payload.modo !== 'reemplazar' && existKeys[key]) { duplicados++; return; }
        existKeys[key] = true;

        allRows.push([
          'CARGA_' + Date.now() + '_' + Math.floor(Math.random() * 9999),
          fechaDate,
          enriquecido.sociedad,
          enriquecido.banco,
          enriquecido.cuenta,
          enriquecido.saldo,
          enriquecido.moneda,
          enriquecido.tc,
          enriquecido.saldoMxn,
          enriquecido.tipo,
          archivo.nombre,
          enriquecido.tipo2,
          enriquecido.abrCobranza
        ]);
      });
    });

    // Modo reemplazar: eliminar filas existentes en HIST para las mismas clave fecha|cuenta
    var reemplazados = 0;
    if (payload.modo === 'reemplazar' && allRows.length > 0) {
      // Construir set de claves que seran reemplazadas
      var replaceKeys = {};
      allRows.forEach(function (r) {
        var fs = r[1] instanceof Date
          ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'yyyy-MM-dd')
          : String(r[1]).substring(0, 10);
        replaceKeys[fs + '|' + r[4]] = true; // r[4] = cuenta ya normalizada
      });

      // Escanear HIST y marcar filas a eliminar
      if (histSheet.getLastRow() > 1) {
        var histData = histSheet.getRange(2, 1, histSheet.getLastRow() - 1, 5).getValues();
        var toDelete = [];
        histData.forEach(function (row, i) {
          var fs = row[1] instanceof Date
            ? Utilities.formatDate(row[1], Session.getScriptTimeZone(), 'yyyy-MM-dd')
            : String(row[1]).substring(0, 10);
          var cn = _normalizarCuenta(row[4]);
          if (replaceKeys[fs + '|' + cn]) toDelete.push(i + 2); // +2: 1-indexed + header
        });
        // Eliminar en orden inverso para que los indices no se desplacen
        toDelete.reverse().forEach(function (rowIdx) {
          histSheet.deleteRow(rowIdx);
        });
        reemplazados = toDelete.length;
      }
    }

    // Escritura batch — una sola llamada a la API de Sheets
    if (allRows.length > 0) {
      var lastRow = histSheet.getLastRow();
      histSheet.getRange(lastRow + 1, 1, allRows.length, 13).setValues(allRows);
    }

    return {
      status:  'success',
      data:    { guardados: allRows.length, reemplazados: reemplazados, duplicados: duplicados, sinMapeo: sinMapeo, errores: errores },
      message: allRows.length + ' registros guardados en HIST_POSICION_BANCARIA'
        + (reemplazados   > 0 ? ' (' + reemplazados   + ' reemplazados)'           : '')
        + (sinMapeo.length > 0 ? ' (' + sinMapeo.length + ' sin mapeo ignoradas)'  : '')
    };

  } catch (e) {
    return { status: 'error', data: {}, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// =============================================
// HISTORIAL DE POSICION — CONSULTA
// =============================================

function getHistorialPosicion(params) {
  try {
    Session.getActiveUser().getEmail();
    var ss    = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'reporte-fechas')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var sheet = ss.getSheetByName('HIST_POSICION_BANCARIA');
    if (!sheet || sheet.getLastRow() <= 1) {
      return { status: 'success', data: [], message: 'Sin datos en historial' };
    }

    var desde = (params && params.desde) ? _parseFechaLocal(params.desde) : null;
    var hasta = (params && params.hasta) ? _parseFechaLocal(params.hasta) : null;
    if (hasta) hasta.setHours(23, 59, 59, 999);

    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
    var rows = [];

    data.forEach(function (row) {
      var fecha = _parseFechaLocal(row[1]); // normaliza Date de Sheets a medianoche local
      if (!fecha || isNaN(fecha.getTime())) return;
      if (desde && fecha < desde) return;
      if (hasta && fecha > hasta)  return;
      rows.push({
        fecha:       Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        sociedad:    String(row[2]  || ''),
        banco:       String(row[3]  || ''),
        cuenta:      String(row[4]  || ''),
        saldo:       parseFloat(row[5]) || 0,
        moneda:      String(row[6]  || 'MXN'),
        tc:          parseFloat(row[7]) || 1,
        saldoMxn:    parseFloat(row[8]) || 0,
        tipo:        String(row[9]  || ''),
        tipo2:       String(row[11] || ''),
        abrCobranza: String(row[12] || '')
      });
    });

    return { status: 'success', data: rows, message: '' };
  } catch (e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

// La funcion temporal diagnosticoCatCuentasSoccer (solo lectura, listaba
// filas de CAT_CUENTAS_MAPEO conteniendo "soccer" para verificar el caso
// real "Soccer Loco Holding" antes de escribir el fallback por nombre en
// _resolverSociedadDesdeNombreEmpresa, Svc_Pagos.js) vivio aqui durante
// esa investigacion y se retiro 2026-08-02 una vez revisado su log.

// Las funciones temporales diagnosticoSaldosHoy (diagnostico de solo
// lectura) y fixSaldoBanorte406605300 (correccion puntual del saldo mal
// cargado el 2026-07-30 para la cuenta 406605300/Banorte, causado por el
// bug de raiz ya corregido en Svc_Parser.js -- ver _validarCuentasUnicas)
// vivieron aqui durante esa ronda y se retiraron 2026-08-01 una vez
// confirmado por el usuario que la correccion se aplico correctamente.

// =============================================
// LISTA DE DISTRIBUCION
// =============================================

function _getOrCreateDistSheet(ss) {
  var sheet = ss.getSheetByName('CONFIG_DISTRIBUCION');
  if (!sheet) {
    sheet = ss.insertSheet('CONFIG_DISTRIBUCION');
    sheet.getRange(1, 1, 1, 4).setValues([['Email', 'Nombre', 'Activo', 'FechaAlta']]);
    sheet.getRange(2, 1, 3, 4).setValues([
      ['vaviles+veva@venturae.com.mx', 'V. Aviles',        true, new Date()],
      ['lcampos@venturae.com.mx',      'L. Campos',        true, new Date()],
      ['lcampos+veva@venturae.com.mx', 'L. Campos (VEVA)', true, new Date()]
    ]);
  }
  return sheet;
}

function getListaDistribucion() {
  try {
    Session.getActiveUser().getEmail();
    var ss      = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    // Control de acceso por vista, Bloque 2 (2026-08-10): igual que
    // enviarReportePosicion, esta funcion sirve Lista de Distribucion --
    // restringida a administradores de CONFIG_TESORERIA.
    if (!_tieneAccesoAVista(ss, 'lista-dist')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Solo un administrador de Tesorería puede verlo.' };
    var sheet   = _getOrCreateDistSheet(ss);
    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { status: 'success', data: [], message: '' };

    var data  = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    var lista = data
      .filter(function (row) { return String(row[0]).trim(); })
      .map(function (row) {
        return {
          email:    String(row[0]).trim(),
          nombre:   String(row[1]).trim(),
          activo:   row[2] === true || String(row[2]).toUpperCase() === 'TRUE',
          fechaAlta: (row[3] instanceof Date)
            ? Utilities.formatDate(row[3], Session.getScriptTimeZone(), 'dd/MM/yyyy')
            : String(row[3]).substring(0, 10)
        };
      });

    return { status: 'success', data: lista, message: '' };
  } catch (e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

function guardarListaDistribucion(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    Session.getActiveUser().getEmail();
    var ss    = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'lista-dist')) return { status: 'error', data: {}, message: 'No tienes acceso a este módulo. Solo un administrador de Tesorería puede verlo.' };
    var sheet = _getOrCreateDistSheet(ss);

    var lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 4).clearContent();

    var lista = payload.lista || [];
    if (lista.length === 0) return { status: 'success', data: {}, message: 'Lista guardada (vacia)' };

    var rows = lista.map(function (item) {
      return [String(item.email || '').trim(), String(item.nombre || '').trim(), item.activo !== false, new Date()];
    });

    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
    return { status: 'success', data: { guardados: rows.length }, message: rows.length + ' destinatarios guardados' };
  } catch (e) {
    return { status: 'error', data: {}, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// =============================================
// ENVIO DE REPORTE POR CORREO
// =============================================

function enviarReportePosicion(payload) {
  try {
    Session.getActiveUser().getEmail();
    // Control de acceso por vista, Bloque 2 (2026-08-10): esta funcion
    // sirve la tab "Reporte de Posicion" de Lista de Distribucion --
    // restringida a los administradores de CONFIG_TESORERIA, no por rol.
    // Antes NO tenia ningun control de acceso real (el Session.getActiveUser
    // de arriba solo forzaba la autenticacion, nunca se usaba para
    // autorizar). ss se abre aqui (antes se abria mas abajo, solo para
    // leer la config del remitente) para poder gatear desde el principio.
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'lista-dist')) return { status: 'error', data: {}, message: 'No tienes acceso a este módulo. Solo un administrador de Tesorería puede verlo.' };

    var scope         = payload.scope         || 'actual';
    var destinatarios = payload.destinatarios || [];
    if (destinatarios.length === 0) {
      return { status: 'error', data: {}, message: 'Sin destinatarios activos' };
    }

    // Una sola call sin fecha — getBankingPosition usa ultimaDisponible internamente
    var posicion = getBankingPosition();
    if (posicion.status !== 'success') {
      return { status: 'error', data: {}, message: 'Error al obtener posicion: ' + posicion.message };
    }
    if (!posicion.data.fecha) {
      return { status: 'error', data: {}, message: 'Sin datos disponibles en historial' };
    }

    var d        = posicion.data;
    var fmtMxn   = function (n) { return '$' + Math.round(n).toLocaleString(); };
    var totalMxn = d.societies.reduce(function (a, s) { return a + s.totalMxn; }, 0);

    // Agrupar por tipo en orden canonico
    var TYPE_ORDER = ['Operativas', 'JV', 'No Operativas', 'Restringidas', 'Otras'];
    var tipoMap = {};
    d.societies.forEach(function (soc) {
      var t = soc.type || 'Otras';
      if (!tipoMap[t]) tipoMap[t] = { total: 0, societies: [] };
      tipoMap[t].societies.push(soc);
      tipoMap[t].total += soc.totalMxn;
    });

    var filas = '';
    TYPE_ORDER.filter(function (t) { return tipoMap[t]; }).forEach(function (t) {
      var g = tipoMap[t];
      filas += '<tr style="background:#f0f2f8;">'
        + '<td style="padding:6px 16px;font-size:11px;font-weight:700;color:#0a2540;text-transform:uppercase;letter-spacing:0.05em;">' + t + '</td>'
        + '<td style="padding:6px 16px;font-size:11px;font-weight:700;text-align:right;color:#0a2540;">' + fmtMxn(g.total) + '</td>'
        + '</tr>';
      g.societies.forEach(function (soc) {
        filas += '<tr style="border-bottom:1px solid #f6f9fc;">'
          + '<td style="padding:4px 16px 4px 28px;font-size:12px;color:#0a2540;">' + soc.name + '</td>'
          + '<td style="padding:4px 16px;font-size:12px;font-weight:600;text-align:right;color:#0a2540;">' + fmtMxn(soc.totalMxn) + '</td>'
          + '</tr>';
      });
    });
    filas += '<tr style="background:#f6f9fc;border-top:2px solid #e6ebf1;">'
      + '<td style="padding:8px 16px;font-size:11px;font-weight:700;color:#697386;text-transform:uppercase;">Total Consolidado</td>'
      + '<td style="padding:8px 16px;font-size:14px;font-weight:700;text-align:right;color:#635bff;">' + fmtMxn(totalMxn) + '</td>'
      + '</tr>';

    var scopeLabel = scope === 'actual' ? 'Actual'
      : scope === 'anterior' ? 'Actual + Anterior'
      : 'Ultimos ' + scope.replace('dias', '') + ' dias';

    var html = '<div style="font-family:Arial,sans-serif;max-width:640px;margin:0 auto;">'
      + '<div style="background:#0a2540;padding:24px 32px;border-radius:12px 12px 0 0;">'
      + '<span style="background:#635bff;color:#fff;font-weight:900;font-size:20px;padding:6px 14px;border-radius:8px;">V</span>'
      + '<span style="color:#fff;font-weight:700;font-size:18px;margin-left:12px;">Ventura Entertainment</span>'
      + '</div>'
      + '<div style="background:#fff;padding:24px 32px;border:1px solid #e6ebf1;border-top:none;">'
      + '<p style="font-size:14px;color:#697386;margin:0 0 4px;">Posicion Bancaria Consolidada &mdash; ' + scopeLabel + '</p>'
      + '<p style="font-size:24px;font-weight:700;color:#635bff;margin:0 0 4px;">' + fmtMxn(totalMxn) + ' MXN</p>'
      + '<p style="font-size:12px;color:#697386;margin:0 0 24px;">Datos al ' + d.fecha + ' &mdash; TC USD: ' + d.tc_usd + '</p>'
      + '<table style="width:100%;border-collapse:collapse;border:1px solid #e6ebf1;">'
      + '<thead><tr style="background:#f6f9fc;">'
      + '<th style="padding:10px 16px;text-align:left;font-size:11px;color:#697386;font-weight:700;text-transform:uppercase;">Clasificacion / Sociedad</th>'
      + '<th style="padding:10px 16px;text-align:right;font-size:11px;color:#697386;font-weight:700;text-transform:uppercase;">Saldo MXN</th>'
      + '</tr></thead>'
      + '<tbody>' + filas + '</tbody>'
      + '</table>'
      + '</div>'
      + '<div style="background:#f6f9fc;padding:12px 32px;border:1px solid #e6ebf1;border-top:none;border-radius:0 0 12px 12px;">'
      + '<p style="font-size:11px;color:#697386;margin:0;">Generado por Ventura Entertainment &mdash; '
      + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm') + '</p>'
      + '</div>'
      + '</div>';

    // Remitente 2026-08-10: ya no fijo -- viene de CONFIG_TESORERIA
    // (Svc_Pagos.js, editable desde Lista de Distribucion por los
    // administradores listados ahi). GmailApp SI permite fijar "from",
    // pero solo si ya es un alias "Enviar correo como" verificado de
    // quien ejecuta el envio (executeAs:USER_ACCESSING) -- responsabilidad
    // de quien configure el valor, el codigo no puede validarlo de
    // antemano.
    var configTesoreria = _leerConfigTesoreria(ss);
    GmailApp.sendEmail(destinatarios.join(','), 'Tesorer\u00eda VLMM | Posicion Bancaria ' + scopeLabel + ' \u2014 ' + d.fecha, '', {
      htmlBody: html,
      from: configTesoreria.emailTesoreria,
      name: 'Tesorer\u00eda VLMM'
    });

    return { status: 'success', data: { enviados: destinatarios.length }, message: 'Reporte enviado a ' + destinatarios.length + ' destinatario(s)' };
  } catch (e) {
    return { status: 'error', data: {}, message: e.toString() };
  }
}

// =============================================
// DIAGNOSTICO
// =============================================

/**
 * Valida que todas las cuentas en HIST_POSICION_BANCARIA tengan mapeo completo en CAT_CUENTAS_MAPEO.
 * Ejecutar desde el editor de GAS y revisar el log.
 */
function validarMapeo() {
  var ss      = SpreadsheetApp.openById(SALDOS_SHEET_ID);
  var catMap  = _buildCatalogMap(ss);
  var hist    = ss.getSheetByName('HIST_POSICION_BANCARIA');
  if (!hist || hist.getLastRow() <= 1) { Logger.log('HIST vacio'); return; }

  var data    = hist.getRange(2, 1, hist.getLastRow() - 1, 5).getValues();
  var vistas  = {};

  data.forEach(function (row) {
    var cuenta = String(row[4] || '').trim();
    if (!cuenta || vistas[cuenta]) return;
    vistas[cuenta] = true;

    var cat = catMap[cuenta.replace(/\D/g, '')];
    if (!cat) {
      Logger.log('SIN MAPEO    cuenta=' + cuenta);
    } else {
      var campos = [];
      if (!cat.nombreSociedad) campos.push('sociedad');
      if (!cat.banco)          campos.push('banco');
      if (!cat.moneda)         campos.push('moneda');
      if (!cat.tipoCuenta)     campos.push('tipo');
      if (campos.length > 0) {
        Logger.log('INCOMPLETO   cuenta=' + cuenta + '  faltan=' + campos.join(','));
      } else {
        Logger.log('OK           cuenta=' + cuenta + '  soc=' + cat.nombreSociedad + '  banco=' + cat.banco);
      }
    }
  });
  Logger.log('Validacion terminada. Cuentas unicas revisadas: ' + Object.keys(vistas).length);
}

/**
 * Diagnostico profundo: muestra HIST vs CAT para cuentas especificas.
 * Agrega o quita sufijos en el array SUFIJOS segun se necesite.
 */
function diagnosticoCuenta() {
  var SUFIJOS = ['3628', '7073', '2292'];
  var ss      = SpreadsheetApp.openById(SALDOS_SHEET_ID);
  var catMap  = _buildCatalogMap(ss);
  var hist    = ss.getSheetByName('HIST_POSICION_BANCARIA');
  if (!hist || hist.getLastRow() <= 1) { Logger.log('HIST vacio'); return; }

  var data = hist.getRange(2, 1, hist.getLastRow() - 1, 11).getValues();
  var vistas = {};

  data.forEach(function (row) {
    var cuenta = String(row[4] || '').trim();
    var match  = SUFIJOS.some(function (s) { return cuenta.endsWith(s); });
    if (!match || vistas[cuenta]) return;
    vistas[cuenta] = true;

    Logger.log('--- cuenta=' + cuenta);
    Logger.log('    HIST  sociedad="' + row[2] + '"  banco="' + row[3] + '"  tipo_hist="' + row[9] + '"');
    var cat = catMap[cuenta.replace(/\D/g, '')];
    if (cat) {
      Logger.log('    CAT   soc="' + cat.nombreSociedad + '"  banco="' + cat.banco
        + '"  moneda="' + cat.moneda + '"  tipo="' + cat.tipoCuenta
        + '"  tipo2="' + cat.tipoCuenta2 + '"');
    } else {
      Logger.log('    CAT   SIN MAPEO');
    }
  });
  Logger.log('Fin diagnostico. Cuentas revisadas: ' + Object.keys(vistas).length);
}

function testHeaders() {
  var ss      = SpreadsheetApp.openById(SALDOS_SHEET_ID);
  var sheet   = ss.getSheets()[0];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  Logger.log('Hoja: ' + sheet.getName());
  Logger.log('Columnas: ' + JSON.stringify(headers));
  var sample = sheet.getLastRow() > 1
    ? sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0]
    : [];
  Logger.log('Fila 2: ' + JSON.stringify(sample));
}

// =============================================
// MOVIMIENTOS BANCARIOS — CATALOGOS
// =============================================

function _buildCatCuentasArray(ss) {
  var catSheet = ss.getSheetByName('CAT_CUENTAS_MAPEO');
  if (!catSheet || catSheet.getLastRow() <= 1) return [];
  var lastCol = catSheet.getLastColumn();
  var headers = catSheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var hi = {};
  headers.forEach(function(h, i) { if (h) hi[String(h).trim().toUpperCase()] = i; });
  function col(name, fallback) { return hi[name] !== undefined ? hi[name] : fallback; }

  var iCuenta      = col('NUMERO_CUENTA', 1);
  var iBanco       = col('BANCO', 2);
  var iIdSoc       = col('ID_SOCIEDAD', 3);
  var iNombreSoc   = col('NOMBRE_SOCIEDAD', 4);
  var iMoneda      = col('MONEDA', 5);
  var iTipo        = col('TIPO_CUENTA', 6);
  var iTipo2       = col('TIPO_CUENTA2', 7);
  var iNombreCorto = col('NOMBRE_CORTO', 8);
  var iAbr         = col('ABR_COBRANZA', 11);

  var data = catSheet.getRange(2, 1, catSheet.getLastRow() - 1, lastCol).getValues();
  var arr = [];
  data.forEach(function(row) {
    var num = _normalizarCuenta(row[iCuenta]);
    if (!num) return;
    arr.push({
      cuenta:        num,
      banco:         String(row[iBanco]       || '').trim(),
      idSociedad:    String(row[iIdSoc]       || '').trim(),
      nombreSociedad:String(row[iNombreSoc]   || '').trim(),
      moneda:        String(row[iMoneda]      || 'MXN').trim(),
      tipoCuenta:    String(row[iTipo]        || '').trim(),
      tipoCuenta2:   String(row[iTipo2]       || '').trim(),
      nombreCorto:   String(row[iNombreCorto] || '').trim(),
      abrCobranza:   String(row[iAbr]         || '').trim()
    });
  });
  return arr;
}

function _buildReglasArray(ss) {
  var sheet = ss.getSheetByName('REGLAS_PARSER');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var hi = {};
  headers.forEach(function(h, i) {
    if (!h) return;
    var k = String(h).trim().toUpperCase()
      .replace(/[ÁÀÂÄ]/g, 'A').replace(/[ÉÈÊË]/g, 'E')
      .replace(/[ÍÌÎÏ]/g, 'I').replace(/[ÓÒÔÖ]/g, 'O')
      .replace(/[ÚÙÛÜ]/g, 'U').replace(/Ñ/g, 'N');
    hi[k] = i;
  });
  if (hi['MATCH'] !== undefined && hi['TIPO_MATCH'] === undefined) hi['TIPO_MATCH'] = hi['MATCH'];
  function col(name, fallback) { return hi[name] !== undefined ? hi[name] : fallback; }

  var iFase   = col('FASE', 1);
  var iId     = col('REGLA_ID', 2);
  var iTipo   = col('TIPO', 3);
  var iClasif = col('CLASIFICACION', 4);
  var iSub    = col('SUBCATEGORIA', 5);
  var iPatron = col('PATRON', 6);
  var iMatch  = col('TIPO_MATCH', 7);
  var iCampo  = col('CAMPO_ESTANDAR', 8);
  var iLookup = col('CAMPO_LOOKUP', 9);
  var iAccion = col('ACCION', 10);

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  var arr = [];
  data.forEach(function(row) {
    var id = String(row[iId] || '').trim();
    if (!id) return;
    arr.push({
      fase:          String(row[iFase]   || '').trim().toUpperCase(),
      reglaId:       id,
      tipo:          String(row[iTipo]   || '').trim().toUpperCase(),
      clasificacion: String(row[iClasif] || '').trim(),
      subcategoria:  String(row[iSub]    || '').trim(),
      patron:        String(row[iPatron] || '').trim(),
      tipoMatch:     String(row[iMatch]  || '').trim().toUpperCase(),
      campoEstandar: String(row[iCampo]  || '').trim().toUpperCase(),
      campoLookup:   String(row[iLookup] || '').trim().toUpperCase(),
      accion:        String(row[iAccion] || '').trim().toUpperCase()
    });
  });
  return arr;
}

function _buildCatClientesArray(ss) {
  var sheet = ss.getSheetByName('CAT_CLIENTES');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var hi = {};
  headers.forEach(function(h, i) { if (h) hi[String(h).trim().toUpperCase()] = i; });
  function col(name, fallback) { return hi[name] !== undefined ? hi[name] : fallback; }

  var iPatron    = col('PATRON', 0);
  var iTipoMatch = col('TIPO_MATCH', 1);
  var iNombre    = col('NOMBRE', 2);
  var iActivo    = col('ACTIVO', 3);

  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  var arr = [];
  data.forEach(function(row) {
    var patron = String(row[iPatron] || '').trim();
    if (!patron) return;
    var activo = row[iActivo];
    if (activo === false || String(activo).toUpperCase() === 'FALSE' || String(activo).toUpperCase() === 'NO') return;
    arr.push({
      patron:    patron,
      tipoMatch: String(row[iTipoMatch] || 'CONTAINS').trim().toUpperCase(),
      nombre:    String(row[iNombre]    || patron).trim()
    });
  });
  return arr;
}

// HIST_MOVIMIENTOS — 22 columnas (A-V)
// A=ID_MOV B=FECHA C=SOCIEDAD D=BANCO E=CUENTA F=DESCRIPCION
// G=CARGO H=ABONO I=SALDO J=REFERENCIA K=CLAVE_RASTREO L=CONTRAPARTE
// M=TIPO N=CLASIFICACION O=SUBCATEGORIA P=REGLA_ID Q=ES_COBRANZA R=ARCHIVO_FUENTE
// S=CLASIFICACION_MANUAL T=ID_CONTRA U=CONCEPTO V=CLIENTE_COBRANZA
var _HIST_COLS = 24;

// SUM_EFE — 11 columnas (A-K): agregados por cuenta+periodo+tipo+clasif2+cliente
// A=CUENTA B=BANCO C=SOCIEDAD D=PERIODO E=TIPO(INGRESO/EGRESO)
// F=CLASIFICACION G=CLASIF2 H=CLIENTE_COBRANZA I=MONTO J=CONTEO K=FECHA_ACTUALIZACION
var _SUM_EFE_COLS = 11;

// Convierte el valor de una celda PERIODO (puede ser string, Date o serial numérico) a 'YYYY-MM'.
function _periodoStr(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM');
  return String(val).trim().slice(0, 7);
}

// SUM_MOV — 21 columnas (A-U)
// A=CUENTA B=BANCO C=SOCIEDAD D=PERIODO E=TIPO_CUENTA
// F=SALDO_INICIAL G=TOTAL_ABONOS H=TOTAL_CARGOS
// I=ABONOS_INTERNOS J=CARGOS_INTERNOS
// K=ABONOS_NETOS L=CARGOS_NETOS M=SALDO_FINAL_CALC
// N=TOTAL_COBRANZA O=CONTEO P=FECHA_ACTUALIZACION Q=SALDO_REPORTADO
// R=TOTAL_COBRANZA_MXN S=TOTAL_COBRANZA_USD
// T=SALDO_BANCO_FINAL (ultimo SALDO_BANCO del periodo — para saldo inicial siguiente mes)
// U=ES_PARTIDA (SI cuando clasif es TRASPASOS/INTERCIAS/INVERSION)
var _SUM_COLS = 21;

function _ensureHistMovSheet(ss) {
  var headers = [
    'ID_MOV', 'FECHA', 'SOCIEDAD', 'BANCO', 'CUENTA', 'DESCRIPCION',
    'CARGO', 'ABONO', 'SALDO_BANCO', 'REFERENCIA', 'CLAVE_RASTREO', 'CONTRAPARTE',
    'TIPO', 'CLASIFICACION', 'SUBCATEGORIA', 'REGLA_ID', 'ES_COBRANZA', 'ARCHIVO_FUENTE',
    'CLASIFICACION_MANUAL', 'ID_CONTRA', 'CONCEPTO', 'CLIENTE_COBRANZA',
    'MONEDA_ORIGINAL', 'TC_APLICADO'
  ];
  var sheet = ss.getSheetByName('HIST_MOVIMIENTOS');
  if (!sheet) {
    sheet = ss.insertSheet('HIST_MOVIMIENTOS');
    sheet.getRange(1, 1, 1, _HIST_COLS).setValues([headers]);
    sheet.setFrozenRows(1);
  } else {
    var h9 = sheet.getRange(1, 9).getValue();
    if (h9 === 'SALDO') sheet.getRange(1, 9).setValue('SALDO_BANCO');
    var lastCol = sheet.getLastColumn();
    if (lastCol < 23) sheet.getRange(1, 23).setValue('MONEDA_ORIGINAL');
    if (lastCol < 24) sheet.getRange(1, 24).setValue('TC_APLICADO');
  }
  return sheet;
}

function _ensureSumEfeSheet(ss) {
  var sheet = ss.getSheetByName('SUM_EFE');
  if (!sheet) {
    sheet = ss.insertSheet('SUM_EFE');
    sheet.getRange(1, 1, 1, _SUM_EFE_COLS).setValues([[
      'CUENTA', 'BANCO', 'SOCIEDAD', 'PERIODO', 'TIPO',
      'CLASIFICACION', 'CLASIF2', 'CLIENTE_COBRANZA', 'MONTO', 'CONTEO', 'FECHA_ACTUALIZACION'
    ]]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  // Migrar schema 10→11: insertar CLIENTE_COBRANZA en col H (indice 7), desplazar MONTO/CONTEO/FECHA
  if (sheet.getLastColumn() === 10) {
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      var old = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
      var migrated = [];
      for (var mi = 0; mi < old.length; mi++) {
        var r = old[mi];
        migrated.push([r[0], r[1], r[2], r[3], r[4], r[5], r[6], '', r[7], r[8], r[9]]);
      }
      sheet.getRange(2, 1, lastRow - 1, 10).clearContent();
      sheet.getRange(2, 1, migrated.length, _SUM_EFE_COLS).setValues(migrated);
    }
    sheet.getRange(1, 1, 1, _SUM_EFE_COLS).setValues([[
      'CUENTA', 'BANCO', 'SOCIEDAD', 'PERIODO', 'TIPO',
      'CLASIFICACION', 'CLASIF2', 'CLIENTE_COBRANZA', 'MONTO', 'CONTEO', 'FECHA_ACTUALIZACION'
    ]]);
  }
  return sheet;
}

function _actualizarSumEfe(ss, histData, afectados, catMap) {
  var mapeoFlujo = _leerMapeoFlujo(ss);
  var efeMap = {};

  for (var i = 0; i < histData.length; i++) {
    var row      = histData[i];
    var cuenta   = String(row[4] || '').trim();
    var fechaRaw = row[1];
    var fechaStr = fechaRaw instanceof Date
      ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(fechaRaw).slice(0, 10);
    var periodo  = fechaStr.slice(0, 7);
    if (!afectados[cuenta + '|' + periodo]) continue;

    var tipo = String(row[12] || '').trim();
    if (tipo !== 'INGRESO' && tipo !== 'EGRESO') continue;

    var clasif  = String(row[13] || '').trim();
    var sub     = String(row[14] || '').trim();
    var cargo   = parseFloat(row[6])  || 0;
    var abono   = parseFloat(row[7])  || 0;
    var tc      = parseFloat(row[23]) || 1;
    var monto   = (tipo === 'INGRESO' ? abono : cargo) * tc;
    if (monto <= 0) continue;

    var resolved = _resolverClasif(clasif, sub, tipo, mapeoFlujo);
    var sec      = resolved.clasif;
    var c2       = resolved.clasif2;
    var cat      = catMap[cuenta] || {};
    var soc      = String(row[2]  || '').trim();
    var banco    = String(row[3]  || '').trim() || cat.banco || '';
    var cliente  = String(row[21] || '').trim();
    var eKey     = cuenta + '|' + periodo + '|' + tipo + '|' + c2 + '|' + cliente;

    if (!efeMap[eKey]) {
      efeMap[eKey] = { cuenta: cuenta, banco: banco, sociedad: soc,
        periodo: periodo, tipo: tipo, clasif: sec, clasif2: c2, cliente: cliente, monto: 0, conteo: 0 };
    }
    efeMap[eKey].monto  += monto;
    efeMap[eKey].conteo += 1;
  }

  if (Object.keys(efeMap).length === 0) return;

  var ahora      = new Date();
  var nuevasFilas = [];
  var eKeys      = Object.keys(efeMap);
  for (var ki = 0; ki < eKeys.length; ki++) {
    var e = efeMap[eKeys[ki]];
    nuevasFilas.push([
      e.cuenta, e.banco, e.sociedad, e.periodo, e.tipo,
      e.clasif, e.clasif2, e.cliente, e.monto, e.conteo, ahora
    ]);
  }

  var efeSheet     = _ensureSumEfeSheet(ss);
  var filasConservar = [];
  if (efeSheet.getLastRow() > 1) {
    var existData = efeSheet.getRange(2, 1, efeSheet.getLastRow() - 1, _SUM_EFE_COLS).getValues();
    for (var ei = 0; ei < existData.length; ei++) {
      var eCta = String(existData[ei][0] || '').trim();
      var ePer = _periodoStr(existData[ei][3]);
      if (!afectados[eCta + '|' + ePer]) filasConservar.push(existData[ei]);
    }
  }

  var todasLasFilas = filasConservar.concat(nuevasFilas);
  var clearRows     = efeSheet.getLastRow() - 1;
  if (clearRows > 0) efeSheet.getRange(2, 1, clearRows, _SUM_EFE_COLS).clearContent();
  if (todasLasFilas.length > 0) {
    efeSheet.getRange(2, 1, todasLasFilas.length, _SUM_EFE_COLS).setValues(todasLasFilas);
    efeSheet.getRange(2, 4, todasLasFilas.length, 1).setNumberFormat('@');
  }
}

function _ensureSumMovSheet(ss) {
  var sheet = ss.getSheetByName('SUM_MOV');
  if (sheet) {
    // Migrar a 21 cols si existe con schema anterior (19 cols)
    var lastCol = sheet.getLastColumn();
    if (lastCol < 21 && sheet.getLastRow() >= 1) {
      if (lastCol < 20) sheet.getRange(1, 20).setValue('SALDO_BANCO_FINAL');
      if (lastCol < 21) sheet.getRange(1, 21).setValue('ES_PARTIDA');
    }
    return sheet;
  }
  sheet = ss.insertSheet('SUM_MOV');
  sheet.getRange(1, 1, 1, _SUM_COLS).setValues([[
    'CUENTA', 'BANCO', 'SOCIEDAD', 'PERIODO', 'TIPO_CUENTA',
    'SALDO_INICIAL', 'TOTAL_ABONOS', 'TOTAL_CARGOS',
    'ABONOS_INTERNOS', 'CARGOS_INTERNOS',
    'ABONOS_NETOS', 'CARGOS_NETOS', 'SALDO_FINAL_CALC',
    'TOTAL_COBRANZA', 'CONTEO', 'FECHA_ACTUALIZACION', 'SALDO_REPORTADO',
    'TOTAL_COBRANZA_MXN', 'TOTAL_COBRANZA_USD',
    'SALDO_BANCO_FINAL', 'ES_PARTIDA'
  ]]);
  sheet.setFrozenRows(1);
  return sheet;
}

// =============================================
// DIAS INHABILES — HELPERS
// =============================================

function _ensureDiasInhabilesSheet(ss) {
  var sheet = ss.getSheetByName('CAT_DIAS_INHABILES');
  if (sheet) return sheet;
  sheet = ss.insertSheet('CAT_DIAS_INHABILES');
  sheet.getRange(1, 1, 1, 3).setValues([['FECHA', 'DESCRIPCION', 'ACTIVO']]);
  var h = [
    ['2025-01-01','Año Nuevo',true], ['2025-02-03','Constitucion',true],
    ['2025-03-17','Juarez',true],    ['2025-04-17','Jueves Santo',true],
    ['2025-04-18','Viernes Santo',true], ['2025-05-01','Trabajo',true],
    ['2025-09-16','Independencia',true], ['2025-11-17','Revolucion',true],
    ['2025-12-25','Navidad',true],
    ['2026-01-01','Año Nuevo',true], ['2026-02-02','Constitucion',true],
    ['2026-03-16','Juarez',true],    ['2026-04-02','Jueves Santo',true],
    ['2026-04-03','Viernes Santo',true], ['2026-05-01','Trabajo',true],
    ['2026-09-16','Independencia',true], ['2026-11-16','Revolucion',true],
    ['2026-12-25','Navidad',true]
  ];
  sheet.getRange(2, 1, h.length, 3).setValues(h);
  sheet.setFrozenRows(1);
  return sheet;
}

function _getDiasInhabiles(ss) {
  _ensureDiasInhabilesSheet(ss);
  var sheet = ss.getSheetByName('CAT_DIAS_INHABILES');
  var data  = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  var set   = {};
  for (var i = 0; i < data.length; i++) {
    var activo = data[i][2];
    if (activo === false || String(activo).toLowerCase() === 'false') continue;
    var fRaw = data[i][0];
    var fStr = fRaw instanceof Date
      ? Utilities.formatDate(fRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(fRaw || '').trim().slice(0, 10);
    if (fStr.length === 10) set[fStr] = true;
  }
  return set;
}

// Devuelve array de fechas YYYY-MM-DD para reportar cobranza dado el dia de la posicion.
// Regla: si D-1 es dia habil → [D-1].
// Si D-1 es fin de semana o inhabil → recorre hacia atras todos los dias no habiles
// consecutivos y los devuelve (sin incluir el dia habil anterior a esa racha).
function _prevBusinessDayDates(fechaDate, diasInhabiles) {
  var tz    = Session.getScriptTimeZone();
  var d     = new Date(fechaDate.getTime());
  d.setDate(d.getDate() - 1);
  var dates = [];
  while (true) {
    var ds     = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    var dow    = d.getDay();
    var nonBiz = (dow === 0 || dow === 6 || !!diasInhabiles[ds]);
    if (nonBiz) {
      dates.push(ds);
      d.setDate(d.getDate() - 1);
    } else {
      if (dates.length === 0) dates.push(ds); // dia habil directo
      break;
    }
  }
  return dates;
}

// TC USD mas reciente desde HIST_POSICION_BANCARIA
function _getLatestTcUsd(ss) {
  var sheet = ss.getSheetByName('HIST_POSICION_BANCARIA');
  if (!sheet || sheet.getLastRow() <= 1) return 1;
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var hi = {};
  headers.forEach(function(h, i) { if (h) hi[String(h).trim()] = i; });
  var data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  var bestTc = 1, bestTs = -1;
  var iMon = hi['MONEDA_ORIGINAL'] !== undefined ? hi['MONEDA_ORIGINAL'] : 5;
  var iTc  = hi['TC_APLICADO']     !== undefined ? hi['TC_APLICADO']     : 7;
  var iId  = hi['ID_CARGA']        !== undefined ? hi['ID_CARGA']        : -1;
  for (var i = 0; i < data.length; i++) {
    var mon = String(data[i][iMon] || '').trim().toUpperCase();
    var tc  = parseFloat(data[i][iTc]) || 0;
    if (mon !== 'USD' || tc <= 1) continue;
    var ts = (iId >= 0)
      ? parseInt((String(data[i][iId] || '').match(/CARGA_(\d+)_/) || [])[1] || '-1', 10)
      : -1;
    if (ts > bestTs) { bestTc = tc; bestTs = ts; }
  }
  return bestTc;
}

function _ensureCatClientesSheet(ss) {
  var sheet = ss.getSheetByName('CAT_CLIENTES');
  if (sheet) return sheet;
  sheet = ss.insertSheet('CAT_CLIENTES');
  sheet.getRange(1, 1, 1, 5).setValues([['PATRON', 'TIPO_MATCH', 'NOMBRE', 'ACTIVO', 'NOTA']]);
  sheet.getRange(2, 1, 10, 5).setValues([
    ['CARNIVAL',    'CONTAINS',    'Carnival Corp.',      true, ''],
    ['MSC',         'CONTAINS',    'MSC Cruises',         true, ''],
    ['ROYAL CARIE', 'CONTAINS',    'Royal Caribbean',     true, ''],
    ['PRINCESS CR', 'CONTAINS',    'Princess Cruises',    true, ''],
    ['NCL CORPOR',  'CONTAINS',    'Norwegian Cruise',    true, ''],
    ['CLASSICA C',  'CONTAINS',    'Classica Cruises',    true, ''],
    ['DISNEY WOR',  'CONTAINS',    'Disney World',        true, ''],
    ['AFIL-',       'STARTS_WITH', 'Afiliacion',          true, 'Prefijo afiliaciones'],
    ['F',           'STARTS_WITH', 'Cliente F (factura)', true, 'Clientes sin referencia mas que factura'],
    ['ABONO FACT',  'STARTS_WITH', 'Abono Factura',       true, '']
  ]);
  sheet.setFrozenRows(1);
  return sheet;
}

function _ensureSaldoInicialSheet(ss) {
  var sheet = ss.getSheetByName('SALDO_INICIAL');
  if (sheet) return sheet;
  sheet = ss.insertSheet('SALDO_INICIAL');
  sheet.getRange(1, 1, 1, 7).setValues([[
    'CUENTA', 'BANCO', 'SOCIEDAD', 'PERIODO', 'SALDO', 'BLOQUEADO', 'FECHA_ACTUALIZACION'
  ]]);
  sheet.setFrozenRows(1);
  return sheet;
}

function _ensureMapeoFlujoSheet(ss) {
  var sheet = ss.getSheetByName('MAPEO_FLUJO');
  if (sheet) return sheet;
  sheet = ss.insertSheet('MAPEO_FLUJO');
  sheet.getRange(1, 1, 1, 5).setValues([['TIPO', 'CLASIFICACION', 'CLASIFICACION2', 'ORDEN', 'ACTIVO']]);
  var data = [
    ['INGRESOS', 'OPERATIVO',       'Cliente',                       10, true],
    ['INGRESOS', 'OPERATIVO',       'Cobranza',                      20, true],
    ['INGRESOS', 'OPERATIVO',       'Patrocinios',                   30, true],
    ['INGRESOS', 'INTERCIAS',       'Aportaciones capital Intercia', 10, true],
    ['INGRESOS', 'INTERCIAS',       'Prestamos Intercias',           20, true],
    ['INGRESOS', 'INTERCIAS',       'Intereses partes relacionadas', 30, true],
    ['INGRESOS', 'INTERCIAS',       'Venta de Activo Fijo Intercias',40, true],
    ['INGRESOS', 'NO OPERATIVO',    'Aportaciones capital',          10, true],
    ['INGRESOS', 'NO OPERATIVO',    'Traspasos entre cuentas',       20, true],
    ['INGRESOS', 'NO OPERATIVO',    'Arrendadoras',                  30, true],
    ['INGRESOS', 'NO OPERATIVO',    'Venta de Activo Fijo',          40, true],
    ['INGRESOS', 'NO OPERATIVO',    'Intercambios',                  50, true],
    ['INGRESOS', 'NO OPERATIVO',    'Devoluciones pagos',            60, true],
    ['INGRESOS', 'FINANCIEROS',     'Prestamos bancarios',           10, true],
    ['INGRESOS', 'FINANCIEROS',     'Inversiones',                   20, true],
    ['INGRESOS', 'FINANCIEROS',     'Intereses',                     30, true],
    ['INGRESOS', 'FINANCIEROS',     'Factoraje',                     40, true],
    ['EGRESOS',  'OPERATIVO',       'Proveedores',                   10, true],
    ['EGRESOS',  'OPERATIVO',       'Varios',                        20, true],
    ['EGRESOS',  'OPERATIVO',       'Comisiones Bancarias',          30, true],
    ['EGRESOS',  'OPERATIVO',       'Nominas',                       40, true],
    ['EGRESOS',  'OPERATIVO',       'Regalias',                      50, true],
    ['EGRESOS',  'OPERATIVO',       'Arrendamiento Inmuebles',       60, true],
    ['EGRESOS',  'OPERATIVO',       'Arrendamientos',                70, true],
    ['EGRESOS',  'NO OPERATIVO',    'Traspasos entre cuentas',       10, true],
    ['EGRESOS',  'NO OPERATIVO',    'Intercambios',                  20, true],
    ['EGRESOS',  'CAPEX',           'Capex',                         10, true],
    ['EGRESOS',  'FINANCIEROS',     'Capital Deuda',                 10, true],
    ['EGRESOS',  'FINANCIEROS',     'Intereses deuda',               20, true],
    ['EGRESOS',  'FINANCIEROS',     'Inversiones',                   30, true],
    ['EGRESOS',  'FINANCIEROS',     'Intereses',                     40, true],
    ['EGRESOS',  'FINANCIEROS',     'Factoraje',                     50, true],
    ['EGRESOS',  'FINANCIEROS',     'TDC',                           60, true],
    ['EGRESOS',  'INTERCIAS',       'Inversiones en subsidiaria',    10, true],
    ['EGRESOS',  'INTERCIAS',       'Prestamos Intercias',           20, true],
    ['EGRESOS',  'INTERCIAS',       'Intereses partes relacionadas', 30, true],
    ['EGRESOS',  'INTERCIAS',       'Venta de Activo Fijo',          40, true],
    ['EGRESOS',  'IMPUESTOS FEDE',  'IVA',                           10, true],
    ['EGRESOS',  'IMPUESTOS FEDE',  'ISR',                           20, true],
    ['EGRESOS',  'IMPUESTOS FEDE',  'Retenciones',                   30, true],
    ['EGRESOS',  'IMPUESTOS LOCA',  'IMSS SAR and INFONAVIT',        10, true],
    ['EGRESOS',  'IMPUESTOS LOCA',  'ISN',                           20, true]
  ];
  sheet.getRange(2, 1, data.length, 5).setValues(data);
  sheet.setFrozenRows(1);
  return sheet;
}

// =============================================
// MOVIMIENTOS BANCARIOS — HELPERS INTERNOS
// =============================================

/**
 * Matching de transferencias internas en memoria.
 * Modifica histData en el lugar; devuelve el conteo de pares encontrados.
 * Solo procesa periodos incluidos en el set 'afectados' (cuenta|YYYY-MM).
 * Algoritmo: hash por monto centavos → busca pares EGRESO↔INGRESO misma fecha ±1 dia,
 * distinta cuenta, ambas cuentas internas (en cuentasInternas).
 */
function _matchearTransferenciasEnMemoria(histData, afectados, cuentasInternas, catCuentasArray) {
  // indices para cada monto: { [montoKey]: { [fechaStr]: [{rowIdx, cuenta, idMov}] } }
  var mapIngresos = {};
  var mapEgresos  = {};

  for (var i = 0; i < histData.length; i++) {
    var row    = histData[i];
    var cuenta = String(row[4] || '').trim();
    if (!cuentasInternas[cuenta]) continue;

    var fechaRaw = row[1];
    var fechaStr = fechaRaw instanceof Date
      ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(fechaRaw).slice(0, 10);
    var periodo  = fechaStr.slice(0, 7);
    if (!afectados[cuenta + '|' + periodo]) continue;

    var clasif = String(row[13] || '');
    if (clasif === 'TRASPASOS_INTERNOS' || clasif === 'INTERCIAS' || clasif === 'TRANSFERENCIA_INTERNA') continue; // ya matcheado

    var cargo = parseFloat(row[6]) || 0;
    var abono = parseFloat(row[7]) || 0;
    var idMov = String(row[0] || '');
    if (!idMov) continue;

    var entry = { rowIdx: i, cuenta: cuenta, idMov: idMov };

    if (abono > 0 && cargo === 0) {
      var keyI = Math.round(abono * 100);
      if (!mapIngresos[keyI]) mapIngresos[keyI] = {};
      if (!mapIngresos[keyI][fechaStr]) mapIngresos[keyI][fechaStr] = [];
      mapIngresos[keyI][fechaStr].push(entry);
    } else if (cargo > 0 && abono === 0) {
      var keyE = Math.round(cargo * 100);
      if (!mapEgresos[keyE]) mapEgresos[keyE] = {};
      if (!mapEgresos[keyE][fechaStr]) mapEgresos[keyE][fechaStr] = [];
      mapEgresos[keyE][fechaStr].push(entry);
    }
  }

  // Construye las 3 fechas a verificar (misma fecha + ±1 dia) para tolerancia
  function _fechasPlusMinusOne(fechaStr) {
    var d    = new Date(fechaStr + 'T12:00:00');
    var dm1  = new Date(d.getTime() - 86400000);
    var dp1  = new Date(d.getTime() + 86400000);
    function fmt(x) {
      return x.getFullYear() + '-' +
        ('0' + (x.getMonth() + 1)).slice(-2) + '-' +
        ('0' + x.getDate()).slice(-2);
    }
    return [fechaStr, fmt(dm1), fmt(dp1)];
  }

  var usados      = {};
  var totalPares  = 0;
  var egresoKeys  = Object.keys(mapEgresos);

  for (var ek = 0; ek < egresoKeys.length; ek++) {
    var montoKey  = +egresoKeys[ek];
    var fechasEg  = Object.keys(mapEgresos[montoKey]);

    for (var ef = 0; ef < fechasEg.length; ef++) {
      var fechaEg = fechasEg[ef];
      var listaEg = mapEgresos[montoKey][fechaEg];

      for (var ei = 0; ei < listaEg.length; ei++) {
        var eg = listaEg[ei];
        if (usados[eg.rowIdx]) continue;

        // Buscar ingreso coincidente: mismo monto (clave exacta), fecha ±1 dia
        var fechasCheck = _fechasPlusMinusOne(fechaEg);
        var encontrado  = null;

        for (var fc = 0; fc < fechasCheck.length && !encontrado; fc++) {
          var lista = mapIngresos[montoKey] && mapIngresos[montoKey][fechasCheck[fc]];
          if (!lista) continue;
          for (var li = 0; li < lista.length; li++) {
            var ing = lista[li];
            if (usados[ing.rowIdx]) continue;
            if (ing.cuenta === eg.cuenta) continue; // misma cuenta = no es transferencia
            encontrado = ing;
            break;
          }
        }

        if (encontrado) {
          usados[eg.rowIdx]          = true;
          usados[encontrado.rowIdx]  = true;
          // Clasificar segun sociedad: misma sociedad=TRASPASOS_INTERNOS, distinta=INTERCIAS
          var socEg  = String(histData[eg.rowIdx][2] || '').trim();
          var socIng = String(histData[encontrado.rowIdx][2] || '').trim();
          var clasMatch = (socEg && socIng && socEg === socIng) ? 'TRASPASOS_INTERNOS' : 'INTERCIAS';
          // Columnas: N(13)=CLASIFICACION, S(18)=CLASIFICACION_MANUAL, T(19)=ID_CONTRA
          histData[eg.rowIdx][13]           = clasMatch;
          histData[eg.rowIdx][18]           = 'AUTO';
          histData[eg.rowIdx][19]           = encontrado.idMov;
          histData[encontrado.rowIdx][13]   = clasMatch;
          histData[encontrado.rowIdx][18]   = 'AUTO';
          histData[encontrado.rowIdx][19]   = eg.idMov;
          totalPares++;
        }
      }
    }
  }

  // Segunda pasada: clasificar por nombre de contraparte cuando coincide con sociedad del grupo
  if (catCuentasArray && catCuentasArray.length) {
    var relacionMap = {};
    for (var ri = 0; ri < catCuentasArray.length; ri++) {
      var rc = catCuentasArray[ri];
      var rnombres = [rc.nombreCorto, rc.nombreSociedad];
      for (var rni = 0; rni < rnombres.length; rni++) {
        var rnom = (rnombres[rni] || '').trim().toUpperCase();
        if (rnom.length >= 4) relacionMap[rnom] = rc.idSociedad || rc.nombreSociedad || '';
      }
    }
    var rKeys = Object.keys(relacionMap);
    for (var i2 = 0; i2 < histData.length; i2++) {
      if (usados[i2]) continue;
      var row2 = histData[i2];
      var cta2 = String(row2[4] || '').trim();
      if (!cuentasInternas[cta2]) continue;
      var fr2  = row2[1];
      var fs2  = fr2 instanceof Date ? Utilities.formatDate(fr2, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(fr2).slice(0, 10);
      var per2 = fs2.slice(0, 7);
      if (!afectados[cta2 + '|' + per2]) continue;
      var cl2 = String(row2[13] || '');
      if (cl2 === 'TRASPASOS_INTERNOS' || cl2 === 'INTERCIAS' || cl2 === 'TRANSFERENCIA_INTERNA') continue;
      if (String(row2[18] || '') === 'MANUAL') continue;
      var contrap2 = String(row2[11] || '').trim().toUpperCase();
      if (!contrap2) continue;
      var socFound = null;
      for (var rk = 0; rk < rKeys.length && !socFound; rk++) {
        if (contrap2.indexOf(rKeys[rk]) >= 0) socFound = relacionMap[rKeys[rk]];
      }
      if (socFound !== null) {
        var soc2      = String(row2[2] || '').trim();
        var clasRel   = (soc2 && socFound && soc2 === socFound) ? 'TRASPASOS_INTERNOS' : 'INTERCIAS';
        histData[i2][13] = clasRel;
        histData[i2][18] = 'AUTO_RELAC';
        totalPares++;
      }
    }
  }

  return totalPares;
}

/**
 * Recalcula SUM_MOV solo para los periodos afectados (cuenta|YYYY-MM).
 * Lee histData completo (ya actualizado), agrega en memoria y reescribe SUM_MOV.
 */
function _actualizarSumMov(ss, histData, afectados, catMap) {
  var sumMap = {};

  for (var i = 0; i < histData.length; i++) {
    var row    = histData[i];
    var cuenta = String(row[4] || '').trim();

    var fechaRaw = row[1];
    var fechaStr = fechaRaw instanceof Date
      ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
      : String(fechaRaw).slice(0, 10);
    var periodo  = fechaStr.slice(0, 7);
    var key      = cuenta + '|' + periodo;
    if (!afectados[key]) continue;

    var cargo      = parseFloat(row[6])  || 0;
    var abono      = parseFloat(row[7])  || 0;
    var saldo      = parseFloat(row[8]);
    var clasif     = String(row[13] || '');
    var esInterna  = (clasif === 'TRASPASOS_INTERNOS' || clasif === 'INTERCIAS' || clasif === 'TRANSFERENCIA_INTERNA');
    var esPartida  = (clasif === 'TRASPASOS_INTERNOS' || clasif === 'INTERCIAS' ||
                      clasif === 'TRANSFERENCIA_INTERNA' || clasif === 'IN_INVERSION' ||
                      clasif === 'EG_INVERSION' || clasif === 'INVERSIONES' || clasif === 'PRESTAMOS');
    var esCobranza = (String(row[16] || '') === 'SI');

    if (!sumMap[key]) {
      var cat       = catMap[cuenta] || {};
      sumMap[key] = {
        cuenta:    cuenta,
        banco:     String(row[3] || '').trim() || cat.banco || '',
        sociedad:  String(row[2] || '').trim(),
        periodo:   periodo,
        tipoCuenta: cat.tipoCuenta || '',
        totalAbonos:    0, totalCargos:    0,
        abonosInternos: 0, cargosInternos: 0,
        totalCobranza:  0, totalCobranzaMxn: 0, totalCobranzaUsd: 0, conteo: 0,
        tienePartida:   false,
        _firstDate: null, _firstSaldo: null, _firstCargo: null, _firstAbono: null,
        _lastDate:  null, _lastSaldo:  null, _lastTc: 1
      };
    }

    var s = sumMap[key];
    s.totalAbonos += abono;
    s.totalCargos += cargo;
    if (esInterna)  { s.abonosInternos += abono; s.cargosInternos += cargo; }
    if (esPartida)  { s.tienePartida = true; }
    if (esCobranza) {
      var monedaCta = (catMap[cuenta] || {}).moneda || 'MXN';
      s.totalCobranza += abono;
      if (monedaCta === 'USD') { s.totalCobranzaUsd += abono; }
      else                     { s.totalCobranzaMxn += abono; }
    }
    s.conteo++;

    if (!s._firstDate || fechaStr < s._firstDate) {
      s._firstDate  = fechaStr;
      s._firstSaldo = isNaN(saldo) ? null : saldo;
      s._firstCargo = cargo;
      s._firstAbono = abono;
    }
    if (!s._lastDate || fechaStr >= s._lastDate) {
      s._lastDate  = fechaStr;
      s._lastSaldo = isNaN(saldo) ? null : saldo;
      s._lastTc    = parseFloat(row[23]) || 1;
    }
  }

  if (Object.keys(sumMap).length === 0) return;

  var ahora = new Date();
  var nuevasFilas = [];

  var keys = Object.keys(sumMap);
  for (var ki = 0; ki < keys.length; ki++) {
    var s = sumMap[keys[ki]];
    // saldo_inicial = saldo de primera fila + cargo - abono (saldo antes de ese movimiento)
    var saldoInicial   = (s._firstSaldo !== null && s._firstSaldo !== 0)
      ? (s._firstSaldo + s._firstCargo - s._firstAbono) : null;
    var abonosNetos    = s.totalAbonos   - s.abonosInternos;
    var cargosNetos    = s.totalCargos   - s.cargosInternos;
    var saldoFinalCalc = saldoInicial !== null ? (saldoInicial + abonosNetos - cargosNetos) : '';

    nuevasFilas.push([
      s.cuenta, s.banco, s.sociedad, s.periodo, s.tipoCuenta,
      saldoInicial !== null ? saldoInicial : '',
      s.totalAbonos, s.totalCargos,
      s.abonosInternos, s.cargosInternos,
      abonosNetos, cargosNetos, saldoFinalCalc,
      s.totalCobranza, s.conteo, ahora,
      s._lastSaldo !== null ? s._lastSaldo : '',
      s.totalCobranzaMxn, s.totalCobranzaUsd,
      s._lastSaldo !== null ? s._lastSaldo * (s._lastTc || 1) : '',  // T: SALDO_BANCO_FINAL en MXN
      s.tienePartida ? 'SI' : 'NO'                                   // U: ES_PARTIDA
    ]);
  }

  var sumSheet = _ensureSumMovSheet(ss);

  // Leer filas existentes y excluir las que seran reemplazadas
  var filasConservar = [];
  if (sumSheet.getLastRow() > 1) {
    var existData = sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, _SUM_COLS).getValues();
    for (var ei = 0; ei < existData.length; ei++) {
      var ekCuenta     = String(existData[ei][0] || '').trim();
      var ekPeriodoRaw = existData[ei][3];
      var ekPeriodo    = (ekPeriodoRaw instanceof Date)
        ? Utilities.formatDate(ekPeriodoRaw, Session.getScriptTimeZone(), 'yyyy-MM')
        : String(ekPeriodoRaw || '').trim().slice(0, 7);
      if (!afectados[ekCuenta + '|' + ekPeriodo]) {
        var fila = existData[ei].slice();
        if (ekPeriodoRaw instanceof Date) fila[3] = ekPeriodo; // normalizar Date → string
        filasConservar.push(fila);
      }
    }
  }

  var todasLasFilas = filasConservar.concat(nuevasFilas);

  // Pre-generar filas vacias para todos los periodos 2026 de cuentas del catalogo
  var existKeys = {};
  for (var ti = 0; ti < todasLasFilas.length; ti++) {
    existKeys[String(todasLasFilas[ti][0]||'').trim() + '|' + String(todasLasFilas[ti][3]||'').trim()] = true;
  }
  var meses2026 = ['2026-01','2026-02','2026-03','2026-04','2026-05','2026-06',
                   '2026-07','2026-08','2026-09','2026-10','2026-11','2026-12'];
  for (var vCta in catMap) {
    var vCat = catMap[vCta];
    for (var vmi = 0; vmi < meses2026.length; vmi++) {
      var vKey = vCta + '|' + meses2026[vmi];
      if (!existKeys[vKey]) {
        todasLasFilas.push([
          vCta, vCat.banco || '', vCat.nombreCorto || vCat.nombreSociedad || '',
          meses2026[vmi], vCat.tipoCuenta || '',
          '', 0, 0, 0, 0, 0, 0, '', 0, 0, ahora, '', 0, 0, '', 'NO'
        ]);
        existKeys[vKey] = true;
      }
    }
  }

  // Limpiar hoja de datos y reescribir en un solo batch
  var clearRows = sumSheet.getLastRow() - 1;
  if (clearRows > 0) sumSheet.getRange(2, 1, clearRows, _SUM_COLS).clearContent();
  if (todasLasFilas.length > 0) {
    sumSheet.getRange(2, 1, todasLasFilas.length, _SUM_COLS).setValues(todasLasFilas);
    sumSheet.getRange(2, 4, todasLasFilas.length, 1).setNumberFormat('@');
  }
}

/**
 * Registra saldo inicial por cuenta|periodo en hoja SALDO_INICIAL.
 * Solo escribe entradas no bloqueadas; entradas BLOQUEADO=SI nunca se sobreescriben.
 * Debe llamarse DESPUES de _actualizarSumMov (lee SUM_MOV para obtener el saldo calculado).
 */
function _registrarSaldoInicial(ss, afectados, catMap) {
  if (!afectados || !Object.keys(afectados).length) return;
  var sumSheet = ss.getSheetByName('SUM_MOV');
  var siSheet  = _ensureSaldoInicialSheet(ss);
  if (!sumSheet || sumSheet.getLastRow() <= 1) return;

  var sumData = sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, _SUM_COLS).getValues();

  // Mapa de entradas existentes en SALDO_INICIAL
  var siData  = siSheet.getLastRow() > 1
    ? siSheet.getRange(2, 1, siSheet.getLastRow() - 1, 7).getValues()
    : [];
  var siMap = {};
  for (var si = 0; si < siData.length; si++) {
    var scta = String(siData[si][0] || '').trim();
    var sper = _periodoStr(siData[si][3]);
    if (scta && sper) siMap[scta + '|' + sper] = (siData[si][5] === true || String(siData[si][5]).toUpperCase() === 'SI');
  }

  var ahora  = new Date();
  var nuevas = [];
  for (var ri = 0; ri < sumData.length; ri++) {
    var row    = sumData[ri];
    var cuenta = String(row[0] || '').trim();
    var periodo = _periodoStr(row[3]);
    if (!cuenta || !periodo) continue;
    if (!afectados[cuenta + '|' + periodo]) continue;
    if (siMap[cuenta + '|' + periodo] === true) continue; // bloqueado — no tocar
    var saldo = row[5]; // SUM_MOV col F = SALDO_INICIAL
    if (saldo === '' || saldo === null || saldo === undefined) continue;
    var cat = catMap ? (catMap[cuenta] || {}) : {};
    nuevas.push([cuenta, String(row[1] || cat.banco || ''), String(row[2] || ''), periodo, saldo, 'NO', ahora]);
  }
  if (!nuevas.length) return;

  // Conservar entradas no afectadas o bloqueadas, reemplazar el resto
  var conservar = [];
  for (var ei = 0; ei < siData.length; ei++) {
    var ecta = String(siData[ei][0] || '').trim();
    var eper = _periodoStr(siData[ei][3]);
    var ekey = ecta + '|' + eper;
    var bloq = (siData[ei][5] === true || String(siData[ei][5]).toUpperCase() === 'SI');
    if (!afectados[ekey] || bloq) conservar.push(siData[ei]);
  }

  var todas = conservar.concat(nuevas);
  if (siSheet.getLastRow() > 1) siSheet.getRange(2, 1, siSheet.getLastRow() - 1, 7).clearContent();
  if (todas.length) siSheet.getRange(2, 1, todas.length, 7).setValues(todas);
}

// =============================================
// MOVIMIENTOS BANCARIOS — API
// =============================================

function previsualizarMovimientos(payload) {
  try {
    Session.getActiveUser().getEmail();
    var ss          = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'carga-mov')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var catCuentas  = _buildCatCuentasArray(ss);
    var reglas      = _buildReglasArray(ss);
    var catClientes = _buildCatClientesArray(ss);
    var resultados  = [];

    payload.archivos.forEach(function(archivo) {
      var sheets = archivo.sheets || [];
      var banco  = archivo.banco  || '';
      var parsed = ParserMovimientos.procesarArchivo(banco, sheets, catCuentas, reglas, catClientes);
      resultados.push({
        nombre:         archivo.nombre,
        success:        parsed.success,
        error:          parsed.error || null,
        total:          parsed.movimientos ? parsed.movimientos.length : 0,
        sinClasif:      parsed.movimientos
          ? parsed.movimientos.filter(function(m) { return m.clasificacion === 'SIN_CLASIFICAR'; }).length : 0,
        cobranza:       parsed.movimientos
          ? parsed.movimientos.filter(function(m) { return m.esCobranza; }).length : 0,
        intercias:      parsed.stats ? (parsed.stats.intercias || 0) : 0,
        saldoInicial:   parsed.saldoInicial || null,
        validadorSaldo: parsed.validadorSaldo || null,
        advertencias:   parsed.advertencias || [],
        stats:          parsed.stats || {},
        movimientosSaldo: parsed.movimientos
          ? parsed.movimientos.map(function(m) {
              return {
                fecha:          m.fecha || '',
                descripcion:    (m.descripcion || '').slice(0, 50),
                cargo:          m.cargo  || 0,
                abono:          m.abono  || 0,
                saldo:          (m.saldo !== null && m.saldo !== undefined) ? m.saldo : null,
                saldoCalculado: (m.saldoCalculado !== undefined)            ? m.saldoCalculado : null
              };
            })
          : []
      });
    });

    return { status: 'success', data: resultados, message: '' };
  } catch(e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

/**
 * Carga movimientos en HIST_MOVIMIENTOS.
 * Post-procesamiento en una sola lectura de la hoja:
 *   1. Matching de transferencias internas (modifica en memoria)
 *   2. Reescritura de HIST si hubo pares matched
 *   3. Actualizacion de SUM_MOV para los periodos afectados
 */
function cargarMovimientos(payload) {
  // Declaradas fuera del lock para que _actualizarSumMov y _registrarSaldoInicial corran despues
  var ss          = null;
  var catMap      = null;
  var catCuentas  = null;
  var histData    = null;
  var afectados   = {};
  var doSumUpdate = false;
  var lockResult;

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    Session.getActiveUser().getEmail();
    ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'carga-mov')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };

    var histSheet = _ensureHistMovSheet(ss);
    _ensureCatClientesSheet(ss);
    _ensureSumMovSheet(ss);
    _ensureSaldoInicialSheet(ss);
    _ensureMapeoFlujoSheet(ss);

    catCuentas = _buildCatCuentasArray(ss);
    var reglas      = _buildReglasArray(ss);
    var catClientes = _buildCatClientesArray(ss);
    catMap          = _buildCatalogMap(ss);

    // Indice de cuentas internas para matching de transferencias
    var cuentasInternas = {};
    for (var ci = 0; ci < catCuentas.length; ci++) {
      if (catCuentas[ci].cuenta) cuentasInternas[catCuentas[ci].cuenta] = true;
    }

    // Leer IDs existentes para dedup por hash MD5
    var existIds = {};
    if (histSheet.getLastRow() > 1) {
      var idData = histSheet.getRange(2, 1, histSheet.getLastRow() - 1, 1).getValues();
      for (var ii = 0; ii < idData.length; ii++) {
        var id = String(idData[ii][0] || '').trim();
        if (id) existIds[id] = true;
      }
    }

    var allRows    = [];
    var errores    = [];
    var advertencias = [];
    var duplicados = 0;
    var tcMapCarga = _buildTcMap(ss);
    // afectados declarado en scope externo (visible despues del lock)

    payload.archivos.forEach(function(archivo) {
      var sheets = archivo.sheets || [];
      var banco  = archivo.banco  || '';
      var parsed = ParserMovimientos.procesarArchivo(banco, sheets, catCuentas, reglas, catClientes);

      if (!parsed.success) {
        errores.push({ archivo: archivo.nombre, error: parsed.error });
        return;
      }
      if (parsed.advertencias) {
        for (var ai = 0; ai < parsed.advertencias.length; ai++) {
          advertencias.push({ archivo: archivo.nombre, msg: parsed.advertencias[ai] });
        }
      }

      for (var mi = 0; mi < parsed.movimientos.length; mi++) {
        var mov = parsed.movimientos[mi];
        if (!mov.idMov) continue;
        if (existIds[mov.idMov]) { duplicados++; continue; }
        existIds[mov.idMov] = true;

        var fechaDate = mov.fecha ? new Date(mov.fecha + 'T00:00:00') : new Date();
        var periodo   = mov.fecha ? mov.fecha.slice(0, 7) : '';
        if (mov.cuenta && periodo) afectados[mov.cuenta + '|' + periodo] = true;

        var movMoneda = ((catMap[mov.cuenta] || {}).moneda || 'MXN').toUpperCase();
        if (movMoneda === 'MXP') movMoneda = 'MXN';
        var movTc = movMoneda === 'MXN' ? 1 : _tcFromMap(tcMapCarga, mov.fecha || '');
        allRows.push([
          mov.idMov,
          fechaDate,
          mov.sociedad      || '',
          mov.banco         || '',
          mov.cuenta        || '',
          mov.descripcion   || '',
          mov.cargo         || 0,
          mov.abono         || 0,
          mov.saldo         || 0,
          mov.referencia    || '',
          mov.claveRastreo  || '',
          mov.contraparte   || '',
          mov.tipo          || '',
          mov.clasificacion || 'SIN_CLASIFICAR',
          mov.subcategoria  || '',
          mov.reglaId       || '',
          mov.esCobranza    ? 'SI' : 'NO',
          archivo.nombre,
          'NO',  // CLASIFICACION_MANUAL
          '',    // ID_CONTRA
          mov.concepto         || '',
          mov.clienteCobranza  || '',
          movMoneda,
          movTc
        ]);
      }
    });

    // 1. Escritura batch de filas nuevas
    if (allRows.length > 0) {
      histSheet.getRange(histSheet.getLastRow() + 1, 1, allRows.length, _HIST_COLS).setValues(allRows);
    }

    // 2. Post-procesamiento: una sola lectura de HIST, matching de transferencias
    // _actualizarSumMov se ejecuta FUERA del lock (idempotente — no requiere candado)
    var totalTransferencias = 0;
    if (allRows.length > 0 && Object.keys(afectados).length > 0) {
      var fullLast = histSheet.getLastRow();
      if (fullLast > 1) {
        histData = histSheet.getRange(2, 1, fullLast - 1, _HIST_COLS).getValues();

        // Matching solo sobre cuentas del batch actual — scope completo via sincronizarTransferenciasHist
        totalTransferencias = _matchearTransferenciasEnMemoria(histData, afectados, cuentasInternas, catCuentas);
        if (totalTransferencias > 0) {
          histSheet.getRange(2, 1, fullLast - 1, _HIST_COLS).setValues(histData);
        }
        doSumUpdate = true;
      }
    }

    lockResult = {
      status:  'success',
      data:    {
        guardados:              allRows.length,
        duplicados:             duplicados,
        transferenciasInternas: totalTransferencias,
        errores:                errores,
        advertencias:           advertencias
      },
      message: allRows.length + ' movimientos guardados'
        + (duplicados           > 0 ? ' (' + duplicados           + ' duplicados omitidos)'        : '')
        + (totalTransferencias  > 0 ? ' (' + totalTransferencias  + ' transferencias internas)'    : '')
        + (errores.length       > 0 ? ' (' + errores.length       + ' archivos con error)'         : '')
    };
  } catch(e) {
    lockResult = { status: 'error', data: {}, message: e.toString() };
  } finally {
    lock.releaseLock();
  }

  // Fuera del lock: actualizar SUM_MOV, SUM_EFE y saldo inicial (idempotentes)
  if (doSumUpdate && histData && catMap) {
    try { _actualizarSumMov(ss, histData, afectados, catMap);           } catch(e2) {}
    try { _actualizarSumEfe(ss, histData, afectados, catMap);           } catch(e3) {}
    try { _registrarSaldoInicial(ss, afectados, catMap);                } catch(e4) {}
  }

  return lockResult;
}

/**
 * Elimina movimientos de HIST y SUM_MOV para un rango de meses YYYY-MM.
 * payload: { desde: 'YYYY-MM', hasta: 'YYYY-MM' }
 * Si solo se pasa { periodo }, limpia ese mes exacto.
 * Si se pasa { todo: true }, limpia todo HIST y SUM_MOV.
 */
function limpiarPeriodoMovimientos(payload) {
  try {
    Session.getActiveUser().getEmail();
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      var ss  = SpreadsheetApp.openById(SALDOS_SHEET_ID);
      if (!_tieneAccesoAVista(ss, 'carga-mov')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
      var tz  = Session.getScriptTimeZone();
      var todo = payload && payload.todo === true;
      var desde = '', hasta = '';

      if (!todo) {
        desde = String((payload && (payload.desde || payload.periodo)) || '').trim().slice(0, 7);
        hasta = String((payload && (payload.hasta || payload.periodo)) || '').trim().slice(0, 7);
        if (!desde || !hasta || !/^\d{4}-\d{2}$/.test(desde) || !/^\d{4}-\d{2}$/.test(hasta)) {
          return { status: 'error', data: {}, message: 'desde y hasta requeridos (YYYY-MM)' };
        }
        if (hasta < desde) hasta = desde;
      }

      var eliminadosHist = 0;
      var eliminadosSum  = 0;

      // HIST_MOVIMIENTOS
      var histSheet = ss.getSheetByName('HIST_MOVIMIENTOS');
      if (histSheet && histSheet.getLastRow() > 1) {
        if (todo) {
          histSheet.getRange(2, 1, histSheet.getLastRow() - 1, _HIST_COLS).clearContent();
          eliminadosHist = histSheet.getLastRow() - 1;
        } else {
          var histData = histSheet.getRange(2, 1, histSheet.getLastRow() - 1, _HIST_COLS).getValues();
          var keepHist = [];
          for (var i = 0; i < histData.length; i++) {
            var fRaw = histData[i][1];
            var fStr = fRaw instanceof Date
              ? Utilities.formatDate(fRaw, tz, 'yyyy-MM-dd')
              : String(fRaw || '').slice(0, 10);
            var fPer = fStr.slice(0, 7);
            if (fPer >= desde && fPer <= hasta) { eliminadosHist++; continue; }
            keepHist.push(histData[i]);
          }
          histSheet.getRange(2, 1, histSheet.getLastRow() - 1, _HIST_COLS).clearContent();
          if (keepHist.length) histSheet.getRange(2, 1, keepHist.length, _HIST_COLS).setValues(keepHist);
        }
      }

      // SUM_MOV
      var sumSheet = ss.getSheetByName('SUM_MOV');
      if (sumSheet && sumSheet.getLastRow() > 1) {
        if (todo) {
          sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, _SUM_COLS).clearContent();
          eliminadosSum = sumSheet.getLastRow() - 1;
        } else {
          var sumData = sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, _SUM_COLS).getValues();
          var keepSum = [];
          for (var si = 0; si < sumData.length; si++) {
            var sPer = _periodoStr(sumData[si][3]);
            if (sPer >= desde && sPer <= hasta) { eliminadosSum++; continue; }
            keepSum.push(sumData[si]);
          }
          sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, _SUM_COLS).clearContent();
          if (keepSum.length) sumSheet.getRange(2, 1, keepSum.length, _SUM_COLS).setValues(keepSum);
        }
      }

      // SUM_EFE
      var eliminadosEfe = 0;
      var efeSheet = ss.getSheetByName('SUM_EFE');
      if (efeSheet && efeSheet.getLastRow() > 1) {
        if (todo) {
          efeSheet.getRange(2, 1, efeSheet.getLastRow() - 1, _SUM_EFE_COLS).clearContent();
          eliminadosEfe = efeSheet.getLastRow() - 1;
        } else {
          var efeData = efeSheet.getRange(2, 1, efeSheet.getLastRow() - 1, _SUM_EFE_COLS).getValues();
          var keepEfe = [];
          for (var ei2 = 0; ei2 < efeData.length; ei2++) {
            var ePer2 = _periodoStr(efeData[ei2][3]);
            if (ePer2 >= desde && ePer2 <= hasta) { eliminadosEfe++; continue; }
            keepEfe.push(efeData[ei2]);
          }
          efeSheet.getRange(2, 1, efeSheet.getLastRow() - 1, _SUM_EFE_COLS).clearContent();
          if (keepEfe.length) efeSheet.getRange(2, 1, keepEfe.length, _SUM_EFE_COLS).setValues(keepEfe);
        }
      }

      var label = todo ? 'HIST completo' : (desde === hasta ? desde : desde + ' a ' + hasta);
      return {
        status:  'success',
        data:    { eliminadosHist: eliminadosHist, eliminadosSum: eliminadosSum, eliminadosEfe: eliminadosEfe },
        message: eliminadosHist + ' movimientos eliminados (' + label + ')'
      };
    } finally {
      lock.releaseLock();
    }
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  }
}

/**
 * Consulta filtrada de HIST_MOVIMIENTOS. Max 500 filas por llamada.
 * params: { desde, hasta, sociedad, banco, tipo, clasificacion, pagina, porPagina }
 */
function getMovimientos(params) {
  try {
    Session.getActiveUser().getEmail();
    var ss        = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'clasificador')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var histSheet = ss.getSheetByName('HIST_MOVIMIENTOS');
    if (!histSheet || histSheet.getLastRow() <= 1) {
      return { status: 'success', data: { rows: [], total: 0, paginas: 0 }, message: '' };
    }

    var desde      = (params && params.desde) ? _parseFechaLocal(params.desde) : null;
    var hasta      = (params && params.hasta) ? _parseFechaLocal(params.hasta) : null;
    var filtTipo      = (params && params.tipo)  ? String(params.tipo).trim().toUpperCase()  : '';
    var porPagina  = (params && params.porPagina) ? Math.min(+params.porPagina, 5000) : 200;
    var pagina     = (params && params.pagina)    ? Math.max(1, +params.pagina)      : 1;
    if (hasta) hasta.setHours(23, 59, 59, 999);

    // Filtros MultiSelect: acepta array o string simple (retrocompatible)
    function _toArr(v) {
      if (!v) return [];
      if (Array.isArray(v)) return v.filter(function(x) { return x; });
      return String(v).trim() ? [String(v).trim()] : [];
    }
    var filtSocArr    = _toArr(params && params.sociedad);
    var filtBancoArr  = _toArr(params && params.banco);
    var filtCuentaArr = _toArr(params && params.cuenta);
    var filtClasifArr = _toArr(params && params.clasificacion);

    var raw = histSheet.getRange(2, 1, histSheet.getLastRow() - 1, _HIST_COLS).getValues();

    // Colectar opciones unicas de todo el dataset (sin aplicar filtros)
    var opSoc = {}, opBan = {}, opCta = {}, opCla = {};
    for (var oi = 0; oi < raw.length; oi++) {
      if (raw[oi][2]) opSoc[String(raw[oi][2])] = true;
      if (raw[oi][3]) opBan[String(raw[oi][3])] = true;
      if (raw[oi][4]) opCta['*' + String(raw[oi][4]).slice(-4)] = true;
      if (raw[oi][13]) opCla[String(raw[oi][13])] = true;
    }

    var rows = [];
    for (var i = 0; i < raw.length; i++) {
      var r     = raw[i];
      var fecha = _parseFechaLocal(r[1]);
      if (!fecha) continue;
      if (desde && fecha < desde) continue;
      if (hasta && fecha > hasta) continue;
      if (filtSocArr.length    > 0 && filtSocArr.indexOf(String(r[2] || ''))    < 0) continue;
      if (filtBancoArr.length  > 0 && filtBancoArr.indexOf(String(r[3] || ''))  < 0) continue;
      if (filtCuentaArr.length > 0) {
        var ctaFull = String(r[4] || '');
        var ctaOk = false;
        for (var ci = 0; ci < filtCuentaArr.length; ci++) {
          var fc = String(filtCuentaArr[ci]);
          if (fc.charAt(0) === '*' ? ctaFull.slice(-4) === fc.slice(1) : ctaFull === fc) { ctaOk = true; break; }
        }
        if (!ctaOk) continue;
      }
      if (filtTipo                  && String(r[12] || '').toUpperCase() !== filtTipo)                   continue;
      if (filtClasifArr.length > 0  && filtClasifArr.indexOf(String(r[13] || '')) < 0)                   continue;

      rows.push({
        idMov:           String(r[0]  || ''),
        fecha:           Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        sociedad:        String(r[2]  || ''),
        banco:           String(r[3]  || ''),
        cuenta:          String(r[4]  || ''),
        descripcion:     String(r[5]  || ''),
        cargo:           parseFloat(r[6])  || 0,
        abono:           parseFloat(r[7])  || 0,
        saldo:           parseFloat(r[8])  || 0,
        referencia:      String(r[9]  || ''),
        contraparte:     String(r[11] || ''),
        tipo:            String(r[12] || ''),
        clasificacion:   String(r[13] || ''),
        subcategoria:    String(r[14] || ''),
        esCobranza:      String(r[16] || '') === 'SI',
        clasManual:      String(r[18] || '') === 'SI',
        idContra:        String(r[19] || ''),
        concepto:        String(r[20] || ''),
        clienteCobranza: String(r[21] || '')
      });
    }

    var total   = rows.length;
    var offset  = (pagina - 1) * porPagina;
    var paginas = Math.ceil(total / porPagina);

    var opciones = {
      sociedades:      Object.keys(opSoc).sort(),
      bancos:          Object.keys(opBan).sort(),
      cuentas:         Object.keys(opCta).sort(),
      clasificaciones: Object.keys(opCla).sort()
    };

    return {
      status:  'success',
      data:    { rows: rows.slice(offset, offset + porPagina), total: total, pagina: pagina, paginas: paginas, opciones: opciones },
      message: ''
    };
  } catch(e) {
    return { status: 'error', data: { rows: [], total: 0, paginas: 0 }, message: e.toString() };
  }
}

/**
 * Actualiza la clasificacion de movimientos individuales por ID_MOV.
 * payload: { items: [{idMov, clasificacion, subcategoria}] }
 * Marca CLASIFICACION_MANUAL='SI' y recalcula SUM_MOV para los periodos afectados.
 */
function clasificarMovimientos(payload) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    Session.getActiveUser().getEmail();
    var ss        = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'clasificador')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var histSheet = ss.getSheetByName('HIST_MOVIMIENTOS');
    if (!histSheet || histSheet.getLastRow() <= 1) {
      return { status: 'error', data: {}, message: 'HIST_MOVIMIENTOS vacia o no existe' };
    }

    var items = (payload && payload.items) ? payload.items : [];
    if (items.length === 0) return { status: 'success', data: { actualizados: 0 }, message: '' };

    // Indice por idMov
    var updateIdx = {};
    for (var ii = 0; ii < items.length; ii++) {
      var item = items[ii];
      updateIdx[String(item.idMov)] = {
        clasificacion: String(item.clasificacion || 'SIN_CLASIFICAR'),
        subcategoria:  String(item.subcategoria  || ''),
        reglaId:       'MANUAL'
      };
    }

    var lastRow   = histSheet.getLastRow();
    var histData  = histSheet.getRange(2, 1, lastRow - 1, _HIST_COLS).getValues();
    var afectados = {};
    var actualizados = 0;

    for (var i = 0; i < histData.length; i++) {
      var idMov = String(histData[i][0] || '');
      if (!updateIdx[idMov]) continue;

      var upd = updateIdx[idMov];
      histData[i][13] = upd.clasificacion;
      histData[i][14] = upd.subcategoria;
      histData[i][15] = upd.reglaId;
      histData[i][16] = (upd.clasificacion === 'COBRANZA' || (upd.clasificacion || '').toUpperCase() === 'OPERATIVOS') ? 'SI' : 'NO';  // ES_COBRANZA sincronizado con clasificacion manual
      histData[i][18] = 'SI';  // CLASIFICACION_MANUAL

      var fechaRaw = histData[i][1];
      var fechaStr = fechaRaw instanceof Date
        ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(fechaRaw).slice(0, 10);
      var cuenta  = String(histData[i][4] || '').trim();
      var periodo = fechaStr.slice(0, 7);
      if (cuenta && periodo) afectados[cuenta + '|' + periodo] = true;
      actualizados++;
    }

    // Una sola escritura batch de HIST_MOVIMIENTOS
    histSheet.getRange(2, 1, lastRow - 1, _HIST_COLS).setValues(histData);

    // Recalcular SUM_MOV y SUM_EFE para periodos afectados
    if (Object.keys(afectados).length > 0) {
      var catMap = _buildCatalogMap(ss);
      _actualizarSumMov(ss, histData, afectados, catMap);
      _actualizarSumEfe(ss, histData, afectados, catMap);
    }

    return {
      status:  'success',
      data:    { actualizados: actualizados },
      message: actualizados + ' movimientos actualizados'
    };
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

function sincronizarCobranzaHist() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    Session.getActiveUser().getEmail();
    var ss        = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'carga-mov')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var histSheet = ss.getSheetByName('HIST_MOVIMIENTOS');
    if (!histSheet || histSheet.getLastRow() <= 1) {
      return { status: 'success', data: { actualizados: 0 }, message: 'HIST_MOVIMIENTOS vacia' };
    }
    var lastRow  = histSheet.getLastRow();
    var histData = histSheet.getRange(2, 1, lastRow - 1, _HIST_COLS).getValues();
    var afectados    = {};
    var actualizados = 0;
    for (var i = 0; i < histData.length; i++) {
      var row    = histData[i];
      var tipo   = String(row[12] || '').trim();
      var clasif = String(row[13] || '').trim();
      var esCol  = String(row[16] || '').trim();
      if (tipo !== 'INGRESO' || (clasif !== 'COBRANZA' && clasif.toUpperCase() !== 'OPERATIVOS') || esCol === 'SI') continue;
      histData[i][16] = 'SI';
      if (!String(row[21] || '').trim()) {
        var contraparte  = String(row[11] || '').trim();
        var subcategoria = String(row[14] || '').trim();
        histData[i][21] = contraparte || subcategoria || '';
      }
      var fechaRaw = row[1];
      var fechaStr = fechaRaw instanceof Date
        ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(fechaRaw).slice(0, 10);
      var cuenta  = String(row[4] || '').trim();
      var periodo = fechaStr.slice(0, 7);
      if (cuenta && periodo) afectados[cuenta + '|' + periodo] = true;
      actualizados++;
    }
    if (actualizados === 0) {
      return { status: 'success', data: { actualizados: 0 }, message: 'Sin movimientos que corregir' };
    }
    histSheet.getRange(2, 1, lastRow - 1, _HIST_COLS).setValues(histData);
    var catMap = _buildCatalogMap(ss);
    _actualizarSumMov(ss, histData, afectados, catMap);
    return {
      status:  'success',
      data:    { actualizados: actualizados },
      message: actualizados + ' movimientos corregidos'
    };
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Re-clasifica traspasos e intercias en HIST_MOVIMIENTOS.
 * Scope completo: TODAS las cuentas internas en TODOS sus periodos.
 * Util cuando los dos lados de un traspaso se cargaron en sesiones separadas.
 */
function sincronizarTransferenciasHist() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    Session.getActiveUser().getEmail();
    var ss        = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'carga-mov')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var histSheet = ss.getSheetByName('HIST_MOVIMIENTOS');
    if (!histSheet || histSheet.getLastRow() <= 1) {
      return { status: 'success', data: { pares: 0 }, message: 'HIST_MOVIMIENTOS vacia' };
    }

    var catCuentas = _buildCatCuentasArray(ss);
    var catMap     = _buildCatalogMap(ss);

    var cuentasInternas = {};
    for (var ci = 0; ci < catCuentas.length; ci++) {
      if (catCuentas[ci].cuenta) cuentasInternas[catCuentas[ci].cuenta] = true;
    }

    var lastRow  = histSheet.getLastRow();
    var histData = histSheet.getRange(2, 1, lastRow - 1, _HIST_COLS).getValues();

    // Construir afectados = todos los periodos de todas las cuentas internas
    var afectados = {};
    for (var i = 0; i < histData.length; i++) {
      var cuenta = String(histData[i][4] || '').trim();
      if (!cuentasInternas[cuenta]) continue;
      var fechaRaw = histData[i][1];
      var fechaStr = fechaRaw instanceof Date
        ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(fechaRaw).slice(0, 10);
      var periodo = fechaStr.slice(0, 7);
      if (cuenta && periodo) afectados[cuenta + '|' + periodo] = true;
    }

    if (!Object.keys(afectados).length) {
      return { status: 'success', data: { pares: 0 }, message: 'Sin cuentas internas en HIST' };
    }

    var totalPares = _matchearTransferenciasEnMemoria(histData, afectados, cuentasInternas, catCuentas);

    if (totalPares > 0) {
      histSheet.getRange(2, 1, lastRow - 1, _HIST_COLS).setValues(histData);
      _actualizarSumMov(ss, histData, afectados, catMap);
    }

    return {
      status:  'success',
      data:    { pares: totalPares },
      message: totalPares > 0
        ? totalPares + ' pares de traspasos/intercias sincronizados'
        : 'Sin pares nuevos que sincronizar'
    };
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Retorna SUM_MOV para el cuadre de saldos.
 * params: { periodo } — 'YYYY-MM' o null para todos.
 */
function getCuadre(params) {
  try {
    Session.getActiveUser().getEmail();
    var ss       = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'validacion')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var sumSheet = ss.getSheetByName('SUM_MOV');
    if (!sumSheet || sumSheet.getLastRow() <= 1) {
      return { status: 'success', data: [], message: 'SUM_MOV sin datos — cargar movimientos primero' };
    }

    var filtroPeriodo = (params && params.periodo) ? String(params.periodo).trim() : '';
    var raw  = sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, _SUM_COLS).getValues();
    var rows = [];

    for (var i = 0; i < raw.length; i++) {
      var r       = raw[i];
      var periodoRaw = r[3];
      var periodo    = (periodoRaw instanceof Date)
        ? Utilities.formatDate(periodoRaw, Session.getScriptTimeZone(), 'yyyy-MM')
        : String(periodoRaw || '').trim().slice(0, 7);
      if (filtroPeriodo && periodo !== filtroPeriodo) continue;

      var saldoIni       = r[5]  !== '' ? (parseFloat(r[5])  || 0) : null;
      var saldoFinal     = r[12] !== '' ? (parseFloat(r[12]) || 0) : null;
      var saldoReportado = r[16] !== '' ? (parseFloat(r[16]) || 0) : null;
      var abonosNetos    = parseFloat(r[10]) || 0;
      var cargosNetos    = parseFloat(r[11]) || 0;

      rows.push({
        cuenta:          String(r[0] || ''),
        banco:           String(r[1] || ''),
        sociedad:        String(r[2] || ''),
        periodo:         periodo,
        tipoCuenta:      String(r[4] || ''),
        saldoInicial:    saldoIni,
        totalAbonos:     parseFloat(r[6]) || 0,
        totalCargos:     parseFloat(r[7]) || 0,
        abonosInternos:  parseFloat(r[8]) || 0,
        cargosInternos:  parseFloat(r[9]) || 0,
        abonosNetos:     abonosNetos,
        cargosNetos:     cargosNetos,
        saldoFinalCalc:  saldoFinal,
        saldoReportado:  saldoReportado,
        totalCobranza:   parseFloat(r[13]) || 0,
        conteo:          parseFloat(r[14]) || 0,
        tieneBase:       saldoIni !== null
      });
    }

    rows.sort(function(a, b) {
      var ka = a.sociedad + a.banco + a.cuenta + a.periodo;
      var kb = b.sociedad + b.banco + b.cuenta + b.periodo;
      return ka < kb ? -1 : ka > kb ? 1 : 0;
    });

    // Cuentas en catálogo sin movimientos para el periodo filtrado
    var catMap = _buildCatalogMap(ss);
    var cuentasConMovs = {};
    for (var ci = 0; ci < rows.length; ci++) cuentasConMovs[rows[ci].cuenta] = true;
    var sinMovimientos = [];
    for (var catCta in catMap) {
      if (!cuentasConMovs[catCta]) {
        var cat = catMap[catCta];
        sinMovimientos.push({ cuenta: '*' + catCta.slice(-4), banco: cat.banco || '', sociedad: cat.nombreCorto || cat.nombreSociedad || '' });
      }
    }

    return { status: 'success', data: { rows: rows, sinMovimientos: sinMovimientos }, message: '' };
  } catch(e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

function _derivarTipoMov(descripcion, claveRastreo, clasif) {
  var cl = (clasif || '').trim().toUpperCase();
  if (cl === 'TRASPASOS_INTERNOS' || cl === 'INTERCIAS' || cl === 'TRANSFERENCIA_INTERNA') return 'Transferencia';
  var clave = (claveRastreo || '').trim();
  if (clave) return 'SPEI';
  var desc = (descripcion || '').toUpperCase();
  if (desc.indexOf('TPV')           >= 0 || desc.indexOf('TERMINAL')       >= 0 ||
      desc.indexOf('PUNTO DE VENTA') >= 0 || desc.indexOf('AMEXCO')         >= 0 ||
      desc.indexOf('BNET')          >= 0 || desc.indexOf('OPENMX')         >= 0) return 'TPV';
  if (desc.indexOf('EFECTIVO') >= 0) return 'Efectivo';
  if (desc.indexOf('SPEI')     >= 0) return 'SPEI';
  return 'Otro';
}

/**
 * Retorna movimientos de cobranza / todos los ingresos de cuentas operativas.
 * params: { desde, hasta, hoy, soloCobranza }
 *   soloCobranza=true (default): solo ES_COBRANZA='SI'
 *   soloCobranza=false: todos los abonos excepto TRASPASOS_INTERNOS / INTERCIAS / PRESTAMOS
 *   hoy=true: filtra solo el dia actual
 */
function getCobranza(params) {
  try {
    Session.getActiveUser().getEmail();
    var ss        = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'cobranza')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var histSheet = ss.getSheetByName('HIST_MOVIMIENTOS');
    if (!histSheet || histSheet.getLastRow() <= 1) {
      return { status: 'success', data: { registros: [], total: 0, porSociedad: [], tcUsd: 1 }, message: '' };
    }

    var soloCobranza = (params && params.soloCobranza !== undefined) ? !!params.soloCobranza : true;
    var desde = null, hasta = null;
    if (params && params.hoy) {
      var hoyStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      desde = _parseFechaLocal(hoyStr);
      hasta = _parseFechaLocal(hoyStr);
    } else {
      desde = (params && params.desde) ? _parseFechaLocal(params.desde) : null;
      hasta = (params && params.hasta) ? _parseFechaLocal(params.hasta) : null;
    }
    if (hasta) hasta.setHours(23, 59, 59, 999);

    var catMap = _buildCatalogMap(ss);
    var tcUsd  = _getLatestTcUsd(ss);
    var raw    = histSheet.getRange(2, 1, histSheet.getLastRow() - 1, _HIST_COLS).getValues();

    var EXCLUIR = { TRASPASOS_INTERNOS: true, INTERCIAS: true, TRANSFERENCIA_INTERNA: true, PRESTAMOS: true, PRESTAMOS_INTERNOS: true, INVERSION: true };
    var registros = [];
    var totalMxn  = 0;

    for (var i = 0; i < raw.length; i++) {
      var r      = raw[i];
      var abono  = parseFloat(r[7]) || 0;
      if (abono <= 0) continue;

      var cuenta     = String(r[4] || '').trim();
      var cat        = catMap[cuenta] || {};
      if (!cat.tipoCuenta || !cat.tipoCuenta.toUpperCase().includes('OPERAT')) continue;

      var esCobranza = (String(r[16] || '') === 'SI');
      var clasif     = String(r[13] || '').trim().toUpperCase();

      if (soloCobranza) {
        if (!esCobranza) continue;
      } else {
        if (EXCLUIR[clasif]) continue;
      }

      var fecha = _parseFechaLocal(r[1]);
      if (!fecha) continue;
      if (desde && fecha < desde) continue;
      if (hasta && fecha > hasta) continue;

      var moneda   = (cat.moneda || 'MXN').toUpperCase();
      var tc       = (moneda === 'USD' && tcUsd > 1) ? tcUsd : 1;
      var abonoMxn = abono * tc;
      totalMxn    += abonoMxn;

      registros.push({
        fecha:           Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
        sociedad:        String(r[2]  || ''),
        banco:           String(r[3]  || ''),
        cuenta:          cuenta,
        descripcion:     String(r[5]  || ''),
        abono:           abono,
        moneda:          moneda,
        tc:              moneda !== 'MXN' ? tc : null,
        abonoMxn:        abonoMxn,
        clasificacion:   String(r[13] || ''),
        esCobranza:      esCobranza,
        abrCobranza:     cat.abrCobranza || '',
        concepto:        String(r[20] || ''),
        clienteCobranza: String(r[21] || ''),
        tipoMov:         _derivarTipoMov(String(r[5] || ''), String(r[10] || ''), String(r[13] || ''))
      });
    }

    var gruposMap = {};
    for (var ri = 0; ri < registros.length; ri++) {
      var rec = registros[ri];
      var abr = rec.abrCobranza || rec.sociedad.substring(0, 3).toUpperCase();
      var gk  = abr + '||' + rec.sociedad;
      if (!gruposMap[gk]) gruposMap[gk] = { abr: abr, sociedad: rec.sociedad, totalMxn: 0, cuentas: {} };
      gruposMap[gk].totalMxn += rec.abonoMxn;
      var ck = rec.banco + '|' + rec.cuenta;
      if (!gruposMap[gk].cuentas[ck]) gruposMap[gk].cuentas[ck] = { banco: rec.banco, cuenta: rec.cuenta, importeMxn: 0, moneda: rec.moneda };
      gruposMap[gk].cuentas[ck].importeMxn += rec.abonoMxn;
    }
    var porSociedad = Object.keys(gruposMap).map(function(gk) {
      var g = gruposMap[gk];
      return { abr: g.abr, sociedad: g.sociedad, totalMxn: g.totalMxn, cuentas: Object.values ? Object.values(g.cuentas) : Object.keys(g.cuentas).map(function(k) { return g.cuentas[k]; }) };
    });
    porSociedad.sort(function(a, b) { return b.totalMxn - a.totalMxn; });

    return {
      status:  'success',
      data:    { registros: registros, total: totalMxn, porSociedad: porSociedad, tcUsd: tcUsd },
      message: ''
    };
  } catch(e) {
    return { status: 'error', data: { registros: [], total: 0, porSociedad: [], tcUsd: 1 }, message: e.toString() };
  }
}

// =============================================
// CONCILIACION MENSUAL — PIVOT DE MOVIMIENTOS
// =============================================

/**
 * Agrega HIST_MOVIMIENTOS por sociedad/banco/cuenta para el pivot de Conciliacion Mensual.
 * params: { periodo } — 'YYYY-MM'. Sin periodo devuelve el mes mas reciente.
 */
function getConciliacion(params) {
  try {
    Session.getActiveUser().getEmail();
    var ss        = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'conciliacion')) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var histSheet = ss.getSheetByName('HIST_MOVIMIENTOS');
    if (!histSheet || histSheet.getLastRow() <= 1) {
      return { status: 'success', data: { filas: [], periodos: [], sinMovimientos: [] }, message: '' };
    }

    function normIngreso(c) {
      var k = (c || '').toUpperCase().replace(/\s+/g, '_').trim();
      if (k === 'COBRANZA' || k === 'OPERATIVOS' || k === 'DEPOSITOS')                                                    return 'OPERATIVOS';
      if (k === 'INTERCIAS')                                                                                               return 'INTERCIAS';
      if (k === 'TRASPASOS_INTERNOS' || k === 'TRASPASOS' || k === 'TRASPASOS_INTERCIA' || k === 'TRANSFERENCIA_INTERNA') return 'TRASPASOS';
      if (k === 'INVERSION' || k === 'INVERSIONES' || k === 'INTERESES')                                                  return 'INVERSION';
      if (k === 'PRESTAMOS' || k === 'PRESTAMOS_INTERNOS')                                                                return 'PRESTAMOS';
      if (k === 'OTROS')                                                                                                   return 'OTROS';
      return 'OTROS';
    }
    function normEgreso(c) {
      var k = (c || '').toUpperCase().replace(/\s+/g, '_').trim();
      if (k === 'PROVEEDORES')                                                                                             return 'PROVEEDORES';
      if (k === 'NOMINA')                                                                                                  return 'NOMINA';
      if (k === 'IMPUESTOS' || k === 'IVA' || k === 'ISR' || k === 'COMISIONES')                                         return 'IMPUESTOS';
      if (k === 'INTERCIAS')                                                                                               return 'INTERCIAS';
      if (k === 'TRASPASOS_INTERNOS' || k === 'TRASPASOS' || k === 'TRASPASOS_INTERCIA' || k === 'TRANSFERENCIA_INTERNA') return 'TRASPASOS';
      if (k === 'INVERSION' || k === 'INVERSIONES')                                                                       return 'INVERSION';
      if (k === 'PRESTAMOS' || k === 'PRESTAMOS_INTERNOS')                                                                return 'PRESTAMOS';
      return 'OTROS';
    }

    var raw = histSheet.getRange(2, 1, histSheet.getLastRow() - 1, _HIST_COLS).getValues();

    var periodosMap = {};
    for (var pi = 0; pi < raw.length; pi++) {
      var pr = raw[pi][1];
      var ps = pr instanceof Date
        ? Utilities.formatDate(pr, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(pr).slice(0, 10);
      var pp = ps.slice(0, 7);
      if (pp) periodosMap[pp] = true;
    }
    var periodos = Object.keys(periodosMap).sort().reverse();

    var filtroDesde = (params && params.desde)
      ? String(params.desde).trim()
      : ((params && params.periodo) ? String(params.periodo).trim() : (periodos.length > 0 ? periodos[0] : ''));
    var filtroHasta = (params && params.hasta)
      ? String(params.hasta).trim()
      : filtroDesde;

    function tipoOrden(tc) {
      var t = (tc || '').toUpperCase();
      if (t.includes('OPERAT')) return 1;
      if (t === 'JV')           return 2;
      if (t.includes('NO') && t.includes('OPERAT')) return 3;
      if (t.includes('RESTRI')) return 4;
      return 5;
    }

    var catMap = _buildCatalogMap(ss);

    var filtSoc = (params && params.sociedad) ? (Array.isArray(params.sociedad) ? params.sociedad : [params.sociedad]) : [];
    var filtBan = (params && params.banco)    ? (Array.isArray(params.banco)    ? params.banco    : [params.banco])    : [];
    var filtCta = (params && params.cuenta)   ? (Array.isArray(params.cuenta)   ? params.cuenta   : [params.cuenta])   : [];

    var aggMap = {};
    for (var i = 0; i < raw.length; i++) {
      var row = raw[i];
      var fr  = row[1];
      var fs  = fr instanceof Date
        ? Utilities.formatDate(fr, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(fr).slice(0, 10);
      var pp = fs.slice(0, 7);
      if (pp < filtroDesde || pp > filtroHasta) continue;

      var soc    = String(row[2]  || '').trim();
      var ban    = String(row[3]  || '').trim();
      var cta    = String(row[4]  || '').trim();
      if (filtSoc.length && filtSoc.indexOf(soc) === -1) continue;
      if (filtBan.length && filtBan.indexOf(ban) === -1) continue;
      if (filtCta.length && filtCta.indexOf(cta) === -1) continue;

      var cargo  = parseFloat(row[6])  || 0;
      var abono  = parseFloat(row[7])  || 0;
      var tipo   = String(row[12] || '').trim();
      var clasif = String(row[13] || '').trim();
      var key    = soc + '|' + ban + '|' + cta;

      if (!aggMap[key]) {
        var catEntry = catMap[cta] || {};
        var tc = catEntry.tipoCuenta || '';
        aggMap[key] = {
          sociedad: soc, banco: ban, cuenta: cta,
          tipoCuenta: tc, tipoOrder: tipoOrden(tc),
          IN_OPERATIVOS: 0, IN_INTERCIAS: 0, IN_TRASPASOS: 0, IN_INVERSION: 0, IN_PRESTAMOS: 0, IN_OTROS: 0,
          EG_PROVEEDORES: 0, EG_NOMINA: 0, EG_IMPUESTOS: 0, EG_INTERCIAS: 0, EG_TRASPASOS: 0, EG_INVERSION: 0, EG_PRESTAMOS: 0, EG_OTROS: 0
        };
      }
      var s = aggMap[key];
      if (tipo === 'INGRESO' && abono > 0) {
        s['IN_' + normIngreso(clasif)] += abono;
      } else if (tipo === 'EGRESO' && cargo > 0) {
        s['EG_' + normEgreso(clasif)] += cargo;
      } else if (abono > 0) {
        s['IN_OTROS'] += abono;
      } else if (cargo > 0) {
        s['EG_OTROS'] += cargo;
      }
    }

    // Read SUM_MOV for saldo final and saldo reportado per cuenta/periodo
    var sumMap = {};
    var sumSheet = ss.getSheetByName('SUM_MOV');
    if (sumSheet && sumSheet.getLastRow() > 1) {
      var sumRaw = sumSheet.getRange(2, 1, sumSheet.getLastRow() - 1, _SUM_COLS).getValues();
      for (var si = 0; si < sumRaw.length; si++) {
        var sr = sumRaw[si];
        var srPer = (sr[3] instanceof Date)
          ? Utilities.formatDate(sr[3], Session.getScriptTimeZone(), 'yyyy-MM')
          : String(sr[3] || '').trim().slice(0, 7);
        if (srPer !== filtroHasta) continue;
        var srCta = String(sr[0] || '').trim();
        sumMap[srCta] = {
          saldoFinalCalc:  sr[12] !== '' ? (parseFloat(sr[12]) || 0) : null,
          saldoReportado:  sr[16] !== '' ? (parseFloat(sr[16]) || 0) : null
        };
      }
    }

    var filas = [];
    var allKeys = Object.keys(aggMap);
    for (var ki = 0; ki < allKeys.length; ki++) {
      var s = aggMap[allKeys[ki]];
      // Fix 2026-07-25: totalIn/totalEg y el objeto de salida omitian IN_PRESTAMOS,
      // EG_INVERSION y EG_PRESTAMOS, y usaban la llave inexistente IN_INVERSIONES
      // (el acumulador real es IN_INVERSION, singular) -- undefined + numero = NaN,
      // enmascarado como $0 por el "n || 0" de fmt() en el frontend.
      var totalIn  = s.IN_OPERATIVOS + s.IN_INTERCIAS + s.IN_TRASPASOS + s.IN_INVERSION + s.IN_PRESTAMOS + s.IN_OTROS;
      var totalEg  = s.EG_PROVEEDORES + s.EG_NOMINA + s.EG_IMPUESTOS + s.EG_INTERCIAS + s.EG_TRASPASOS + s.EG_INVERSION + s.EG_PRESTAMOS + s.EG_OTROS;
      var saldos   = sumMap[s.cuenta] || {};
      filas.push({
        sociedad: s.sociedad, banco: s.banco, cuenta: s.cuenta,
        tipoCuenta: s.tipoCuenta, tipoOrder: s.tipoOrder,
        IN_OPERATIVOS: s.IN_OPERATIVOS, IN_INTERCIAS: s.IN_INTERCIAS,
        IN_TRASPASOS:  s.IN_TRASPASOS,  IN_INVERSION: s.IN_INVERSION, IN_PRESTAMOS: s.IN_PRESTAMOS, IN_OTROS: s.IN_OTROS,
        EG_PROVEEDORES: s.EG_PROVEEDORES, EG_NOMINA: s.EG_NOMINA,
        EG_IMPUESTOS:   s.EG_IMPUESTOS,   EG_INTERCIAS: s.EG_INTERCIAS,
        EG_TRASPASOS:   s.EG_TRASPASOS,   EG_INVERSION: s.EG_INVERSION, EG_PRESTAMOS: s.EG_PRESTAMOS, EG_OTROS: s.EG_OTROS,
        totalIngresos:   totalIn, totalEgresos: totalEg, neto: totalIn - totalEg,
        saldoFinalCalc:  saldos.saldoFinalCalc  !== undefined ? saldos.saldoFinalCalc  : null,
        saldoReportado:  saldos.saldoReportado  !== undefined ? saldos.saldoReportado  : null
      });
    }
    filas.sort(function(a, b) {
      if (a.tipoOrder !== b.tipoOrder) return a.tipoOrder - b.tipoOrder;
      var ka = a.sociedad + a.banco + a.cuenta;
      var kb = b.sociedad + b.banco + b.cuenta;
      return ka < kb ? -1 : ka > kb ? 1 : 0;
    });

    var conMovs = {};
    for (var fi = 0; fi < filas.length; fi++) conMovs[filas[fi].cuenta] = true;
    var sinMovimientos = [];
    for (var catCta in catMap) {
      if (!conMovs[catCta]) {
        var cat = catMap[catCta];
        sinMovimientos.push({
          cuenta:   '*' + catCta.slice(-4),
          banco:    cat.banco || '',
          sociedad: cat.nombreCorto || cat.nombreSociedad || ''
        });
      }
    }

    return {
      status:  'success',
      data:    { filas: filas, periodos: periodos, sinMovimientos: sinMovimientos, periodoActivo: filtroHasta },
      message: ''
    };
  } catch(e) {
    return { status: 'error', data: { filas: [], periodos: [], sinMovimientos: [] }, message: e.toString() };
  }
}

/**
 * Recalcula SUM_MOV y SUM_EFE para los periodos indicados a partir de HIST_MOVIMIENTOS.
 * Util para poblar meses que se cargaron antes de que existiera SUM_EFE.
 * params: { periodos: ['2026-02', '2026-04'] }
 */
function recalcularSumPeriodos(params) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    Session.getActiveUser().getEmail();
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'admin-flujo')) return { status: 'error', data: {}, message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };

    var rawPeriodos = (params && Array.isArray(params.periodos)) ? params.periodos : [];
    if (!rawPeriodos.length) return { status: 'error', data: {}, message: 'Se requiere al menos un periodo YYYY-MM' };

    var perSet = {};
    for (var pi = 0; pi < rawPeriodos.length; pi++) {
      var p = String(rawPeriodos[pi]).trim().slice(0, 7);
      if (/^\d{4}-\d{2}$/.test(p)) perSet[p] = true;
    }
    if (!Object.keys(perSet).length) return { status: 'error', data: {}, message: 'Periodos invalidos (formato YYYY-MM)' };

    var histSheet = ss.getSheetByName('HIST_MOVIMIENTOS');
    if (!histSheet || histSheet.getLastRow() <= 1) {
      return { status: 'error', data: {}, message: 'HIST_MOVIMIENTOS vacia' };
    }

    var histData = histSheet.getRange(2, 1, histSheet.getLastRow() - 1, _HIST_COLS).getValues();

    var afectados = {};
    for (var i = 0; i < histData.length; i++) {
      var cuenta = String(histData[i][4] || '').trim();
      if (!cuenta) continue;
      var fechaRaw = histData[i][1];
      var fechaStr = fechaRaw instanceof Date
        ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(fechaRaw).slice(0, 10);
      var periodo = fechaStr.slice(0, 7);
      if (perSet[periodo]) afectados[cuenta + '|' + periodo] = true;
    }

    if (!Object.keys(afectados).length) {
      return { status: 'error', data: {}, message: 'Sin movimientos en HIST para los periodos indicados' };
    }

    var catMap = _buildCatalogMap(ss);
    _actualizarSumMov(ss, histData, afectados, catMap);
    _actualizarSumEfe(ss, histData, afectados, catMap);

    var perList = Object.keys(perSet).sort().join(', ');
    return {
      status:  'success',
      data:    { periodos: Object.keys(perSet).sort() },
      message: 'SUM_MOV y SUM_EFE actualizados para: ' + perList
    };
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Re-aplica REGLAS_PARSER a HIST_MOVIMIENTOS para los periodos indicados y
 * reconstruye SUM_MOV + SUM_EFE. Omite filas con CLASIFICACION_MANUAL=SI o P1/P2.
 * params: { desde: 'YYYY-MM', hasta: 'YYYY-MM' }
 */
function recalcularMovPeriodos(params) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    Session.getActiveUser().getEmail();
    var desde = String((params && params.desde) || '').trim().slice(0, 7);
    var hasta  = String((params && params.hasta)  || '').trim().slice(0, 7);
    if (!desde || !hasta || desde > hasta)
      return { status: 'error', data: {}, message: 'Se requiere desde y hasta (YYYY-MM)' };

    var ss          = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'admin-flujo')) return { status: 'error', data: {}, message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var catCuentas  = _buildCatCuentasArray(ss);
    var reglas      = _buildReglasArray(ss);
    var catClientes = _buildCatClientesArray(ss);
    var catMap      = _buildCatalogMap(ss);

    var histSheet = ss.getSheetByName('HIST_MOVIMIENTOS');
    if (!histSheet || histSheet.getLastRow() <= 1)
      return { status: 'error', data: {}, message: 'HIST_MOVIMIENTOS vacia' };

    var histData = histSheet.getRange(2, 1, histSheet.getLastRow() - 1, _HIST_COLS).getValues();
    var idx          = ParserMovimientos.buildIdx(catCuentas, catClientes); // una sola vez
    var actualizados = 0;
    var omitidos     = 0;
    var afectados    = {};

    for (var i = 0; i < histData.length; i++) {
      var row    = histData[i];
      var cuenta = String(row[4] || '').trim();
      if (!cuenta) continue;
      var fechaRaw = row[1];
      var fechaStr = fechaRaw instanceof Date
        ? Utilities.formatDate(fechaRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd')
        : String(fechaRaw).slice(0, 10);
      var periodo = fechaStr.slice(0, 7);
      if (periodo < desde || periodo > hasta) continue;

      // Saltar clasificados manualmente o intercias detectadas por cuenta
      var reglaId  = String(row[15] || '').trim();
      var esManual = String(row[18] || '').trim().toUpperCase() === 'SI';
      if (esManual || reglaId === 'P1_TRASPASO_INTERNO' || reglaId === 'P2_INTERCIA_CUENTA_GRUPO') {
        omitidos++;
        afectados[cuenta + '|' + periodo] = true;
        continue;
      }

      var rc = ParserMovimientos.reclasificarFila(row, idx, reglas);
      row[12] = rc.tipo;
      row[13] = rc.clasificacion;
      row[14] = rc.subcategoria;
      row[15] = rc.reglaId;
      row[16] = rc.esCobranza ? 'SI' : 'NO';
      row[21] = rc.clienteCobranza;
      afectados[cuenta + '|' + periodo] = true;
      actualizados++;
    }

    if (!Object.keys(afectados).length)
      return { status: 'error', data: {}, message: 'Sin movimientos en HIST para el rango indicado' };

    histSheet.getRange(2, 1, histData.length, _HIST_COLS).setValues(histData);
    _actualizarSumMov(ss, histData, afectados, catMap);
    _actualizarSumEfe(ss, histData, afectados, catMap);

    return {
      status:  'success',
      data:    { actualizados: actualizados, omitidos: omitidos },
      message: actualizados + ' movimientos reclasificados, ' + omitidos + ' omitidos (manual/intercia). SUM actualizado.'
    };
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Genera la hoja AUDITORIA_SUM con formulas vivas referenciando HIST_MOVIMIENTOS y SUM_MOV.
 * Una fila por cuenta x periodo 2026 (12 meses x todas las cuentas del catalogo).
 * Cols A-D: estaticos. E-K: formulas SUMPRODUCT para validar fuentes.
 */
function generarHojaAuditoria() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    Session.getActiveUser().getEmail();
    var ss     = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'carga-mov')) return { status: 'error', data: {}, message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var catMap = _buildCatalogMap(ss);

    var histSheet = ss.getSheetByName('HIST_MOVIMIENTOS');
    if (!histSheet) {
      return { status: 'error', data: {}, message: 'HIST_MOVIMIENTOS no encontrada' };
    }
    var histData = histSheet.getDataRange().getValues();
    // B=fecha(idx1), E=cuenta(idx4)
    var pairsMap = {};
    for (var i = 1; i < histData.length; i++) {
      var row    = histData[i];
      var fecha  = row[1];
      var cuenta = String(row[4] || '').trim();
      if (!fecha || !cuenta) continue;
      var fechaStr = Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var key = fechaStr + '|' + cuenta;
      if (!pairsMap[key]) pairsMap[key] = { fecha: fechaStr, cuenta: cuenta };
    }

    var pairs = Object.keys(pairsMap).map(function(k) { return pairsMap[k]; })
      .sort(function(a, b) {
        var ka = a.cuenta + '|' + a.fecha, kb = b.cuenta + '|' + b.fecha;
        return ka < kb ? -1 : ka > kb ? 1 : 0;
      });

    if (pairs.length === 0) {
      return { status: 'error', data: {}, message: 'Sin datos en HIST_MOVIMIENTOS' };
    }

    var auditSheet = ss.getSheetByName('AUDITORIA_SUM');
    if (auditSheet) { auditSheet.clearContents(); } else { auditSheet = ss.insertSheet('AUDITORIA_SUM'); }

    var headers = ['FECHA','CUENTA','BANCO','SOCIEDAD',
                   'MOV_HIST','ABONOS_HIST','CARGOS_HIST',
                   'ABONOS_SUM_MES','CARGOS_SUM_MES'];
    auditSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    var hn      = Math.max(histData.length + 10, 2000);
    var sn      = 1000;
    var staticV = [];
    var fmls    = [];

    for (var pi = 0; pi < pairs.length; pi++) {
      var p   = pairs[pi];
      var cat = catMap[p.cuenta] || {};
      var r   = pi + 2;

      staticV.push([p.fecha, p.cuenta, cat.banco || '', cat.nombreCorto || cat.nombreSociedad || '']);

      // &"" forces string comparison to avoid type mismatch (account nums can be numeric in sheet)
      var H    = 'HIST_MOVIMIENTOS!';
      var cond = '((' + H + '$E$2:$E$' + hn + ')&""=B' + r + '&"")' +
                 '*(TEXT(' + H + '$B$2:$B$' + hn + ',"YYYY-MM-DD")=A' + r + ')';
      var fMov = '=SUMPRODUCT(' + cond + ')';
      var fAbo = '=SUMPRODUCT(' + cond + '*(' + H + '$H$2:$H$' + hn + '))';
      var fCar = '=SUMPRODUCT(' + cond + '*(' + H + '$G$2:$G$' + hn + '))';

      // SUM_MOV monthly reference: match by LEFT 7 chars of date (YYYY-MM)
      var S     = 'SUM_MOV!';
      var sc    = '((' + S + '$A$2:$A$' + sn + ')&""=B' + r + '&"")' +
                  '*(LEFT(' + S + '$D$2:$D$' + sn + ',7)=LEFT(A' + r + ',7))';
      var fAboS = '=SUMPRODUCT(' + sc + '*(' + S + '$G$2:$G$' + sn + '))';
      var fCarS = '=SUMPRODUCT(' + sc + '*(' + S + '$H$2:$H$' + sn + '))';

      fmls.push([fMov, fAbo, fCar, fAboS, fCarS]);
    }

    var total = staticV.length;
    auditSheet.getRange(2, 1, total, 4).setValues(staticV);
    auditSheet.getRange(2, 5, total, 5).setFormulas(fmls);
    auditSheet.setFrozenRows(1);
    auditSheet.setFrozenColumns(4);
    auditSheet.getRange(1, 1, 1, headers.length)
      .setBackground('#635bff').setFontColor('#ffffff').setFontWeight('bold');

    return {
      status:  'success',
      data:    { filas: total },
      message: total + ' filas diarias generadas en AUDITORIA_SUM'
    };
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Genera la hoja AUDITORIA_COBRANZA con formulas vivas referenciando HIST_MOVIMIENTOS y SUM_MOV.
 * Una fila por cuenta x periodo 2026. Valida que ES_COBRANZA='SI' en HIST coincida con SUM_MOV.
 */
function generarHojaAuditoriaCobranza() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    Session.getActiveUser().getEmail();
    var ss     = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoAVista(ss, 'carga-mov')) return { status: 'error', data: {}, message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var catMap = _buildCatalogMap(ss);

    var histSheet = ss.getSheetByName('HIST_MOVIMIENTOS');
    if (!histSheet) {
      return { status: 'error', data: {}, message: 'HIST_MOVIMIENTOS no encontrada' };
    }
    var histData = histSheet.getDataRange().getValues();
    // B=fecha(idx1), E=cuenta(idx4), Q=ES_COBRANZA(idx16)
    var pairsMap = {};
    for (var i = 1; i < histData.length; i++) {
      var row        = histData[i];
      var fecha      = row[1];
      var cuenta     = String(row[4]  || '').trim();
      var esCobranza = String(row[16] || '').toUpperCase();
      if (!fecha || !cuenta || esCobranza !== 'SI') continue;
      var fechaStr = Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      var key = fechaStr + '|' + cuenta;
      if (!pairsMap[key]) pairsMap[key] = { fecha: fechaStr, cuenta: cuenta };
    }

    var pairs = Object.keys(pairsMap).map(function(k) { return pairsMap[k]; })
      .sort(function(a, b) {
        var ka = a.cuenta + '|' + a.fecha, kb = b.cuenta + '|' + b.fecha;
        return ka < kb ? -1 : ka > kb ? 1 : 0;
      });

    if (pairs.length === 0) {
      return { status: 'error', data: {}, message: 'Sin movimientos con ES_COBRANZA=SI en HIST' };
    }

    var audSheet = ss.getSheetByName('AUDITORIA_COBRANZA');
    if (audSheet) { audSheet.clearContents(); } else { audSheet = ss.insertSheet('AUDITORIA_COBRANZA'); }

    var headers = ['FECHA','CUENTA','BANCO','SOCIEDAD',
                   'MOV_COBRANZA','TOTAL_COB_HIST_DIA',
                   'TOTAL_COB_SUM_MES'];
    audSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    var hn      = Math.max(histData.length + 10, 2000);
    var sn      = 1000;
    var staticV = [];
    var fmls    = [];

    for (var pi = 0; pi < pairs.length; pi++) {
      var p   = pairs[pi];
      var cat = catMap[p.cuenta] || {};
      var r   = pi + 2;

      staticV.push([p.fecha, p.cuenta, cat.banco || '', cat.nombreCorto || cat.nombreSociedad || '']);

      // &"" forces string comparison; Q=ES_COBRANZA col (idx 17 → col Q)
      var H    = 'HIST_MOVIMIENTOS!';
      var cond = '((' + H + '$E$2:$E$' + hn + ')&""=B' + r + '&"")' +
                 '*(TEXT(' + H + '$B$2:$B$' + hn + ',"YYYY-MM-DD")=A' + r + ')' +
                 '*((' + H + '$Q$2:$Q$' + hn + ')="SI")';
      var fMov = '=SUMPRODUCT(' + cond + ')';
      var fAbo = '=SUMPRODUCT(' + cond + '*(' + H + '$H$2:$H$' + hn + '))';

      // SUM_MOV col R = TOTAL_COBRANZA_MXN; match by month
      var S     = 'SUM_MOV!';
      var sc    = '((' + S + '$A$2:$A$' + sn + ')&""=B' + r + '&"")' +
                  '*(LEFT(' + S + '$D$2:$D$' + sn + ',7)=LEFT(A' + r + ',7))';
      var fCobS = '=SUMPRODUCT(' + sc + '*(' + S + '$R$2:$R$' + sn + '))';

      fmls.push([fMov, fAbo, fCobS]);
    }

    var total = staticV.length;
    audSheet.getRange(2, 1, total, 4).setValues(staticV);
    audSheet.getRange(2, 5, total, 3).setFormulas(fmls);
    audSheet.setFrozenRows(1);
    audSheet.setFrozenColumns(4);
    audSheet.getRange(1, 1, 1, headers.length)
      .setBackground('#10b981').setFontColor('#ffffff').setFontWeight('bold');

    return {
      status:  'success',
      data:    { filas: total },
      message: total + ' filas diarias generadas en AUDITORIA_COBRANZA'
    };
  } catch(e) {
    return { status: 'error', data: {}, message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// =============================================
// ACTIVOS FIJOS
// =============================================

function saveAssetRequest(data) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      // Fix 2026-07-25: getActiveSpreadsheet() devuelve null en un proyecto
      // standalone ejecutado como app web (sin hoja contenedora) -- la
      // funcion fallaba siempre. Usa la misma BD que el resto del backend.
      var ss    = SpreadsheetApp.openById(SALDOS_SHEET_ID);
      if (!_tieneAccesoAVista(ss, 'solicitud')) return { success: false, error: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
      var sheet = ss.getSheetByName('SolicitudesAF') || ss.insertSheet('SolicitudesAF');
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['Fecha', 'Folio', 'Tipo', 'Sociedad', 'Monto', 'Moneda', 'Descripcion', 'Justificacion']);
      }
      var folio = 'AF-' + Math.floor(1000 + Math.random() * 9000);
      sheet.appendRow([new Date(), folio, data.tipo, data.sociedad, data.monto, data.moneda, data.descripcion, data.justificacion]);
      return { success: true, folio: folio };
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}
