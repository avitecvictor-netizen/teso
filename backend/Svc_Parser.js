/**
 * VEVA Enterprise — Svc_Parser.js
 * v0.3.0
 *
 * Mejoras sobre v14:
 * - Configuracion por banco (Santander, Banorte, BBVA, Inbursa, Banamex)
 * - Deteccion de banco desde nombre de archivo
 * - Scoring para seleccion de columnas (no flags fragiles)
 * - limpiarCuenta preserva ceros iniciales
 * - Mapeo enriquecido devuelto para previsualizar antes de guardar
 * - Parser movimientos con scoring identico al de saldos
 * - Mensajes de error con contexto (archivo + fila)
 */

var ParserBancario = (function () {

  // =========================================================================
  // CONFIGURACION POR BANCO
  // Cada banco define sus palabras clave en orden de prioridad
  // =========================================================================
  var BANK_CFG = {
    Santander: {
      cuenta: ['CUENTA', 'NUMERO DE CUENTA', 'NO. DE CUENTA'],
      saldo:  ['DISPONIBLE', 'SALDO DISPONIBLE'],
      moneda: ['DIVISA', 'MONEDA']
    },
    Banorte: {
      cuenta: ['CONTRATO', 'CUENTA', 'NUMERO'],
      saldo:  ['DISPONIBLE', 'SALDO DISPONIBLE', 'SALDO'],
      moneda: ['MONEDA', 'TIPO MONEDA']
    },
    BBVA: {
      cuenta: ['CUENTA', 'NUMERO DE CUENTA', 'NO CUENTA'],
      saldo:  ['DISPONIBLE', 'TOTAL', 'SALDO'],
      moneda: ['MONEDA', 'DIVISA']
    },
    Inbursa: {
      cuenta: ['TITULAR', 'CUENTA', 'NO. CUENTA'],
      saldo:  ['SALDO DISPONIBLE', 'DISPONIBLE', 'SALDO'],
      moneda: ['MONEDA']
    },
    Banamex: {
      cuenta: ['CUENTA', 'NUMERO DE CUENTA', 'CLABE'],
      saldo:  ['SALDO DISPONIBLE', 'DISPONIBLE', 'SALDO'],
      moneda: ['MONEDA', 'DIVISA']
    }
  };

  // Fallback universal cuando el banco no se identifica
  var GENERIC_CFG = {
    cuenta: ['CUENTA', 'NUMERO DE CUENTA', 'CONTRATO', 'TITULAR', 'NO. DE CUENTA', 'NO CUENTA', 'NUMERO', 'CLABE'],
    saldo:  ['DISPONIBLE', 'SALDO DISPONIBLE', 'SALDO CONTABLE', 'TOTAL', 'SALDO', 'SALDO ACTUAL', 'SALDO AL DIA'],
    moneda: ['DIVISA', 'MONEDA', 'CURRENCY', 'TIPO MONEDA']
  };

  var FECHA_KEYS = ['FECHA', 'FECHA VALOR', 'FECHA OPERACION', 'FECHA MOV', 'DIA', 'F. OPERACION'];
  var DESC_KEYS  = ['DESCRIPCION', 'CONCEPTO', 'DETALLE', 'REFERENCIA', 'MOVIMIENTO', 'BENEFICIARIO'];
  var CARGO_KEYS = ['CARGO', 'RETIRO', 'DEBITO', 'EGRESOS', 'IMPORTE CARGO'];
  var ABONO_KEYS = ['ABONO', 'DEPOSITO', 'CREDITO', 'INGRESOS', 'IMPORTE ABONO'];
  var SALDO_MOV_KEYS = ['SALDO', 'SALDO ACTUAL', 'SALDO FINAL'];

  // =========================================================================
  // UTILIDADES
  // =========================================================================

  function limpiarMonto(valor) {
    if (valor === null || valor === undefined || valor === '') return 0;
    if (typeof valor === 'number') return valor;

    var s = String(valor).trim();
    // Negativos entre parentesis: (10,000.50) -> -10000.50
    var negativo = s[0] === '-' || (s.indexOf('(') >= 0 && s.indexOf(')') > s.indexOf('('));

    // Eliminar: $, €, espacios, parentesis, comillas. Preservar: digitos, punto, guion.
    var limpio = s.replace(/[$€\s()',"]/g, '').replace(/[^0-9.\-]/g, '');
    if (limpio === '' || limpio === '-') return 0;

    var n = parseFloat(limpio) || 0;
    return negativo ? -Math.abs(n) : n;
  }

  function limpiarCuenta(valor) {
    if (!valor) return '';
    var s = String(valor).trim().replace(/['"´ \-]/g, '');
    // Solo digitos — NO quitar ceros iniciales (son parte del numero de cuenta)
    return s.replace(/\D/g, '');
  }

  function detectarBanco(nombre) {
    var n = String(nombre).toUpperCase();
    if (n.indexOf('SANTANDER') >= 0) return 'Santander';
    if (n.indexOf('BANORTE')   >= 0) return 'Banorte';
    if (n.indexOf('BBVA')      >= 0) return 'BBVA';
    if (n.indexOf('INBURSA')   >= 0) return 'Inbursa';
    if (n.indexOf('BANAMEX')   >= 0 || n.indexOf('CITI') >= 0) return 'Banamex';
    return null;
  }

  function parsearFecha(valor) {
    if (!valor) return null;
    if (valor instanceof Date) {
      return isNaN(valor.getTime()) ? null
        : new Date(valor.getFullYear(), valor.getMonth(), valor.getDate(), 0, 0, 0, 0);
    }

    var s = String(valor).trim();
    if (!s) return null;

    var m;
    // YYYY-MM-DD o YYYY/MM/DD — debe ir ANTES del fallback new Date() para evitar parse UTC
    m = s.match(/^(\d{4})[\/\-](\d{2})[\/\-](\d{2})$/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3], 0, 0, 0, 0);

    // DD/MM/YYYY o DD-MM-YYYY
    m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1], 0, 0, 0, 0);

    // YYYYMMDD
    m = s.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3], 0, 0, 0, 0);

    // Formato no reconocido — no usar new Date(s) como fallback (comportamiento UTC no predecible)
    return null;
  }

  function extraerFechaDeNombre(nombre) {
    var n = String(nombre);
    var m;

    // DD_MM_YYYY o DD-MM-YYYY — patron mas especifico primero (ej. Santander: "14_04_2026 09_30_29")
    m = n.match(/(\d{2})[_\-](\d{2})[_\-](\d{4})/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1], 0, 0, 0, 0);

    // 8 digitos contiguos — intentar YYYYMMDD primero, luego DDMMYYYY
    m = n.match(/(?<!\d)(\d{8})(?!\d)/);
    if (m) {
      var raw = m[1];
      var dISO = new Date(+raw.slice(0, 4), +raw.slice(4, 6) - 1, +raw.slice(6, 8), 0, 0, 0, 0);
      var dDMY = new Date(+raw.slice(4, 8), +raw.slice(2, 4) - 1, +raw.slice(0, 2), 0, 0, 0, 0);
      var yISO = dISO.getFullYear(), yDMY = dDMY.getFullYear();
      if (yISO >= 2000 && yISO <= 2100 && !isNaN(dISO.getTime())) return dISO;
      if (yDMY >= 2000 && yDMY <= 2100 && !isNaN(dDMY.getTime())) return dDMY;
    }

    // Sin patron reconocible — el caller debe manejar null explicitamente
    return null;
  }

  /**
   * Scoring de columnas.
   * Exact match en posicion 0 de claves = 100, decrece con el indice.
   * Partial match = 50 - indice.
   * Devuelve el indice de columna con mayor score, o -1 si ninguno supera 0.
   */
  function encontrarColumna(filaNorm, claves) {
    var mejor = -1;
    var score = 0;

    filaNorm.forEach(function (celda, i) {
      var c = String(celda || '').toUpperCase().trim();
      claves.forEach(function (k, ki) {
        var s = 0;
        if (c === k) s = 100 - ki;
        else if (c.indexOf(k) >= 0 && k.length >= 4) s = 50 - ki;
        if (s > score) { score = s; mejor = i; }
      });
    });

    return score > 0 ? mejor : -1;
  }

  function parsearCSV(contenido) {
    if (Array.isArray(contenido)) return contenido;
    try {
      return Utilities.parseCsv(String(contenido));
    } catch (e) {
      Logger.log('parseCsv fallback: ' + e.message);
      return String(contenido).split(/\r\n|\n/).map(function (l) { return l.split(','); });
    }
  }

  // =========================================================================
  // PARSER SALDOS
  // =========================================================================

  /**
   * Un archivo de saldos real tiene maximo 1 fila por cuenta. Una cuenta
   * repetida es la señal de que en realidad es un archivo de MOVIMIENTOS
   * cargado por error via el flujo de saldos (incidente real: cuenta
   * 406605300, 2026-07-30 — el dedup por fecha+cuenta en
   * guardarSaldosEnHistorial se quedo silenciosamente con la primera
   * transaccion del mes en vez del saldo de cierre real).
   * Devuelve un mensaje de error, o null si no hay cuentas duplicadas.
   */
  function _validarCuentasUnicas(resultados, nombreArchivo) {
    var conteoPorCuenta = {};
    for (var ri = 0; ri < resultados.length; ri++) {
      var cta = resultados[ri].cuenta;
      conteoPorCuenta[cta] = (conteoPorCuenta[cta] || 0) + 1;
    }

    var cuentasDuplicadas = [];
    for (var ctaKey in conteoPorCuenta) {
      if (conteoPorCuenta[ctaKey] > 1) {
        cuentasDuplicadas.push(ctaKey + ' (' + conteoPorCuenta[ctaKey] + ' veces)');
      }
    }

    if (cuentasDuplicadas.length === 0) return null;

    return 'Archivo "' + nombreArchivo + '" rechazado: la(s) cuenta(s) ' + cuentasDuplicadas.join(', ')
      + ' aparece(n) mas de una vez (un archivo de saldos debe tener maximo 1 fila por cuenta). '
      + 'Esto sugiere que es un archivo de MOVIMIENTOS, no de SALDOS — verifica el archivo o cargalo '
      + 'desde el flujo de carga de movimientos correspondiente.';
  }

  function procesarArchivoSaldos(contenido, nombreArchivo, bancoManual, fechaManual, tcManual) {
    try {
      var banco     = bancoManual || detectarBanco(nombreArchivo);
      var fechaBase = fechaManual ? parsearFecha(fechaManual) || extraerFechaDeNombre(nombreArchivo)
                                 : extraerFechaDeNombre(nombreArchivo);
      var tc        = parseFloat(tcManual) || 1;
      var cfg       = BANK_CFG[banco] || GENERIC_CFG;
      var lineas    = parsearCSV(contenido);

      // Buscar fila de encabezados (primeras 30 filas)
      var idxCuenta = -1, idxSaldo = -1, idxMoneda = -1, filaInicio = -1;

      for (var i = 0; i < Math.min(30, lineas.length); i++) {
        var norm = lineas[i].map(function (c) { return String(c || '').toUpperCase().trim(); });

        var iC = encontrarColumna(norm, cfg.cuenta);
        var iS = encontrarColumna(norm, cfg.saldo);
        var iM = encontrarColumna(norm, cfg.moneda);

        if (iC >= 0 && iS >= 0 && iC !== iS) {
          idxCuenta  = iC;
          idxSaldo   = iS;
          idxMoneda  = iM;
          filaInicio = i + 1;
          break;
        }
      }

      if (filaInicio === -1) {
        return {
          success: false,
          error: 'Columnas no detectadas en "' + nombreArchivo + '"'
            + (banco ? ' (banco: ' + banco + ')' : ' — intenta especificar el banco manualmente')
        };
      }

      var resultados   = [];
      var advertencias = [];

      for (var r = filaInicio; r < lineas.length; r++) {
        var row = lineas[r];
        if (!row || row.length <= Math.max(idxCuenta, idxSaldo)) continue;

        var cuentaRaw = String(row[idxCuenta] || '').trim();
        var cuentaUpper = cuentaRaw.toUpperCase();
        if (!cuentaRaw || cuentaUpper === 'TOTAL' || cuentaUpper === 'CUENTA' || cuentaUpper === 'SUBTOTAL') continue;

        var cuenta = limpiarCuenta(cuentaRaw);
        if (cuenta.length < 5) continue;

        var saldo = limpiarMonto(row[idxSaldo]);

        // Guard: saldo no puede coincidir numericamente con la cuenta
        var cuentaNum = parseFloat(cuenta);
        if (!isNaN(cuentaNum) && cuenta.length > 8 && Math.abs(saldo - cuentaNum) < 1) {
          advertencias.push('Fila ' + (r + 1) + ': saldo descartado por coincidencia con numero de cuenta');
          continue;
        }

        var moneda = 'MXN';
        if (idxMoneda >= 0 && row[idxMoneda]) {
          var mv = String(row[idxMoneda]).toUpperCase();
          if (mv.indexOf('USD') >= 0 || mv.indexOf('DOLAR') >= 0) moneda = 'USD';
          else if (mv.indexOf('EUR') >= 0) moneda = 'EUR';
        }

        resultados.push({
          fecha:   fechaBase,
          banco:   banco || 'Desconocido',
          cuenta:  cuenta,
          saldo:   saldo,
          moneda:  moneda,
          tc:      moneda === 'MXN' ? 1 : tc,
          archivo: nombreArchivo
        });
      }

      // Guard de raiz: un archivo de saldos real tiene maximo 1 fila por
      // cuenta (incidente real: cuenta 406605300, 2026-07-30). Ver
      // _validarCuentasUnicas.
      var errorDuplicados = _validarCuentasUnicas(resultados, nombreArchivo);
      if (errorDuplicados) {
        return { success: false, error: errorDuplicados };
      }

      return { success: true, datos: resultados, advertencias: advertencias };

    } catch (e) {
      return { success: false, error: 'Error parser saldos "' + nombreArchivo + '": ' + e.message };
    }
  }

  // =========================================================================
  // PARSER MOVIMIENTOS
  // =========================================================================

  function procesarArchivoMovimientos(contenido, nombreArchivo, bancoManual) {
    try {
      var banco  = bancoManual || detectarBanco(nombreArchivo);
      var lineas = parsearCSV(contenido);

      var col = { fecha: -1, desc: -1, cargo: -1, abono: -1, saldo: -1 };
      var encabezadoOk = false;
      var movimientos  = [];

      for (var i = 0; i < lineas.length; i++) {
        var norm = lineas[i].map(function (c) { return String(c || '').toUpperCase().trim(); });

        if (!encabezadoOk) {
          var iF = encontrarColumna(norm, FECHA_KEYS);
          var iD = encontrarColumna(norm, DESC_KEYS);
          if (iF >= 0 && iD >= 0) {
            col.fecha  = iF;
            col.desc   = iD;
            col.cargo  = encontrarColumna(norm, CARGO_KEYS);
            col.abono  = encontrarColumna(norm, ABONO_KEYS);
            col.saldo  = encontrarColumna(norm, SALDO_MOV_KEYS);
            // Si no hay cargo/abono separado, buscar IMPORTE/MONTO
            if (col.cargo < 0 && col.abono < 0) {
              col.cargo = encontrarColumna(norm, ['IMPORTE', 'MONTO', 'VALOR']);
            }
            encabezadoOk = true;
          }
          continue;
        }

        var row      = lineas[i];
        var fechaVal = row[col.fecha];
        if (!fechaVal || String(fechaVal).trim() === '') continue;

        var fechaMov = parsearFecha(fechaVal);
        if (!fechaMov) continue;

        var cargo = 0, abono = 0;
        if (col.cargo >= 0 && col.abono >= 0) {
          cargo = limpiarMonto(row[col.cargo]);
          abono = limpiarMonto(row[col.abono]);
        } else if (col.cargo >= 0) {
          var m = limpiarMonto(row[col.cargo]);
          if (m < 0) cargo = Math.abs(m); else abono = m;
        }
        if (cargo === 0 && abono === 0) continue;

        movimientos.push({
          fecha:       Utilities.formatDate(fechaMov, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
          descripcion: String(row[col.desc] || '').trim(),
          cargo:       cargo,
          abono:       abono,
          saldo:       col.saldo >= 0 ? limpiarMonto(row[col.saldo]) : null,
          banco:       banco || 'Desconocido',
          archivo:     nombreArchivo
        });
      }

      if (!encabezadoOk) {
        return { success: false, error: 'Encabezados de movimientos no detectados en "' + nombreArchivo + '"' };
      }

      return { success: true, datos: movimientos };

    } catch (e) {
      return { success: false, error: 'Error parser movimientos "' + nombreArchivo + '": ' + e.message };
    }
  }

  // =========================================================================
  // EXPORTS
  // =========================================================================
  return {
    procesarArchivoSaldos:      procesarArchivoSaldos,
    procesarArchivoMovimientos: procesarArchivoMovimientos,
    detectarBanco:              detectarBanco
  };

})();
