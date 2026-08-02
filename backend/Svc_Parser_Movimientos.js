/**
 * VEVA Enterprise — Svc_Parser_Movimientos.js  v2.0.0
 * Parser de movimientos bancarios + motor de clasificacion por reglas.
 * Bancos: Santander (multi-hoja), BBVA (multi-hoja), Banorte, Inbursa.
 *
 * Cambios v2:
 *  - Banorte: captura nomContraparte desde col DESCRIPCION DETALLADA (L)
 *  - BBVA: captura nomContraparte desde col CONCEPTO (C)
 *  - _buildIdx: corregido check de activo (el arreglo ya viene filtrado)
 *  - _matchCliente: soporta STARTS_WITH / CONTAINS / EQUALS (era solo PREFIJO)
 *  - _clasificarMov: cobranza se evalua para cualquier INGRESO (no solo Operativos)
 *    y actualiza clasificacion si el movimiento no tenia una asignada
 *  - _build: expone clasificacionManual e idContra para columnas S y T de HIST
 */
var ParserMovimientos = (function () {

  // =========================================================================
  // UTILIDADES
  // =========================================================================

  function _limpiarMonto(v) {
    if (v === null || v === undefined || v === '' || v === '-') return 0;
    if (typeof v === 'number') return isNaN(v) ? 0 : Math.abs(v);
    var s = String(v).trim().replace(/[$,\s()']/g, '');
    var n = parseFloat(s.replace(/[^0-9.\-]/g, ''));
    return isNaN(n) ? 0 : Math.abs(n);
  }

  function _limpiarCuenta(v) {
    if (!v && v !== 0) return '';
    var s = String(v).trim().replace(/[^0-9]/g, '');
    return s.replace(/^0+(\d)/, '$1') || s;
  }

  function _fmtFecha(d) {
    if (!d || isNaN(d.getTime())) return '';
    return d.getFullYear() + '-' +
      ('0' + (d.getMonth() + 1)).slice(-2) + '-' +
      ('0' + d.getDate()).slice(-2);
  }

  function _parseFecha(v) {
    if (!v && v !== 0) return null;
    if (v instanceof Date) {
      if (isNaN(v.getTime())) return null;
      var yy = v.getFullYear();
      if (yy < 2000 || yy > 2100) return null;
      return new Date(yy, v.getMonth(), v.getDate(), 0, 0, 0, 0);
    }
    var s = String(v).trim().replace(/'/g, '');
    if (!s) return null;
    var m;
    // 8 digitos sin separador: DDMMAAAA (Santander) o YYYYMMDD
    m = s.match(/^(\d{8})$/);
    if (m) {
      var d1 = +s.slice(0,2), m1 = +s.slice(2,4), y1 = +s.slice(4,8);
      if (y1 >= 2000 && y1 <= 2100 && m1 >= 1 && m1 <= 12 && d1 >= 1 && d1 <= 31)
        return new Date(y1, m1 - 1, d1, 0, 0, 0, 0);
      var y2 = +s.slice(0,4), m2 = +s.slice(4,6), d2 = +s.slice(6,8);
      if (y2 >= 2000 && y2 <= 2100 && m2 >= 1 && m2 <= 12 && d2 >= 1 && d2 <= 31)
        return new Date(y2, m2 - 1, d2, 0, 0, 0, 0);
    }
    // DD/MM/YYYY o D/M/YYYY (4 digitos de año)
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1], 0, 0, 0, 0);
    // M/D/YY o D/M/YY con año de 2 digitos (BBVA y Banorte exportan en este formato)
    // Desambiguar: si el segundo segmento > 12 es el dia → M/D/YY; si el primero > 12 → D/M/YY
    m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);
    if (m) {
      var pa = +m[1], pb = +m[2], pyy = +m[3] + 2000;
      if (pa >= 1 && pa <= 12 && pb >= 1 && pb <= 31) {
        // Si pb > 12 solo puede ser dia (M/D/YY); si pa > 12 solo puede ser dia de D/M/YY
        // Cuando ambos <= 12 asumir M/D/YY (formato US que usan BBVA y Banorte)
        if (pa > 12) return new Date(pyy, pb - 1, pa, 0, 0, 0, 0); // D/M/YY
        return new Date(pyy, pa - 1, pb, 0, 0, 0, 0);               // M/D/YY
      }
    }
    // YYYY-MM-DD
    m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3], 0, 0, 0, 0);
    // Excel serial (BBVA/Banorte cuando XLSX convierte sin formato fecha)
    var n = parseFloat(s);
    if (!isNaN(n) && n > 40000 && n < 60000) {
      var utc = new Date((n - 25569) * 86400000);
      return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate(), 0, 0, 0, 0);
    }
    return null;
  }

  function _hashMov(fecha, cuenta, cargo, abono, desc, saldo) {
    var raw = [
      String(fecha  || ''),
      String(cuenta || ''),
      Math.round((cargo  || 0) * 100),
      Math.round((abono  || 0) * 100),
      String(desc || '').toUpperCase().trim().slice(0, 80),
      Math.round((saldo  || 0) * 100)
    ].join('|');
    var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw);
    return bytes.map(function (b) {
      return ('0' + (b & 0xFF).toString(16)).slice(-2);
    }).join('').slice(0, 16);
  }

  function _norm(v) { return String(v || '').toUpperCase().trim(); }

  function _parsearCSV(csv) {
    if (Array.isArray(csv)) return csv;
    try   { return Utilities.parseCsv(String(csv)); }
    catch (e) {
      return String(csv).split(/\r\n|\n/).map(function (l) { return l.split(','); });
    }
  }

  // Busca el primer indice cuyo header normalizado coincide con alguna clave
  function _ci(hdr, keys) {
    for (var k = 0; k < keys.length; k++) {
      var idx = hdr.indexOf(keys[k]);
      if (idx >= 0) return idx;
    }
    return -1;
  }

  // =========================================================================
  // PARSERS POR BANCO
  // Entrada: rows (array de arrays leido desde CSV)
  // Salida:  { movs[], saldoInicial, advertencias[] }
  // =========================================================================

  function _parsearSantander(sheetName, rows) {
    if (!rows || rows.length < 2) return { movs: [], saldoInicial: null, advertencias: [] };
    var hdr       = rows[0].map(_norm);
    var iCuenta   = _ci(hdr, ['CUENTA']);
    var iFecha    = _ci(hdr, ['FECHA']);
    var iSigno    = _ci(hdr, ['CARGO/ABONO']);
    var iImporte  = _ci(hdr, ['IMPORTE']);
    var iSaldo    = _ci(hdr, ['SALDO']);
    var iRef      = _ci(hdr, ['REFERENCIA']);
    var iConcepto = _ci(hdr, ['CONCEPTO']);
    var iDescLarga= _ci(hdr, ['DESCRIPCION LARGA']);
    var iDescCorta= _ci(hdr, ['DESCRIPCION']);
    var iNomBen   = _ci(hdr, ['NOMBRE BENEFICIARIO']);
    var iNomOrd   = _ci(hdr, ['NOMBRE ORDENANTE']);
    var iClabeBen = _ci(hdr, ['CLABE BENEFICIARIO']);
    var iClabeOrd = _ci(hdr, ['CTA ORDENANTE']);
    var iRfcBen   = _ci(hdr, ['RFC BENEFICIARIO']);
    var iRfcOrd   = _ci(hdr, ['RFC ORDENANTE']);
    var iClave    = _ci(hdr, ['CLAVE DE RASTREO']);

    if (iFecha < 0 || iImporte < 0) {
      return { movs: [], saldoInicial: null, advertencias: ['Santander: cols no encontradas en ' + sheetName] };
    }

    var movs = [];
    for (var i = 1; i < rows.length; i++) {
      var r     = rows[i];
      var fd    = _parseFecha(r[iFecha]);
      if (!fd) continue;
      var imp   = _limpiarMonto(r[iImporte]);
      if (imp === 0) continue;
      var signo = _norm(iSigno >= 0 ? r[iSigno] : '');
      var cargo = (signo === '-') ? imp : 0;
      var abono = (signo === '+') ? imp : 0;
      if (cargo === 0 && abono === 0) continue;

      // Descripcion: combinar Descripcion Larga > Concepto > Descripcion Corta
      var desc = '';
      if (iDescLarga >= 0 && String(r[iDescLarga] || '').trim()) {
        desc = String(r[iDescLarga]).trim();
      } else if (iConcepto >= 0 && String(r[iConcepto] || '').trim()) {
        desc = String(r[iConcepto]).trim();
      } else if (iDescCorta >= 0) {
        desc = String(r[iDescCorta] || '').trim();
      }

      // Contraparte: quien envia (INGRESO=Ordenante) o quien recibe (EGRESO=Beneficiario)
      var nomContrap = abono > 0
        ? String(r[iNomOrd]  >= 0 ? r[iNomOrd]  || '' : '').trim()
        : String(r[iNomBen]  >= 0 ? r[iNomBen]  || '' : '').trim();
      var ctaContrap = _limpiarCuenta(abono > 0
        ? (iClabeOrd >= 0 ? r[iClabeOrd] || '' : '')
        : (iClabeBen >= 0 ? r[iClabeBen] || '' : ''));
      var rfcContrap = abono > 0
        ? String(iRfcOrd >= 0 ? r[iRfcOrd] || '' : '').trim()
        : String(iRfcBen >= 0 ? r[iRfcBen] || '' : '').trim();

      movs.push({
        fecha:          _fmtFecha(fd),
        cuenta:         iCuenta >= 0 ? _limpiarCuenta(r[iCuenta]) : '',
        descripcion:    desc,
        cargo:          cargo,
        abono:          abono,
        saldo:          iSaldo >= 0  ? _limpiarMonto(r[iSaldo])   : null,
        referencia:     iRef   >= 0  ? String(r[iRef]   || '').trim().replace(/'/g, '') : '',
        claveRastreo:   iClave >= 0  ? String(r[iClave] || '').trim() : '',
        nomContraparte: nomContrap,
        ctaContraparte: ctaContrap,
        rfcContraparte: rfcContrap,
        concepto:       iConcepto >= 0 ? String(r[iConcepto] || '').trim() : ''
      });
    }
    return { movs: movs, saldoInicial: null, advertencias: [] };
  }

  function _parsearBBVA(sheetName, rows) {
    if (!rows || rows.length < 2) return { movs: [], saldoInicial: null, advertencias: [] };

    // BBVA puede tener hasta 7 filas de metadatos antes de los headers — escanear hasta fila 10.
    // Acepta cualquier alias de fecha + cualquier alias de monto para maxima compatibilidad.
    var FECHAS_BBVA  = ['FECHA OPERACION', 'FECHA OPERACIÓN',
                        'FECHA DE OPERACION', 'FECHA DE OPERACIÓN',
                        'FECHA VALOR', 'FECHA'];
    var MONTOS_BBVA  = ['CONCEPTO', 'MOVIMIENTO', 'DESCRIPCION', 'DESCRIPCIÓN',
                        'DETALLE', 'CARGO', 'ABONO', 'IMPORTE'];
    var hdrIdx = -1;
    for (var hi = 0; hi < Math.min(10, rows.length); hi++) {
      var hn = rows[hi].map(_norm);
      if (_ci(hn, FECHAS_BBVA) >= 0 && _ci(hn, MONTOS_BBVA) >= 0) {
        hdrIdx = hi;
        break;
      }
    }
    if (hdrIdx < 0) {
      var muestra = rows[0].map(_norm).filter(function(v) { return v; }).slice(0, 6).join(' | ');
      return { movs: [], saldoInicial: null, advertencias: [
        'BBVA: encabezados no encontrados en "' + sheetName + '". Fila 0: ' + muestra
      ]};
    }

    // Cuenta: buscar primero patron explicito "Cuenta: XXXX", luego numero de 9-18 digitos
    var cuenta = '';
    for (var mi = 0; mi < hdrIdx && !cuenta; mi++) {
      for (var mj = 0; mj < rows[mi].length && !cuenta; mj++) {
        var cellStr = String(rows[mi][mj] || '').trim();
        var mc = cellStr.match(/(?:cuenta|n[uú]mero\s+(?:de\s+)?cuenta|no\.?\s*cta)[\s:.]+(\d+)/i);
        if (mc && mc[1].length >= 7) { cuenta = _limpiarCuenta(mc[1]); break; }
      }
    }
    if (!cuenta) {
      for (var mi2 = 0; mi2 < hdrIdx && !cuenta; mi2++) {
        for (var mj2 = 0; mj2 < rows[mi2].length && !cuenta; mj2++) {
          var cl = _limpiarCuenta(rows[mi2][mj2]);
          if (cl && cl.length >= 9 && cl.length <= 18) cuenta = cl;
        }
      }
    }
    if (!cuenta) cuenta = _limpiarCuenta(rows[0][1]); // ultimo recurso: celda B1

    var hdr    = rows[hdrIdx].map(_norm);
    var iFecha = _ci(hdr, FECHAS_BBVA);
    var iConc  = _ci(hdr, ['CONCEPTO', 'MOVIMIENTO', 'DESCRIPCION', 'DESCRIPCIÓN', 'DETALLE']);
    var iRef   = _ci(hdr, ['REFERENCIA AMPLIADA', 'REFERENCIA']);
    var iCargo = _ci(hdr, ['CARGO']);
    var iAbono = _ci(hdr, ['ABONO', 'DEPOSITO', 'DEPÓSITO', 'DEPOSITOS', 'DEPÓSITOS']);
    var iImp   = _ci(hdr, ['IMPORTE']);
    var iSaldo = _ci(hdr, ['SALDO']);

    // Solo se requiere fecha; cargo/abono/importe se manejan abajo
    if (iFecha < 0) {
      return { movs: [], saldoInicial: null, advertencias: [
        'BBVA: columna FECHA no encontrada en "' + sheetName + '". Headers: ' +
        hdr.filter(function(v) { return v; }).join(' | ')
      ]};
    }

    var movs = [];
    var primerFd = null, ultimaFd = null;
    for (var i = hdrIdx + 1; i < rows.length; i++) {
      var r  = rows[i];
      var fd = _parseFecha(r[iFecha]);
      if (!fd) continue;
      if (!primerFd) primerFd = fd;
      ultimaFd = fd;

      var cargo = 0, abono = 0;
      if (iCargo >= 0 || iAbono >= 0) {
        // Formato con columnas separadas CARGO / ABONO
        cargo = iCargo >= 0 ? _limpiarMonto(r[iCargo]) : 0;
        abono = iAbono >= 0 ? _limpiarMonto(r[iAbono]) : 0;
      } else if (iImp >= 0) {
        // Formato con columna IMPORTE con signo: negativo=cargo, positivo=abono
        var rawImp = String(r[iImp] || '').trim().replace(/[$,\s()']/g, '');
        var impVal = parseFloat(rawImp);
        if (!isNaN(impVal)) {
          if (impVal < 0) cargo = Math.abs(impVal);
          else            abono = impVal;
        }
      }
      if (cargo === 0 && abono === 0) continue;

      var concepto = iConc >= 0 ? String(r[iConc] || '').trim() : '';
      var refAmp   = iRef  >= 0 ? String(r[iRef]  || '').trim() : '';

      // Concepto hex (deposito referenciado BBVA): usar REFERENCIA AMPLIADA como nombre y descripcion
      var nomContrapBBVA = concepto;
      var descBBVA       = concepto;
      if (/^[0-9a-fA-F]{14,20}$/.test(concepto) && refAmp) {
        nomContrapBBVA = refAmp;
        descBBVA       = refAmp;
      }

      movs.push({
        fecha:          _fmtFecha(fd),
        cuenta:         cuenta,
        descripcion:    descBBVA,
        cargo:          cargo,
        abono:          abono,
        saldo:          iSaldo >= 0 ? _limpiarMonto(r[iSaldo]) : null,
        referencia:     refAmp,
        claveRastreo:   '',
        nomContraparte: nomContrapBBVA,
        ctaContraparte: '',
        rfcContraparte: '',
        concepto:       ''
      });
    }

    // Invertir solo si el archivo viene descendente (primero=mas reciente)
    if (primerFd && ultimaFd && primerFd > ultimaFd) movs.reverse();

    var advsB = [];
    if (movs.length === 0) {
      var sRowB = rows.length > hdrIdx + 1 ? rows[hdrIdx + 1] : [];
      advsB.push(
        'BBVA: 0 movimientos en "' + sheetName + '". ' +
        'Headers[fila ' + hdrIdx + ']: ' + hdr.filter(function(v) { return v; }).join(' | ') + '. ' +
        'iFecha=' + iFecha + ' iCargo=' + iCargo + ' iAbono=' + iAbono + ' iImp=' + iImp + '. ' +
        'Muestra fila datos: ' + sRowB.slice(0, 8).map(function(v) { return String(v || '(vacio)'); }).join(' | ')
      );
    }
    return { movs: movs, saldoInicial: null, advertencias: advsB };
  }

  function _parsearBanorte(sheetName, rows) {
    if (!rows || rows.length < 2) return { movs: [], saldoInicial: null, advertencias: [] };

    // Banorte puede tener filas de metadatos antes del header — detectar dinamicamente.
    var FECHAS_BNT = ['FECHA DE OPERACIÓN', 'FECHA DE OPERACION',
                      'FECHA DE OPER', 'FECHA OPERACION', 'FECHA OPERACIÓN', 'FECHA'];
    var MONTOS_BNT = ['DEPÓSITOS', 'DEPOSITOS', 'DEPÓSITO', 'DEPOSITO',
                      'RETIROS', 'RETIRO', 'ABONO', 'ABONOS',
                      'CARGO', 'CARGOS', 'IMPORTE'];
    var hdrIdx = -1;
    for (var hi = 0; hi < Math.min(8, rows.length); hi++) {
      var hn = rows[hi].map(_norm);
      if (_ci(hn, FECHAS_BNT) >= 0 && _ci(hn, MONTOS_BNT) >= 0) {
        hdrIdx = hi;
        break;
      }
    }
    if (hdrIdx < 0) {
      var muestra = rows[0].map(_norm).filter(function(v) { return v; }).slice(0, 6).join(' | ');
      return { movs: [], saldoInicial: null, advertencias: [
        'Banorte: encabezados no encontrados en "' + sheetName + '". Fila 0: ' + muestra
      ]};
    }

    var hdr       = rows[hdrIdx].map(_norm);
    var iCta      = _ci(hdr, ['CUENTA', 'NO. CUENTA', 'NUMERO DE CUENTA', 'NÚMERO DE CUENTA']);
    var iFecha    = _ci(hdr, FECHAS_BNT);
    var iDescCorta= _ci(hdr, ['DESCRIPCIÓN', 'DESCRIPCION']);
    var iDescDet  = _ci(hdr, ['DESCRIPCIÓN DETALLADA', 'DESCRIPCION DETALLADA']);
    var iDep      = _ci(hdr, ['DEPÓSITOS', 'DEPOSITOS', 'DEPÓSITO', 'DEPOSITO', 'ABONO', 'ABONOS']);
    var iRet      = _ci(hdr, ['RETIROS', 'RETIRO', 'CARGO', 'CARGOS']);
    var iImp      = _ci(hdr, ['IMPORTE']);
    var iSaldo    = _ci(hdr, ['SALDO']);
    var iRef      = _ci(hdr, ['REFERENCIA']);

    // Si iDescDet no se encontro, buscar por substring 'DETALLADA' en el header
    if (iDescDet < 0) {
      for (var hj = 0; hj < hdr.length; hj++) {
        if (hdr[hj].indexOf('DETALLADA') >= 0 || hdr[hj].indexOf('DETAIL') >= 0) {
          iDescDet = hj; break;
        }
      }
    }

    // Requiere fecha y al menos una columna de monto
    if (iFecha < 0 || (iDep < 0 && iRet < 0 && iImp < 0)) {
      return { movs: [], saldoInicial: null, advertencias: [
        'Banorte: columnas no encontradas en "' + sheetName + '". Headers: ' +
        hdr.filter(function(v) { return v; }).join(' | ')
      ]};
    }

    // Cuenta: primero buscar en filas de metadata; fallback a primera fila de datos
    var cuenta = '';
    for (var mi = 0; mi < hdrIdx && !cuenta; mi++) {
      for (var mj = 0; mj < rows[mi].length && !cuenta; mj++) {
        var cellStr = String(rows[mi][mj] || '').trim();
        var mc = cellStr.match(/(?:cuenta|n[uú]mero\s+(?:de\s+)?cuenta|no\.?\s*cta)[\s:.]+(\d+)/i);
        if (mc && mc[1].length >= 7) { cuenta = _limpiarCuenta(mc[1]); break; }
      }
    }
    if (!cuenta) {
      for (var mi2 = 0; mi2 < hdrIdx && !cuenta; mi2++) {
        for (var mj2 = 0; mj2 < rows[mi2].length && !cuenta; mj2++) {
          var cl = _limpiarCuenta(rows[mi2][mj2]);
          if (cl && cl.length >= 9 && cl.length <= 18) cuenta = cl;
        }
      }
    }

    var movs = [];
    for (var i = hdrIdx + 1; i < rows.length; i++) {
      var r  = rows[i];
      // Fallback: tomar cuenta de la columna CUENTA en filas de datos
      if (!cuenta && iCta >= 0) {
        var ctaVal = _limpiarCuenta(r[iCta]);
        if (ctaVal) cuenta = ctaVal;
      }
      var fd = _parseFecha(r[iFecha]);
      if (!fd) continue;

      var cargo = 0, abono = 0;
      if (iDep >= 0 || iRet >= 0) {
        var rawDep = iDep >= 0 ? String(r[iDep] || '').trim() : '';
        var rawRet = iRet >= 0 ? String(r[iRet] || '').trim() : '';
        abono = (rawDep && rawDep !== '-') ? _limpiarMonto(rawDep) : 0;
        cargo = (rawRet && rawRet !== '-') ? _limpiarMonto(rawRet) : 0;
      } else if (iImp >= 0) {
        var rawImp = String(r[iImp] || '').trim().replace(/[$,\s()']/g, '');
        var impVal = parseFloat(rawImp);
        if (!isNaN(impVal)) {
          if (impVal < 0) cargo = Math.abs(impVal);
          else            abono = impVal;
        }
      }
      if (cargo === 0 && abono === 0) continue;

      var descCorta = iDescCorta >= 0 ? String(r[iDescCorta] || '').trim() : '';
      var descDet   = iDescDet   >= 0 ? String(r[iDescDet]   || '').trim() : '';

      movs.push({
        fecha:          _fmtFecha(fd),
        cuenta:         cuenta,
        descripcion:    descCorta || descDet,
        cargo:          cargo,
        abono:          abono,
        saldo:          iSaldo >= 0 ? _limpiarMonto(r[iSaldo]) : null,
        referencia:     iRef   >= 0 ? String(r[iRef] || '').trim().replace(/'/g, '') : '',
        claveRastreo:   '',
        nomContraparte: descDet || descCorta,
        ctaContraparte: '',
        rfcContraparte: '',
        concepto:       ''
      });
    }
    var advsN = [];
    if (movs.length === 0) {
      var sRowN = rows.length > hdrIdx + 1 ? rows[hdrIdx + 1] : [];
      advsN.push(
        'Banorte: 0 movimientos en "' + sheetName + '". ' +
        'Headers[fila ' + hdrIdx + ']: ' + hdr.filter(function(v) { return v; }).join(' | ') + '. ' +
        'iFecha=' + iFecha + ' iDep=' + iDep + ' iRet=' + iRet + ' iImp=' + iImp + '. ' +
        'Muestra fila datos: ' + sRowN.slice(0, 8).map(function(v) { return String(v || '(vacio)'); }).join(' | ')
      );
    }
    return { movs: movs, saldoInicial: null, advertencias: advsN };
  }

  function _parsearInbursa(sheetName, rows) {
    if (!rows || rows.length < 10) return { movs: [], saldoInicial: null, advertencias: [] };

    // Extraer numero de cuenta de las primeras 8 filas de metadatos
    var cuenta = '';
    for (var mi = 0; mi < Math.min(8, rows.length); mi++) {
      for (var ci = 0; ci < rows[mi].length; ci++) {
        var mc = String(rows[mi][ci] || '').match(/Cuenta[:\s]+(\d{10,20})/i);
        if (mc) {
          // CLABE mexicana = 18 digitos: [3 banco][3 ciudad][11 cuenta][1 control]
          // Extraer los 11 digitos de cuenta corta (pos 6-16) para coincidir con el catalogo
          cuenta = mc[1].length === 18 ? mc[1].slice(6, 17) : mc[1];
          break;
        }
      }
      if (cuenta) break;
    }

    var hdr    = rows[8].map(_norm);           // headers en fila 9 (indice 8)
    var iFecha = _ci(hdr, ['FECHA']);
    var iMovim = _ci(hdr, ['MOVIMIENTO']);
    var iCargo = _ci(hdr, ['CARGO']);
    var iAbono = _ci(hdr, ['ABONO']);
    var iRef   = _ci(hdr, ['REFERENCIA']);
    var iRefLey= _ci(hdr, ['REFERENCIA LEYENDA']);
    var iOrden = _ci(hdr, ['ORDENANTE']);
    var iClave = _ci(hdr, ['CLAVE DE RASTREO']);

    // Saldo inicial explicito en fila 10 (indice 9): fila "SALDO INICIAL"
    var saldoInicial = null;
    var r9 = rows[9] || [];
    if (iMovim >= 0 && _norm(r9[iMovim] || '') === 'SALDO INICIAL' && iAbono >= 0) {
      saldoInicial = _limpiarMonto(r9[iAbono]);
    }

    if (iFecha < 0) {
      return { movs: [], saldoInicial: saldoInicial, advertencias: ['Inbursa: cols no encontradas en ' + sheetName] };
    }

    var dataStart = (saldoInicial !== null) ? 10 : 9;
    var movs = [];
    for (var i = dataStart; i < rows.length; i++) {
      var r  = rows[i];
      var fd = _parseFecha(r[iFecha]);
      if (!fd) continue;
      var cargo = iCargo >= 0 ? _limpiarMonto(r[iCargo]) : 0;
      var abono = iAbono >= 0 ? _limpiarMonto(r[iAbono]) : 0;
      if (cargo === 0 && abono === 0) continue;

      var movDesc = iMovim >= 0 ? String(r[iMovim] || '').trim() : '';
      var refLey  = iRefLey >= 0 ? String(r[iRefLey] || '').trim() : '';
      var desc    = movDesc + (refLey ? ' / ' + refLey : '');

      movs.push({
        fecha:          _fmtFecha(fd),
        cuenta:         cuenta,
        descripcion:    desc,
        cargo:          cargo,
        abono:          abono,
        saldo:          null,
        referencia:     iRef   >= 0 ? String(r[iRef]   || '').trim() : '',
        claveRastreo:   iClave >= 0 ? String(r[iClave] || '').trim() : '',
        nomContraparte: iOrden >= 0 ? String(r[iOrden] || '').trim() : refLey,
        ctaContraparte: '',
        rfcContraparte: '',
        concepto:       ''
      });
    }
    // Derivar saldo corriente por movimiento desde saldoInicial (Inbursa no expone saldo por fila)
    if (saldoInicial !== null) {
      var runSaldo = saldoInicial;
      for (var j = 0; j < movs.length; j++) {
        runSaldo += (movs[j].abono || 0) - (movs[j].cargo || 0);
        movs[j].saldo = runSaldo;
      }
    }
    return { movs: movs, saldoInicial: saldoInicial, advertencias: [] };
  }

  // Dispatcher: elige parser por banco; fallback prueba los 4 y elige el que mas movimientos produce
  function _parsearHoja(sheetName, csv, banco) {
    var rows = _parsearCSV(csv);
    if (banco === 'Santander') return _parsearSantander(sheetName, rows);
    if (banco === 'BBVA')      return _parsearBBVA(sheetName, rows);
    if (banco === 'Banorte')   return _parsearBanorte(sheetName, rows);
    if (banco === 'Inbursa')   return _parsearInbursa(sheetName, rows);
    var candidatos = [
      _parsearSantander(sheetName, rows),
      _parsearBBVA(sheetName, rows),
      _parsearBanorte(sheetName, rows),
      _parsearInbursa(sheetName, rows)
    ];
    var mejor = candidatos[0];
    for (var i = 1; i < candidatos.length; i++) {
      if (candidatos[i].movs.length > mejor.movs.length) mejor = candidatos[i];
    }
    return mejor;
  }

  // =========================================================================
  // MOTOR DE CLASIFICACION
  // =========================================================================

  function _buildIdx(catCuentas, catClientes) {
    var porCuenta = {};
    var porNombre = [];
    for (var i = 0; i < catCuentas.length; i++) {
      var e  = catCuentas[i];
      var ct = _limpiarCuenta(e.cuenta);
      if (ct) porCuenta[ct] = e;
      if (e.nombreCorto)     porNombre.push({ pat: e.nombreCorto.toUpperCase(),     ent: e });
      if (e.nombreSociedad)  porNombre.push({ pat: e.nombreSociedad.toUpperCase(),  ent: e });
    }
    // catClientes ya viene filtrado por activo=true desde _buildCatClientesArray
    var clientes = [];
    for (var k = 0; k < catClientes.length; k++) {
      var cl = catClientes[k];
      if (cl.patron) {
        clientes.push({
          pat:    cl.patron.toUpperCase(),
          tipo:   cl.tipoMatch || 'CONTAINS',
          nombre: cl.nombre
        });
      }
    }
    return { porCuenta: porCuenta, porNombre: porNombre, clientes: clientes };
  }

  function _matchCliente(desc, contrap, clientes) {
    var d = (desc    || '').toUpperCase();
    var c = (contrap || '').toUpperCase();
    for (var i = 0; i < clientes.length; i++) {
      var cl  = clientes[i];
      var pat = cl.pat;
      var md, mc;
      if (cl.tipo === 'STARTS_WITH' || cl.tipo === 'PREFIJO') {
        md = (d.indexOf(pat) === 0);
        mc = (c.indexOf(pat) === 0);
      } else if (cl.tipo === 'EQUALS') {
        md = (d === pat);
        mc = (c === pat);
      } else {
        // CONTAINS (default)
        md = (d.indexOf(pat) >= 0);
        mc = (c.indexOf(pat) >= 0);
      }
      if (md || mc) return cl.nombre;
    }
    return null;
  }

  function _matchRegla(desc, reglas) {
    var d = (desc || '').toUpperCase().trim();
    for (var i = 0; i < reglas.length; i++) {
      var r   = reglas[i];
      if (r.fase !== 'CLASIFICACION' || !r.patron) continue;
      var pat = String(r.patron).toUpperCase();
      var ok  = r.tipoMatch === 'EQUALS'      ? (d === pat)
              : r.tipoMatch === 'STARTS_WITH' ? (d.indexOf(pat) === 0)
              : r.tipoMatch === 'CONTAINS'    ? (d.indexOf(pat) >= 0)
              : false;
      if (ok) return r;
    }
    return null;
  }

  function _detectarIntercia(mov, propia, idx) {
    // P1/P2: contraparte es cuenta conocida (por numero)
    var ctaContrap = mov.ctaContraparte;
    if (ctaContrap && idx.porCuenta[ctaContrap]) {
      var cp      = idx.porCuenta[ctaContrap];
      var interna = propia && (
        (propia.idSociedad     && cp.idSociedad     === propia.idSociedad)     ||
        (propia.nombreCorto    && cp.nombreCorto    === propia.nombreCorto)    ||
        (propia.nombreSociedad && cp.nombreSociedad === propia.nombreSociedad)
      );
      return {
        tipo:          mov.cargo > 0 ? 'EGRESO' : 'INGRESO',
        clasificacion: interna ? 'TRASPASOS_INTERNOS' : 'INTERCIAS',
        subcategoria:  interna ? 'Misma sociedad'     : 'Entre empresas del grupo',
        reglaId:       interna ? 'P1_TRASPASO_INTERNO': 'P2_INTERCIA_CUENTA_GRUPO'
      };
    }
    // P4: contraparte conocida por nombre
    if (mov.nomContraparte) {
      var nom = mov.nomContraparte.toUpperCase();
      for (var n = 0; n < idx.porNombre.length; n++) {
        if (nom.indexOf(idx.porNombre[n].pat) >= 0) {
          return {
            tipo:          mov.cargo > 0 ? 'EGRESO' : 'INGRESO',
            clasificacion: 'INTERCIAS',
            subcategoria:  'Por nombre contraparte',
            reglaId:       'P4_INTERCIA_NOMBRE'
          };
        }
      }
    }
    return null;
  }

  function _clasificarMov(mov, idx, reglas) {
    var cta    = _limpiarCuenta(mov.cuenta);
    var propia = idx.porCuenta[cta] || null;
    var soc    = propia ? (propia.nombreCorto || propia.nombreSociedad || '') : '';
    var banco  = propia ? (propia.banco || '') : '';

    // P1/P2/P4: intercias — maxima prioridad
    var intercia = _detectarIntercia(mov, propia, idx);
    if (intercia) {
      return _build(mov, soc, banco,
        intercia.tipo, intercia.clasificacion, intercia.subcategoria, intercia.reglaId, false, mov.concepto || '', '');
    }

    // P5: clasificacion por descripcion via REGLAS_PARSER
    var tipo = 'REVISION';
    var clas = 'SIN_CLASIFICAR';
    var sub  = 'Pendiente revision manual';
    var rid  = 'P99_NO_MATCH';
    var regla = _matchRegla(mov.descripcion, reglas);
    if (regla) {
      tipo = regla.tipo === 'INGRESO' ? 'INGRESO' : 'EGRESO';
      clas = regla.clasificacion || clas;
      sub  = regla.subcategoria  || sub;
      rid  = regla.reglaId       || rid;
    } else if (mov.abono > 0 && mov.cargo === 0) {
      tipo = 'INGRESO';
    } else if (mov.cargo > 0 && mov.abono === 0) {
      tipo = 'EGRESO';
    }

    // Cobranza: cualquier INGRESO cuya concepto (o desc) coincide con CAT_CLIENTES
    var esCobranza = false;
    var clienteCobranza = '';
    if (tipo === 'INGRESO') {
      var buscarEn = (mov.concepto && mov.concepto.trim()) ? mov.concepto : mov.descripcion;
      var cliente = _matchCliente(buscarEn, mov.nomContraparte, idx.clientes);
      if (cliente) {
        esCobranza = true;
        clienteCobranza = cliente;
        if (clas === 'SIN_CLASIFICAR') {
          clas = 'COBRANZA';
          sub  = 'Cobranza ' + cliente;
          rid  = 'COBRANZA_CLIENTE';
        }
      }
    }

    // Sincronizar: INGRESO clasificado COBRANZA o Operativos por REGLAS_PARSER activa el flag
    if (tipo === 'INGRESO' && (clas === 'COBRANZA' || clas.toUpperCase() === 'OPERATIVOS')) {
      esCobranza = true;
      if (!clienteCobranza) {
        clienteCobranza = (mov.nomContraparte && mov.nomContraparte.trim())
          ? mov.nomContraparte.trim()
          : (sub && sub !== 'Pendiente revision manual' ? sub : '');
      }
    }

    return _build(mov, soc, banco, tipo, clas, sub, rid, esCobranza, mov.concepto || '', clienteCobranza);
  }

  function _build(mov, soc, banco, tipo, clas, sub, rid, esCobranza, concepto, clienteCobranza) {
    return {
      idMov:              _hashMov(mov.fecha, mov.cuenta, mov.cargo, mov.abono, mov.descripcion, mov.saldo),
      fecha:              mov.fecha,
      sociedad:           soc,
      banco:              banco,
      cuenta:             mov.cuenta,
      descripcion:        mov.descripcion,
      cargo:              mov.cargo,
      abono:              mov.abono,
      saldo:              mov.saldo,
      referencia:         mov.referencia,
      claveRastreo:       mov.claveRastreo,
      contraparte:        mov.nomContraparte,
      tipo:               tipo,
      clasificacion:      clas,
      subcategoria:       sub,
      reglaId:            rid,
      esCobranza:         esCobranza,
      clasificacionManual: false,
      idContra:           '',
      concepto:           concepto || '',
      clienteCobranza:    clienteCobranza || ''
    };
  }

  // =========================================================================
  // API PUBLICA
  // =========================================================================

  /**
   * Procesa todas las hojas de un archivo bancario y clasifica los movimientos.
   * @param {string} banco       — 'Santander'|'BBVA'|'Banorte'|'Inbursa'|''(auto)
   * @param {Array}  sheets      — [{nombre, csv}] una entrada por hoja del XLSX
   * @param {Array}  catCuentas  — [{cuenta, idSociedad, nombreCorto, nombreSociedad, banco}]
   * @param {Array}  reglas      — [{fase, reglaId, tipo, clasificacion, patron, tipoMatch}]
   * @param {Array}  catClientes — [{patron, tipoMatch, nombre}] ya filtrado por activo
   * @returns {{success, movimientos, saldoInicial, advertencias, stats}}
   */
  function procesarArchivo(banco, sheets, catCuentas, reglas, catClientes) {
    try {
      var idx          = _buildIdx(catCuentas, catClientes);
      var todosMovs    = [];
      var advertencias = [];
      var saldoInicial = null;

      for (var i = 0; i < sheets.length; i++) {
        var res = _parsearHoja(sheets[i].nombre, sheets[i].csv, banco);
        if (res.saldoInicial !== null && saldoInicial === null) saldoInicial = res.saldoInicial;
        for (var j = 0; j < res.advertencias.length; j++) advertencias.push(res.advertencias[j]);
        for (var k = 0; k < res.movs.length; k++)           todosMovs.push(res.movs[k]);
      }

      var clasificados = [];
      for (var m = 0; m < todosMovs.length; m++) {
        clasificados.push(_clasificarMov(todosMovs[m], idx, reglas));
      }

      // Validador de saldo: saldo_inicial + sum(abonos) - sum(cargos) debe igualar saldo final reportado
      // BBVA ya viene invertido al orden cronologico antes de llegar aqui (ver _parsearBBVA reverse)
      var validadorSaldo = null;
      if (clasificados.length > 0) {
        var vSaldoIni = saldoInicial;
        if (vSaldoIni === null) {
          var vPm = clasificados[0];
          if (vPm.saldo !== null && vPm.saldo !== undefined && vPm.saldo !== 0) {
            vSaldoIni = Math.round((vPm.saldo - (vPm.abono || 0) + (vPm.cargo || 0)) * 100) / 100;
          }
        }
        if (vSaldoIni !== null) {
          var vRun = vSaldoIni;
          var vFallas = 0;
          for (var vv = 0; vv < clasificados.length; vv++) {
            vRun = Math.round((vRun + (clasificados[vv].abono || 0) - (clasificados[vv].cargo || 0)) * 100) / 100;
            clasificados[vv].saldoCalculado = vRun;
            var vSaldoRow = clasificados[vv].saldo;
            if (vSaldoRow !== null && vSaldoRow !== undefined && vSaldoRow !== 0) {
              if (Math.abs(vRun - vSaldoRow) > 0.02) vFallas++;
            }
          }
          var vSaldoFinalRep = null;
          for (var vu = clasificados.length - 1; vu >= 0; vu--) {
            var vSu = clasificados[vu].saldo;
            if (vSu !== null && vSu !== undefined && vSu !== 0) { vSaldoFinalRep = vSu; break; }
          }
          var vDif = vSaldoFinalRep !== null ? Math.round((vRun - vSaldoFinalRep) * 100) / 100 : null;
          validadorSaldo = {
            saldoInicial:   Math.round(vSaldoIni * 100) / 100,
            saldoCalculado: vRun,
            saldoReportado: vSaldoFinalRep,
            diferencia:     vDif,
            ok:             vDif !== null ? Math.abs(vDif) < 1.0 : null,
            filasFallidas:  vFallas
          };
        }
      }

      var stats = { total: clasificados.length, ingreso: 0, egreso: 0, revision: 0, cobranza: 0, intercias: 0 };
      for (var s = 0; s < clasificados.length; s++) {
        var c = clasificados[s];
        if (c.tipo === 'INGRESO')  stats.ingreso++;
        if (c.tipo === 'EGRESO')   stats.egreso++;
        if (c.tipo === 'REVISION') stats.revision++;
        if (c.esCobranza)          stats.cobranza++;
        if (c.clasificacion === 'INTERCIAS' || c.clasificacion === 'TRASPASOS_INTERNOS') stats.intercias++;
      }

      return {
        success:        true,
        movimientos:    clasificados,
        saldoInicial:   saldoInicial,
        validadorSaldo: validadorSaldo,
        advertencias:   advertencias,
        stats:          stats
      };
    } catch (e) {
      return {
        success: false, movimientos: [], saldoInicial: null,
        validadorSaldo: null, advertencias: [e.message], stats: {}, error: e.message
      };
    }
  }

  /**
   * Construye el indice de busqueda una sola vez para uso en lotes.
   * Llamar antes del loop y pasar idx a reclasificarFila.
   */
  function buildIdx(catCuentas, catClientes) {
    return _buildIdx(catCuentas, catClientes);
  }

  /**
   * Re-clasifica una fila de HIST_MOVIMIENTOS usando las reglas actuales.
   * Recibe idx pre-construido (buildIdx) para evitar reconstruirlo por fila.
   * Omitir filas con CLASIFICACION_MANUAL='SI' o reglaId P1/P2 antes de llamar.
   */
  function reclasificarFila(histRow, idx, reglas) {
    var mov = {
      fecha:          histRow[1],
      cuenta:         String(histRow[4]  || '').trim(),
      descripcion:    String(histRow[5]  || '').trim(),
      cargo:          parseFloat(histRow[6])  || 0,
      abono:          parseFloat(histRow[7])  || 0,
      saldo:          parseFloat(histRow[8])  || 0,
      referencia:     String(histRow[9]  || '').trim(),
      claveRastreo:   String(histRow[10] || '').trim(),
      nomContraparte: String(histRow[11] || '').trim(),
      concepto:       String(histRow[20] || '').trim(),
      ctaContraparte: ''
    };
    var r = _clasificarMov(mov, idx, reglas);
    return {
      tipo:            r.tipo,
      clasificacion:   r.clasificacion,
      subcategoria:    r.subcategoria,
      reglaId:         r.reglaId,
      esCobranza:      r.esCobranza,
      clienteCobranza: r.clienteCobranza
    };
  }

  return { procesarArchivo: procesarArchivo, buildIdx: buildIdx, reclasificarFila: reclasificarFila };

})();
