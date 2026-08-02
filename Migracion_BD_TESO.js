/**
 * Script de migracion UNICA (correr una sola vez desde el editor de Apps
 * Script, menu Ejecutar > migrarABdTeso). No es parte del backend de TESO,
 * es una utileria de traspaso: lee la BD vieja de VEVA (solo lectura) y
 * escribe la BD nueva limpia de TESO.
 *
 * Decisiones tomadas con el usuario (2026-07-25):
 * - HIST_MOVIMIENTOS: 3 encabezados en blanco se nombran segun lo que el
 *   propio backend de VEVA ya espera (Codigo.gs:1078-1084): CLASIFICACION_MANUAL,
 *   ID_CONTRA, CONCEPTO. Se quita la 25a columna (fuera de _HIST_COLS=24, sin uso).
 * - SUM_MOV: 3 encabezados en blanco se nombran igual (Codigo.gs:1219-1226):
 *   SALDO_REPORTADO, TOTAL_COBRANZA_MXN, TOTAL_COBRANZA_USD.
 * - REGLAS_PARSER: tenia 2 juegos de columnas para lo mismo (bug: el backend
 *   normaliza encabezados y el segundo juego pisaba al primero en el mapa,
 *   Codigo.gs:974-982). Decision: TIPO = el juego que el backend ya usaba en
 *   la practica (columna "Tipo"). El juego viejo de TIPO (columna "TIPO",
 *   con valores mas especificos como INTERCIA_INTERNA/INTERCIA_GRUPO) se
 *   conserva como SUBCLASIFICACION_INTERCIA -- PENDIENTE DE VALIDAR como se
 *   usara, no esta conectada a ninguna logica todavia. Los otros 4 pares
 *   duplicados (CLASIFICACION, SUBCATEGORIA, PATRON, TIPO_MATCH) tenian
 *   valores identicos en ambos lados: se dejo uno solo.
 * - CAT_CUENTAS_MAPEO: columna "X" se quita (confirmado: es un duplicado
 *   exacto de NUMERO_CUENTA en las filas que la tienen). La 13a columna sin
 *   nombre tenia texto real (CAPTADORA PESOS, PAGADORA CANCUN, etc.) y se
 *   nombra NOTA. "ABR_ COBRANZA" (con espacio) se corrige a ABR_COBRANZA.
 * - HIST_POSICION_BANCARIA: columna sin nombre con codigos tipo VCN/DCZ/TMT
 *   se nombra ABR_COBRANZA -- CONFIRMADO en Codigo.gs:660, getHistorialPosicion
 *   ya devuelve ese mismo campo como "abrCobranza: String(row[12] || '')".
 * - TC: no tenia fila de encabezados (bug: _buildTcMap en Codigo.gs:372 se
 *   saltaba la fila 1 asumiendo que era encabezado, cuando en realidad era
 *   el tipo de cambio del 2025-01-01 -- ese dato se perdia en cada calculo).
 *   Se agrega encabezado FECHA/TC_USD y se conserva esa fila.
 * - "Claude Log" no migra (no la usa ningun codigo, es log de otra sesion).
 * - SALDO_INICIAL y AUDITORIA_COBRANZA se crean vacias con su esquema ya
 *   definido en Codigo.gs (las autogenera el backend en su primer uso).
 */
var OLD_SHEET_ID = '1OBFYLVawRKy5sLbgAExaL4Uozec5t4gIzl_w9Wqic08'; // VEVA (solo lectura, referencia)
var NEW_SHEET_ID = '1-DzPpYYViWVG0RHuGGldkDo4LUyYcCLAjQOIlp2DdS8'; // BD_TESO (destino)

function migrarABdTeso() {
  var oldSs = SpreadsheetApp.openById(OLD_SHEET_ID);
  var newSs = SpreadsheetApp.openById(NEW_SHEET_ID);

  ['SUM_EFE', 'MAPEO_FLUJO', 'CAT_DIAS_INHABILES', 'CAT_CLIENTES', 'CONFIG_DISTRIBUCION', 'AUDITORIA_SUM']
    .forEach(function (nombre) { _copiarVerbatim(oldSs, newSs, nombre); });

  _migrarHistMovimientos(oldSs, newSs);
  _migrarSumMov(oldSs, newSs);
  _migrarReglasParser(oldSs, newSs);
  _migrarCatCuentasMapeo(oldSs, newSs);
  _migrarHistPosicionBancaria(oldSs, newSs);
  _migrarTC(oldSs, newSs);

  _crearVacia(newSs, 'SALDO_INICIAL', ['CUENTA', 'BANCO', 'SOCIEDAD', 'PERIODO', 'SALDO', 'BLOQUEADO', 'FECHA_ACTUALIZACION']);
  _crearVacia(newSs, 'AUDITORIA_COBRANZA', ['FECHA', 'CUENTA', 'BANCO', 'SOCIEDAD', 'MOV_COBRANZA', 'TOTAL_COB_HIST_DIA', 'TOTAL_COB_SUM_MES']);

  Logger.log('Migracion completa.');
}

function _getOrCreate(ss, nombre) {
  var sheet = ss.getSheetByName(nombre);
  if (sheet) { sheet.clearContents(); } else { sheet = ss.insertSheet(nombre); }
  return sheet;
}

function _copiarVerbatim(oldSs, newSs, nombre) {
  var oldSheet = oldSs.getSheetByName(nombre);
  if (!oldSheet || oldSheet.getLastRow() < 1) return;
  var data = oldSheet.getRange(1, 1, oldSheet.getLastRow(), oldSheet.getLastColumn()).getValues();
  var sheet = _getOrCreate(newSs, nombre);
  sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  sheet.setFrozenRows(1);
}

function _crearVacia(newSs, nombre, headers) {
  var sheet = _getOrCreate(newSs, nombre);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

function _migrarHistMovimientos(oldSs, newSs) {
  var old = oldSs.getSheetByName('HIST_MOVIMIENTOS');
  if (!old || old.getLastRow() < 2) return;
  var headers = [
    'ID_MOV', 'FECHA', 'SOCIEDAD', 'BANCO', 'CUENTA', 'DESCRIPCION', 'CARGO', 'ABONO',
    'SALDO_BANCO', 'REFERENCIA', 'CLAVE_RASTREO', 'CONTRAPARTE', 'TIPO', 'CLASIFICACION',
    'SUBCATEGORIA', 'REGLA_ID', 'ES_COBRANZA', 'ARCHIVO_FUENTE', 'CLASIFICACION_MANUAL',
    'ID_CONTRA', 'CONCEPTO', 'CLIENTE_COBRANZA', 'MONEDA_ORIGINAL', 'TC_APLICADO'
  ];
  var data = old.getRange(2, 1, old.getLastRow() - 1, headers.length).getValues(); // corta la 25a col
  var sheet = _getOrCreate(newSs, 'HIST_MOVIMIENTOS');
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  sheet.setFrozenRows(1);
}

function _migrarSumMov(oldSs, newSs) {
  var old = oldSs.getSheetByName('SUM_MOV');
  if (!old || old.getLastRow() < 2) return;
  var headers = [
    'CUENTA', 'BANCO', 'SOCIEDAD', 'PERIODO', 'TIPO_CUENTA', 'SALDO_INICIAL', 'TOTAL_ABONOS',
    'TOTAL_CARGOS', 'ABONOS_INTERNOS', 'CARGOS_INTERNOS', 'ABONOS_NETOS', 'CARGOS_NETOS',
    'SALDO_FINAL_CALC', 'TOTAL_COBRANZA', 'CONTEO', 'FECHA_ACTUALIZACION', 'SALDO_REPORTADO',
    'TOTAL_COBRANZA_MXN', 'TOTAL_COBRANZA_USD', 'SALDO_BANCO_FINAL', 'ES_PARTIDA'
  ];
  var data = old.getRange(2, 1, old.getLastRow() - 1, headers.length).getValues();
  var sheet = _getOrCreate(newSs, 'SUM_MOV');
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  sheet.setFrozenRows(1);
}

function _migrarReglasParser(oldSs, newSs) {
  var old = oldSs.getSheetByName('REGLAS_PARSER');
  if (!old || old.getLastRow() < 2) return;
  var raw = old.getRange(2, 1, old.getLastRow() - 1, 18).getValues();
  var headers = [
    'PRIORIDAD', 'FASE', 'REGLA_ID', 'TIPO', 'SUBCLASIFICACION_INTERCIA', 'CLASIFICACION',
    'SUBCATEGORIA', 'PATRON', 'TIPO_MATCH', 'CAMPO_ESTANDAR', 'CAMPO_LOOKUP', 'ACCION',
    'DESCRIPCION', 'NOTA_CODIFICACION'
  ];
  var data = raw.map(function (r) {
    return [r[0], r[1], r[2], r[13], r[3], r[14], r[15], r[16], r[17], r[8], r[9], r[10], r[11], r[12]];
  });
  var sheet = _getOrCreate(newSs, 'REGLAS_PARSER');
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  sheet.setFrozenRows(1);
}

function _migrarCatCuentasMapeo(oldSs, newSs) {
  var old = oldSs.getSheetByName('CAT_CUENTAS_MAPEO');
  if (!old || old.getLastRow() < 2) return;
  var raw = old.getRange(2, 1, old.getLastRow() - 1, 13).getValues();
  var headers = [
    'LLAVE_CUENTA', 'NUMERO_CUENTA', 'BANCO_ORIGEN', 'ID_SOCIEDAD', 'NOMBRE_SOCIEDAD', 'MONEDA',
    'TIPO_CUENTA', 'TIPO_CUENTA2', 'NOMBRE_CORTO', 'ULTIMOS_DIGITOS', 'ABR_COBRANZA', 'NOTA'
  ];
  var data = raw.map(function (r) {
    return [r[0], r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8], r[9], r[11], r[12]]; // r[10]='X' se descarta (duplicado de NUMERO_CUENTA)
  });
  var sheet = _getOrCreate(newSs, 'CAT_CUENTAS_MAPEO');
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  sheet.setFrozenRows(1);
}

function _migrarHistPosicionBancaria(oldSs, newSs) {
  var old = oldSs.getSheetByName('HIST_POSICION_BANCARIA');
  if (!old || old.getLastRow() < 2) return;
  var headers = [
    'ID_CARGA', 'FECHA_POSICION', 'NOMBRE_CORTO', 'BANCO', 'CUENTA', 'SALDO_ORIGINAL',
    'MONEDA_ORIGINAL', 'TC_APLICADO', 'SALDO_MXN_FINAL', 'TIPO_CUENTA', 'ARCHIVO_FUENTE',
    'TIPO_CUENTA2', 'ABR_COBRANZA' // confirmado: Codigo.gs:660 ya lo llama "abrCobranza"
  ];
  var data = old.getRange(2, 1, old.getLastRow() - 1, headers.length).getValues();
  var sheet = _getOrCreate(newSs, 'HIST_POSICION_BANCARIA');
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  sheet.setFrozenRows(1);
}

function _migrarTC(oldSs, newSs) {
  var old = oldSs.getSheetByName('TC');
  if (!old || old.getLastRow() < 1) return;
  var data = old.getRange(1, 1, old.getLastRow(), 2).getValues(); // incluye la fila 1 (antes se perdia)
  var sheet = _getOrCreate(newSs, 'TC');
  sheet.getRange(1, 1, 1, 2).setValues([['FECHA', 'TC_USD']]);
  sheet.getRange(2, 1, data.length, 2).setValues(data);
  sheet.setFrozenRows(1);
}
