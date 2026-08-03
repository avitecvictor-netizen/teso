/**
 * TESO — Svc_Pagos.js
 * Backend del Workflow de Propuestas de Pago. Puerto de un prototipo
 * standalone entregado por el usuario (_source/workflow_pagos/Code.gs),
 * adaptado a la convencion de nombres de TESO y a BD_TESO como unica base
 * (ver PROMPT_MIGRACION_VEVA seccion 5-bis y PROMPT_TESORERIA360_PAGOS_1
 * para la interpretacion completa).
 *
 * Cambios vs el original (documentados, no silenciosos):
 * - SpreadsheetApp.getActiveSpreadsheet() -> openById(SALDOS_SHEET_ID):
 *   el original asumia un Sheet contenedor; TESO es standalone (mismo
 *   fix ya aplicado en saveAssetRequest, Codigo.js).
 * - Nombres de hoja PROPUESTAS_PAGO / PARTIDAS_PAGO (con sufijo _PAGO
 *   para no chocar con nada existente) en vez de "Propuestas"/"Partidas".
 * - Encabezados en MAYUSCULAS_CON_GUION para igualar el resto de hojas
 *   de BD_TESO, en vez de camelCase.
 * - api_getBoard/api_updateLine/... renombrados sin prefijo "api_".
 * - Las hojas se auto-crean vacias (solo encabezados) la primera vez que
 *   se usan -- NO se siembran datos de ejemplo (son ficticios de demo,
 *   no deben aparecer en produccion).
 *
 * Segunda ronda (Tesoreria 360, ver PROMPT_TESORERIA360_PAGOS_1 secc. 6):
 * - FASE y PROVEEDOR se mueven de PROPUESTAS_PAGO (una por propuesta
 *   completa) a PARTIDAS_PAGO (por linea) -- una propuesta puede tener
 *   varios proveedores, y el mismo proveedor puede repetirse en varias
 *   lineas. setFasePropuestaPago() queda retirado: el avance de fase
 *   ahora es por partida, vía el mismo patch que ya aplican los botones
 *   de accion existentes (autorizar/liberar/subir/aplicar) -- no se creo
 *   un boton nuevo de "avanzar fase" por separado.
 * - Nuevo campo ESTADO ('activa' | 'rechazada' | 'caducada'), separado
 *   de FASE: rechazar/caducar NUNCA pisa FASE (se congela donde iba).
 *   'caducada' se calcula de forma perezosa en cada lectura del tablero
 *   (getPropuestasPago), sin trigger programado: vence el sabado
 *   siguiente a FECHA_CARGA de la propuesta.
 * - STATUS (ya existia: autorizada/liberada) deja de usar el valor
 *   'rechazada' -- ese eje vive ahora solo en ESTADO, para que no
 *   compitan dos campos por el mismo significado.
 * - Nuevo campo MONEDA por partida (dato real de la columna "Moneda"
 *   de la propuesta subida, no inventado aqui).
 * - Nuevo campo NOTA_APLICACION, separado de COMENTARIO: COMENTARIO
 *   sigue siendo "por que sigue bloqueado" (se oculta al terminar,
 *   ver View_Pagos.html); NOTA_APLICACION es "como se resolvio al
 *   aplicar" (tipo de cambio, compra de divisa, etc.), visible solo
 *   cuando la linea ya se marco pagada.
 * - Nueva hoja CAT_USUARIOS (EMAIL/NOMBRE/ROL/ACTIVO) + getUsuarioActual(),
 *   catalogo real de roles (Solicitante/Contador/Tesoreria/Director/
 *   Auditor, matriz de permisos tal cual la definio el usuario) usado
 *   para: (a) decidir si "solicitar reactivacion" de una linea rechazada
 *   es instantaneo (lo pide alguien con rol Tesoreria) o queda pendiente
 *   de confirmar; (b) que el frontend oculte los botones de accion que
 *   no le corresponden al rol del usuario actual. Si el catalogo esta
 *   vacio (aun no lo llena nadie) se falla ABIERTO -- se muestran todos
 *   los botones -- para no bloquear el primer despliegue; si el
 *   catalogo ya tiene filas pero el usuario actual no esta en el,
 *   se falla CERRADO para ese usuario (rol = null, sin botones de accion).
 *
 * Fases en orden de flujo (sin cambios en esta ronda -- conciliacion y
 * comprobantes quedan fuera de este incremento, ver seccion 3 de
 * PROMPT_TESORERIA360_PAGOS_1):
 *   contabilidad -> tesoreria -> liberacion2 -> banca -> aplicado
 */

var SHEET_PROPUESTAS_PAGO = 'PROPUESTAS_PAGO';
var SHEET_PARTIDAS_PAGO = 'PARTIDAS_PAGO';
var SHEET_CAT_USUARIOS = 'CAT_USUARIOS';
var SHEET_CONCILIACION_PAGOS = 'CONCILIACION_PAGOS';
var SHEET_COMPROBANTES_PAGO = 'COMPROBANTES_PAGO';
var SHEET_COMPROBANTES_PAGO_PARTIDAS = 'COMPROBANTES_PAGO_PARTIDAS';
var SHEET_CONFIG_CONCILIACION = 'CONFIG_CONCILIACION';
var SHEET_APLICACIONES_PAGO = 'APLICACIONES_PAGO';
// Carpeta de Drive designada por el usuario (2026-07-26) para comprobantes
// bancarios (PDF/imagen del portal, NO CFDI fiscal) -- no se crea una
// carpeta nueva por codigo, ya existe.
var COMPROBANTES_DRIVE_FOLDER_ID = '10E6ekjx1BWCUTAEUjv6HPIttmHzA6kLn';
var SHEET_CONFIG_NOTIFICACIONES_PAGOS = 'CONFIG_NOTIFICACIONES_PAGOS';

// SOLICITANTE se movio de aqui a PARTIDAS_PAGO (2026-07-26, carga de
// propuesta): en el archivo real, el solicitante varia por linea dentro
// de una misma sociedad/propuesta (hasta 14 personas distintas en un
// mismo lote) -- un solo valor por propuesta perderia ese dato real.
// CONTADOR/TESORERO retirados (2026-07-26, retroalimentacion en vivo):
// eran codigo muerto, nunca se leian ni escribian en ningun lado -- el
// tracking real de quien-hizo-que ahora vive por partida y por fase
// (ver REVISADO/AUTORIZADO/LIBERADO/SUBIDO/APLICADO en PARTIDAS_PAGO).
var PROPUESTAS_PAGO_HEADERS = ['ID', 'FOLIO', 'SOCIEDAD', 'CONCEPTO', 'FECHA_CARGA', 'NOTA_GENERAL'];
// SOLICITANTE, UUID, RFC, NUMERO_DOCUMENTO agregados 2026-07-26 para la
// carga de propuesta real (PROPUESTA.xlsx) -- UUID/RFC/NUMERO_DOCUMENTO
// son el dato fiscal de origen de cada factura, se conservan explicitos
// (no se descartan) por decision del usuario. NUMERO_PROVEEDOR/REFERENCIA
// agregados en la misma ronda que lo de abajo (ver notas ahi).
//
// ULTIMA_ACCION_POR/FECHA (generico, se sobreescribia en cada transicion)
// RETIRADO 2026-07-26 -- hallazgo real: en cuanto una partida pasaba de
// autorizada a liberada, se perdia el rastro de quien la autorizo. Se
// reemplaza por un par POR/FECHA especifico de cada fase real del
// pipeline, que nunca se pisan entre si:
//   REVISADO_POR/FECHA    -- avanzo de contabilidad a tesoreria (Contador)
//   AUTORIZADO_POR/FECHA  -- autorizo en tesoreria (Tesoreria)
//   LIBERADO_POR/FECHA    -- libero en liberacion contable (Contador)
//   SUBIDO_POR/FECHA      -- subio a banca (Tesoreria)
//   APLICADO_POR/FECHA    -- aplico el pago (Tesoreria)
// Se estampan con el usuario real de sesion (getUsuarioActual), nunca con
// SOLICITANTE (que es el dato del Excel, quien PIDIO el pago, no quien
// ejecuto cada paso del workflow -- son cosas distintas, confundirlas fue
// un hallazgo explicito de la retroalimentacion en vivo).
// CUENTA_PAGO/MONEDA_PAGO_REAL/TC_APLICADO agregados 2026-07-26
// (conciliacion + comprobantes, fases 6-7) -- a proposito AL FINAL del
// esquema (append puro), a diferencia de FASE/PROVEEDOR/ESTADO/MONEDA que
// se insertaron EN MEDIO en la ronda anterior. Se capturan en la pestaña
// "Aplicado" (tab==='aplicado' en View_Pagos.html), gateadas a Tesoreria
// -- mismo rol que ya controla el toggle de PAGADO en esa pestaña
// (esTesoreria(), ver canToggleAplicado); Contador NO captura estos
// campos (distinto de conciliar partidas/subir comprobantes en las
// pestañas nuevas, que si son Tesoreria+Contador, ver
// _puedeConciliarOComprobantes):
//   CUENTA_PAGO       -- de que cuenta bancaria (CAT_CUENTAS_MAPEO) se
//                        pago, para poder agrupar Sociedad->Banco en la
//                        pestaña de conciliacion ANTES de que exista un
//                        match real contra HIST_MOVIMIENTOS. OBLIGATORIA
//                        antes de poder marcar la partida como pagada
//                        (2026-07-26, ver validacion en
//                        _updatePartidaPagoRow) -- se captura ANTES del
//                        pago, no despues; el selector de la pestaña
//                        Aplicado esta disponible mientras la partida
//                        sigue sin pagar, justo cuando hace falta.
//   MONEDA_PAGO_REAL   -- solo relevante si difiere de MONEDA (ej. una
//                        partida en EUR pagada en MXN "por acuerdo").
//   TC_APLICADO        -- tipo de cambio usado para esa conversion, para
//                        que el buscador de candidatos de conciliacion
//                        sepa que monto/moneda buscar en el banco.
// Ver PROMPT_TESORERIA360_PAGOS_1.md seccion 3.3 para el detalle completo.
// RECHAZADO_POR/FECHA, REACTIVADO_POR/FECHA, FECHA_REACTIVACION
// (2026-07-27, hallazgo real de auditoria pedido por el usuario) --
// rechazar/reactivar una partida no dejaba NINGUN rastro consultable
// desde la hoja (solo el pie del correo de ese momento). FECHA_REACTIVACION
// ademas resuelve un bug real de negocio: _fechaCaducidad se calculaba
// siempre contra FECHA_CARGA de la PROPUESTA original -- una partida
// rechazada y reactivada DESPUES de esa fecha limite volvia a marcarse
// 'caducada' sola en la siguiente lectura (getPropuestasPago), y
// 'caducada' no tiene ninguna accion disponible en la UI: un callejon sin
// salida real. Con esta columna, getPropuestasPago recalcula la
// caducidad contra FECHA_REACTIVACION cuando existe, en vez de contra
// FECHA_CARGA -- una partida que nunca se rechazo sigue calculando
// exactamente igual que antes (columna vacia, mismo dato de entrada).
var PARTIDAS_PAGO_HEADERS = [
  'ID', 'PROPUESTA_ID', 'PROVEEDOR', 'NUMERO_PROVEEDOR', 'SOLICITANTE', 'DESC', 'REFERENCIA',
  'MONTO', 'MONEDA', 'UUID', 'RFC', 'NUMERO_DOCUMENTO',
  'FASE', 'ESTADO', 'REACTIVACION_SOLICITADA', 'CHECKED', 'APPROVAL_TYPE', 'MONTO_AUTORIZADO', 'STATUS',
  'SUBIDO', 'PAGADO', 'COMENTARIO', 'NOTA_APLICACION',
  'REVISADO_POR', 'REVISADO_FECHA', 'AUTORIZADO_POR', 'AUTORIZADO_FECHA',
  'LIBERADO_POR', 'LIBERADO_FECHA', 'SUBIDO_POR', 'SUBIDO_FECHA', 'APLICADO_POR', 'APLICADO_FECHA',
  'CUENTA_PAGO', 'MONEDA_PAGO_REAL', 'TC_APLICADO',
  'RECHAZADO_POR', 'RECHAZADO_FECHA', 'REACTIVADO_POR', 'REACTIVADO_FECHA', 'FECHA_REACTIVACION'
];
var CAT_USUARIOS_HEADERS = ['EMAIL', 'NOMBRE', 'ROL', 'ACTIVO'];

// CONCILIACION_PAGOS: vinculo PARTIDA<->MOVIMIENTO (N:M real -- una fila
// por match). Split (1 partida pagada en varias exhibiciones) = varias
// filas con la misma PARTIDA_ID; distintos ID_MOV. MONTO_APLICADO es el
// monto de ESE movimiento especifico que se atribuye a la partida (para
// poder sumar y validar tolerancia < $1 contra el monto efectivo).
var CONCILIACION_PAGOS_HEADERS = ['ID', 'PARTIDA_ID', 'ID_MOV', 'MONTO_APLICADO', 'CONCILIADO_POR', 'CONCILIADO_FECHA'];

// COMPROBANTES_PAGO (el archivo) + COMPROBANTES_PAGO_PARTIDAS (el vinculo,
// hoja aparte a proposito): un PDF consolidado del banco (resumen + varios
// pagos individuales adentro) se sube UNA vez y se vincula a N partidas.
// Desvincular una partida es borrar su fila en la hoja de vinculo -- nunca
// toca COMPROBANTES_PAGO ni el archivo real en Drive. Identificar el
// monto/pagina exacto de cada pago DENTRO del PDF consolidado queda fuera
// de este incremento (iteracion futura, ver PROMPT_TESORERIA360_PAGOS_1).
var COMPROBANTES_PAGO_HEADERS = ['ID', 'ARCHIVO_NOMBRE', 'DRIVE_FILE_ID', 'DRIVE_URL', 'MIME_TYPE', 'SUBIDO_POR', 'SUBIDO_FECHA'];
var COMPROBANTES_PAGO_PARTIDAS_HEADERS = ['ID', 'COMPROBANTE_ID', 'PARTIDA_ID', 'VINCULADO_POR', 'VINCULADO_FECHA'];

// APLICACIONES_PAGO (2026-07-27, Bloque B): una fila por cada liberacion
// PARCIAL de pago sobre una partida -- nunca un campo acumulado sin
// historial en PARTIDAS_PAGO, para conservar auditoria real de quien/
// cuando de CADA exhibicion (decision explicita del usuario, mismo
// patron ya establecido por CONCILIACION_PAGOS para sus propios splits).
// PAGADO/APLICADO_POR/APLICADO_FECHA en PARTIDAS_PAGO siguen siendo el
// estado FINAL derivado (se estampan solo cuando la suma de esta hoja
// alcanza el monto efectivo completo, ver aplicarPagoParcial) -- esta
// hoja es la fuente de verdad del detalle, PARTIDAS_PAGO es el resumen.
var APLICACIONES_PAGO_HEADERS = ['ID', 'PARTIDA_ID', 'MONTO', 'APLICADO_POR', 'APLICADO_FECHA'];

// CONFIG_CONCILIACION: hoja tipo llave/valor editable a mano por finanzas
// (mismo patron ya establecido por CONFIG_DISTRIBUCION en Codigo.js) --
// controla la tolerancia de BUSQUEDA de candidatos a conciliar (la
// tolerancia de CONFIRMAR un match es fija, <$1 exacto, no configurable).
// Pedido explicito del usuario 2026-07-26: que sea editable facil por
// admin sin tocar codigo.
var CONFIG_CONCILIACION_HEADERS = ['CLAVE', 'VALOR', 'DESCRIPCION'];
var CONFIG_CONCILIACION_DEFAULTS = [
  ['TOLERANCIA_MONTO_PCT', 2, 'Porcentaje de tolerancia para BUSCAR movimientos candidatos (no aplica a confirmar el match, que exige <$1 de diferencia exacta).'],
  ['TOLERANCIA_DIAS', 2, 'Dias +/- alrededor de APLICADO_FECHA de la partida para buscar movimientos candidatos.']
];

// CONFIG_NOTIFICACIONES_PAGOS: lista de distribucion de correo POR
// SOCIEDAD (2026-07-26, Bloque 1) -- una fila por destinatario, nunca una
// lista-CSV en una sola celda, para poder dar de alta/baja individual y
// llevar FECHA_ALTA por persona. Vacia al crearse, nunca se siembra un
// correo inventado (mismo criterio que el resto de hojas CAT_ y CONFIG_
// de TESO). Deliberadamente NO reusa CONFIG_DISTRIBUCION (Codigo.js) --
// esa es una lista unica global sin columna de sociedad, de otro dominio
// (reporte de posicion, disparo manual por boton), no cabe el requisito
// real de este bloque.
var CONFIG_NOTIFICACIONES_PAGOS_HEADERS = ['SOCIEDAD', 'EMAIL', 'NOMBRE', 'ACTIVO', 'FECHA_ALTA'];

// CONFIG_CC_PAGOS (2026-07-27, Bloque E): lista de copia (CC) POR
// SOCIEDAD, permanente, aplicada a los 7 tipos de correo de pagos sin
// excepcion -- mismo shape/patron que CONFIG_NOTIFICACIONES_PAGOS
// (una fila por destinatario), pero es un grupo DISTINTO: esa hoja es el
// destinatario principal (TO), esta es la copia (CC). Vive dentro de la
// pantalla "Lista de Distribucion" (decision explicita del usuario --
// a diferencia de Notificaciones de Pagos, que se hizo aparte).
var SHEET_CONFIG_CC_PAGOS = 'CONFIG_CC_PAGOS';
var CONFIG_CC_PAGOS_HEADERS = ['SOCIEDAD', 'EMAIL', 'NOMBRE', 'ACTIVO', 'FECHA_ALTA'];

// CC_EXTRA_PROPUESTA (2026-07-27, Bloque E): copia EXTRAORDINARIA por
// PROPUESTA puntual (no por sociedad) -- capturada al cargar el Excel,
// vía buscador del directorio de Google Workspace. Hoja de auditoria
// dedicada (mismo patron que APLICACIONES_PAGO), tipicamente 0-1 fila
// por propuesta pero sin bloquear el caso raro de mas de 1. Sin logica
// de expiracion: una propuesta en fase final deja de generar
// transiciones, asi que deja de consultarse en la practica -- no hace
// falta borrar nada a mano (decision explicita del usuario).
var SHEET_CC_EXTRA_PROPUESTA = 'CC_EXTRA_PROPUESTA';
var CC_EXTRA_PROPUESTA_HEADERS = ['ID', 'PROPUESTA_ID', 'EMAIL', 'NOMBRE', 'AGREGADO_POR', 'AGREGADO_FECHA'];

// Pieza 2 (2026-07-28, pedido explicito del usuario): hoja de auditoria
// dedicada para cada vez que alguien fuerza la carga de una linea con
// UUID ya existente en PARTIDAS_PAGO -- mismo patron "hoja de evento,
// auto-creada vacia" que CC_EXTRA_PROPUESTA/APLICACIONES_PAGO, nunca se
// pisa ni se borra, un registro por linea forzada.
var SHEET_OVERRIDE_DUPLICADOS_PAGO = 'OVERRIDE_DUPLICADOS_PAGO';
var OVERRIDE_DUPLICADOS_PAGO_HEADERS = ['ID', 'PROPUESTA_ID', 'PARTIDA_ID', 'UUID', 'PROVEEDOR', 'SOCIEDAD', 'MOTIVO', 'FORZADO_POR', 'FORZADO_FECHA'];

/** Dado un patch de mutacion, decide que par POR/FECHA de auditoria por
 * fase le corresponde (si alguno) y devuelve los campos extra a
 * persistir junto con el patch original. Cada transicion real del
 * pipeline tiene una firma de patch unica e inconfundible (ver
 * bulkAction/autorizarSeleccionadas/toggleBool en View_Pagos.html):
 * fase:'tesoreria' = revision contable completada; status:'autorizada' =
 * tesoreria autorizo; status:'liberada' = contabilidad libero;
 * subido:true = tesoreria subio a banca; pagado:true = tesoreria aplico;
 * estado:'rechazada' = tesoreria rechazo (2026-07-27, hallazgo de
 * auditoria); reactivacion instantanea (estado:'activa' +
 * reactivacionSolicitada:false, firma exclusiva de esa rama en
 * solicitarReactivacion -- ninguna otra funcion del archivo escribe ese
 * par exacto) = tesoreria reactivo, y de paso se estampa
 * fechaReactivacion (usada por getPropuestasPago para no volver a
 * calcular 'caducada' contra la fecha de carga original). */
function _camposAuditoriaPorFase(patch, email, ahora) {
  var extra = {};
  if (patch.fase === 'tesoreria') { extra.revisadoPor = email; extra.revisadoFecha = ahora; }
  if (patch.status === 'autorizada') { extra.autorizadoPor = email; extra.autorizadoFecha = ahora; }
  if (patch.status === 'liberada') { extra.liberadoPor = email; extra.liberadoFecha = ahora; }
  if (patch.subido === true) { extra.subidoPor = email; extra.subidoFecha = ahora; }
  if (patch.pagado === true) { extra.aplicadoPor = email; extra.aplicadoFecha = ahora; }
  if (patch.estado === 'rechazada') { extra.rechazadoPor = email; extra.rechazadoFecha = ahora; }
  if (patch.estado === 'activa' && patch.reactivacionSolicitada === false) {
    extra.reactivadoPor = email; extra.reactivadoFecha = ahora; extra.fechaReactivacion = ahora;
  }
  return extra;
}

/** Detecta si un patch corresponde a una de las 6 transiciones
 * notificables por correo (Bloque 1, 2026-07-26) -- mismas firmas de
 * patch que ya usa _camposAuditoriaPorFase para auditoria (deliberado,
 * para que no se desalineen: si se agrega una fase nueva hay que tocar
 * ambas funciones), mas el caso de rechazo, que esa funcion no cubre
 * porque no tiene par POR/FECHA propio. Devuelve null si el patch no es
 * ninguna transicion notificable (comentario suelto, checkbox, monto
 * autorizado sin mas, etc.) -- nunca manda correo por esos casos.
 *
 * `etiqueta` (2026-07-27, mapeo confirmado con el usuario) es el texto en
 * mayusculas que va al final del asunto del correo (PROPUESTA DE PAGOS
 * "SOCIEDAD" FECHA "DD/MM/AAAA" <etiqueta>). Vive en la misma rama que ya
 * define `verbo` para que nunca se desalinee de la transicion real --
 * nunca una tabla de mapeo aparte.
 *
 * La etiqueta CARGADA (fase 1, carga inicial de la propuesta) tiene su
 * propia constante -- `TRANSICION_CARGADA` -- porque ese evento NO es un
 * patch sobre una partida existente (es la creacion de PARTIDAS_PAGO en
 * confirmarCargaPropuestaPago), esta funcion no le aplica. Bloque D
 * (2026-07-27): ya se dispara, ver confirmarCargaPropuestaPago. */
function _tipoTransicionNotificable(patch) {
  if (patch.fase === 'tesoreria') return { tipo: 'revisado', verbo: 'revisada a Tesorería', etiqueta: 'CONFIRMADA' };
  if (patch.status === 'autorizada') return { tipo: 'autorizado', verbo: 'autorizada', etiqueta: 'AUTORIZADA' };
  if (patch.estado === 'rechazada') return { tipo: 'rechazado', verbo: 'rechazada', etiqueta: 'RECHAZADA' };
  if (patch.status === 'liberada') return { tipo: 'liberado', verbo: 'liberada a Banco', etiqueta: 'LIBERADA A BANCO' };
  if (patch.subido === true) return { tipo: 'subido', verbo: 'subida al banco', etiqueta: 'SUBIDA EN BANCA' };
  if (patch.pagado === true) return { tipo: 'aplicado', verbo: 'pagada', etiqueta: 'APLICADA EN BANCA' };
  return null;
}

/** Transicion de la carga inicial de una propuesta (Bloque D,
 * 2026-07-27) -- mismo shape {tipo,verbo,etiqueta} que devuelve
 * _tipoTransicionNotificable, pero como constante literal porque no hay
 * ningun patch de PARTIDAS_PAGO que la represente (es la creacion misma
 * de las filas, no una mutacion). Usada solo por confirmarCargaPropuestaPago. */
var TRANSICION_CARGADA = { tipo: 'cargado', verbo: 'cargada', etiqueta: 'CARGADA' };

/** Compara el encabezado real de una hoja ya existente contra el esperado.
 * PROVEEDOR/FASE/ESTADO/MONEDA se insertaron EN MEDIO del orden de
 * columnas en esta ronda (no solo al final), asi que una hoja creada por
 * una version anterior del esquema NO se puede seguir leyendo
 * posicionalmente sin corromper datos en silencio (col. 4 pasaria a
 * leerse como otra cosa). Mejor fallar con un mensaje claro que adivinar
 * una migracion automatica. */
function _verificarEncabezado(sh, headersEsperados, nombreHoja) {
  var actual = sh.getRange(1, 1, 1, headersEsperados.length).getValues()[0];
  var coincide = headersEsperados.every(function (h, i) { return actual[i] === h; });
  if (!coincide) {
    throw new Error(nombreHoja + ' tiene un encabezado distinto al que espera esta version del codigo (el esquema de Tesoreria 360 movio columnas, no solo agrego al final). Verifica/migra la hoja a mano antes de continuar: esperado [' + headersEsperados.join(', ') + '].');
  }
}

function _ensurePropuestasPagoSheet(ss) {
  var sh = ss.getSheetByName(SHEET_PROPUESTAS_PAGO);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PROPUESTAS_PAGO);
    sh.getRange(1, 1, 1, PROPUESTAS_PAGO_HEADERS.length).setValues([PROPUESTAS_PAGO_HEADERS]);
    sh.setFrozenRows(1);
    return sh;
  }
  _verificarEncabezado(sh, PROPUESTAS_PAGO_HEADERS, SHEET_PROPUESTAS_PAGO);
  return sh;
}

/** Migracion ADITIVA (2026-07-26): CUENTA_PAGO/MONEDA_PAGO_REAL/
 * TC_APLICADO se agregaron al FINAL de PARTIDAS_PAGO_HEADERS. A
 * diferencia de la migracion anterior (columnas insertadas EN MEDIO, que
 * por eso _verificarEncabezado exige migrar la hoja a mano), un append
 * puro se puede completar solo -- mismo patron que ya usa
 * _ensureHistMovSheet (Codigo.js) para MONEDA_ORIGINAL/TC_APLICADO de
 * HIST_MOVIMIENTOS. Solo escribe encabezados en columnas que hoy no
 * existen; nunca toca columnas ya presentes. _verificarEncabezado sigue
 * bloqueando duro para cualquier OTRO mismatch real (typo, columna
 * movida) -- esto solo la vuelve tolerante a "faltan columnas nuevas al
 * final". */
function _migrarColumnasFinalesPartidasPago(sh) {
  // Bug real encontrado en produccion (2026-07-26): NO usar getLastColumn()
  // para decidir que columnas "ya existen" -- Apps Script puede reportar
  // un lastColumn mayor al real si alguna celda mas alla del contenido
  // real trae solo formato (borde, color, formato de numero) sin valor,
  // cosa muy comun en hojas que finanzas ha estado formateando a mano.
  // Eso hacia que la migracion se saltara CUENTA_PAGO/MONEDA_PAGO_REAL/
  // TC_APLICADO creyendo que ya existian, dejandolas en blanco y
  // rompiendo _verificarEncabezado. Fix: leer el VALOR real de cada
  // celda esperada y solo llenar las que esten vacias, sin importar
  // cuantas columnas reporte getLastColumn().
  var totalCols = Math.max(sh.getLastColumn(), PARTIDAS_PAGO_HEADERS.length);
  var actual = sh.getRange(1, 1, 1, totalCols).getValues()[0];
  PARTIDAS_PAGO_HEADERS.forEach(function (h, i) {
    if (!actual[i]) sh.getRange(1, i + 1).setValue(h);
  });
}

/** Migracion puntual (2026-07-26): la hoja PARTIDAS_PAGO real en
 * produccion se encontro con el esquema de ANTES de la "tercera ronda"
 * (retroalimentacion en vivo, ver historial arriba) -- le faltaban
 * NUMERO_PROVEEDOR/REFERENCIA (se insertaron EN MEDIO en esa ronda) y
 * todavia traia el par generico ULTIMA_ACCION_POR/ULTIMA_ACCION_FECHA
 * que esa misma ronda retiro a favor de los 5 pares especificos por
 * fase. No se sabe si la migracion de esa ronda nunca se completo en la
 * hoja real o se revirtio por accidente -- de cualquier forma no se
 * puede corregir borrando y recreando la hoja (la unica migracion "en
 * medio" usada hasta ahora) porque ya hay partidas reales cargadas.
 * Se detecta por firma EXACTA de las primeras 23 columnas (para no
 * disparar sobre ningun otro esquema inesperado -- si la firma no
 * coincide, no hace nada y deja que _verificarEncabezado atrape
 * cualquier otro problema con su mensaje de siempre) y se corrige con
 * insertColumnBefore/deleteColumns, que reacomodan los DATOS de cada
 * fila junto con el encabezado -- no solo la fila 1. Perdida de datos
 * conocida y aceptada: el contenido de ULTIMA_ACCION_POR/FECHA se borra
 * (ese campo ya estaba retirado del codigo desde la ronda anterior, nada
 * lo lee ni lo escribe). */
function _migrarEsquemaMedioPartidasPago(sh) {
  var ESQUEMA_ANTES_TERCERA_RONDA = [
    'ID', 'PROPUESTA_ID', 'PROVEEDOR', 'SOLICITANTE', 'DESC',
    'MONTO', 'MONEDA', 'UUID', 'RFC', 'NUMERO_DOCUMENTO',
    'FASE', 'ESTADO', 'REACTIVACION_SOLICITADA', 'CHECKED', 'APPROVAL_TYPE', 'MONTO_AUTORIZADO', 'STATUS',
    'SUBIDO', 'PAGADO', 'COMENTARIO', 'NOTA_APLICACION',
    'ULTIMA_ACCION_POR', 'ULTIMA_ACCION_FECHA'
  ];
  var actual = sh.getRange(1, 1, 1, ESQUEMA_ANTES_TERCERA_RONDA.length).getValues()[0];
  var esEsquemaViejo = ESQUEMA_ANTES_TERCERA_RONDA.every(function (h, i) { return actual[i] === h; });
  if (!esEsquemaViejo) return;

  sh.insertColumnBefore(4);
  sh.getRange(1, 4).setValue('NUMERO_PROVEEDOR');
  sh.insertColumnBefore(7);
  sh.getRange(1, 7).setValue('REFERENCIA');
  sh.deleteColumns(24, 2);
}

function _ensurePartidasPagoSheet(ss) {
  var sh = ss.getSheetByName(SHEET_PARTIDAS_PAGO);
  if (!sh) {
    sh = ss.insertSheet(SHEET_PARTIDAS_PAGO);
    sh.getRange(1, 1, 1, PARTIDAS_PAGO_HEADERS.length).setValues([PARTIDAS_PAGO_HEADERS]);
    sh.setFrozenRows(1);
    return sh;
  }
  _migrarEsquemaMedioPartidasPago(sh);
  _migrarColumnasFinalesPartidasPago(sh);
  _verificarEncabezado(sh, PARTIDAS_PAGO_HEADERS, SHEET_PARTIDAS_PAGO);
  return sh;
}

function _ensureConciliacionPagosSheet(ss) {
  var sh = ss.getSheetByName(SHEET_CONCILIACION_PAGOS);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_CONCILIACION_PAGOS);
  sh.getRange(1, 1, 1, CONCILIACION_PAGOS_HEADERS.length).setValues([CONCILIACION_PAGOS_HEADERS]);
  sh.setFrozenRows(1);
  return sh;
}

/** Hoja nueva (Bloque B, 2026-07-27), sin datos previos que migrar --
 * mismo patron minimo que _ensureConciliacionPagosSheet, no requiere
 * _verificarEncabezado porque nunca hubo un esquema anterior distinto. */
function _ensureAplicacionesPagoSheet(ss) {
  var sh = ss.getSheetByName(SHEET_APLICACIONES_PAGO);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_APLICACIONES_PAGO);
  sh.getRange(1, 1, 1, APLICACIONES_PAGO_HEADERS.length).setValues([APLICACIONES_PAGO_HEADERS]);
  sh.setFrozenRows(1);
  return sh;
}

/** Todas las filas de APLICACIONES_PAGO de una partida, ordenadas tal
 * cual vienen de la hoja (orden de captura). Helper compartido por
 * getConciliacionPagos/buscarCandidatosMovimiento/conciliarPartida/
 * aplicarPagoParcial -- una sola forma de leer/agrupar, nunca duplicada. */
function _aplicacionesPorPartida(ss, partidaId) {
  var todas = _pagoSheetToObjects(_ensureAplicacionesPagoSheet(ss), APLICACIONES_PAGO_HEADERS);
  return todas.filter(function (a) { return a.partidaId === partidaId; });
}

/** Mapa {partidaId: montoAcumulado} de TODAS las partidas de un jalon,
 * en una sola lectura de APLICACIONES_PAGO -- hallazgo real del revisor:
 * getPropuestasPago y getConciliacionPagos repetian este mismo forEach
 * de agrupado idéntico, mismo criterio anti-duplicacion que ya motivo
 * extraer _montoEfectivoPartida. */
function _acumuladoAplicadoPorPartida(ss) {
  var aplicRows = _pagoSheetToObjects(_ensureAplicacionesPagoSheet(ss), APLICACIONES_PAGO_HEADERS);
  var acumulado = {};
  aplicRows.forEach(function (a) { acumulado[a.partidaId] = (acumulado[a.partidaId] || 0) + (Number(a.monto) || 0); });
  return acumulado;
}

/** Monto efectivo TOTAL que hay que pagar por una partida (el objetivo
 * final, no lo ya aplicado) -- respeta autorizacion parcial
 * (approvalType/montoAutorizado) y conversion de moneda de pago real
 * (monedaPagoReal/tcAplicado). Hallazgo real de planificacion (2026-07-27):
 * getConciliacionPagos/buscarCandidatosMovimiento/conciliarPartida
 * calculaban esto 3 veces de forma inconsistente, SIEMPRE contra
 * Number(l.monto) (el nominal de la factura), sin considerar
 * approvalType==='parcial' -- a diferencia de montoEfectivo(l) del
 * frontend (View_Pagos.html), que si lo hace. Bug real preexistente,
 * corregido aqui de una vez con un unico helper reusado en los 3 sitios
 * mas la funcion nueva de aplicar pago parcial. */
function _montoEfectivoPartida(l) {
  var base = (l.approvalType === 'parcial') ? (Number(l.montoAutorizado) || 0) : Number(l.monto);
  return (l.monedaPagoReal && l.monedaPagoReal !== l.moneda && l.tcAplicado) ? base * Number(l.tcAplicado) : base;
}

function _ensureComprobantesPagoSheets(ss) {
  var shComp = ss.getSheetByName(SHEET_COMPROBANTES_PAGO);
  if (!shComp) {
    shComp = ss.insertSheet(SHEET_COMPROBANTES_PAGO);
    shComp.getRange(1, 1, 1, COMPROBANTES_PAGO_HEADERS.length).setValues([COMPROBANTES_PAGO_HEADERS]);
    shComp.setFrozenRows(1);
  }
  var shVinc = ss.getSheetByName(SHEET_COMPROBANTES_PAGO_PARTIDAS);
  if (!shVinc) {
    shVinc = ss.insertSheet(SHEET_COMPROBANTES_PAGO_PARTIDAS);
    shVinc.getRange(1, 1, 1, COMPROBANTES_PAGO_PARTIDAS_HEADERS.length).setValues([COMPROBANTES_PAGO_PARTIDAS_HEADERS]);
    shVinc.setFrozenRows(1);
  }
  return { comprobantes: shComp, vinculos: shVinc };
}

function _ensureConfigConciliacionSheet(ss) {
  var sh = ss.getSheetByName(SHEET_CONFIG_CONCILIACION);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_CONFIG_CONCILIACION);
  sh.getRange(1, 1, 1, CONFIG_CONCILIACION_HEADERS.length).setValues([CONFIG_CONCILIACION_HEADERS]);
  sh.getRange(2, 1, CONFIG_CONCILIACION_DEFAULTS.length, CONFIG_CONCILIACION_HEADERS.length).setValues(CONFIG_CONCILIACION_DEFAULTS);
  sh.setFrozenRows(1);
  return sh;
}

/** Lee CONFIG_CONCILIACION con defaults de respaldo (2%/2 dias) si la
 * hoja esta vacia o finanzas borro una fila por error -- nunca truena por
 * config faltante. Editable a mano por finanzas, como el resto de hojas
 * CAT_ y CONFIG_ de TESO, sin pantalla de admin dedicada. */
function _leerConfigConciliacion(ss) {
  var sh = _ensureConfigConciliacionSheet(ss);
  var defaults = { TOLERANCIA_MONTO_PCT: 2, TOLERANCIA_DIAS: 2 };
  if (sh.getLastRow() < 2) return { toleranciaMontoPct: defaults.TOLERANCIA_MONTO_PCT, toleranciaDias: defaults.TOLERANCIA_DIAS };
  var vals = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  var map = {};
  vals.forEach(function (r) { if (r[0]) map[String(r[0]).trim()] = r[1]; });
  var pct = Number(map.TOLERANCIA_MONTO_PCT);
  var dias = Number(map.TOLERANCIA_DIAS);
  return {
    toleranciaMontoPct: (!isNaN(pct) && pct >= 0) ? pct : defaults.TOLERANCIA_MONTO_PCT,
    toleranciaDias: (!isNaN(dias) && dias >= 0) ? dias : defaults.TOLERANCIA_DIAS
  };
}

function _ensureConfigNotificacionesPagoSheet(ss) {
  var sh = ss.getSheetByName(SHEET_CONFIG_NOTIFICACIONES_PAGOS);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_CONFIG_NOTIFICACIONES_PAGOS);
  sh.getRange(1, 1, 1, CONFIG_NOTIFICACIONES_PAGOS_HEADERS.length).setValues([CONFIG_NOTIFICACIONES_PAGOS_HEADERS]);
  sh.setFrozenRows(1);
  return sh;
}

function _ensureConfigCcPagosSheet(ss) {
  var sh = ss.getSheetByName(SHEET_CONFIG_CC_PAGOS);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_CONFIG_CC_PAGOS);
  sh.getRange(1, 1, 1, CONFIG_CC_PAGOS_HEADERS.length).setValues([CONFIG_CC_PAGOS_HEADERS]);
  sh.setFrozenRows(1);
  return sh;
}

function _ensureCcExtraPropuestaSheet(ss) {
  var sh = ss.getSheetByName(SHEET_CC_EXTRA_PROPUESTA);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_CC_EXTRA_PROPUESTA);
  sh.getRange(1, 1, 1, CC_EXTRA_PROPUESTA_HEADERS.length).setValues([CC_EXTRA_PROPUESTA_HEADERS]);
  sh.setFrozenRows(1);
  return sh;
}

function _ensureOverrideDuplicadosPagoSheet(ss) {
  var sh = ss.getSheetByName(SHEET_OVERRIDE_DUPLICADOS_PAGO);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_OVERRIDE_DUPLICADOS_PAGO);
  sh.getRange(1, 1, 1, OVERRIDE_DUPLICADOS_PAGO_HEADERS.length).setValues([OVERRIDE_DUPLICADOS_PAGO_HEADERS]);
  sh.setFrozenRows(1);
  return sh;
}

/** Fila de auditoria para una linea con UUID ya existente, cargada de
 * todas formas por decision explicita del usuario (Pieza 2, 2026-07-28).
 * Extraida de confirmarCargaPropuestaPago (hallazgo del revisor: esa
 * funcion ya excedia por mucho la convencion de funciones cortas del
 * proyecto). */
function _filaOverrideDuplicado(propId, lineaId, l, sociedad, motivoForzado, emailQuienCarga, ahora) {
  return OVERRIDE_DUPLICADOS_PAGO_HEADERS.map(function (h) {
    if (h === 'ID') return Utilities.getUuid();
    if (h === 'PROPUESTA_ID') return propId;
    if (h === 'PARTIDA_ID') return lineaId;
    if (h === 'UUID') return l.uuid;
    if (h === 'PROVEEDOR') return l.proveedor || '';
    if (h === 'SOCIEDAD') return sociedad;
    if (h === 'MOTIVO') return motivoForzado;
    if (h === 'FORZADO_POR') return emailQuienCarga;
    if (h === 'FORZADO_FECHA') return ahora;
    return '';
  });
}

/** Mapa {propuestaId: [emails]} de TODAS las propuestas con CC
 * extraordinario, en una sola lectura de CC_EXTRA_PROPUESTA -- mismo
 * criterio de rendimiento que _acumuladoAplicadoPorPartida (una lectura
 * para todo el lote, no una consulta por propuesta dentro de un forEach).
 * Usado por _notificarTransicionPagos para sub-agrupar el correo de una
 * sociedad por propuesta cuando alguna tiene CC extra y otra no. */
function _ccExtraPorPropuestaMap(ss) {
  var todas = _pagoSheetToObjects(_ensureCcExtraPropuestaSheet(ss), CC_EXTRA_PROPUESTA_HEADERS);
  var map = {};
  todas.forEach(function (c) {
    if (!map[c.propuestaId]) map[c.propuestaId] = [];
    map[c.propuestaId].push(c.email);
  });
  return map;
}

/** Roles esperados (matriz de permisos definida por el usuario). Se usa
 * tanto para la validacion por lista en la hoja como para las
 * comparaciones de rol -- una sola fuente de verdad para el texto exacto,
 * en vez de escribirlo suelto en cada sitio (hallazgo de QA: el
 * comentario original decia "Tesoreria" sin acento mientras el codigo
 * comparaba contra "Tesorería" con acento -- si finanzas tecleaba el rol
 * sin acento en la hoja, la comparacion fallaba en silencio). Ademas,
 * toda comparacion de rol pasa por _normalizarRol (sin acentos, sin
 * mayusculas, sin espacios de sobra) para que un typo de captura no
 * rompa el gateo de permisos. */
var ROLES_VALIDOS = ['Solicitante', 'Contador', 'Tesorería', 'Director', 'Auditor'];
var ROL_TESORERIA = 'Tesorería';
var ROL_CONTADOR = 'Contador';

function _normalizarRol(s) {
  // Rango Unicode escrito con codigos hex (backslash-u), no como
  // caracteres sueltos: este archivo se copia y pega al editor de Apps
  // Script, y un caracter combinante literal en el codigo fuente es
  // fragil ante encoding/copy-paste. El escape hex es puro ASCII.
  var SIN_ACENTOS = new RegExp(String.fromCharCode(91) + String.fromCharCode(0x0300) + '-' + String.fromCharCode(0x036f) + String.fromCharCode(93), 'g');
  return String(s || '').trim().toLowerCase().normalize('NFD').replace(SIN_ACENTOS, '');
}
function _esRol(rolActual, rolEsperado) {
  return _normalizarRol(rolActual) === _normalizarRol(rolEsperado);
}

/** Bug real encontrado 2026-07-26 (bloque de acceso): ACTIVO en
 * CAT_USUARIOS se llena a mano como texto libre (confirmado con el
 * usuario), y la comparacion "truthy" de JS que se usaba antes
 * (`&& u.activo`) trataba CUALQUIER texto no vacio como activo --
 * incluida la palabra "No" escrita literal. Mientras ACTIVO solo
 * ocultaba botones no importaba tanto; ahora que se usa para bloquear
 * acceso de verdad, se normaliza explicitamente. Acepta boolean TRUE de
 * un checkbox real de Sheets tambien, no solo texto. */
var VALORES_ACTIVO = ['si', 'true', '1', 'activo', 'activa'];
function _esActivo(valor) {
  if (valor === true) return true;
  if (valor === false || valor == null || valor === '') return false;
  return VALORES_ACTIVO.indexOf(_normalizarRol(valor)) >= 0;
}

/** Conciliar partidas y subir/vincular comprobantes: permitido a
 * Tesoreria Y Contador (decidido con el usuario 2026-07-26 -- en la
 * practica Contabilidad tambien recibe comprobantes del banco). Mismo
 * criterio "falla ABIERTO si CAT_USUARIOS esta vacio" que ya usa el resto
 * del modulo (ver getUsuarioActual) para no bloquear el primer
 * despliegue. */
function _puedeConciliarOComprobantes(ss) {
  var usuarios = _pagoSheetToObjects(_ensureCatUsuariosSheet(ss), CAT_USUARIOS_HEADERS);
  if (usuarios.length === 0) return true;
  return _tieneAlgunRol(ss, [ROL_TESORERIA, ROL_CONTADOR]);
}

/** Acceso real al modulo de pagos -- no solo "oculta botones", bloquea
 * lectura Y escritura (2026-07-26, pedido explicito del usuario:
 * "bloquear a quien no este en CAT_USUARIOS"). Mismo criterio "falla
 * ABIERTO si el catalogo esta vacio" que ya usa el resto del modulo
 * (para no bloquear el primer despliegue) -- eso NO cambia. Lo que
 * cambia es el caso "catalogo con filas pero el usuario no aparece o su
 * fila no esta activa": antes solo ocultaba botones en el frontend,
 * ahora bloquea de verdad, en cliente y servidor. */
function _tieneAccesoValido(ss) {
  var usuarios = _pagoSheetToObjects(_ensureCatUsuariosSheet(ss), CAT_USUARIOS_HEADERS);
  if (usuarios.length === 0) return true;
  return _rolesUsuarioActual(ss).length > 0;
}

/** Catalogo real de usuarios/roles. Se llena a mano por finanzas (una fila
 * por persona) -- igual que el resto de hojas CAT_* de TESO, sin tocar
 * codigo. La columna ROL lleva validacion de lista (ROLES_VALIDOS) para
 * reducir el riesgo de typos, ademas de la normalizacion de arriba. */
function _ensureCatUsuariosSheet(ss) {
  var sh = ss.getSheetByName(SHEET_CAT_USUARIOS);
  if (sh) return sh;
  sh = ss.insertSheet(SHEET_CAT_USUARIOS);
  sh.getRange(1, 1, 1, CAT_USUARIOS_HEADERS.length).setValues([CAT_USUARIOS_HEADERS]);
  sh.setFrozenRows(1);
  var rolCol = CAT_USUARIOS_HEADERS.indexOf('ROL') + 1;
  var regla = SpreadsheetApp.newDataValidation().requireValueInList(ROLES_VALIDOS, true).setAllowInvalid(false).build();
  sh.getRange(2, rolCol, 500, 1).setDataValidation(regla);
  return sh;
}

/** BUG REAL corregido 2026-07-26: con hojas vacias/de prueba
 * getPropuestasPago siempre funciono; en cuanto hubo datos reales
 * (113 partidas importadas de PROPUESTA.xlsx, cada una con
 * ULTIMA_ACCION_FECHA, mas FECHA_CARGA por propuesta) el cliente
 * empezo a recibir `null` en vez del objeto real -- confirmado con
 * `Logger.log`: el servidor SI devolvia los datos completos
 * (ejecucion "Completada"), pero google.script.run resolvia null del
 * lado del navegador. Causa: `sheet.getRange(...).getValues()`
 * devuelve objetos Date nativos de Apps Script para columnas con
 * fecha, y google.script.run transporta mal las respuestas grandes
 * con muchos Date anidados (confirmado que solo fallaba con volumen
 * real de datos, nunca con las hojas vacias). Fix: convertir todo
 * Date a string ISO aqui mismo, una sola vez, para que cualquier
 * funcion que use este helper quede protegida -- `new Date(...)` en
 * el propio servidor (ej. _fechaCaducidad) sigue funcionando igual
 * con un string ISO como entrada. */
function _pagoSheetToObjects(sheet, headers) {
  if (sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  return values.filter(function (row) { return row[0] !== ''; }).map(function (row) {
    var obj = {};
    headers.forEach(function (h, i) {
      var key = h.toLowerCase().replace(/_([a-z])/g, function (_, c) { return c.toUpperCase(); });
      var val = row[i];
      if (val instanceof Date) val = val.toISOString();
      obj[key] = val;
    });
    return obj;
  });
}

/** Todos los roles ACTIVOS del usuario de la sesion actual -- array,
 * nunca un solo rol (hallazgo real 2026-07-26: una persona puede tener
 * MAS DE UNA fila en CAT_USUARIOS con el mismo email, una por rol --
 * ej. alguien que es Tesoreria Y Director a la vez, exactamente el caso
 * real del usuario). La version anterior de esta funcion tomaba solo
 * "la primera fila que coincidiera" (`.filter(...)[0]`), asi que una
 * persona con varias filas quedaba atrapada con el PRIMER rol de la
 * hoja segun el orden de captura -- casi siempre no el que necesitaba
 * en ese momento. Array vacio si el email no aparece en ninguna fila
 * activa. */
function _rolesUsuarioActual(ss) {
  var email = Session.getActiveUser().getEmail();
  var usuarios = _pagoSheetToObjects(_ensureCatUsuariosSheet(ss), CAT_USUARIOS_HEADERS);
  return usuarios.filter(function (u) { return u.email === email && _esActivo(u.activo); }).map(function (u) { return u.rol; });
}

/** true si el usuario tiene AL MENOS UNO de los roles esperados, entre
 * sus posiblemente varios roles activos. Comparacion via _esRol (ya
 * normaliza acentos/mayusculas) para cada combinacion. */
function _tieneAlgunRol(ss, rolesEsperados) {
  var roles = _rolesUsuarioActual(ss);
  return roles.some(function (r) { return rolesEsperados.some(function (re) { return _esRol(r, re); }); });
}

/** Usuario/roles de la sesion actual, para que el frontend oculte
 * botones que no le correspondan. catalogoVacio=true mientras nadie haya
 * llenado CAT_USUARIOS -- el frontend debe interpretarlo como "mostrar
 * todo". `roles` es un ARRAY (2026-07-26, ver _rolesUsuarioActual) --
 * una persona puede tener varias filas con el mismo email, una por rol.
 * `rol` (singular) se retira a proposito, no se deja como alias de
 * "roles[0]" -- eso reintroduciria el mismo bug que se esta corrigiendo
 * (asumir un solo rol "principal" arbitrario). Todo consumidor de este
 * campo debe revisar el array completo. */
function getUsuarioActual() {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    var usuarios = _pagoSheetToObjects(_ensureCatUsuariosSheet(ss), CAT_USUARIOS_HEADERS);
    var email = Session.getActiveUser().getEmail();
    var propios = usuarios.filter(function (u) { return u.email === email && _esActivo(u.activo); });
    return {
      status: 'success',
      data: {
        email: email,
        nombre: propios.length ? propios[0].nombre : '',
        roles: propios.map(function (u) { return u.rol; }),
        catalogoVacio: usuarios.length === 0
      },
      message: ''
    };
  } catch (e) {
    return { status: 'error', data: {}, message: e.toString() };
  }
}

/** TC mas reciente disponible para el "gran total" convertido que pide
 * el usuario (punto 4 de la retroalimentacion 2026-07-26): sumar
 * MXN + USD*TC en una sola cifra. Reutiliza _getTcForDate (Codigo.js,
 * ya usado por carga-mov) -- mismo fallback a la fecha anterior mas
 * cercana si no hay TC exacto para hoy. No se crea un sistema de TC
 * nuevo, la hoja TC ya existe y ya se alimenta. */
function getTcActual() {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    var tc = _getTcForDate(ss, new Date());
    return { status: 'success', data: { tc: tc }, message: '' };
  } catch (e) {
    return { status: 'error', data: { tc: 1 }, message: e.toString() };
  }
}

/** Sabado siguiente a fechaCarga, 23:59:59 (limite de caducidad). Si
 * fechaCarga cae en sabado, el limite es el sabado DE LA SEMANA
 * SIGUIENTE (7 dias), no el mismo dia -- "el sabado siguiente" se lee
 * como el proximo, no el actual. */
function _fechaCaducidad(fechaCarga) {
  var d = new Date(fechaCarga);
  d.setHours(0, 0, 0, 0);
  var diasHastaSabado = (6 - d.getDay() + 7) % 7;
  if (diasHastaSabado === 0) diasHastaSabado = 7;
  var limite = new Date(d.getTime());
  limite.setDate(limite.getDate() + diasHastaSabado);
  limite.setHours(23, 59, 59, 999);
  return limite;
}

/** Devuelve todas las propuestas con sus partidas anidadas (tablero completo).
 * Calcula ESTADO='caducada' al vuelo (sin trigger) para partidas que
 * sigan 'activa' pero ya pasaron su fecha limite -- no se escribe de
 * vuelta a la hoja, es un calculo de lectura. */
function getPropuestasPago() {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) {
      return { status: 'sin_acceso', data: [], message: 'Tu cuenta no está registrada con un rol activo en CAT_USUARIOS. Contacta a finanzas para que te den de alta.' };
    }
    var propSheet = _ensurePropuestasPagoSheet(ss);
    var lineSheet = _ensurePartidasPagoSheet(ss);

    var propRows = _pagoSheetToObjects(propSheet, PROPUESTAS_PAGO_HEADERS);
    var lineRows = _pagoSheetToObjects(lineSheet, PARTIDAS_PAGO_HEADERS);

    // Acumulado aplicado por partida (Bloque B, 2026-07-27) -- una sola
    // lectura de APLICACIONES_PAGO para todas las lineas, no una consulta
    // por partida (mismo criterio de rendimiento que roles/catalogoVacio
    // en updatePartidaPago/bulkUpdatePartidasPago).
    var acumuladoPorPartida = _acumuladoAplicadoPorPartida(ss);

    var propById = {};
    propRows.forEach(function (p) { propById[p.id] = p; p.lineas = []; });

    var ahora = new Date();
    lineRows.forEach(function (l) {
      var prop = propById[l.propuestaId];
      if (prop && l.estado === 'activa') {
        // Hallazgo real 2026-07-27: si la partida fue rechazada y luego
        // reactivada, calcular SIEMPRE contra prop.fechaCarga (la carga
        // original de la PROPUESTA) hacia que una partida reactivada
        // DESPUES de esa fecha limite volviera a marcarse 'caducada' en
        // la siguiente lectura -- y 'caducada' no tiene ninguna accion
        // disponible en la UI, un callejon sin salida real. l.fechaReactivacion
        // (estampada por _camposAuditoriaPorFase al reactivar) tiene
        // prioridad cuando existe; una partida que nunca se rechazo no
        // tiene ese campo y calcula exactamente igual que antes.
        var limite = _fechaCaducidad(l.fechaReactivacion || prop.fechaCarga);
        if (ahora > limite) l.estado = 'caducada';
      }
      // Progreso de aplicacion parcial (Bloque B) -- montoEfectivoTotal es
      // el objetivo final (respeta autorizacion parcial), montoAplicadoAcumulado
      // es la suma de APLICACIONES_PAGO hasta ahora, saldoPendiente la resta.
      // Usado por la pestaña "Aplicado" de View_Pagos.html para mostrar
      // "$X de $Y aplicado" y precargar el input de monto.
      l.montoEfectivoTotal = _montoEfectivoPartida(l);
      l.montoAplicadoAcumulado = acumuladoPorPartida[l.id] || 0;
      l.saldoPendienteAplicar = Math.max(0, l.montoEfectivoTotal - l.montoAplicadoAcumulado);
      if (prop) prop.lineas.push(l);
    });

    return { status: 'success', data: propRows, message: '' };
  } catch (e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

/** Hallazgo real 2026-07-27 (auditoria de workflow pedida por el
 * usuario): el gateo de rol por transicion vivia SOLO en el cliente
 * (puedeActuar()/esTesoreria() por boton, View_Pagos.html) -- el
 * servidor solo validaba "¿tiene ALGUN rol activo?" (_tieneAccesoValido),
 * nunca el rol especifico ni la fase de origen de la partida. Cualquiera
 * con acceso valido a CUALQUIER rol podia, llamando updatePartidaPago/
 * bulkUpdatePartidasPago directo (fuera de los botones), disparar
 * cualquier transicion o saltarse el orden de fases.
 *
 * Esta funcion repite en el servidor EXACTAMENTE lo que el frontend ya
 * exige por boton -- no inventa ninguna regla de negocio nueva:
 *   fase:'tesoreria'    -> Contador,   partida debia estar en 'contabilidad'
 *   status:'autorizada' -> Tesoreria,  partida debia estar en 'tesoreria'
 *   estado:'rechazada'  -> Tesoreria,  la partida debia estar 'activa'
 *   status:'liberada'   -> Contador,   partida debia estar en 'liberacion2'
 *   subido               -> Tesoreria,  si pasa a true, debia estar en 'banca'
 *   pagado                -> Tesoreria,  si pasa a true, debia estar en 'aplicado'
 *                          Y el acumulado de APLICACIONES_PAGO debe cubrir
 *                          el monto efectivo completo (Bloque B)
 *   approvalType/montoAutorizado (fuera de la transicion 'autorizada')
 *                        -> Tesoreria (ver setType/autorizarSeleccionadas)
 *   cuentaPago/monedaPagoReal/tcAplicado -> Tesoreria (ver setCuentaPago/
 *                          setPagoReal, unicos gateados a esTesoreria() en
 *                          la pestaña Aplicado)
 *   reactivacion instantanea -> Tesoreria, la partida debia estar
 *                          'rechazada' (ver mas abajo)
 *
 * `estadoActual==='activa'` (2026-07-27, hallazgo real del usuario
 * durante la planificacion del Bloque B): esta funcion SOLO revisaba
 * FASE, nunca ESTADO -- una partida rechazada conserva su ultima FASE
 * tal cual (nunca se pisa al rechazar, decision historica), asi que
 * llamando updatePartidaPago directo con un patch que calzara esa MISMA
 * fase (ej. autorizar una partida rechazada mientras seguia en fase
 * 'tesoreria') el gate de fase sola NUNCA lo detectaba -- "Cancelar"
 * existe especificamente para decir "no le des mas dinero a esto", asi
 * que CUALQUIER transicion de avance debe exigir que la partida siga
 * 'activa'. ESTADO solo toma 'activa'/'rechazada' de verdad en la hoja
 * ('caducada' es un calculo de lectura en getPropuestasPago, nunca se
 * persiste) -- por eso este chequeo en la practica es "¿esta rechazada?".
 *
 * catalogoVacio=true se salta TODO el gate (mismo criterio "falla
 * ABIERTO" que _tieneAccesoValido/_puedeConciliarOComprobantes -- no
 * bloquear el primer despliegue con CAT_USUARIOS sin llenar). roles debe
 * venir ya calculado UNA vez por request (no recalculado por partida
 * dentro de un bulk), igual que catalogoVacio. */
function _validarTransicionPermitida(roles, catalogoVacio, faseActual, estadoActual, patch) {
  if (catalogoVacio) return { ok: true };
  function tiene(rolesEsperados) {
    return roles.some(function (r) { return rolesEsperados.some(function (re) { return _esRol(r, re); }); });
  }
  if (patch.fase === 'tesoreria') {
    if (!tiene([ROL_CONTADOR])) return { ok: false, message: 'Solo Contador puede liberar esta partida a Tesorería.' };
    if (estadoActual !== 'activa') return { ok: false, message: 'La partida no está activa (fue rechazada).' };
    if (faseActual !== 'contabilidad') return { ok: false, message: 'La partida ya no está en revisión contable.' };
  }
  if (patch.status === 'autorizada') {
    if (!tiene([ROL_TESORERIA])) return { ok: false, message: 'Solo Tesorería puede autorizar el monto.' };
    if (estadoActual !== 'activa') return { ok: false, message: 'La partida no está activa (fue rechazada).' };
    if (faseActual !== 'tesoreria') return { ok: false, message: 'La partida no está en fase de autorización de Tesorería.' };
  }
  if (patch.estado === 'rechazada') {
    if (!tiene([ROL_TESORERIA])) return { ok: false, message: 'Solo Tesorería puede rechazar una partida.' };
    if (estadoActual !== 'activa') return { ok: false, message: 'Solo se puede rechazar una partida activa.' };
  }
  if (patch.status === 'liberada') {
    if (!tiene([ROL_CONTADOR])) return { ok: false, message: 'Solo Contador puede liberar la partida a banco.' };
    if (estadoActual !== 'activa') return { ok: false, message: 'La partida no está activa (fue rechazada).' };
    if (faseActual !== 'liberacion2') return { ok: false, message: 'La partida no está en fase de liberación contable.' };
  }
  if (patch.subido !== undefined) {
    if (!tiene([ROL_TESORERIA])) return { ok: false, message: 'Solo Tesorería puede modificar el estado de "subida a banco".' };
    if (estadoActual !== 'activa') return { ok: false, message: 'La partida no está activa (fue rechazada).' };
    if (patch.subido === true && faseActual !== 'banca') return { ok: false, message: 'La partida no está en fase de carga a banca.' };
  }
  if (patch.pagado !== undefined) {
    if (!tiene([ROL_TESORERIA])) return { ok: false, message: 'Solo Tesorería puede modificar el estado de pago.' };
    if (estadoActual !== 'activa') return { ok: false, message: 'La partida no está activa (fue rechazada).' };
    if (patch.pagado === true && faseActual !== 'aplicado') return { ok: false, message: 'La partida no está en fase de pago aplicado.' };
  }
  if (patch.approvalType !== undefined || patch.montoAutorizado !== undefined) {
    if (!tiene([ROL_TESORERIA])) return { ok: false, message: 'Solo Tesorería puede definir el tipo/monto autorizado.' };
    if (estadoActual !== 'activa') return { ok: false, message: 'La partida no está activa (fue rechazada).' };
  }
  if (patch.cuentaPago !== undefined || patch.monedaPagoReal !== undefined || patch.tcAplicado !== undefined) {
    if (!tiene([ROL_TESORERIA])) return { ok: false, message: 'Solo Tesorería puede capturar la cuenta/moneda de pago.' };
    if (estadoActual !== 'activa') return { ok: false, message: 'La partida no está activa (fue rechazada).' };
  }
  // C-1 (auditoria 2026-07-28, hallazgo real): un patch con `fase` que no
  // fuera 'tesoreria' NO tenia ningun chequeo de que ese valor coincidiera
  // con la transicion real que lo autoriza -- los bloques de arriba
  // (status:'autorizada', status:'liberada', subido) validan rol/fase-de-
  // origen/estado, pero nunca revisaban que patch.fase (si venia en el
  // mismo patch) fuera el destino correcto. Un patch directo al backend
  // como {fase:'aplicado'} solo, o {status:'autorizada', fase:'aplicado'}
  // (mezclando una transicion real con un destino de fase ajeno), pasaba
  // sin objecion porque ninguna regla evaluaba ese valor -- saltandose
  // revision/autorizacion/liberacion completas, sin dejar auditoria
  // (REVISADO/AUTORIZADO/LIBERADO_POR vacios) ni disparar correo. Este
  // bloque cierra eso: si patch.fase viene y no es 'tesoreria' (ya
  // cubierto arriba), debe ser EXACTAMENTE el destino que corresponde a
  // la transicion real que el mismo patch tambien trae.
  if (patch.fase !== undefined && patch.fase !== 'tesoreria') {
    var destinoFaseValido =
      (patch.fase === 'liberacion2' && patch.status === 'autorizada') ||
      (patch.fase === 'banca' && patch.status === 'liberada') ||
      (patch.fase === 'aplicado' && patch.subido === true);
    if (!destinoFaseValido) return { ok: false, message: 'Combinación de fase/transición no reconocida.' };
  }
  // Hallazgo real de `revisor` (2026-07-27, ronda de correccion de este
  // mismo bloque): la reactivacion INSTANTANEA (firma exclusiva
  // estado:'activa'+reactivacionSolicitada:false, ver solicitarReactivacion)
  // no tenia gate aqui -- cualquier usuario con acceso valido a CUALQUIER
  // rol podia llamar updatePartidaPago directo con ese patch y reactivar
  // una partida rechazada sin ser Tesoreria, dejando ademas una auditoria
  // (REACTIVADO_POR) que aparentaba una aprobacion legitima. `solicitarReactivacion`
  // YA hace su propio chequeo de rol antes de escribir (_tieneAlgunRol),
  // pero eso no protege una llamada directa a updatePartidaPago/
  // bulkUpdatePartidasPago con la misma firma de patch -- el gate real
  // tiene que vivir aqui, igual que las demas transiciones. Se agrega
  // ademas (2026-07-27, segunda ronda) exigir estadoActual==='rechazada'
  // -- solo se puede reactivar algo que este rechazado, mismo chequeo que
  // solicitarReactivacion ya hace por su cuenta antes de llegar aqui,
  // repetido en el choke point real para que una llamada directa al
  // backend no se lo salte.
  // C-1 (auditoria 2026-07-28, hallazgo real): la rama de abajo solo
  // reaccionaba a la firma COMPLETA de reactivacion instantanea
  // (estado:'activa' + reactivacionSolicitada:false) -- un patch
  // {estado:'activa'} SOLO (sin el segundo campo) no calzaba esa firma y
  // caia directo al `return {ok:true}` final sin ningun gate, dejando
  // reactivar una partida rechazada a CUALQUIER rol valido, sin que
  // _camposAuditoriaPorFase estampe REACTIVADO_POR (esa funcion tambien
  // exige la firma pareada) -- una reactivacion silenciosa, sin rastro.
  // Ahora CUALQUIER patch con estado:'activa' pasa por este gate: con la
  // firma completa se aplica el chequeo real de Tesoreria+rechazada de
  // siempre; cualquier otra combinacion se rechaza explicito en vez de
  // caer al default.
  if (patch.estado === 'activa') {
    if (patch.reactivacionSolicitada === false) {
      if (!tiene([ROL_TESORERIA])) return { ok: false, message: 'Solo Tesorería puede reactivar una partida.' };
      if (estadoActual !== 'rechazada') return { ok: false, message: 'Solo se puede reactivar una partida rechazada.' };
    } else {
      return { ok: false, message: 'Transición a estado "activa" no reconocida (falta la firma de reactivación).' };
    }
  }
  return { ok: true };
}

/** Devuelve {ok:true} o {ok:false, message} -- nunca lanza por datos
 * invalidos, para que el llamador pueda reportar el motivo exacto al
 * usuario en vez de fallar generico. Valida rol+fase+estado de origen
 * contra la transicion pedida (ver _validarTransicionPermitida) antes que
 * cualquier otra cosa. Valida MONTO_AUTORIZADO contra el MONTO original
 * (hallazgo del revisor: antes se aceptaba cualquier numero, incluyendo
 * negativos o por encima de la factura). Tambien estampa el par
 * POR/FECHA de auditoria que corresponda segun la forma del patch (ver
 * _camposAuditoriaPorFase) -- no para cambios triviales como un checkbox
 * o un comentario.
 *
 * roles/catalogoVacio: calculados UNA vez por el llamador (updatePartidaPago/
 * bulkUpdatePartidasPago/setPartidasPagoChecked/solicitarReactivacion/
 * aplicarPagoParcial), nunca recalculados aqui adentro -- evita releer
 * CAT_USUARIOS por cada partida de un lote.
 *
 * `ss` (2026-07-27, Bloque B): necesario SOLO para validar `patch.pagado
 * === true` contra el acumulado real de APLICACIONES_PAGO -- ninguna
 * otra validacion de esta funcion lo necesita. Hallazgo real: sin esto,
 * cualquiera con rol Tesoreria podria llamar updatePartidaPago/
 * bulkUpdatePartidasPago directo con {pagado:true} y marcar una partida
 * como pagada sin que el acumulado de aplicaciones parciales la respalde
 * -- `pagado` debe ser un estado DERIVADO, nunca forzable directo. */
function _updatePartidaPagoRow(sheet, partidaId, patch, roles, catalogoVacio, ss) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: false, message: 'Sin partidas' };
  var idCol = PARTIDAS_PAGO_HEADERS.indexOf('ID') + 1;
  var faseCol = PARTIDAS_PAGO_HEADERS.indexOf('FASE') + 1;
  var estadoCol = PARTIDAS_PAGO_HEADERS.indexOf('ESTADO') + 1;
  var ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (ids[i][0] === partidaId) {
      var rowIndex = i + 2;

      var faseActual = sheet.getRange(rowIndex, faseCol).getValue();
      var estadoActual = sheet.getRange(rowIndex, estadoCol).getValue();
      var chk = _validarTransicionPermitida(roles || [], !!catalogoVacio, faseActual, estadoActual, patch);
      if (!chk.ok) return { ok: false, message: chk.message + ' (partida ' + partidaId + ').' };

      if (patch.pagado === true && ss) {
        var partidaCompleta = _pagoSheetToObjects(sheet, PARTIDAS_PAGO_HEADERS).filter(function (p) { return p.id === partidaId; })[0];
        var objetivo = _montoEfectivoPartida(partidaCompleta);
        var acumulado = _aplicacionesPorPartida(ss, partidaId).reduce(function (s, a) { return s + (Number(a.monto) || 0); }, 0);
        if (objetivo - acumulado >= 1) {
          return { ok: false, message: 'No se puede marcar como pagada: aplicado $' + acumulado.toFixed(2) + ' de $' + objetivo.toFixed(2) + ' (partida ' + partidaId + '). Usa aplicarPagoParcial para completar el saldo.' };
        }
      }

      if (patch.montoAutorizado !== undefined) {
        var montoCol = PARTIDAS_PAGO_HEADERS.indexOf('MONTO') + 1;
        var monto = Number(sheet.getRange(rowIndex, montoCol).getValue());
        var val = Number(patch.montoAutorizado);
        if (isNaN(val) || val <= 0) return { ok: false, message: 'Monto autorizado debe ser mayor a cero (partida ' + partidaId + ').' };
        if (val > monto) return { ok: false, message: 'Monto autorizado no puede exceder ' + monto + ' (partida ' + partidaId + ').' };
      }

      // Validaciones de los campos nuevos de conciliacion (2026-07-26):
      // monedaPagoReal solo tiene sentido si trae un valor real (no se
      // fuerza a estar siempre presente, MXN->MXN no manda el campo);
      // tcAplicado, si se manda, debe ser positivo -- un TC de 0 o
      // negativo rompe silenciosamente la busqueda de candidatos de
      // conciliacion (monto buscado = monto * tc).
      if (patch.monedaPagoReal !== undefined && patch.monedaPagoReal !== '' && MONEDAS_VALIDAS.indexOf(patch.monedaPagoReal) < 0) {
        return { ok: false, message: 'Moneda de pago real invalida (' + patch.monedaPagoReal + '), partida ' + partidaId + '.' };
      }
      if (patch.tcAplicado !== undefined && patch.tcAplicado !== '') {
        var tcVal = Number(patch.tcAplicado);
        if (isNaN(tcVal) || tcVal <= 0) return { ok: false, message: 'TC aplicado debe ser mayor a cero (partida ' + partidaId + ').' };
      }

      // CUENTA_PAGO obligatoria antes de marcar Pagada (2026-07-26,
      // pedido explicito del usuario -- sin esto la conciliacion no
      // puede agrupar la partida por banco). Solo se exige al pasar a
      // pagado:true, nunca al desmarcar (revertir un error no debe
      // quedar bloqueado por un dato que ya no importa). Lee el valor
      // ya persistido en la hoja si el patch no trae cuentaPago en la
      // misma llamada (caso normal: se captura antes, en un click
      // aparte, ver setCuentaPago en View_Pagos.html).
      if (patch.pagado === true) {
        var cuentaCol = PARTIDAS_PAGO_HEADERS.indexOf('CUENTA_PAGO') + 1;
        var cuentaActual = patch.cuentaPago !== undefined ? patch.cuentaPago : sheet.getRange(rowIndex, cuentaCol).getValue();
        if (!cuentaActual) return { ok: false, message: 'No se puede marcar como pagada sin capturar la cuenta de pago (partida ' + partidaId + ').' };
      }

      var extra = _camposAuditoriaPorFase(patch, Session.getActiveUser().getEmail(), new Date());
      var full = Object.keys(extra).length ? Object.assign({}, patch, extra) : patch;
      Object.keys(full).forEach(function (key) {
        var headerKey = key.replace(/([A-Z])/g, '_$1').toUpperCase();
        var col = PARTIDAS_PAGO_HEADERS.indexOf(headerKey) + 1;
        if (col > 0) sheet.getRange(rowIndex, col).setValue(full[key]);
      });
      return { ok: true };
    }
  }
  return { ok: false, message: 'Partida no encontrada: ' + partidaId };
}

/** payload = { partidaId, patch }. patch = { checked, approvalType,
 * montoAutorizado, status, subido, pagado, comentario, notaAplicacion,
 * fase, estado }.
 *
 * BUG REAL corregido 2026-07-26: esta funcion (y bulkUpdatePartidasPago/
 * setPartidasPagoChecked) antes recibian (partidaId, patch) como DOS
 * parametros separados, y el frontend las llamaba con
 * VEVA.gas('updatePartidaPago', l.id, patch) -- TRES argumentos. Pero
 * VEVA.gas (Shared_Core.html) tiene la firma function(fnName, payload):
 * solo reenvia UN payload a google.script.run. El segundo argumento real
 * (`patch`) nunca llegaba al backend -- se perdia en silencio, sin error
 * visible, porque JS no truena por argumentos de sobra en una llamada.
 * Resultado: autorizar/rechazar/liberar/marcar subida-pagado/comentarios/
 * checkboxes nunca escribian su patch real (solo el `partidaId`), aunque
 * la UI mostraba el cambio de forma optimista. Se detecto al probar la
 * carga de propuesta end-to-end, no lo podia atrapar `revisor`/`qa`
 * (leen codigo, no ejecutan el navegador real). Fix: un solo objeto de
 * payload, mismo patron que ya usa el resto de TESO (getFlujoEfectivo,
 * getMovimientos, etc. -- todos pasan un objeto, nunca argumentos
 * posicionales sueltos). */
function updatePartidaPago(payload) {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var usuarios = _pagoSheetToObjects(_ensureCatUsuariosSheet(ss), CAT_USUARIOS_HEADERS);
    var res = _updatePartidaPagoRow(_ensurePartidasPagoSheet(ss), payload.partidaId, payload.patch, _rolesUsuarioActual(ss), usuarios.length === 0, ss);
    if (!res.ok) return { status: 'error', message: res.message };
    _notificarTransicionPagos(ss, [payload.partidaId], _tipoTransicionNotificable(payload.patch), Session.getActiveUser().getEmail());
    return { status: 'success', message: '' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/** payload = { partidaIds, patch }. Aplica el mismo patch a varias
 * partidas (seleccion masiva por fase). Este es tambien el mecanismo de
 * avance de fase: los botones de accion del frontend (autorizar/liberar/
 * subir/aplicar) incluyen `fase` en el patch cuando ese paso efectivamente
 * mueve la partida a la siguiente fase -- no existe una funcion ni un
 * boton de "avanzar fase" aparte. */
function bulkUpdatePartidasPago(payload) {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var sheet = _ensurePartidasPagoSheet(ss);
    // roles/catalogoVacio se calculan UNA vez para todo el lote (no por
    // partida dentro del forEach) -- releer CAT_USUARIOS en cada
    // iteracion seria lento y, si alguien editara su fila a mitad de un
    // lote grande, daria resultados inconsistentes entre partidas del
    // mismo request.
    var usuarios = _pagoSheetToObjects(_ensureCatUsuariosSheet(ss), CAT_USUARIOS_HEADERS);
    var roles = _rolesUsuarioActual(ss);
    var catalogoVacio = usuarios.length === 0;
    var errores = [];
    var idsExitosos = [];
    (payload.partidaIds || []).forEach(function (id) {
      var res = _updatePartidaPagoRow(sheet, id, payload.patch, roles, catalogoVacio, ss);
      if (!res.ok) errores.push(res.message); else idsExitosos.push(id);
    });
    if (idsExitosos.length) _notificarTransicionPagos(ss, idsExitosos, _tipoTransicionNotificable(payload.patch), Session.getActiveUser().getEmail());
    // Bug real corregido en Bloque F (2026-07-28): antes, si UNA partida
    // del lote fallaba validacion, se devolvia status:'error' para TODO
    // el lote aunque las demas ya hubieran quedado escritas con exito
    // (esta funcion nunca fue atomica a nivel de lote, solo por fila) --
    // sin decir cuales pasaron y cuales no. Ahora el contrato distingue
    // los 3 casos reales: 'success' (todas pasaron), 'parcial' (algunas
    // si, algunas no -- exitosas ya quedaron guardadas), 'error' (ninguna
    // paso). El detalle por partida sigue disponible en `data`.
    var status = errores.length === 0 ? 'success' : (idsExitosos.length > 0 ? 'parcial' : 'error');
    return { status: status, data: { idsExitosos: idsExitosos, errores: errores }, message: errores.join(' ') };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/** payload = { partidaIds, checked }. Marca/desmarca varias partidas
 * (checkbox "seleccionar todas"). */
function setPartidasPagoChecked(payload) {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var sheet = _ensurePartidasPagoSheet(ss);
    // { checked } no dispara ningun gate de rol en _validarTransicionPermitida
    // (fuera de alcance a proposito, igual que en el frontend) -- se
    // calculan roles/catalogoVacio solo por consistencia de firma con
    // los demas llamadores, sin efecto real aqui.
    var usuarios = _pagoSheetToObjects(_ensureCatUsuariosSheet(ss), CAT_USUARIOS_HEADERS);
    var roles = _rolesUsuarioActual(ss);
    var catalogoVacio = usuarios.length === 0;
    (payload.partidaIds || []).forEach(function (id) { _updatePartidaPagoRow(sheet, id, { checked: payload.checked }, roles, catalogoVacio, ss); });
    return { status: 'success', message: '' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/** Un solo boton "Solicitar reactivacion" para una partida rechazada:
 * si quien lo pide tiene rol Tesoreria, se reactiva al instante (vuelve
 * a ESTADO='activa'); si lo pide cualquier otro rol, se marca
 * REACTIVACION_SOLICITADA=true y queda pendiente hasta que alguien con
 * rol Tesoreria llame a esta misma funcion sobre la misma partida.
 * Una partida 'caducada' nunca es reactivable (regla explicita, sin
 * excepcion). */
function solicitarReactivacion(partidaId) {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    // Hallazgo real de `revisor` (2026-07-27): la reactivacion instantanea
    // ahora tiene gate propio dentro de _validarTransicionPermitida (solo
    // Tesoreria) -- hay que calcular roles/catalogoVacio aqui igual que
    // en updatePartidaPago/bulkUpdatePartidasPago/setPartidasPagoChecked,
    // si no, _updatePartidaPagoRow recibiria roles=[] por default y
    // bloquearia incluso a un Tesoreria legitimo reactivando desde el
    // boton real.
    var usuarios = _pagoSheetToObjects(_ensureCatUsuariosSheet(ss), CAT_USUARIOS_HEADERS);
    var roles = _rolesUsuarioActual(ss);
    var catalogoVacio = usuarios.length === 0;
    var sheet = _ensurePartidasPagoSheet(ss);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { status: 'error', message: 'Sin partidas' };

    var idCol = PARTIDAS_PAGO_HEADERS.indexOf('ID') + 1;
    var estadoCol = PARTIDAS_PAGO_HEADERS.indexOf('ESTADO') + 1;
    var ids = sheet.getRange(2, idCol, lastRow - 1, 1).getValues();

    for (var i = 0; i < ids.length; i++) {
      if (ids[i][0] !== partidaId) continue;
      var rowIndex = i + 2;
      var estadoActual = sheet.getRange(rowIndex, estadoCol).getValue();

      if (estadoActual === 'caducada') {
        return { status: 'error', message: 'No se puede reactivar: la partida ya caducó.' };
      }
      if (estadoActual !== 'rechazada') {
        return { status: 'error', message: 'Solo se puede solicitar reactivación de una partida rechazada.' };
      }

      // Se escribe via _updatePartidaPagoRow (no setValue directo) para
      // que quede la misma auditoria de quien/cuando que el resto de
      // acciones de negocio (ver ACCIONES_AUDITABLES). _tieneAlgunRol ya
      // normaliza acentos/mayusculas y revisa TODOS los roles activos
      // del usuario, no solo uno -- hallazgo real 2026-07-26: alguien
      // con varias filas en CAT_USUARIOS (una por rol) quedaba atrapado
      // con el primer rol de la hoja si se comparaba un solo valor.
      if (_tieneAlgunRol(ss, [ROL_TESORERIA])) {
        _updatePartidaPagoRow(sheet, partidaId, { estado: 'activa', reactivacionSolicitada: false }, roles, catalogoVacio, ss);
        return { status: 'success', data: { reactivadaAlInstante: true }, message: 'Partida reactivada.' };
      }
      _updatePartidaPagoRow(sheet, partidaId, { reactivacionSolicitada: true }, roles, catalogoVacio, ss);
      return { status: 'success', data: { reactivadaAlInstante: false }, message: 'Solicitud enviada, pendiente de confirmación de Tesorería.' };
    }
    return { status: 'error', message: 'Partida no encontrada' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/** payload = { partidaId, monto }. Libera un pago PARCIAL sobre una
 * partida en fase 'aplicado' -- inserta una fila nueva en
 * APLICACIONES_PAGO (auditoria completa de CADA liberacion, nunca un
 * campo acumulado sin historial, decision explicita del usuario Bloque B
 * 2026-07-27) y, si el acumulado alcanza el monto efectivo completo
 * (_montoEfectivoPartida, respeta autorizacion parcial), marca la
 * partida como pagada de verdad (PAGADO=true, APLICADO_POR/FECHA) via
 * _updatePartidaPagoRow -- ese es el UNICO camino real para que PAGADO
 * pase a true, nunca se fuerza aqui a mano.
 *
 * LockService: dos llamadas concurrentes sobre la MISMA partida podrian,
 * sin este candado, leer el acumulado ANTES de que la otra escriba su
 * fila y terminar sumando por encima del saldo pendiente -- riesgo real
 * de colision que no existia antes de este bloque (las demas mutaciones
 * de PARTIDAS_PAGO tocan una sola fila por ID; esta ademas suma contra un
 * acumulado que puede crecer entre lectura y escritura).
 *
 * estadoActual==='activa' se revisa aqui TAMBIEN (no solo en
 * _validarTransicionPermitida) porque una aplicacion parcial que NO
 * completa el 100% nunca llega a llamar _updatePartidaPagoRow -- si no
 * se revisara aqui, se podria seguir aplicando dinero a una partida ya
 * rechazada sin pasar nunca por ese gate (hallazgo real del usuario:
 * "Cancelar" existe para decir "no le des mas dinero a esto"). */
/** Nucleo de aplicarPagoParcial (extraido en Bloque F, 2026-07-28, para
 * reusarlo desde bulkResolverAplicado sin duplicar ninguna validacion).
 * NO adquiere/libera LockService ni notifica -- eso lo hace cada
 * llamador (aplicarPagoParcial con su propio lock individual;
 * bulkResolverAplicado con un lock por partida DENTRO del loop, mismo
 * patron, para no bloquear el modulo completo durante un lote grande).
 * Devuelve { resultado, partidaIdCompletada, emailActor } -- resultado
 * es exactamente el mismo shape {status,data,message} que antes
 * devolvia aplicarPagoParcial completo. */
function _aplicarPagoParcialCore(ss, partidaId, montoCrudo, roles, catalogoVacio) {
  var sheet = _ensurePartidasPagoSheet(ss);
  var partida = _pagoSheetToObjects(sheet, PARTIDAS_PAGO_HEADERS).filter(function (p) { return p.id === partidaId; })[0];
  if (!partida) return { resultado: { status: 'error', message: 'Partida no encontrada: ' + partidaId }, partidaIdCompletada: null };

  if (!catalogoVacio && partida.estado !== 'activa') {
    return { resultado: { status: 'error', message: 'La partida no está activa (fue rechazada) -- no se puede aplicar más pago.' }, partidaIdCompletada: null };
  }
  if (partida.fase !== 'aplicado') {
    return { resultado: { status: 'error', message: 'La partida no está en fase de pago aplicado.' }, partidaIdCompletada: null };
  }
  if (!partida.cuentaPago) {
    return { resultado: { status: 'error', message: 'No se puede aplicar pago sin capturar la cuenta de pago.' }, partidaIdCompletada: null };
  }

  var monto = Number(montoCrudo);
  if (isNaN(monto) || monto <= 0) return { resultado: { status: 'error', message: 'El monto a aplicar debe ser mayor a cero.' }, partidaIdCompletada: null };

  var objetivo = _montoEfectivoPartida(partida);
  var acumuladoAnterior = _aplicacionesPorPartida(ss, partida.id).reduce(function (s, a) { return s + (Number(a.monto) || 0); }, 0);
  var saldoPendiente = objetivo - acumuladoAnterior;
  // Tolerancia de $0.01 (no <$1 como la conciliacion banco<->partida):
  // aqui no se esta cuadrando contra un tercero con redondeos de tipo
  // de cambio, es una resta exacta contra un limite de negocio propio.
  if (monto - saldoPendiente > 0.01) {
    return { resultado: { status: 'error', message: 'El monto ($' + monto.toFixed(2) + ') excede el saldo pendiente ($' + saldoPendiente.toFixed(2) + ') de la partida.' }, partidaIdCompletada: null };
  }

  var emailActor = Session.getActiveUser().getEmail();
  var ahora = new Date();
  var aplicSheet = _ensureAplicacionesPagoSheet(ss);
  aplicSheet.getRange(aplicSheet.getLastRow() + 1, 1, 1, APLICACIONES_PAGO_HEADERS.length).setValues([[
    Utilities.getUuid(), partida.id, monto, emailActor, ahora
  ]]);

  var nuevoAcumulado = acumuladoAnterior + monto;
  // Mismo criterio de tolerancia <$1 que el resto del modulo usa para
  // "ya cuadra" (ver conciliarPartida/getConciliacionPagos).
  var completo = (objetivo - nuevoAcumulado) < 1;
  var partidaIdCompletada = null;
  if (completo) {
    var res = _updatePartidaPagoRow(sheet, partida.id, { pagado: true }, roles, catalogoVacio, ss);
    if (!res.ok) {
      // No deberia pasar (rol/fase/estado ya se validaron arriba y el
      // acumulado recien escrito ya cubre el objetivo) -- pero si pasa,
      // la aplicacion parcial YA quedo registrada en APLICACIONES_PAGO
      // (no se revierte, es el registro real de que el dinero salio);
      // se reporta el error para que alguien revise manualmente por
      // que no se pudo cerrar la partida.
      return { resultado: { status: 'error', message: 'Pago parcial registrado, pero no se pudo cerrar la partida: ' + res.message }, partidaIdCompletada: null, emailActor: emailActor };
    }
    partidaIdCompletada = partida.id;
  }

  return {
    resultado: {
      status: 'success',
      data: { completo: completo, montoAplicadoAcumulado: nuevoAcumulado, saldoPendiente: Math.max(0, objetivo - nuevoAcumulado) },
      message: completo ? 'Partida aplicada por completo.' : ('Pago parcial de $' + monto.toFixed(2) + ' registrado.')
    },
    partidaIdCompletada: partidaIdCompletada,
    emailActor: emailActor
  };
}

function aplicarPagoParcial(payload) {
  var lock = LockService.getScriptLock();
  var partidaIdCompletada = null, emailActor = '', ss;
  try {
    lock.waitLock(30000);
    ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };

    var usuarios = _pagoSheetToObjects(_ensureCatUsuariosSheet(ss), CAT_USUARIOS_HEADERS);
    var catalogoVacio = usuarios.length === 0;
    if (!catalogoVacio && !_tieneAlgunRol(ss, [ROL_TESORERIA])) {
      return { status: 'error', message: 'Solo Tesorería puede aplicar pagos.' };
    }

    var out = _aplicarPagoParcialCore(ss, payload.partidaId, payload.monto, _rolesUsuarioActual(ss), catalogoVacio);
    partidaIdCompletada = out.partidaIdCompletada;
    emailActor = out.emailActor || '';
    return out.resultado;
  } catch (e) {
    return { status: 'error', message: e.toString() };
  } finally {
    lock.releaseLock();
    // Notificacion FUERA del lock (mismo patron que updatePartidaPago) --
    // _notificarTransicionPagos ya nunca lanza (try/catch interno propio),
    // no hace falta envolverla aqui otra vez.
    if (partidaIdCompletada && ss) {
      _notificarTransicionPagos(ss, [partidaIdCompletada], _tipoTransicionNotificable({ pagado: true }), emailActor);
    }
  }
}

/** payload = { aplicar:[{partidaId,monto}], idsNoPagado:[...] }. Bloque
 * F (2026-07-28, rediseño de checkbox pedido por el usuario): resuelve
 * TODAS las partidas marcadas en la pestaña "Aplicado" de una sola vez.
 * `aplicar[]` reusa el nucleo de aplicarPagoParcial (_aplicarPagoParcialCore)
 * con LockService POR PARTIDA dentro del loop -- mismo criterio que
 * aplicarPagoParcial individual, nunca un lock unico para todo el lote
 * (bloquearia el modulo completo durante un lote grande). `idsNoPagado[]`
 * marca {estado:'rechazada'} via _updatePartidaPagoRow, CON el mismo
 * LockService por partida que la rama `aplicar` (hallazgo real de
 * auditoria 2026-07-28, D-9: sin este candado, alguien marcando "No
 * pagado" sobre una partida podia correr sin exclusion mutua contra otro
 * usuario aplicandole un parcial simultaneo a la MISMA partida -- el
 * lock de la rama `aplicar` no protege contra un escritor que no lo
 * toma). Si una linea no tiene cuenta de pago capturada, el frontend ya
 * bloquea el envio antes de llegar aqui (resolverAplicado,
 * View_Pagos.html) -- el backend igual la rechazaria via
 * _aplicarPagoParcialCore como defensa en profundidad. Notifica UNA sola
 * vez al final, con las 2 secciones (aplicadas primero, no pagadas
 * despues con su motivo -- ver COMENTARIO en PARTIDAS_PAGO). */
function bulkResolverAplicado(payload) {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };

    var usuarios = _pagoSheetToObjects(_ensureCatUsuariosSheet(ss), CAT_USUARIOS_HEADERS);
    var catalogoVacio = usuarios.length === 0;
    if (!catalogoVacio && !_tieneAlgunRol(ss, [ROL_TESORERIA])) {
      return { status: 'error', message: 'Solo Tesorería puede resolver esta fase.' };
    }
    var roles = _rolesUsuarioActual(ss);

    var idsAplicadosCompletados = [];
    var exitososAplicar = 0;
    var errores = [];
    (payload.aplicar || []).forEach(function (item) {
      var lock = LockService.getScriptLock();
      try {
        lock.waitLock(30000);
        var out = _aplicarPagoParcialCore(ss, item.partidaId, item.monto, roles, catalogoVacio);
        if (out.resultado.status === 'success') {
          exitososAplicar++;
          if (out.partidaIdCompletada) idsAplicadosCompletados.push(out.partidaIdCompletada);
        } else {
          errores.push(item.partidaId + ': ' + out.resultado.message);
        }
      } catch (e) {
        errores.push(item.partidaId + ': ' + e.toString());
      } finally {
        lock.releaseLock();
      }
    });

    var sheet = _ensurePartidasPagoSheet(ss);
    var idsRechazadosExitosos = [];
    (payload.idsNoPagado || []).forEach(function (id) {
      var lock = LockService.getScriptLock();
      try {
        lock.waitLock(30000);
        var res = _updatePartidaPagoRow(sheet, id, { estado: 'rechazada' }, roles, catalogoVacio, ss);
        if (res.ok) idsRechazadosExitosos.push(id); else errores.push(id + ': ' + res.message);
      } catch (e) {
        errores.push(id + ': ' + e.toString());
      } finally {
        lock.releaseLock();
      }
    });

    var grupos = [];
    if (idsAplicadosCompletados.length) grupos.push({ idsExitosos: idsAplicadosCompletados, transicion: _tipoTransicionNotificable({ pagado: true }) });
    if (idsRechazadosExitosos.length) grupos.push({ idsExitosos: idsRechazadosExitosos, transicion: _tipoTransicionNotificable({ estado: 'rechazada' }) });
    if (grupos.length) _notificarTransicionPagosMultiple(ss, grupos, Session.getActiveUser().getEmail());

    var totalExitosos = exitososAplicar + idsRechazadosExitosos.length;
    var status = errores.length === 0 ? 'success' : (totalExitosos > 0 ? 'parcial' : 'error');
    return { status: status, data: { exitosos: totalExitosos, fallidos: errores.length }, message: errores.length ? errores.join(' · ') : '' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/** payload = { idsPositivos:[...], idsNegativos:[...] }. Bloque F
 * (2026-07-28): resuelve TODAS las partidas marcadas en la pestaña
 * "Banca" de una sola vez -- positivas avanzan {subido:true,
 * fase:'aplicado'} (mismo avance atomico que ya usaba el boton "Marcar
 * subidas"), negativas quedan {estado:'rechazada'}. Sin LockService por
 * partida (mismo patron ya establecido por bulkUpdatePartidasPago para
 * mutaciones via _updatePartidaPagoRow que no tocan ningun libro
 * contable aparte). Notifica UNA sola vez al final con las 2 secciones. */
function bulkResolverBanca(payload) {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };

    var usuarios = _pagoSheetToObjects(_ensureCatUsuariosSheet(ss), CAT_USUARIOS_HEADERS);
    var catalogoVacio = usuarios.length === 0;
    if (!catalogoVacio && !_tieneAlgunRol(ss, [ROL_TESORERIA])) {
      return { status: 'error', message: 'Solo Tesorería puede resolver esta fase.' };
    }
    var roles = _rolesUsuarioActual(ss);
    var sheet = _ensurePartidasPagoSheet(ss);

    function resolverLote(ids, patch) {
      var exitosos = [], errs = [];
      ids.forEach(function (id) {
        var res = _updatePartidaPagoRow(sheet, id, patch, roles, catalogoVacio, ss);
        if (res.ok) exitosos.push(id); else errs.push(id + ': ' + res.message);
      });
      return { exitosos: exitosos, errores: errs };
    }

    var resPos = resolverLote(payload.idsPositivos || [], { subido: true, fase: 'aplicado' });
    var resNeg = resolverLote(payload.idsNegativos || [], { estado: 'rechazada' });

    var grupos = [];
    if (resPos.exitosos.length) grupos.push({ idsExitosos: resPos.exitosos, transicion: _tipoTransicionNotificable({ subido: true }) });
    if (resNeg.exitosos.length) grupos.push({ idsExitosos: resNeg.exitosos, transicion: _tipoTransicionNotificable({ estado: 'rechazada' }) });
    if (grupos.length) _notificarTransicionPagosMultiple(ss, grupos, Session.getActiveUser().getEmail());

    var errores = resPos.errores.concat(resNeg.errores);
    var totalExitosos = resPos.exitosos.length + resNeg.exitosos.length;
    var status = errores.length === 0 ? 'success' : (totalExitosos > 0 ? 'parcial' : 'error');
    return { status: status, data: { exitosos: totalExitosos, fallidos: errores.length }, message: errores.length ? errores.join(' · ') : '' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/**
 * Carga de propuesta (2026-07-26) -- importar el Excel real de propuesta
 * de pago (formato "PROPUESTA CONSOLIDADA", confirmado contra el archivo
 * real, no una plantilla asumida) y crear PROPUESTAS_PAGO/PARTIDAS_PAGO
 * automaticamente. Patron de dos pasos igual que carga-mov
 * (previsualizarMovimientos/cargarMovimientos en Codigo.js): primero
 * previsualizarPropuestaPago (nunca escribe) para que el usuario revise
 * y excluya lo que no quiera, despues confirmarCargaPropuestaPago
 * (si escribe) solo con lo confirmado.
 *
 * Decisiones basadas en inspeccion real de PROPUESTA.xlsx + BD.xlsx
 * (no supuestos):
 * - Cada fila real trae el importe en UNA sola de las columnas
 *   Importe MXN/USD/GBP -- nunca en mas de una a la vez, una vez que se
 *   excluyen las filas de subtotal (el archivo real trae 6 filas de
 *   "GRAN TOTAL" mezcladas con las 113 filas reales de factura).
 * - El numero al final de "Nombre empresa" (ej. "Blau Life SAPI de CV
 *   8332") es el campo ULTIMOS_DIGITOS de una cuenta especifica en
 *   CAT_CUENTAS_MAPEO -- NO son los ultimos digitos literales de
 *   NUMERO_CUENTA (verificado: no coinciden), es una columna curada
 *   aparte. Se resuelve por ULTIMOS_DIGITOS, nunca se adivina el nombre
 *   de sociedad.
 * - La hoja "INTERCIAS" (traspasos entre cuentas propias, sin proveedor
 *   ni factura) queda fuera de este importador por decision explicita
 *   del usuario -- no encaja en el modelo Propuestas/Partidas.
 */
// 'Importe EUR' NO se agrega aqui a proposito (a diferencia de MXN/USD/GBP,
// que si son columnas obligatorias aunque una fila las deje vacias): el
// archivo real de la propuesta hoy nunca ha traido esa columna, y
// exigirla en la lista de requeridos romperia CUALQUIER carga futura que
// siga usando la plantilla actual, hasta que finanzas actualice su
// Excel. Se lee como OPCIONAL mas abajo (si la columna existe, se usa;
// si no, simplemente no hay importes en EUR posibles).
var PROPUESTA_HEADERS_REQUERIDOS = ['Nombre empresa', 'Proveedor', 'Nombre del Proveedor', 'Importe MXN', 'Importe USD', 'Importe GBP', 'Referencia', 'Texto', 'Solicitante', 'Nº documento', 'UUID', 'RFC'];

function _esHojaPropuestaConsolidada(headerRow) {
  var set = (headerRow || []).map(function (h) { return String(h || '').trim(); });
  return set.indexOf('Nombre empresa') >= 0 && set.indexOf('Proveedor') >= 0 && set.indexOf('Importe MXN') >= 0;
}

/** Lee CAT_CUENTAS_MAPEO buscando NOMBRE_CORTO, ULTIMOS_DIGITOS y
 * NOMBRE_SOCIEDAD (este ultimo agregado 2026-08-02 para el fallback por
 * nombre completo, ver _resolverSociedadDesdeNombreEmpresa). Distinto a
 * proposito de _buildCatCuentasArray (Codigo.js): esa funcion sirve para
 * matchear movimientos bancarios por NUMERO_CUENTA completo y no expone
 * ULTIMOS_DIGITOS -- forzarla aqui hubiera requerido derivar
 * ULTIMOS_DIGITOS del numero de cuenta, y ya se verifico con datos
 * reales que NO son los ultimos digitos literales del numero completo.
 *
 * Ya no filtra por ULTIMOS_DIGITOS no vacio (antes si) -- una sociedad
 * puede tener filas de cuenta sin ese dato y aun asi ser resoluble por
 * NOMBRE_SOCIEDAD; el filtro viejo las hubiera descartado sin razon
 * para ese caso. No cambia el comportamiento de la busqueda por codigo
 * (ultimosDigitos vacio nunca coincide con un codigo real, que siempre
 * son digitos). */
function _leerCatCuentasPorUltimosDigitos(ss) {
  var sheet = ss.getSheetByName('CAT_CUENTAS_MAPEO');
  if (!sheet || sheet.getLastRow() < 2) return [];
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h || '').trim().toUpperCase(); });
  var iUlt = headers.indexOf('ULTIMOS_DIGITOS');
  var iCorto = headers.indexOf('NOMBRE_CORTO');
  var iSoc = headers.indexOf('NOMBRE_SOCIEDAD');
  if (iCorto < 0) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues()
    .filter(function (row) { return String(row[iCorto] || '').trim(); })
    .map(function (row) {
      return {
        ultimosDigitos: iUlt >= 0 ? String(row[iUlt] || '').trim() : '',
        nombreCorto: String(row[iCorto] || '').trim(),
        nombreSociedad: iSoc >= 0 ? String(row[iSoc] || '').trim() : ''
      };
    });
}

/** El nombre trae el codigo pegado al final (ej. "Blau Life SAPI de CV
 * 8332") -- se extrae el numero final y se busca contra ULTIMOS_DIGITOS.
 * Nunca inventa una sociedad si no hay match.
 *
 * Fallback por nombre completo (2026-08-02, pedido explicito del
 * usuario, caso real "Soccer Loco Holding" -- llega SIN numero al final
 * en el Excel real, a diferencia del resto de sociedades). Aplica a
 * CUALQUIER propuesta, no solo a esa sociedad: si "Nombre empresa" no
 * trae ningun numero al final, se compara contra NOMBRE_SOCIEDAD
 * (trim() de extremos + mayusculas -- NO colapsa espacios dobles
 * internos ni normaliza acentos; si eso llega a fallar para una
 * sociedad nueva, cae al mismo mensaje "no reconocida" de siempre, no
 * en silencio). Si SI trae numero pero
 * ese numero no matchea ningun ULTIMOS_DIGITOS, NO cae a este fallback
 * -- eso seguiria siendo "no reconocida" como antes (un numero
 * equivocado es una señal real de error, no debe enmascararse
 * intentando adivinar por nombre). Una sociedad real puede tener varias
 * filas/cuentas en el catalogo (ej. Soccer Loco: MXN/USD/nomina) --
 * todas comparten el mismo NOMBRE_CORTO, asi que basta la primera
 * coincidencia. */
function _resolverSociedadDesdeNombreEmpresa(nombreEmpresa, catCuentas) {
  var nombreLimpio = String(nombreEmpresa || '').trim();
  if (!nombreLimpio) return null;

  var m = nombreLimpio.match(/(\d+)\s*$/);
  if (m) {
    var codigo = m[1];
    var matchPorCodigo = catCuentas.filter(function (c) { return c.ultimosDigitos === codigo; })[0];
    return matchPorCodigo ? matchPorCodigo.nombreCorto : null;
  }

  var nombreUpper = nombreLimpio.toUpperCase();
  var matchPorNombre = catCuentas.filter(function (c) { return c.nombreSociedad && c.nombreSociedad.toUpperCase() === nombreUpper; })[0];
  return matchPorNombre ? matchPorNombre.nombreCorto : null;
}

/** Mismo saneo que _limpiarMonto en Svc_Parser_Movimientos.js (no
 * reusable directamente: vive dentro del closure privado de
 * ParserMovimientos, no se expone en su API publica) -- necesario porque
 * XLSX.utils.sheet_to_csv exporta los importes con el formato de celda
 * de Excel (ej. "12,500.00"), y Utilities.parseCsv conserva la coma
 * dentro del campo. parseFloat("12,500.00") solo, sin sanear, devuelve
 * 12 -- trunca el monto real en silencio (hallazgo real del revisor). */
function _limpiarImporte(v) {
  if (v === null || v === undefined || v === '' || v === '-') return 0;
  if (typeof v === 'number') return isNaN(v) ? 0 : v;
  var s = String(v).trim().replace(/[$,\s()']/g, '');
  var n = parseFloat(s.replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? 0 : n;
}

var MONEDAS_VALIDAS = ['MXN', 'USD', 'GBP', 'EUR'];

function _uuidsExistentesEnPartidas(ss) {
  var sheet = _ensurePartidasPagoSheet(ss);
  if (sheet.getLastRow() < 2) return [];
  var col = PARTIDAS_PAGO_HEADERS.indexOf('UUID') + 1;
  return sheet.getRange(2, col, sheet.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; }).filter(function (v) { return v; });
}

/** sheets = [{nombre, csv}] -- mismo formato que ya arma el frontend con
 * SheetJS (readFileAsAllSheets) para carga-mov. No escribe nada, solo
 * arma la vista previa: propuestas candidatas agrupadas por sociedad
 * resuelta, mas errores por fila (nunca se descarta una fila en
 * silencio). */
function previsualizarPropuestaPago(sheets) {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var catCuentas = _leerCatCuentasPorUltimosDigitos(ss);
    var uuidsExistentes = _uuidsExistentesEnPartidas(ss);

    var hojaValida = null;
    (sheets || []).forEach(function (s) {
      if (hojaValida) return;
      var filas = Utilities.parseCsv(s.csv);
      if (filas.length && _esHojaPropuestaConsolidada(filas[0])) hojaValida = { nombre: s.nombre, filas: filas };
    });
    if (!hojaValida) {
      return { status: 'error', message: 'Ninguna hoja del archivo coincide con el formato esperado (columnas "Nombre empresa" / "Proveedor" / "Importe MXN"). Hojas en el archivo: ' + (sheets || []).map(function (s) { return s.nombre; }).join(', ') + '.' };
    }

    var headers = hojaValida.filas[0].map(function (h) { return String(h).trim(); });
    var hi = {}; headers.forEach(function (h, i) { hi[h] = i; });
    var faltantes = PROPUESTA_HEADERS_REQUERIDOS.filter(function (h) { return hi[h] === undefined; });
    if (faltantes.length) {
      return { status: 'error', message: 'La hoja "' + hojaValida.nombre + '" no trae las columnas: ' + faltantes.join(', ') + '. Encabezados encontrados: ' + headers.join(', ') };
    }

    var gruposPorSociedad = {};
    var ordenSociedades = [];
    var erroresFila = [];
    var filasOmitidasSubtotal = 0;
    var totalFilasValidas = 0;

    hojaValida.filas.slice(1).forEach(function (fila, i) {
      if (fila.every(function (v) { return v === ''; })) return;
      var numFilaExcel = i + 2;

      var nombreEmpresa = fila[hi['Nombre empresa']];
      var proveedor = fila[hi['Proveedor']];
      var nombreProveedor = fila[hi['Nombre del Proveedor']];
      var esSubtotal = (!nombreEmpresa && !proveedor) || /TOTAL/i.test(nombreProveedor || '');
      if (esSubtotal) { filasOmitidasSubtotal++; return; }

      var mxn = _limpiarImporte(fila[hi['Importe MXN']]);
      var usd = _limpiarImporte(fila[hi['Importe USD']]);
      var gbp = _limpiarImporte(fila[hi['Importe GBP']]);
      // 'Importe EUR' es opcional (ver comentario de PROPUESTA_HEADERS_REQUERIDOS)
      // -- solo se lee si la columna existe en este archivo.
      var eur = hi['Importe EUR'] !== undefined ? _limpiarImporte(fila[hi['Importe EUR']]) : 0;
      var divisas = [];
      if (mxn) divisas.push({ moneda: 'MXN', monto: mxn });
      if (usd) divisas.push({ moneda: 'USD', monto: usd });
      if (gbp) divisas.push({ moneda: 'GBP', monto: gbp });
      if (eur) divisas.push({ moneda: 'EUR', monto: eur });

      if (divisas.length === 0) {
        erroresFila.push({ fila: numFilaExcel, motivo: 'Sin importe en MXN, USD, GBP ni EUR.' });
        return;
      }
      if (divisas.length > 1) {
        erroresFila.push({ fila: numFilaExcel, motivo: 'Trae importe en mas de una divisa a la vez (' + divisas.map(function (d) { return d.moneda; }).join(' + ') + ') -- no se puede determinar una sola moneda para esta linea.' });
        return;
      }

      var nombreCorto = _resolverSociedadDesdeNombreEmpresa(nombreEmpresa, catCuentas);
      if (!nombreCorto) {
        erroresFila.push({ fila: numFilaExcel, motivo: 'Sociedad no reconocida en CAT_CUENTAS_MAPEO: "' + nombreEmpresa + '".' });
        return;
      }

      var uuid = fila[hi['UUID']] || '';
      var linea = {
        tempId: Utilities.getUuid(),
        proveedor: nombreProveedor || '',
        numeroProveedor: proveedor || '', // codigo SAP del proveedor (columna "Proveedor"), distinto del nombre -- antes nunca se capturaba
        solicitante: fila[hi['Solicitante']] || '',
        desc: fila[hi['Texto']] || '',
        referencia: fila[hi['Referencia']] || '', // campo propio, ya no se mezcla con DESC (antes: Texto||Referencia perdia Referencia por completo si Texto venia lleno)
        monto: divisas[0].monto,
        moneda: divisas[0].moneda,
        uuid: uuid,
        rfc: fila[hi['RFC']] || '',
        numeroDocumento: fila[hi['Nº documento']] || '',
        duplicado: !!uuid && uuidsExistentes.indexOf(uuid) >= 0,
        filaExcel: numFilaExcel
      };

      if (!gruposPorSociedad[nombreCorto]) { gruposPorSociedad[nombreCorto] = []; ordenSociedades.push(nombreCorto); }
      gruposPorSociedad[nombreCorto].push(linea);
      totalFilasValidas++;
    });

    var propuestasPreview = ordenSociedades.sort().map(function (soc) {
      return { sociedad: soc, lineas: gruposPorSociedad[soc] };
    });

    return {
      status: 'success',
      data: {
        hoja: hojaValida.nombre,
        totalFilasValidas: totalFilasValidas,
        filasOmitidasSubtotal: filasOmitidasSubtotal,
        propuestasPreview: propuestasPreview,
        erroresFila: erroresFila
      },
      message: ''
    };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function _generarFolioUnico(propSheet, nombreCorto, fechaStr) {
  var prefijo = nombreCorto + '-' + fechaStr + '-';
  var existentes = [];
  if (propSheet.getLastRow() >= 2) {
    var folioCol = PROPUESTAS_PAGO_HEADERS.indexOf('FOLIO') + 1;
    existentes = propSheet.getRange(2, folioCol, propSheet.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; });
  }
  var n = 1, folio;
  do {
    folio = prefijo + ('0' + n).slice(-2);
    n++;
  } while (existentes.indexOf(folio) >= 0);
  return folio;
}

/** payload = { query }. Busca personas en el DIRECTORIO de Google
 * Workspace de la organizacion (no los contactos personales de Gmail de
 * quien busca -- decision explicita del usuario, Bloque E 2026-07-27,
 * tras explicarle la diferencia de alcance de permisos). Usa el servicio
 * avanzado People (`People.People.searchDirectoryPeople`, habilitado en
 * appsscript.json) -- requiere el permiso OAuth de solo lectura del
 * directorio del dominio, que Apps Script detecta solo (no se declaro
 * `oauthScopes` a mano en el manifest a proposito, para no arriesgar
 * romper los permisos ya usados por el resto del proyecto -- ver
 * appsscript.json). Nunca lanza: si el permiso no esta autorizado, la
 * cuota se agoto, o el directorio esta desactivado a nivel Workspace
 * (fuera del control de este codigo), se degrada a "sin resultados" en
 * vez de tronar la pantalla de carga. Minimo 3 caracteres para buscar
 * (evita golpear la API con cada tecla sin filtro real). */
function buscarContactosDirectorio(payload) {
  var query = String((payload && payload.query) || '').trim();
  if (query.length < 3) return { status: 'success', data: [], message: '' };
  try {
    var res = People.People.searchDirectoryPeople({
      query: query,
      readMask: 'names,emailAddresses',
      sources: ['DIRECTORY_SOURCE_TYPE_DOMAIN_CONTACT', 'DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE']
    });
    var personas = (res && res.people) || [];
    var resultados = personas.map(function (p) {
      var email = (p.emailAddresses && p.emailAddresses[0] && p.emailAddresses[0].value) || '';
      var nombre = (p.names && p.names[0] && p.names[0].displayName) || '';
      return { email: email, nombre: nombre };
    }).filter(function (r) { return r.email; });
    return { status: 'success', data: resultados, message: '' };
  } catch (e) {
    Logger.log('buscarContactosDirectorio fallo (query="' + query + '"): ' + e.toString());
    return { status: 'error', data: [], message: 'No se pudo buscar en el directorio: ' + e.toString() };
  }
}

/** propuestasConfirmadas = [{ sociedad, lineas:[{proveedor, solicitante,
 * desc, monto, moneda, uuid, rfc, numeroDocumento}, ...] }, ...] -- ya
 * filtradas/editadas por el usuario en el preview (las filas con error o
 * los duplicados que decidio excluir nunca llegan aqui). Escritura en
 * bloque (setValues de un arreglo completo), no celda por celda.
 *
 * Defensas agregadas tras hallazgos del revisor (2 rondas):
 * - Revalidacion de MONTO/MONEDA aqui, no solo en el preview: el cliente
 *   podria mandar cualquier cosa a esta funcion directamente (el preview
 *   es una ayuda de UX, no un control de acceso). Si alguna linea no
 *   pasa, se rechaza TODA la carga con el detalle, no se crea nada
 *   parcial. El monto se normaliza a numero (`Number(l.monto)`) AQUI y
 *   esa es la version que se persiste despues -- si se guardara
 *   `l.monto` crudo y llegara como string, Apps Script lo escribe como
 *   texto literal en la celda, no como numero (rompe sumas/formulas).
 * - LockService: sin esto, dos cargas concurrentes para la misma
 *   sociedad el mismo dia podrian generar el mismo FOLIO o pisarse la
 *   escritura una a la otra (ambas leen el mismo getLastRow() antes de
 *   que la otra termine). `waitLock` vive DENTRO del try/catch (si no se
 *   consigue el lock en 30s, la excepcion debe seguir el mismo contrato
 *   {status:'error', message} que el resto del archivo, no propagarse
 *   cruda).
 * - Bloqueo de UUID duplicado server-side: el preview solo marca
 *   `duplicado` como aviso informativo para que el usuario decida
 *   excluir la linea -- no es un control de acceso. Sin este segundo
 *   chequeo aqui, la misma factura podria pagarse dos veces si el
 *   cliente no filtra lo marcado. Se rechaza TODA la carga si algun UUID
 *   no vacio ya existe en PARTIDAS_PAGO -- SALVO que `payload.forzarDuplicados`
 *   venga true con un `payload.motivoForzado` no vacio (Pieza 2,
 *   2026-07-28, pedido explicito del usuario: la MISMA persona que carga
 *   puede confirmar explicitamente "se que esto ya existe, cargalo de
 *   todas formas", nunca sin motivo -- se revalida el motivo tambien
 *   aqui, nunca confiar solo en que el cliente ya lo exigio). Cada linea
 *   duplicada que se fuerza queda registrada en OVERRIDE_DUPLICADOS_PAGO
 *   (quien, cuando, que UUID, por que) -- auditoria real, no solo un
 *   mensaje de exito. */
function confirmarCargaPropuestaPago(payload) {
  var propuestasConfirmadas = (payload || {}).propuestas;
  var forzarDuplicados = !!(payload || {}).forzarDuplicados;
  var motivoForzado = String((payload || {}).motivoForzado || '').trim();

  var erroresValidacion = [];
  (propuestasConfirmadas || []).forEach(function (grupo) {
    (grupo.lineas || []).forEach(function (l) {
      var montoNum = Number(l.monto);
      if (isNaN(montoNum) || montoNum <= 0) erroresValidacion.push('Linea "' + (l.proveedor || '?') + '" (' + grupo.sociedad + '): monto invalido (' + l.monto + ').');
      else l.monto = montoNum; // persistir el numero normalizado, no el valor crudo
      if (MONEDAS_VALIDAS.indexOf(l.moneda) < 0) erroresValidacion.push('Linea "' + (l.proveedor || '?') + '" (' + grupo.sociedad + '): moneda invalida (' + l.moneda + ').');
    });
  });
  if (erroresValidacion.length) {
    return { status: 'error', message: 'No se creo nada -- ' + erroresValidacion.length + ' linea(s) con datos invalidos: ' + erroresValidacion.join(' ') };
  }

  var lock = LockService.getScriptLock();
  // Bloque D (2026-07-27): notificacion de carga disparada DESPUES de
  // liberar el lock (mismo patron ya usado en aplicarPagoParcial) -- leer
  // hojas y mandar correo no debe mantener el lock tomado mas tiempo del
  // necesario, y _notificarTransicionPagos nunca lanza de todas formas.
  var idsParaNotificar = null, emailActorNotificar = '';
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var propSheet = _ensurePropuestasPagoSheet(ss);
    var lineSheet = _ensurePartidasPagoSheet(ss);

    var uuidsExistentes = _uuidsExistentesEnPartidas(ss);
    var uuidsDuplicados = [];
    (propuestasConfirmadas || []).forEach(function (grupo) {
      (grupo.lineas || []).forEach(function (l) {
        if (l.uuid && uuidsExistentes.indexOf(l.uuid) >= 0) uuidsDuplicados.push(l.uuid + ' (' + (l.proveedor || '?') + ', ' + grupo.sociedad + ')');
      });
    });
    if (uuidsDuplicados.length && !forzarDuplicados) {
      return { status: 'error', message: 'No se creo nada -- ' + uuidsDuplicados.length + ' linea(s) con UUID ya cargado antes: ' + uuidsDuplicados.join('; ') + '. Excluyelas en el preview, o confirma que quieres forzar la carga con un motivo.' };
    }
    if (uuidsDuplicados.length && forzarDuplicados && !motivoForzado) {
      return { status: 'error', message: 'Para forzar la carga de línea(s) con UUID ya existente hace falta un motivo.' };
    }

    var ahora = new Date();
    var fechaStr = Utilities.formatDate(ahora, ss.getSpreadsheetTimeZone(), 'yyyyMMdd');

    var nuevasFilasProp = [];
    var nuevasFilasPart = [];
    var resumen = [];
    // Bloque D (2026-07-27): se guardan los IDs de las partidas recien
    // creadas para notificar la carga por correo DESPUES de escribir con
    // exito (ver mas abajo) -- antes el ID se generaba inline dentro del
    // .map() sin quedar accesible fuera de el.
    var idsNuevasPartidas = [];
    // Bloque E (2026-07-27): CC extraordinario por propuesta, opcional --
    // grupo.ccExtra = [{email, nombre}], tipicamente 0-1. Se acumulan
    // filas para CC_EXTRA_PROPUESTA y se escriben junto con lo demas,
    // dentro del mismo lock (el propId se genera aqui mismo, no hace
    // falta un roundtrip aparte).
    var nuevasFilasCcExtra = [];
    // Pieza 2 (2026-07-28): una fila por cada linea forzada (UUID ya
    // existente, cargada de todas formas por decision explicita del
    // usuario) -- reusa lineaId/propId ya generados en el loop de abajo,
    // mismo patron que nuevasFilasCcExtra.
    var nuevasFilasOverride = [];
    var emailQuienCarga = Session.getActiveUser().getEmail();

    (propuestasConfirmadas || []).forEach(function (grupo) {
      if (!grupo.lineas || !grupo.lineas.length) return;
      var propId = Utilities.getUuid();
      var folio = _generarFolioUnico(propSheet, grupo.sociedad, fechaStr);

      nuevasFilasProp.push(PROPUESTAS_PAGO_HEADERS.map(function (h) {
        if (h === 'ID') return propId;
        if (h === 'FOLIO') return folio;
        if (h === 'SOCIEDAD') return grupo.sociedad;
        if (h === 'CONCEPTO') return 'Propuesta de pago ' + grupo.sociedad + ' - ' + fechaStr;
        if (h === 'FECHA_CARGA') return ahora;
        return '';
      }));

      (grupo.ccExtra || []).forEach(function (c) {
        var email = String(c.email || '').trim();
        if (!email) return;
        nuevasFilasCcExtra.push(CC_EXTRA_PROPUESTA_HEADERS.map(function (h) {
          if (h === 'ID') return Utilities.getUuid();
          if (h === 'PROPUESTA_ID') return propId;
          if (h === 'EMAIL') return email;
          if (h === 'NOMBRE') return String(c.nombre || '').trim();
          if (h === 'AGREGADO_POR') return emailQuienCarga;
          if (h === 'AGREGADO_FECHA') return ahora;
          return '';
        }));
      });

      grupo.lineas.forEach(function (l) {
        var lineaId = Utilities.getUuid();
        idsNuevasPartidas.push(lineaId);
        if (l.uuid && uuidsExistentes.indexOf(l.uuid) >= 0) {
          nuevasFilasOverride.push(_filaOverrideDuplicado(propId, lineaId, l, grupo.sociedad, motivoForzado, emailQuienCarga, ahora));
        }
        nuevasFilasPart.push(PARTIDAS_PAGO_HEADERS.map(function (h) {
          if (h === 'ID') return lineaId;
          if (h === 'PROPUESTA_ID') return propId;
          if (h === 'PROVEEDOR') return l.proveedor || '';
          if (h === 'NUMERO_PROVEEDOR') return l.numeroProveedor || '';
          if (h === 'SOLICITANTE') return l.solicitante || '';
          if (h === 'DESC') return l.desc || '';
          if (h === 'REFERENCIA') return l.referencia || '';
          if (h === 'MONTO') return l.monto;
          if (h === 'MONEDA') return l.moneda;
          if (h === 'UUID') return l.uuid || '';
          if (h === 'RFC') return l.rfc || '';
          if (h === 'NUMERO_DOCUMENTO') return l.numeroDocumento || '';
          if (h === 'FASE') return 'contabilidad';
          if (h === 'ESTADO') return 'activa';
          if (h === 'REACTIVACION_SOLICITADA') return false;
          if (h === 'CHECKED') return false;
          if (h === 'SUBIDO') return false;
          if (h === 'PAGADO') return false;
          // Ninguno de los 5 pares POR/FECHA de auditoria por fase aplica
          // aqui: una partida recien importada nace en 'contabilidad' sin
          // haber sido revisada/autorizada/liberada/subida/aplicada
          // todavia -- se estampan solo cuando ocurre la transicion real
          // (ver _camposAuditoriaPorFase).
          return '';
        }));
      });

      resumen.push({ folio: folio, sociedad: grupo.sociedad, lineas: grupo.lineas.length });
    });

    if (!nuevasFilasProp.length) return { status: 'error', message: 'No hay propuestas validas para crear.' };

    propSheet.getRange(propSheet.getLastRow() + 1, 1, nuevasFilasProp.length, PROPUESTAS_PAGO_HEADERS.length).setValues(nuevasFilasProp);
    var filaInicioPart = lineSheet.getLastRow() + 1;
    var rangoNuevasPart = lineSheet.getRange(filaInicioPart, 1, nuevasFilasPart.length, PARTIDAS_PAGO_HEADERS.length);
    // Bug real 2026-07-27 (propuesta "Ventura" invisible + crash de
    // appendChild): codigos SAP puramente numericos en NUMERO_PROVEEDOR
    // se escribian sin formato de texto forzado, asi que Sheets los
    // reinterpretaba como Number al leerlos de vuelta -- _pagoSheetToObjects
    // no los convertia a String, y VEVA.dom.el truena al pasarle un Number
    // como hijo del DOM. Se fuerza texto ANTES de escribir para que la
    // celda nunca se autoconvierta. Mismo riesgo aplica a REFERENCIA
    // (tambien puede llegar como codigo puramente numerico del Excel).
    lineSheet.getRange(filaInicioPart, 4, nuevasFilasPart.length, 1).setNumberFormat('@');
    lineSheet.getRange(filaInicioPart, 7, nuevasFilasPart.length, 1).setNumberFormat('@');
    rangoNuevasPart.setValues(nuevasFilasPart);
    if (nuevasFilasCcExtra.length) {
      var ccExtraSheet = _ensureCcExtraPropuestaSheet(ss);
      ccExtraSheet.getRange(ccExtraSheet.getLastRow() + 1, 1, nuevasFilasCcExtra.length, CC_EXTRA_PROPUESTA_HEADERS.length).setValues(nuevasFilasCcExtra);
    }
    if (nuevasFilasOverride.length) {
      var overrideSheet = _ensureOverrideDuplicadosPagoSheet(ss);
      overrideSheet.getRange(overrideSheet.getLastRow() + 1, 1, nuevasFilasOverride.length, OVERRIDE_DUPLICADOS_PAGO_HEADERS.length).setValues(nuevasFilasOverride);
    }

    // Bloque D (2026-07-27, pedido explicito del usuario -- "no veo que
    // se envien los correos" al cargar): la escritura ya tuvo exito, se
    // marca para notificar en el finally (ver arriba) -- si el correo
    // falla, el negocio ya quedo guardado igual.
    idsParaNotificar = idsNuevasPartidas;
    emailActorNotificar = emailQuienCarga;

    return { status: 'success', data: { propuestasCreadas: resumen }, message: resumen.length + ' propuesta(s) creada(s), ' + nuevasFilasPart.length + ' partida(s) en total.' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  } finally {
    lock.releaseLock();
    if (idsParaNotificar && idsParaNotificar.length) {
      _notificarTransicionPagos(ss, idsParaNotificar, TRANSICION_CARGADA, emailActorNotificar);
    }
  }
}

// =====================================================================
// Conciliacion bancaria (fase 6) + Comprobantes (fase 7) -- 2026-07-26.
// Ver PROMPT_TESORERIA360_PAGOS_1.md seccion 3.3 para el detalle
// completo de las decisiones (match por partida, split 1:N, multi-
// moneda via CUENTA_PAGO/MONEDA_PAGO_REAL/TC_APLICADO en PARTIDAS_PAGO,
// gate de saldo informativo no bloqueante, N:1 fuera de alcance).
// =====================================================================

/** Catalogo de cuentas para los selects del frontend (CUENTA_PAGO en la
 * pestaña Aplicado, agrupacion Sociedad->Banco en Conciliacion). Reusa
 * _buildCatCuentasArray (Codigo.js) tal cual -- no se duplica el
 * catalogo. */
function getCatalogoCuentasPago() {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    return { status: 'success', data: _buildCatCuentasArray(ss), message: '' };
  } catch (e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

/** Lee HIST_MOVIMIENTOS crudo filtrando EGRESO + cuenta + rango de
 * fecha, incluyendo MONEDA_ORIGINAL/TC_APLICADO (columnas 23/24, indices
 * 22/23) -- getMovimientos (Codigo.js) no expone esas 2 columnas en su
 * salida mapeada, y aqui son necesarias para matchear por moneda de
 * origen real del movimiento, no solo por monto. */
function _leerMovimientosEgresoRaw(ss, cuentaNorm, desde, hasta) {
  var sh = ss.getSheetByName('HIST_MOVIMIENTOS');
  if (!sh || sh.getLastRow() < 2) return [];
  var raw = sh.getRange(2, 1, sh.getLastRow() - 1, _HIST_COLS).getValues();
  var out = [];
  raw.forEach(function (r) {
    if (String(r[12] || '').trim() !== 'EGRESO') return;
    if (_normalizarCuenta(r[4]) !== cuentaNorm) return;
    var fecha = _parseFechaLocal(r[1]);
    if (!fecha) return;
    if (desde && fecha < desde) return;
    if (hasta && fecha > hasta) return;
    var cargo = parseFloat(r[6]) || 0;
    if (cargo <= 0) return;
    out.push({
      idMov: String(r[0] || ''),
      fecha: Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
      descripcion: String(r[5] || ''),
      cargo: cargo,
      moneda: String(r[22] || 'MXN').trim() || 'MXN'
    });
  });
  return out;
}

/** Info minima (fecha/descripcion/moneda) de un set de ID_MOV, para
 * pintar los vinculos ya conciliados sin que el frontend tenga que pedir
 * HIST_MOVIMIENTOS aparte. */
function _infoMovimientosPorId(ss, ids) {
  var out = {};
  if (!ids.length) return out;
  var sh = ss.getSheetByName('HIST_MOVIMIENTOS');
  if (!sh || sh.getLastRow() < 2) return out;
  var wanted = {};
  ids.forEach(function (id) { wanted[id] = true; });
  var raw = sh.getRange(2, 1, sh.getLastRow() - 1, _HIST_COLS).getValues();
  raw.forEach(function (r) {
    var id = String(r[0] || '');
    if (!wanted[id]) return;
    var fecha = r[1] instanceof Date ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(r[1] || '').slice(0, 10);
    out[id] = { fecha: fecha, descripcion: String(r[5] || ''), moneda: String(r[22] || 'MXN').trim() || 'MXN' };
  });
  return out;
}

/** Gate de saldo (badge informativo, NUNCA bloqueante): por cuenta, toma
 * el periodo mas reciente de SUM_MOV que ya tenga SALDO_REPORTADO
 * capturado (mismas columnas que ya usa getConciliacion en Codigo.js).
 * Si una cuenta no tiene ninguna fila con saldo reportado, se omite --
 * el frontend la muestra sin badge, nunca como "no cuadra" por falta de
 * dato. */
function _saldoMasRecientePorCuenta(ss) {
  var sh = ss.getSheetByName('SUM_MOV');
  var out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  var raw = sh.getRange(2, 1, sh.getLastRow() - 1, _SUM_COLS).getValues();
  raw.forEach(function (r) {
    var cuenta = _normalizarCuenta(r[0]);
    if (!cuenta) return;
    var reportado = r[16];
    if (reportado === '' || reportado === null) return;
    var periodo = r[3] instanceof Date ? Utilities.formatDate(r[3], Session.getScriptTimeZone(), 'yyyy-MM') : String(r[3] || '').slice(0, 7);
    if (!out[cuenta] || periodo > out[cuenta].periodo) {
      out[cuenta] = { periodo: periodo, saldoFinalCalc: parseFloat(r[12]) || 0, saldoReportado: parseFloat(reportado) || 0 };
    }
  });
  return out;
}

/** Tablero de conciliacion: todas las partidas fase='aplicado' && pagado,
 * con sus vinculos actuales (CONCILIACION_PAGOS), agrupables en el
 * frontend por Sociedad->Banco/Cuenta->Conciliadas/No-conciliadas (el
 * agrupado en si vive en la UI, aqui solo se entregan los datos ya
 * resueltos: banco, estado, montos). */
function getConciliacionPagos() {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', data: { partidas: [], saldoPorCuenta: {} }, message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var propRows = _pagoSheetToObjects(_ensurePropuestasPagoSheet(ss), PROPUESTAS_PAGO_HEADERS);
    var propById = {};
    propRows.forEach(function (p) { propById[p.id] = p; });

    // Bloque B (2026-07-27): la conciliacion ya NO espera a que la
    // partida quede 100% pagada -- cualquier partida en fase 'aplicado'
    // con AL MENOS una liberacion parcial ya es conciliable, para que
    // Tesoreria pueda cuadrar contra el banco cada exhibicion segun va
    // saliendo, sin esperar a la ultima. `montoEfectivo` (mismo nombre de
    // campo que antes, semantica nueva) pasa a ser el ACUMULADO aplicado
    // hasta ahora -- el objetivo VIGENTE de conciliacion, que crece con
    // cada aplicacion parcial nueva. `montoEfectivoTotal` (campo nuevo) es
    // el objetivo FINAL completo, para que la UI muestre progreso.
    var acumuladoPorPartida = _acumuladoAplicadoPorPartida(ss);

    var lineRows = _pagoSheetToObjects(_ensurePartidasPagoSheet(ss), PARTIDAS_PAGO_HEADERS)
      .filter(function (l) { return l.fase === 'aplicado' && (acumuladoPorPartida[l.id] || 0) > 0; });

    var vinculos = _pagoSheetToObjects(_ensureConciliacionPagosSheet(ss), CONCILIACION_PAGOS_HEADERS);
    var vincPorPartida = {};
    vinculos.forEach(function (v) {
      if (!vincPorPartida[v.partidaId]) vincPorPartida[v.partidaId] = [];
      vincPorPartida[v.partidaId].push(v);
    });
    var idsMovNecesarios = {};
    vinculos.forEach(function (v) { idsMovNecesarios[v.idMov] = true; });
    var movInfo = _infoMovimientosPorId(ss, Object.keys(idsMovNecesarios));

    var catCuentas = _buildCatCuentasArray(ss);
    var cuentaInfo = {};
    catCuentas.forEach(function (c) { cuentaInfo[c.cuenta] = c; });
    var saldoPorCuenta = _saldoMasRecientePorCuenta(ss);

    var partidas = lineRows.map(function (l) {
      var prop = propById[l.propuestaId] || {};
      var vincs = (vincPorPartida[l.id] || []).map(function (v) {
        var info = movInfo[v.idMov] || {};
        return { idMov: v.idMov, montoAplicado: v.montoAplicado, fecha: info.fecha || '', descripcion: info.descripcion || '', moneda: info.moneda || '' };
      });
      var montoEfectivoTotal = _montoEfectivoPartida(l);
      var montoEfectivo = acumuladoPorPartida[l.id] || 0; // objetivo vigente (Bloque B)
      var sumaVinculada = vincs.reduce(function (s, v) { return s + (Number(v.montoAplicado) || 0); }, 0);
      var conciliada = vincs.length > 0 && Math.abs(sumaVinculada - montoEfectivo) < 1;
      var cuentaNorm = _normalizarCuenta(l.cuentaPago);
      var cta = cuentaInfo[cuentaNorm] || {};
      return {
        id: l.id, propuestaId: l.propuestaId, folio: prop.folio || '', sociedad: prop.sociedad || '',
        proveedor: l.proveedor, monto: Number(l.monto), moneda: l.moneda,
        monedaPagoReal: l.monedaPagoReal || '', tcAplicado: l.tcAplicado || '',
        montoEfectivo: montoEfectivo, montoEfectivoTotal: montoEfectivoTotal, aplicadoFecha: l.aplicadoFecha || '',
        cuentaPago: cuentaNorm, banco: cta.banco || '', tipoCuenta: cta.tipoCuenta || '',
        vinculos: vincs, sumaVinculada: sumaVinculada,
        estadoConciliacion: !cuentaNorm ? 'sin_cuenta' : (conciliada ? (vincs.length > 1 ? 'split' : 'conciliada') : 'pendiente')
      };
    });

    return { status: 'success', data: { partidas: partidas, saldoPorCuenta: saldoPorCuenta }, message: '' };
  } catch (e) {
    return { status: 'error', data: { partidas: [], saldoPorCuenta: {} }, message: e.toString() };
  }
}

/** payload = { partidaId }. Candidatos de HIST_MOVIMIENTOS para conciliar
 * una partida: mismo CUENTA_PAGO, TIPO=EGRESO, fecha dentro de la ventana
 * de aplicaciones +/- TOLERANCIA_DIAS (CONFIG_CONCILIACION), monto dentro
 * de TOLERANCIA_MONTO_PCT del monto ACUMULADO aplicado (Bloque B), y
 * misma moneda de origen (MONEDA_PAGO_REAL si existe, si no la MONEDA
 * nominal de la partida). Excluye movimientos ya vinculados a OTRA
 * partida (el caso N:1 esta fuera de alcance de este incremento); los
 * vinculos de esta MISMA partida se incluyen marcados yaVinculado para
 * que el frontend los muestre pre-seleccionados al reabrir el panel de
 * match.
 *
 * Ventana de fecha (Bloque B, 2026-07-27): antes se derivaba de
 * PARTIDAS_PAGO.APLICADO_FECHA, que solo se estampa cuando la partida
 * llega al 100% -- con pagos parciales eso queda vacio hasta el final,
 * dejando sin poder buscar candidatos la primera liberacion parcial (la
 * que mas los necesita, para conciliar de inmediato). Ahora se deriva
 * del rango min/max de FECHA en APLICACIONES_PAGO de esta partida. */
function buscarCandidatosMovimiento(payload) {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', data: { candidatos: [] }, message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var partidas = _pagoSheetToObjects(_ensurePartidasPagoSheet(ss), PARTIDAS_PAGO_HEADERS);
    var partida = partidas.filter(function (p) { return p.id === payload.partidaId; })[0];
    if (!partida) return { status: 'error', message: 'Partida no encontrada: ' + payload.partidaId };
    if (!partida.cuentaPago) return { status: 'error', data: { candidatos: [] }, message: 'Esta partida no tiene cuenta de pago capturada (pestaña Aplicado).' };

    var aplicaciones = _aplicacionesPorPartida(ss, partida.id);
    if (!aplicaciones.length) return { status: 'error', data: { candidatos: [] }, message: 'Esta partida todavía no tiene ningún pago aplicado.' };
    var montoBuscado = aplicaciones.reduce(function (s, a) { return s + (Number(a.monto) || 0); }, 0);
    var fechasAplicacion = aplicaciones.map(function (a) { return _parseFechaLocal(String(a.aplicadoFecha).slice(0, 10)); }).filter(Boolean);

    var cfg = _leerConfigConciliacion(ss);
    var fechaMin = new Date(Math.min.apply(null, fechasAplicacion.map(function (f) { return f.getTime(); })));
    var fechaMax = new Date(Math.max.apply(null, fechasAplicacion.map(function (f) { return f.getTime(); })));
    var desde = new Date(fechaMin.getTime());
    desde.setDate(desde.getDate() - cfg.toleranciaDias);
    var hasta = new Date(fechaMax.getTime());
    hasta.setDate(hasta.getDate() + cfg.toleranciaDias);
    hasta.setHours(23, 59, 59, 999);

    var monedaBuscada = partida.monedaPagoReal || partida.moneda;
    var margen = montoBuscado * (cfg.toleranciaMontoPct / 100);

    var cuentaNorm = _normalizarCuenta(partida.cuentaPago);
    var movs = _leerMovimientosEgresoRaw(ss, cuentaNorm, desde, hasta);

    var vinculos = _pagoSheetToObjects(_ensureConciliacionPagosSheet(ss), CONCILIACION_PAGOS_HEADERS);
    var usadosPorOtra = {}, vinculadosAqui = {};
    vinculos.forEach(function (v) {
      if (v.partidaId === partida.id) vinculadosAqui[v.idMov] = v.montoAplicado;
      else usadosPorOtra[v.idMov] = true;
    });

    var candidatos = movs.filter(function (m) {
      if (usadosPorOtra[m.idMov]) return false;
      if (vinculadosAqui[m.idMov] !== undefined) return true;
      if (monedaBuscada && m.moneda !== monedaBuscada) return false;
      return Math.abs(m.cargo - montoBuscado) <= margen;
    }).map(function (m) {
      return {
        idMov: m.idMov, fecha: m.fecha, descripcion: m.descripcion, monto: m.cargo, moneda: m.moneda,
        yaVinculado: vinculadosAqui[m.idMov] !== undefined,
        montoAplicadoActual: vinculadosAqui[m.idMov] !== undefined ? vinculadosAqui[m.idMov] : m.cargo
      };
    });

    // Hallazgo real del revisor: `movs` esta acotado a la ventana de
    // fecha/cuenta ACTUAL (recalculada con CUENTA_PAGO/APLICADO_FECHA/
    // CONFIG_CONCILIACION de hoy). Si un vinculo ya confirmado quedo
    // fuera de esa ventana (ej. alguien redujo TOLERANCIA_DIAS despues
    // de conciliar), no aparecia en absoluto -- y como conciliarPartida
    // reemplaza TODOS los vinculos existentes por lo que el usuario deja
    // seleccionado, reabrir el panel y volver a confirmar podia borrar en
    // silencio un match bancario ya validado. Se fuerza a que TODO
    // vinculo ya confirmado de esta partida aparezca, sin importar si
    // cae dentro de la ventana/tolerancia vigente.
    var idsYaEnCandidatos = {};
    candidatos.forEach(function (c) { idsYaEnCandidatos[c.idMov] = true; });
    var idsFaltantes = Object.keys(vinculadosAqui).filter(function (id) { return !idsYaEnCandidatos[id]; });
    if (idsFaltantes.length) {
      var infoFaltantes = _infoMovimientosPorId(ss, idsFaltantes);
      idsFaltantes.forEach(function (id) {
        var info = infoFaltantes[id] || {};
        candidatos.push({
          idMov: id, fecha: info.fecha || '', descripcion: (info.descripcion || '') + ' (fuera de la ventana de búsqueda actual)',
          monto: vinculadosAqui[id], moneda: info.moneda || '',
          yaVinculado: true, montoAplicadoActual: vinculadosAqui[id]
        });
      });
    }

    return { status: 'success', data: { candidatos: candidatos, montoBuscado: montoBuscado, monedaBuscada: monedaBuscada, tolerancia: cfg }, message: '' };
  } catch (e) {
    return { status: 'error', data: { candidatos: [] }, message: e.toString() };
  }
}

function _borrarVinculosDePartida(vincSheet, partidaId) {
  if (vincSheet.getLastRow() < 2) return;
  var idCol = CONCILIACION_PAGOS_HEADERS.indexOf('PARTIDA_ID') + 1;
  var vals = vincSheet.getRange(2, idCol, vincSheet.getLastRow() - 1, 1).getValues();
  for (var i = vals.length - 1; i >= 0; i--) {
    if (vals[i][0] === partidaId) vincSheet.deleteRow(i + 2);
  }
}

/** payload = { partidaId, movimientos:[{idMov, montoAplicado}] }. Valida
 * suma vs monto ACUMULADO aplicado hasta ahora (Bloque B -- no el objetivo
 * final completo, ver aplicarPagoParcial/getConciliacionPagos; tolerancia
 * < $1, fija, no configurable -- solo la BUSQUEDA de candidatos usa la
 * tolerancia configurable), valida que ningun idMov este ya conciliado
 * contra OTRA partida, y reemplaza los vinculos existentes de esta
 * partida (permite ajustar la seleccion sin dejar residuos huerfanos)
 * bajo LockService. */
function conciliarPartida(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_puedeConciliarOComprobantes(ss)) return { status: 'error', message: 'No tienes permiso para conciliar pagos.' };

    var partidas = _pagoSheetToObjects(_ensurePartidasPagoSheet(ss), PARTIDAS_PAGO_HEADERS);
    var partida = partidas.filter(function (p) { return p.id === payload.partidaId; })[0];
    if (!partida) return { status: 'error', message: 'Partida no encontrada: ' + payload.partidaId };

    var movimientos = payload.movimientos || [];
    if (!movimientos.length) return { status: 'error', message: 'Selecciona al menos un movimiento.' };

    var montoBuscado = _aplicacionesPorPartida(ss, partida.id).reduce(function (s, a) { return s + (Number(a.monto) || 0); }, 0);
    var suma = movimientos.reduce(function (s, m) { return s + (Number(m.montoAplicado) || 0); }, 0);
    if (Math.abs(suma - montoBuscado) >= 1) {
      return { status: 'error', message: 'La suma seleccionada ($' + suma.toFixed(2) + ') no cuadra contra el monto de la partida ($' + montoBuscado.toFixed(2) + ') -- la diferencia debe ser menor a $1.' };
    }

    var vincSheet = _ensureConciliacionPagosSheet(ss);
    var existentes = _pagoSheetToObjects(vincSheet, CONCILIACION_PAGOS_HEADERS);
    var usadosPorOtra = existentes.filter(function (v) { return v.partidaId !== partida.id; }).map(function (v) { return v.idMov; });
    var conflicto = movimientos.filter(function (m) { return usadosPorOtra.indexOf(m.idMov) >= 0; });
    if (conflicto.length) {
      return { status: 'error', message: 'Movimiento(s) ya conciliado(s) contra otra partida: ' + conflicto.map(function (m) { return m.idMov; }).join(', ') + '.' };
    }

    _borrarVinculosDePartida(vincSheet, partida.id);

    var email = Session.getActiveUser().getEmail();
    var ahora = new Date();
    var nuevasFilas = movimientos.map(function (m) {
      return CONCILIACION_PAGOS_HEADERS.map(function (h) {
        if (h === 'ID') return Utilities.getUuid();
        if (h === 'PARTIDA_ID') return partida.id;
        if (h === 'ID_MOV') return m.idMov;
        if (h === 'MONTO_APLICADO') return Number(m.montoAplicado) || 0;
        if (h === 'CONCILIADO_POR') return email;
        if (h === 'CONCILIADO_FECHA') return ahora;
        return '';
      });
    });
    vincSheet.getRange(vincSheet.getLastRow() + 1, 1, nuevasFilas.length, CONCILIACION_PAGOS_HEADERS.length).setValues(nuevasFilas);

    return { status: 'success', message: 'Partida conciliada contra ' + nuevasFilas.length + ' movimiento(s).' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/** payload = { partidaId, idMov? }. Si se manda idMov borra solo ese
 * vinculo (desconciliar un movimiento del split); si no, borra TODOS los
 * vinculos de la partida. */
function desconciliarPartida(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_puedeConciliarOComprobantes(ss)) return { status: 'error', message: 'No tienes permiso para desconciliar pagos.' };
    var vincSheet = _ensureConciliacionPagosSheet(ss);
    if (vincSheet.getLastRow() < 2) return { status: 'success', message: '' };
    var idPCol = CONCILIACION_PAGOS_HEADERS.indexOf('PARTIDA_ID') + 1;
    var idMCol = CONCILIACION_PAGOS_HEADERS.indexOf('ID_MOV') + 1;
    var vals = vincSheet.getRange(2, 1, vincSheet.getLastRow() - 1, CONCILIACION_PAGOS_HEADERS.length).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (vals[i][idPCol - 1] === payload.partidaId && (!payload.idMov || vals[i][idMCol - 1] === payload.idMov)) {
        vincSheet.deleteRow(i + 2);
      }
    }
    return { status: 'success', message: '' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/** Tabla plana para la pestaña Comprobantes: partidas fase='aplicado' &&
 * pagado, con sus comprobantes vinculados (0, 1, o compartiendo uno
 * consolidado con otras partidas -- ver totalPartidas). Independiente de
 * conciliacion a proposito (una partida puede tener comprobante sin
 * estar conciliada, y viceversa). */
function getComprobantesPago() {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var propRows = _pagoSheetToObjects(_ensurePropuestasPagoSheet(ss), PROPUESTAS_PAGO_HEADERS);
    var propById = {};
    propRows.forEach(function (p) { propById[p.id] = p; });

    var lineRows = _pagoSheetToObjects(_ensurePartidasPagoSheet(ss), PARTIDAS_PAGO_HEADERS)
      .filter(function (l) { return l.fase === 'aplicado' && l.pagado === true; });

    var sheets = _ensureComprobantesPagoSheets(ss);
    var comprobantes = _pagoSheetToObjects(sheets.comprobantes, COMPROBANTES_PAGO_HEADERS);
    var compById = {};
    comprobantes.forEach(function (c) { compById[c.id] = c; });

    var vinculos = _pagoSheetToObjects(sheets.vinculos, COMPROBANTES_PAGO_PARTIDAS_HEADERS);
    var vincPorPartida = {};
    var partidasPorComprobante = {};
    vinculos.forEach(function (v) {
      if (!vincPorPartida[v.partidaId]) vincPorPartida[v.partidaId] = [];
      vincPorPartida[v.partidaId].push(v);
      partidasPorComprobante[v.comprobanteId] = (partidasPorComprobante[v.comprobanteId] || 0) + 1;
    });

    var partidas = lineRows.map(function (l) {
      var prop = propById[l.propuestaId] || {};
      var vincs = (vincPorPartida[l.id] || []).map(function (v) {
        var c = compById[v.comprobanteId] || {};
        return {
          comprobanteId: v.comprobanteId, archivoNombre: c.archivoNombre || '', driveUrl: c.driveUrl || '',
          totalPartidas: partidasPorComprobante[v.comprobanteId] || 1
        };
      });
      return {
        id: l.id, folio: prop.folio || '', sociedad: prop.sociedad || '',
        proveedor: l.proveedor, desc: l.desc, monto: Number(l.monto), moneda: l.moneda,
        aplicadoFecha: l.aplicadoFecha || '', comprobantes: vincs
      };
    });

    return { status: 'success', data: partidas, message: '' };
  } catch (e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

/** Techo practico del payload base64 de google.script.run (el limite
 * real de HTML Service ronda ~50MB; se deja margen para el resto del
 * payload + overhead de base64, que infla el tamaño ~33%). */
var COMPROBANTE_MAX_BYTES = 15 * 1024 * 1024;

/** payload = { archivoNombre, base64, mimeType, partidaIds }. Sube UNA
 * vez a Drive (carpeta fija COMPROBANTES_DRIVE_FOLDER_ID) y vincula a
 * TODAS las partidas seleccionadas de un jalon -- caso real: el banco
 * entrega un PDF consolidado con resumen + varios pagos individuales
 * adentro, identificar el monto/pagina exacto de cada uno DENTRO del PDF
 * queda fuera de este incremento (iteracion futura). */
function subirComprobante(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_puedeConciliarOComprobantes(ss)) return { status: 'error', message: 'No tienes permiso para subir comprobantes.' };

    var partidaIds = payload.partidaIds || [];
    if (!partidaIds.length) return { status: 'error', message: 'Selecciona al menos una partida antes de subir el comprobante.' };
    if (!payload.base64) return { status: 'error', message: 'Archivo vacío.' };

    var bytes = Utilities.base64Decode(payload.base64);
    if (bytes.length > COMPROBANTE_MAX_BYTES) {
      return { status: 'error', message: 'El archivo pesa ' + (bytes.length / 1024 / 1024).toFixed(1) + 'MB, el máximo es ' + (COMPROBANTE_MAX_BYTES / 1024 / 1024) + 'MB.' };
    }

    var blob = Utilities.newBlob(bytes, payload.mimeType || 'application/pdf', payload.archivoNombre || 'comprobante');
    var folder = DriveApp.getFolderById(COMPROBANTES_DRIVE_FOLDER_ID);
    var file = folder.createFile(blob);

    var sheets = _ensureComprobantesPagoSheets(ss);
    var email = Session.getActiveUser().getEmail();
    var ahora = new Date();
    var comprobanteId = Utilities.getUuid();

    sheets.comprobantes.getRange(sheets.comprobantes.getLastRow() + 1, 1, 1, COMPROBANTES_PAGO_HEADERS.length).setValues([
      COMPROBANTES_PAGO_HEADERS.map(function (h) {
        if (h === 'ID') return comprobanteId;
        if (h === 'ARCHIVO_NOMBRE') return payload.archivoNombre || file.getName();
        if (h === 'DRIVE_FILE_ID') return file.getId();
        if (h === 'DRIVE_URL') return file.getUrl();
        if (h === 'MIME_TYPE') return payload.mimeType || '';
        if (h === 'SUBIDO_POR') return email;
        if (h === 'SUBIDO_FECHA') return ahora;
        return '';
      })
    ]);

    var filasVinculo = partidaIds.map(function (pid) {
      return COMPROBANTES_PAGO_PARTIDAS_HEADERS.map(function (h) {
        if (h === 'ID') return Utilities.getUuid();
        if (h === 'COMPROBANTE_ID') return comprobanteId;
        if (h === 'PARTIDA_ID') return pid;
        if (h === 'VINCULADO_POR') return email;
        if (h === 'VINCULADO_FECHA') return ahora;
        return '';
      });
    });
    sheets.vinculos.getRange(sheets.vinculos.getLastRow() + 1, 1, filasVinculo.length, COMPROBANTES_PAGO_PARTIDAS_HEADERS.length).setValues(filasVinculo);

    return { status: 'success', data: { comprobanteId: comprobanteId, driveUrl: file.getUrl() }, message: 'Comprobante vinculado a ' + partidaIds.length + ' partida(s).' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/** payload = { comprobanteId, partidaId }. Borra SOLO la fila de vinculo
 * -- nunca el archivo de Drive ni la fila de COMPROBANTES_PAGO (otras
 * partidas pueden seguir usando el mismo comprobante consolidado). */
function desvincularComprobante(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_puedeConciliarOComprobantes(ss)) return { status: 'error', message: 'No tienes permiso para desvincular comprobantes.' };

    var sheets = _ensureComprobantesPagoSheets(ss);
    var shVinc = sheets.vinculos;
    if (shVinc.getLastRow() < 2) return { status: 'success', message: '' };
    var idCCol = COMPROBANTES_PAGO_PARTIDAS_HEADERS.indexOf('COMPROBANTE_ID') + 1;
    var idPCol = COMPROBANTES_PAGO_PARTIDAS_HEADERS.indexOf('PARTIDA_ID') + 1;
    var vals = shVinc.getRange(2, 1, shVinc.getLastRow() - 1, COMPROBANTES_PAGO_PARTIDAS_HEADERS.length).getValues();
    for (var i = vals.length - 1; i >= 0; i--) {
      if (vals[i][idCCol - 1] === payload.comprobanteId && vals[i][idPCol - 1] === payload.partidaId) {
        shVinc.deleteRow(i + 2);
      }
    }
    return { status: 'success', message: '' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// =====================================================================
// Notificaciones por correo (Bloque 1) -- 2026-07-26. Lista de
// distribucion configurable POR SOCIEDAD (CONFIG_NOTIFICACIONES_PAGOS),
// correo agrupado (nunca uno por partida) cuando una partida cambia de
// fase/estado. Ver PROJECT_STATUS.md para el detalle de las 3 decisiones
// de diseño (transiciones, agrupacion, pantalla aparte).
// =====================================================================

/** Orquesta el envio de correo tras una mutacion exitosa de PARTIDAS_PAGO
 * (o, desde Bloque D, tras la creacion de partidas nuevas al cargar una
 * propuesta). Responsabilidad separada de _updatePartidaPagoRow a
 * proposito (SRP): esa funcion solo muta la fila, esta decide si hace
 * falta avisar y a quien. Nunca debe romper la transaccion de negocio ya
 * persistida -- cualquier falla aqui adentro (lectura de hojas, MailApp)
 * se atrapa y solo se registra en Logger, jamas se propaga hacia el
 * llamador (requisito de .claude/DEFINICION_DE_TERMINADO.md, seccion
 * Automatizacion). Se llama DESPUES de que la escritura ya tuvo exito --
 * si el correo falla, el negocio ya quedo guardado igual.
 *
 * `transicion` (2026-07-27, Bloque D): antes este parametro era `patch`
 * y la funcion computaba `_tipoTransicionNotificable(patch)` internamente
 * -- eso no funcionaba para la carga inicial, que no es un patch sobre
 * una partida existente. Ahora el llamador computa la transicion (via
 * `_tipoTransicionNotificable(patch)` para mutaciones reales, o la
 * constante `TRANSICION_CARGADA` para la carga) y la pasa ya resuelta. */
function _notificarTransicionPagos(ss, idsExitosos, transicion, emailActor) {
  if (!transicion || !idsExitosos || !idsExitosos.length) return;
  _notificarTransicionPagosMultiple(ss, [{ idsExitosos: idsExitosos, transicion: transicion }], emailActor);
}

/** Bloque F (2026-07-28): variante que acepta VARIOS grupos con
 * transiciones DISTINTAS en la misma llamada (ej. positivas=LIBERADA A
 * BANCO, negativas=RECHAZADA desde bulkResolverBanca) para que el correo
 * de cada sociedad salga en UN solo mensaje con 2 secciones, nunca 2
 * correos separados. grupos = [{idsExitosos, transicion}, ...], en el
 * orden que deben aparecer las secciones (positivas primero). Los
 * llamadores de un solo grupo siguen usando _notificarTransicionPagos
 * (arriba), que ahora delega aqui -- comportamiento identico a antes de
 * este bloque para todos ellos. */
function _notificarTransicionPagosMultiple(ss, grupos, emailActor) {
  try {
    grupos = (grupos || []).filter(function (g) { return g.transicion && g.idsExitosos && g.idsExitosos.length; });
    if (!grupos.length) return;

    var propRows = _pagoSheetToObjects(_ensurePropuestasPagoSheet(ss), PROPUESTAS_PAGO_HEADERS);
    var propById = {};
    propRows.forEach(function (p) { propById[p.id] = p; });

    var lineRows = _pagoSheetToObjects(_ensurePartidasPagoSheet(ss), PARTIDAS_PAGO_HEADERS);
    var lineById = {};
    lineRows.forEach(function (l) { lineById[l.id] = l; });

    // Bloque E (2026-07-27): el CC extraordinario es POR PROPUESTA, no
    // por sociedad -- si el lote de una sociedad mezcla partidas de 2+
    // propuestas y solo alguna tiene CC extra, NUNCA debe filtrarse ese
    // CC hacia partidas de una propuesta ajena. Se sub-agrupa cada
    // sociedad por "firma de CC extra" (lista de emails extra de la
    // propuesta de cada linea, ordenada y unida): si ninguna propuesta
    // del lote tiene CC extra, todas caen en la firma '' -- sigue siendo
    // 1 solo correo, comportamiento identico a antes de este bloque.
    var ccExtraMap = _ccExtraPorPropuestaMap(ss);

    // porSociedad[sociedad][firmaCc] = { ccExtra, secciones:[{transicion,items}] }
    // -- agrupado por sociedad (2026-07-26, decidido con el usuario: un
    // lote de la misma sociedad manda 1 correo con la lista completa, no
    // 1 correo por partida) Y ADEMAS por transicion (Bloque F), para que
    // cada transicion del lote quede en su propia seccion dentro de ese
    // mismo correo, en vez de mezclarse en una sola tabla.
    var porSociedad = {};
    grupos.forEach(function (grupo) {
      grupo.idsExitosos.forEach(function (id) {
        var l = lineById[id];
        if (!l) return;
        var prop = propById[l.propuestaId];
        var sociedad = prop ? prop.sociedad : '';
        if (!sociedad) return;
        var ccExtra = ccExtraMap[l.propuestaId] || [];
        var firma = ccExtra.slice().sort().join(',');
        if (!porSociedad[sociedad]) porSociedad[sociedad] = {};
        if (!porSociedad[sociedad][firma]) porSociedad[sociedad][firma] = { ccExtra: ccExtra, secciones: [] };
        var secciones = porSociedad[sociedad][firma].secciones;
        var seccion = secciones.filter(function (s) { return s.transicion === grupo.transicion; })[0];
        if (!seccion) { seccion = { transicion: grupo.transicion, items: [] }; secciones.push(seccion); }
        seccion.items.push({ linea: l, folio: prop ? prop.folio : '', propuestaId: l.propuestaId });
      });
    });

    Object.keys(porSociedad).forEach(function (sociedad) {
      Object.keys(porSociedad[sociedad]).forEach(function (firma) {
        var grupoFirma = porSociedad[sociedad][firma];
        // Try/catch por sub-grupo (ademas del try/catch general de
        // afuera): que falle el correo de UN sub-grupo (ej. lista mal
        // capturada) no debe impedir que los demas del mismo lote
        // reciban el suyo.
        try {
          _enviarCorreoNotificacionPagoSeccionado(ss, sociedad, grupoFirma.secciones, emailActor, grupoFirma.ccExtra);
        } catch (eSociedad) {
          Logger.log('Notificacion de pagos fallo para sociedad ' + sociedad + ': ' + eSociedad.toString());
        }
      });
    });
  } catch (e) {
    Logger.log('Notificacion de pagos fallo (no afecta la mutacion ya guardada): ' + e.toString());
  }
}

// Fase de auditoria a mostrar en el correo, UNA sola por transicion (no
// todas las acumuladas -- decision explicita del usuario, 2026-07-31:
// cada correo ya es por una transicion especifica, mostrar solo esa
// columna es mas compacto y legible en Outlook/movil que mostrar las 5
// fases siempre). Mismos nombres camelCase que produce _pagoSheetToObjects
// a partir de PARTIDAS_PAGO_HEADERS. 'cargado' (TRANSICION_CARGADA) no
// tiene POR/FECHA propio en PARTIDAS_PAGO -- queda sin entrada a proposito,
// _textoFaseTransicion devuelve '' para ese caso.
var _FASE_AUDITORIA_POR_TRANSICION = {
  revisado: { label: 'Revisado por', por: 'revisadoPor', fecha: 'revisadoFecha' },
  autorizado: { label: 'Autorizado por', por: 'autorizadoPor', fecha: 'autorizadoFecha' },
  rechazado: { label: 'Rechazado por', por: 'rechazadoPor', fecha: 'rechazadoFecha' },
  liberado: { label: 'Liberado a banco por', por: 'liberadoPor', fecha: 'liberadoFecha' },
  subido: { label: 'Subido a banco por', por: 'subidoPor', fecha: 'subidoFecha' },
  aplicado: { label: 'Aplicado por', por: 'aplicadoPor', fecha: 'aplicadoFecha' }
};

/** "Autorizado por: Juan Perez (26/07/2026)" -- solo la fase de la
 * transicion que motiva ESTE correo. Cadena vacia si esa fase no tiene
 * POR/FECHA definido (cargado) o la partida no lo tiene lleno todavia. */
function _textoFaseTransicion(l, transicion) {
  var f = _FASE_AUDITORIA_POR_TRANSICION[transicion.tipo];
  if (!f || !l[f.por]) return '';
  var fecha = l[f.fecha] ? Utilities.formatDate(new Date(l[f.fecha]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '';
  return f.label + ': ' + l[f.por] + (fecha ? ' (' + fecha + ')' : '');
}

/** Filas HTML de la tabla del correo (Folio/Proveedor/Referencia/UUID/
 * Monto + fila secundaria compacta con Descripcion + fase) -- compartida
 * por _enviarCorreoNotificacionPagoUnica y
 * _enviarCorreoNotificacionPagoSeccionado para no duplicar en 2 lugares
 * el fix del monto (usa _montoEfectivoPartida, respeta autorizacion
 * parcial -- antes mostraba Number(l.monto) crudo, ignorando parciales)
 * ni las columnas nuevas pedidas por el usuario (Referencia, UUID). */
/** Hallazgo real de `qa` (Grupo 3): las celdas de la tabla del correo
 * interpolaban proveedor/referencia/uuid/desc sin escapar -- un nombre de
 * proveedor real con "&" o "<" (ej. "Johnson & Johnson") rompe el HTML
 * del correo. Preexistente en `proveedor` desde antes de este grupo, pero
 * ahora se cierra para las 5 celdas nuevas/tocadas de una vez. */
function _escapeHtmlCorreo(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}
function _filasTablaCorreo(items, transicion) {
  return items.map(function (it) {
    var l = it.linea;
    var monedaMostrar = l.monedaPagoReal || l.moneda || '';
    var montoMostrar = _montoEfectivoPartida(l);
    var filaPrincipal = '<tr>'
      + '<td style="padding:10px 4px 2px;font-size:13px;color:#0a2540;border-top:1px solid #f6f9fc;">' + _escapeHtmlCorreo(it.folio) + '</td>'
      + '<td style="padding:10px 4px 2px;font-size:13px;color:#0a2540;border-top:1px solid #f6f9fc;">' + _escapeHtmlCorreo(l.proveedor) + '</td>'
      + '<td style="padding:10px 4px 2px;font-size:11px;color:#697386;border-top:1px solid #f6f9fc;">' + _escapeHtmlCorreo(l.referencia) + '</td>'
      + '<td style="padding:10px 4px 2px;font-size:11px;color:#697386;border-top:1px solid #f6f9fc;font-family:monospace;">' + _escapeHtmlCorreo(l.uuid) + '</td>'
      + '<td style="padding:10px 4px 2px;font-size:13px;font-weight:600;text-align:right;color:#0a2540;font-variant-numeric:tabular-nums;border-top:1px solid #f6f9fc;">' + _escapeHtmlCorreo(monedaMostrar) + ' ' + montoMostrar.toLocaleString('es-MX', { minimumFractionDigits: 2 }) + '</td>'
      + '</tr>';
    var detalle = [l.desc, _textoFaseTransicion(l, transicion)].filter(Boolean).join(' · ');
    if (!detalle) return filaPrincipal;
    return filaPrincipal + '<tr><td colspan="5" style="padding:0 4px 10px;font-size:11px;color:#94a3b8;">' + _escapeHtmlCorreo(detalle) + '</td></tr>';
  }).join('');
}

/** "MXN $12,345.00 · USD $500.00" -- subtotal por moneda de pago real
 * (o moneda propia si no hay conversion), usando _montoEfectivoPartida
 * (respeta parciales) igual que las filas. Pedido explicito del usuario:
 * ver de un vistazo cuanto se mueve por moneda en cada correo. */
function _subtotalTablaCorreo(items) {
  var map = {};
  items.forEach(function (it) {
    var l = it.linea;
    var m = l.monedaPagoReal || l.moneda || 'SIN_MONEDA';
    map[m] = (map[m] || 0) + _montoEfectivoPartida(l);
  });
  return Object.keys(map).sort().map(function (m) {
    return m + ' ' + map[m].toLocaleString('es-MX', { minimumFractionDigits: 2 });
  }).join(' · ');
}

/** Lee la lista de distribucion ACTIVA de la sociedad y manda un solo
 * correo agrupado con el detalle de las partidas afectadas. Si nadie
 * configuro destinatarios para esa sociedad todavia, no manda nada --
 * no es un error, es un estado valido (nadie lo ha configurado).
 *
 * Rediseño 2026-07-27 (pedido explicito del usuario): marca "Sistema de
 * Pagos" (ya no "TESO Tesorería 360"), asunto
 * PROPUESTA DE PAGOS "SOCIEDAD" FECHA "DD/MM/AAAA" <ETIQUETA>, bloque de
 * auditoria por partida (quien hizo cada fase ya completada) y link de
 * vuelta al sistema. El link es a la URL BASE de la app (no un deep-link
 * a la sociedad/partida especifica -- eso requeriria routing por hash,
 * que hoy no existe en Shared_Core.html; decidido con el usuario dejarlo
 * fuera de este bloque). Se obtiene con ScriptApp.getService().getUrl()
 * en vez de hardcodear el deployment ID, para que nunca quede stale si
 * se vuelve a desplegar.
 *
 * `ccExtraEmails` (Bloque E, 2026-07-27, opcional, default []): emails
 * extraordinarios de la propuesta especifica que este lote de `items`
 * representa (ya resueltos por el llamador -- esta funcion no vuelve a
 * consultar CC_EXTRA_PROPUESTA). Se combinan con el CC ESTANDAR de la
 * sociedad (CONFIG_CC_PAGOS, aplica a los 7 tipos de correo sin
 * excepcion) y se deduplican antes de armar el `cc` final. Si la
 * combinacion queda vacia, no se manda ningun `cc` (mismo criterio que
 * ya usa `boton` para degradarse sin romper el envio). */
/** Bloque F (2026-07-28): despachador seccionado -- si viene 1 sola
 * seccion (el caso normal, todos los llamadores de antes de este
 * bloque), delega a _enviarCorreoNotificacionPagoUnica SIN TOCAR una
 * sola linea de su HTML ya revisado/probado en produccion. Si vienen 2+
 * secciones (bulkResolverBanca/bulkResolverAplicado con positivas Y
 * negativas en el mismo lote), arma UN solo correo con un bloque de
 * tabla por seccion, en el orden recibido -- nunca 2 correos separados.
 * secciones = [{transicion, items}, ...]. El asunto usa la etiqueta de
 * la PRIMERA seccion (el llamador siempre pone las positivas primero). */
function _enviarCorreoNotificacionPagoSeccionado(ss, sociedad, secciones, emailActor, ccExtraEmails) {
  if (secciones.length === 1) {
    return _enviarCorreoNotificacionPagoUnica(ss, sociedad, secciones[0].transicion, secciones[0].items, emailActor, ccExtraEmails);
  }

  var lista = _pagoSheetToObjects(_ensureConfigNotificacionesPagoSheet(ss), CONFIG_NOTIFICACIONES_PAGOS_HEADERS)
    .filter(function (d) { return d.sociedad === sociedad && _esActivo(d.activo); });
  var destinatarios = lista.map(function (d) { return d.email; }).filter(function (e) { return e; });
  if (!destinatarios.length) return;

  var listaCc = _pagoSheetToObjects(_ensureConfigCcPagosSheet(ss), CONFIG_CC_PAGOS_HEADERS)
    .filter(function (d) { return d.sociedad === sociedad && _esActivo(d.activo); })
    .map(function (d) { return d.email; });
  var copiasCrudas = listaCc.concat(ccExtraEmails || []).filter(function (e) { return e; });
  var copias = copiasCrudas.filter(function (e, i) { return copiasCrudas.indexOf(e) === i; });

  var urlApp = '';
  try { urlApp = ScriptApp.getService().getUrl(); } catch (eUrl) { /* correo sale sin boton si getUrl() falla */ }

  // Mismo hallazgo de Outlook/Word ya documentado abajo (flex/gap no se
  // soporta) -- el pill usa margin-left en el span vecino, nunca flex.
  var bloquesSecciones = secciones.map(function (sec) {
    return '<div style="margin-bottom:20px;">'
      + '<div style="margin-bottom:8px;">'
      + '<span style="display:inline-block;background:#635bff14;color:#635bff;font-size:11px;font-weight:700;letter-spacing:0.03em;padding:3px 10px;border-radius:999px;">' + sec.transicion.etiqueta + '</span>'
      + '<span style="font-size:12px;color:#697386;margin-left:8px;">' + sec.items.length + ' partida(s) ' + sec.transicion.verbo + '.</span>'
      + '</div>'
      + '<table style="width:100%;border-collapse:collapse;">'
      + '<thead><tr>'
      + '<th style="padding:0 4px 8px;text-align:left;font-size:10px;color:#94a3b8;font-weight:700;text-transform:uppercase;letter-spacing:0.03em;border-bottom:1px solid #e6ebf1;">Folio</th>'
      + '<th style="padding:0 4px 8px;text-align:left;font-size:10px;color:#94a3b8;font-weight:700;text-transform:uppercase;letter-spacing:0.03em;border-bottom:1px solid #e6ebf1;">Proveedor</th>'
      + '<th style="padding:0 4px 8px;text-align:left;font-size:10px;color:#94a3b8;font-weight:700;text-transform:uppercase;letter-spacing:0.03em;border-bottom:1px solid #e6ebf1;">Referencia</th>'
      + '<th style="padding:0 4px 8px;text-align:left;font-size:10px;color:#94a3b8;font-weight:700;text-transform:uppercase;letter-spacing:0.03em;border-bottom:1px solid #e6ebf1;">UUID</th>'
      + '<th style="padding:0 4px 8px;text-align:right;font-size:10px;color:#94a3b8;font-weight:700;text-transform:uppercase;letter-spacing:0.03em;border-bottom:1px solid #e6ebf1;">Monto</th>'
      + '</tr></thead>'
      + '<tbody>' + _filasTablaCorreo(sec.items, sec.transicion) + '</tbody>'
      + '<tfoot><tr>'
      + '<td colspan="4" style="padding:8px 4px 0;text-align:right;font-size:11px;font-weight:700;color:#0a2540;border-top:1px solid #e6ebf1;">Subtotal</td>'
      + '<td style="padding:8px 4px 0;text-align:right;font-size:11px;font-weight:700;color:#0a2540;border-top:1px solid #e6ebf1;">' + _subtotalTablaCorreo(sec.items) + '</td>'
      + '</tr></tfoot></table>'
      + '</div>';
  }).join('');

  var fechaHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  var fechaSubject = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var asunto = 'PROPUESTA DE PAGOS "' + sociedad + '" FECHA "' + fechaSubject + '" ' + secciones[0].transicion.etiqueta;

  var boton = urlApp
    ? '<div style="margin-top:4px;"><a href="' + urlApp + '" style="display:inline-block;background:#635bff;color:#fff;text-decoration:none;font-size:13px;font-weight:600;padding:10px 20px;border-radius:8px;">Abrir Sistema de Pagos</a></div>'
    : '';

  var html = '<div style="font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;max-width:600px;margin:0 auto;background:#fff;">'
    + '<div style="padding:24px 28px 0;">'
    + '<span style="font-size:11px;font-weight:700;letter-spacing:0.06em;color:#635bff;text-transform:uppercase;">Sistema de Pagos</span>'
    + '<h1 style="margin:8px 0 4px;font-size:19px;font-weight:700;color:#0a2540;">' + sociedad + '</h1>'
    + '<div style="margin-bottom:12px;"><span style="font-size:12px;color:#697386;">' + fechaSubject + '</span></div>'
    + '</div>'
    + '<div style="padding:0 24px 24px;">'
    + bloquesSecciones
    + boton
    + '</div>'
    + '<div style="padding:14px 28px;border-top:1px solid #e6ebf1;">'
    + '<p style="font-size:11px;color:#94a3b8;margin:0;">Realizado por ' + (emailActor || '') + ' — ' + fechaHora + '</p>'
    + '</div></div>';

  var payloadCorreo = { to: destinatarios.join(','), subject: asunto, htmlBody: html };
  if (copias.length) payloadCorreo.cc = copias.join(',');
  MailApp.sendEmail(payloadCorreo);
}

function _enviarCorreoNotificacionPagoUnica(ss, sociedad, transicion, items, emailActor, ccExtraEmails) {
  var lista = _pagoSheetToObjects(_ensureConfigNotificacionesPagoSheet(ss), CONFIG_NOTIFICACIONES_PAGOS_HEADERS)
    .filter(function (d) { return d.sociedad === sociedad && _esActivo(d.activo); });
  var destinatarios = lista.map(function (d) { return d.email; }).filter(function (e) { return e; });
  if (!destinatarios.length) return;

  var listaCc = _pagoSheetToObjects(_ensureConfigCcPagosSheet(ss), CONFIG_CC_PAGOS_HEADERS)
    .filter(function (d) { return d.sociedad === sociedad && _esActivo(d.activo); })
    .map(function (d) { return d.email; });
  var copiasCrudas = listaCc.concat(ccExtraEmails || []).filter(function (e) { return e; });
  var copias = copiasCrudas.filter(function (e, i) { return copiasCrudas.indexOf(e) === i; });

  var urlApp = '';
  try { urlApp = ScriptApp.getService().getUrl(); } catch (eUrl) { /* correo sale sin boton si getUrl() falla */ }

  var fechaHora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  var fechaSubject = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  var asunto = 'PROPUESTA DE PAGOS "' + sociedad + '" FECHA "' + fechaSubject + '" ' + transicion.etiqueta;

  var boton = urlApp
    ? '<div style="margin-top:20px;"><a href="' + urlApp + '" style="display:inline-block;background:#635bff;color:#fff;text-decoration:none;font-size:13px;font-weight:600;padding:10px 20px;border-radius:8px;">Abrir Sistema de Pagos</a></div>'
    : '';

  // Mismo estilo inline que enviarReportePosicion (Codigo.js) -- unico
  // precedente de correo real en el proyecto, no hay helper de plantilla
  // compartido que reusar todavia. Paleta identica a la del resto de la
  // app (VEVA / Index.html): #635bff acento, #0a2540 texto oscuro,
  // #697386 gris secundario, #e6ebf1 bordes, #f6f9fc fondo claro.
  var html = '<div style="font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;max-width:600px;margin:0 auto;background:#fff;">'
    + '<div style="padding:24px 28px 0;">'
    + '<span style="font-size:11px;font-weight:700;letter-spacing:0.06em;color:#635bff;text-transform:uppercase;">Sistema de Pagos</span>'
    + '<h1 style="margin:8px 0 4px;font-size:19px;font-weight:700;color:#0a2540;">' + sociedad + '</h1>'
    // Hallazgo real de revision con lente de usuario (2026-07-27):
    // display:flex + gap NO se soporta de forma confiable en Outlook de
    // escritorio (motor Word) -- el gap se ignora y el pill queda pegado
    // a la fecha sin espacio. Se reemplaza por margin-right en el propio
    // span, patron estandar de compatibilidad en HTML de correo (nunca
    // depender de flex/gap para separar elementos inline).
    + '<div style="margin-bottom:4px;">'
    + '<span style="display:inline-block;margin-right:8px;background:#635bff14;color:#635bff;font-size:11px;font-weight:700;letter-spacing:0.03em;padding:3px 10px;border-radius:999px;">' + transicion.etiqueta + '</span>'
    + '<span style="font-size:12px;color:#697386;">' + fechaSubject + '</span>'
    + '</div>'
    + '<p style="font-size:13px;color:#697386;margin:8px 0 0;">' + items.length + ' partida(s) ' + transicion.verbo + '.</p>'
    + '</div>'
    + '<div style="padding:16px 24px 24px;">'
    + '<table style="width:100%;border-collapse:collapse;">'
    + '<thead><tr>'
    + '<th style="padding:0 4px 8px;text-align:left;font-size:10px;color:#94a3b8;font-weight:700;text-transform:uppercase;letter-spacing:0.03em;border-bottom:1px solid #e6ebf1;">Folio</th>'
    + '<th style="padding:0 4px 8px;text-align:left;font-size:10px;color:#94a3b8;font-weight:700;text-transform:uppercase;letter-spacing:0.03em;border-bottom:1px solid #e6ebf1;">Proveedor</th>'
    + '<th style="padding:0 4px 8px;text-align:left;font-size:10px;color:#94a3b8;font-weight:700;text-transform:uppercase;letter-spacing:0.03em;border-bottom:1px solid #e6ebf1;">Referencia</th>'
    + '<th style="padding:0 4px 8px;text-align:left;font-size:10px;color:#94a3b8;font-weight:700;text-transform:uppercase;letter-spacing:0.03em;border-bottom:1px solid #e6ebf1;">UUID</th>'
    + '<th style="padding:0 4px 8px;text-align:right;font-size:10px;color:#94a3b8;font-weight:700;text-transform:uppercase;letter-spacing:0.03em;border-bottom:1px solid #e6ebf1;">Monto</th>'
    + '</tr></thead>'
    + '<tbody>' + _filasTablaCorreo(items, transicion) + '</tbody>'
    + '<tfoot><tr>'
    + '<td colspan="4" style="padding:8px 4px 0;text-align:right;font-size:11px;font-weight:700;color:#0a2540;border-top:1px solid #e6ebf1;">Subtotal</td>'
    + '<td style="padding:8px 4px 0;text-align:right;font-size:11px;font-weight:700;color:#0a2540;border-top:1px solid #e6ebf1;">' + _subtotalTablaCorreo(items) + '</td>'
    + '</tr></tfoot></table>'
    + boton
    + '</div>'
    + '<div style="padding:14px 28px;border-top:1px solid #e6ebf1;">'
    + '<p style="font-size:11px;color:#94a3b8;margin:0;">Realizado por ' + (emailActor || '') + ' — ' + fechaHora + '</p>'
    + '</div></div>';

  var payloadCorreo = { to: destinatarios.join(','), subject: asunto, htmlBody: html };
  if (copias.length) payloadCorreo.cc = copias.join(',');
  MailApp.sendEmail(payloadCorreo);
}

/** Todas las filas de todas las sociedades -- el frontend agrupa/filtra
 * por sociedad seleccionada. */
function getListaNotificacionesPago() {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var lista = _pagoSheetToObjects(_ensureConfigNotificacionesPagoSheet(ss), CONFIG_NOTIFICACIONES_PAGOS_HEADERS);
    return { status: 'success', data: lista, message: '' };
  } catch (e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

/** payload = { sociedad, lista:[{email,nombre,activo}] }. Reemplaza SOLO
 * las filas de esa sociedad -- nunca borra la configuracion de las
 * demas (a diferencia de guardarListaDistribucion en Codigo.js, que
 * reescribe la hoja completa porque esa lista es unica/global; aqui hay
 * multiples sociedades compartiendo la misma hoja). Gateado igual que
 * conciliar/subir comprobantes (Tesoreria+Contador) -- reusa
 * _puedeConciliarOComprobantes tal cual, no se inventa un gate paralelo. */
/** Reemplaza SOLO las filas de una sociedad en una hoja tipo "lista de
 * correos por sociedad" (shape SOCIEDAD/EMAIL/NOMBRE/ACTIVO/FECHA_ALTA)
 * -- helper compartido por guardarListaNotificacionesPago (el TO) y
 * guardarListaCcPagos (el CC, Bloque E) -- hallazgo real: ambas
 * funciones repetian esta logica identica, incluido el hallazgo previo
 * del revisor sobre preservar FECHA_ALTA por persona (nunca pisarla con
 * "ahora" en cada edicion de CUALQUIER destinatario de la sociedad).
 * Nunca borra la configuracion de otras sociedades. Retorna cuantas
 * filas quedaron para esa sociedad. */
function _guardarListaPorSociedad(sh, headers, sociedad, listaNueva) {
  var actuales = _pagoSheetToObjects(sh, headers);
  var deOtrasSociedades = actuales.filter(function (r) { return r.sociedad !== sociedad; });

  var ahora = new Date();
  var nuevasDeEstaSociedad = (listaNueva || []).map(function (item) {
    return { sociedad: sociedad, email: String(item.email || '').trim(), nombre: String(item.nombre || '').trim(), activo: item.activo !== false, fechaAlta: item.fechaAlta || ahora };
  }).filter(function (r) { return r.email; });

  var todas = deOtrasSociedades.concat(nuevasDeEstaSociedad);
  var filas = todas.map(function (r) {
    return headers.map(function (h) {
      var key = h.toLowerCase().replace(/_([a-z])/g, function (_, c) { return c.toUpperCase(); });
      return r[key] !== undefined ? r[key] : '';
    });
  });

  var lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  if (filas.length) sh.getRange(2, 1, filas.length, headers.length).setValues(filas);

  return nuevasDeEstaSociedad.length;
}

function guardarListaNotificacionesPago(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_puedeConciliarOComprobantes(ss)) return { status: 'error', message: 'No tienes permiso para editar la lista de notificaciones.' };
    if (!payload.sociedad) return { status: 'error', message: 'Falta la sociedad.' };

    var guardados = _guardarListaPorSociedad(_ensureConfigNotificacionesPagoSheet(ss), CONFIG_NOTIFICACIONES_PAGOS_HEADERS, payload.sociedad, payload.lista);
    return { status: 'success', data: { guardados: guardados }, message: guardados + ' destinatario(s) guardados para ' + payload.sociedad + '.' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

/** getListaCcPagos/guardarListaCcPagos (Bloque E, 2026-07-27): mismo
 * patron exacto que getListaNotificacionesPago/guardarListaNotificacionesPago
 * -- unica diferencia real es la hoja/headers (CONFIG_CC_PAGOS en vez de
 * CONFIG_NOTIFICACIONES_PAGOS) y que esta lista es CC, no TO. Mismo gate
 * (Tesoreria+Contador, _puedeConciliarOComprobantes) por consistencia
 * con el resto del modulo -- no se inventa un gate paralelo. */
function getListaCcPagos() {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };
    var lista = _pagoSheetToObjects(_ensureConfigCcPagosSheet(ss), CONFIG_CC_PAGOS_HEADERS);
    return { status: 'success', data: lista, message: '' };
  } catch (e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

function guardarListaCcPagos(payload) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_puedeConciliarOComprobantes(ss)) return { status: 'error', message: 'No tienes permiso para editar la lista de CC de pagos.' };
    if (!payload.sociedad) return { status: 'error', message: 'Falta la sociedad.' };

    var guardados = _guardarListaPorSociedad(_ensureConfigCcPagosSheet(ss), CONFIG_CC_PAGOS_HEADERS, payload.sociedad, payload.lista);
    return { status: 'success', data: { guardados: guardados }, message: guardados + ' contacto(s) en copia guardados para ' + payload.sociedad + '.' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  } finally {
    lock.releaseLock();
  }
}

// Las funciones temporales de diagnostico/reset (_debugInspeccionarHojasPago,
// _debugResetHojasPago) que vivieron aqui durante la depuracion del
// 2026-07-26 se retiraron por hallazgo del revisor: una funcion que borra
// PROPUESTAS_PAGO/PARTIDAS_PAGO por completo no debe quedar invocable de
// forma permanente en produccion, aunque sea con nombre "_debug". La
// migracion de esquema de esta ronda se hizo borrando las hojas a mano
// desde Google Sheets (accion normal de usuario, no codigo desplegado).
//
// CORRECCION 2026-07-26 (incidente post-despliegue de conciliacion +
// comprobantes): el parrafo de arriba resulto INCOMPLETO/FALSO para
// PARTIDAS_PAGO -- la hoja real en produccion todavia tenia el esquema
// DE ANTES de la migracion "en medio" descrita aqui (le faltaban
// NUMERO_PROVEEDOR/REFERENCIA, sobraban ULTIMA_ACCION_POR/FECHA). No se
// sabe si la migracion a mano nunca se completo del todo o se revirtio
// despues al editar la hoja. Se corrigio con una migracion automatica
// nueva (_migrarEsquemaMedioPartidasPago) en vez de repetir el patron de
// "borrar y recrear a mano", porque para entonces ya habia partidas
// reales cargadas que no se podian perder. Ver GUIA_ANTIALUCINACION_
// ANTIESPAGUETI_1.md §A.9: "codigo verificado contra el servidor" nunca
// implica "esquema de datos verificado" -- son cosas distintas.
//
// Las funciones temporales de reset (limpiarPagos, limpiarConciliacionPagos,
// marcador TEMPORAL_PRUEBAS_2026-07-27) que vivieron aqui durante la ronda
// de pruebas del 2026-07-27 se retiraron antes del deploy final de esa
// ronda, decision explicita del usuario -- mismo precedente que
// _debugResetHojasPago arriba. Backup del bloque completo (por si se
// necesita reconstruir para otra ronda de pruebas) en el scratchpad de la
// sesion que las retiro, no en el repo.

// La unica moneda con referencia diaria real hoy es USD (hoja TC, una
// columna de valor). Nunca se reusa ese valor para EUR/GBP -- seria
// mostrar un TC de mercado equivocado en un reporte financiero. Ver
// Grupo 5 (2026-07-31): decision explicita del usuario de dejar el caso
// EUR/GBP sin conversion (raro en la practica) en blanco en vez de
// construir una hoja de referencia diaria nueva para esas 2 monedas.
var TC_DIA_MONEDAS_SOPORTADAS = ['USD'];

/** Igual que _getTcForDate (Codigo.js), pero para este reporte
 * especificamente -- NO se reusa esa funcion compartida porque su
 * contrato (usado por saldos/movimientos) devuelve 1 cuando no hay
 * ningun TC registrado en o antes de la fecha pedida, y aqui eso se
 * mostraria como "Monto MXN = Monto M.E." sin ninguna marca de que es
 * un dato faltante, no una tasa real (hallazgo real de revisor).
 * null = sin dato, nunca se inventa una tasa. */
function _tcDelDiaOrNull(ss, fechaDate) {
  var tcSheet = ss.getSheetByName('TC');
  if (!tcSheet || tcSheet.getLastRow() <= 1) return null;
  var data = tcSheet.getRange(2, 1, tcSheet.getLastRow() - 1, 2).getValues();
  var tc = null, bestTime = -1;
  var targetTime = fechaDate.getTime();
  data.forEach(function (row) {
    var d = row[0] instanceof Date ? row[0] : new Date(row[0]);
    if (isNaN(d.getTime())) return;
    var t = d.getTime();
    var val = parseFloat(row[1]);
    if (t <= targetTime && t > bestTime && val > 0) { tc = val; bestTime = t; }
  });
  return tc;
}

/** Reporte historico de pagos (Grupo 5, 2026-07-31, pedido explicito del
 * usuario: consultar por dia todo lo realmente pagado). Fuente UNICA de
 * verdad: APLICACIONES_PAGO -- una fila por CADA exhibicion real de
 * dinero (total o parcial), confirmado en _aplicarPagoParcialCore
 * (linea ~1306): escribe ahi SIEMPRE, nunca solo en el caso parcial.
 * Enriquecida con PARTIDAS_PAGO (proveedor/moneda/cuenta/referencia/
 * uuid/solicitante/conversion) y PROPUESTAS_PAGO (folio/sociedad).
 * Nunca se calcula un total nuevo -- se suman montos ya persistidos.
 *
 * Logica de TC/monto, confirmada con el usuario (2026-07-31):
 * - MONTO en APLICACIONES_PAGO ya esta en la moneda de PAGO real
 *   (monedaPagoReal si hubo conversion, si no la moneda propia de la
 *   partida) -- es una porcion de _montoEfectivoPartida, que ya aplica
 *   esa conversion antes de repartirse en exhibiciones.
 * - Si hubo conversion real (monedaPagoReal==='MXN', unico destino
 *   posible -- el selector de View_Pagos.html solo ofrece moneda propia
 *   o MXN): el monto ya esta en MXN: "Monto MXN" = monto tal cual,
 *   "Monto M.E." se reconstruye dividiendo entre TC_APLICADO (el real,
 *   digitado por Tesoreria). "TC del dia" no aplica aqui (no hubo
 *   compra al TC de referencia, hubo compra al TC real capturado).
 * - Si NO hubo conversion y la moneda es extranjera (se pago desde una
 *   cuenta en esa misma moneda, ej. USD paga USD): el monto esta en esa
 *   moneda tal cual ("Monto M.E."). "TC del dia" (referencia) se calcula
 *   si esa moneda tiene referencia diaria real (TC_DIA_MONEDAS_SOPORTADAS
 *   -- hoy unicamente USD), solo para mostrarlo aparte como comparacion.
 *   "Monto MXN" da PRIORIDAD al TC real capturado por Tesoreria
 *   (TC_APLICADO, informativo -- se puede llenar aunque no haya
 *   conversion formal, ver View_Pagos.html linea ~932) sobre esa
 *   referencia; solo cae al "TC del dia" si no hay TC_APLICADO. Si
 *   ninguno de los dos existe, "Monto MXN" queda en blanco en vez de
 *   inventar una tasa (hallazgo real 2026-08-02: version anterior
 *   ignoraba TC_APLICADO aqui y usaba la referencia generica, produciendo
 *   un monto MXN incorrecto para partidas con TC informativo capturado).
 * - Si la moneda es MXN nativa: "Monto M.E." queda en blanco, "Monto
 *   MXN" es el monto tal cual, TC no aplica.
 *
 * payload = { fechaDesde, fechaHasta, sociedad, moneda, busqueda } --
 * todos opcionales, string. */
function getHistoricoPagos(payload) {
  try {
    var ss = SpreadsheetApp.openById(SALDOS_SHEET_ID);
    if (!_tieneAccesoValido(ss)) return { status: 'error', data: [], message: 'No tienes acceso a este módulo. Contacta a finanzas para que te den de alta en CAT_USUARIOS.' };

    var aplicaciones = _pagoSheetToObjects(_ensureAplicacionesPagoSheet(ss), APLICACIONES_PAGO_HEADERS);
    if (!aplicaciones.length) return { status: 'success', data: [], message: '' };

    var partidas = _pagoSheetToObjects(_ensurePartidasPagoSheet(ss), PARTIDAS_PAGO_HEADERS);
    var partidaPorId = {};
    partidas.forEach(function (p) { partidaPorId[p.id] = p; });

    var propuestas = _pagoSheetToObjects(_ensurePropuestasPagoSheet(ss), PROPUESTAS_PAGO_HEADERS);
    var propuestaPorId = {};
    propuestas.forEach(function (pr) { propuestaPorId[pr.id] = pr; });

    var catMap = _buildCatalogMap(ss);

    payload = payload || {};
    var filtroDesde = payload.fechaDesde ? _validateFecha(payload.fechaDesde) : null;
    var filtroHasta = payload.fechaHasta ? _validateFecha(payload.fechaHasta) : null;
    var filtroHastaExclusivo = filtroHasta ? new Date(filtroHasta.getFullYear(), filtroHasta.getMonth(), filtroHasta.getDate() + 1) : null;
    var filtroSociedad = String(payload.sociedad || '').trim();
    var filtroMoneda = String(payload.moneda || '').trim().toUpperCase();
    var filtroBusqueda = String(payload.busqueda || '').trim().toLowerCase();

    var filas = aplicaciones.map(function (a) {
      var p = partidaPorId[a.partidaId];
      if (!p) return null; // exhibicion huerfana (partida borrada/inexistente) -- no reventar el reporte por eso, solo omitirla
      var prop = propuestaPorId[p.propuestaId] || {};
      var fecha = a.aplicadoFecha ? new Date(a.aplicadoFecha) : null;

      return {
        fecha: fecha,
        sociedad: prop.sociedad || '',
        folio: prop.folio || '',
        proveedor: p.proveedor || '',
        referencia: p.referencia || '',
        uuid: p.uuid || '',
        cuentaPago: p.cuentaPago || '',
        moneda: p.moneda || '',
        monedaPagoReal: p.monedaPagoReal || '',
        tcAplicado: p.tcAplicado,
        montoAplicado: Number(a.monto) || 0,
        solicitante: p.solicitante || ''
      };
    }).filter(Boolean);

    filas = filas.filter(function (f) {
      if (filtroDesde && (!f.fecha || f.fecha < filtroDesde)) return false;
      if (filtroHastaExclusivo && (!f.fecha || f.fecha >= filtroHastaExclusivo)) return false;
      if (filtroSociedad && f.sociedad !== filtroSociedad) return false;
      if (filtroMoneda && f.moneda !== filtroMoneda) return false;
      if (filtroBusqueda) {
        var hay = [f.proveedor, f.folio, f.uuid, f.referencia, f.cuentaPago].join(' ').toLowerCase();
        if (hay.indexOf(filtroBusqueda) < 0) return false;
      }
      return true;
    });

    filas.forEach(function (f) {
      var catInfo = catMap[_normalizarCuenta(f.cuentaPago)];
      f.banco = catInfo ? (catInfo.banco || '') : '';
      f.cuentaCorta = f.cuentaPago ? String(f.cuentaPago).slice(-4) : '';

      var hayConversion = !!(f.monedaPagoReal && f.monedaPagoReal !== f.moneda);
      var monedaExtranjera = f.moneda !== 'MXN';
      // Hallazgo real 2026-08-02 (reportado por el usuario con caso real:
      // 2 partidas USD, TC 17.4312 capturado, reporte mostraba un monto
      // MXN sin relacion con ese TC): ANTES este calculo solo leia
      // f.tcAplicado cuando hayConversion era true -- pero TC_APLICADO se
      // puede capturar como dato informativo aunque monedaPagoReal NUNCA
      // se haya cambiado (partida pagada desde una cuenta en su propia
      // moneda extranjera, ver comentario de View_Pagos.html linea ~932:
      // "el TC es solo informativo... se puede capturar siempre que la
      // partida no sea MXN"). Gatearlo detras de hayConversion descartaba
      // el TC real capturado por Tesoreria y caia a _tcDelDiaOrNull (una
      // referencia generica de mercado, no lo que de verdad se uso) --
      // ahora se lee independientemente de hayConversion.
      var tcAplicadoNum = (f.tcAplicado !== '' && f.tcAplicado != null && Number(f.tcAplicado) > 0) ? Number(f.tcAplicado) : null;

      f.tcAplicado = tcAplicadoNum;
      f.tcDelDia = null;
      f.montoMonedaExtranjera = null;
      f.montoMxn = null;

      if (hayConversion) {
        // monto ya viene en MXN (ver _montoEfectivoPartida, se convierte
        // ANTES de repartirse en exhibiciones) -- se reconstruye el monto
        // en moneda extranjera dividiendo entre el TC real capturado.
        f.montoMxn = f.montoAplicado;
        f.montoMonedaExtranjera = tcAplicadoNum ? (f.montoAplicado / tcAplicadoNum) : null;
      } else if (monedaExtranjera) {
        f.montoMonedaExtranjera = f.montoAplicado;
        if (TC_DIA_MONEDAS_SOPORTADAS.indexOf(f.moneda) >= 0 && f.fecha) {
          f.tcDelDia = _tcDelDiaOrNull(ss, f.fecha); // referencia, se muestra aparte para comparar aunque no se use para el monto
        }
        // El TC real capturado por Tesoreria SIEMPRE tiene prioridad
        // sobre la referencia generica del dia para calcular Monto MXN --
        // es el dato verificado de esta partida especifica, no un
        // promedio de mercado.
        if (tcAplicadoNum) {
          f.montoMxn = f.montoAplicado * tcAplicadoNum;
        } else if (f.tcDelDia) {
          f.montoMxn = f.montoAplicado * f.tcDelDia;
        }
      } else {
        f.montoMxn = f.montoAplicado;
      }

      // Hallazgo real de revisor: toISOString() convierte a UTC -- un pago
      // aplicado despues de las 18:00 hora Mexico (UTC-6) caia en el dia
      // UTC siguiente, agrupandose bajo el dia equivocado en un reporte
      // cuyo proposito central es "todo lo pagado, agrupado por dia".
      // _sortKey preserva el orden cronologico real (con hora) ANTES de
      // colapsar a solo la fecha local para agrupar/mostrar.
      f._sortKey = f.fecha ? f.fecha.getTime() : 0;
      f.fecha = f.fecha ? Utilities.formatDate(f.fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd') : null;
    });

    filas.sort(function (a, b) { return b._sortKey - a._sortKey; });
    filas.forEach(function (f) { delete f._sortKey; });

    return { status: 'success', data: filas, message: '' };
  } catch (e) {
    return { status: 'error', data: [], message: e.toString() };
  }
}

// La funcion temporal diagnosticoFeedbackPagos20260731 (diagnostico de
// solo lectura contra datos reales de BD_TESO, usada para confirmar/
// descartar 3 puntos de la ronda de retroalimentacion de diseno de esa
// fecha) vivio aqui durante esa ronda y se retiro 2026-08-01 una vez
// revisado su log de ejecucion (ver PROJECT_STATUS.md).
