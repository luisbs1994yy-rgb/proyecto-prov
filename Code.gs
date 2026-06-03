/**
 * ESTADO DE CUENTA - Google Apps Script (compatible con index.HTML v2.73+)
 *
 * Columna Registros: aplicaFacturaId = id del registro FACTURA al que se abona un PAGO.
 * Columna Registros: producto = Etanol / Nafta / Regular (solo Factura).
 *
 * IMPORTANTE:
 * 1) Reemplaza TODO el contenido de Code.gs con este archivo.
 * 2) Ejecuta setupSheets() una vez.
 * 3) Implementar > Administrar implementaciones > Nueva versión (obligatorio).
 *
 * Hojas:
 * - Registros
 * - Proveedores
 * - Logs
 * - PreciosProductos
 */

const SHEET_REGISTROS = 'Registros';
const SHEET_PROVEEDORES = 'Proveedores';
const SHEET_LOGS = 'Logs';
const SHEET_PRECIOS_PRODUCTOS = 'PreciosProductos';

/* ===== Registros ===== */
const REG_HEADERS = ['id', 'fecha', 'proveedor', 'tipo', 'monto', 'factura', 'concepto', 'producto', 'precioLitro', 'litros', 'aplicaFacturaId', 'createdAt'];

/* ===== Proveedores ===== */
const PROV_HEADERS = ['nombre'];

/* ===== Precios por cliente/producto ===== */
const PRODUCTOS_HEADERS = ['cliente', 'producto', 'precio', 'updatedAt'];
const PRODUCTOS_VALIDOS = ['Etanol', 'Nafta', 'Regular'];

/*
  Logs (layout real detectado en tu libro):
  A:id | B:createdAt | C:fecha | D:hora | E:tipo | F:descripcion | G:localId | H:factura | I:concepto
  (En fila 1 pueden aparecer etiquetas viejas; el script escribe por posición fija)
*/
const LOG_COL = {
  id: 1,
  createdAt: 2,
  fecha: 3,
  hora: 4,
  tipo: 5,
  descripcion: 6,
  localId: 7,
  factura: 8,
  concepto: 9,
  registroId: 10
};

const LOG_HEADERS = [
  'id',
  'createdAt',
  'fecha',
  'hora',
  'tipo',
  'descripcion',
  'localId',
  'factura',
  'concepto',
  'registroId'
];

function doGet(e) {
  return handleRequest_(e, 'GET');
}

function doPost(e) {
  return handleRequest_(e, 'POST');
}

function handleRequest_(e, method) {
  try {
    const payload = parseRequest_(e, method);
    const action = String(payload.action || '').trim();
    if (!action) return json_({ ok: false, error: 'Acción requerida' });

    const result = route_(action, payload);
    if (result && typeof result.ok === 'undefined') result.ok = true;
    return json_(result);
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function parseRequest_(e, method) {
  if (method === 'GET') return { action: e && e.parameter ? e.parameter.action : '' };
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
  return JSON.parse(raw);
}

function route_(action, payload) {
  switch (action) {
    case 'getAll':
      return {
        registros: getRegistros_(),
        proveedores: getProveedores_(),
        logs: getLogs_(),
        precios: getPreciosProductos_().precios
      };
    case 'getLogs':
      return { logs: getLogs_() };
    case 'getPreciosProductos':
      return getPreciosProductos_();

    case 'addRegistro':
      return addRegistro_(payload.data || {});
    case 'deleteRegistro':
      return deleteRegistro_(payload.id);
    case 'updateRegistro':
      return updateRegistro_(payload.id, payload.data || {});

    case 'addLog':
      return addLog_(payload.data || {});
    case 'deleteLog':
      return deleteLog_(payload.localId, payload.id);

    case 'addProveedor':
      return addProveedor_(payload.nombre);
    case 'deleteProveedor':
      return deleteProveedor_(payload.nombre);

    case 'setPrecioProducto':
      return setPrecioProducto_(payload.data || {});
    case 'setPreciosProductos':
      return setPreciosProductos_(payload.data || {});

    default:
      return { ok: false, error: 'Acción no soportada: ' + action };
  }
}

/* ===================== REGISTROS ===================== */

function getRegistros_() {
  const sh = sheet_(SHEET_REGISTROS);
  const map = ensureHeaders_(sh, REG_HEADERS);
  const rows = readData_(sh, 2);

  return rows
    .map((row) => ({
      id: cell_(row, map, 'id'),
      fecha: cell_(row, map, 'fecha'),
      proveedor: cell_(row, map, 'proveedor'),
      tipo: cell_(row, map, 'tipo'),
      monto: toNum_(cell_(row, map, 'monto')),
      factura: cell_(row, map, 'factura'),
      concepto: cell_(row, map, 'concepto'),
      producto: cell_(row, map, 'producto'),
      aplicaFacturaId: cell_(row, map, 'aplicaFacturaId'),
      createdAt: cell_(row, map, 'createdAt')
    }))
    .filter((r) => r.id || r.fecha || r.proveedor);
}

function addRegistro_(data) {
  const sh = sheet_(SHEET_REGISTROS);
  const map = ensureHeaders_(sh, REG_HEADERS);

  const id = str_(data.id) || uuid_('reg');
  const fecha = str_(data.fecha);
  const proveedor = str_(data.proveedor);
  const tipo = str_(data.tipo);
  const monto = toNum_(data.monto);
  const factura = str_(data.factura);
  const concepto = str_(data.concepto);
  const producto = normalizeProducto_(data.producto);
  const precioLitroRaw = data.precioLitro;
  const litrosRaw = data.litros;
  const precioLitro = precioLitroRaw === '' || precioLitroRaw == null ? '' : toNum_(precioLitroRaw);
  const litros = litrosRaw === '' || litrosRaw == null ? '' : toNum_(litrosRaw);

  if (!fecha || !proveedor || !tipo || isNaN(monto)) {
    return { ok: false, error: 'Datos incompletos para registro' };
  }

  const aplicaFacturaId = str_(data.aplicaFacturaId);
  const errAplica = validateAplicaFacturaId_(aplicaFacturaId, proveedor, tipo);
  if (errAplica) return { ok: false, error: errAplica };

  appendByMap_(sh, map, {
    id: id,
    fecha: fecha,
    proveedor: proveedor,
    tipo: tipo,
    monto: monto,
    factura: factura,
    concepto: concepto,
    producto: producto,
    precioLitro: precioLitro === '' || isNaN(precioLitro) ? '' : precioLitro,
    litros: litros === '' || isNaN(litros) ? '' : litros,
    aplicaFacturaId: aplicaFacturaId,
    createdAt: new Date().toISOString()
  });

  return { ok: true, id: id };
}

function updateRegistro_(id, data) {
  const target = str_(id);
  if (!target) return { ok: false, error: 'ID requerido' };

  const sh = sheet_(SHEET_REGISTROS);
  const map = ensureHeaders_(sh, REG_HEADERS);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: false, error: 'Registro no encontrado' };

  const colId = map.id;
  if (!colId) return { ok: false, error: 'Columna id no encontrada' };

  const ids = sh.getRange(2, colId, lastRow - 1, 1).getValues();
  let rowNum = -1;
  for (let i = 0; i < ids.length; i++) {
    if (str_(ids[i][0]) === target) {
      rowNum = i + 2;
      break;
    }
  }
  if (rowNum < 0) return { ok: false, error: 'Registro no encontrado' };

  if (Object.prototype.hasOwnProperty.call(data, 'aplicaFacturaId')) {
    const tipoActual = str_(sh.getRange(rowNum, map.tipo).getValue());
    const provActual = str_(sh.getRange(rowNum, map.proveedor).getValue());
    const aplicaVal = str_(data.aplicaFacturaId);
    const errAplica = validateAplicaFacturaId_(aplicaVal, provActual, tipoActual);
    if (errAplica) return { ok: false, error: errAplica };
    const col = map.aplicafacturaid;
    if (col) sh.getRange(rowNum, col).setValue(aplicaVal);
  }
  if (Object.prototype.hasOwnProperty.call(data, 'concepto')) {
    const col = map.concepto;
    if (col) sh.getRange(rowNum, col).setValue(str_(data.concepto));
  }
  if (Object.prototype.hasOwnProperty.call(data, 'factura')) {
    const col = map.factura;
    if (col) sh.getRange(rowNum, col).setValue(str_(data.factura));
  }
  if (Object.prototype.hasOwnProperty.call(data, 'producto')) {
    const col = map.producto;
    if (col) sh.getRange(rowNum, col).setValue(normalizeProducto_(data.producto));
  }
  if (Object.prototype.hasOwnProperty.call(data, 'monto')) {
    const col = map.monto;
    if (col) sh.getRange(rowNum, col).setValue(toNum_(data.monto));
  }

  return { ok: true, id: target };
}

function deleteRegistro_(id) {
  const target = str_(id);
  if (!target) return { ok: false, error: 'ID requerido' };

  const sh = sheet_(SHEET_REGISTROS);
  const map = ensureHeaders_(sh, REG_HEADERS);
  const deleted = deleteByField_(sh, map, 'id', target);

  return deleted ? { ok: true } : { ok: false, error: 'Registro no encontrado' };
}

/** true si el tipo es pago (incluye "pago sin comprobante", etc.) */
function isTipoPago_(tipo) {
  return str_(tipo).toLowerCase().indexOf('pago') !== -1;
}

/**
 * Valida que aplicaFacturaId apunte a una factura del mismo proveedor.
 * Solo los pagos pueden tener aplicaFacturaId.
 */
function validateAplicaFacturaId_(aplicaFacturaId, proveedor, tipo) {
  const apl = str_(aplicaFacturaId);
  if (!apl) return '';
  if (!isTipoPago_(tipo)) {
    return 'Solo un pago puede aplicarse a una factura';
  }
  const regs = getRegistros_();
  const fact = regs.find(function (r) {
    return str_(r.id) === apl;
  });
  if (!fact) return 'Factura no encontrada: ' + apl;
  if (isTipoPago_(fact.tipo)) return 'El destino debe ser una factura, no un pago';
  if (str_(fact.proveedor).toUpperCase() !== str_(proveedor).toUpperCase()) {
    return 'La factura debe ser del mismo proveedor';
  }
  return '';
}

/**
 * Ejecutar una vez: agrega columna aplicaFacturaId/producto en hoja Registros si falta.
 */
function repairRegistrosAplicaColumn() {
  const sh = sheet_(SHEET_REGISTROS);
  const map = ensureHeaders_(sh, REG_HEADERS);
  return {
    ok: true,
    column: map.aplicafacturaid || 0,
    productoColumn: map.producto || 0,
    headers: REG_HEADERS.join(', ')
  };
}

/* ===================== PRECIOS PRODUCTOS ===================== */

function getPreciosProductos_() {
  const sh = sheet_(SHEET_PRECIOS_PRODUCTOS);
  const map = ensureHeaders_(sh, PRODUCTOS_HEADERS);
  const rows = readData_(sh, 2);

  const precios = rows
    .map((row) => {
      const cliente = cell_(row, map, 'cliente');
      const producto = normalizeProducto_(cell_(row, map, 'producto'));
      const precio = toNum_(cell_(row, map, 'precio'));
      return {
        cliente: cliente,
        proveedor: cliente,
        producto: producto,
        precio: isNaN(precio) ? '' : precio,
        updatedAt: cell_(row, map, 'updatedAt')
      };
    })
    .filter((p) => p.cliente && p.producto);

  return { ok: true, precios: precios };
}

function setPrecioProducto_(data) {
  const cliente = str_(data.cliente || data.proveedor).toUpperCase();
  const producto = normalizeProducto_(data.producto);
  const precioRaw = data.precio;
  const precio = precioRaw === '' || precioRaw == null ? '' : toNum_(precioRaw);

  if (!cliente) return { ok: false, error: 'Cliente requerido' };
  if (!producto) return { ok: false, error: 'Producto requerido' };
  if (PRODUCTOS_VALIDOS.indexOf(producto) === -1) {
    return { ok: false, error: 'Producto inválido: ' + producto };
  }
  if (precio !== '' && (isNaN(precio) || precio < 0)) {
    return { ok: false, error: 'Precio inválido' };
  }

  const sh = sheet_(SHEET_PRECIOS_PRODUCTOS);
  const map = ensureHeaders_(sh, PRODUCTOS_HEADERS);
  const lastRow = sh.getLastRow();
  let rowNum = -1;

  if (lastRow >= 2) {
    const rows = sh.getRange(2, 1, lastRow - 1, Math.max(sh.getLastColumn(), PRODUCTOS_HEADERS.length)).getValues();
    for (let i = 0; i < rows.length; i++) {
      const currCliente = cell_(rows[i], map, 'cliente').toUpperCase();
      const currProducto = normalizeProducto_(cell_(rows[i], map, 'producto'));
      if (currCliente === cliente && currProducto === producto) {
        rowNum = i + 2;
        break;
      }
    }
  }

  const payload = {
    cliente: cliente,
    producto: producto,
    precio: precio,
    updatedAt: new Date().toISOString()
  };

  if (precio === '') {
    if (rowNum > 0) sh.deleteRow(rowNum);
    return { ok: true, precio: { cliente: cliente, proveedor: cliente, producto: producto, precio: '' } };
  }

  if (rowNum > 0) {
    Object.keys(payload).forEach((k) => {
      const col = map[k.toLowerCase()];
      if (col) sh.getRange(rowNum, col).setValue(payload[k]);
    });
  } else {
    appendByMap_(sh, map, payload);
  }

  return { ok: true, precio: { cliente: cliente, proveedor: cliente, producto: producto, precio: precio } };
}

function setPreciosProductos_(data) {
  const list = Array.isArray(data.precios) ? data.precios : [];
  if (!list.length) return { ok: false, error: 'Lista de precios requerida' };

  const saved = [];
  list.forEach((item) => {
    const res = setPrecioProducto_(item || {});
    if (!res || res.ok === false) {
      throw new Error(res && res.error ? res.error : 'No se pudo guardar precio');
    }
    saved.push(res.precio);
  });

  return { ok: true, precios: getPreciosProductos_().precios, saved: saved.length };
}

function normalizeProducto_(producto) {
  const p = str_(producto);
  if (!p) return '';
  if (p.toLowerCase() === ('etan' + 'o')) return 'Etanol';
  const found = PRODUCTOS_VALIDOS.find((x) => x.toLowerCase() === p.toLowerCase());
  return found || p;
}

/* ===================== PROVEEDORES ===================== */

function getProveedores_() {
  const sh = sheet_(SHEET_PROVEEDORES);
  ensureProveedoresLayout_(sh);
  const map = ensureHeaders_(sh, PROV_HEADERS);
  const rows = readData_(sh, 2);

  let list = rows.map((r) => cell_(r, map, 'nombre')).filter(Boolean);
  if (!list.length) list = readProveedoresColA_(sh);
  return Array.from(new Set(list)).sort((a, b) => a.localeCompare(b, 'es'));
}

/**
 * Hojas viejas guardan nombres en columna A sin encabezado "nombre".
 * Si la fila 1 no es encabezado, inserta "nombre" arriba.
 */
function ensureProveedoresLayout_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 1) return;

  const a1 = str_(sh.getRange(1, 1).getValue()).toLowerCase();
  if (isProveedorHeader_(a1)) return;

  const hasData = lastRow >= 1 && str_(sh.getRange(1, 1).getValue());
  if (!hasData) return;

  sh.insertRowBefore(1);
  sh.getRange(1, 1).setValue('nombre');
  sh.setFrozenRows(1);
}

function isProveedorHeader_(value) {
  return /^(nombre|proveedor|proveedores|cliente|clientes|name)$/i.test(str_(value));
}

function readProveedoresColA_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 1) return [];

  const vals = sh.getRange(1, 1, lastRow, 1).getValues().flat().map(str_).filter(Boolean);
  if (!vals.length) return [];

  let start = 0;
  if (isProveedorHeader_(vals[0])) start = 1;

  return vals.slice(start);
}

function addProveedor_(nombre) {
  const value = str_(nombre);
  if (!value) return { ok: false, error: 'Nombre de cliente requerido' };

  const exists = getProveedores_().some((p) => p.toUpperCase() === value.toUpperCase());
  if (!exists) {
    const sh = sheet_(SHEET_PROVEEDORES);
    const map = ensureHeaders_(sh, PROV_HEADERS);
    appendByMap_(sh, map, { nombre: value });
  }

  return { ok: true };
}

function deleteProveedor_(nombre) {
  const value = str_(nombre);
  if (!value) return { ok: false, error: 'Nombre de cliente requerido' };

  const sh = sheet_(SHEET_PROVEEDORES);
  const map = ensureHeaders_(sh, PROV_HEADERS);
  const deleted = deleteByField_(sh, map, 'nombre', value, true);

  return deleted ? { ok: true } : { ok: false, error: 'Cliente no encontrado' };
}

/* ===================== LOGS ===================== */

function getLogs_() {
  const sh = sheet_(SHEET_LOGS);
  ensureLogsLayout_(sh);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const width = Math.max(sh.getLastColumn(), LOG_HEADERS.length);
  const rows = sh.getRange(2, 1, lastRow - 1, width).getValues();
  const logs = [];

  rows.forEach((row) => {
    const descripcion = str_(row[LOG_COL.descripcion - 1]);
    const parsed = parseLogExtra_(descripcion, {
      factura: str_(row[LOG_COL.factura - 1]),
      concepto: str_(row[LOG_COL.concepto - 1])
    });

    const item = {
      id: str_(row[LOG_COL.id - 1]),
      createdAt: str_(row[LOG_COL.createdAt - 1]),
      fecha: str_(row[LOG_COL.fecha - 1]),
      hora: str_(row[LOG_COL.hora - 1]),
      tipo: str_(row[LOG_COL.tipo - 1]),
      descripcion: descripcion,
      localId: str_(row[LOG_COL.localId - 1]),
      factura: parsed.factura,
      concepto: parsed.concepto,
      registroId: str_(row[LOG_COL.registroId - 1])
    };

    if (item.id || item.localId || item.descripcion || item.tipo) logs.push(item);
  });

  logs.sort((a, b) => {
    const ta = Date.parse(a.createdAt || '') || 0;
    const tb = Date.parse(b.createdAt || '') || 0;
    return tb - ta;
  });

  return logs;
}

function addLog_(data) {
  const sh = sheet_(SHEET_LOGS);
  ensureLogsLayout_(sh);

  const now = new Date();
  const createdAt = str_(data.createdAt) || now.toISOString();
  const localId = str_(data.localId) || uuid_('log');
  const id = str_(data.id) || uuid_('logrow');

  const descripcion = str_(data.descripcion);
  const parsed = parseLogExtra_(descripcion, data);
  const facturaVal = parsed.factura;
  const conceptoVal = parsed.concepto;

  const fecha = str_(data.fecha);
  const hora = str_(data.hora) || formatHora_(new Date(createdAt));
  const tipo = str_(data.tipo);
  const registroId = str_(data.registroId);

  // Escritura celda por celda (evita que factura/concepto queden vacíos en H/I)
  const nextRow = sh.getLastRow() + 1;
  sh.getRange(nextRow, LOG_COL.id).setValue(id);
  sh.getRange(nextRow, LOG_COL.createdAt).setValue(createdAt);
  sh.getRange(nextRow, LOG_COL.fecha).setValue(fecha);
  sh.getRange(nextRow, LOG_COL.hora).setValue(hora);
  sh.getRange(nextRow, LOG_COL.tipo).setValue(tipo);
  sh.getRange(nextRow, LOG_COL.descripcion).setValue(descripcion);
  sh.getRange(nextRow, LOG_COL.localId).setValue(localId);
  sh.getRange(nextRow, LOG_COL.factura).setValue(facturaVal);
  sh.getRange(nextRow, LOG_COL.concepto).setValue(conceptoVal);
  sh.getRange(nextRow, LOG_COL.registroId).setValue(registroId);

  return { ok: true, id: id, factura: facturaVal, concepto: conceptoVal };
}

function deleteLog_(localId, id) {
  const sh = sheet_(SHEET_LOGS);
  ensureLogsLayout_(sh);

  const local = str_(localId);
  const rowId = str_(id);
  if (!local && !rowId) return { ok: false, error: 'localId o id requerido' };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: false, error: 'Log no encontrado' };

  const toDelete = [];
  const valuesId = sh.getRange(2, LOG_COL.id, lastRow - 1, 1).getValues();
  const valuesLocal = sh.getRange(2, LOG_COL.localId, lastRow - 1, 1).getValues();

  for (let i = 0; i < valuesId.length; i++) {
    const currId = str_(valuesId[i][0]);
    const currLocal = str_(valuesLocal[i][0]);
    if ((rowId && currId === rowId) || (local && currLocal === local)) {
      toDelete.push(i + 2);
    }
  }

  if (!toDelete.length) return { ok: false, error: 'Log no encontrado' };

  toDelete.sort((a, b) => b - a).forEach((r) => sh.deleteRow(r));
  return { ok: true };
}

function ensureLogsLayout_(sh) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  } else {
    sh.getRange(1, LOG_COL.factura).setValue('factura');
    sh.getRange(1, LOG_COL.concepto).setValue('concepto');
    sh.getRange(1, LOG_COL.registroId).setValue('registroId');
    sh.getRange(1, LOG_COL.localId).setValue('localId');
    sh.getRange(1, LOG_COL.createdAt).setValue('createdAt');
    sh.getRange(1, LOG_COL.descripcion).setValue('descripcion');
    sh.getRange(1, LOG_COL.tipo).setValue('tipo');
    sh.getRange(1, LOG_COL.fecha).setValue('fecha');
    sh.getRange(1, LOG_COL.hora).setValue('hora');
    sh.getRange(1, LOG_COL.id).setValue('id');
  }

  sh.setFrozenRows(1);
}

function parseLogExtra_(descripcion, data) {
  let factura = str_(data && data.factura);
  let concepto = str_(data && data.concepto);

  if (factura || concepto) return { factura: factura, concepto: concepto };

  const desc = str_(descripcion);
  if (!desc) return { factura: '', concepto: '' };

  const mf = desc.match(/(?:^|·)\s*Factura:\s*([^·]+)(?:\s*·|$)/i);
  const mc = desc.match(/(?:^|·)\s*Concepto:\s*(.+)$/i);
  if (mf) factura = mf[1].trim();
  if (mc) concepto = mc[1].trim();

  if (!factura || !concepto) {
    const parts = desc.split('·').map((x) => x.trim()).filter(Boolean);
    const tail = parts.slice(2);
    const isReserved = (v) => /^(pago|factura|eliminar|proveedor)$/i.test(v);

    if (!factura && tail[0] && !isReserved(tail[0]) && !/^concepto\s*:/i.test(tail[0])) {
      factura = tail[0];
    }
    if (!concepto && tail.length > 1) {
      const c = tail
        .slice(1)
        .join(' · ')
        .replace(/^concepto\s*:/i, '')
        .trim();
      if (c && !isReserved(c)) concepto = c;
    }
  }

  return { factura: factura, concepto: concepto };
}

/* ===================== HELPERS ===================== */

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function ensureHeaders_(sh, headers) {
  const width = Math.max(sh.getLastColumn(), headers.length, 1);
  const row1 = sh.getRange(1, 1, 1, width).getValues()[0];
  const map = {};

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    headers.forEach((h, i) => (map[h.toLowerCase()] = i + 1));
    sh.setFrozenRows(1);
    return map;
  }

  const existing = row1.map((h) => str_(h).toLowerCase());
  let col = sh.getLastColumn();

  headers.forEach((h) => {
    const key = h.toLowerCase();
    if (!existing.includes(key)) {
      col += 1;
      sh.getRange(1, col).setValue(h);
      existing.push(key);
    }
  });

  const refreshed = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  refreshed.forEach((h, i) => {
    const key = str_(h).toLowerCase();
    if (key) map[key] = i + 1;
  });

  sh.setFrozenRows(1);
  return map;
}

function readData_(sh, startRow) {
  const lastRow = sh.getLastRow();
  if (lastRow < startRow) return [];
  const width = sh.getLastColumn();
  return sh.getRange(startRow, 1, lastRow - startRow + 1, width).getValues();
}

function appendByMap_(sh, map, obj) {
  const width = Math.max(sh.getLastColumn(), Object.keys(map).length);
  const row = new Array(width).fill('');

  Object.keys(obj).forEach((k) => {
    const col = map[k.toLowerCase()];
    if (col) row[col - 1] = obj[k];
  });

  sh.appendRow(row);
}

function deleteByField_(sh, map, field, value, insensitive) {
  const col = map[field.toLowerCase()];
  if (!col) return false;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return false;

  const target = str_(value);
  const vals = sh.getRange(2, col, lastRow - 1, 1).getValues();
  const rows = [];

  for (let i = 0; i < vals.length; i++) {
    const cell = str_(vals[i][0]);
    const match = insensitive ? cell.toUpperCase() === target.toUpperCase() : cell === target;
    if (match) rows.push(i + 2);
  }

  if (!rows.length) return false;
  rows.sort((a, b) => b - a).forEach((r) => sh.deleteRow(r));
  return true;
}

function cell_(row, map, field) {
  const col = map[field.toLowerCase()];
  if (!col) return '';
  return str_(row[col - 1]);
}

function str_(v) {
  return String(v == null ? '' : v).trim();
}

function toNum_(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(str_(v).replace(/,/g, ''));
  return isNaN(n) ? NaN : n;
}

function uuid_(prefix) {
  return prefix + '_' + Utilities.getUuid();
}

function formatHora_(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone() || 'America/Mexico_City',
    'h:mm:ss a'
  );
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

/**
 * Ejecutar una vez para rellenar columnas H (factura) e I (concepto)
 * desde la descripción en filas viejas.
 */
function repairLogsFacturaConcepto() {
  const sh = sheet_(SHEET_LOGS);
  ensureLogsLayout_(sh);

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, updated: 0 };

  let updated = 0;
  for (let r = 2; r <= lastRow; r++) {
    const desc = str_(sh.getRange(r, LOG_COL.descripcion).getValue());
    const currentFact = str_(sh.getRange(r, LOG_COL.factura).getValue());
    const currentCon = str_(sh.getRange(r, LOG_COL.concepto).getValue());
    const parsed = parseLogExtra_(desc, { factura: currentFact, concepto: currentCon });

    if (parsed.factura && parsed.factura !== currentFact) {
      sh.getRange(r, LOG_COL.factura).setValue(parsed.factura);
      updated++;
    }
    if (parsed.concepto && parsed.concepto !== currentCon) {
      sh.getRange(r, LOG_COL.concepto).setValue(parsed.concepto);
      updated++;
    }
  }

  return { ok: true, updated: updated };
}

/** Ejecutar una vez */
function setupSheets() {
  const regMap = ensureHeaders_(sheet_(SHEET_REGISTROS), REG_HEADERS);
  const provSh = sheet_(SHEET_PROVEEDORES);
  ensureProveedoresLayout_(provSh);
  ensureHeaders_(provSh, PROV_HEADERS);
  ensureLogsLayout_(sheet_(SHEET_LOGS));
  ensureHeaders_(sheet_(SHEET_PRECIOS_PRODUCTOS), PRODUCTOS_HEADERS);
  const logsRepair = repairLogsFacturaConcepto();
  return {
    ok: true,
    registrosColumnaAplica: regMap.aplicafacturaid || 0,
    registrosColumnaProducto: regMap.producto || 0,
    preciosProductos: SHEET_PRECIOS_PRODUCTOS,
    logs: logsRepair
  };
}

/** Ejecutar una vez si la hoja Proveedores quedó con encabezado en columna B */
function repairProveedoresSheet() {
  const sh = sheet_(SHEET_PROVEEDORES);
  const lastRow = sh.getLastRow();
  if (lastRow < 1) return { ok: true, moved: 0 };

  let moved = 0;
  const b1 = str_(sh.getRange(1, 2).getValue()).toLowerCase();
  if (b1 === 'nombre' && !isProveedorHeader_(str_(sh.getRange(1, 1).getValue()))) {
    const names = readProveedoresColA_(sh);
    sh.clear();
    sh.getRange(1, 1).setValue('nombre');
    names.forEach((n, i) => sh.getRange(i + 2, 1).setValue(n));
    sh.setFrozenRows(1);
    moved = names.length;
  } else {
    ensureProveedoresLayout_(sh);
    ensureHeaders_(sh, PROV_HEADERS);
  }

  return { ok: true, moved: moved, total: getProveedores_().length };
}
