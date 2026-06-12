/* ===================== REGISTROS CLIENTES ===================== */

function getRegistros_(filtro) {
  return getRegistrosFromSheet_(sheet_(SHEET_REGISTROS), filtro);
}

function addRegistro_(data) {
  return addRegistroToSheet_(sheet_(SHEET_REGISTROS), data, getRegistros_);
}

function updateRegistro_(id, data) {
  return updateRegistroInSheet_(sheet_(SHEET_REGISTROS), id, data, getRegistros_);
}

function deleteRegistro_(id) {
  return deleteRegistroFromSheet_(sheet_(SHEET_REGISTROS), id);
}

/* ===================== REGISTROS PROVEEDORES ===================== */

function getRegistrosProveedores_(filtro) {
  return getRegistrosFromSheet_(sheet_(SHEET_REGISTROS_PROV), filtro);
}

function addRegistroProveedor_(data) {
  return addRegistroToSheet_(sheet_(SHEET_REGISTROS_PROV), data, getRegistrosProveedores_);
}

function updateRegistroProveedor_(id, data) {
  return updateRegistroInSheet_(sheet_(SHEET_REGISTROS_PROV), id, data, getRegistrosProveedores_);
}

function deleteRegistroProveedor_(id) {
  return deleteRegistroFromSheet_(sheet_(SHEET_REGISTROS_PROV), id);
}

/* ===================== SHARED LOGIC ===================== */

function getRegistrosFromSheet_(sh, filtro) {
  const map = ensureHeaders_(sh, REG_HEADERS);
  const rows = readData_(sh, 2);

  const desde = filtro && filtro.desde ? filtro.desde : null;
  const hasta = filtro && filtro.hasta ? filtro.hasta : null;

  return rows
    .map((row) => {
      const precioLitro = toNum_(cell_(row, map, 'precioLitro'));
      const litros = toNum_(cell_(row, map, 'litros'));
      return {
        id: cell_(row, map, 'id'),
        fecha: cell_(row, map, 'fecha'),
        proveedor: cell_(row, map, 'proveedor'),
        tipo: cell_(row, map, 'tipo'),
        monto: toNum_(cell_(row, map, 'monto')),
        factura: cell_(row, map, 'factura'),
        concepto: cell_(row, map, 'concepto'),
        producto: cell_(row, map, 'producto'),
        precioLitro: isNaN(precioLitro) ? '' : precioLitro,
        litros: isNaN(litros) ? '' : litros,
        aplicaFacturaId: cell_(row, map, 'aplicaFacturaId'),
        createdAt: cell_(row, map, 'createdAt')
      };
    })
    .filter((r) => r.id || r.fecha || r.proveedor)
    .filter((r) => {
      if (!desde && !hasta) return true;
      const f = str_(r.fecha).substring(0, 10);
      if (desde && f < desde) return false;
      if (hasta && f > hasta) return false;
      return true;
    });
}

function addRegistroToSheet_(sh, data, getRegsFn) {
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
  const errAplica = validateAplicaFacturaId_(aplicaFacturaId, proveedor, tipo, getRegsFn);
  if (errAplica) return { ok: false, error: errAplica };

  if (factura && !isTipoPago_(tipo) && !data._forzar) {
    const dup = checkFacturaDuplicada_(factura, proveedor, getRegsFn);
    if (dup) return { ok: false, error: dup, duplicado: true };
  }

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

function updateRegistroInSheet_(sh, id, data, getRegsFn) {
  const target = str_(id);
  if (!target) return { ok: false, error: 'ID requerido' };

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
    const errAplica = validateAplicaFacturaId_(aplicaVal, provActual, tipoActual, getRegsFn);
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
  if (Object.prototype.hasOwnProperty.call(data, 'fecha')) {
    const col = map.fecha;
    if (col) sh.getRange(rowNum, col).setValue(str_(data.fecha));
  }

  return { ok: true, id: target };
}

function deleteRegistroFromSheet_(sh, id) {
  const target = str_(id);
  if (!target) return { ok: false, error: 'ID requerido' };

  const map = ensureHeaders_(sh, REG_HEADERS);
  const deleted = deleteByField_(sh, map, 'id', target);

  return deleted ? { ok: true } : { ok: false, error: 'Registro no encontrado' };
}

function isTipoPago_(tipo) {
  return str_(tipo).toLowerCase().indexOf('pago') !== -1;
}

function validateAplicaFacturaId_(aplicaFacturaId, proveedor, tipo, getRegsFn) {
  const apl = str_(aplicaFacturaId);
  if (!apl) return '';
  if (!isTipoPago_(tipo)) {
    return 'Solo un pago puede aplicarse a una factura';
  }
  const regs = getRegsFn();
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

function checkFacturaDuplicada_(factura, proveedor, getRegsFn) {
  const numFact = str_(factura).trim().toUpperCase();
  if (!numFact) return '';
  const provUpper = str_(proveedor).toUpperCase();
  const regs = getRegsFn();
  const dup = regs.find(function (r) {
    return str_(r.factura).trim().toUpperCase() === numFact &&
           str_(r.proveedor).toUpperCase() === provUpper &&
           !isTipoPago_(r.tipo);
  });
  if (!dup) return '';
  return 'La factura ' + factura + ' ya está registrada para ' + proveedor + ' (fecha: ' + str_(dup.fecha).substring(0, 10) + ')';
}

function repairRegistrosAplicaColumn() {
  const sh = sheet_(SHEET_REGISTROS);
  const map = ensureHeaders_(sh, REG_HEADERS);
  const shProv = sheet_(SHEET_REGISTROS_PROV);
  const mapProv = ensureHeaders_(shProv, REG_HEADERS);
  return {
    ok: true,
    clientes: { column: map.aplicafacturaid || 0, productoColumn: map.producto || 0 },
    proveedores: { column: mapProv.aplicafacturaid || 0, productoColumn: mapProv.producto || 0 },
    headers: REG_HEADERS.join(', ')
  };
}
