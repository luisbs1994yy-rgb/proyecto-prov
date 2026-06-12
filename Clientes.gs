/* ===================== CLIENTES (hoja Proveedores) ===================== */

function getProveedores_() {
  return getNombresFromSheet_(sheet_(SHEET_PROVEEDORES), PROV_HEADERS, ensureProveedoresLayout_);
}

function addProveedor_(nombre) {
  return addNombreToSheet_(sheet_(SHEET_PROVEEDORES), PROV_HEADERS, nombre, 'Nombre de cliente requerido', getProveedores_, ensureProveedoresLayout_);
}

function deleteProveedor_(nombre) {
  const value = str_(nombre);
  if (!value) return { ok: false, error: 'Nombre de cliente requerido' };

  const deletedIds = deleteRegistrosByNombreFromSheet_(sheet_(SHEET_REGISTROS), value);
  const preciosEliminados = deletePreciosByCliente_(value);
  const logsEliminados = deleteLogsByCliente_(value, deletedIds);

  const sh = sheet_(SHEET_PROVEEDORES);
  ensureProveedoresLayout_(sh);
  const map = ensureHeaders_(sh, PROV_HEADERS);
  const deleted = deleteByField_(sh, map, 'nombre', value, true);

  if (!deleted && !deletedIds.length) {
    return { ok: false, error: 'Cliente no encontrado' };
  }

  return {
    ok: true,
    registrosEliminados: deletedIds.length,
    logsEliminados: logsEliminados,
    preciosEliminados: preciosEliminados
  };
}

function deleteRegistrosByNombreFromSheet_(sh, nombre) {
  const map = ensureHeaders_(sh, REG_HEADERS);
  const colProv = map.proveedor;
  if (!colProv) return [];

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const target = str_(nombre).toUpperCase();
  const width = Math.max(sh.getLastColumn(), REG_HEADERS.length);
  const vals = sh.getRange(2, 1, lastRow - 1, width).getValues();
  const rowsToDelete = [];
  const deletedIds = [];

  for (let i = 0; i < vals.length; i++) {
    const prov = str_(cell_(vals[i], map, 'proveedor')).toUpperCase();
    if (prov === target) {
      rowsToDelete.push(i + 2);
      const id = str_(cell_(vals[i], map, 'id'));
      if (id) deletedIds.push(id);
    }
  }

  rowsToDelete.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  return deletedIds;
}

function deletePreciosByCliente_(cliente) {
  const sh = sheet_(SHEET_PRECIOS_PRODUCTOS);
  const map = ensureHeaders_(sh, PRODUCTOS_HEADERS);
  const col = map.cliente;
  if (!col) return 0;

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  const target = str_(cliente).toUpperCase();
  const vals = sh.getRange(2, col, lastRow - 1, 1).getValues();
  const rows = [];

  for (let i = 0; i < vals.length; i++) {
    if (str_(vals[i][0]).toUpperCase() === target) rows.push(i + 2);
  }

  rows.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  return rows.length;
}

function deleteLogsByCliente_(nombre, registroIds) {
  const sh = sheet_(SHEET_LOGS);
  ensureLogsLayout_(sh);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;

  const target = str_(nombre).toUpperCase();
  const idSet = {};
  (registroIds || []).forEach(function (id) { idSet[str_(id)] = true; });

  const width = Math.max(sh.getLastColumn(), LOG_HEADERS.length);
  const vals = sh.getRange(2, 1, lastRow - 1, width).getValues();
  const toDelete = [];

  for (let i = 0; i < vals.length; i++) {
    const row = vals[i];
    const tipo = str_(row[LOG_COL.tipo - 1]).toLowerCase();
    const regId = str_(row[LOG_COL.registroId - 1]);
    const desc = str_(row[LOG_COL.descripcion - 1]);

    if (regId && idSet[regId]) {
      toDelete.push(i + 2);
      continue;
    }
    if (tipo === 'proveedor' || tipo === 'eliminar') continue;

    const parts = desc.split(' · ');
    if (parts.length && str_(parts[0]).toUpperCase() === target) {
      toDelete.push(i + 2);
    }
  }

  toDelete.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
  return toDelete.length;
}
