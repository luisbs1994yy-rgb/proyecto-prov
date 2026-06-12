/* ===================== SETUP Y REPARACIONES ===================== */

/**
 * Ejecutar una vez después de pegar los archivos .gs en Google Apps Script.
 * También ejecutar después de cualquier migración de estructura de hojas.
 */
function setupSheets() {
  const regMap = ensureHeaders_(sheet_(SHEET_REGISTROS), REG_HEADERS);
  const provSh = sheet_(SHEET_PROVEEDORES);
  ensureProveedoresLayout_(provSh);
  ensureHeaders_(provSh, PROV_HEADERS);
  const regProvMap = ensureHeaders_(sheet_(SHEET_REGISTROS_PROV), REG_HEADERS);
  const catProvSh = sheet_(SHEET_CATALOGO_PROV);
  ensureCatalogoProveedoresLayout_(catProvSh);
  ensureHeaders_(catProvSh, CAT_PROV_HEADERS);
  ensureLogsLayout_(sheet_(SHEET_LOGS));
  ensureHeaders_(sheet_(SHEET_PRECIOS_PRODUCTOS), PRODUCTOS_HEADERS);
  const logsRepair = repairLogsFacturaConcepto();
  return {
    ok: true,
    version: APP_VERSION,
    registrosClientes: { aplica: regMap.aplicafacturaid || 0, producto: regMap.producto || 0 },
    registrosProveedores: { aplica: regProvMap.aplicafacturaid || 0, producto: regProvMap.producto || 0 },
    hojas: [SHEET_REGISTROS, SHEET_PROVEEDORES, SHEET_REGISTROS_PROV, SHEET_CATALOGO_PROV, SHEET_LOGS, SHEET_PRECIOS_PRODUCTOS],
    logs: logsRepair
  };
}

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
