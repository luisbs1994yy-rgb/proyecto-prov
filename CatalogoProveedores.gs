/* ===================== CATÁLOGO PROVEEDORES ===================== */

function getCatalogoProvRows_() {
  const sh = sheet_(SHEET_CATALOGO_PROV);
  ensureCatalogoProveedoresLayout_(sh);
  const map = ensureHeaders_(sh, CAT_PROV_HEADERS);
  const rows = readData_(sh, 2);
  let list = rows
    .map((r) => ({
      nombre: cell_(r, map, 'nombre'),
      litrosCobrados: cell_(r, map, 'litroscobrados')
    }))
    .filter((r) => r.nombre);
  if (!list.length) {
    list = readProveedoresColA_(sh).map((nombre) => ({ nombre: nombre, litrosCobrados: '' }));
  }
  return list;
}

function getProvLitrosCobrados_() {
  const map = {};
  getCatalogoProvRows_().forEach((r) => {
    if (r.litrosCobrados) map[r.nombre] = Number(r.litrosCobrados) || r.litrosCobrados;
  });
  return map;
}

function setProvLitrosCobrados_(proveedor, litrosCobrados) {
  const nombre = str_(proveedor);
  if (!nombre) return { ok: false, error: 'Proveedor requerido' };

  const sh = sheet_(SHEET_CATALOGO_PROV);
  ensureCatalogoProveedoresLayout_(sh);
  const map = ensureHeaders_(sh, CAT_PROV_HEADERS);
  const colNombre = map.nombre;
  const colLc = map.litroscobrados;
  if (!colNombre) return { ok: false, error: 'Hoja catálogo sin columna nombre' };

  const value = litrosCobrados == null || litrosCobrados === '' ? '' : str_(litrosCobrados);
  const lastRow = sh.getLastRow();
  let found = false;

  if (lastRow >= 2) {
    const nombres = sh.getRange(2, colNombre, lastRow - 1, 1).getValues();
    for (let i = 0; i < nombres.length; i++) {
      if (str_(nombres[i][0]).toUpperCase() === nombre.toUpperCase()) {
        if (colLc) sh.getRange(i + 2, colLc).setValue(value);
        found = true;
        break;
      }
    }
  }

  if (!found) {
    appendByMap_(sh, map, { nombre: nombre, litrosCobrados: value });
  }

  return { ok: true, provLitrosCobrados: getProvLitrosCobrados_() };
}

function getCatalogoProveedores_() {
  const rows = getCatalogoProvRows_();
  if (rows.length) {
    return Array.from(new Set(rows.map((r) => r.nombre))).sort((a, b) => a.localeCompare(b, 'es'));
  }
  return getNombresFromSheet_(sheet_(SHEET_CATALOGO_PROV), CAT_PROV_HEADERS, ensureCatalogoProveedoresLayout_);
}

function addCatalogoProveedor_(nombre) {
  return addNombreToSheet_(sheet_(SHEET_CATALOGO_PROV), CAT_PROV_HEADERS, nombre, 'Nombre de proveedor requerido', getCatalogoProveedores_, ensureCatalogoProveedoresLayout_);
}

function deleteCatalogoProveedor_(nombre) {
  return deleteNombreFromSheet_(sheet_(SHEET_CATALOGO_PROV), CAT_PROV_HEADERS, nombre, 'Nombre de proveedor requerido', 'Proveedor no encontrado', ensureCatalogoProveedoresLayout_);
}

function getNombresFromSheet_(sh, headers, layoutFn) {
  if (layoutFn) layoutFn(sh);
  const map = ensureHeaders_(sh, headers);
  const rows = readData_(sh, 2);

  let list = rows.map((r) => cell_(r, map, 'nombre')).filter(Boolean);
  if (!list.length) list = readProveedoresColA_(sh);
  return Array.from(new Set(list)).sort((a, b) => a.localeCompare(b, 'es'));
}

function addNombreToSheet_(sh, headers, nombre, errEmpty, getListFn, layoutFn) {
  const value = str_(nombre);
  if (!value) return { ok: false, error: errEmpty };

  if (layoutFn) layoutFn(sh);
  const exists = getListFn().some((p) => p.toUpperCase() === value.toUpperCase());
  if (!exists) {
    const map = ensureHeaders_(sh, headers);
    appendByMap_(sh, map, { nombre: value });
  }

  return { ok: true };
}

function deleteNombreFromSheet_(sh, headers, nombre, errEmpty, errNotFound, layoutFn) {
  const value = str_(nombre);
  if (!value) return { ok: false, error: errEmpty };

  if (layoutFn) layoutFn(sh);
  const map = ensureHeaders_(sh, headers);
  const deleted = deleteByField_(sh, map, 'nombre', value, true);

  return deleted ? { ok: true } : { ok: false, error: errNotFound };
}

function ensureProveedoresLayout_(sh) {
  ensureNombreSheetLayout_(sh);
}

function ensureCatalogoProveedoresLayout_(sh) {
  ensureNombreSheetLayout_(sh);
}

function ensureNombreSheetLayout_(sh) {
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
