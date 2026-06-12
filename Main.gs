/* ===================== ENTRY POINTS ===================== */

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
  if (method === 'GET') {
    const p = e && e.parameter ? e.parameter : {};
    return { action: p.action || '', desde: p.desde || '', hasta: p.hasta || '' };
  }
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
  return JSON.parse(raw);
}

function route_(action, payload) {
  switch (action) {
    case 'getAll': {
      const filtro = buildFiltro_(payload);
      return {
        registros: getRegistros_(filtro),
        proveedores: getProveedores_(),
        catalogoProveedores: getCatalogoProveedores_(),
        provLitrosCobrados: getProvLitrosCobrados_(),
        registrosProveedores: getRegistrosProveedores_(filtro),
        logs: getLogs_(),
        precios: getPreciosProductos_().precios,
        _filtro: filtro
      };
    }
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

    case 'addRegistroProveedor':
      return addRegistroProveedor_(payload.data || {});
    case 'deleteRegistroProveedor':
      return deleteRegistroProveedor_(payload.id);
    case 'updateRegistroProveedor':
      return updateRegistroProveedor_(payload.id, payload.data || {});

    case 'addLog':
      return addLog_(payload.data || {});
    case 'deleteLog':
      return deleteLog_(payload.localId, payload.id);
    case 'updateLog':
      return updateLog_(payload.localId, payload.id, payload.data || {});

    case 'addProveedor':
      return addProveedor_(payload.nombre);
    case 'deleteProveedor':
      return deleteProveedor_(payload.nombre);

    case 'addCatalogoProveedor':
      return addCatalogoProveedor_(payload.nombre);
    case 'deleteCatalogoProveedor':
      return deleteCatalogoProveedor_(payload.nombre);

    case 'setProvLitrosCobrados':
      return setProvLitrosCobrados_(payload.proveedor, payload.litrosCobrados);

    case 'setPrecioProducto':
      return setPrecioProducto_(payload.data || {});
    case 'setPreciosProductos':
      return setPreciosProductos_(payload.data || {});

    case 'ocrImage':
      return ocrImage_(payload.base64, payload.mimeType);

    default:
      return { ok: false, error: 'Acción no soportada: ' + action };
  }
}

function buildFiltro_(payload) {
  let desde = str_(payload.desde || '');
  let hasta = str_(payload.hasta || '');
  if (!desde && !hasta) {
    const tz = Session.getScriptTimeZone();
    const hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
    const d = new Date();
    d.setDate(d.getDate() - 90);
    desde = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    hasta = hoy;
  }
  return { desde: desde, hasta: hasta };
}
