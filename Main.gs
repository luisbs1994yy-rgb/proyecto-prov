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
        catalogoProveedores: getCatalogoProveedores_(),
        provLitrosCobrados: getProvLitrosCobrados_(),
        registrosProveedores: getRegistrosProveedores_(),
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
