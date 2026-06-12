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
