/*
  LISTA DE PRECIOS - Apps Script
  Pegar estas funciones en tu proyecto de Apps Script.

  Tambien agrega estos casos a tus dispatchers existentes:

  doGet(e):
    if (action === 'getPreciosProductos') return json_(getPreciosProductos_());

  doPost(e):
    if (body.action === 'setPrecioProducto') return json_(setPrecioProducto_(body.data || {}));

  Si tu helper para responder JSON no se llama json_, usa:
    ContentService.createTextOutput(JSON.stringify(obj))
      .setMimeType(ContentService.MimeType.JSON)
*/

var PRECIOS_PRODUCTOS_SHEET = 'PreciosProductos';
var PRODUCTOS_PRECIOS = ['Etanol', 'Nafta', 'Regular'];

function ensurePreciosProductosSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(PRECIOS_PRODUCTOS_SHEET);
  if (!sh) sh = ss.insertSheet(PRECIOS_PRODUCTOS_SHEET);

  var headers = ['Cliente', 'Producto', 'Precio', 'UpdatedAt'];
  var current = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  var needsHeader = headers.some(function (h, i) { return String(current[i] || '').trim() !== h; });
  if (needsHeader) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function normalizeProductoPrecio_(producto) {
  var p = String(producto || '').trim();
  if (p.toLowerCase() === ('etan' + 'o')) return 'Etanol';
  var found = PRODUCTOS_PRECIOS.filter(function (x) { return x.toLowerCase() === p.toLowerCase(); })[0];
  return found || p;
}

function getPreciosProductos_() {
  var sh = ensurePreciosProductosSheet_();
  var last = sh.getLastRow();
  if (last < 2) return { precios: [] };

  var values = sh.getRange(2, 1, last - 1, 4).getValues();
  var precios = values
    .filter(function (r) { return String(r[0] || '').trim() && String(r[1] || '').trim(); })
    .map(function (r) {
      return {
        cliente: String(r[0] || '').trim(),
        proveedor: String(r[0] || '').trim(), // compatibilidad con el HTML actual
        producto: normalizeProductoPrecio_(r[1]),
        precio: Number(r[2]) || 0,
        updatedAt: r[3] || ''
      };
    });

  return { precios: precios };
}

function setPrecioProducto_(data) {
  data = data || {};
  var cliente = String(data.cliente || data.proveedor || '').trim().toUpperCase();
  var producto = normalizeProductoPrecio_(data.producto);
  var precioRaw = data.precio;
  var precio = precioRaw === '' || precioRaw == null ? '' : Number(precioRaw);

  if (!cliente) throw new Error('Cliente requerido');
  if (!producto) throw new Error('Producto requerido');
  if (PRODUCTOS_PRECIOS.indexOf(producto) === -1) throw new Error('Producto invalido: ' + producto);
  if (precio !== '' && (isNaN(precio) || precio < 0)) throw new Error('Precio invalido');

  var sh = ensurePreciosProductosSheet_();
  var last = sh.getLastRow();
  var rowToUpdate = 0;

  if (last >= 2) {
    var values = sh.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < values.length; i++) {
      var c = String(values[i][0] || '').trim().toUpperCase();
      var p = normalizeProductoPrecio_(values[i][1]);
      if (c === cliente && p === producto) {
        rowToUpdate = i + 2;
        break;
      }
    }
  }

  if (rowToUpdate) {
    sh.getRange(rowToUpdate, 1, 1, 4).setValues([[cliente, producto, precio, new Date()]]);
  } else {
    sh.appendRow([cliente, producto, precio, new Date()]);
  }

  return {
    ok: true,
    precio: {
      cliente: cliente,
      proveedor: cliente,
      producto: producto,
      precio: precio
    }
  };
}
