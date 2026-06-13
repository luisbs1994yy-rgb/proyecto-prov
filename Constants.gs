/**
 * ESTADO DE CUENTA  |  Versión 3.93
 * Compatible con index.HTML v3.93+
 *
 * Clientes (quienes te deben): hoja Proveedores + Registros (campo proveedor = cliente)
 * Proveedores (a quienes les debes): hoja CatalogoProveedores + RegistrosProveedores
 *
 * Hojas:
 * - Registros               (movimientos de clientes)
 * - Proveedores             (lista de clientes)
 * - RegistrosProveedores    (movimientos de proveedores)
 * - CatalogoProveedores     (lista de proveedores)
 * - Logs
 * - PreciosProductos
 */

const APP_VERSION = '3.93';

const SHEET_REGISTROS         = 'Registros';
const SHEET_PROVEEDORES       = 'Proveedores';
const SHEET_REGISTROS_PROV    = 'RegistrosProveedores';
const SHEET_CATALOGO_PROV     = 'CatalogoProveedores';
const SHEET_LOGS              = 'Logs';
const SHEET_PRECIOS_PRODUCTOS = 'PreciosProductos';
const SHEET_HISTORIAL_PRECIOS = 'HistorialPrecios';

/* ===== Registros (clientes y proveedores comparten columnas) ===== */
const REG_HEADERS = [
  'id', 'fecha', 'proveedor', 'tipo', 'monto',
  'factura', 'concepto', 'producto', 'precioLitro',
  'litros', 'aplicaFacturaId', 'createdAt'
];

/* ===== Lista clientes (hoja Proveedores — nombre histórico) ===== */
const PROV_HEADERS = ['nombre'];

/* ===== Lista proveedores (catálogo) ===== */
const CAT_PROV_HEADERS = ['nombre', 'litrosCobrados'];

/* ===== Precios por cliente/producto ===== */
const PRODUCTOS_HEADERS = ['cliente', 'producto', 'precio', 'updatedAt'];
const PRODUCTOS_VALIDOS = ['Etanol', 'Nafta', 'Regular'];

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
  registroId: 10,
  comentario: 11,
  pagoGrupoId: 12
};

const LOG_HEADERS = [
  'id', 'createdAt', 'fecha', 'hora', 'tipo',
  'descripcion', 'localId', 'factura', 'concepto',
  'registroId', 'comentario', 'pagoGrupoId'
];
