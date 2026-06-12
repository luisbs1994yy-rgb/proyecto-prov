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
      registroId: str_(row[LOG_COL.registroId - 1]),
      comentario: str_(row[LOG_COL.comentario - 1]),
      pagoGrupoId: str_(row[LOG_COL.pagoGrupoId - 1])
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
  const comentario = str_(data.comentario);
  const pagoGrupoId = str_(data.pagoGrupoId);

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
  sh.getRange(nextRow, LOG_COL.comentario).setValue(comentario);
  sh.getRange(nextRow, LOG_COL.pagoGrupoId).setValue(pagoGrupoId);

  return { ok: true, id: id, factura: facturaVal, concepto: conceptoVal };
}

function updateLog_(localId, id, data) {
  const sh = sheet_(SHEET_LOGS);
  ensureLogsLayout_(sh);

  const local = str_(localId);
  const rowId = str_(id);
  if (!local && !rowId) return { ok: false, error: 'localId o id requerido' };

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: false, error: 'Log no encontrado' };

  let targetRow = 0;
  const valuesId = sh.getRange(2, LOG_COL.id, lastRow - 1, 1).getValues();
  const valuesLocal = sh.getRange(2, LOG_COL.localId, lastRow - 1, 1).getValues();

  for (let i = 0; i < valuesId.length; i++) {
    const currId = str_(valuesId[i][0]);
    const currLocal = str_(valuesLocal[i][0]);
    if ((rowId && currId === rowId) || (local && currLocal === local)) {
      targetRow = i + 2;
      break;
    }
  }

  if (!targetRow) return { ok: false, error: 'Log no encontrado' };

  if (data.descripcion !== undefined) {
    sh.getRange(targetRow, LOG_COL.descripcion).setValue(str_(data.descripcion));
  }
  if (data.concepto !== undefined) {
    sh.getRange(targetRow, LOG_COL.concepto).setValue(str_(data.concepto));
  }
  if (data.comentario !== undefined) {
    sh.getRange(targetRow, LOG_COL.comentario).setValue(str_(data.comentario));
  }
  if (data.fecha !== undefined) {
    sh.getRange(targetRow, LOG_COL.fecha).setValue(str_(data.fecha));
  }
  if (data.hora !== undefined) {
    sh.getRange(targetRow, LOG_COL.hora).setValue(str_(data.hora));
  }
  if (data.createdAt !== undefined) {
    sh.getRange(targetRow, LOG_COL.createdAt).setValue(str_(data.createdAt));
  }

  return { ok: true };
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
    sh.getRange(1, LOG_COL.comentario).setValue('comentario');
    sh.getRange(1, LOG_COL.pagoGrupoId).setValue('pagoGrupoId');
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
