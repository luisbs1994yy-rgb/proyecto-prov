/* ===================== OCR — Registro Rápido =====================
 * Requiere: Drive API (Advanced Service) habilitado en GAS.
 * GAS → Servicios → Drive API → Agregar
 */

function ocrImage_(base64Data, mimeType) {
  const decoded = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(decoded, mimeType || 'image/jpeg', 'rr_ocr_temp');

  // Insertar en Drive con conversión a Google Doc (activa OCR automáticamente)
  const resource = {
    title: 'rr_ocr_temp',
    mimeType: 'application/vnd.google-apps.document'
  };
  const file = Drive.Files.insert(resource, blob, { convert: true, ocr: true });

  let text = '';
  try {
    const doc = DocumentApp.openById(file.getId());
    text = doc.getBody().getText();
  } finally {
    try { Drive.Files.remove(file.getId()); } catch (e) {}
  }

  return { ok: true, text: text.trim() };
}
