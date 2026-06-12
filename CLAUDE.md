# Reglas del proyecto

## Versión — OBLIGATORIO en cada cambio

Cada commit, por pequeño que sea, DEBE aumentar la versión en +0.01:

1. `Constants.gs` — `APP_VERSION = 'X.XX'` y el comentario `Versión X.XX` en el header
2. `index.HTML` — atributo `data-app-version="X.XX"` en la etiqueta `<html>`

La versión actual se puede ver con: `grep APP_VERSION Constants.gs`

Ejemplo: si la versión actual es `3.72`, el próximo cambio la deja en `3.73`.

## Setup de Drive API (para Registro Rápido con foto/OCR)

El usuario debe habilitar una sola vez en Google Apps Script:
- Editor de GAS → **Servicios** (ícono +) → **Drive API** → Agregar

Sin esto, la función `ocrImage_` en `OCR.gs` lanzará un error.
