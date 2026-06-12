/**
 * Backend de la app de registro de imágenes.
 * Se pega en script.google.com y se despliega como Web App:
 *   - Ejecutar como: Yo
 *   - Acceso: Cualquier persona
 *
 * Guarda cada imagen en la carpeta de Drive FOLDER_NAME
 * y registra los datos en una hoja de cálculo SHEET_NAME.
 */

const FOLDER_NAME = 'Registro Imagenes App';
const SHEET_NAME  = 'Registro Imagenes App';

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    // 1. Carpeta destino en Drive (se crea si no existe)
    const folder = getOrCreateFolder_(FOLDER_NAME);

    // 2. Decodificar y guardar la imagen
    const bytes = Utilities.base64Decode(body.data);
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm-ss');
    const name  = stamp + '_' + (body.fileName || 'imagen.jpg');
    const blob  = Utilities.newBlob(bytes, body.mimeType || 'image/jpeg', name);
    const file  = folder.createFile(blob);

    // 3. Registrar en la hoja de cálculo
    const sheet = getOrCreateSheet_(SHEET_NAME);
    sheet.appendRow([
      new Date(),
      body.autor || '',
      body.categoria || '',
      body.notas || '',
      name,
      file.getUrl()
    ]);

    return jsonOut_({ status: 'ok', fileUrl: file.getUrl() });
  } catch (err) {
    return jsonOut_({ status: 'error', message: err.message });
  }
}

function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function getOrCreateSheet_(name) {
  const it = DriveApp.getFilesByName(name);
  let ss;
  if (it.hasNext()) {
    ss = SpreadsheetApp.open(it.next());
  } else {
    ss = SpreadsheetApp.create(name);
    ss.getActiveSheet().appendRow(['Fecha', 'Autor', 'Categoría', 'Notas', 'Archivo', 'Link']);
  }
  return ss.getActiveSheet();
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
