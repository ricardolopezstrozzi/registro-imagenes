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
const NOTIFICATION_EMAIL = 'primesincro@gmail.com';

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
    const fileUrl = file.getUrl();

    // 3. Registrar en la hoja de cálculo
    const sheet = getOrCreateSheet_(SHEET_NAME);
    sheet.appendRow([
      new Date(),
      body.autor || '',
      body.categoria || '',
      body.notas || '',
      name,
      fileUrl
    ]);

    let emailSent = false;
    let emailError = '';
    try {
      sendNotificationEmail_(body, name, fileUrl);
      emailSent = true;
    } catch (mailErr) {
      emailError = mailErr.message;
      console.error('No se pudo enviar el correo: ' + emailError);
    }

    return jsonOut_({ status: 'ok', fileUrl: fileUrl, emailSent: emailSent, emailError: emailError });
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

function sendNotificationEmail_(body, fileName, fileUrl) {
  const autor = body.autor || 'Sin nombre';
  const categoria = body.categoria || 'General';
  const notas = body.notas || '';
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: 'Nuevo registro de imagen',
    htmlBody: [
      '<h2>Nuevo registro de imagen</h2>',
      '<p><strong>Fecha:</strong> ' + escapeHtml_(fecha) + '</p>',
      '<p><strong>Autor:</strong> ' + escapeHtml_(autor) + '</p>',
      '<p><strong>Categoria:</strong> ' + escapeHtml_(categoria) + '</p>',
      '<p><strong>Notas:</strong> ' + escapeHtml_(notas) + '</p>',
      '<p><strong>Archivo:</strong> ' + escapeHtml_(fileName) + '</p>',
      '<p><a href="' + fileUrl + '">Ver imagen en Drive</a></p>'
    ].join('')
  });
}

function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
