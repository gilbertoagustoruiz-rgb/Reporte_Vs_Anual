/**
 * Puente de envío del Reporte Vs Anual por Gmail. Versión 2 (con gráficos dentro del correo).
 *
 * Render (plan gratis) no deja conectarse a los servidores de correo, así que el reporte
 * le pide a este script de Google que envíe el correo. El script corre con TU cuenta de Gmail:
 * el correo sale de tu dirección y queda en tus Enviados.
 *
 * Instalación (una sola vez):
 *  1. Entra a https://script.google.com con tu cuenta de belowgroup → Nuevo proyecto.
 *  2. Borra lo que aparece y pega todo este archivo.
 *  3. En CLAVE escribe una clave larga inventada por ti (la misma irá en Render como MAIL_WEBHOOK_SECRET).
 *  4. Implementar → Nueva implementación → tipo "Aplicación web".
 *     - Ejecutar como: Yo.
 *     - Quién tiene acceso: Cualquier usuario.
 *     Autoriza los permisos que pide Google (enviar correo como tú).
 *  5. Copia la URL de la aplicación web (termina en /exec) y ponla en Render como MAIL_WEBHOOK_URL.
 */
const CLAVE = 'ESCRIBE-AQUI-UNA-CLAVE-LARGA';

// El reporte consulta la versión para saber si puede mandar gráficos.
function doGet() {
  return salida({ ok: true, version: 2 });
}

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    if (!CLAVE || CLAVE.indexOf('ESCRIBE-AQUI') === 0 || d.secret !== CLAVE) return salida({ ok: false, error: 'La clave no coincide con la del script.' });
    const opciones = { htmlBody: d.html };
    if (d.nombre) opciones.name = d.nombre;
    if (d.cc) opciones.cc = d.cc;
    if (d.adjunto && d.adjunto.base64) {
      opciones.attachments = [Utilities.newBlob(Utilities.base64Decode(d.adjunto.base64), d.adjunto.tipo, d.adjunto.nombre)];
    }
    if (d.imagenes && d.imagenes.length) {
      opciones.inlineImages = {};
      d.imagenes.forEach(function (im) {
        opciones.inlineImages[im.cid] = Utilities.newBlob(Utilities.base64Decode(im.base64), 'image/png', im.cid + '.png');
      });
    }
    GmailApp.sendEmail(d.para, d.asunto, 'Este correo tiene formato. Si no lo ves, ábrelo en Gmail.', opciones);
    return salida({ ok: true, desde: Session.getActiveUser().getEmail() });
  } catch (err) {
    return salida({ ok: false, error: String(err && err.message || err) });
  }
}

function salida(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
