/**
 * Agrega este archivo como un .gs adicional al proyecto de VISITAS.
 * En el doPost existente, justo después del parseo de `datos` (incluido su
 * fallback) y antes de procesar un ingreso normal, agrega:
 *
 * var respuestaVale = procesarSolicitudEntregaVales_(datos);
 * if (respuestaVale !== null) return responderJson(respuestaVale);
 *
 * No declares otro doPost ni otro doGet. Crea la propiedad de script
 * VALE_API_TOKEN y vuelve a desplegar la aplicación web como nueva versión.
 */
var VALE_DELIVERY_SPREADSHEET_ID = '1sZX7LQSlDEWJwRMlbz4SMFBXv2FCS1ztbxdW3EwCVT8';
var VALE_DELIVERY_API_VERSION = '2026-10-17-tiempo-espera-6';

function procesarSolicitudEntregaVales_(payload) {
  if (!payload || !['verificarConexion', 'registrarEntrega'].includes(payload.action)) {
    return null;
  }

  try {
    var expectedToken = PropertiesService.getScriptProperties().getProperty('VALE_API_TOKEN');
    if (!expectedToken || payload.apiToken !== expectedToken) {
      return { ok: false, error: 'No autorizado.' };
    }
    if (payload.action === 'verificarConexion') {
      return { ok: true, version: VALE_DELIVERY_API_VERSION };
    }

    var codigo = String(payload.codigoFarmacia || '').replace(/\D/g, '');
    var persona = String(payload.persona || '').trim();
    var cantidad = Number(payload.cantidadPedidos);
    var operario = String(payload.operarioLogueado || '').trim();
    if (!codigo) throw new Error('El código de farmacia no es válido.');
    if (!persona) throw new Error('No se recibió el nombre de la persona seleccionada.');
    if (!Number.isSafeInteger(cantidad) || cantidad < 0 || cantidad > 10000) {
      throw new Error('La cantidad de pedidos debe ser un número entero entre 0 y 10000.');
    }
    if (!operario) throw new Error('No se recibió el operario de la sesión. Cierra sesión e ingresa nuevamente.');

    var ss = SpreadsheetApp.openById(VALE_DELIVERY_SPREADSHEET_ID);
    var hojaRespuestas = ss.getSheetByName('Respuestas de formulario 1');
    var hojaRegistro = ss.getSheetByName('REGISTRO');
    if (!hojaRespuestas) throw new Error('No existe la pestaña "Respuestas de formulario 1".');
    if (!hojaRegistro) throw new Error('No existe la pestaña "REGISTRO".');

    var rows = hojaRespuestas.getDataRange().getDisplayValues();
    var authorizedName = '';
    var pharmacyName = '';
    var pharmacyEmail = '';
    var personaDni = extraerDniPersonaVale_(persona);
    for (var rowIndex = 1; rowIndex < rows.length; rowIndex++) {
      var row = rows[rowIndex];
      if (String(row[3] || '').replace(/\D/g, '') !== codigo) continue;
      pharmacyName = String(row[2] || '').trim();
      for (var personIndex = 4; personIndex <= 9; personIndex++) {
        var nameInSheet = String(row[personIndex] || '').trim();
        var dniInSheet = String(row[11 + personIndex - 4] || '').replace(/\D/g, '');
        if (personaDni && dniInSheet === personaDni) {
          authorizedName = nameInSheet;
          pharmacyEmail = String(row[1] || '').trim();
          break;
        }
      }
      if (authorizedName) break;
    }

    var authorized = Boolean(authorizedName);
    var state = authorized ? 'AUTORIZADO' : 'DENEGADO';
    var registeredName = authorized ? authorizedName : persona + ' (No autorizado)';
    var fechaRegistro = new Date();
    var tiempoEspera = formatearTiempoEsperaVale_(payload.inicioEspera, fechaRegistro);
    hojaRegistro.appendRow([
      fechaRegistro,
      codigo,
      registeredName,
      authorized ? cantidad : 0,
      state,
      operario,
      tiempoEspera
    ]);
    var avisoCorreo = '';
    if (authorized) {
      avisoCorreo = enviarAvisoRetiroValeIntegrado_(
        pharmacyEmail, pharmacyName, codigo, authorizedName, cantidad, operario, fechaRegistro, ss.getSpreadsheetTimeZone()
      );
    }

    return {
      ok: true,
      estado: state,
      mensaje: authorized
        ? 'Acceso autorizado para: ' + authorizedName + '. Registro guardado.'
        : 'Acceso denegado: el nombre no figura autorizado para esta farmacia.',
      persona: registeredName,
      farmacia: pharmacyName,
      tiempoEspera: tiempoEspera,
      avisoCorreo: avisoCorreo
    };
  } catch (error) {
    console.error('Error al procesar entrega de vales: ' + error);
    return { ok: false, error: String(error && error.message || error) };
  }
}

function enviarAvisoRetiroValeIntegrado_(email, farmacia, codigo, persona, cantidad, operario, fecha, zonaHoraria) {
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    console.error('El retiro quedó registrado, pero no se envió el aviso: falta un correo válido en la columna B.');
    return 'El retiro quedó registrado, pero no se pudo enviar el aviso: falta un correo válido de la farmacia en la columna B del formulario.';
  }

  var fechaTexto = Utilities.formatDate(fecha, zonaHoraria, 'dd/MM/yyyy HH:mm:ss');
  var asunto = 'Aviso de retiro de vales - ' + farmacia + ' (' + codigo + ')';
  var mensaje = 'AVISO DE ENTREGA REGISTRADA\n\n' +
    'Farmacia: ' + farmacia + ' (código ' + codigo + ')\n' +
    'Persona: ' + persona + '\n' +
    'Cantidad de vales: ' + cantidad + '\n' +
    'Operario: ' + operario + '\n' +
    'Fecha y hora: ' + fechaTexto + '\n\n' +
    'Mensaje automático del sistema de entregas. Por favor, no responda a este correo.';
  var cuerpoHTML = '<div style="margin:0;padding:32px 16px;background-color:#f3f6fa;font-family:Arial,Helvetica,sans-serif;color:#243247;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="max-width:600px;margin:0 auto;background-color:#ffffff;border:1px solid #e2e8f0;border-radius:12px;overflow:hidden;">' +
    '<tr><td style="padding:24px 32px;background-color:#17324d;color:#ffffff;">' +
    '<p style="margin:0 0 8px;font-size:12px;letter-spacing:1.5px;text-transform:uppercase;color:#b9d5e8;">Sistema de entregas</p>' +
    '<h1 style="margin:0;font-size:22px;line-height:1.3;font-weight:700;">Entrega registrada</h1>' +
    '</td></tr>' +
    '<tr><td style="padding:28px 32px 12px;">' +
    '<p style="margin:0 0 12px;font-size:15px;line-height:1.6;">Estimado equipo de <strong>' + escaparHTMLAvisoRetiroValeIntegrado_(farmacia) + '</strong>:</p>' +
    '<p style="margin:0;font-size:14px;line-height:1.6;color:#526174;">Se registró correctamente el retiro de vales correspondiente a su farmacia.</p>' +
    '</td></tr>' +
    '<tr><td style="padding:16px 32px 28px;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;background-color:#f7f9fc;border:1px solid #e7edf4;border-radius:8px;">' +
    '<tr><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;font-size:13px;color:#68778a;">Código de farmacia</td><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;text-align:right;font-size:14px;font-weight:700;color:#243247;">' + escaparHTMLAvisoRetiroValeIntegrado_(codigo) + '</td></tr>' +
    '<tr><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;font-size:13px;color:#68778a;">Persona autorizada</td><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;text-align:right;font-size:14px;font-weight:700;color:#243247;">' + escaparHTMLAvisoRetiroValeIntegrado_(persona) + '</td></tr>' +
    '<tr><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;font-size:13px;color:#68778a;">Cantidad de vales</td><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;text-align:right;font-size:14px;font-weight:700;color:#243247;">' + escaparHTMLAvisoRetiroValeIntegrado_(cantidad) + '</td></tr>' +
    '<tr><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;font-size:13px;color:#68778a;">Operario</td><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;text-align:right;font-size:14px;font-weight:700;color:#243247;">' + escaparHTMLAvisoRetiroValeIntegrado_(operario) + '</td></tr>' +
    '<tr><td style="padding:12px 16px;font-size:13px;color:#68778a;">Fecha y hora</td><td style="padding:12px 16px;text-align:right;font-size:14px;font-weight:700;color:#243247;">' + escaparHTMLAvisoRetiroValeIntegrado_(fechaTexto) + '</td></tr>' +
    '</table></td></tr>' +
    '<tr><td style="padding:16px 32px 24px;border-top:1px solid #edf1f5;">' +
    '<p style="margin:0;text-align:center;font-size:11px;line-height:1.5;color:#8793a2;">Este es un mensaje automático del sistema de entregas. Por favor, no responda a este correo.</p>' +
    '</td></tr></table></div>';
  try {
    MailApp.sendEmail({ to: email, subject: asunto, body: mensaje, htmlBody: cuerpoHTML });
    return '';
  } catch (emailError) {
    console.error('El retiro quedó registrado, pero no se pudo enviar el aviso a ' + email + ': ' + emailError);
    return 'El retiro quedó registrado, pero Gmail no pudo enviar el aviso al correo de la farmacia. Revisa los permisos y la cuota de envío de Apps Script.';
  }
}

function escaparHTMLAvisoRetiroValeIntegrado_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatearTiempoEsperaVale_(inicioEspera, fechaRegistro) {
  var inicio = new Date(String(inicioEspera || ''));
  var inicioMs = inicio.getTime();
  var finMs = fechaRegistro.getTime();
  if (!Number.isFinite(inicioMs) || inicioMs > finMs) {
    throw new Error('La fecha y hora de ingreso de la persona no son válidas.');
  }
  var segundos = Math.floor((finMs - inicioMs) / 1000);
  if (segundos > 12 * 60 * 60) return 'EXEDIDO';
  var horas = String(Math.floor(segundos / 3600)).padStart(2, '0');
  var minutos = String(Math.floor((segundos % 3600) / 60)).padStart(2, '0');
  return horas + ':' + minutos;
}

function normalizarPersonaVale_(value) {
  return String(value == null ? '' : value)
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLocaleLowerCase()
    .replace(/\s+/g, ' ');
}

function extraerDniPersonaVale_(value) {
  var match = String(value == null ? '' : value).match(/(?:^|\D)(\d{7,8})(?=\D|$)/);
  return match ? match[1] : '';
}
