/**
 * Integración del portal de socios y del control de entregas.
 *
 * procesarFormularioSocios es llamada por el formulario HTML mediante
 * google.script.run. doPost recibe las entregas y solicitudes de reporte de server.js.
 * Para doPost, crear VALE_API_TOKEN en las propiedades del script, desplegar
 * como aplicación web ejecutada como propietario y definir REGISTRO_WRITE_URL
 * y REGISTRO_API_TOKEN en el entorno de Node.js.
 */
var VALE_API_VERSION = '2026-10-20-maestro-multi-additions-7';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('FormularioSocios')
    .setTitle('Portal de Socios - Autorización de Retiro')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    var payload = JSON.parse(e && e.postData && e.postData.contents || '{}');
    var expectedToken = PropertiesService.getScriptProperties().getProperty('VALE_API_TOKEN');
    if (!expectedToken || payload.apiToken !== expectedToken) {
      return jsonResponse({ ok: false, error: 'No autorizado.' });
    }
    if (payload.action === 'verificarConexion') {
      return jsonResponse({ ok: true, version: VALE_API_VERSION });
    }
    if (payload.action === 'enviarReporteGerencial') {
      return jsonResponse(enviarReporteGerencialMostrador_(payload));
    }
    if (!['registrarEntrega', 'registrarEntregas', 'consultarAgregados', 'registrarAgregados', 'registrarAgregado'].includes(payload.action)) {
      return jsonResponse({ ok: false, error: 'Acción no válida.' });
    }

    if (payload.action === 'consultarAgregados') {
      return jsonResponse({
        ok: true,
        entregas: consultarAgregadosMostrador(payload.filaOriginal)
      });
    }
    if (payload.action === 'registrarAgregados' || payload.action === 'registrarAgregado') {
      var additions = payload.action === 'registrarAgregado'
        ? [registrarAgregadoMostrador(payload)]
        : registrarAgregadosMostrador(payload);
      return jsonResponse({
        ok: true,
        entregas: additions
      });
    }

    var entries = payload.action === 'registrarEntrega'
      ? [{ codigoFarmacia: payload.codigoFarmacia, cantidadPedidos: payload.cantidadPedidos }]
      : payload.entregas;
    var results = registrarEntregasMostrador(
      entries, payload.persona, payload.operarioLogueado, payload.inicioEspera
    );
    if (payload.action === 'registrarEntrega') {
      var result = results[0];
      return jsonResponse({
        ok: true,
        estado: result.estado,
        mensaje: result.mensaje,
        persona: result.persona,
        farmacia: result.farmacia,
        tiempoEspera: result.tiempoEspera,
        avisoCorreo: result.avisoCorreo
      });
    }
    return jsonResponse({
      ok: true,
      entregas: results
    });
  } catch (error) {
    return jsonResponse({ ok: false, error: String(error && error.message || error) });
  }
}

function jsonResponse(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

function enviarReporteGerencialMostrador_(payload) {
  var rowNumber = Number(payload.filaRegistro);
  var email = String(payload.email || '').trim();
  var solicitadoPor = String(payload.solicitadoPor || '').trim();
  if (!Number.isSafeInteger(rowNumber) || rowNumber < 2) {
    throw new Error('La fila del reporte no es válida.');
  }
  if (email.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('Ingresa una dirección de correo electrónico válida para gerencia.');
  }

  var hojaRegistro = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('REGISTRO');
  if (!hojaRegistro) throw new Error('No existe la pestaña "REGISTRO".');
  if (rowNumber > hojaRegistro.getLastRow()) throw new Error('No se encontró la entrega solicitada en REGISTRO.');
  var registro = hojaRegistro.getRange(rowNumber, 1, 1, 9);
  var values = registro.getDisplayValues()[0];
  var ingreso = String(values[0] || '').trim();
  var codigo = String(values[1] || '').trim();
  var persona = String(values[2] || '').trim();
  var cantidad = String(values[3] || '').trim();
  var estado = String(values[4] || '').trim();
  var operario = String(values[5] || '').trim();
  var tiempoEspera = obtenerTiempoEsperaRegistro_(hojaRegistro, rowNumber);
  var egreso = String(values[8] || '').trim();
  if (!ingreso || !egreso || !tiempoEspera || !persona || !codigo) {
    throw new Error('Faltan datos de ingreso, egreso, persona o tiempo de espera en esta fila; no se envió un reporte incompleto.');
  }

  var safe = escaparHTML;
  var generatedAt = Utilities.formatDate(
    new Date(),
    SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'America/Argentina/Buenos_Aires',
    'dd/MM/yyyy HH:mm:ss'
  );
  var subject = 'Reporte ejecutivo de atención - Registro ' + rowNumber;
  var plainText = [
    'REPORTE EJECUTIVO DE ATENCIÓN',
    'Persona: ' + persona,
    'Código de farmacia: ' + codigo,
    'Estado: ' + estado,
    'Ingreso: ' + ingreso,
    'Egreso: ' + egreso,
    'Tiempo de espera: ' + tiempoEspera,
    'Pedidos: ' + (cantidad || '0'),
    'Operario: ' + (operario || 'No disponible'),
    'Solicitado por: ' + (solicitadoPor || 'Operador de mostrador'),
    'Emitido: ' + generatedAt
  ].join('\n');
  var details = [
    ['Persona', persona],
    ['Código de farmacia', codigo],
    ['Estado de atención', estado],
    ['Cantidad de pedidos', cantidad || '0'],
    ['Operario responsable', operario || 'No disponible'],
    ['Hora de ingreso', ingreso],
    ['Hora de egreso', egreso],
    ['Tiempo de espera', tiempoEspera]
  ];
  var rows = details.map(function(item) {
    return '<tr><td style="padding:13px 16px;border-bottom:1px solid #e8edf2;color:#64748b;font-size:13px;">' +
      safe(item[0]) + '</td><td style="padding:13px 16px;border-bottom:1px solid #e8edf2;text-align:right;color:#172b45;font-size:14px;font-weight:700;">' +
      safe(item[1]) + '</td></tr>';
  }).join('');
  var htmlBody =
    '<div style="margin:0;padding:32px 12px;background:#f1f5f9;font-family:Arial,Helvetica,sans-serif;color:#172b45;">' +
    '<table role="presentation" style="width:100%;max-width:640px;margin:0 auto;border-collapse:collapse;background:#fff;border:1px solid #dfe7ef;border-radius:12px;overflow:hidden;">' +
    '<tr><td style="padding:28px 32px;background:#10263d;color:#fff;">' +
    '<div style="font-size:11px;letter-spacing:2px;color:#72d6cf;font-weight:700;">VALE · CONTROL DE MOSTRADOR</div>' +
    '<h1 style="margin:12px 0 4px;font-size:24px;line-height:1.25;">Reporte ejecutivo</h1>' +
    '<div style="font-size:13px;color:#c5d3e0;">Resumen de atención y permanencia en mostrador</div>' +
    '</td></tr><tr><td style="padding:24px 32px 8px;">' +
    '<div style="display:inline-block;padding:6px 10px;border-radius:20px;background:#e8f7f5;color:#087e78;font-size:11px;font-weight:700;letter-spacing:1px;">REGISTRO ' +
    safe(rowNumber) + '</div><p style="margin:14px 0 8px;color:#526174;font-size:14px;line-height:1.6;">Se presenta el detalle solicitado para revisión de gerencia.</p>' +
    '</td></tr><tr><td style="padding:8px 32px 24px;"><table role="presentation" style="width:100%;border-collapse:collapse;border:1px solid #e8edf2;">' +
    rows + '</table></td></tr><tr><td style="padding:18px 32px;background:#f8fafc;border-top:1px solid #e8edf2;">' +
    '<div style="font-size:11px;line-height:1.7;color:#718096;">Solicitado por ' + safe(solicitadoPor || 'operador de mostrador') +
    ' · Emitido ' + safe(generatedAt) + '<br>El egreso corresponde a la hora del registro de entrega en mostrador.</div>' +
    '</td></tr></table></div>';

  MailApp.sendEmail({ to: email, subject: subject, body: plainText, htmlBody: htmlBody });
  return { ok: true, email: email };
}

function procesarFormularioSocios(datos) {
  try {
    if (!datos || typeof datos !== 'object') {
      throw new Error('No se recibieron los datos del formulario.');
    }

    var email = String(datos.email || '').trim();
    var farmacia = String(datos.farmacia || '').trim();
    var codigo = String(datos.codigo || '').trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
      throw new Error('Ingresa una dirección de correo electrónico válida.');
    }
    if (!farmacia) throw new Error('Ingresa el nombre de la farmacia.');
    if (!/^\d+$/.test(codigo)) throw new Error('El código de farmacia debe contener solo números.');

    var personas = [];
    var dnis = [];
    for (var i = 1; i <= 6; i++) {
      var persona = String(datos['p' + i] || '').trim();
      var dniOriginal = String(datos['dni' + i] || '').trim();
      var dni = soloDigitos(dniOriginal);
      if (persona && (!dni || !/^[\d\s.-]+$/.test(dniOriginal))) {
        throw new Error('Completa el DNI con números para el personal autorizado ' + i + '.');
      }
      if (!persona && dni) {
        throw new Error('Completa el nombre del personal autorizado ' + i + ' o deja su DNI vacío.');
      }
      personas.push(persona);
      dnis.push(dni);
    }
    if (!personas.some(function(persona) { return Boolean(persona); })) {
      throw new Error('Ingresa al menos una persona autorizada con su DNI.');
    }

    var ss = SpreadsheetApp.openById('1sZX7LQSlDEWJwRMlbz4SMFBXv2FCS1ztbxdW3EwCVT8');
    var hoja = ss.getSheetByName('Respuestas de formulario 1');
    if (!hoja) throw new Error('No existe la pestaña "Respuestas de formulario 1".');

    // La hoja usa A:D para los datos de la farmacia, E:J para los nombres,
    // K como separador y L:Q para los DNIs 1:6.
    var fila = new Array(17).fill('');
    fila[0] = new Date();
    fila[1] = email;
    fila[2] = farmacia;
    fila[3] = codigo;
    for (var j = 0; j < 6; j++) {
      fila[4 + j] = personas[j];
      fila[11 + j] = dnis[j];
    }
    hoja.appendRow(fila);

    var autorizados = personas
      .map(function(persona, index) {
        return persona ? '<li>' + escaparHTML(persona) + ' (DNI: ' + escaparHTML(dnis[index]) + ')</li>' : '';
      })
      .filter(function(item) { return Boolean(item); })
      .join('');
    var asunto = 'Confirmación de autorización de retiro de vales - Código: ' + codigo;
    var cuerpoHTML = '<div style="font-family:Arial,sans-serif;color:#333;max-width:600px;margin:0 auto;padding:20px;border:1px solid #e0e0e0;border-radius:8px;">' +
      '<h2 style="color:#2c3e50;border-bottom:2px solid #3498db;padding-bottom:10px;">Confirmación de registro de autorizados</h2>' +
      '<p>Estimado equipo de <strong>' + escaparHTML(farmacia) + '</strong>:</p>' +
      '<p>Recibimos el listado de personal autorizado para retirar pedidos bajo el código <strong>' + escaparHTML(codigo) + '</strong>.</p>' +
      '<ul style="background-color:#f9f9f9;padding:15px 20px;border-radius:5px;list-style-type:square;">' + autorizados + '</ul>' +
      '<p>El registro fue recibido y guardado para las gestiones en mostrador.</p>' +
      '<p>Atentamente,<br><strong>Departamento de Logística y Entregas</strong></p>' +
      '<hr style="border:none;border-top:1px solid #eee;margin-top:20px;">' +
      '<p style="font-size:11px;color:#999;text-align:center;">Mensaje automático; por favor, no responda a este correo.</p></div>';

    try {
      MailApp.sendEmail({ to: email, subject: asunto, htmlBody: cuerpoHTML });
      return { exito: true };
    } catch (emailError) {
      Logger.log('La autorización se guardó, pero no se pudo enviar el correo: ' + emailError);
      return {
        exito: true,
        aviso: 'La autorización se guardó en la hoja, pero no se pudo enviar el correo de confirmación.'
      };
    }
  } catch (error) {
    Logger.log('Error al procesar el formulario de socios: ' + error);
    return { exito: false, error: String(error && error.message || error) };
  }
}

function escaparHTML(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function registrarEntregasMostrador(entregas, personaIngresada, operarioLogueado, inicioEspera) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaMaestro = ss.getSheetByName('MAESTRO');
  var hojaRegistro = ss.getSheetByName('REGISTRO');
  if (!hojaMaestro) throw new Error('No existe la pestaña "MAESTRO".');
  if (!hojaRegistro) throw new Error('No existe la pestaña "REGISTRO".');

  if (!Array.isArray(entregas) || entregas.length < 1 || entregas.length > 10) {
    throw new Error('Selecciona entre una y diez farmacias.');
  }
  var persona = String(personaIngresada || '').trim();
  var operario = String(operarioLogueado || '').trim();
  if (!persona) throw new Error('No se recibió el nombre de la persona seleccionada.');
  if (!operario) throw new Error('No se recibió el operario de la sesión. Cierra sesión e ingresa nuevamente.');

  var personaDni = extraerDniPersona(persona);
  var rows = hojaMaestro.getDataRange().getDisplayValues();
  var fechaRegistro = new Date();
  var fechaVisita = new Date(inicioEspera);
  if (isNaN(fechaVisita.getTime()) || fechaVisita.getTime() > fechaRegistro.getTime()) {
    throw new Error('La fecha y hora de ingreso en MOVIMIENTOS no es válida.');
  }
  var tiempoEspera = formatearTiempoEspera_(fechaVisita, fechaRegistro);
  var seenCodes = {};
  var results = entregas.map(function(entrega) {
    var codigo = soloDigitos(entrega && entrega.codigoFarmacia);
    var cantidad = Number(entrega && entrega.cantidadPedidos);
    if (!codigo) throw new Error('El código de farmacia no es válido.');
    if (seenCodes[codigo]) throw new Error('No repitas el mismo código de farmacia.');
    seenCodes[codigo] = true;
    if (!Number.isSafeInteger(cantidad) || cantidad < 0 || cantidad > 10000) {
      throw new Error('La cantidad de pedidos debe ser un número entero entre 0 y 10000.');
    }

    var authorizedName = '';
    var pharmacyName = '';
    for (var rowIndex = 1; rowIndex < rows.length; rowIndex++) {
      var row = rows[rowIndex];
      if (soloDigitos(row[0]) !== codigo) continue;

      pharmacyName = String(row[1] || '').trim();
      for (var personIndex = 2; personIndex <= 5; personIndex++) {
        var authorizedEntry = String(row[personIndex] || '').trim();
        var dniInSheet = extraerDniPersona(authorizedEntry);
        if (personaDni && dniInSheet === personaDni) {
          authorizedName = authorizedEntry
            .replace(dniInSheet, ' ')
            .replace(/\bD\.?N\.?I\.?\b/gi, ' ')
            .replace(/^\s*[-–—,:;]+|[-–—,:;]+\s*$/g, '')
            .trim();
          break;
        }
      }
      if (authorizedName) break;
    }

    var authorized = Boolean(authorizedName);
    var state = authorized ? 'AUTORIZADO' : 'DENEGADO';
    var registeredName = authorized ? authorizedName : persona + ' (No autorizado)';
    return {
      codigo: codigo,
      farmacia: pharmacyName,
      persona: registeredName,
      cantidad: authorized ? cantidad : 0,
      estado: state,
      tiempoEspera: tiempoEspera,
      values: [fechaVisita, codigo, registeredName, authorized ? cantidad : 0, state, operario, tiempoEspera]
    };
  });

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    prepararColumnaDniRegistro_(hojaRegistro);
    var firstRow = hojaRegistro.getLastRow() + 1;
    hojaRegistro.getRange(firstRow, 1, results.length, 9)
      .setValues(results.map(function(result) {
        return result.values.concat([result.estado === 'AUTORIZADO' ? personaDni : '', fechaRegistro]);
      }));
    aplicarFormatoFechaRegistro_(hojaRegistro, firstRow, results.length);
    return results.map(function(result, index) {
      return {
        estado: result.estado,
        mensaje: result.estado === 'AUTORIZADO'
          ? 'Acceso autorizado para: ' + result.persona + '. Registro guardado.'
          : 'Acceso denegado: el nombre no figura autorizado para esta farmacia.',
        persona: result.persona,
        farmacia: result.farmacia,
        tiempoEspera: result.tiempoEspera,
        avisoCorreo: result.estado === 'AUTORIZADO'
          ? 'No se envió aviso por correo: MAESTRO no contiene una dirección de email.'
          : '',
        filaRegistro: firstRow + index
      };
    });
  } finally {
    lock.releaseLock();
  }
}

function prepararColumnaDniRegistro_(hojaRegistro) {
  var dniHeader = String(hojaRegistro.getRange(1, 8).getDisplayValue() || '').trim();
  if (dniHeader && dniHeader !== 'DNI_PERSONA_SISTEMA') {
    throw new Error('La columna H de REGISTRO ya contiene datos. No se modificó el registro; libera la columna H para el DNI interno o define otra columna técnica.');
  }
  if (!dniHeader && hojaRegistro.getLastRow() > 1) {
    var existingValues = hojaRegistro.getRange(2, 8, hojaRegistro.getLastRow() - 1, 1).getDisplayValues();
    if (existingValues.some(function(row) { return String(row[0] || '').trim() !== ''; })) {
      throw new Error('La columna H de REGISTRO ya contiene datos. No se modificó el registro; libera la columna H para el DNI interno o define otra columna técnica.');
    }
  }
  if (!dniHeader) hojaRegistro.getRange(1, 8).setValue('DNI_PERSONA_SISTEMA');
  var timestampHeader = String(hojaRegistro.getRange(1, 9).getDisplayValue() || '').trim();
  if (timestampHeader && timestampHeader !== 'REGISTRADO_EN_SISTEMA') {
    throw new Error('La columna I de REGISTRO ya contiene datos. No se modificó el registro; libera la columna I para la hora interna de registro o define otra columna técnica.');
  }
  if (!timestampHeader && hojaRegistro.getLastRow() > 1) {
    var timestampValues = hojaRegistro.getRange(2, 9, hojaRegistro.getLastRow() - 1, 1).getDisplayValues();
    if (timestampValues.some(function(row) { return String(row[0] || '').trim() !== ''; })) {
      throw new Error('La columna I de REGISTRO ya contiene datos. No se modificó el registro; libera la columna I para la hora interna de registro o define otra columna técnica.');
    }
  }
  if (!timestampHeader) hojaRegistro.getRange(1, 9).setValue('REGISTRADO_EN_SISTEMA');
  hojaRegistro.hideColumns(8, 2);
}

function aplicarFormatoFechaRegistro_(hojaRegistro, firstRow, rowCount) {
  if (hojaRegistro.getLastRow() > 1) {
    hojaRegistro.getRange(2, 1, hojaRegistro.getLastRow() - 1, 1)
      .setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }
  if (rowCount > 0) {
    hojaRegistro.getRange(firstRow, 1, rowCount, 1)
      .setNumberFormat('dd/MM/yyyy HH:mm:ss');
    hojaRegistro.getRange(firstRow, 9, rowCount, 1)
      .setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }
}

function obtenerFechaRegistroInterna_(hojaRegistro, rowNumber, fechaVisita) {
  var registeredAt = hojaRegistro.getRange(rowNumber, 9).getValue();
  var value = registeredAt instanceof Date ? registeredAt : new Date(registeredAt);
  if (registeredAt && !isNaN(value.getTime())) return value;
  return fechaVisita instanceof Date ? fechaVisita : new Date(fechaVisita);
}

function obtenerTiempoEsperaRegistro_(hojaRegistro, rowNumber) {
  var displayValue = String(hojaRegistro.getRange(rowNumber, 7).getDisplayValue() || '').trim();
  if (/^EXEDIDO$/i.test(displayValue)) return 'EXEDIDO';
  var timeMatch = /^(\d{1,2}):([0-5]\d)(?::[0-5]\d)?$/.exec(displayValue);
  if (timeMatch) {
    return String(Number(timeMatch[1])).padStart(2, '0') + ':' + timeMatch[2];
  }
  var storedValue = hojaRegistro.getRange(rowNumber, 7).getValue();
  var totalMinutes = NaN;
  if (typeof storedValue === 'number' && Number.isFinite(storedValue) && storedValue >= 0) {
    totalMinutes = Math.round(storedValue * 24 * 60);
  } else if (typeof storedValue === 'string') {
    var serializedDate = /^Date\(\d{4},\d{1,2},\d{1,2},(\d{1,2}),(\d{1,2})(?:,\d{1,2})?\)/.exec(storedValue);
    if (serializedDate && Number(serializedDate[2]) < 60) {
      totalMinutes = Number(serializedDate[1]) * 60 + Number(serializedDate[2]);
    }
  }
  if (Number.isSafeInteger(totalMinutes) && totalMinutes >= 0 && totalMinutes <= 12 * 60) {
    return String(Math.floor(totalMinutes / 60)).padStart(2, '0') + ':' +
      String(totalMinutes % 60).padStart(2, '0');
  }
  if (storedValue instanceof Date) {
    var timezone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    var formatted = Utilities.formatDate(storedValue, timezone, 'HH:mm');
    if (/^\d{2}:[0-5]\d$/.test(formatted)) return formatted;
  }
  throw new Error('El tiempo de espera de la entrega original no tiene un formato válido.');
}

function registrarEntregaMostrador(codigoFarmaciaIngresado, personaIngresada, cantidadPedidos, operarioLogueado, inicioEspera) {
  return registrarEntregasMostrador(
    [{ codigoFarmacia: codigoFarmaciaIngresado, cantidadPedidos: cantidadPedidos }],
    personaIngresada, operarioLogueado, inicioEspera
  )[0];
}

function obtenerDniDeRegistro_(hojaRegistro, rowNumber, person, masterRows) {
  var saved = String(hojaRegistro.getRange(rowNumber, 8).getDisplayValue() || '').trim();
  if (/^\d{7,8}$/.test(saved)) return saved;

  var normalizedPerson = normalizarPersona(person);
  var candidate = '';
  for (var rowIndex = 1; rowIndex < masterRows.length; rowIndex++) {
    for (var personIndex = 2; personIndex <= 5; personIndex++) {
      var authorizedEntry = String(masterRows[rowIndex][personIndex] || '').trim();
      var dni = extraerDniPersona(authorizedEntry);
      var authorizedName = authorizedEntry
        .replace(dni, ' ')
        .replace(/\bD\.?N\.?I\.?\b/gi, ' ')
        .replace(/^\s*[-–—,:;]+|[-–—,:;]+\s*$/g, '')
        .trim();
      if (!dni || normalizarPersona(authorizedName) !== normalizedPerson) continue;
      if (candidate && candidate !== dni) return '';
      candidate = dni;
    }
  }
  return candidate;
}

function consultarAgregadosMostrador(rowNumberIngresado) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaRegistro = ss.getSheetByName('REGISTRO');
  var hojaMaestro = ss.getSheetByName('MAESTRO');
  if (!hojaRegistro) throw new Error('No existe la pestaña "REGISTRO".');
  if (!hojaMaestro) throw new Error('No existe la pestaña "MAESTRO".');
  var rowNumber = Number(rowNumberIngresado);
  if (!Number.isSafeInteger(rowNumber) || rowNumber < 2 || rowNumber > hojaRegistro.getLastRow()) {
    throw new Error('No se encontró la entrega original en REGISTRO.');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var original = hojaRegistro.getRange(rowNumber, 1, 1, 9).getValues()[0];
    var now = new Date();
    var originalDate = obtenerFechaRegistroInterna_(hojaRegistro, rowNumber, original[0]);
    var elapsed = now.getTime() - originalDate.getTime();
    if (String(original[4]).trim() !== 'AUTORIZADO' || !Number.isFinite(elapsed) ||
      elapsed < 0 || elapsed > 40 * 60 * 1000) {
      throw new Error('Solo se pueden agregar pedidos durante los 40 minutos posteriores a una entrega autorizada.');
    }
    var masterRows = hojaMaestro.getDataRange().getDisplayValues();
    var dni = obtenerDniDeRegistro_(hojaRegistro, rowNumber, String(original[2] || ''), masterRows);
    if (!dni) {
      throw new Error('No se pudo identificar de forma única a la persona de esta entrega. No se habilitaron agregados.');
    }
    return masterRows.slice(1).flatMap(function(row) {
      var authorized = row.slice(2, 6).some(function(value) {
        return Boolean(dni) && extraerDniPersona(value) === dni;
      });
      var code = soloDigitos(row[0]);
      if (!authorized || !code) return [];
      return [{ codigoFarmacia: code, farmacia: String(row[1] || '').trim() }];
    });
  } finally {
    lock.releaseLock();
  }
}

function registrarAgregadosMostrador(payload) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hojaRegistro = ss.getSheetByName('REGISTRO');
  var hojaMaestro = ss.getSheetByName('MAESTRO');
  if (!hojaRegistro) throw new Error('No existe la pestaña "REGISTRO".');
  if (!hojaMaestro) throw new Error('No existe la pestaña "MAESTRO".');
  var rowNumber = Number(payload.filaOriginal);
  var operario = String(payload.operarioLogueado || '').trim();
  if (!Number.isSafeInteger(rowNumber) || rowNumber < 2 || rowNumber > hojaRegistro.getLastRow()) {
    throw new Error('No se encontró la entrega original en REGISTRO.');
  }
  var entries = payload.entregas;
  if (!operario || !Array.isArray(entries) || entries.length < 1 || entries.length > 10) {
    throw new Error('Selecciona entre una y diez farmacias y vuelve a ingresar si falta el operario.');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var original = hojaRegistro.getRange(rowNumber, 1, 1, 9).getValues()[0];
    var now = new Date();
    var originalDate = obtenerFechaRegistroInterna_(hojaRegistro, rowNumber, original[0]);
    var elapsed = now.getTime() - originalDate.getTime();
    if (String(original[4]).trim() !== 'AUTORIZADO' || !Number.isFinite(elapsed) ||
      elapsed < 0 || elapsed > 40 * 60 * 1000) {
      throw new Error('Solo se pueden agregar pedidos a una entrega autorizada durante los 40 minutos posteriores.');
    }
    var masterRows = hojaMaestro.getDataRange().getDisplayValues();
    var person = String(original[2] || '');
    var dni = obtenerDniDeRegistro_(hojaRegistro, rowNumber, person, masterRows);
    if (!dni) throw new Error('No se pudo identificar de forma única a la persona de esta entrega.');
    var waitingTime = obtenerTiempoEsperaRegistro_(hojaRegistro, rowNumber);
    var seenCodes = {};
    var results = entries.map(function(entry) {
      var code = soloDigitos(entry && entry.codigoFarmacia);
      var quantity = Number(entry && entry.cantidadPedidos);
      if (!code || seenCodes[code]) throw new Error('Los códigos de farmacia deben ser válidos y no repetirse.');
      seenCodes[code] = true;
      if (!Number.isSafeInteger(quantity) || quantity < 0 || quantity > 10000) {
        throw new Error('La cantidad agregada debe ser un número entero entre 0 y 10000.');
      }
      var pharmacyName = '';
      var authorized = false;
      for (var masterIndex = 1; masterIndex < masterRows.length; masterIndex++) {
        var masterRow = masterRows[masterIndex];
        if (soloDigitos(masterRow[0]) !== code) continue;
        pharmacyName = String(masterRow[1] || '').trim();
        authorized = masterRow.slice(2, 6).some(function(value) {
          return extraerDniPersona(value) === dni;
        });
        break;
      }
      if (!authorized) throw new Error('El código ' + code + ' no está autorizado para esta persona en MAESTRO.');
      return {
        codigoFarmacia: code,
        farmacia: pharmacyName,
        cantidad: quantity,
        values: [original[0], code, person, quantity, 'AGREGADO', operario, waitingTime]
      };
    });

    prepararColumnaDniRegistro_(hojaRegistro);
    var firstRow = hojaRegistro.getLastRow() + 1;
    hojaRegistro.getRange(firstRow, 7, results.length, 1).setNumberFormat('@');
    hojaRegistro.getRange(firstRow, 1, results.length, 9)
      .setValues(results.map(function(result) { return result.values.concat([dni, now]); }));
    aplicarFormatoFechaRegistro_(hojaRegistro, firstRow, results.length);
    return results.map(function(result, index) {
      return {
        estado: 'AGREGADO',
        codigoFarmacia: result.codigoFarmacia,
        farmacia: result.farmacia,
        persona: person,
        cantidadPedidos: result.cantidad,
        tiempoEspera: String(waitingTime || ''),
        filaRegistro: firstRow + index
      };
    });
  } finally {
    lock.releaseLock();
  }
}

function registrarAgregadoMostrador(payload) {
  var result = registrarAgregadosMostrador({
    filaOriginal: payload.filaOriginal,
    operarioLogueado: payload.operarioLogueado,
    entregas: [{
      codigoFarmacia: payload.codigoFarmacia,
      cantidadPedidos: payload.cantidadPedidos
    }]
  })[0];
  return {
    estado: result.estado,
    persona: result.persona,
    farmacia: result.farmacia,
    tiempoEspera: result.tiempoEspera,
    filaRegistro: result.filaRegistro
  };
}

function enviarAvisoRetiroVale_(email, farmacia, codigo, persona, cantidad, operario, fecha, zonaHoraria) {
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    Logger.log('El retiro quedó registrado, pero no se envió el aviso: la farmacia no tiene un correo válido en la columna B.');
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
    '<p style="margin:0 0 12px;font-size:15px;line-height:1.6;">Estimado equipo de <strong>' + escaparHTML(farmacia) + '</strong>:</p>' +
    '<p style="margin:0;font-size:14px;line-height:1.6;color:#526174;">Se registró correctamente el retiro de vales correspondiente a su farmacia.</p>' +
    '</td></tr>' +
    '<tr><td style="padding:16px 32px 28px;">' +
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="100%" style="border-collapse:collapse;background-color:#f7f9fc;border:1px solid #e7edf4;border-radius:8px;">' +
    '<tr><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;font-size:13px;color:#68778a;">Código de farmacia</td><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;text-align:right;font-size:14px;font-weight:700;color:#243247;">' + escaparHTML(codigo) + '</td></tr>' +
    '<tr><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;font-size:13px;color:#68778a;">Persona autorizada</td><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;text-align:right;font-size:14px;font-weight:700;color:#243247;">' + escaparHTML(persona) + '</td></tr>' +
    '<tr><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;font-size:13px;color:#68778a;">Cantidad de vales</td><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;text-align:right;font-size:14px;font-weight:700;color:#243247;">' + escaparHTML(cantidad) + '</td></tr>' +
    '<tr><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;font-size:13px;color:#68778a;">Operario</td><td style="padding:12px 16px;border-bottom:1px solid #e7edf4;text-align:right;font-size:14px;font-weight:700;color:#243247;">' + escaparHTML(operario) + '</td></tr>' +
    '<tr><td style="padding:12px 16px;font-size:13px;color:#68778a;">Fecha y hora</td><td style="padding:12px 16px;text-align:right;font-size:14px;font-weight:700;color:#243247;">' + escaparHTML(fechaTexto) + '</td></tr>' +
    '</table></td></tr>' +
    '<tr><td style="padding:16px 32px 24px;border-top:1px solid #edf1f5;">' +
    '<p style="margin:0;text-align:center;font-size:11px;line-height:1.5;color:#8793a2;">Este es un mensaje automático del sistema de entregas. Por favor, no responda a este correo.</p>' +
    '</td></tr></table></div>';
  try {
    MailApp.sendEmail({ to: email, subject: asunto, body: mensaje, htmlBody: cuerpoHTML });
    return '';
  } catch (emailError) {
    Logger.log('El retiro quedó registrado, pero no se pudo enviar el aviso a ' + email + ': ' + emailError);
    return 'El retiro quedó registrado, pero Gmail no pudo enviar el aviso al correo de la farmacia. Revisa los permisos y la cuota de envío de Apps Script.';
  }
}

function formatearTiempoEspera_(inicioEspera, fechaRegistro) {
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

function normalizarPersona(value) {
  return String(value == null ? '' : value)
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLocaleLowerCase()
    .replace(/\s+/g, ' ');
}

function extraerDniPersona(value) {
  var match = String(value == null ? '' : value).match(/(?:^|\D)(\d{7,8})(?=\D|$)/);
  return match ? match[1] : '';
}

function soloDigitos(value) {
  return String(value == null ? '' : value).replace(/\D/g, '');
}
