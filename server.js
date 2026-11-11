'use strict';

const express = require('express');
const bodyParser = require('body-parser');
const session = require('express-session');
const crypto = require('node:crypto');
const fs = require('node:fs/promises');
const path = require('node:path');

const app = express();
const IS_PRODUCTION = process.env.NODE_ENV === 'production';
if (IS_PRODUCTION) app.set('trust proxy', 1);
const PORT = Number(process.env.PORT) || 3000;
const DATA_FILE = process.env.DATA_FILE || path.join(__dirname, 'vales-data.json');
const OPERATOR_USERNAME = process.env.OPERATOR_USERNAME || 'operario';
const OPERATOR_PASSWORD = process.env.OPERATOR_PASSWORD || (IS_PRODUCTION ? '' : 'vales2026');
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || '';
const ALLOW_OPERATOR_REGISTRATION = process.env.ALLOW_OPERATOR_REGISTRATION !== 'false';
const SESSION_SECRET = process.env.SESSION_SECRET || (IS_PRODUCTION ? '' : crypto.randomBytes(32).toString('hex'));
const REGISTRO_WRITE_URL = process.env.REGISTRO_WRITE_URL || 'https://script.google.com/macros/s/AKfycbw-kgQZNcY6IDxGpu-HAAp61YuPpnu1pGI_iyVFLYjcq6CMhh6Bnepp8dGrSjlcoLyU/exec';
const REGISTRO_API_TOKEN = process.env.REGISTRO_API_TOKEN || '';
const AUTH_SHEET_ID = process.env.AUTH_SHEET_ID || '1sZX7LQSlDEWJwRMlbz4SMFBXv2FCS1ztbxdW3EwCVT8';
const SESSION_LIFETIME = 8 * 60 * 60 * 1000;
const REMEMBER_LIFETIME = 30 * 24 * 60 * 60 * 1000;
const VISITS_SHEET_ID = process.env.VISITS_SHEET_ID || '1jSwHR5eaN7NP29-4JJSNTC2owf_ufrcB47jLMzUsNLk';
const VISITS_SHEET_TAB = process.env.VISITS_SHEET_TAB || 'MOVIMIENTOS';
const BUSINESS_TIME_ZONE = process.env.BUSINESS_TIME_ZONE || 'America/Argentina/Buenos_Aires';
const VALE_API_VERSION = '2026-10-20-maestro-multi-additions-7';

if (IS_PRODUCTION && (!OPERATOR_PASSWORD || OPERATOR_PASSWORD === 'vales2026' || !ADMIN_PASSWORD || !SESSION_SECRET || !REGISTRO_API_TOKEN)) {
  throw new Error('Production requires OPERATOR_PASSWORD, ADMIN_PASSWORD, SESSION_SECRET, and REGISTRO_API_TOKEN to be configured securely.');
}
if (!IS_PRODUCTION && OPERATOR_USERNAME === 'operario' && OPERATOR_PASSWORD === 'vales2026') {
  console.warn('AVISO: usando credenciales de demostración. Define OPERATOR_USERNAME y OPERATOR_PASSWORD antes de publicar.');
}
if (!IS_PRODUCTION && !process.env.SESSION_SECRET) {
  console.warn('AVISO: SESSION_SECRET no está definido; las sesiones se invalidarán al reiniciar el servidor.');
}

app.disable('x-powered-by');
app.use(bodyParser.json({ limit: '16kb' }));
app.use(bodyParser.urlencoded({ extended: false, limit: '16kb' }));
app.use(session({
  name: 'vale.sid',
  secret: SESSION_SECRET,
  resave: false,
  saveUninitialized: false,
  cookie: {
    httpOnly: true,
    sameSite: 'strict',
    secure: process.env.NODE_ENV === 'production',
    maxAge: SESSION_LIFETIME
  }
}));

const normalize = value => String(value == null ? '' : value)
  .trim()
  .normalize('NFD')
  .replace(/[\u0300-\u036f]/g, '')
  .toLocaleLowerCase('es')
  .replace(/\s+/g, ' ');
const digits = value => String(value == null ? '' : value).replace(/\D/g, '');
let store;
let saveQueue = Promise.resolve();

async function loadStore() {
  try {
    const contents = await fs.readFile(DATA_FILE, 'utf8');
    const parsed = JSON.parse(contents);
    if (!parsed || !Array.isArray(parsed.deliveries) || !Array.isArray(parsed.sessions)) {
      throw new Error('El archivo de datos no tiene el formato esperado.');
    }
    if (!Array.isArray(parsed.operators)) parsed.operators = [];
    if (!Array.isArray(parsed.completedVisits)) parsed.completedVisits = [];
    return parsed;
  } catch (error) {
    if (error.code === 'ENOENT') return { deliveries: [], sessions: [], operators: [], completedVisits: [] };
    throw error;
  }
}

function persistStore() {
  const snapshot = JSON.stringify(store, null, 2);
  const temporaryFile = `${DATA_FILE}.tmp`;
  const write = saveQueue.then(async () => {
    await fs.writeFile(temporaryFile, snapshot, 'utf8');
    await fs.rename(temporaryFile, DATA_FILE);
  });
  saveQueue = write.catch(() => {});
  return write;
}

function requireLogin(req, res, next) {
  if (req.session && req.session.operator) return next();
  return res.status(401).json({ error: 'Inicia sesión para continuar.' });
}

function requireAdmin(req, res, next) {
  if (req.session && req.session.operator && req.session.adminAuthenticated) return next();
  return res.status(403).json({ error: 'Desbloquea el panel con la contraseña de administrador.' });
}

function safeEqual(left, right) {
  const leftBuffer = Buffer.from(String(left));
  const rightBuffer = Buffer.from(String(right));
  return leftBuffer.length === rightBuffer.length && crypto.timingSafeEqual(leftBuffer, rightBuffer);
}

function hashPassword(password, salt) {
  return new Promise((resolve, reject) => {
    crypto.scrypt(password, salt, 64, (error, derivedKey) => {
      if (error) return reject(error);
      return resolve(derivedKey.toString('hex'));
    });
  });
}

function businessDayKey(date = new Date()) {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: BUSINESS_TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  }).formatToParts(date);
  const part = type => parts.find(value => value.type === type).value;
  return `${part('year')}-${part('month')}-${part('day')}`;
}

function spreadsheetDateKey(value) {
  const match = /^Date\((\d{4}),(\d{1,2}),(\d{1,2})/.exec(String(value || ''));
  if (!match) return '';
  return `${match[1]}-${String(Number(match[2]) + 1).padStart(2, '0')}-${String(match[3]).padStart(2, '0')}`;
}

function spreadsheetTimestamp(value) {
  const match = /^Date\((\d{4}),(\d{1,2}),(\d{1,2})(?:,(\d{1,2}),(\d{1,2}),(\d{1,2})(?:,(\d{1,3}))?)?\)$/.exec(String(value || ''));
  if (!match) return '';
  const [, year, month, day, hours = '0', minutes = '0', seconds = '0', milliseconds = '0'] = match;
  const wallClock = Date.UTC(
    Number(year), Number(month), Number(day), Number(hours), Number(minutes), Number(seconds),
    Number(String(milliseconds).padEnd(3, '0'))
  );
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: BUSINESS_TIME_ZONE,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hourCycle: 'h23'
  }).formatToParts(new Date(wallClock));
  const part = type => Number(parts.find(value => value.type === type).value);
  const timezoneOffset = Date.UTC(
    part('year'), part('month') - 1, part('day'), part('hour'), part('minute'), part('second')
  ) - wallClock;
  return new Date(wallClock - timezoneOffset).toISOString();
}

async function fetchSheetTable(sheetId, tab, query) {
  const url = new URL(`https://docs.google.com/spreadsheets/d/${encodeURIComponent(sheetId)}/gviz/tq`);
  url.searchParams.set('tqx', 'out:json');
  url.searchParams.set('sheet', tab);
  url.searchParams.set('tq', query);
  const response = await fetch(url, { cache: 'no-store' });
  if (!response.ok) throw new Error(`Google Sheets respondió HTTP ${response.status}.`);
  const text = await response.text();
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start < 0 || end < start) throw new Error('Google Sheets devolvió una respuesta que no es JSON.');
  const payload = JSON.parse(text.slice(start, end + 1));
  if (payload.status !== 'ok' || !payload.table) {
    const detail = payload.errors && payload.errors[0] && payload.errors[0].detailed_message;
    throw new Error(detail || 'Google Sheets no devolvió una tabla válida.');
  }
  return payload.table;
}

function extractDni(value) {
  const match = String(value == null ? '' : value).match(/(?:^|\D)(\d{7,8})(?=\D|$)/);
  return match ? match[1] : '';
}

function getAuthorizationsByDni(table) {
  const authorizationsByDni = new Map();
  for (const row of table.rows || []) {
    const cells = row.c || [];
    const code = digits(cells[0] && (cells[0].f ?? cells[0].v));
    if (!code) continue;
    const pharmacy = String(cells[1] && (cells[1].v ?? cells[1].f) || '').trim();
    for (const cell of cells.slice(2, 6)) {
      const value = String(cell && (cell.v ?? cell.f) || '').trim();
      const dni = extractDni(value);
      if (!dni) continue;
      const name = value.replace(dni, ' ')
        .replace(/\bD\.?N\.?I\.?\b/gi, ' ')
        .replace(/[\s,;:–—-]+$/g, '')
        .trim();
      const authorizations = authorizationsByDni.get(dni) || [];
      if (!authorizations.some(item => item.code === code)) authorizations.push({ code, pharmacy, name });
      authorizationsByDni.set(dni, authorizations);
    }
  }
  return authorizationsByDni;
}

async function getWaitingVisits(filters = {}) {
  const codeFilter = digits(filters.code);
  const dniFilter = digits(filters.dni);
  const visitIdFilter = String(filters.visitId || '').trim();
  if (!codeFilter && !dniFilter && !visitIdFilter) return [];

  const day = businessDayKey();
  const [table, master] = await Promise.all([
    fetchSheetTable(
      VISITS_SHEET_ID,
      VISITS_SHEET_TAB,
      `select A,B,D,E,I where A >= datetime '${day} 00:00:00'`
    ),
    fetchSheetTable(AUTH_SHEET_ID, 'MAESTRO', 'select A,B,C,D,E,F')
  ]);
  const authorizedCodesByDni = getAuthorizationsByDni(master);

  return (table.rows || []).flatMap(row => {
    const cells = row.c || [];
    const dateCell = cells[0] || {};
    const person = String((cells[1] && (cells[1].v ?? cells[1].f)) || '').trim();
    const sector = String((cells[2] && (cells[2].v ?? cells[2].f)) || '').trim();
    const host = String((cells[3] && (cells[3].v ?? cells[3].f)) || '').trim();
    const code = digits((cells[4] && (cells[4].f ?? cells[4].v)) || '');
    const dateKey = spreadsheetDateKey(dateCell.v);
    const waitingSince = spreadsheetTimestamp(dateCell.v);
    if (!person || dateKey !== day ||
      (!normalize(sector).includes('mostrador') && !normalize(host).includes('mostrador'))) return [];
    const dni = extractDni(person);
    const authorizations = dni ? (authorizedCodesByDni.get(dni) || []) : [];
    const matchesCode = !codeFilter || code === codeFilter || authorizations.some(item => item.code === codeFilter);
    const matchesDni = !dniFilter || dni === dniFilter;
    if (!matchesCode || !matchesDni) return [];
    const identity = [dateCell.v, person, sector, host].join('\u001f');
    const id = crypto.createHash('sha256').update(identity).digest('hex').slice(0, 24);
    if (visitIdFilter && id !== visitIdFilter) return [];
    if (store.completedVisits.includes(id)) return [];
    const authorized = codeFilter
      ? authorizations.some(item => item.code === codeFilter)
      : authorizations.length > 0;
    return [{
      id,
      person,
      dni,
      code,
      authorized,
      authorizedForCode: codeFilter ? authorizations.some(item => item.code === codeFilter) : false,
      authorizations,
      time: dateCell.f || '',
      waitingSince,
      sector,
      host
    }];
  });
}

async function getAuthorizedPeopleForCode(code) {
  const table = await fetchSheetTable(AUTH_SHEET_ID, 'MAESTRO', 'select A,B,C,D,E,F');
  const authorizationsByDni = getAuthorizationsByDni(table);
  const people = new Map();
  for (const row of table.rows || []) {
    const cells = row.c || [];
    const rowCode = digits(cells[0] && (cells[0].f ?? cells[0].v));
    if (rowCode !== code) continue;
    const pharmacy = String(cells[1] && (cells[1].v ?? cells[1].f) || '').trim();
    for (const cell of cells.slice(2, 6)) {
      const person = String(cell && (cell.v ?? cell.f) || '').trim();
      const dni = extractDni(person);
      if (!person || !dni) continue;
      const authorizations = authorizationsByDni.get(dni) || [];
      people.set(dni, {
        dni,
        person,
        name: authorizations.find(item => item.code === code)?.name || person,
        code,
        pharmacy,
        authorizations
      });
    }
  }
  return [...people.values()];
}

async function createOperatorSession(req, username, remember) {
  await new Promise((resolve, reject) => {
    req.session.regenerate(error => error ? reject(error) : resolve());
  });
  req.session.operator = username;
  req.session.cookie.maxAge = remember ? REMEMBER_LIFETIME : SESSION_LIFETIME;
  store.sessions.push({ date: new Date().toISOString(), operator: username, remember: Boolean(remember) });
  await persistStore();
  await new Promise((resolve, reject) => {
    req.session.save(error => error ? reject(error) : resolve());
  });
}

async function spreadsheetWriteRequest(payload) {
  if (!REGISTRO_WRITE_URL || !REGISTRO_API_TOKEN) {
    const error = new Error('El guardado en Google Sheets no está configurado. Publica GoogleAppsScript.gs como aplicación web y define REGISTRO_WRITE_URL y REGISTRO_API_TOKEN antes de registrar entregas.');
    error.status = 503;
    throw error;
  }
  const endpointStatus = await verifySpreadsheetEndpoint();
  if (!endpointStatus.spreadsheetConnected) {
    const error = new Error(
      'No se envió el registro porque Apps Script no reconoce la versión necesaria para guardar el tiempo de espera. ' +
      (endpointStatus.spreadsheetError || 'Actualiza el despliegue y verifica REGISTRO_WRITE_URL.') +
      ' No vuelvas a intentar hasta actualizar el despliegue.'
    );
    error.status = 503;
    throw error;
  }
  const response = await fetch(REGISTRO_WRITE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
    body: JSON.stringify({ ...payload, apiToken: REGISTRO_API_TOKEN }),
    cache: 'no-store'
  });
  const text = await response.text();
  let result;
  try {
    result = JSON.parse(text);
  } catch {
    throw new Error(`El endpoint de Google Sheets no devolvió JSON válido (HTTP ${response.status}).`);
  }
  if (!response.ok || result.ok !== true) {
    if (result.error === 'Acción no válida.') {
      const error = new Error(
        'Apps Script está accesible, pero su implementación activa no incluye el envío de reportes. ' +
        'Actualiza la implementación existente con el código actual de GoogleAppsScript.gs: ' +
        'Implementar > Gestionar implementaciones > Editar > Nueva versión > Implementar.'
      );
      error.status = 503;
      throw error;
    }
    throw new Error(result.error || `Google Sheets no confirmó el registro (HTTP ${response.status}).`);
  }
  return result;
}

async function writeDeliveriesToSpreadsheet(delivery) {
  const result = await spreadsheetWriteRequest({
    action: 'registrarEntregas',
    persona: delivery.person,
    operarioLogueado: delivery.operator,
    inicioEspera: delivery.waitingSince,
    entregas: delivery.entries.map(entry => ({
      codigoFarmacia: entry.code,
      cantidadPedidos: entry.quantity
    }))
  });
  if (!Array.isArray(result.entregas) || result.entregas.length !== delivery.entries.length ||
    result.entregas.some(item => !item || !['AUTORIZADO', 'DENEGADO'].includes(item.estado) ||
      !/^(?:EXEDIDO|(?:0\d|1[01]):[0-5]\d|12:00)$/.test(String(item.tiempoEspera || '')))) {
    throw new Error('Google Sheets pudo haber guardado la entrega pero no confirmó el tiempo de espera. Verifica REGISTRO antes de volver a intentarlo.');
  }
  return result.entregas;
}

async function getAdditionOptionsFromSpreadsheet(rowNumber) {
  const result = await spreadsheetWriteRequest({
    action: 'consultarAgregados',
    filaOriginal: rowNumber
  });
  if (!Array.isArray(result.entregas)) {
    throw new Error('Google Sheets no devolvió las farmacias autorizadas para editar esta entrega.');
  }
  return result.entregas;
}

async function writeAdditionsToSpreadsheet(addition) {
  const result = await spreadsheetWriteRequest({
    action: 'registrarAgregados',
    filaOriginal: addition.rowNumber,
    entregas: addition.entries.map(entry => ({
      codigoFarmacia: entry.code,
      cantidadPedidos: entry.quantity
    })),
    operarioLogueado: addition.operator
  });
  if (!Array.isArray(result.entregas) || result.entregas.length !== addition.entries.length ||
    result.entregas.some(item => !item || item.estado !== 'AGREGADO' ||
      !/^(?:EXEDIDO|(?:0\d|1[01]):[0-5]\d|12:00)$/.test(String(item.tiempoEspera || '')))) {
    throw new Error('Google Sheets no confirmó todos los agregados. Verifica REGISTRO antes de volver a intentarlo.');
  }
  return result.entregas;
}

async function readDeliveriesFromSpreadsheet() {
  const table = await fetchSheetTable(AUTH_SHEET_ID, 'REGISTRO', 'select *');
  return (table.rows || []).flatMap((row, index) => {
    const cells = row.c || [];
    const values = Array.from({ length: 9 }, (_, index) => {
      const cell = cells[index];
      return cell ? (cell.v ?? cell.f ?? '') : '';
    });
    if (!values.some(value => value !== '' && value != null)) return [];
    if (normalize(values[0]).includes('fecha') && normalize(values[1]).includes('codigo')) return [];

    const date = spreadsheetDateTimeValue(values[0]);
    const registeredAt = spreadsheetDateTimeValue(values[8]);

    return [{
      date,
      dateHasTime: spreadsheetDateHasTime(values[0]),
      registeredAt,
      code: String(values[1] == null ? '' : values[1]),
      person: String(values[2] == null ? '' : values[2]),
      quantity: Number(values[3]) || 0,
      state: String(values[4] == null ? '' : values[4]),
      operator: String(values[5] == null ? '' : values[5]),
      waitingTime: formatSpreadsheetWaitingTime(values[6]),
      rowNumber: index + 2
    }];
  }).reverse();
}

function spreadsheetDateTimeValue(value) {
  const rawDate = String(value || '');
  if (/^Date\(\d{4},\d{1,2},\d{1,2}/.test(rawDate)) return spreadsheetTimestamp(rawDate);
  if (rawDate && !Number.isNaN(new Date(rawDate).getTime())) return new Date(rawDate).toISOString();
  return '';
}

function spreadsheetDateHasTime(value) {
  const rawDate = String(value || '');
  if (/^Date\(\d{4},\d{1,2},\d{1,2},\d{1,2},\d{1,2},\d{1,2}/.test(rawDate)) return true;
  return /^\d{4}-\d{2}-\d{2}[T ]\d{2}:\d{2}/.test(rawDate);
}

async function verifySpreadsheetEndpoint() {
  if (!REGISTRO_WRITE_URL || !REGISTRO_API_TOKEN) {
    return {
      spreadsheetConfigured: false,
      spreadsheetConnected: false,
      spreadsheetError: 'Falta configurar REGISTRO_API_TOKEN. Usa el mismo valor que la propiedad VALE_API_TOKEN de Apps Script.'
    };
  }
  try {
    const response = await fetch(REGISTRO_WRITE_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
      body: JSON.stringify({ action: 'verificarConexion', apiToken: REGISTRO_API_TOKEN }),
      cache: 'no-store'
    });
    const text = await response.text();
    let result;
    try {
      result = JSON.parse(text);
    } catch {
      throw new Error(`El despliegue respondió algo que no es JSON (HTTP ${response.status}). Actualiza la aplicación web con el router unificado.`);
    }
    if (!response.ok || result.ok !== true || result.version !== VALE_API_VERSION) {
      const deployedVersion = result.version
        ? ` La versión publicada que respondió es "${result.version}".`
        : ' La respuesta no incluye un número de versión.';
      throw new Error(
        (result.error || `El despliegue no reconoce la versión unificada (HTTP ${response.status}).`) +
        deployedVersion + ' Actualiza la aplicación web desde Apps Script y verifica REGISTRO_WRITE_URL.'
      );
    }
    return {
      spreadsheetConfigured: true,
      spreadsheetConnected: true,
      spreadsheetVersion: result.version
    };
  } catch (error) {
    return {
      spreadsheetConfigured: true,
      spreadsheetConnected: false,
      spreadsheetError: error.message
    };
  }
}

function formatSpreadsheetWaitingTime(value) {
  const raw = String(value == null ? '' : value).trim();
  if (!raw) return '';
  if (raw.toLocaleUpperCase('es') === 'EXEDIDO') return 'EXEDIDO';

  const dateValue = /^Date\((\d{4}),(\d{1,2}),(\d{1,2})(?:,(\d{1,2}),(\d{1,2}),(\d{1,2}))?\)$/.exec(raw);
  let totalSeconds;
  if (dateValue) {
    const [, year, month, day, hours = '0', minutes = '0', seconds = '0'] = dateValue;
    const start = Date.UTC(1899, 11, 30);
    const end = Date.UTC(Number(year), Number(month), Number(day));
    totalSeconds = Math.floor((end - start) / 1000) + Number(hours) * 3600 +
      Number(minutes) * 60 + Number(seconds);
  } else {
    const timeValue = /^(\d+):([0-5]\d)(?::([0-5]\d))?$/.exec(raw);
    if (!timeValue) return raw;
    totalSeconds = Number(timeValue[1]) * 3600 + Number(timeValue[2]) * 60 + Number(timeValue[3] || 0);
  }

  if (!Number.isSafeInteger(totalSeconds) || totalSeconds < 0) return raw;
  if (totalSeconds > 12 * 60 * 60) return 'EXEDIDO';
  const hours = String(Math.floor(totalSeconds / 3600)).padStart(2, '0');
  const minutes = String(Math.floor((totalSeconds % 3600) / 60)).padStart(2, '0');
  return `${hours}:${minutes}`;
}

app.get('/', (req, res) => res.set('Cache-Control', 'no-store').type('html').send(HTML));

app.get('/api/session', (req, res) => {
  res.json({ operator: req.session.operator || null });
});

app.post('/api/login', async (req, res, next) => {
  const username = String(req.body.username || '').trim();
  const password = String(req.body.password || '');
  if (!username || !password) return res.status(400).json({ error: 'Completa usuario y contraseña.' });
  const account = store.operators.find(operator => normalize(operator.username) === normalize(username));
  let valid = false;
  try {
    if (account) {
      const submittedHash = await hashPassword(password, account.passwordSalt);
      valid = safeEqual(submittedHash, account.passwordHash);
    } else {
      valid = safeEqual(username, OPERATOR_USERNAME) && safeEqual(password, OPERATOR_PASSWORD);
    }
  } catch (error) {
    return next(error);
  }
  if (!valid) {
    return res.status(401).json({ error: 'Usuario o contraseña incorrectos.' });
  }

  try {
    await createOperatorSession(req, account ? account.username : OPERATOR_USERNAME, req.body.remember);
    return res.json({ operator: req.session.operator });
  } catch (error) {
    return next(error);
  }
});

app.post('/api/register', async (req, res, next) => {
  if (!ALLOW_OPERATOR_REGISTRATION) {
    return res.status(403).json({ error: 'El registro de operarios está deshabilitado.' });
  }
  const username = String(req.body.username || '').trim();
  const password = String(req.body.password || '');
  if (!/^[a-zA-Z0-9._-]{3,32}$/.test(username)) {
    return res.status(400).json({ error: 'El usuario debe tener 3–32 caracteres: letras, números, punto, guion o guion bajo.' });
  }
  if (password.length < 8 || password.length > 128) {
    return res.status(400).json({ error: 'La contraseña debe tener entre 8 y 128 caracteres.' });
  }
  const duplicate = store.operators.some(operator => normalize(operator.username) === normalize(username)) ||
    normalize(username) === normalize(OPERATOR_USERNAME);
  if (duplicate) return res.status(409).json({ error: 'Ese nombre de operario ya está registrado.' });

  try {
    const passwordSalt = crypto.randomBytes(16).toString('hex');
    const passwordHash = await hashPassword(password, passwordSalt);
    if (store.operators.some(operator => normalize(operator.username) === normalize(username))) {
      return res.status(409).json({ error: 'Ese nombre de operario ya está registrado.' });
    }
    store.operators.push({
      username,
      passwordSalt,
      passwordHash,
      createdAt: new Date().toISOString()
    });
    await persistStore();
    await createOperatorSession(req, username, false);
    return res.status(201).json({ operator: username });
  } catch (error) {
    return next(error);
  }
});

app.post('/api/logout', requireLogin, (req, res, next) => {
  req.session.destroy(error => {
    if (error) return next(error);
    res.clearCookie('vale.sid', { httpOnly: true, sameSite: 'strict', secure: process.env.NODE_ENV === 'production' });
    res.json({ ok: true });
  });
});

app.get('/api/admin/session', requireLogin, (req, res) => {
  res.set('Cache-Control', 'no-store').json({ authenticated: Boolean(req.session.adminAuthenticated) });
});

app.post('/api/admin/unlock', requireLogin, (req, res) => {
  if (!ADMIN_PASSWORD) {
    return res.status(503).json({ error: 'El panel de administrador está bloqueado: configura ADMIN_PASSWORD en el entorno del servidor y reinícialo.' });
  }
  const now = Date.now();
  if (req.session.adminLockoutUntil > now) {
    return res.status(429).json({ error: 'Demasiados intentos. Espera unos minutos antes de volver a intentar.' });
  }
  const submittedPassword = String(req.body.password || '');
  if (!safeEqual(submittedPassword, ADMIN_PASSWORD)) {
    req.session.adminFailedAttempts = (req.session.adminFailedAttempts || 0) + 1;
    if (req.session.adminFailedAttempts >= 5) {
      req.session.adminFailedAttempts = 0;
      req.session.adminLockoutUntil = now + 5 * 60 * 1000;
    }
    return res.status(401).json({ error: 'Contraseña de administrador incorrecta.' });
  }
  req.session.adminAuthenticated = true;
  req.session.adminFailedAttempts = 0;
  delete req.session.adminLockoutUntil;
  return res.json({ authenticated: true });
});

app.post('/api/admin/lock', requireLogin, (req, res) => {
  delete req.session.adminAuthenticated;
  return res.json({ authenticated: false });
});

app.get('/api/deliveries', requireLogin, async (req, res, next) => {
  try {
    return res.json(await readDeliveriesFromSpreadsheet());
  } catch (error) {
    return next(error);
  }
});

app.post('/api/deliveries/:rowNumber/report', requireLogin, async (req, res, next) => {
  const rowNumber = Number(req.params.rowNumber);
  const email = String(req.body.email || '').trim();
  if (!Number.isSafeInteger(rowNumber) || rowNumber < 2) {
    return res.status(400).json({ error: 'La fila de la entrega no es válida.' });
  }
  if (email.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    return res.status(400).json({ error: 'Ingresa una dirección de correo electrónico válida para gerencia.' });
  }

  try {
    const result = await spreadsheetWriteRequest({
      action: 'enviarReporteGerencial',
      filaRegistro: rowNumber,
      email,
      solicitadoPor: req.session.operator
    });
    return res.json({ sent: true, email: result.email });
  } catch (error) {
    return next(error);
  }
});

app.get('/api/status', requireLogin, requireAdmin, async (req, res) => {
  const spreadsheet = await verifySpreadsheetEndpoint();
  res.set('Cache-Control', 'no-store').json({
    ...spreadsheet,
    localFile: path.basename(DATA_FILE),
    authorizationFormUrl: 'https://script.google.com/macros/s/AKfycbw-kgQZNcY6IDxGpu-HAAp61YuPpnu1pGI_iyVFLYjcq6CMhh6Bnepp8dGrSjlcoLyU/exec',
    spreadsheetUrl: 'https://docs.google.com/spreadsheets/d/1sZX7LQSlDEWJwRMlbz4SMFBXv2FCS1ztbxdW3EwCVT8/edit?usp=sharing'
  });
});

app.get('/api/waiting-visits', requireLogin, async (req, res, next) => {
  try {
    return res.json(await getWaitingVisits({ code: req.query.code, dni: req.query.dni }));
  } catch (error) {
    return next(error);
  }
});

app.get('/api/authorized-people', requireLogin, async (req, res, next) => {
  const code = digits(req.query.code);
  if (!code || code.length > 5) {
    return res.status(400).json({ error: 'Ingresa un código de farmacia válido para consultar los autorizados.' });
  }
  try {
    return res.json({ people: await getAuthorizedPeopleForCode(code) });
  } catch (error) {
    return next(error);
  }
});

app.get('/api/deliveries/:rowNumber/additions', requireLogin, async (req, res, next) => {
  const rowNumber = Number(req.params.rowNumber);
  if (!Number.isSafeInteger(rowNumber) || rowNumber < 2) {
    return res.status(400).json({ error: 'La fila de la entrega no es válida.' });
  }

  try {
    const history = await readDeliveriesFromSpreadsheet();
    const original = history.find(delivery => delivery.rowNumber === rowNumber);
    const elapsed = original ? Date.now() - Date.parse(original.date) : NaN;
    if (!original || original.state !== 'AUTORIZADO' || !Number.isFinite(elapsed) ||
      elapsed < 0 || elapsed > 40 * 60 * 1000) {
      return res.status(409).json({ error: 'El plazo de edición de 40 minutos venció o la entrega original no está disponible.' });
    }
    const authorizations = await getAdditionOptionsFromSpreadsheet(rowNumber);
    return res.json({ authorizations });
  } catch (error) {
    return next(error);
  }
});

app.post('/api/deliveries', requireLogin, async (req, res, next) => {
  const visitId = String(req.body.visitId || '').trim();
  const manualEntry = req.body.manualEntry === true;
  const manualCode = digits(req.body.manualCode);
  const personDni = digits(req.body.personDni);
  const entries = Array.isArray(req.body.entries) ? req.body.entries.map(entry => ({
    code: digits(entry && entry.code),
    quantity: Number(entry && entry.quantity)
  })) : [];
  if ((!visitId && !manualEntry) || (manualEntry && (!manualCode || manualCode.length > 5 || !/^\d{7,8}$/.test(personDni))) ||
    entries.length < 1 || entries.length > 10 ||
    entries.some(entry => !entry.code || !Number.isSafeInteger(entry.quantity) || entry.quantity < 0 || entry.quantity > 10000) ||
    new Set(entries.map(entry => entry.code)).size !== entries.length) {
    return res.status(400).json({ error: 'Selecciona una persona, al menos una farmacia y una cantidad válida por código (0–10000).' });
  }

  try {
    let waitingVisit;
    if (manualEntry) {
      const person = (await getAuthorizedPeopleForCode(manualCode)).find(item => item.dni === personDni);
      if (!person) {
        return res.status(409).json({ error: 'La persona ya no figura autorizada para ese código en MAESTRO. Actualiza la lista.' });
      }
      waitingVisit = {
        id: `manual-${crypto.randomUUID()}`,
        person: person.person,
        displayPerson: person.name,
        dni: person.dni,
        code: manualCode,
        authorizations: person.authorizations,
        waitingSince: new Date().toISOString(),
        time: 'INGRESO REGISTRADO AHORA',
        manualEntry: true
      };
    } else {
      waitingVisit = (await getWaitingVisits({ visitId })).find(visit => visit.id === visitId);
      if (!waitingVisit) {
        return res.status(409).json({ error: 'Esta visita ya no está pendiente en la lista de hoy. Actualiza la lista antes de continuar.' });
      }
    }
    const authorizedCodes = new Set(waitingVisit.authorizations.map(item => item.code));
    const unauthorizedEntries = entries.filter(entry => !authorizedCodes.has(entry.code));
    const requestedCode = digits(req.body.requestedCode);
    if (manualEntry && unauthorizedEntries.length) {
      return res.status(403).json({ error: 'La persona seleccionada no está autorizada para uno o más códigos elegidos.' });
    }
    if (!manualEntry && unauthorizedEntries.length &&
      (entries.length !== 1 || unauthorizedEntries[0].code !== requestedCode)) {
      return res.status(400).json({ error: 'Solo puedes registrar como denegado el código exacto utilizado en la búsqueda.' });
    }
    if (!waitingVisit.waitingSince) {
      return res.status(409).json({ error: 'La visita seleccionada no tiene una fecha y hora de ingreso válidas en MOVIMIENTOS. Corrige la fila antes de registrarla.' });
    }
    const results = await writeDeliveriesToSpreadsheet({
      entries, person: waitingVisit.person, operator: req.session.operator,
      waitingSince: waitingVisit.waitingSince
    });
    const deliveries = results.map((sheetResult, index) => ({
      date: new Date().toISOString(),
      code: entries[index].code,
      person: sheetResult.persona || waitingVisit.person,
      quantity: sheetResult.estado === 'AUTORIZADO' ? entries[index].quantity : 0,
      state: sheetResult.estado,
      operator: req.session.operator,
      waitingTime: sheetResult.tiempoEspera || '',
      pharmacy: sheetResult.farmacia || '',
      emailWarning: sheetResult.avisoCorreo || '',
      visitId,
      spreadsheetSynced: true
    }));
    store.deliveries.push(...deliveries);
    if (visitId && !manualEntry) store.completedVisits.push(visitId);
    await persistStore();
    return res.status(201).json({ deliveries });
  } catch (error) {
    return next(error);
  }
});

app.post('/api/deliveries/:rowNumber/addition', requireLogin, async (req, res, next) => {
  const rowNumber = Number(req.params.rowNumber);
  const entries = Array.isArray(req.body.entries) ? req.body.entries.map(entry => ({
    code: digits(entry && entry.code),
    quantity: Number(entry && entry.quantity)
  })) : [];
  if (!Number.isSafeInteger(rowNumber) || rowNumber < 2 ||
    entries.length < 1 || entries.length > 10 ||
    entries.some(entry => !entry.code || !Number.isSafeInteger(entry.quantity) || entry.quantity < 0 || entry.quantity > 10000) ||
    new Set(entries.map(entry => entry.code)).size !== entries.length) {
    return res.status(400).json({ error: 'Selecciona una o más farmacias sin repetir y define una cantidad válida por código.' });
  }

  try {
    const history = await readDeliveriesFromSpreadsheet();
    const original = history.find(delivery => delivery.rowNumber === rowNumber);
    const elapsed = original ? Date.now() - Date.parse(original.date) : NaN;
    if (!original || original.state !== 'AUTORIZADO' || !Number.isFinite(elapsed) || elapsed < 0 || elapsed > 40 * 60 * 1000) {
      return res.status(409).json({ error: 'El plazo de edición de 40 minutos venció o la entrega original no está disponible. Actualiza el historial.' });
    }
    const sheetResults = await writeAdditionsToSpreadsheet({
      rowNumber, entries, operator: req.session.operator
    });
    const deliveries = sheetResults.map((sheetResult, index) => ({
      date: new Date().toISOString(),
      code: entries[index].code,
      person: sheetResult.persona || original.person,
      quantity: entries[index].quantity,
      state: sheetResult.estado,
      operator: req.session.operator,
      waitingTime: sheetResult.tiempoEspera,
      rowNumber: sheetResult.filaRegistro,
      addition: true,
      spreadsheetSynced: true
    }));
    store.deliveries.push(...deliveries);
    await persistStore();
    return res.status(201).json({ deliveries });
  } catch (error) {
    return next(error);
  }
});

app.use((error, req, res, next) => {
  console.error(error);
  if (res.headersSent) return next(error);
  return res.status(error.status || 500).json({ error: error.message || 'Error interno del servidor.' });
});

const HTML = String.raw`<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <meta name="theme-color" content="#071014">
  <title>Vale // Control de mostrador</title>
  <style>
    :root{color-scheme:dark;--bg:#071014;--panel:#0c191f;--panel2:#102128;--line:#1b3942;--text:#d8f4f1;--muted:#78969a;--green:#00ffcc;--blue:#0099ff;--red:#ff3366;--mono:"Cascadia Code","SFMono-Regular",Consolas,monospace}
    *{box-sizing:border-box}body{margin:0;min-height:100vh;background:radial-gradient(ellipse at 50% -20%,#10343a 0,transparent 55%),linear-gradient(135deg,#071014,#08151a 55%,#071014);color:var(--text);font-family:var(--mono);font-size:14px}
    body:before{position:fixed;inset:0;pointer-events:none;content:"";opacity:.09;background:repeating-linear-gradient(0deg,transparent 0 3px,#80fff3 4px);z-index:5}
    .shell{width:min(1120px,calc(100% - 36px));margin:auto}.topbar{border-bottom:1px solid var(--line);background:#071014d9}.top-inner{min-height:72px;display:flex;align-items:center;justify-content:space-between;gap:20px}.brand{display:flex;align-items:center;gap:12px;color:var(--green);font-weight:800;letter-spacing:.14em}.brand-mark{width:30px;height:30px;border:1px solid var(--green);display:grid;place-items:center;box-shadow:0 0 13px #00ffcc42}.top-meta{font-size:11px;color:var(--muted);letter-spacing:.1em}.live{color:var(--green)}
    main{padding:46px 0 70px}.eyebrow{color:var(--blue);font-size:11px;letter-spacing:.2em;text-transform:uppercase}.hero{display:flex;justify-content:space-between;align-items:end;gap:20px;margin-bottom:28px}.hero h1{margin:8px 0 0;font-size:clamp(25px,5vw,42px);letter-spacing:.05em;line-height:1.1}.hero p{color:var(--muted);max-width:600px;line-height:1.7;margin:10px 0 0;font-size:12px}.panel{background:linear-gradient(145deg,#0d1b21ed,#091419ed);border:1px solid var(--line);box-shadow:0 10px 34px #0005,0 0 20px #0099ff08;position:relative;padding:24px;margin:18px 0}.panel:before{content:"";position:absolute;left:-1px;top:-1px;width:50px;height:2px;background:var(--blue);box-shadow:0 0 10px var(--blue)}.panel h2{font-size:14px;letter-spacing:.12em;margin:0 0 20px;text-transform:uppercase}.grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:17px}.field label{display:block;color:#91b5b7;font-size:11px;letter-spacing:.1em;margin-bottom:8px;text-transform:uppercase}.field input{width:100%;border:1px solid #23434c;background:#071217;color:var(--text);padding:13px 14px;font:inherit;outline:none;transition:border-color .2s,box-shadow .2s}.field input:focus{border-color:var(--blue);box-shadow:0 0 0 2px #0099ff24,0 0 14px #0099ff1a}.actions{display:flex;gap:10px;align-items:center;flex-wrap:wrap;margin-top:21px}.button{font:inherit;font-weight:700;letter-spacing:.08em;padding:12px 17px;border:1px solid;cursor:pointer;transition:transform .18s,box-shadow .18s,background .18s}.button:hover{transform:translateY(-2px)}.button:active{transform:translateY(1px)}.button:disabled{opacity:.55;cursor:wait;transform:none}.primary{color:#001b16;border-color:var(--green);background:var(--green);box-shadow:0 0 16px #00ffcc30}.primary:hover{box-shadow:0 0 25px #00ffcc70}.secondary{background:transparent;color:#a9c8ca;border-color:#31525b}.secondary:hover{border-color:var(--blue);box-shadow:0 0 15px #0099ff30}.status{display:none;margin-top:20px;padding:17px 19px;border:1px solid;line-height:1.6}.status.show{display:block}.status strong{display:block;font-size:15px;letter-spacing:.12em}.status span{display:block;margin-top:5px;font-size:12px}.status.ok{border-color:var(--green);color:var(--green);background:#00ffcc0d;box-shadow:inset 3px 0 var(--green),0 0 16px #00ffcc12}.status.denied{position:fixed;z-index:20;isolation:isolate;top:50%;left:50%;width:min(680px,calc(100% - 32px));margin:0;padding:clamp(24px,6vw,48px);transform:translate(-50%,-50%);border:2px solid #ff557d;color:#fff;background:linear-gradient(145deg,#291019,#140b11 72%);box-shadow:inset 8px 0 var(--red),0 0 0 1px #ff33664d,0 0 55px #ff174799,0 24px 90px #000;animation:denied-enter .65s cubic-bezier(.2,.9,.25,1.3) both,denied-glow 1.4s ease-in-out .65s 4 alternate}.status.denied:before{position:fixed;z-index:-1;inset:50% auto auto 50%;width:100vw;height:100vh;transform:translate(-50%,-50%);content:"";background:#020609dc;backdrop-filter:blur(5px)}.status.denied strong{display:flex;align-items:center;gap:16px;color:#ff557d;font-size:clamp(38px,9vw,64px);line-height:1.05;letter-spacing:.13em;text-shadow:0 0 20px #ff174799}.status.denied strong:before{content:"⚠";font-size:.9em;animation:denied-icon .55s ease-in-out 3}.status.denied span{margin-top:20px;color:#ffe4ea;font-size:clamp(14px,3vw,18px);line-height:1.7}.status.denied .denied-close{display:block;margin:26px 0 0 auto;padding:12px 20px;border:1px solid #ff557d;background:#ff174722;color:#fff;font:inherit;font-weight:800;letter-spacing:.12em;cursor:pointer;box-shadow:0 0 18px #ff17474d}.status.denied .denied-close:hover{background:#ff174755;box-shadow:0 0 28px #ff174799}.status.info{border-color:var(--blue);color:#80caff;background:#0099ff0b}@keyframes denied-enter{0%{opacity:0;transform:translate(-50%,-44%) scale(.82)}65%{opacity:1;transform:translate(-50%,-52%) scale(1.025)}100%{opacity:1;transform:translate(-50%,-50%) scale(1)}}@keyframes denied-glow{to{box-shadow:inset 8px 0 var(--red),0 0 0 1px #ff336699,0 0 75px #ff1747cc,0 24px 90px #000}}@keyframes denied-icon{50%{transform:scale(1.2) rotate(-8deg)}}@media(prefers-reduced-motion:reduce){.status.denied,.status.denied strong:before{animation:none}}
    .table-head{display:flex;align-items:center;justify-content:space-between;gap:16px;margin-bottom:16px}.table-head h2{margin:0}.filter{width:min(260px,100%);padding:10px 12px;border:1px solid #23434c;background:#071217;color:var(--text);font:inherit;outline:none}.field select{width:100%;border:1px solid #23434c;background:#071217;color:var(--text);padding:13px 14px;font:inherit;outline:none}.field select:focus{border-color:var(--blue);box-shadow:0 0 0 2px #0099ff24}.table-wrap{overflow:auto;border:1px solid #18323a}table{border-collapse:collapse;width:100%;min-width:720px}th,td{text-align:left;padding:12px 14px;border-bottom:1px solid #19323a;font-size:11px}th{color:var(--blue);letter-spacing:.1em;text-transform:uppercase;background:#0b1b21}td{color:#b8d2d1}tr:last-child td{border-bottom:0}.state{font-weight:800;letter-spacing:.08em}.state.authorized{color:var(--green)}.state.denied{color:var(--red)}.empty{text-align:center;color:var(--muted);padding:26px}.badge{font-size:10px;color:var(--green);border:1px solid #00ffcc42;padding:7px 10px;letter-spacing:.08em}.hidden{display:none!important}.login-wrap{max-width:480px;margin:48px auto}.login-wrap .panel{padding:30px}.login-title{font-size:22px;letter-spacing:.1em;margin:9px 0 8px}.login-copy{color:var(--muted);font-size:12px;line-height:1.6;margin:0 0 22px}.remember{display:flex;gap:9px;align-items:center;color:#9ebbbc;font-size:11px;margin-top:15px}.remember input{accent-color:var(--green)}.login-error{color:var(--red);font-size:12px;min-height:18px;margin-top:12px}.logout{color:var(--red);border-color:#ff33665c}.processing{position:fixed;z-index:10;inset:0;background:#050c0feF;display:grid;place-items:center;padding:20px}.processing-box{width:min(480px,100%);border:1px solid #1f4c56;background:#0b1b20;padding:28px;box-shadow:0 0 40px #0099ff1a}.processing-title{color:var(--green);font-size:12px;letter-spacing:.15em;margin-bottom:17px}.processing-message{min-height:20px;color:#c2e8e7;font-size:12px}.progress{height:5px;background:#1b343b;margin-top:18px;overflow:hidden}.progress span{display:block;width:35%;height:100%;background:var(--green);box-shadow:0 0 12px var(--green);animation:progress 1.2s ease-in-out infinite alternate}@keyframes progress{from{transform:translateX(-30%)}to{transform:translateX(220%)}}.footer{color:#547076;text-align:center;font-size:10px;margin-top:28px;letter-spacing:.08em}
    .guard-notice{margin:18px 0 0;padding:14px 16px;border:1px solid #8a6b32;background:#211b0d;color:#ffe3a3;font-size:12px;line-height:1.6;text-align:center}
    .history-actions{display:flex;align-items:center;gap:8px}.report-button{display:grid;place-items:center;width:40px;height:40px;padding:0;border:1px solid #6c8ca0;border-radius:8px;background:linear-gradient(145deg,#173344,#10232c);color:#b9e8ff;font:700 22px/1 "Segoe UI",sans-serif;cursor:pointer;transition:border-color .2s,box-shadow .2s,transform .2s}.report-button:hover{transform:translateY(-2px);border-color:#72c9ff;box-shadow:0 0 16px #0099ff42}.report-button:focus-visible{outline:2px solid #72c9ff;outline-offset:2px}.report-button:disabled{opacity:.55;cursor:wait;transform:none}.report-editor{display:grid;gap:9px;min-width:220px}.report-editor label{color:#91b5b7;font-size:11px;letter-spacing:.06em;text-transform:uppercase}.report-editor input{width:100%;padding:10px 11px;border:1px solid #31525b;background:#071217;color:var(--text);font:inherit}.report-editor .actions{margin-top:0}
    .account-switch{display:block;margin:16px auto 0;padding:8px;border:0;background:none;color:var(--blue);font:inherit;font-size:11px;cursor:pointer}.account-switch:hover{color:var(--green);text-decoration:underline}
    body{font-size:16px;line-height:1.55}main{padding-top:34px}.hero h1{font-size:clamp(30px,5vw,46px)}.hero p,.login-copy{font-size:15px}.panel{border-radius:12px;backdrop-filter:blur(8px);transition:border-color .2s,box-shadow .2s}.panel:hover{border-color:#28505b;box-shadow:0 14px 38px #0006,0 0 24px #0099ff12}.panel h2{font-size:17px;letter-spacing:.07em}.button{font-size:14px;border-radius:7px}.badge{font-size:12px}.field label{font-size:13px}.field input,.filter{font-size:16px}.lookup-fields{display:grid;grid-template-columns:minmax(150px,210px) minmax(220px,300px);gap:18px;align-items:end}.lookup-field input{height:58px;font-size:25px;font-weight:800;letter-spacing:.14em;text-align:center;border-color:#00ffcc66;box-shadow:inset 0 0 18px #00ffcc0b,0 0 16px #00ffcc0a}.lookup-field input:focus{border-color:var(--green);box-shadow:0 0 0 3px #00ffcc25,0 0 24px #00ffcc20}.lookup-field input::placeholder{font-size:14px;letter-spacing:0;font-weight:400}.lookup-label{color:#c2f5ed!important;font-size:13px!important}.queue-person{font-size:18px}.queue-meta{font-size:13px;line-height:1.6}.queue-timer{font-size:14px}.authorization-option{font-size:15px}.authorization-option .pharmacy-code{font-size:12px}.history-editor,.history-addition{font-size:14px}.history-addition .button,.history-editor .button{font-size:12px}th,td{font-size:14px;padding:14px}.table-wrap{border-radius:8px}.table-wrap tbody tr:hover{background:#102128}.state{font-size:13px}.empty{font-size:15px}.admin-lock-form{display:flex;gap:10px;align-items:end;flex-wrap:wrap}.admin-lock-form .field{width:min(340px,100%)}.admin-lock-form input{min-height:48px}.admin-help{color:var(--muted);font-size:14px;max-width:720px}.status.show{border-radius:8px}.lookup-help{font-size:13px;color:var(--muted);margin:12px 0 0}@media(max-width:650px){.lookup-fields{grid-template-columns:1fr 1fr;gap:10px}.lookup-field input{height:54px;font-size:22px}.lookup-field input::placeholder{font-size:12px}.queue-row{align-items:flex-start;flex-direction:column}.queue-row>.button{width:100%}th,td{font-size:13px}}
    .queue-list{display:grid;gap:10px}.queue-row{display:flex;justify-content:space-between;align-items:center;gap:16px;padding:14px;border:1px solid #23434c;background:#08151a}.queue-person{font-weight:700;line-height:1.5;overflow-wrap:anywhere}.queue-meta{margin-top:5px;color:var(--muted);font-size:10px}.queue-timer{margin-top:6px;color:var(--green);font-size:11px;font-weight:700;letter-spacing:.04em}.queue-selected{border-color:var(--green);box-shadow:inset 3px 0 var(--green),0 0 14px #00ffcc12}.selected-visit{margin:14px 0 0;padding:12px 14px;border:1px solid #00ffcc50;color:var(--green);font-size:12px;line-height:1.6}.authorization-option{display:grid;grid-template-columns:auto minmax(0,1fr) minmax(110px,180px);align-items:center;gap:12px;padding:12px;border:1px solid #23434c;background:#08151a}.authorization-option input[type=checkbox]{accent-color:var(--green);width:18px;height:18px}.authorization-option .pharmacy-name{font-weight:700}.authorization-option .pharmacy-code{display:block;margin-top:4px;color:var(--muted);font-size:10px}.authorization-option input[type=number]{width:100%;min-width:0;border:1px solid #23434c;background:#071217;color:var(--text);padding:10px;font:inherit}.history-addition{display:flex;gap:7px;align-items:center;min-width:210px;flex-wrap:wrap}.history-addition input{width:72px;border:1px solid #23434c;background:#071217;color:var(--text);padding:9px;font:inherit}.history-addition .button{padding:9px 11px}.history-editor{display:grid;gap:8px;min-width:240px}.history-editor .authorization-option{grid-template-columns:auto minmax(0,1fr) 78px;padding:8px}.history-editor .authorization-option input[type=number]{padding:7px}
    @media(max-width:650px){.shell{width:calc(100% - 24px)}main{padding-top:30px}.top-inner{min-height:60px}.top-meta{display:none}.hero{display:block}.grid{grid-template-columns:1fr}.panel{padding:18px}.table-head{align-items:flex-start;flex-direction:column}.filter{width:100%}.login-wrap{margin:24px auto}}
  </style>
</head>
<body>
  <header class="topbar"><div class="shell top-inner"><div class="brand"><span class="brand-mark">V/</span> VALE // MOSTRADOR</div><div class="top-meta"><span class="live">● SISTEMA EN LÍNEA</span> &nbsp;·&nbsp; CONTROL DE ENTREGAS</div></div></header>
  <main class="shell">
    <section id="login-screen" class="login-wrap">
      <div class="panel"><div class="eyebrow">Acceso de operador / 01</div><h1 class="login-title" id="auth-title">INICIAR SESIÓN</h1><p class="login-copy" id="auth-copy">Identifícate para abrir el terminal de control de entregas.</p>
        <form id="login-form"><div class="field"><label for="username">Usuario</label><input id="username" name="username" autocomplete="username" required></div><div class="field" style="margin-top:16px"><label for="password">Contraseña</label><input id="password" name="password" type="password" autocomplete="current-password" required></div><label class="remember"><input type="checkbox" name="remember"> Recordar esta sesión en este dispositivo</label><div class="login-error" id="login-error" role="alert"></div><button class="button primary" id="login-button" type="submit">CONECTAR TERMINAL</button></form>
        <button class="account-switch${ALLOW_OPERATOR_REGISTRATION ? '' : ' hidden'}" id="show-register" type="button">¿Primera vez? Registrar operario</button>
        <form id="register-form" class="hidden"><div class="field"><label for="new-username">Nombre de operario / usuario</label><input id="new-username" name="username" autocomplete="username" minlength="3" maxlength="32" pattern="[A-Za-z0-9._\-]+" required></div><div class="field" style="margin-top:16px"><label for="new-password">Contraseña (mínimo 8 caracteres)</label><input id="new-password" name="password" type="password" autocomplete="new-password" minlength="8" maxlength="128" required></div><div class="field" style="margin-top:16px"><label for="confirm-password">Confirmar contraseña</label><input id="confirm-password" name="confirmPassword" type="password" autocomplete="new-password" minlength="8" maxlength="128" required></div><div class="login-error" id="register-error" role="alert"></div><button class="button primary" id="register-button" type="submit">CREAR OPERARIO</button><button class="account-switch" id="show-login" type="button">Volver a iniciar sesión</button></form>
      </div><div class="footer">ACCESO RESTRINGIDO · SESIÓN AUDITADA</div>
    </section>
    <section id="dashboard" class="hidden">
      <div class="hero"><div><div class="eyebrow">Terminal de control / 02</div><h1>ENTREGA DE VALES</h1><p>Verifica la identidad autorizada y registra cada operación de mostrador.</p></div><div class="actions"><span class="badge" id="operator-badge"></span><button class="button secondary" id="refresh-page" type="button">ACTUALIZAR PÁGINA</button><button class="button secondary logout" id="logout-button" type="button">CERRAR SESIÓN</button></div></div>
      <section class="panel"><div class="table-head"><div><h2>▸ Buscar persona</h2><p class="sub" style="margin:7px 0 0;color:var(--muted);font-size:14px">Busca los ingresos de hoy por código de farmacia o DNI. Usa el DNI si no sabes si está autorizado.</p></div><div class="actions" style="margin:0"><span class="badge" id="waiting-count">0 RESULTADOS</span><button class="button secondary" id="refresh-waiting" type="button">BUSCAR</button></div></div><div class="lookup-fields"><div class="field lookup-field"><label class="lookup-label" for="lookup-code">Código farmacia · MAESTRO A</label><input class="filter" id="lookup-code" type="search" inputmode="numeric" maxlength="5" autocomplete="off" placeholder="ID (5 dígitos)" aria-label="Filtrar por código de farmacia"></div><div class="field lookup-field"><label class="lookup-label" for="lookup-dni">DNI de la persona</label><input class="filter" id="lookup-dni" type="search" inputmode="numeric" maxlength="8" autocomplete="off" placeholder="DNI (7 u 8 dígitos)" aria-label="Filtrar por DNI"></div></div><p class="lookup-help">Puedes usar uno de los filtros o combinar ambos para acotar la búsqueda.</p><button class="button secondary hidden" id="show-authorized-people" type="button">DESPLEGAR LISTA DE AUTORIZADOS</button><div class="queue-list" id="waiting-list" aria-live="polite"></div><div class="queue-list hidden" id="authorized-people-list" aria-live="polite"></div><div class="login-error" id="waiting-error" role="alert"></div></section>
      <section class="panel"><h2>▸ Registrar entrega</h2><p class="login-copy">Selecciona a la persona en la lista. Marca cada farmacia autorizada y registra la cantidad de pedidos para cada código seleccionado.</p><form id="delivery-form"><div id="selected-visit" class="selected-visit hidden"></div><div class="queue-list" id="authorization-options"></div><div class="actions"><button class="button primary" id="verify-button" type="submit">VERIFICAR Y REGISTRAR</button><button class="button secondary" type="reset">LIMPIAR</button><button class="button secondary hidden" id="clear-visit" type="button">DESVINCULAR PERSONA</button></div></form><div id="result" class="status" role="status" aria-live="polite"></div></section>
      <section class="panel"><div class="table-head"><div><h2>▸ Historial desde REGISTRO</h2><p class="sub" style="margin:7px 0 0;color:var(--muted);font-size:13px">Actualización automática cada 30 segundos.</p></div><div class="actions" style="margin:0"><input class="filter" id="filter-code" type="search" placeholder="Filtrar código de farmacia…" aria-label="Filtrar por código de farmacia"><button class="button secondary" id="refresh-history" type="button">ACTUALIZAR HISTORIAL</button><span class="badge" id="count">0 REGISTROS</span></div></div><div class="table-wrap"><table><thead><tr><th>Fecha / hora</th><th>Código</th><th>Persona</th><th>Pedidos</th><th>Estado</th><th>Operario</th><th>Tiempo de espera</th><th>Acción</th></tr></thead><tbody id="history"></tbody></table></div><p class="login-error" id="history-error" role="alert"></p></section>
      <div class="guard-notice" role="note">Si se requiere autorizar a una persona que figura como DENEGADO, comuníquese con el PAC.</div>
      <div class="footer">HISTORIAL LEÍDO DESDE GOOGLE SHEETS · PESTAÑA REGISTRO</div>
      <section class="panel admin-panel"><h2>▸ Destino de los registros · ADMIN</h2><div id="admin-locked"><p class="admin-help">Panel protegido. Ingresa la contraseña de administrador para consultar el estado de conexión y abrir los recursos de registros.</p><form id="admin-unlock-form" class="admin-lock-form"><div class="field"><label for="admin-password">Contraseña de administrador</label><input id="admin-password" type="password" autocomplete="current-password" required></div><button class="button secondary" id="admin-unlock-button" type="submit">DESBLOQUEAR</button></form><p class="login-error" id="admin-error" role="alert"></p></div><div id="admin-content" class="hidden"><p class="login-copy" id="storage-status">Comprobando conexión con la hoja...</p><div class="actions"><a class="button secondary" id="authorization-form-link" target="_blank" rel="noopener">FORMULARIO DE AUTORIZADOS</a><a class="button secondary" id="spreadsheet-link" target="_blank" rel="noopener">ABRIR HOJA DE VALES</a><button class="button secondary" id="admin-lock-button" type="button">BLOQUEAR PANEL</button></div></div></section>
    </section>
  </main>
  <div id="processing" class="processing hidden" aria-live="polite"><div class="processing-box"><div class="processing-title">INICIALIZANDO SESIÓN</div><div class="processing-message" id="processing-message">Estableciendo conexión...</div><div class="progress"><span></span></div></div></div>
  <script>
    const $ = selector => document.querySelector(selector);
    const loginScreen = $('#login-screen'), dashboard = $('#dashboard'), processing = $('#processing');
    let deliveries = [];
    let historyError = '';
    let waitingVisits = [];
    let selectedWaitingVisit = null;
    let authorizedFallbackPeople = null;
    let authorizedFallbackCode = '';
    let waitingRefreshTimer = null;
    let waitingClockTimer = null;
    let historyRefreshTimer = null;
    let historyRefreshInProgress = false;
    async function api(url, options = {}) {
      const response = await fetch(url, { credentials: 'same-origin', ...options, headers: { 'Content-Type': 'application/json', ...(options.headers || {}) } });
      const data = await response.json();
      if (!response.ok) throw new Error(data.error || 'No se pudo completar la operación.');
      return data;
    }
    function showResult(state, title, message) {
      const result = $('#result');
      result.className = 'status show ' + state;
      result.setAttribute('role', state === 'denied' ? 'alertdialog' : 'status');
      result.setAttribute('aria-modal', state === 'denied' ? 'true' : 'false');
      result.replaceChildren();
      const heading = document.createElement('strong'), text = document.createElement('span');
      heading.textContent = title; text.textContent = message; result.append(heading, text);
      if (state === 'denied') {
        const authorizationForm = document.createElement('a');
        authorizationForm.href = 'https://script.google.com/macros/s/AKfycbw-kgQZNcY6IDxGpu-HAAp61YuPpnu1pGI_iyVFLYjcq6CMhh6Bnepp8dGrSjlcoLyU/exec';
        authorizationForm.target = '_blank';
        authorizationForm.rel = 'noopener';
        authorizationForm.className = 'denied-close';
        authorizationForm.textContent = 'ABRIR FORMULARIO DE AUTORIZADOS';
        const close = document.createElement('button');
        close.type = 'button';
        close.className = 'denied-close';
        close.textContent = 'ENTENDIDO · CERRAR';
        close.addEventListener('click', () => {
          result.className = 'status';
          result.setAttribute('role', 'status');
          result.setAttribute('aria-modal', 'false');
          $('#verify-button').focus();
        });
        result.append(authorizationForm, close);
        close.focus();
      }
    }
    function renderHistory() {
      const filter = $('#filter-code').value.replace(/\D/g, '');
      const rows = deliveries.filter(item => String(item.code).includes(filter));
      $('#count').textContent = rows.length + (rows.length === 1 ? ' REGISTRO' : ' REGISTROS');
      const body = $('#history'); body.replaceChildren();
      if (historyError) {
        const row = document.createElement('tr'), cell = document.createElement('td');
        cell.colSpan = 8; cell.className = 'empty'; cell.textContent = 'No se pudo leer el historial desde REGISTRO: ' + historyError;
        row.append(cell); body.append(row); return;
      }
      if (!rows.length) {
        const row = document.createElement('tr'), cell = document.createElement('td');
        cell.colSpan = 8; cell.className = 'empty'; cell.textContent = filter ? 'No hay resultados para ese código.' : 'Aún no hay entregas registradas.';
        row.append(cell); body.append(row); return;
      }
      for (const item of rows) {
        const row = document.createElement('tr');
        const date = new Date(item.date);
        const dateText = item.dateHasTime
          ? date.toLocaleString('es-AR', {
            timeZone: 'America/Argentina/Buenos_Aires',
            hourCycle: 'h23'
          })
          : date.toLocaleDateString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires' }) + ' · HORA NO DISPONIBLE';
        const values = [dateText, item.code, item.person || '—', item.quantity, item.state, item.operator, item.waitingTime || '—'];
        values.forEach((value, index) => {
          const cell = document.createElement('td');
          cell.textContent = value;
          if (index === 4) cell.className = 'state ' + (['AUTORIZADO', 'AGREGADO'].includes(item.state) ? 'authorized' : 'denied');
          row.append(cell);
        });
        const actionCell = document.createElement('td');
        const actions = document.createElement('div');
        actions.className = 'history-actions';
        const elapsed = Date.now() - Date.parse(item.registeredAt || item.date);
        if (item.state === 'AUTORIZADO' && Number.isFinite(elapsed) && elapsed >= 0 && elapsed <= 40 * 60 * 1000) {
          const button = document.createElement('button');
          button.type = 'button'; button.className = 'button secondary'; button.dataset.editRow = String(item.rowNumber);
          button.textContent = 'EDITAR / AGREGAR PEDIDOS';
          actions.append(button);
        } else if (item.state === 'AUTORIZADO') {
          const expired = document.createElement('span');
          expired.textContent = 'PLAZO VENCIDO · 40 MIN';
          expired.title = 'Los pedidos adicionales solo se pueden agregar durante los 40 minutos posteriores al registro.';
          actions.append(expired);
        } else {
          const unavailable = document.createElement('span');
          unavailable.textContent = '—';
          actions.append(unavailable);
        }
        const reportButton = document.createElement('button');
        reportButton.type = 'button';
        reportButton.className = 'report-button';
        reportButton.dataset.reportRow = String(item.rowNumber);
        reportButton.setAttribute('aria-label', 'Enviar reporte ejecutivo de ' + (item.person || 'la persona'));
        reportButton.title = 'Enviar reporte ejecutivo a gerencia';
        reportButton.textContent = '↗';
        actions.append(reportButton);
        actionCell.append(actions);
        row.append(actionCell);
        body.append(row);
      }
    }
    async function refreshDeliveryHistory() {
      if (historyRefreshInProgress) return;
      historyRefreshInProgress = true;
      const button = $('#refresh-history');
      if (button) button.disabled = true;
      historyError = '';
      try {
        deliveries = await api('/api/deliveries');
        $('#history-error').textContent = '';
      } catch (error) {
        deliveries = [];
        historyError = error.message;
        $('#history-error').textContent = 'No se pudo cargar el historial de la pestaña REGISTRO: ' + error.message;
      } finally {
        renderHistory();
        historyRefreshInProgress = false;
        if (button) button.disabled = false;
      }
    }
    function formatWaitingTime(waitingSince) {
      const startedAt = Date.parse(waitingSince);
      if (!Number.isFinite(startedAt)) return 'SIN HORA DE INGRESO';
      const elapsedSeconds = Math.max(0, Math.floor((Date.now() - startedAt) / 1000));
      if (elapsedSeconds > 12 * 60 * 60) return 'EXEDIDO';
      const hours = String(Math.floor(elapsedSeconds / 3600)).padStart(2, '0');
      const minutes = String(Math.floor((elapsedSeconds % 3600) / 60)).padStart(2, '0');
      return hours + ':' + minutes;
    }
    function describeSelectedVisit(visit) {
      return 'PERSONA EN ESPERA SELECCIONADA: ' + (visit.displayPerson || visit.person) +
        (visit.time ? ' · INGRESO: ' + visit.time : '') +
        ' · TIEMPO DE ESPERA: ' + (visit.waitingStoppedTime || formatWaitingTime(visit.waitingSince));
    }
    function renderAuthorizationOptions(visit) {
      const container = $('#authorization-options');
      container.replaceChildren();
      const requestedCode = $('#lookup-code').value.replace(/\D/g, '');
      const matchingAuthorization = visit.authorizations.find(item => item.code === requestedCode);
      const options = requestedCode
        ? (matchingAuthorization ? visit.authorizations : [{ code: requestedCode, pharmacy: 'ID consultado', unauthorized: true }])
        : (visit.authorizations.length ? visit.authorizations : [{ code: '', pharmacy: 'Ingresar código de farmacia', unauthorized: true }]);
      for (const authorization of options) {
        const option = document.createElement('label');
        option.className = 'authorization-option';
        const manualCode = authorization.unauthorized && !authorization.code;
        const checkbox = document.createElement('input');
        checkbox.type = 'checkbox';
        checkbox.value = authorization.code;
        checkbox.checked = manualCode || (Boolean(requestedCode) && authorization.code === requestedCode);
        checkbox.dataset.authorized = authorization.unauthorized ? 'false' : 'true';
        const pharmacy = document.createElement('span');
        pharmacy.className = 'pharmacy-name';
        pharmacy.textContent = (authorization.pharmacy || 'Farmacia') +
          (authorization.unauthorized ? ' · NO AUTORIZADO' : '');
        const code = document.createElement('span');
        code.className = 'pharmacy-code';
        code.textContent = manualCode ? '' : 'CÓDIGO: ' + authorization.code;
        pharmacy.append(code);
        let codeInput = null;
        if (manualCode) {
          codeInput = document.createElement('input');
          codeInput.type = 'text'; codeInput.inputMode = 'numeric'; codeInput.maxLength = 5;
          codeInput.placeholder = 'Código (máx. 5 dígitos)'; codeInput.required = true;
          codeInput.dataset.manualCode = 'true';
          codeInput.setAttribute('aria-label', 'Código de farmacia a verificar');
          codeInput.addEventListener('input', () => { checkbox.value = codeInput.value.replace(/\D/g, ''); });
        }
        const quantity = document.createElement('input');
        quantity.type = 'number'; quantity.min = '0'; quantity.max = '10000'; quantity.step = '1';
        quantity.value = '0'; quantity.disabled = !checkbox.checked; quantity.required = checkbox.checked;
        quantity.setAttribute('aria-label', 'Cantidad para el código ' + authorization.code);
        checkbox.addEventListener('change', () => {
          quantity.disabled = !checkbox.checked;
          quantity.required = checkbox.checked;
          if (codeInput) codeInput.disabled = !checkbox.checked;
        });
        option.append(checkbox, pharmacy);
        if (codeInput) option.append(codeInput);
        option.append(quantity);
        container.append(option);
      }
      $('#selected-visit').textContent = describeSelectedVisit(visit);
      $('#selected-visit').classList.remove('hidden');
      $('#clear-visit').classList.remove('hidden');
    }
    function updateWaitingTimers() {
      document.querySelectorAll('.queue-timer[data-waiting-since]').forEach(timer => {
        timer.textContent = 'ESPERA: ' + (timer.dataset.stoppedTime || formatWaitingTime(timer.dataset.waitingSince));
      });
      if (selectedWaitingVisit && !$('#selected-visit').classList.contains('hidden')) {
        $('#selected-visit').textContent = describeSelectedVisit(selectedWaitingVisit);
      }
    }
    function renderWaitingVisits() {
      const list = $('#waiting-list');
      list.replaceChildren();
      const code = $('#lookup-code').value.replace(/\D/g, '');
      $('#show-authorized-people').classList.toggle('hidden', code.length === 0);
      $('#waiting-count').textContent = waitingVisits.length + ' RESULTADOS';
      if (!waitingVisits.length) {
        const empty = document.createElement('div');
        const hasFilter = $('#lookup-code').value.trim() || $('#lookup-dni').value.trim();
        empty.className = 'empty';
        empty.textContent = hasFilter
            ? 'No hay personas de hoy que coincidan con los filtros.'
            : 'Escribe un código de farmacia o DNI para buscar personas ingresadas hoy.';
        list.append(empty);
        return;
      }
      for (const visit of waitingVisits) {
        const row = document.createElement('div');
        row.className = 'queue-row' + (selectedWaitingVisit && selectedWaitingVisit.id === visit.id ? ' queue-selected' : '');
        const detail = document.createElement('div');
        const person = document.createElement('div'); person.className = 'queue-person'; person.textContent = visit.person;
        const meta = document.createElement('div'); meta.className = 'queue-meta'; meta.textContent = [
          visit.time,
          visit.dni ? 'DNI: ' + visit.dni : 'NO SE DETECTÓ DNI EN VISITAS',
          visit.code ? 'CÓDIGO DE INGRESO: ' + visit.code : '',
          visit.authorizations.length
            ? 'AUTORIZADO PARA: ' + visit.authorizations.map(item => item.code + (item.pharmacy ? ' · ' + item.pharmacy : '')).join(' / ')
            : 'SIN AUTORIZACIÓN EN MAESTRO'
        ].filter(Boolean).join(' · ');
        const authorization = document.createElement('div');
        authorization.className = 'queue-timer ' + (visit.authorized ? 'state authorized' : 'state denied');
        authorization.textContent = visit.authorized ? 'AUTORIZADO EN MAESTRO' : 'DENEGADO · DNI NO ASOCIADO EN MAESTRO';
        const timer = document.createElement('div'); timer.className = 'queue-timer'; timer.dataset.waitingSince = visit.waitingSince || ''; timer.dataset.visitId = visit.id;
        timer.textContent = 'ESPERA: ' + formatWaitingTime(visit.waitingSince);
        detail.append(person, meta, authorization, timer);
        const action = document.createElement('button');
        action.className = 'button secondary'; action.type = 'button'; action.dataset.visitId = visit.id;
        action.textContent = selectedWaitingVisit && selectedWaitingVisit.id === visit.id ? 'SELECCIONADA' : 'ATENDER';
        row.append(detail, action); list.append(row);
      }
    }
    function renderAuthorizedFallbackPeople() {
      const list = $('#authorized-people-list');
      list.replaceChildren();
      list.classList.remove('hidden');
      const selectedDni = selectedWaitingVisit && selectedWaitingVisit.manualEntry ? selectedWaitingVisit.dni : '';
      const notice = document.createElement('div');
      notice.className = 'empty';
      notice.textContent = 'Autorizados de MAESTRO para el código ' + authorizedFallbackCode +
        '. Al elegir a una persona, el ingreso se registrará desde ahora y la espera comenzará en 00:00.';
      list.append(notice);
      for (const [index, person] of authorizedFallbackPeople.entries()) {
        const row = document.createElement('div');
        row.className = 'queue-row' + (selectedDni === person.dni ? ' queue-selected' : '');
        const detail = document.createElement('div');
        const name = document.createElement('div');
        name.className = 'queue-person';
        name.textContent = person.name;
        const meta = document.createElement('div');
        meta.className = 'queue-meta';
        meta.textContent = 'DNI: ' + person.dni + ' · CÓDIGO: ' + person.code +
          (person.pharmacy ? ' · ' + person.pharmacy : '');
        const note = document.createElement('div');
        note.className = 'queue-timer';
        note.textContent = 'INGRESO: SE REGISTRARÁ AHORA · ESPERA: 00:00';
        detail.append(name, meta, note);
        const action = document.createElement('button');
        action.type = 'button';
        action.className = 'button secondary';
        action.dataset.authorizedPerson = String(index);
        action.textContent = selectedDni === person.dni ? 'SELECCIONADA' : 'ELEGIR PERSONA';
        row.append(detail, action);
        list.append(row);
      }
      if (!authorizedFallbackPeople.length) {
        const empty = document.createElement('div');
        empty.className = 'empty';
        empty.textContent = 'No hay personas autorizadas en MAESTRO para este código.';
        list.append(empty);
      }
    }
    async function refreshWaitingVisits() {
      const button = $('#refresh-waiting');
      if (button) button.disabled = true;
      try {
        const query = new URLSearchParams({
          code: $('#lookup-code').value,
          dni: $('#lookup-dni').value
        });
        waitingVisits = await api('/api/waiting-visits?' + query.toString());
        if (selectedWaitingVisit && !selectedWaitingVisit.manualEntry &&
          !waitingVisits.some(visit => visit.id === selectedWaitingVisit.id)) {
          selectedWaitingVisit = null;
          $('#selected-visit').classList.add('hidden');
          $('#clear-visit').classList.add('hidden');
          $('#authorization-options').replaceChildren();
        }
        $('#waiting-error').textContent = '';
        if (authorizedFallbackPeople && $('#lookup-code').value.replace(/\D/g, '') === authorizedFallbackCode) {
          renderAuthorizedFallbackPeople();
        } else {
          authorizedFallbackPeople = null;
          authorizedFallbackCode = '';
          $('#authorized-people-list').replaceChildren();
          $('#authorized-people-list').classList.add('hidden');
          $('#show-authorized-people').textContent = 'DESPLEGAR LISTA DE AUTORIZADOS';
          renderWaitingVisits();
        }
      } catch (error) {
        $('#waiting-error').textContent = 'No se pudo cargar la lista de ingresos: ' + error.message;
      } finally {
        if (button) button.disabled = false;
      }
    }
    async function openDashboard(operator) {
      $('#operator-badge').textContent = 'OPERARIO: ' + operator;
      loginScreen.classList.add('hidden'); dashboard.classList.remove('hidden');
      const adminSession = await api('/api/admin/session');
      setAdminPanel(adminSession.authenticated);
      await Promise.all([
        refreshDeliveryHistory(),
        refreshWaitingVisits(),
        adminSession.authenticated ? loadAdminStatus() : Promise.resolve()
      ]);
      if (waitingRefreshTimer) clearInterval(waitingRefreshTimer);
      waitingRefreshTimer = setInterval(() => {
        if ($('#lookup-code').value.trim() || $('#lookup-dni').value.trim()) refreshWaitingVisits();
      }, 30000);
      if (historyRefreshTimer) clearInterval(historyRefreshTimer);
      historyRefreshTimer = setInterval(() => {
        if (!document.hidden) refreshDeliveryHistory();
      }, 30000);
      if (waitingClockTimer) clearInterval(waitingClockTimer);
      waitingClockTimer = setInterval(updateWaitingTimers, 1000);
    }
    function setAdminPanel(unlocked) {
      $('#admin-locked').classList.toggle('hidden', unlocked);
      $('#admin-content').classList.toggle('hidden', !unlocked);
      if (!unlocked) {
        $('#storage-status').textContent = '';
        $('#authorization-form-link').removeAttribute('href');
        $('#spreadsheet-link').removeAttribute('href');
      }
    }
    async function loadAdminStatus() {
      try {
        const status = await api('/api/status');
        $('#authorization-form-link').href = status.authorizationFormUrl;
        $('#spreadsheet-link').href = status.spreadsheetUrl;
        $('#storage-status').textContent = status.spreadsheetConnected
          ? 'Apps Script conectado (' + status.spreadsheetVersion + '). Las entregas se guardan y el historial se consulta en la pestaña REGISTRO.'
          : 'Historial: Google Sheets, pestaña REGISTRO. Guardado de entregas no conectado: ' + (status.spreadsheetError || 'revisa el despliegue de Apps Script.') + ' ' + status.localFile + ' no se usa como historial.';
      } catch (error) {
        $('#storage-status').textContent = 'No se pudo verificar el destino de registros: ' + error.message;
      }
    }
    $('#login-form').addEventListener('submit', async event => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      const button = $('#login-button');
      const stages = ['Estableciendo conexión...', 'Validando credenciales...', 'Sincronizando sesión...'];
      processing.classList.remove('hidden'); button.disabled = true; $('#login-error').textContent = '';
      let stage = 0;
      $('#processing-message').textContent = stages[stage];
      const timer = setInterval(() => { stage = Math.min(stage + 1, stages.length - 1); $('#processing-message').textContent = stages[stage]; }, 450);
      try {
        const result = await api('/api/login', { method: 'POST', body: JSON.stringify({ username: form.get('username'), password: form.get('password'), remember: form.has('remember') }) });
        await new Promise(resolve => setTimeout(resolve, 1400));
        clearInterval(timer); processing.classList.add('hidden');
        await openDashboard(result.operator);
      } catch (error) {
        clearInterval(timer); processing.classList.add('hidden'); $('#login-error').textContent = error.message;
      } finally { button.disabled = false; }
    });
    $('#show-register').addEventListener('click', () => {
      $('#login-form').classList.add('hidden'); $('#show-register').classList.add('hidden');
      $('#register-form').classList.remove('hidden');
      $('#auth-title').textContent = 'REGISTRAR OPERARIO';
      $('#auth-copy').textContent = 'Crea una cuenta para acceder al terminal de control de entregas.';
      $('#register-error').textContent = '';
    });
    $('#show-login').addEventListener('click', () => {
      $('#register-form').classList.add('hidden'); $('#login-form').classList.remove('hidden');
      $('#show-register').classList.remove('hidden');
      $('#auth-title').textContent = 'INICIAR SESIÓN';
      $('#auth-copy').textContent = 'Identifícate para abrir el terminal de control de entregas.';
      $('#register-error').textContent = '';
    });
    $('#register-form').addEventListener('submit', async event => {
      event.preventDefault();
      const form = new FormData(event.currentTarget);
      const password = form.get('password');
      const button = $('#register-button');
      $('#register-error').textContent = '';
      if (password !== form.get('confirmPassword')) {
        $('#register-error').textContent = 'Las contraseñas no coinciden.';
        return;
      }
      button.disabled = true; button.textContent = 'CREANDO CUENTA…';
      try {
        const result = await api('/api/register', { method: 'POST', body: JSON.stringify({ username: form.get('username'), password }) });
        await openDashboard(result.operator);
      } catch (error) { $('#register-error').textContent = error.message; }
      finally { button.disabled = false; button.textContent = 'CREAR OPERARIO'; }
    });
    $('#waiting-list').addEventListener('click', event => {
      const manualButton = event.target.closest('button[data-authorized-person]');
      if (manualButton) {
        const person = authorizedFallbackPeople && authorizedFallbackPeople[Number(manualButton.dataset.authorizedPerson)];
        if (!person) return;
        const waitingSince = new Date().toISOString();
        selectedWaitingVisit = {
          id: 'manual-' + person.dni + '-' + Date.now(),
          person: person.person,
          displayPerson: person.name,
          dni: person.dni,
          code: person.code,
          authorizations: person.authorizations,
          waitingSince,
          time: 'INGRESO REGISTRADO AHORA',
          manualEntry: true,
          manualCode: person.code
        };
        renderAuthorizationOptions(selectedWaitingVisit);
        renderAuthorizedFallbackPeople();
        $('#authorization-options input[type="checkbox"]').focus();
        return;
      }
      const button = event.target.closest('button[data-visit-id]');
      if (!button) return;
      selectedWaitingVisit = waitingVisits.find(visit => visit.id === button.dataset.visitId) || null;
      if (!selectedWaitingVisit) return;
      renderAuthorizationOptions(selectedWaitingVisit);
      renderWaitingVisits();
      $('#authorization-options input[type="checkbox"]').focus();
    });
    $('#admin-unlock-form').addEventListener('submit', async event => {
      event.preventDefault();
      const button = $('#admin-unlock-button');
      button.disabled = true;
      $('#admin-error').textContent = '';
      try {
        await api('/api/admin/unlock', { method: 'POST', body: JSON.stringify({ password: $('#admin-password').value }) });
        $('#admin-password').value = '';
        setAdminPanel(true);
        await loadAdminStatus();
      } catch (error) {
        $('#admin-error').textContent = error.message;
      } finally {
        button.disabled = false;
      }
    });
    $('#admin-lock-button').addEventListener('click', async () => {
      try {
        await api('/api/admin/lock', { method: 'POST' });
        setAdminPanel(false);
      } catch (error) {
        $('#admin-error').textContent = error.message;
      }
    });
    let waitingSearchTimer = null;
    function scheduleWaitingSearch() {
      clearTimeout(waitingSearchTimer);
      authorizedFallbackPeople = null;
      authorizedFallbackCode = '';
      waitingSearchTimer = setTimeout(refreshWaitingVisits, 300);
    }
    $('#lookup-code').addEventListener('input', scheduleWaitingSearch);
    $('#lookup-dni').addEventListener('input', scheduleWaitingSearch);
    $('#refresh-waiting').addEventListener('click', refreshWaitingVisits);
    $('#show-authorized-people').addEventListener('click', async event => {
      const button = event.currentTarget;
      const code = $('#lookup-code').value.replace(/\D/g, '');
      if (!code) return;
      if (authorizedFallbackPeople !== null) {
        authorizedFallbackPeople = null;
        authorizedFallbackCode = '';
        $('#authorized-people-list').replaceChildren();
        $('#authorized-people-list').classList.add('hidden');
        button.textContent = 'DESPLEGAR LISTA DE AUTORIZADOS';
        return;
      }
      button.disabled = true;
      $('#waiting-error').textContent = '';
      try {
        const result = await api('/api/authorized-people?code=' + encodeURIComponent(code));
        authorizedFallbackPeople = result.people;
        authorizedFallbackCode = code;
        button.textContent = 'OCULTAR LISTA DE AUTORIZADOS';
        renderAuthorizedFallbackPeople();
      } catch (error) {
        $('#waiting-error').textContent = 'No se pudo cargar la lista de autorizados: ' + error.message;
      } finally {
        button.disabled = false;
      }
    });
    $('#refresh-page').addEventListener('click', () => window.location.reload());
    $('#clear-visit').addEventListener('click', () => {
      selectedWaitingVisit = null;
      $('#selected-visit').classList.add('hidden'); $('#clear-visit').classList.add('hidden');
      $('#authorization-options').replaceChildren();
      if (authorizedFallbackPeople) renderAuthorizedFallbackPeople();
      else renderWaitingVisits();
    });
    $('#delivery-form').addEventListener('submit', async event => {
      event.preventDefault();
      const deliveryForm = event.currentTarget;
      const button = $('#verify-button');
      if (!selectedWaitingVisit) {
        showResult('info', 'SELECCIONA UNA PERSONA', 'Elige ATENDER en una persona de la lista de espera de hoy.');
        return;
      }
      const entries = [...$('#authorization-options').querySelectorAll('.authorization-option')]
        .filter(option => option.querySelector('input[type="checkbox"]').checked)
        .map(option => ({
          code: option.querySelector('input[type="checkbox"]').value,
          quantity: Number(option.querySelector('input[type="number"]').value),
          authorized: option.querySelector('input[type="checkbox"]').dataset.authorized === 'true'
        }));
      if (!entries.length) {
        showResult('info', 'SELECCIONA UNA FARMACIA', 'Marca al menos un código de farmacia y escribe la cantidad correspondiente.');
        return;
      }
      if (entries.some(entry => !/^\d+$/.test(entry.code) ||
        !Number.isSafeInteger(entry.quantity) || entry.quantity < 0 || entry.quantity > 10000)) {
        showResult('info', 'CANTIDAD NO VÁLIDA', 'Cada código seleccionado debe tener una cantidad entera entre 0 y 10000.');
        return;
      }
      if (entries.some(entry => !entry.authorized) && entries.length > 1) {
        showResult('info', 'INTENTO DENEGADO POR SEPARADO', 'El código no autorizado debe registrarse por separado de los códigos autorizados.');
        return;
      }
      button.disabled = true; button.textContent = 'CONSULTANDO AUTORIZACIÓN…'; $('#result').className = 'status';
      try {
        const visitId = selectedWaitingVisit.id;
        const deniedEntry = entries.find(entry => !entry.authorized);
        const data = await api('/api/deliveries', { method: 'POST', body: JSON.stringify({
          entries: entries.map(({ code, quantity }) => ({ code, quantity })),
          requestedCode: deniedEntry ? deniedEntry.code : $('#lookup-code').value,
          visitId,
          manualEntry: selectedWaitingVisit.manualEntry === true,
          manualCode: selectedWaitingVisit.manualCode || '',
          personDni: selectedWaitingVisit.manualEntry ? selectedWaitingVisit.dni : ''
        }) });
        const outcomes = data.deliveries || [];
        const denied = outcomes.some(delivery => delivery.state !== 'AUTORIZADO');
        const firstDelivery = outcomes[0];
        if (firstDelivery && firstDelivery.waitingTime) {
          selectedWaitingVisit.waitingStoppedTime = firstDelivery.waitingTime;
          const queueTimer = [...document.querySelectorAll('.queue-timer')].find(timer => timer.dataset.visitId === visitId);
          if (queueTimer) {
            queueTimer.dataset.stoppedTime = firstDelivery.waitingTime;
            queueTimer.textContent = 'ESPERA: ' + firstDelivery.waitingTime;
          }
          $('#selected-visit').textContent = describeSelectedVisit(selectedWaitingVisit);
        }
        await refreshDeliveryHistory();
        deliveryForm.reset();
        selectedWaitingVisit = null;
        authorizedFallbackPeople = null;
        authorizedFallbackCode = '';
          $('#authorized-people-list').replaceChildren();
          $('#authorized-people-list').classList.add('hidden');
          $('#show-authorized-people').textContent = 'DESPLEGAR LISTA DE AUTORIZADOS';
        $('#selected-visit').classList.add('hidden'); $('#clear-visit').classList.add('hidden');
        $('#authorization-options').replaceChildren();
        await refreshWaitingVisits();
        const summary = outcomes.map(delivery =>
          delivery.code + ': ' + delivery.state + ' (' + delivery.quantity + ' pedidos)'
        ).join(' · ');
        showResult(denied ? 'denied' : 'ok', denied ? 'DENEGADO' : 'AUTORIZADO', summary || 'Registro guardado correctamente.');
      } catch (error) { showResult('info', 'NO SE PUDO VERIFICAR', error.message); }
      finally { button.disabled = false; button.textContent = 'VERIFICAR Y REGISTRAR'; }
    });
    $('#history').addEventListener('click', async event => {
      const button = event.target.closest('button[data-report-row], button[data-send-report], button[data-cancel-report], button[data-edit-row], button[data-save-additions], button[data-cancel-additions]');
      if (!button) return;
      const row = button.closest('tr');
      const editorForAction = button.closest('.history-editor');
      const reportEditor = button.closest('.report-editor');
      const rowNumber = Number(button.dataset.reportRow || button.dataset.editRow ||
        (editorForAction && editorForAction.dataset.rowNumber) ||
        (reportEditor && reportEditor.dataset.rowNumber));
      if (button.dataset.reportRow !== undefined) {
        const editor = document.createElement('div');
        editor.className = 'report-editor';
        editor.dataset.rowNumber = String(rowNumber);
        const label = document.createElement('label');
        label.textContent = 'Correo de gerencia';
        const emailInput = document.createElement('input');
        emailInput.type = 'email';
        emailInput.required = true;
        emailInput.maxLength = 254;
        emailInput.autocomplete = 'email';
        emailInput.placeholder = 'gerencia@empresa.com';
        emailInput.setAttribute('aria-label', 'Correo electrónico de gerencia');
        const controls = document.createElement('div');
        controls.className = 'actions';
        const sendButton = document.createElement('button');
        sendButton.type = 'button';
        sendButton.className = 'button primary';
        sendButton.dataset.sendReport = 'true';
        sendButton.textContent = 'ENVIAR REPORTE';
        const cancelButton = document.createElement('button');
        cancelButton.type = 'button';
        cancelButton.className = 'button secondary';
        cancelButton.dataset.cancelReport = 'true';
        cancelButton.textContent = 'CANCELAR';
        controls.append(sendButton, cancelButton);
        editor.append(label, emailInput, controls);
        row.lastElementChild.replaceChildren(editor);
        emailInput.focus();
        return;
      }
      if (button.dataset.cancelReport !== undefined) {
        refreshDeliveryHistory();
        return;
      }
      if (button.dataset.sendReport !== undefined) {
        const email = reportEditor.querySelector('input[type="email"]').value.trim();
        if (email.length > 254 || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
          showResult('info', 'CORREO NO VÁLIDO', 'Ingresa una dirección de correo electrónico válida para gerencia.');
          return;
        }
        button.disabled = true;
        try {
          const result = await api('/api/deliveries/' + encodeURIComponent(rowNumber) + '/report', {
            method: 'POST', body: JSON.stringify({ email: email.trim() })
          });
          await refreshDeliveryHistory();
          showResult('ok', 'REPORTE ENVIADO', 'El reporte ejecutivo se envió a ' + result.email + '.');
        } catch (error) {
          showResult('info', 'NO SE PUDO ENVIAR EL REPORTE', error.message);
          button.disabled = false;
        }
        return;
      }
      if (button.dataset.cancelAdditions !== undefined) {
        refreshDeliveryHistory();
        return;
      }
      if (button.dataset.editRow !== undefined) {
        button.disabled = true;
        try {
          const data = await api('/api/deliveries/' + encodeURIComponent(rowNumber) + '/additions');
          if (!data.authorizations.length) {
            throw new Error('Esta persona no tiene códigos autorizados disponibles en MAESTRO.');
          }
          const actionCell = row.lastElementChild;
          const editor = document.createElement('div');
          editor.className = 'history-editor';
          editor.dataset.rowNumber = String(rowNumber);
          for (const authorization of data.authorizations) {
            const option = document.createElement('label');
            option.className = 'authorization-option';
            const checkbox = document.createElement('input');
            checkbox.type = 'checkbox'; checkbox.value = authorization.codigoFarmacia;
            const pharmacy = document.createElement('span');
            pharmacy.className = 'pharmacy-name';
            pharmacy.textContent = authorization.farmacia || 'Farmacia';
            const code = document.createElement('span');
            code.className = 'pharmacy-code';
            code.textContent = 'CÓDIGO: ' + authorization.codigoFarmacia;
            pharmacy.append(code);
            const quantity = document.createElement('input');
            quantity.type = 'number'; quantity.min = '0'; quantity.max = '10000'; quantity.step = '1';
            quantity.value = '0'; quantity.disabled = true;
            quantity.setAttribute('aria-label', 'Cantidad para código ' + authorization.codigoFarmacia);
            checkbox.addEventListener('change', () => { quantity.disabled = !checkbox.checked; });
            option.append(checkbox, pharmacy, quantity);
            editor.append(option);
          }
          const controls = document.createElement('div');
          controls.className = 'actions';
          const save = document.createElement('button');
          save.type = 'button'; save.className = 'button primary'; save.dataset.saveAdditions = 'true';
          save.textContent = 'GUARDAR AGREGADOS';
          const cancel = document.createElement('button');
          cancel.type = 'button'; cancel.className = 'button secondary'; cancel.dataset.cancelAdditions = 'true';
          cancel.textContent = 'CANCELAR';
          controls.append(save, cancel);
          editor.append(controls);
          actionCell.replaceChildren(editor);
        } catch (error) {
          showResult('info', 'NO SE PUDO ABRIR LA EDICIÓN', error.message);
          button.disabled = false;
        }
        return;
      }

      const editor = button.closest('.history-editor');
      const entries = [...editor.querySelectorAll('.authorization-option')]
        .filter(option => option.querySelector('input[type="checkbox"]').checked)
        .map(option => ({
          code: option.querySelector('input[type="checkbox"]').value,
          quantity: Number(option.querySelector('input[type="number"]').value)
        }));
      if (!entries.length || entries.some(entry => !Number.isSafeInteger(entry.quantity) || entry.quantity < 0 || entry.quantity > 10000)) {
        showResult('info', 'CANTIDAD NO VÁLIDA', 'Marca al menos una farmacia y escribe una cantidad entera entre 0 y 10000 para cada código.');
        return;
      }
      button.disabled = true;
      try {
        const data = await api('/api/deliveries/' + encodeURIComponent(rowNumber) + '/addition', {
          method: 'POST', body: JSON.stringify({ entries })
        });
        await refreshDeliveryHistory();
        showResult('ok', 'AGREGADOS', data.deliveries.map(delivery =>
          delivery.code + ': ' + delivery.quantity + ' pedido(s)'
        ).join(' · '));
      } catch (error) {
        showResult('info', 'NO SE PUDIERON GUARDAR LOS AGREGADOS', error.message);
        button.disabled = false;
      }
    });
    $('#filter-code').addEventListener('input', renderHistory);
    $('#refresh-history').addEventListener('click', refreshDeliveryHistory);
    $('#delivery-form').addEventListener('reset', () => {
      $('#result').className = 'status';
      $('#result').setAttribute('role', 'status');
      $('#result').setAttribute('aria-modal', 'false');
      selectedWaitingVisit = null;
      $('#selected-visit').classList.add('hidden');
      $('#clear-visit').classList.add('hidden');
      $('#authorization-options').replaceChildren();
      renderWaitingVisits();
    });
    $('#logout-button').addEventListener('click', async () => {
      try { await api('/api/logout', { method: 'POST', body: '{}' }); }
      catch (error) { showResult('info', 'NO SE PUDO CERRAR SESIÓN', error.message); return; }
      if (waitingRefreshTimer) clearInterval(waitingRefreshTimer);
      if (waitingClockTimer) clearInterval(waitingClockTimer);
      if (historyRefreshTimer) clearInterval(historyRefreshTimer);
      waitingRefreshTimer = null; waitingClockTimer = null; historyRefreshTimer = null;
      selectedWaitingVisit = null; waitingVisits = [];
      dashboard.classList.add('hidden'); loginScreen.classList.remove('hidden'); $('#login-form').reset(); deliveries = [];
    });
    api('/api/session').then(data => { if (data.operator) openDashboard(data.operator).catch(error => showResult('info', 'ERROR DE SESIÓN', error.message)); }).catch(() => {});
  </script>
</body>
</html>`;

loadStore().then(loaded => {
  store = loaded;
  app.listen(PORT, () => console.log(`Vale // Mostrador disponible en http://localhost:${PORT}`));
}).catch(error => {
  console.error('No se pudieron cargar los registros de vales:', error);
  process.exitCode = 1;
});
