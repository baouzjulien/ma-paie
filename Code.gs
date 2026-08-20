const SHEETS = {
  paie: 'PAIE_MENSUELLE',
  audit: 'AUDIT'
};

const HEADERS = [
  'MONTH_KEY',
  'UPDATED_AT',
  'DEVICE_ID',
  'YEAR',
  'MONTH_INDEX',
  'TAUX_H',
  'H_BASE',
  'H_SUP_25_CONTRAT',
  'H_SUP_25_SAISIES',
  'H_SUP_50',
  'HEURES_NON_MAJOREES',
  'HEURES_NUIT',
  'PRIME_POLYVALENCE',
  'PRIME_ENTRETIEN',
  'PRIME_TELEPHONE',
  'REPAS',
  'FRAIS_DEPLACEMENT',
  'ABSENCES',
  'NET_ESTIME',
  'BRUT_ESTIME',
  'JSON'
];

function doGet(e) {
  const action = String((e.parameter && e.parameter.action) || '').trim();
  try {
    ensureSheets();
    if (action === 'getYear') return json(getYear(e), e);
    if (action === 'saveMonth') {
      const payload = JSON.parse(String(e.parameter.payload || '{}'));
      return json(saveMonth(payload, e), e);
    }
    return json({ success:false, error:'Action inconnue' }, e);
  } catch (err) {
    audit('ERREUR_GET', String(err), e);
    return json({ success:false, error:String(err) }, e);
  }
}

function doPost(e) {
  try {
    ensureSheets();
    const payload = JSON.parse((e.postData && e.postData.contents) || '{}');
    if (payload.action === 'saveMonth') return json(saveMonth(payload, e));
    return json({ success:false, error:'Action inconnue' });
  } catch (err) {
    audit('ERREUR_POST', String(err), e);
    return json({ success:false, error:String(err) });
  }
}

function json(obj, event) {
  const callback = event && event.parameter && String(event.parameter.callback || '').trim();
  const body = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet(ss, SHEETS.paie, HEADERS);
  ensureSheet(ss, SHEETS.audit, ['TIMESTAMP', 'ACTION', 'DETAILS', 'USER_AGENT']);
}

function ensureSheet(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  const current = sh.getRange(1, 1, 1, headers.length).getValues()[0];
  const missingHeaders = headers.some((h, i) => String(current[i] || '') !== h);
  if (missingHeaders) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  return sh;
}

function getYear(e) {
  const year = Number(e.parameter.year);
  if (!year) throw new Error('Annee manquante');

  const values = sheet(SHEETS.paie).getDataRange().getValues();
  if (values.length < 2) return { success:true, months:[] };

  const headers = values[0].map(String);
  const jsonIndex = headers.indexOf('JSON');
  const yearIndex = headers.indexOf('YEAR');
  const months = values.slice(1)
    .filter(row => Number(row[yearIndex]) === year)
    .map(row => parseMonthJson(row[jsonIndex]))
    .filter(Boolean);

  audit('GET_YEAR', 'Annee ' + year + ', mois ' + months.length, e);
  return { success:true, months };
}

function saveMonth(payload, e) {
  const month = payload.month || {};
  const year = Number(month._year);
  const monthIndex = Number(month._month);
  const monthKey = String(month._monthKey || (year + '-' + String(monthIndex + 1).padStart(2, '0')));
  if (!year || isNaN(monthIndex) || !monthKey) throw new Error('Mois invalide');

  const incoming = {
    ...month,
    _year: year,
    _month: monthIndex,
    _monthKey: monthKey,
    _deviceId: String(payload.deviceId || month._deviceId || ''),
    _updatedAt: month._updatedAt || new Date().toISOString()
  };

  const sh = sheet(SHEETS.paie);
  const rowNumber = findMonthRow(sh, monthKey);
  if (rowNumber) {
    const existing = parseMonthJson(sh.getRange(rowNumber, HEADERS.indexOf('JSON') + 1).getValue());
    const existingTime = Date.parse(existing && existing._updatedAt || '') || 0;
    const incomingTime = Date.parse(incoming._updatedAt || '') || 0;
    if (existingTime > incomingTime) {
      audit('CONFLIT_IGNORE', monthKey + ' conserve: Sheet plus recent', e);
      return { success:false, conflict:true, remote:existing };
    }
    sh.getRange(rowNumber, 1, 1, HEADERS.length).setValues([toRow(incoming)]);
  } else {
    sh.appendRow(toRow(incoming));
  }

  audit('SAVE_MONTH', monthKey + ' depuis ' + incoming._deviceId, e);
  return { success:true, remote:incoming };
}

function toRow(month) {
  return [
    month._monthKey || '',
    month._updatedAt || '',
    month._deviceId || '',
    month._year || '',
    month._month || '',
    numberOrBlank(month.tauxH),
    numberOrBlank(month.hBase),
    numberOrBlank(month.hSup25),
    numberOrBlank(month.hSup25Var),
    numberOrBlank(month.hSup50),
    numberOrBlank(month.compHeures),
    numberOrBlank(month.hNuit),
    numberOrBlank(month.primePoliv),
    numberOrBlank(month.primeVeh),
    numberOrBlank(month.primeTel),
    numberOrBlank(month.nbJours),
    numberOrBlank(month.fraisDep),
    numberOrBlank(month.absences),
    numberOrBlank(month._net),
    numberOrBlank(month._brut),
    JSON.stringify(month)
  ];
}

function findMonthRow(sh, monthKey) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  const values = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === monthKey) return i + 2;
  }
  return 0;
}

function parseMonthJson(value) {
  try { return JSON.parse(String(value || '')); }
  catch(e) { return null; }
}

function sheet(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Onglet manquant: ' + name);
  return sh;
}

function numberOrBlank(value) {
  if (value === '' || value === null || value === undefined) return '';
  const n = Number(String(value).replace(',', '.'));
  return isNaN(n) ? value : n;
}

function audit(action, details, event) {
  try {
    const ua = event && event.parameter ? String(event.parameter.userAgent || '') : '';
    sheet(SHEETS.audit).appendRow([new Date(), action || '', details || '', ua]);
  } catch (err) {
    // L'audit ne doit jamais bloquer la synchro.
  }
}
