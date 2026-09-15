const SPREADSHEET_ID = 'PASTE_GOOGLE_SHEET_ID_HERE';
const API_TOKEN = 'PASTE_SHARED_API_TOKEN_HERE';
const SCHEMA_VERSION = 'online_1_2';

const ACTIVE_STATUSES = ['Action Required', 'Monitoring'];
const TABLES = {
  issues: [
    'id', 'created_at', 'issue_time', 'resolved_time', 'line', 'instrument', 'worker',
    'category', 'subcategory', 'title', 'description', 'status', 'resolution_notes',
    'updated_at', 'deleted_at', 'client_request_id'
  ],
  version_templates: [
    'id', 'group_name', 'sw_version', 'algo_version', 'description', 'sw_description',
    'algo_description', 'worker', 'created_at', 'updated_at', 'deleted_at', 'client_request_id'
  ],
  version_history: [
    'id', 'created_at', 'update_time', 'group_name', 'line', 'instrument', 'sw_version',
    'algo_version', 'description', 'sw_description', 'algo_description', 'sw_touched',
    'algo_touched', 'worker', 'created_issue_id', 'updated_at', 'deleted_at', 'client_request_id'
  ],
  dl_trained_models: [
    'id', 'created_at', 'trained_time', 'model_family', 'model_version', 'change_type',
    'target_machines', 'scope', 'description', 'worker', 'status', 'created_issue_id',
    'updated_at', 'deleted_at', 'client_request_id'
  ],
  dl_model_applications: [
    'id', 'created_at', 'applied_time', 'model_family', 'model_version', 'change_type',
    'line', 'polarity', 'instrument', 'machine', 'scope', 'description', 'worker',
    'created_issue_id', 'updated_at', 'deleted_at', 'client_request_id'
  ]
};

function doGet(e) {
  return jsonOutput({ ok: true, service: 'Vision Issue Tracker Online API', version: SCHEMA_VERSION });
}

function doPost(e) {
  try {
    const request = parseRequest_(e);
    authorize_(request);
    ensureDatabaseReady_();

    switch (request.action) {
      case 'ping':
        return jsonOutput({ ok: true, message: 'pong', server_time: serverNow_() });
      case 'init':
        return jsonOutput({ ok: true, tables: Object.keys(TABLES), schema_version: SCHEMA_VERSION, server_time: serverNow_() });
      case 'bootstrap':
        return jsonOutput(bootstrap_());
      case 'changesSince':
        return jsonOutput(changesSince_(request.since || ''));
      case 'activeIssues':
        return jsonOutput({ ok: true, rows: activeIssues_(), server_time: serverNow_() });
      case 'issueBounds':
        return jsonOutput({ ok: true, bounds: issueTimeBounds_(), server_time: serverNow_() });
      case 'searchIssues':
        return jsonOutput({ ok: true, rows: searchIssues_(request.filters || {}), server_time: serverNow_() });
      case 'list':
        return jsonOutput({ ok: true, rows: listRows_(request.table), server_time: serverNow_() });
      case 'get':
        return jsonOutput({ ok: true, row: getRow_(request.table, request.id), server_time: serverNow_() });
      case 'append':
        return jsonOutput({ ok: true, row: appendRow_(request.table, request.row || {}), server_time: serverNow_() });
      case 'bulkAppend':
        return jsonOutput({ ok: true, rows: bulkAppendRows_(request.table, request.rows || []), server_time: serverNow_() });
      case 'update':
        return jsonOutput({ ok: true, row: updateRow_(request.table, request.id, request.row || {}), server_time: serverNow_() });
      case 'delete':
        return jsonOutput({ ok: true, row: deleteRow_(request.table, request.id), server_time: serverNow_() });
      default:
        throw new Error('Unknown action: ' + request.action);
    }
  } catch (error) {
    return jsonOutput({ ok: false, error: String(error && error.message ? error.message : error) });
  }
}

function parseRequest_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error('Missing JSON request body.');
  }
  return JSON.parse(e.postData.contents);
}

function authorize_(request) {
  if (!API_TOKEN || API_TOKEN === 'PASTE_SHARED_API_TOKEN_HERE') {
    throw new Error('API token is not configured in Apps Script.');
  }
  if (!request || request.token !== API_TOKEN) {
    throw new Error('Unauthorized.');
  }
}

function jsonOutput(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}

function serverNow_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function spreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function ensureDatabaseReady_() {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('schema_version') === SCHEMA_VERSION) {
    return;
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (properties.getProperty('schema_version') !== SCHEMA_VERSION) {
      initializeDatabase_();
      properties.setProperty('schema_version', SCHEMA_VERSION);
    }
  } finally {
    lock.releaseLock();
  }
}

function initializeDatabase_() {
  const ss = spreadsheet_();
  Object.keys(TABLES).forEach(function(table) {
    const headers = TABLES[table];
    let sheet = ss.getSheetByName(table);
    if (!sheet) {
      sheet = ss.insertSheet(table);
    }
    const currentWidth = Math.max(sheet.getLastColumn(), headers.length);
    const current = sheet.getRange(1, 1, 1, currentWidth).getValues()[0];
    const needsHeader = current.join('') === '' || headers.some(function(header, index) {
      return current[index] !== header;
    });
    sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 1), headers.length).setNumberFormat('@');
    if (needsHeader) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
    backfillMetadata_(sheet, headers);
  });
}

function backfillMetadata_(sheet, headers) {
  const lastRow = sheet.getLastRow();
  const updatedIndex = headers.indexOf('updated_at');
  if (lastRow < 2 || updatedIndex < 0) {
    return;
  }
  const createdIndex = headers.indexOf('created_at');
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  let changed = false;
  values.forEach(function(row) {
    if (row.some(function(value) { return value !== ''; }) && !row[updatedIndex]) {
      row[updatedIndex] = createdIndex >= 0 && row[createdIndex] ? row[createdIndex] : serverNow_();
      changed = true;
    }
  });
  if (changed) {
    const range = sheet.getRange(2, 1, values.length, headers.length);
    range.setNumberFormat('@');
    range.setValues(values.map(function(row) {
      return row.map(function(value) { return String(value === undefined || value === null ? '' : value); });
    }));
  }
}

function requireTable_(table) {
  if (!TABLES[table]) {
    throw new Error('Invalid table: ' + table);
  }
  return TABLES[table];
}

function sheetFor_(table) {
  requireTable_(table);
  const sheet = spreadsheet_().getSheetByName(table);
  if (!sheet) {
    throw new Error('Missing sheet: ' + table);
  }
  return sheet;
}

function bootstrap_() {
  return {
    ok: true,
    schema_version: SCHEMA_VERSION,
    server_time: serverNow_(),
    tables: {
      issues: listRows_('issues'),
      version_templates: listRows_('version_templates'),
      version_history: listRows_('version_history'),
      dl_trained_models: listRows_('dl_trained_models'),
      dl_model_applications: listRows_('dl_model_applications')
    },
    issue_bounds: issueTimeBounds_()
  };
}

function changesSince_(since) {
  return {
    ok: true,
    schema_version: SCHEMA_VERSION,
    server_time: serverNow_(),
    tables: {
      issues: changedRowsSince_('issues', since),
      version_templates: changedRowsSince_('version_templates', since),
      version_history: changedRowsSince_('version_history', since),
      dl_trained_models: changedRowsSince_('dl_trained_models', since),
      dl_model_applications: changedRowsSince_('dl_model_applications', since)
    },
    issue_bounds: issueTimeBounds_()
  };
}

function listRows_(table, includeDeleted) {
  const headers = requireTable_(table);
  const sheet = sheetFor_(table);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values
    .filter(function(row) { return row.some(function(value) { return value !== ''; }); })
    .map(function(row) { return objectFromRow_(headers, row); })
    .filter(function(row) { return includeDeleted || !row.deleted_at; });
}

function changedRowsSince_(table, since) {
  return listRows_(table, true).filter(function(row) {
    return Boolean(since) && ((row.updated_at && row.updated_at > since) || (row.deleted_at && row.deleted_at > since));
  });
}

function activeIssues_() {
  return listRows_('issues').filter(function(row) {
    return ACTIVE_STATUSES.indexOf(row.status) >= 0;
  });
}

function issueTimeBounds_() {
  const times = listRows_('issues').map(function(row) { return row.issue_time || ''; }).filter(String);
  if (!times.length) {
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    return { first: today + ' 00:00', latest: today + ' 23:59' };
  }
  times.sort();
  return { first: times[0], latest: times[times.length - 1] };
}

function searchIssues_(filters) {
  return listRows_('issues').filter(function(row) {
    return issueMatchesFilters_(row, filters || {});
  });
}

function issueMatchesFilters_(row, filters) {
  const exactFields = ['status', 'line', 'category', 'subcategory', 'worker'];
  for (let i = 0; i < exactFields.length; i++) {
    const field = exactFields[i];
    const value = String(filters[field] || '').trim();
    if (value && String(row[field] || '') !== value) {
      return false;
    }
  }
  const selectedInstruments = splitInstruments_(String(filters.instrument || '').trim());
  if (selectedInstruments.length) {
    const rowInstruments = splitInstruments_(String(row.instrument || ''));
    const matches = selectedInstruments.some(function(instrument) {
      return rowInstruments.indexOf(instrument) >= 0;
    });
    if (!matches) {
      return false;
    }
  }
  const dateFrom = String(filters.date_from || '').trim();
  const dateTo = String(filters.date_to || '').trim();
  const issueTime = String(row.issue_time || '');
  if (dateFrom && issueTime < dateFrom) {
    return false;
  }
  if (dateTo && issueTime > dateTo) {
    return false;
  }
  const keyword = String(filters.keyword || '').trim().toLowerCase();
  if (keyword) {
    const text = [row.title, row.description, row.resolution_notes].join(' ').toLowerCase();
    if (text.indexOf(keyword) < 0) {
      return false;
    }
  }
  return true;
}

function splitInstruments_(value) {
  if (!value) {
    return [];
  }
  return value.split('/').map(function(part) { return part.trim(); }).filter(String);
}

function getRow_(table, id, includeDeleted) {
  const rows = listRows_(table, Boolean(includeDeleted));
  const target = String(id);
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i].id) === target) {
      return rows[i];
    }
  }
  return null;
}

function appendRow_(table, row) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const headers = requireTable_(table);
    const sheet = sheetFor_(table);
    const requestId = String(row.client_request_id || '').trim();
    if (requestId) {
      const duplicate = findRowByValue_(sheet, headers, 'client_request_id', requestId);
      if (duplicate) {
        return duplicate;
      }
    }
    const now = serverNow_();
    const nextId = row.id ? Number(row.id) : nextId_(sheet, headers);
    const normalized = normalizeRow_(headers, Object.assign({}, row, {
      id: nextId,
      updated_at: row.updated_at || now,
      deleted_at: row.deleted_at || ''
    }));
    writeRows_(sheet, headers, [normalized]);
    return normalized;
  } finally {
    lock.releaseLock();
  }
}

function bulkAppendRows_(table, rows) {
  if (!Array.isArray(rows)) {
    throw new Error('rows must be an array.');
  }
  const created = [];
  const newRows = [];
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const headers = requireTable_(table);
    const sheet = sheetFor_(table);
    const now = serverNow_();
    let nextId = nextId_(sheet, headers);
    rows.forEach(function(row) {
      const requestId = String(row.client_request_id || '').trim();
      const duplicate = requestId ? findRowByValue_(sheet, headers, 'client_request_id', requestId) : null;
      if (duplicate) {
        created.push(duplicate);
        return;
      }
      const id = row.id ? Number(row.id) : nextId++;
      const normalized = normalizeRow_(headers, Object.assign({}, row, {
        id: id,
        updated_at: row.updated_at || now,
        deleted_at: row.deleted_at || ''
      }));
      created.push(normalized);
      newRows.push(normalized);
    });
    if (newRows.length) {
      writeRows_(sheet, headers, newRows);
    }
    return created;
  } finally {
    lock.releaseLock();
  }
}

function updateRow_(table, id, updates) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const headers = requireTable_(table);
    const sheet = sheetFor_(table);
    const rowIndex = findRowIndex_(sheet, headers, id);
    if (rowIndex < 0) {
      throw new Error('Row not found: ' + id);
    }
    const currentValues = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const current = objectFromRow_(headers, currentValues);
    const normalized = normalizeRow_(headers, Object.assign({}, current, updates, {
      id: current.id,
      updated_at: updates.updated_at || serverNow_()
    }));
    const range = sheet.getRange(rowIndex, 1, 1, headers.length);
    range.setNumberFormat('@');
    range.setValues([headers.map(function(header) { return String(normalized[header] === undefined || normalized[header] === null ? '' : normalized[header]); })]);
    return normalized;
  } finally {
    lock.releaseLock();
  }
}

function deleteRow_(table, id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const headers = requireTable_(table);
    const sheet = sheetFor_(table);
    const rowIndex = findRowIndex_(sheet, headers, id);
    if (rowIndex < 0) {
      return null;
    }
    const currentValues = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const current = objectFromRow_(headers, currentValues);
    const now = serverNow_();
    const normalized = normalizeRow_(headers, Object.assign({}, current, {
      id: current.id,
      updated_at: now,
      deleted_at: now
    }));
    const range = sheet.getRange(rowIndex, 1, 1, headers.length);
    range.setNumberFormat('@');
    range.setValues([headers.map(function(header) { return String(normalized[header] === undefined || normalized[header] === null ? '' : normalized[header]); })]);
    return normalized;
  } finally {
    lock.releaseLock();
  }
}

function findRowIndex_(sheet, headers, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return -1;
  }
  const idIndex = headers.indexOf('id');
  const deletedIndex = headers.indexOf('deleted_at');
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  const target = String(id);
  let deletedMatch = -1;
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][idIndex]) !== target) {
      continue;
    }
    if (deletedIndex >= 0 && values[i][deletedIndex]) {
      if (deletedMatch < 0) {
        deletedMatch = i + 2;
      }
      continue;
    }
    return i + 2;
  }
  return deletedMatch;
}

function findRowByValue_(sheet, headers, columnName, value) {
  const column = headers.indexOf(columnName) + 1;
  if (column <= 0 || !value) {
    return null;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return null;
  }
  const values = sheet.getRange(2, column, lastRow - 1, 1).getValues();
  const target = String(value);
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === target) {
      const row = sheet.getRange(i + 2, 1, 1, headers.length).getValues()[0];
      return objectFromRow_(headers, row);
    }
  }
  return null;
}

function nextId_(sheet, headers) {
  const idColumn = headers.indexOf('id') + 1;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 1;
  }
  const values = sheet.getRange(2, idColumn, lastRow - 1, 1).getValues();
  let maxId = 0;
  values.forEach(function(row) {
    const value = Number(row[0]);
    if (!isNaN(value)) {
      maxId = Math.max(maxId, value);
    }
  });
  return maxId + 1;
}

function writeRows_(sheet, headers, rows) {
  const startRow = sheet.getLastRow() + 1;
  const values = rows.map(function(row) {
    return headers.map(function(header) { return String(row[header] === undefined || row[header] === null ? '' : row[header]); });
  });
  const range = sheet.getRange(startRow, 1, values.length, headers.length);
  range.setNumberFormat('@');
  range.setValues(values);
}

function objectFromRow_(headers, row) {
  const object = {};
  headers.forEach(function(header, index) {
    const value = row[index];
    object[header] = value instanceof Date ? formatDate_(value) : value;
  });
  return object;
}

function normalizeRow_(headers, row) {
  const normalized = {};
  headers.forEach(function(header) {
    normalized[header] = row[header] === undefined || row[header] === null ? '' : row[header];
  });
  return normalized;
}

function formatDate_(value) {
  return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}
