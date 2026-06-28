const SHEET_ID = '1oTwT7mM5jnYLJfwQp21GydqWjVR29SyrNFbOtdbTT8k';

function doGet(e) {
  const action = e.parameter.action;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let result;

  switch (action) {
    case 'loadAll':
      result = {
        certs: getSheet(ss, 'cert_master'),
        materials: getSheet(ss, 'materials'),
        materialItems: getSheet(ss, 'material_items'),
        studyLogs: getSheet(ss, 'study_log'),
        reviewLogs: getSheet(ss, 'review_log'),
        dailySummary: getSheet(ss, 'daily_summary')
      };
      break;
    default:
      result = { error: 'unknown action' };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let result = { ok: true };

  switch (action) {
    case 'appendStudyLog':
      appendRow(ss, 'study_log', body.row);
      break;
    case 'appendReviewLog':
      appendRow(ss, 'review_log', body.row);
      break;
    case 'updateMaterialProgress':
      updateCell(ss, 'materials', body.rowIndex, 7, body.value);
      break;
    case 'addMaterial':
      appendRow(ss, 'materials', body.row);
      break;
    case 'updateCert':
      updateRow(ss, 'cert_master', body.rowIndex, body.row);
      break;
    case 'appendMaterialItem':
      appendRow(ss, 'material_items', body.row);
      break;
    case 'updateReviewLog':
      updateRow(ss, 'review_log', body.rowIndex, body.row);
      break;
    case 'updateMaterialRow':
      updateRow(ss, 'materials', body.rowIndex, body.row);
      break;
    case 'aggregateDaily':
      aggregateDaily(ss);
      break;
    default:
      result = { error: 'unknown action' };
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let v = row[i];
      if (v instanceof Date) v = Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
      obj[h] = v;
    });
    return obj;
  });
}

function appendRow(ss, name, row) {
  ss.getSheetByName(name).appendRow(row);
}

function updateCell(ss, name, rowIndex, col, value) {
  ss.getSheetByName(name).getRange(rowIndex, col).setValue(value);
}

function updateRow(ss, name, rowIndex, row) {
  const sheet = ss.getSheetByName(name);
  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
}

function aggregateDaily(ss) {
  const logSheet = ss.getSheetByName('study_log');
  const sumSheet = ss.getSheetByName('daily_summary');
  const logs = logSheet.getDataRange().getValues();
  if (logs.length <= 1) return;

  const headers = logs[0];
  const tsIdx = 0, certIdx = 1, minIdx = 3;

  const daily = {};
  for (let i = 1; i < logs.length; i++) {
    const ts = logs[i][tsIdx];
    if (!ts) continue;
    const d = (ts instanceof Date) ? Utilities.formatDate(ts, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(ts).slice(0, 10);
    if (!daily[d]) daily[d] = { total: 0, certs: {}, sessions: 0 };
    daily[d].total += Number(logs[i][minIdx]) || 0;
    const cid = logs[i][certIdx];
    daily[d].certs[cid] = (daily[d].certs[cid] || 0) + (Number(logs[i][minIdx]) || 0);
    daily[d].sessions++;
  }

  const existing = sumSheet.getDataRange().getValues();
  const existDates = new Set();
  for (let i = 1; i < existing.length; i++) {
    const ed = existing[i][0];
    if (ed instanceof Date) existDates.add(Utilities.formatDate(ed, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
    else existDates.add(String(ed));
  }

  const sorted = Object.keys(daily).sort();
  let streak = 0;
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  for (let i = sorted.length - 1; i >= 0; i--) {
    const expected = new Date(today);
    expected.setDate(expected.getDate() - (sorted.length - 1 - i));
    const exp = Utilities.formatDate(expected, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (sorted[i] === exp) streak++;
    else break;
  }

  const newRows = [];
  sorted.forEach(d => {
    if (!existDates.has(d)) {
      const breakdown = Object.entries(daily[d].certs).map(([k, v]) => `${k}:${v}`).join(',');
      newRows.push([d, daily[d].total, breakdown, daily[d].sessions, streak]);
    }
  });

  if (newRows.length) {
    sumSheet.getRange(sumSheet.getLastRow() + 1, 1, newRows.length, 5).setValues(newRows);
  }
}
