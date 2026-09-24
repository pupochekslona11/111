const SHEET_NAME = 'Операции';
const CAPITAL_SHEET_NAME = 'Капитал';
const CAPITAL_HISTORY_SHEET_NAME = 'История капитала';

function setup() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty('SPREADSHEET_ID');
  let ss;

  if (spreadsheetId) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } else {
    ss = SpreadsheetApp.create('Мои финансы — учёт доходов и расходов');
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }

  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    const sheets = ss.getSheets();
    if (sheets.length === 1 && sheets[0].getLastRow() === 0) {
      sh = sheets[0];
      sh.setName(SHEET_NAME);
    } else {
      sh = ss.insertSheet(SHEET_NAME);
    }
  }

  if (sh.getLastRow() === 0) {
    sh.appendRow(['ID','Дата','Тип','Сумма','Категория','Исходная команда','Создано']);
    sh.setFrozenRows(1);
    sh.getRange('A1:G1').setFontWeight('bold');
    sh.getRange('B:B').setNumberFormat('dd.mm.yyyy');
    sh.getRange('D:D').setNumberFormat('#,##0.00');
    sh.autoResizeColumns(1, 7);
  }

  Logger.log('Таблица: ' + ss.getUrl());
  return ss.getUrl();
}

function generateSecuritySecret() {
  const props = PropertiesService.getScriptProperties();
  const secret = [
    Utilities.getUuid().replace(/-/g, ''),
    Utilities.getUuid().replace(/-/g, '')
  ].join('');

  props.setProperty('SECURITY_SECRET', secret);
  Logger.log('SECURITY_SECRET=' + secret);
  return secret;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = String(p.callback || '').trim();

  if (!p.action) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, service: 'voice-finance-api-v2' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const payload = handleApi_(p);

  if (callback) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
      return ContentService
        .createTextOutput('/* invalid callback */')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleApi_(p) {
  try {
    verifySignedRequest_(p);

    const action = String(p.action || '');
    const month = String(p.month || '');

    if (action === 'dashboard') {
      return { ok: true, data: getDashboard(month) };
    }

    if (action === 'add') {
      addOperation({
        id: String(p.id || Utilities.getUuid()),
        date: String(p.date || ''),
        type: String(p.type || ''),
        amount: Number(p.amount),
        category: String(p.category || 'Без категории'),
        raw: String(p.raw || '')
      });
      return { ok: true, data: getDashboard(month) };
    }

    if (action === 'delete') {
      deleteOperation(String(p.id || ''));
      return { ok: true, data: getDashboard(month) };
    }

    if (action === 'capital_dashboard') {
      return { ok: true, data: getCapitalDashboard() };
    }

    if (action === 'capital_add' || action === 'capital_update') {
      upsertCapitalSource_({
        id: String(p.id || Utilities.getUuid()),
        sourceType: String(p.sourceType || 'Другое'),
        name: String(p.name || 'Без названия'),
        amount: Number(p.amount),
        rate: Number(p.rate || 0),
        note: String(p.note || '')
      });
      saveCapitalSnapshot_();
      return { ok: true, data: getCapitalDashboard() };
    }

    if (action === 'capital_delete') {
      deleteCapitalSource_(String(p.id || ''));
      saveCapitalSnapshot_();
      return { ok: true, data: getCapitalDashboard() };
    }

    if (action === 'capital_settings') {
      setMonthlyContribution_(Number(p.monthlyContribution || 0));
      return { ok: true, data: getCapitalDashboard() };
    }

    throw new Error('Неизвестная команда API');
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
}

function verifySignedRequest_(p) {
  const secret = PropertiesService.getScriptProperties().getProperty('SECURITY_SECRET');
  if (!secret) throw new Error('Защита не настроена: запусти generateSecuritySecret().');

  const ts = Number(p.ts || 0);
  const nonce = String(p.nonce || '');
  const sig = String(p.sig || '').toLowerCase();

  if (!Number.isFinite(ts) || Math.abs(Date.now() - ts) > 5 * 60 * 1000) {
    throw new Error('Запрос устарел');
  }

  if (!/^[a-zA-Z0-9]{16,128}$/.test(nonce)) {
    throw new Error('Некорректный nonce');
  }

  if (!/^[0-9a-f]{64}$/.test(sig)) {
    throw new Error('Некорректная подпись');
  }

  const expected = bytesToHex_(
    Utilities.computeHmacSha256Signature(
      canonicalForSign_(p),
      secret,
      Utilities.Charset.UTF_8
    )
  );

  if (!constantTimeEqual_(sig, expected)) {
    throw new Error('Неверная подпись');
  }

  const cache = CacheService.getScriptCache();
  const nonceKey = 'nonce:' + nonce;
  if (cache.get(nonceKey)) {
    throw new Error('Повторный запрос');
  }
  cache.put(nonceKey, '1', 600);
}

function canonicalForSign_(p) {
  const order = ['action','id','date','type','amount','category','raw','month','sourceType','name','rate','note','monthlyContribution','ts','nonce'];
  return order
    .filter(k => p[k] !== undefined && p[k] !== null)
    .map(k => k + '=' + encodeURIComponent(String(p[k])))
    .join('&');
}

function bytesToHex_(bytes) {
  return bytes
    .map(b => {
      const v = b < 0 ? b + 256 : b;
      return ('0' + v.toString(16)).slice(-2);
    })
    .join('');
}

function constantTimeEqual_(a, b) {
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return diff === 0;
}

function getSheet() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Сначала запусти setup().');
  const ss = SpreadsheetApp.openById(id);
  let sh = ss.getSheetByName(SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(['ID','Дата','Тип','Сумма','Категория','Исходная команда','Создано']);
  }

  return sh;
}

function addOperation(op) {
  if (!op) throw new Error('Нет данных операции.');

  const amount = Number(op.amount);
  if (!Number.isFinite(amount) || amount <= 0) throw new Error('Некорректная сумма.');
  if (op.type !== 'income' && op.type !== 'expense') throw new Error('Некорректный тип операции.');

  const sh = getSheet();
  const id = op.id || Utilities.getUuid();
  const now = new Date();
  const date = op.date ? new Date(op.date + 'T12:00:00') : now;

  if (sh.getLastRow() > 1) {
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().flat();
    if (ids.includes(String(id))) return { ok: true, duplicate: true, id: id };
  }

  sh.appendRow([
    id,
    date,
    op.type === 'income' ? 'Доход' : 'Расход',
    amount,
    String(op.category || 'Без категории').slice(0, 100),
    String(op.raw || '').slice(0, 300),
    now
  ]);

  return { ok: true, id: id };
}

function deleteOperation(id) {
  const sh = getSheet();
  if (sh.getLastRow() < 2) return { ok: false };

  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().flat();
  const index = ids.indexOf(String(id));
  if (index === -1) return { ok: false };

  sh.deleteRow(index + 2);
  return { ok: true };
}

function getDashboard(month) {
  const sh = getSheet();

  if (sh.getLastRow() < 2) {
    return { income: 0, expense: 0, balance: 0, operations: [] };
  }

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  let income = 0;
  let expense = 0;
  const operations = [];

  rows.forEach(row => {
    const date = new Date(row[1]);
    if (isNaN(date)) return;

    const ym = Utilities.formatDate(date, 'Europe/Moscow', 'yyyy-MM');
    if (month && ym !== month) return;

    const type = row[2] === 'Доход' ? 'income' : 'expense';
    const amount = Number(row[3]) || 0;

    if (type === 'income') income += amount;
    else expense += amount;

    operations.push({
      id: String(row[0]),
      date: Utilities.formatDate(date, 'Europe/Moscow', 'yyyy-MM-dd'),
      type: type,
      amount: amount,
      category: String(row[4] || 'Без категории'),
      raw: String(row[5] || '')
    });
  });

  operations.reverse();

  return {
    income: income,
    expense: expense,
    balance: income - expense,
    operations: operations
  };
}


function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Сначала запусти setup().');
  return SpreadsheetApp.openById(id);
}

function getCapitalSheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(CAPITAL_SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(CAPITAL_SHEET_NAME);
    sh.appendRow(['ID','Тип','Название','Сумма','Доходность % годовых','Комментарий','Создано','Обновлено']);
    sh.setFrozenRows(1);
    sh.getRange('A1:H1').setFontWeight('bold');
    sh.getRange('D:D').setNumberFormat('#,##0.00');
    sh.getRange('E:E').setNumberFormat('0.00');
    sh.autoResizeColumns(1, 8);
  }

  return sh;
}

function getCapitalHistorySheet_() {
  const ss = getSpreadsheet_();
  let sh = ss.getSheetByName(CAPITAL_HISTORY_SHEET_NAME);

  if (!sh) {
    sh = ss.insertSheet(CAPITAL_HISTORY_SHEET_NAME);
    sh.appendRow(['Дата','Капитал']);
    sh.setFrozenRows(1);
    sh.getRange('A1:B1').setFontWeight('bold');
    sh.getRange('A:A').setNumberFormat('dd.mm.yyyy');
    sh.getRange('B:B').setNumberFormat('#,##0.00');
    sh.autoResizeColumns(1, 2);
  }

  return sh;
}

function getCapitalSources_() {
  const sh = getCapitalSheet_();
  if (sh.getLastRow() < 2) return [];

  return sh.getRange(2, 1, sh.getLastRow() - 1, 8).getValues()
    .filter(row => String(row[0] || '').trim())
    .map(row => ({
      id: String(row[0]),
      sourceType: String(row[1] || 'Другое'),
      name: String(row[2] || 'Без названия'),
      amount: Number(row[3]) || 0,
      rate: Number(row[4]) || 0,
      note: String(row[5] || ''),
      created: row[6] ? Utilities.formatDate(new Date(row[6]), 'Europe/Moscow', 'yyyy-MM-dd') : '',
      updated: row[7] ? Utilities.formatDate(new Date(row[7]), 'Europe/Moscow', 'yyyy-MM-dd') : ''
    }));
}

function upsertCapitalSource_(source) {
  const amount = Number(source.amount);
  const rate = Number(source.rate || 0);

  if (!Number.isFinite(amount) || amount < 0) throw new Error('Некорректная сумма капитала.');
  if (!Number.isFinite(rate) || rate < -100 || rate > 10000) throw new Error('Некорректная доходность.');

  const sh = getCapitalSheet_();
  const now = new Date();
  const id = String(source.id || Utilities.getUuid());

  let rowIndex = -1;
  if (sh.getLastRow() > 1) {
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().flat();
    const idx = ids.indexOf(id);
    if (idx >= 0) rowIndex = idx + 2;
  }

  if (rowIndex > 0) {
    const created = sh.getRange(rowIndex, 7).getValue() || now;
    sh.getRange(rowIndex, 1, 1, 8).setValues([[
      id,
      String(source.sourceType || 'Другое').slice(0, 80),
      String(source.name || 'Без названия').slice(0, 120),
      amount,
      rate,
      String(source.note || '').slice(0, 300),
      created,
      now
    ]]);
  } else {
    sh.appendRow([
      id,
      String(source.sourceType || 'Другое').slice(0, 80),
      String(source.name || 'Без названия').slice(0, 120),
      amount,
      rate,
      String(source.note || '').slice(0, 300),
      now,
      now
    ]);
  }

  return { ok: true, id: id };
}

function deleteCapitalSource_(id) {
  const sh = getCapitalSheet_();
  if (sh.getLastRow() < 2) return { ok: false };

  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues().flat();
  const index = ids.indexOf(String(id));
  if (index < 0) return { ok: false };

  sh.deleteRow(index + 2);
  return { ok: true };
}

function capitalTotal_() {
  return getCapitalSources_().reduce((sum, source) => sum + (Number(source.amount) || 0), 0);
}

function saveCapitalSnapshot_() {
  const sh = getCapitalHistorySheet_();
  const today = Utilities.formatDate(new Date(), 'Europe/Moscow', 'yyyy-MM-dd');
  const total = capitalTotal_();

  if (sh.getLastRow() > 1) {
    const dates = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = dates.length - 1; i >= 0; i--) {
      const d = dates[i][0];
      if (!d) continue;
      const key = Utilities.formatDate(new Date(d), 'Europe/Moscow', 'yyyy-MM-dd');
      if (key === today) {
        sh.getRange(i + 2, 2).setValue(total);
        return;
      }
    }
  }

  sh.appendRow([new Date(today + 'T12:00:00'), total]);
}

function getCapitalHistory_() {
  const sh = getCapitalHistorySheet_();
  if (sh.getLastRow() < 2) return [];

  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .filter(row => row[0])
    .map(row => ({
      date: Utilities.formatDate(new Date(row[0]), 'Europe/Moscow', 'yyyy-MM-dd'),
      total: Number(row[1]) || 0
    }))
    .sort((a, b) => a.date.localeCompare(b.date));
}

function setMonthlyContribution_(value) {
  const v = Number(value);
  if (!Number.isFinite(v) || v < 0) throw new Error('Некорректное ежемесячное пополнение.');
  PropertiesService.getScriptProperties().setProperty('MONTHLY_CONTRIBUTION', String(v));
}

function getCapitalDashboard() {
  const sources = getCapitalSources_();
  const history = getCapitalHistory_();
  const monthlyContribution = Number(
    PropertiesService.getScriptProperties().getProperty('MONTHLY_CONTRIBUTION') || 0
  ) || 0;

  const total = sources.reduce((sum, source) => sum + source.amount, 0);
  const expectedAnnualIncome = sources.reduce(
    (sum, source) => sum + source.amount * source.rate / 100,
    0
  );

  return {
    sources: sources,
    history: history,
    total: total,
    expectedAnnualIncome: expectedAnnualIncome,
    expectedMonthlyIncome: expectedAnnualIncome / 12,
    monthlyContribution: monthlyContribution
  };
}
