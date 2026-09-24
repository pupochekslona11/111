const SHEET_NAME = 'Операции';
const API_KEY = 'PASTE_YOUR_PRIVATE_KEY_HERE';

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

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = String(p.callback || '').trim();

  if (!p.action) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, service: 'voice-finance-api' }))
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
    if (String(p.key || '') !== API_KEY) {
      throw new Error('Неверный ключ доступа');
    }

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

    throw new Error('Неизвестная команда API');
  } catch (err) {
    return { ok: false, error: err && err.message ? err.message : String(err) };
  }
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
