const SHEET_NAME = 'Операции';
const CAPITAL_SHEET_NAME = 'Капитал';
const CAPITAL_HISTORY_SHEET_NAME = 'История капитала';
const CAPITAL_FLOW_SHEET_NAME = 'Движение капитала';
const LIABILITY_SHEET_NAME = 'Обязательства';
const GOALS_SHEET_NAME = 'Цели';
const RECURRING_SHEET_NAME = 'Регулярные операции';
const TZ = 'Europe/Moscow';

function setup() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty('SPREADSHEET_ID');
  let ss;

  if (spreadsheetId) ss = SpreadsheetApp.openById(spreadsheetId);
  else {
    ss = SpreadsheetApp.create('Мои финансы — учёт доходов и расходов');
    props.setProperty('SPREADSHEET_ID', ss.getId());
  }

  ensureOperationsSheet_();
  getCapitalSheet_();
  getCapitalHistorySheet_();
  getCapitalFlowSheet_();
  getLiabilitySheet_();
  getGoalsSheet_();
  getRecurringSheet_();
  ensureDailyTrigger_();

  Logger.log('Таблица: ' + ss.getUrl());
  return ss.getUrl();
}

function generateSecuritySecret() {
  const secret = [
    Utilities.getUuid().replace(/-/g, ''),
    Utilities.getUuid().replace(/-/g, '')
  ].join('');
  PropertiesService.getScriptProperties().setProperty('SECURITY_SECRET', secret);
  Logger.log('SECURITY_SECRET=' + secret);
  return secret;
}

function doGet(e) {
  const p = (e && e.parameter) || {};
  const callback = String(p.callback || '').trim();

  if (!p.action) {
    return ContentService.createTextOutput(JSON.stringify({
      ok:true, service:'voice-finance-api-v3'
    })).setMimeType(ContentService.MimeType.JSON);
  }

  const payload = handleApi_(p);

  if (callback) {
    if (!/^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
      return ContentService.createTextOutput('/* invalid callback */')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(payload) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleApi_(p) {
  try {
    verifySignedRequest_(p);
    const action = String(p.action || '');
    const month = String(p.month || '');

    if (action === 'dashboard') {
      applyRecurringForCurrentMonth_();
      return {ok:true, data:getDashboard(month)};
    }
    if (action === 'add') {
      addOperation({
        id:String(p.id || Utilities.getUuid()),
        date:String(p.date || ''),
        type:String(p.type || ''),
        amount:Number(p.amount),
        category:String(p.category || 'Без категории'),
        raw:String(p.raw || '')
      });
      return {ok:true, data:getDashboard(month)};
    }
    if (action === 'delete') {
      deleteOperation(String(p.id || ''));
      return {ok:true, data:getDashboard(month)};
    }

    if (action === 'capital_dashboard') return {ok:true, data:getCapitalDashboard()};
    if (action === 'capital_add' || action === 'capital_update') {
      upsertCapitalSource_({
        id:String(p.id || Utilities.getUuid()),
        sourceType:String(p.sourceType || 'Другое'),
        name:String(p.name || 'Без названия'),
        amount:Number(p.amount),
        rate:Number(p.rate || 0),
        note:String(p.note || ''),
        maturityDate:String(p.maturityDate || ''),
        liquidity:String(p.liquidity || '')
      });
      saveCapitalSnapshot_();
      return {ok:true, data:getCapitalDashboard()};
    }
    if (action === 'capital_delete') {
      deleteCapitalSource_(String(p.id || ''));
      saveCapitalSnapshot_();
      return {ok:true, data:getCapitalDashboard()};
    }
    if (action === 'capital_settings') {
      setMonthlyContribution_(Number(p.monthlyContribution || 0));
      return {ok:true, data:getCapitalDashboard()};
    }
    if (action === 'capital_flow_add') {
      addCapitalFlow_({
        id:String(p.id || Utilities.getUuid()),
        sourceId:String(p.sourceId || ''),
        date:String(p.date || ''),
        amount:Number(p.flowAmount),
        note:String(p.flowNote || '')
      });
      return {ok:true, data:getCapitalDashboard()};
    }

    if (action === 'liabilities_dashboard') return {ok:true, data:getLiabilities_()};
    if (action === 'liability_add' || action === 'liability_update') {
      upsertLiability_({
        id:String(p.id || Utilities.getUuid()),
        liabilityType:String(p.liabilityType || 'Другое'),
        name:String(p.name || 'Без названия'),
        balance:Number(p.balance),
        rate:Number(p.rate || 0),
        dueDate:String(p.dueDate || ''),
        graceEnd:String(p.graceEnd || ''),
        minPayment:Number(p.minPayment || 0),
        note:String(p.note || '')
      });
      return {ok:true, data:getLiabilities_()};
    }
    if (action === 'liability_delete') {
      deleteById_(getLiabilitySheet_(), String(p.id || ''));
      return {ok:true, data:getLiabilities_()};
    }

    if (action === 'goals_dashboard') return {ok:true, data:getGoals_()};
    if (action === 'goal_add' || action === 'goal_update') {
      upsertGoal_({
        id:String(p.id || Utilities.getUuid()),
        name:String(p.name || 'Цель'),
        targetAmount:Number(p.targetAmount),
        targetDate:String(p.targetDate || ''),
        note:String(p.note || '')
      });
      return {ok:true, data:getGoals_()};
    }
    if (action === 'goal_delete') {
      deleteById_(getGoalsSheet_(), String(p.id || ''));
      return {ok:true, data:getGoals_()};
    }

    if (action === 'recurring_dashboard') return {ok:true, data:getRecurring_()};
    if (action === 'recurring_add' || action === 'recurring_update') {
      upsertRecurring_({
        id:String(p.id || Utilities.getUuid()),
        name:String(p.name || 'Регулярная операция'),
        type:String(p.type || 'expense'),
        amount:Number(p.amount),
        category:String(p.category || 'Без категории'),
        dayOfMonth:Number(p.dayOfMonth || 1),
        enabled:String(p.enabled || 'true') === 'true'
      });
      return {ok:true, data:getRecurring_()};
    }
    if (action === 'recurring_delete') {
      deleteById_(getRecurringSheet_(), String(p.id || ''));
      return {ok:true, data:getRecurring_()};
    }

    if (action === 'stats_dashboard') {
      applyRecurringForCurrentMonth_();
      return {ok:true, data:getStatsDashboard_()};
    }

    throw new Error('Неизвестная команда API');
  } catch (err) {
    return {ok:false, error:err && err.message ? err.message : String(err)};
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
  if (!/^[a-zA-Z0-9]{16,128}$/.test(nonce)) throw new Error('Некорректный nonce');
  if (!/^[0-9a-f]{64}$/.test(sig)) throw new Error('Некорректная подпись');

  const expected = bytesToHex_(
    Utilities.computeHmacSha256Signature(
      canonicalForSign_(p), secret, Utilities.Charset.UTF_8
    )
  );
  if (!constantTimeEqual_(sig, expected)) throw new Error('Неверная подпись');

  const cache = CacheService.getScriptCache();
  const nonceKey = 'nonce:' + nonce;
  if (cache.get(nonceKey)) throw new Error('Повторный запрос');
  cache.put(nonceKey, '1', 600);
}

function canonicalForSign_(p) {
  const order = [
    'action','id','date','type','amount','category','raw','month',
    'sourceType','sourceId','name','rate','note','maturityDate','liquidity',
    'liabilityType','balance','dueDate','graceEnd','minPayment',
    'targetAmount','targetDate','dayOfMonth','enabled',
    'monthlyContribution','flowAmount','flowNote','ts','nonce'
  ];
  return order.filter(k => p[k] !== undefined && p[k] !== null)
    .map(k => k + '=' + encodeURIComponent(String(p[k]))).join('&');
}

function bytesToHex_(bytes) {
  return bytes.map(b => {
    const v = b < 0 ? b + 256 : b;
    return ('0' + v.toString(16)).slice(-2);
  }).join('');
}

function constantTimeEqual_(a,b) {
  if (a.length !== b.length) return false;
  let diff=0;
  for (let i=0;i<a.length;i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function ss_() {
  const id=PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Сначала запусти setup().');
  return SpreadsheetApp.openById(id);
}

function ensureSheet_(name, headers) {
  const ss=ss_();
  let sh=ss.getSheetByName(name);
  if (!sh) {
    sh=ss.insertSheet(name);
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,headers.length).setFontWeight('bold');
    sh.autoResizeColumns(1,headers.length);
  } else {
    for (let i=0;i<headers.length;i++) {
      if (!String(sh.getRange(1,i+1).getValue() || '').trim()) {
        sh.getRange(1,i+1).setValue(headers[i]).setFontWeight('bold');
      }
    }
  }
  return sh;
}

function ensureOperationsSheet_() {
  const ss=ss_();
  let sh=ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    const sheets=ss.getSheets();
    if (sheets.length===1 && sheets[0].getLastRow()===0) {
      sh=sheets[0]; sh.setName(SHEET_NAME);
    } else sh=ss.insertSheet(SHEET_NAME);
  }
  if (sh.getLastRow()===0) {
    sh.appendRow(['ID','Дата','Тип','Сумма','Категория','Исходная команда','Создано']);
    sh.setFrozenRows(1);
    sh.getRange('A1:G1').setFontWeight('bold');
  }
  sh.getRange('B:B').setNumberFormat('dd.mm.yyyy');
  sh.getRange('D:D').setNumberFormat('#,##0.00');
  return sh;
}

function addOperation(op) {
  const amount=Number(op.amount);
  if (!Number.isFinite(amount) || amount<=0) throw new Error('Некорректная сумма.');
  if (op.type!=='income' && op.type!=='expense') throw new Error('Некорректный тип операции.');

  const sh=ensureOperationsSheet_();
  const id=op.id || Utilities.getUuid();
  if (sh.getLastRow()>1) {
    const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues().flat();
    if (ids.includes(String(id))) return {ok:true,duplicate:true,id:id};
  }
  const now=new Date();
  const date=op.date ? new Date(op.date+'T12:00:00') : now;
  sh.appendRow([
    id,date,op.type==='income'?'Доход':'Расход',amount,
    String(op.category||'Без категории').slice(0,100),
    String(op.raw||'').slice(0,300),now
  ]);
  return {ok:true,id:id};
}

function deleteOperation(id) {
  return deleteById_(ensureOperationsSheet_(),id);
}

function getOperations_() {
  const sh=ensureOperationsSheet_();
  if (sh.getLastRow()<2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,7).getValues().map(row=>{
    const date=new Date(row[1]);
    return {
      id:String(row[0]),
      date:isNaN(date)?'':Utilities.formatDate(date,TZ,'yyyy-MM-dd'),
      type:row[2]==='Доход'?'income':'expense',
      amount:Number(row[3])||0,
      category:String(row[4]||'Без категории'),
      raw:String(row[5]||'')
    };
  }).filter(o=>o.date);
}

function getDashboard(month) {
  let income=0,expense=0;
  const operations=[];
  getOperations_().forEach(o=>{
    if (month && o.date.slice(0,7)!==month) return;
    if (o.type==='income') income+=o.amount; else expense+=o.amount;
    operations.push(o);
  });
  operations.sort((a,b)=>b.date.localeCompare(a.date));
  return {income:income,expense:expense,balance:income-expense,operations:operations};
}

function getCapitalSheet_() {
  const sh=ensureSheet_(CAPITAL_SHEET_NAME,[
    'ID','Тип','Название','Сумма','Доходность % годовых','Комментарий',
    'Создано','Обновлено','Дата погашения / окончания','Ликвидность'
  ]);
  sh.getRange('D:D').setNumberFormat('#,##0.00');
  sh.getRange('E:E').setNumberFormat('0.00');
  sh.getRange('I:I').setNumberFormat('dd.mm.yyyy');
  return sh;
}

function defaultLiquidity_(type) {
  if (['Вклад','Облигации','ЦФА'].includes(type)) return 'До срока';
  if (['Недвижимость','Бизнес'].includes(type)) return 'Неликвидно';
  return 'Мгновенно';
}

function getCapitalSources_() {
  const sh=getCapitalSheet_();
  if (sh.getLastRow()<2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,10).getValues()
    .filter(r=>String(r[0]||'').trim())
    .map(r=>({
      id:String(r[0]),
      sourceType:String(r[1]||'Другое'),
      name:String(r[2]||'Без названия'),
      amount:Number(r[3])||0,
      rate:Number(r[4])||0,
      note:String(r[5]||''),
      created:r[6]?Utilities.formatDate(new Date(r[6]),TZ,'yyyy-MM-dd'):'',
      updated:r[7]?Utilities.formatDate(new Date(r[7]),TZ,'yyyy-MM-dd'):'',
      maturityDate:r[8]?Utilities.formatDate(new Date(r[8]),TZ,'yyyy-MM-dd'):'',
      liquidity:String(r[9]||defaultLiquidity_(String(r[1]||'')))
    }));
}

function upsertCapitalSource_(s) {
  const amount=Number(s.amount), rate=Number(s.rate||0);
  if (!Number.isFinite(amount)||amount<0) throw new Error('Некорректная сумма капитала.');
  if (!Number.isFinite(rate)||rate<-100||rate>10000) throw new Error('Некорректная доходность.');

  let maturity='';
  if (String(s.maturityDate||'').trim()) {
    maturity=new Date(String(s.maturityDate)+'T12:00:00');
    if (isNaN(maturity)) throw new Error('Некорректная дата погашения.');
  }

  const sh=getCapitalSheet_(), now=new Date(), id=String(s.id||Utilities.getUuid());
  const row=findRowById_(sh,id);
  const values=[[
    id,String(s.sourceType||'Другое').slice(0,80),
    String(s.name||'Без названия').slice(0,120),
    amount,rate,String(s.note||'').slice(0,300),
    row>0?(sh.getRange(row,7).getValue()||now):now,
    now,maturity,String(s.liquidity||defaultLiquidity_(s.sourceType)).slice(0,40)
  ]];
  if (row>0) sh.getRange(row,1,1,10).setValues(values);
  else sh.appendRow(values[0]);
}

function deleteCapitalSource_(id) {
  return deleteById_(getCapitalSheet_(),id);
}

function getCapitalHistorySheet_() {
  const sh=ensureSheet_(CAPITAL_HISTORY_SHEET_NAME,['Дата','Капитал']);
  sh.getRange('A:A').setNumberFormat('dd.mm.yyyy');
  sh.getRange('B:B').setNumberFormat('#,##0.00');
  return sh;
}

function saveCapitalSnapshot_() {
  const sh=getCapitalHistorySheet_();
  const today=Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd');
  const total=capitalTotal_();
  if (sh.getLastRow()>1) {
    const rows=sh.getRange(2,1,sh.getLastRow()-1,1).getValues();
    for (let i=rows.length-1;i>=0;i--) {
      if (!rows[i][0]) continue;
      if (Utilities.formatDate(new Date(rows[i][0]),TZ,'yyyy-MM-dd')===today) {
        sh.getRange(i+2,2).setValue(total); return;
      }
    }
  }
  sh.appendRow([new Date(today+'T12:00:00'),total]);
}

function getCapitalHistory_() {
  const sh=getCapitalHistorySheet_();
  if (sh.getLastRow()<2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,2).getValues()
    .filter(r=>r[0]).map(r=>({
      date:Utilities.formatDate(new Date(r[0]),TZ,'yyyy-MM-dd'),
      total:Number(r[1])||0
    })).sort((a,b)=>a.date.localeCompare(b.date));
}

function getCapitalFlowSheet_() {
  const sh=ensureSheet_(CAPITAL_FLOW_SHEET_NAME,['ID','Дата','Источник ID','Сумма','Комментарий','Создано']);
  sh.getRange('B:B').setNumberFormat('dd.mm.yyyy');
  sh.getRange('D:D').setNumberFormat('#,##0.00');
  return sh;
}

function getCapitalFlows_() {
  const sh=getCapitalFlowSheet_();
  if (sh.getLastRow()<2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,6).getValues()
    .filter(r=>String(r[0]||'').trim()).map(r=>({
      id:String(r[0]),
      date:r[1]?Utilities.formatDate(new Date(r[1]),TZ,'yyyy-MM-dd'):'',
      sourceId:String(r[2]||''),
      amount:Number(r[3])||0,
      note:String(r[4]||'')
    }));
}

function addCapitalFlow_(flow) {
  const amount=Number(flow.amount);
  if (!Number.isFinite(amount)||amount===0) throw new Error('Сумма пополнения/вывода не может быть 0.');

  const sources=getCapitalSources_();
  const source=sources.find(s=>s.id===flow.sourceId);
  if (!source) throw new Error('Источник капитала не найден.');
  if (source.amount+amount<0) throw new Error('После вывода сумма источника не может быть отрицательной.');

  source.amount+=amount;
  upsertCapitalSource_(source);

  const sh=getCapitalFlowSheet_();
  const date=flow.date?new Date(flow.date+'T12:00:00'):new Date();
  sh.appendRow([
    flow.id||Utilities.getUuid(),date,flow.sourceId,amount,
    String(flow.note||'').slice(0,300),new Date()
  ]);
  saveCapitalSnapshot_();
}

function capitalTotal_() {
  return getCapitalSources_().reduce((s,x)=>s+(Number(x.amount)||0),0);
}

function setMonthlyContribution_(value) {
  const v=Number(value);
  if (!Number.isFinite(v)||v<0) throw new Error('Некорректное ежемесячное пополнение.');
  PropertiesService.getScriptProperties().setProperty('MONTHLY_CONTRIBUTION',String(v));
}

function capitalPerformance_(history,flows,total) {
  if (!history.length) return {gain:0,netFlows:0,startCapital:total};
  const start=history[0];
  const netFlows=flows
    .filter(f=>f.date>start.date)
    .reduce((s,f)=>s+f.amount,0);
  return {
    startCapital:start.total,
    netFlows:netFlows,
    gain:total-start.total-netFlows
  };
}

function getCapitalDashboard() {
  const sources=getCapitalSources_();
  const history=getCapitalHistory_();
  const flows=getCapitalFlows_();
  const monthlyContribution=Number(PropertiesService.getScriptProperties().getProperty('MONTHLY_CONTRIBUTION')||0)||0;
  const total=sources.reduce((s,x)=>s+x.amount,0);
  const expectedAnnualIncome=sources.reduce((s,x)=>s+x.amount*x.rate/100,0);
  const liabilities=getLiabilities_();
  const liabilityTotal=liabilities.reduce((s,x)=>s+x.balance,0);
  const perf=capitalPerformance_(history,flows,total);

  return {
    sources:sources,history:history,flows:flows,total:total,
    expectedAnnualIncome:expectedAnnualIncome,
    expectedMonthlyIncome:expectedAnnualIncome/12,
    monthlyContribution:monthlyContribution,
    liabilitiesTotal:liabilityTotal,
    netWorth:total-liabilityTotal,
    performance:perf
  };
}

function getLiabilitySheet_() {
  const sh=ensureSheet_(LIABILITY_SHEET_NAME,[
    'ID','Тип','Название','Остаток долга','Ставка %','Дата платежа/погашения',
    'Конец грейса','Мин. платёж','Комментарий','Создано','Обновлено'
  ]);
  sh.getRange('D:D').setNumberFormat('#,##0.00');
  sh.getRange('E:E').setNumberFormat('0.00');
  sh.getRange('F:G').setNumberFormat('dd.mm.yyyy');
  sh.getRange('H:H').setNumberFormat('#,##0.00');
  return sh;
}

function getLiabilities_() {
  const sh=getLiabilitySheet_();
  if (sh.getLastRow()<2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,11).getValues()
    .filter(r=>String(r[0]||'').trim()).map(r=>({
      id:String(r[0]), liabilityType:String(r[1]||'Другое'),
      name:String(r[2]||'Без названия'), balance:Number(r[3])||0,
      rate:Number(r[4])||0,
      dueDate:r[5]?Utilities.formatDate(new Date(r[5]),TZ,'yyyy-MM-dd'):'',
      graceEnd:r[6]?Utilities.formatDate(new Date(r[6]),TZ,'yyyy-MM-dd'):'',
      minPayment:Number(r[7])||0, note:String(r[8]||'')
    }));
}

function upsertLiability_(x) {
  if (!Number.isFinite(x.balance)||x.balance<0) throw new Error('Некорректный остаток долга.');
  if (!Number.isFinite(x.rate)) throw new Error('Некорректная ставка.');
  const sh=getLiabilitySheet_(), now=new Date(), id=String(x.id||Utilities.getUuid());
  const row=findRowById_(sh,id);
  const due=x.dueDate?new Date(x.dueDate+'T12:00:00'):'';
  const grace=x.graceEnd?new Date(x.graceEnd+'T12:00:00'):'';
  const vals=[[
    id,String(x.liabilityType||'Другое'),String(x.name||'Без названия').slice(0,120),
    x.balance,x.rate,due,grace,Number(x.minPayment)||0,
    String(x.note||'').slice(0,300),
    row>0?(sh.getRange(row,10).getValue()||now):now,now
  ]];
  if (row>0) sh.getRange(row,1,1,11).setValues(vals); else sh.appendRow(vals[0]);
}

function getGoalsSheet_() {
  const sh=ensureSheet_(GOALS_SHEET_NAME,['ID','Название','Целевая сумма','Целевая дата','Комментарий','Создано','Обновлено']);
  sh.getRange('C:C').setNumberFormat('#,##0.00');
  sh.getRange('D:D').setNumberFormat('dd.mm.yyyy');
  return sh;
}

function getGoals_() {
  const sh=getGoalsSheet_();
  if (sh.getLastRow()<2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,7).getValues()
    .filter(r=>String(r[0]||'').trim()).map(r=>({
      id:String(r[0]),name:String(r[1]||'Цель'),targetAmount:Number(r[2])||0,
      targetDate:r[3]?Utilities.formatDate(new Date(r[3]),TZ,'yyyy-MM-dd'):'',
      note:String(r[4]||'')
    }));
}

function upsertGoal_(x) {
  if (!Number.isFinite(x.targetAmount)||x.targetAmount<=0) throw new Error('Некорректная сумма цели.');
  const sh=getGoalsSheet_(),now=new Date(),id=String(x.id||Utilities.getUuid());
  const row=findRowById_(sh,id);
  const td=x.targetDate?new Date(x.targetDate+'T12:00:00'):'';
  const vals=[[id,String(x.name||'Цель').slice(0,120),x.targetAmount,td,
    String(x.note||'').slice(0,300),row>0?(sh.getRange(row,6).getValue()||now):now,now]];
  if (row>0) sh.getRange(row,1,1,7).setValues(vals); else sh.appendRow(vals[0]);
}

function getRecurringSheet_() {
  const sh=ensureSheet_(RECURRING_SHEET_NAME,['ID','Название','Тип','Сумма','Категория','День месяца','Активно','Создано','Обновлено']);
  sh.getRange('D:D').setNumberFormat('#,##0.00');
  return sh;
}

function getRecurring_() {
  const sh=getRecurringSheet_();
  if (sh.getLastRow()<2) return [];
  return sh.getRange(2,1,sh.getLastRow()-1,9).getValues()
    .filter(r=>String(r[0]||'').trim()).map(r=>({
      id:String(r[0]),name:String(r[1]||'Регулярная операция'),
      type:String(r[2]||'expense'),amount:Number(r[3])||0,
      category:String(r[4]||'Без категории'),dayOfMonth:Number(r[5])||1,
      enabled:String(r[6]).toLowerCase()!=='false'
    }));
}

function upsertRecurring_(x) {
  if (!Number.isFinite(x.amount)||x.amount<=0) throw new Error('Некорректная сумма.');
  if (!Number.isFinite(x.dayOfMonth)||x.dayOfMonth<1||x.dayOfMonth>31) throw new Error('День месяца должен быть от 1 до 31.');
  if (x.type!=='income'&&x.type!=='expense') throw new Error('Некорректный тип.');
  const sh=getRecurringSheet_(),now=new Date(),id=String(x.id||Utilities.getUuid());
  const row=findRowById_(sh,id);
  const vals=[[id,String(x.name||'Регулярная операция').slice(0,120),x.type,x.amount,
    String(x.category||'Без категории').slice(0,100),Math.floor(x.dayOfMonth),
    !!x.enabled,row>0?(sh.getRange(row,8).getValue()||now):now,now]];
  if (row>0) sh.getRange(row,1,1,9).setValues(vals); else sh.appendRow(vals[0]);
}

function applyRecurringForCurrentMonth_() {
  const now=new Date();
  const ym=Utilities.formatDate(now,TZ,'yyyy-MM');
  const today=Number(Utilities.formatDate(now,TZ,'d'));
  getRecurring_().filter(r=>r.enabled&&r.dayOfMonth<=today).forEach(r=>{
    const lastDay=new Date(now.getFullYear(),now.getMonth()+1,0).getDate();
    const day=Math.min(r.dayOfMonth,lastDay);
    const date=ym+'-'+String(day).padStart(2,'0');
    addOperation({
      id:'rec-'+r.id+'-'+ym,
      date:date,type:r.type,amount:r.amount,category:r.category,
      raw:'Регулярная: '+r.name
    });
  });
}

function getStatsDashboard_() {
  const ops=getOperations_();
  const now=new Date();
  const monthly=[];
  for (let i=11;i>=0;i--) {
    const d=new Date(now.getFullYear(),now.getMonth()-i,1);
    const key=Utilities.formatDate(d,TZ,'yyyy-MM');
    const label=Utilities.formatDate(d,TZ,'MM.yy');
    let income=0,expense=0;
    ops.forEach(o=>{
      if (o.date.slice(0,7)!==key) return;
      if (o.type==='income') income+=o.amount; else expense+=o.amount;
    });
    monthly.push({month:key,label:label,income:income,expense:expense,saved:income-expense});
  }

  const currentMonth=Utilities.formatDate(now,TZ,'yyyy-MM');
  const categories={};
  ops.filter(o=>o.type==='expense'&&o.date.slice(0,7)===currentMonth).forEach(o=>{
    categories[o.category]=(categories[o.category]||0)+o.amount;
  });
  const topCategories=Object.keys(categories).map(k=>({name:k,value:categories[k]}))
    .sort((a,b)=>b.value-a.value).slice(0,8);

  const last3=monthly.slice(-3);
  const avgExpense3=last3.reduce((s,x)=>s+x.expense,0)/Math.max(last3.length,1);
  const capital=getCapitalDashboard();
  const instant=capital.sources.filter(s=>s.liquidity==='Мгновенно').reduce((s,x)=>s+x.amount,0);
  const locked=capital.sources.filter(s=>s.liquidity==='До срока').reduce((s,x)=>s+x.amount,0);
  const illiquid=capital.sources.filter(s=>s.liquidity==='Неликвидно').reduce((s,x)=>s+x.amount,0);
  const current=monthly[monthly.length-1]||{income:0,expense:0,saved:0};
  const savingsRate=current.income>0?current.saved/current.income*100:0;

  return {
    monthly:monthly,topCategories:topCategories,
    currentMonth:current,savingsRate:savingsRate,avgExpense3:avgExpense3,
    liquidity:{instant:instant,locked:locked,illiquid:illiquid,runwayMonths:avgExpense3>0?instant/avgExpense3:0},
    capitalHistory:capital.history,capitalPerformance:capital.performance,
    capitalTotal:capital.total,liabilitiesTotal:capital.liabilitiesTotal,netWorth:capital.netWorth,
    calendar:getMoneyCalendar_(),
    goals:getGoals_()
  };
}

function getMoneyCalendar_() {
  const today=new Date();
  const start=new Date(today.getFullYear(),today.getMonth(),today.getDate());
  const end=new Date(start.getTime()+120*86400000);
  const events=[];

  getCapitalSources_().forEach(s=>{
    if (!s.maturityDate) return;
    const d=new Date(s.maturityDate+'T12:00:00');
    if (d>=start&&d<=end) events.push({
      date:s.maturityDate,type:'Погашение',name:s.name,amount:s.amount
    });
  });

  getLiabilities_().forEach(x=>{
    if (x.dueDate) {
      const d=new Date(x.dueDate+'T12:00:00');
      if (d>=start&&d<=end) events.push({date:x.dueDate,type:'Платёж / погашение',name:x.name,amount:x.balance});
    }
    if (x.graceEnd) {
      const d=new Date(x.graceEnd+'T12:00:00');
      if (d>=start&&d<=end) events.push({date:x.graceEnd,type:'Конец грейса',name:x.name,amount:x.balance});
    }
  });

  return events.sort((a,b)=>a.date.localeCompare(b.date));
}

function findRowById_(sh,id) {
  if (sh.getLastRow()<2) return -1;
  const ids=sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues().flat();
  const idx=ids.indexOf(String(id));
  return idx<0?-1:idx+2;
}

function deleteById_(sh,id) {
  const row=findRowById_(sh,id);
  if (row<0) return {ok:false};
  sh.deleteRow(row);
  return {ok:true};
}


function ensureDailyTrigger_() {
  const exists = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'dailyMaintenance');
  if (!exists) {
    ScriptApp.newTrigger('dailyMaintenance')
      .timeBased()
      .everyDays(1)
      .atHour(6)
      .create();
  }
}

function dailyMaintenance() {
  applyRecurringForCurrentMonth_();
  saveCapitalSnapshot_();
}
