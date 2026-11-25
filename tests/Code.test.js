const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'gas', 'Code.js'), 'utf8');
const headers = ['報告日時', '報告者名', '取引先ID', '取引先名', '商品ID', '商品名', '販売個数', '販売単価', '合計金額'];
const rosterHeaders = ['社員名', '部署', '役職', 'メールアドレス'];
const customerHeaders = ['取引先ID', '取引先名', '取引先区分'];
const productHeaders = ['商品ID', '商品名', '標準単価'];
const validAnswers = {報告者名: '佐藤 花子', 取引先名: 'C001｜取引先A', 商品名: 'P001｜どら焼き', 販売個数: '3'};
const timestamp = new Date('2026-10-31T00:00:00Z');

function formatDate(date, timezone, pattern) {
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en-US', {
    timeZone: timezone, year: 'numeric', month: '2-digit', day: '2-digit'
  }).formatToParts(date).map(part => [part.type, part.value]));
  if (pattern === 'yyyy') return parts.year;
  if (pattern === 'M') return String(Number(parts.month));
  if (pattern === 'yyyy-MM') return parts.year + '-' + parts.month;
  throw new Error('Unsupported pattern: ' + pattern);
}

function makeSheet(name, headings, rows) {
  return {
    name, headings, rows,
    getName() { return this.name; },
    getRange(row, column, count, width) {
      assert.deepEqual([row, column, count], [1, 1, 1]);
      return {getValues: () => [this.headings.slice(0, width)]};
    },
    getDataRange() { return {getValues: () => [this.headings, ...this.rows]}; },
    appendRow(row) { this.rows.push(row); }
  };
}

function fixture() {
  const sheets = {
    '売上報告集計シート': makeSheet('売上報告集計シート', headers, []),
    '社員名簿': makeSheet('社員名簿', rosterHeaders, [
      ['佐藤 花子', '営業部', '担当', 'sato@example.com'],
      ['鈴木 一郎', '営業部', '部長', 'director@example.com'],
      ['田中 三郎', '総務部', '担当', 'tanaka@example.com']
    ]),
    '取引先マスタ': makeSheet('取引先マスタ', customerHeaders, [['C001', '取引先A', '法人']]),
    '商品マスタ': makeSheet('商品マスタ', productHeaders, [['P001', 'どら焼き', 250]])
  };
  const spreadsheet = {getSheetByName: name => sheets[name] || null};
  const sent = [];
  const lockCalls = [];
  const context = vm.createContext({
    Date, Utilities: {formatDate},
    SpreadsheetApp: {getActiveSpreadsheet: () => spreadsheet},
    LockService: {getScriptLock: () => ({waitLock: ms => lockCalls.push(ms), releaseLock: () => lockCalls.push('release')})},
    MailApp: {sendEmail: (...args) => sent.push(args)}
  });
  vm.runInContext(source, context, {filename: 'Code.js'});
  return {context, sheets, spreadsheet, sent, lockCalls};
}

function response(answers = validAnswers) {
  return {
    getTimestamp: () => timestamp,
    getItemResponses: () => Object.entries(answers).map(([title, value]) => ({
      getItem: () => ({getTitle: () => title}), getResponse: () => value
    }))
  };
}

function makeForm() {
  const items = [];
  function item(title, type) {
    return {
      title, type, choices: [], required: false,
      getTitle() { return this.title; }, getType() { return this.type; },
      getIndex() { return items.indexOf(this); },
      setTitle(value) { this.title = value; return this; },
      setRequired(value) { this.required = value; return this; },
      setChoiceValues(values) { this.choices = [...values]; return this; },
      setValidation(value) { this.validation = value; return this; },
      asListItem() { return this; }, asTextItem() { return this; }
    };
  }
  ['報告者名', '取引先名', '商品名', '販売個数', '販売単価'].forEach(title => items.push(item(title, 'TEXT')));
  const form = {
    getId: () => 'sales-form', getItems: () => items.slice(),
    addListItem() { const added = item('', 'LIST'); items.push(added); return added; },
    moveItem(from, to) { items.splice(to, 0, items.splice(from, 1)[0]); },
    deleteItem(value) { items.splice(items.indexOf(value), 1); }
  };
  return {form, items, item};
}

function attachForm(context, form) {
  context.PropertiesService = {getScriptProperties: () => ({getProperty: key => {
    assert.equal(key, 'NAGOMI_SALES_FORM_ID'); return 'sales-form';
  }})};
  context.FormApp = {
    ItemType: {TEXT: 'TEXT', LIST: 'LIST'},
    openById(id) { assert.equal(id, 'sales-form'); return form; },
    create() { throw new Error('フォームを新規作成してはいけません'); },
    createTextValidation: () => ({
      requireWholeNumber() { return this; }, requireNumberGreaterThan(value) { assert.equal(value, 0); return this; },
      setHelpText() { return this; }, build() { return {}; }
    })
  };
}

test('メニューはマスタ選択肢の手動更新だけを表示する', () => {
  const {context} = fixture();
  const entries = [];
  context.SpreadsheetApp.getUi = () => ({createMenu: () => ({
    addItem(label, handler) { entries.push([label, handler]); return this; }, addToUi() {}
  })});
  context.onOpen();
  assert.deepEqual(entries, [['売上報告フォームの選択肢を更新', 'updateSalesFormChoicesFromMenu']]);
});

test('フォーム更新の確認を取り消すと設問を変更しない', () => {
  const {context} = fixture();
  const {form, items} = makeForm();
  attachForm(context, form);
  context.SpreadsheetApp.getUi = () => ({
    ButtonSet: {OK_CANCEL: 'set'}, Button: {OK: 'ok'},
    alert: () => 'cancel'
  });
  context.updateSalesFormChoicesFromMenu();
  assert.equal(items.length, 5);
  assert.equal(items[0].type, 'TEXT');
});

test('既存フォームの4設問を構成し、再実行で選択肢を更新する', () => {
  const {context, sheets} = fixture();
  const {form, items} = makeForm();
  attachForm(context, form);
  context.updateSalesFormChoices();
  assert.deepEqual(items.map(item => item.title), ['報告者名', '取引先名', '商品名', '販売個数']);
  assert.deepEqual(items.slice(0, 3).map(item => item.type), ['LIST', 'LIST', 'LIST']);
  assert.deepEqual(items[0].choices, ['佐藤 花子', '鈴木 一郎']);
  assert.deepEqual(items[1].choices, ['C001｜取引先A']);
  assert.deepEqual(items[2].choices, ['P001｜どら焼き']);
  assert.equal(items[3].required, true);
  assert.ok(items[3].validation);
  sheets['商品マスタ'].rows.push(['P002', '大福', 180]);
  context.updateSalesFormChoices();
  assert.equal(items.length, 4);
  assert.deepEqual(items[2].choices, ['P001｜どら焼き', 'P002｜大福']);
});

test('想定外の設問があればフォームを変更しない', () => {
  const {context} = fixture();
  const {form, items, item} = makeForm();
  items.push(item('売上日', 'TEXT'));
  attachForm(context, form);
  assert.throws(() => context.updateSalesFormChoices(), /想定外/);
  assert.equal(items[0].type, 'TEXT');
});

test('9列へマスタ値と送信時点の単価を数値で記録する', () => {
  const {context, sheets, lockCalls} = fixture();
  context.onSalesFormSubmit({response: response()});
  const row = sheets['売上報告集計シート'].rows[0];
  assert.deepEqual(Array.from(row), [timestamp, '佐藤 花子', 'C001', '取引先A', 'P001', 'どら焼き', 3, 250, 750]);
  assert.deepEqual(lockCalls, [30000, 'release']);
  sheets['商品マスタ'].rows[0][2] = 300;
  assert.equal(row[7], 250);
  assert.equal(row[8], 750);
});

test('変更された商品単価は次の報告から反映する', () => {
  const {context, sheets} = fixture();
  context.onSalesFormSubmit({response: response()});
  sheets['商品マスタ'].rows[0][2] = 300;
  context.onSalesFormSubmit({response: response()});
  assert.deepEqual(sheets['売上報告集計シート'].rows.map(row => row[8]), [750, 900]);
});

test('未登録IDと古い名称は追記せず、ロックを解放する', () => {
  const {context, sheets, lockCalls} = fixture();
  for (const change of [{取引先名: 'C999｜不明'}, {商品名: 'P001｜旧名称'}]) {
    assert.throws(() => context.onSalesFormSubmit({response: response({...validAnswers, ...change})}), /マスタ|一致/);
  }
  assert.equal(sheets['売上報告集計シート'].rows.length, 0);
  assert.deepEqual(lockCalls, [30000, 'release', 30000, 'release']);
});

test('報告者と販売個数を検証する', () => {
  const {context, spreadsheet} = fixture();
  for (const change of [{報告者名: '田中 三郎'}, {販売個数: '0'}, {販売個数: '1.5'}, {販売個数: 'abc'}]) {
    assert.throws(() => context.buildSalesRow_({...validAnswers, ...change}, timestamp, spreadsheet));
  }
  assert.throws(() => context.buildSalesRow_(validAnswers, new Date('invalid'), spreadsheet), /報告日時/);
});

test('マスタの重複IDと空欄、区切り文字を拒否する', () => {
  const {context, spreadsheet, sheets} = fixture();
  const customer = sheets['取引先マスタ'];
  customer.rows.push(['C001', '別の取引先', '法人']);
  assert.throws(() => context.readMaster_(spreadsheet, '取引先マスタ', customerHeaders), /重複/);
  customer.rows.pop();
  customer.rows[0][0] = '';
  assert.throws(() => context.readMaster_(spreadsheet, '取引先マスタ', customerHeaders), /ID・名称/);
  customer.rows[0][0] = 'C｜001';
  assert.throws(() => context.readMaster_(spreadsheet, '取引先マスタ', customerHeaders), /ID・名称/);
});

test('商品マスタの不正な単価を拒否する', () => {
  const {context, spreadsheet, sheets} = fixture();
  for (const price of ['', ' ', -1, 'abc', 12.5, false, '1e2']) {
    sheets['商品マスタ'].rows[0][2] = price;
    assert.throws(() => context.readMaster_(spreadsheet, '商品マスタ', productHeaders), /標準単価/);
  }
  sheets['商品マスタ'].rows[0][2] = 0;
  assert.equal(context.readMaster_(spreadsheet, '商品マスタ', productHeaders)[0].price, 0);
});

test('月次報告は報告日時で前月を選び9列目の金額を合計する', () => {
  const {context} = fixture();
  const row = (date, amount) => [date, '佐藤', 'C', '取引先', 'P', '商品', 1, amount, amount];
  const rows = [
    row(new Date('2025-11-30T14:59:59Z'), 999),
    row(new Date('2025-11-30T15:00:00Z'), 100),
    row(new Date('2025-12-31T14:59:59Z'), 250),
    row(new Date('2025-12-31T15:00:00Z'), 888)
  ];
  const report = context.summarizePreviousMonth_(rows, new Date('2025-12-31T15:05:00Z'));
  assert.deepEqual([report.monthKey, report.count, report.total], ['2025-12', 2, 350]);
  assert.match(report.body, /350円/);
  assert.deepEqual([context.summarizePreviousMonth_([], new Date('2026-02-01T00:00:00Z')).count], [0]);
});

test('月次報告は営業部長の名簿メールへ送る', () => {
  const {context, sheets, sent} = fixture();
  sheets['売上報告集計シート'].rows.push([new Date('2026-09-15T00:00:00Z'), '佐藤', 'C', 'A', 'P', '商品', 2, 100, 200]);
  context.currentDate_ = () => new Date('2026-10-01T00:00:00Z');
  context.sendMonthlyReport();
  assert.equal(sent.length, 1);
  assert.equal(sent[0][0], 'director@example.com');
  assert.match(sent[0][2], /200円/);
});

test('実メール試験は9列の月次報告だけを指定アドレスへ1通送る', () => {
  const {context, sheets, sent} = fixture();
  sheets['売上報告集計シート'].rows.push([new Date('2026-09-15T00:00:00Z'), '佐藤', 'C', 'A', 'P', '商品', 2, 100, 200]);
  context.currentDate_ = () => new Date('2026-10-01T00:00:00Z');
  context.SpreadsheetApp.getUi = () => ({
    ButtonSet: {OK_CANCEL: 'set'}, Button: {OK: 'ok'},
    prompt: () => ({getSelectedButton: () => 'ok', getResponseText: () => 'test@valid.jp'}),
    alert() {}
  });
  context.runEmailDeliveryTest();
  assert.equal(sent.length, 1);
  assert.equal(sent[0][0], 'test@valid.jp');
  assert.match(sent[0][2], /200円/);
});

test('9列の初期設定は変更を予告し、確認後にトリガーを重複なく設定する', () => {
  const {context} = fixture();
  const {form} = makeForm();
  attachForm(context, form);
  context.updateSalesFormChoices();
  const triggers = [];
  const prompts = [];
  let confirmation = 'cancel';
  context.ScriptApp = {
    getProjectTriggers: () => triggers,
    newTrigger(handler) { return {
      source: null,
      forForm(value) { this.source = value.getId(); return this; }, onFormSubmit() { return this; },
      timeBased() { return this; }, onMonthDay() { return this; }, atHour() { return this; },
      nearMinute() { return this; }, inTimezone() { return this; },
      create() { const source = this.source; triggers.push({getHandlerFunction: () => handler, getTriggerSourceId: () => source}); }
    }; }
  };
  context.SpreadsheetApp.getUi = () => ({
    ButtonSet: {OK_CANCEL: 'set'}, Button: {OK: 'ok'},
    alert(...args) { if (args.length === 3) { prompts.push(args[1]); return confirmation; } }
  });
  context.setupProject();
  assert.equal(triggers.length, 0);
  assert.match(prompts[0], /送信トリガー/);
  confirmation = 'ok';
  context.setupProject();
  context.setupProject();
  assert.equal(triggers.length, 2);
  assert.equal(prompts.length, 2);
});

test('初期設定は7列や不正なマスタ見出しを拒否する', () => {
  const {context, sheets} = fixture();
  const {form} = makeForm();
  attachForm(context, form);
  sheets['売上報告集計シート'].headings = ['報告日時', '報告者名', '取引先名', '商品名', '販売個数', '販売単価', '合計金額'];
  assert.throws(() => context.setupProject(), /見出し/);
  sheets['売上報告集計シート'].headings = headers;
  sheets['商品マスタ'].headings = ['不明'];
  assert.throws(() => context.setupProject(), /見出し/);
});

test('フォーム更新前の初期設定はトリガーを作らない', () => {
  const {context} = fixture();
  const {form} = makeForm();
  attachForm(context, form);
  context.ScriptApp = {newTrigger() { throw new Error('トリガーを作成してはいけません'); }};
  assert.throws(() => context.setupProject(), /4設問が未対応/);
});

test('不正な集計金額と月次合計の桁あふれを検出する', () => {
  const {context} = fixture();
  const date = new Date('2026-09-15T00:00:00Z');
  const row = amount => [date, '佐藤', 'C', '取引先', 'P', '商品', 1, 100, amount];
  const now = new Date('2026-10-01T00:00:00Z');
  assert.throws(() => context.summarizePreviousMonth_([row('100')], now), /数値ではありません/);
  assert.throws(() => context.summarizePreviousMonth_([row(Number.MAX_SAFE_INTEGER), row(1)], now), /範囲を超え/);
});

test('商品単価と個数の積が扱える範囲を超える場合は記録しない', () => {
  const {context, sheets} = fixture();
  sheets['商品マスタ'].rows[0][2] = Number.MAX_SAFE_INTEGER;
  assert.throws(() => context.onSalesFormSubmit({response: response()}), /合計金額/);
  assert.equal(sheets['売上報告集計シート'].rows.length, 0);
});
