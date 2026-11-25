/** 和み製菓 営業部プロトタイプ。既存スプレッドシートに紐づけて使用する。 */
const NAGOMI = Object.freeze({
  salesSheetNames: ['売上報告集計シート', '売上報告集計'],
  rosterSheetName: '社員名簿',
  salesHeaders: ['報告日時', '報告者名', '取引先ID', '取引先名', '商品ID', '商品名', '販売個数', '販売単価', '合計金額'],
  rosterHeaders: ['社員名', '部署', '役職', 'メールアドレス'],
  customerSheetName: '取引先マスタ',
  customerHeaders: ['取引先ID', '取引先名', '取引先区分'],
  productSheetName: '商品マスタ',
  productHeaders: ['商品ID', '商品名', '標準単価'],
  salesFormProperty: 'NAGOMI_SALES_FORM_ID',
  timezone: 'Asia/Tokyo'
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('営業部ツール')
    .addItem('売上報告フォームの選択肢を更新', 'updateSalesFormChoicesFromMenu')
    .addToUi();
}

function updateSalesFormChoicesFromMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const formId = PropertiesService.getScriptProperties().getProperty(NAGOMI.salesFormProperty);
    if (!formId) throw new Error('既存の売上報告フォームIDが設定されていません。');
    const answer = ui.alert('売上フォームの更新', 'フォームID: ' + formId + '\n報告者名・取引先名・商品名をマスタの選択肢で更新します。テキスト設問はプルダウンへ置き換え、販売個数を正の整数に制限し、販売単価の設問を削除します。', ui.ButtonSet.OK_CANCEL);
    if (answer !== ui.Button.OK) return;
    updateSalesFormChoices();
    ui.alert('売上報告フォームの選択肢と設問を更新しました。');
  } catch (error) {
    ui.alert('売上報告フォームを更新できませんでした。\n' + (error.message || String(error)));
  }
}

/** 既存の9列シート、マスタ、フォームを確認して不足トリガーを登録する。 */
function setupProject() {
  const spreadsheet = requireSpreadsheet_();
  const salesSheet = requireSalesSheet_(spreadsheet);
  requireHeaders_(salesSheet, NAGOMI.salesHeaders);
  requireHeaders_(requireRosterSheet_(spreadsheet), NAGOMI.rosterHeaders);
  getSalesEmployeeNames_();
  readMaster_(spreadsheet, NAGOMI.customerSheetName, NAGOMI.customerHeaders);
  readMaster_(spreadsheet, NAGOMI.productSheetName, NAGOMI.productHeaders);
  const salesForm = requireSalesForm_();
  requireOperationalForm_(salesForm);
  const ui = SpreadsheetApp.getUi();
  const triggers = ScriptApp.getProjectTriggers();
  const addSales = !hasSalesTrigger_(triggers, salesForm);
  const addMonthly = !hasMonthlyTrigger_(triggers);
  const changes = [
    addSales && '売上報告フォームの送信トリガーを1件作成',
    addMonthly && '毎月1日9時ごろの月次報告トリガーを1件作成'
  ].filter(Boolean);
  if (changes.length === 0) {
    ui.alert('既存のフォーム・シート・必要なトリガーを確認しました。Google側の変更はありません。');
    return;
  }
  const answer = ui.alert('初期設定', '既存のフォーム・シートは変更しません。実行する変更:\n' + changes.join('\n'), ui.ButtonSet.OK_CANCEL);
  if (answer !== ui.Button.OK) return;
  if (addSales) ensureSalesTrigger_(salesForm);
  if (addMonthly) ensureMonthlyTrigger_();
  ui.alert('必要なトリガーを確認・登録しました。');
}

function requireSalesForm_() {
  const formId = PropertiesService.getScriptProperties().getProperty(NAGOMI.salesFormProperty);
  if (!formId) throw new Error('既存の売上報告フォームIDが設定されていません。');
  return FormApp.openById(formId);
}

function requireOperationalForm_(form) {
  const expected = {
    '報告者名': FormApp.ItemType.LIST,
    '取引先名': FormApp.ItemType.LIST,
    '商品名': FormApp.ItemType.LIST,
    '販売個数': FormApp.ItemType.TEXT
  };
  const items = form.getItems();
  if (items.length !== 4 || Object.keys(expected).some(title =>
    items.filter(item => item.getTitle() === title && item.getType() === expected[title]).length !== 1
  )) {
    throw new Error('売上報告フォームの4設問が未対応です。先にメニューから選択肢を更新してください。');
  }
}

/** マスタ変更後に手動実行する。既存フォームの4設問だけを使用する。 */
function updateSalesFormChoices() {
  const spreadsheet = requireSpreadsheet_();
  const choices = {
    '報告者名': getSalesEmployeeNames_(),
    '取引先名': readMaster_(spreadsheet, NAGOMI.customerSheetName, NAGOMI.customerHeaders).map(item => masterChoice_(item)),
    '商品名': readMaster_(spreadsheet, NAGOMI.productSheetName, NAGOMI.productHeaders).map(item => masterChoice_(item))
  };
  const form = requireSalesForm_();
  const items = form.getItems();
  Object.keys(choices).forEach(title => validateChoiceItem_(items, title));
  const quantityItems = items.filter(item => item.getTitle() === '販売個数');
  if (quantityItems.length !== 1 || quantityItems[0].getType() !== FormApp.ItemType.TEXT) {
    throw new Error('「販売個数」はテキスト設問を1つだけ用意してください。');
  }
  const priceItems = items.filter(item => item.getTitle() === '販売単価');
  if (priceItems.length > 1) throw new Error('「販売単価」設問が重複しています。');
  const allowed = [...Object.keys(choices), '販売個数', '販売単価'];
  if (items.some(item => !allowed.includes(item.getTitle()))) {
    throw new Error('売上報告フォームに想定外の設問があります。設問を確認してください。');
  }
  Object.keys(choices).forEach(title => updateChoiceItem_(form, title, choices[title]));
  quantityItems[0].asTextItem().setRequired(true).setValidation(
    FormApp.createTextValidation().requireWholeNumber().requireNumberGreaterThan(0)
      .setHelpText('1以上の整数を入力してください。').build()
  );
  if (priceItems.length === 1) form.deleteItem(priceItems[0]);
}

function validateChoiceItem_(items, title) {
  const matches = items.filter(item => item.getTitle() === title);
  const textItems = matches.filter(item => item.getType() === FormApp.ItemType.TEXT);
  const listItems = matches.filter(item => item.getType() === FormApp.ItemType.LIST);
  if (!((matches.length === 1 && listItems.length === 1) ||
      (textItems.length === 1 && listItems.length <= 1 && matches.length <= 2))) {
    throw new Error('売上報告フォームの「' + title + '」設問の種類と重複を確認してください。');
  }
}

function updateChoiceItem_(form, title, values) {
  const matches = form.getItems().filter(item => item.getTitle() === title);
  const textItems = matches.filter(item => item.getType() === FormApp.ItemType.TEXT);
  const listItems = matches.filter(item => item.getType() === FormApp.ItemType.LIST);
  if (matches.length === 1 && listItems.length === 1) {
    listItems[0].asListItem().setChoiceValues(values).setRequired(true);
  } else {
    // 設問追加後に失敗しても、次回は追加済みのプルダウンを再利用する。
    const replacement = listItems.length === 1 ? listItems[0].asListItem() : form.addListItem().setTitle(title);
    replacement.setRequired(true).setChoiceValues(values);
    form.moveItem(replacement.getIndex(), textItems[0].getIndex());
    form.deleteItem(textItems[0]);
  }
}

function getSalesEmployeeNames_() {
  const sheet = requireRosterSheet_(requireSpreadsheet_());
  requireHeaders_(sheet, NAGOMI.rosterHeaders);
  const names = sheet.getDataRange().getValues().slice(1)
    .filter(row => String(row[1]).trim() === '営業部')
    .map(row => String(row[0]).trim())
    .filter(Boolean);
  if (names.length === 0) throw new Error('社員名簿に営業部の社員名がありません。');
  if (new Set(names).size !== names.length) throw new Error('社員名簿の営業部に同じ社員名が複数あります。');
  return names;
}

function readMaster_(spreadsheet, sheetName, headers) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new Error('「' + sheetName + '」シートが見つかりません。');
  requireHeaders_(sheet, headers);
  const seen = new Set();
  const entries = sheet.getDataRange().getValues().slice(1)
    .map((row, index) => ({row, rowNumber: index + 2}))
    .filter(({row}) => row.some(value => String(value == null ? '' : value).trim()));
  if (entries.length === 0) throw new Error('「' + sheetName + '」にデータがありません。');
  return entries.map(({row, rowNumber}) => {
    const id = String(row[0] == null ? '' : row[0]).trim();
    const name = String(row[1] == null ? '' : row[1]).trim();
    if (!id || !name || id.includes('｜') || name.includes('｜')) {
      throw new Error('「' + sheetName + '」の' + rowNumber + '行目のID・名称を確認してください。');
    }
    if (seen.has(id)) throw new Error('「' + sheetName + '」のID「' + id + '」が重複しています。');
    seen.add(id);
    const entry = {id, name};
    if (sheetName === NAGOMI.productSheetName) {
      const rawPrice = row[2];
      const priceText = String(rawPrice == null ? '' : rawPrice).trim();
      const price = Number(rawPrice);
      if (!/^(0|[1-9]\d*)$/.test(priceText) || !Number.isSafeInteger(price) || price < 0) {
        throw new Error('「商品マスタ」の' + rowNumber + '行目の標準単価を確認してください。');
      }
      entry.price = price;
    }
    return entry;
  });
}

function masterChoice_(entry) {
  return entry.id + '｜' + entry.name;
}

function resolveMasterChoice_(value, entries, label) {
  const selected = String(value == null ? '' : value).trim();
  if (!selected) throw new Error(label + 'が空欄です。');
  const separator = selected.indexOf('｜');
  if (separator < 1) throw new Error(label + 'は「ID｜名称」形式で選択してください。');
  const id = selected.slice(0, separator);
  const entry = entries.find(item => item.id === id);
  if (!entry) throw new Error(label + 'のID「' + id + '」がマスタに登録されていません。');
  if (selected !== masterChoice_(entry)) throw new Error(label + 'の名称がマスタと一致しません。選択肢を更新してください。');
  return entry;
}

function hasSalesTrigger_(triggers, form) {
  return triggers.some(trigger =>
    trigger.getHandlerFunction() === 'onSalesFormSubmit' &&
    trigger.getTriggerSourceId() === form.getId()
  );
}

function hasMonthlyTrigger_(triggers) {
  return triggers.some(trigger => trigger.getHandlerFunction() === 'sendMonthlyReport');
}

function ensureSalesTrigger_(form) {
  if (!hasSalesTrigger_(ScriptApp.getProjectTriggers(), form)) {
    ScriptApp.newTrigger('onSalesFormSubmit').forForm(form).onFormSubmit().create();
  }
}

function ensureMonthlyTrigger_() {
  if (!hasMonthlyTrigger_(ScriptApp.getProjectTriggers())) {
    ScriptApp.newTrigger('sendMonthlyReport').timeBased()
      .onMonthDay(1).atHour(9).nearMinute(0).inTimezone(NAGOMI.timezone).create();
  }
}

/** 売上フォーム専用のインストール型トリガー。フォームの回答シートは作らない。 */
function onSalesFormSubmit(event) {
  if (!event || !event.response) throw new Error('売上フォームの送信イベントがありません。');
  const response = event.response;
  const answers = {};
  response.getItemResponses().forEach(itemResponse => {
    answers[itemResponse.getItem().getTitle()] = itemResponse.getResponse();
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const spreadsheet = requireSpreadsheet_();
    const sheet = requireSalesSheet_(spreadsheet);
    requireHeaders_(sheet, NAGOMI.salesHeaders);
    const row = buildSalesRow_(answers, response.getTimestamp(), spreadsheet);
    // 文字列の先頭が「=」の場合、シートの数式として評価させない。
    sheet.appendRow(row.map((value, index) =>
      index >= 1 && index <= 5 && value.startsWith('=') ? "'" + value : value
    ));
  } finally {
    lock.releaseLock();
  }
}

function buildSalesRow_(answers, timestamp, spreadsheet) {
  if (!(timestamp instanceof Date) || isNaN(timestamp.getTime())) {
    throw new Error('報告日時が取得できません。');
  }
  const reporter = String(answers['報告者名'] == null ? '' : answers['報告者名']).trim();
  if (!reporter || !getSalesEmployeeNames_().includes(reporter)) {
    throw new Error('報告者名が社員名簿の営業部員に登録されていません。');
  }
  const customer = resolveMasterChoice_(answers['取引先名'],
    readMaster_(spreadsheet, NAGOMI.customerSheetName, NAGOMI.customerHeaders), '取引先名');
  const product = resolveMasterChoice_(answers['商品名'],
    readMaster_(spreadsheet, NAGOMI.productSheetName, NAGOMI.productHeaders), '商品名');
  const quantity = parseNonnegativeNumber_(answers['販売個数'], '販売個数');
  if (!Number.isSafeInteger(quantity) || quantity < 1) throw new Error('販売個数は1以上の整数にしてください。');
  const amount = quantity * product.price;
  if (!Number.isSafeInteger(amount)) {
    throw new Error('合計金額が扱える範囲を超えています。');
  }
  return [timestamp, reporter, customer.id, customer.name, product.id, product.name, quantity, product.price, amount];
}

function parseNonnegativeNumber_(value, label) {
  const text = String(value == null ? '' : value).trim();
  if (!text) throw new Error(label + 'が空欄です。');
  const number = Number(text);
  if (!Number.isFinite(number) || number < 0) throw new Error(label + 'は0以上の数値にしてください。');
  return number;
}

/** 毎月1日、前月分を営業部長へ送る。 */
function sendMonthlyReport() {
  const spreadsheet = requireSpreadsheet_();
  const sheet = requireSalesSheet_(spreadsheet);
  requireHeaders_(sheet, NAGOMI.salesHeaders);
  const rows = sheet.getDataRange().getValues().slice(1);
  const report = summarizePreviousMonth_(rows, currentDate_());
  const director = findSalesDirector_(requireRosterSheet_(spreadsheet));
  sendMonthlyReportTo_(director.email, report);
}

function currentDate_() {
  return new Date();
}

function sendMonthlyReportTo_(email, report) {
  MailApp.sendEmail(email, report.subject, report.body);
}

function summarizePreviousMonth_(rows, now) {
  const currentYear = Number(Utilities.formatDate(now, NAGOMI.timezone, 'yyyy'));
  const currentMonth = Number(Utilities.formatDate(now, NAGOMI.timezone, 'M'));
  const previous = new Date(Date.UTC(currentYear, currentMonth - 2, 1));
  const monthKey = Utilities.formatDate(previous, 'UTC', 'yyyy-MM');
  let count = 0;
  let total = 0;
  rows.forEach((row, index) => {
    const date = row[0];
    if (!(date instanceof Date) || isNaN(date.getTime())) return;
    if (Utilities.formatDate(date, NAGOMI.timezone, 'yyyy-MM') !== monthKey) return;
    const amount = row[8];
    if (typeof amount !== 'number' || !Number.isSafeInteger(amount) || amount < 0) {
      throw new Error('売上報告集計の' + (index + 2) + '行目の合計金額が数値ではありません。');
    }
    count += 1;
    total += amount;
    if (!Number.isSafeInteger(total)) throw new Error('月次の総売上額が扱える範囲を超えています。');
  });
  const label = monthKey.replace('-', '年') + '月';
  return {
    monthKey,
    count,
    total,
    subject: '【和み製菓】' + label + ' 売上月次報告',
    body: label + 'の売上報告です。\n\n総売上額: ' + total.toLocaleString('ja-JP') + '円\n売上件数: ' + count + '件'
  };
}

/** GASエディタから手動実行する月次報告の実メール試験。 */
function runEmailDeliveryTest() {
  const ui = SpreadsheetApp.getUi();
  const answer = ui.prompt('月次報告メール試験', '有効なテスト用メールアドレスを入力してください。月次報告を1通送ります。', ui.ButtonSet.OK_CANCEL);
  if (answer.getSelectedButton() !== ui.Button.OK) return;
  const email = String(answer.getResponseText()).trim();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) || /(?:@|\.)example\.com$/i.test(email)) {
    throw new Error('有効なテスト用メールアドレスを入力してください。');
  }
  const spreadsheet = requireSpreadsheet_();
  const salesSheet = requireSalesSheet_(spreadsheet);
  requireHeaders_(salesSheet, NAGOMI.salesHeaders);
  const report = summarizePreviousMonth_(salesSheet.getDataRange().getValues().slice(1), currentDate_());
  sendMonthlyReportTo_(email, report);
  ui.alert('テスト用アドレスへ月次報告を1通送信しました。受信を確認してください。');
}

function findSalesDirector_(sheet) {
  requireHeaders_(sheet, NAGOMI.rosterHeaders);
  const matches = sheet.getDataRange().getValues().slice(1).filter(row =>
    String(row[1]).trim() === '営業部' && String(row[2]).trim() === '部長'
  );
  if (matches.length !== 1) throw new Error('社員名簿の営業部長は1名である必要があります。');
  return employeeFromRow_(matches[0]);
}

function employeeFromRow_(row) {
  const name = String(row[0]).trim();
  const email = String(row[3]).trim();
  if (!email || !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
    throw new Error('社員名簿の「' + name + '」のメールアドレスを確認してください。');
  }
  return {name, email};
}

function requireSpreadsheet_() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error('既存スプレッドシートに紐づくGASプロジェクトで実行してください。');
  return spreadsheet;
}

function requireSalesSheet_(spreadsheet) {
  const matches = NAGOMI.salesSheetNames.map(name => spreadsheet.getSheetByName(name)).filter(Boolean);
  if (matches.length !== 1) throw new Error('売上報告集計シートが見つからないか、候補名のシートが2つあります。');
  return matches[0];
}

function requireRosterSheet_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(NAGOMI.rosterSheetName);
  if (!sheet) throw new Error('社員名簿シートが見つかりません。');
  return sheet;
}

function requireHeaders_(sheet, expected) {
  const actual = sheet.getRange(1, 1, 1, expected.length).getValues()[0];
  if (expected.some((value, index) => String(actual[index]).trim() !== value)) {
    throw new Error('「' + sheet.getName() + '」の見出しが想定した並びと異なります。');
  }
}
