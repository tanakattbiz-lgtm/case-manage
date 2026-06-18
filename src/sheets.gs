let spreadsheetCache_;

function getSpreadsheet_() {
  if (!spreadsheetCache_) {
    spreadsheetCache_ = SpreadsheetApp.openById(getRequiredProperty_(SCRIPT_PROPERTY_KEYS.spreadsheetId));
  }
  return spreadsheetCache_;
}

function getSheet_(sheetName) {
  const definition = SHEET_DEFINITIONS[sheetName];
  if (!definition) {
    throwAppError_('SHEET_NOT_DEFINED', '未定義のシートです: ' + sheetName);
  }

  const spreadsheet = getSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    initializeSheet_(sheet, definition.headers, definition.hidden);
  }

  if (sheetName === SHEET_NAMES.clients) {
    migrateLegacyClientSheet_(sheet);
  }
  if (sheetName === SHEET_NAMES.projects) {
    migrateLegacyProjectSheet_(sheet);
  }

  ensureSheetHeaders_(sheet, definition.headers);
  if (definition.hidden) {
    try {
      sheet.hideSheet();
    } catch (error) {
    }
  }
  return sheet;
}

function initializeSheet_(sheet, headers, hidden) {
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1f2937')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  if (hidden) sheet.hideSheet();
}

function ensureSheetHeaders_(sheet, headers) {
  const range = sheet.getRange(1, 1, 1, headers.length);
  const currentHeaders = range.getValues()[0];
  const needsSync = headers.some(function (header, index) {
    return currentHeaders[index] !== header;
  });

  if (!needsSync) return;

  range
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1f2937')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);
}

function migrateLegacyClientSheet_(sheet) {
  const legacyHeaders = ['ID', 'クライアント名', '担当者', 'メール', '電話', '備考', '登録日'];
  const currentHeaders = sheet.getRange(1, 1, 1, legacyHeaders.length).getValues()[0];
  const isLegacy = legacyHeaders.every(function (header, index) {
    return currentHeaders[index] === header;
  });
  if (!isLegacy) return;

  sheet.insertColumnAfter(2);
  sheet.getRange(1, 3).setValue('販売手数料');
}

function migrateProjectCostColumn_(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  if (headers.indexOf('売上原価') >= 0) return;
  const profitIndex = headers.indexOf('利益') + 1;
  if (!profitIndex) return;
  sheet.insertColumnAfter(profitIndex);
  sheet.getRange(1, profitIndex + 1).setValue('売上原価');
}

function migrateLegacyProjectSheet_(sheet) {
  migrateDraftProjectSplitSheet_(sheet);
  migrateProjectIntegrationSheet_(sheet);
  migrateProjectDepositDateColumn_(sheet);
  migrateProjectCostColumn_(sheet);

  const legacyHeaders = [
    'ID',
    '案件名',
    'クライアントID',
    'クライアント名',
    '売上',
    '利益',
    'ステータス',
    '完了日',
    '備考',
    '登録日',
    '更新日',
  ];
  const currentHeaders = sheet.getRange(1, 1, 1, legacyHeaders.length).getValues()[0];
  const isLegacy = legacyHeaders.every(function (header, index) {
    return currentHeaders[index] === header;
  });
  if (!isLegacy) return;

  sheet.insertColumnsAfter(7, 4);
  sheet.getRange(1, 8, 1, 4).setValues([['親案件ID', 'フェーズ名', '着手金', '着手金入金日']]);
  migrateProjectDepositDateColumn_(sheet);
}

function migrateDraftProjectSplitSheet_(sheet) {
  const draftHeaders = [
    'ID',
    '案件名',
    'クライアントID',
    'クライアント名',
    '売上',
    '利益',
    'ステータス',
    '分割区分',
    '親案件ID',
    'フェーズ名',
    '着手金',
  ];
  const currentHeaders = sheet.getRange(1, 1, 1, draftHeaders.length).getValues()[0];
  const isDraft = draftHeaders.every(function (header, index) {
    return currentHeaders[index] === header;
  });
  if (!isDraft) return;

  sheet.deleteColumn(8);
  sheet.getRange(1, 8).setValue('親案件ID');
  migrateProjectDepositDateColumn_(sheet);
}

function migrateProjectIntegrationSheet_(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  const integrationColumnIndex = headers.indexOf('統合案件ID') + 1;
  if (!integrationColumnIndex) return;
  sheet.getRange(1, integrationColumnIndex).setValue('親案件ID');
  migrateProjectDepositDateColumn_(sheet);
}

function migrateProjectDepositDateColumn_(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  if (headers.indexOf('着手金入金日') >= 0) return;
  const depositColumnIndex = headers.indexOf('着手金') + 1;
  if (!depositColumnIndex) return;
  sheet.insertColumnAfter(depositColumnIndex);
  sheet.getRange(1, depositColumnIndex + 1).setValue('着手金入金日');
}

function getProjectsSheet_() {
  return getSheet_(SHEET_NAMES.projects);
}

function getClientsSheet_() {
  return getSheet_(SHEET_NAMES.clients);
}

function getUsersSheet_() {
  return getSheet_(SHEET_NAMES.users);
}

function getSessionsSheet_() {
  return getSheet_(SHEET_NAMES.sessions);
}

function getAuditSheet_() {
  return getSheet_(SHEET_NAMES.auditLogs);
}
