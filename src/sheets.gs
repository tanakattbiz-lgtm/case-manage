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
  sheet.getRange(1, 3).setValue('既定利益率');
}

function migrateLegacyProjectSheet_(sheet) {
  migrateDraftProjectSplitSheet_(sheet);

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

  sheet.insertColumnsAfter(7, 3);
  sheet.getRange(1, 8, 1, 3).setValues([['統合案件ID', 'フェーズ名', '着手金']]);
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
  sheet.getRange(1, 8).setValue('統合案件ID');
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
