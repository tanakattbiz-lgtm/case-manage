function getSheet(name, headers, options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  const opts = options || {};

  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1f2937')
      .setFontColor('#fff');
    sh.setFrozenRows(1);
    if (opts.hidden) sh.hideSheet();
  }

  if (name === SN.clients) migrateLegacyClientSheet_(sh);
  ensureSheetHeaders_(sh, headers);

  return sh;
}

function migrateLegacyClientSheet_(sh) {
  const legacyHeaders = ['ID', 'クライアント名', '担当者', 'メール', '電話', '備考', '登録日'];
  const currentHeaders = sh.getRange(1, 1, 1, legacyHeaders.length).getValues()[0];
  const isLegacySchema = legacyHeaders.every((header, index) => currentHeaders[index] === header);

  if (!isLegacySchema) return;

  sh.insertColumnAfter(2);
  sh.getRange(1, 3).setValue('既定利益率');
}

function ensureSheetHeaders_(sh, headers) {
  const headerRange = sh.getRange(1, 1, 1, headers.length);
  const currentHeaders = headerRange.getValues()[0];
  const needsSync = headers.some((header, index) => currentHeaders[index] !== header);

  if (!needsSync) return;

  headerRange
    .setValues([headers])
    .setFontWeight('bold')
    .setBackground('#1f2937')
    .setFontColor('#fff');
  sh.setFrozenRows(1);
}

const projSheet = () => getSheet(SN.projects, PROJ_COLS);
const clientSheet = () => getSheet(SN.clients, CLI_COLS);
const userSheet = () => getSheet(SN.users, USER_COLS, { hidden: true });
const sessionSheet = () => getSheet(SN.sessions, SESSION_COLS, { hidden: true });
const auditSheet = () => getSheet(SN.auditLogs, AUDIT_COLS, { hidden: true });
