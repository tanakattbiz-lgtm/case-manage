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

  return sh;
}

const projSheet = () => getSheet(SN.projects, PROJ_COLS);
const clientSheet = () => getSheet(SN.clients, CLI_COLS);
const userSheet = () => getSheet(SN.users, USER_COLS, { hidden: true });
const sessionSheet = () => getSheet(SN.sessions, SESSION_COLS, { hidden: true });
const auditSheet = () => getSheet(SN.auditLogs, AUDIT_COLS, { hidden: true });
