function getSheet(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#1f2937')
      .setFontColor('#fff');
    sh.setFrozenRows(1);
  }

  return sh;
}

const projSheet = () => getSheet(SN.projects, PROJ_COLS);
const clientSheet = () => getSheet(SN.clients, CLI_COLS);
