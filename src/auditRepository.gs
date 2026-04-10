function listAuditRecords_() {
  return sheetToObjects_(getAuditSheet_());
}

function appendAuditRecord_(record) {
  getAuditSheet_().appendRow(toRowValues_(AUDIT_COLUMNS, record));
  return record;
}
