function listClientRecords_() {
  return sheetToObjects_(getClientsSheet_());
}

function findClientRecordById_(clientId) {
  return listClientRecords_().find(function (record) {
    return record.ID === clientId;
  }) || null;
}

function appendClientRecord_(record) {
  getClientsSheet_().appendRow(toRowValues_(CLIENT_COLUMNS, record));
  return findClientRecordById_(record.ID);
}

function updateClientRecordById_(clientId, record) {
  const existing = findClientRecordById_(clientId);
  if (!existing) {
    throwAppError_('CLIENT_NOT_FOUND', 'クライアントが見つかりません。');
  }
  getClientsSheet_()
    .getRange(existing._row, 1, 1, CLIENT_COLUMNS.length)
    .setValues([toRowValues_(CLIENT_COLUMNS, record)]);
  return Object.assign({ _row: existing._row }, record);
}

function deleteClientRecordById_(clientId) {
  const existing = findClientRecordById_(clientId);
  if (!existing) {
    throwAppError_('CLIENT_NOT_FOUND', 'クライアントが見つかりません。');
  }
  getClientsSheet_().deleteRow(existing._row);
  return true;
}
