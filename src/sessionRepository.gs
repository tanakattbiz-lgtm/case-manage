function listSessionRecords_() {
  return sheetToObjects_(getSessionsSheet_());
}

function findSessionRecordById_(sessionId) {
  return listSessionRecords_().find(function (record) {
    return record.ID === sessionId;
  }) || null;
}

function findSessionRecordByHash_(tokenHash) {
  return listSessionRecords_().find(function (record) {
    return record['トークンハッシュ'] === tokenHash;
  }) || null;
}

function listSessionRecordsByUserId_(userId) {
  return listSessionRecords_().filter(function (record) {
    return record['ユーザーID'] === userId;
  });
}

function appendSessionRecord_(record) {
  getSessionsSheet_().appendRow(toRowValues_(SESSION_COLUMNS, record));
  return findSessionRecordById_(record.ID);
}

function updateSessionRecordById_(sessionId, record) {
  const existing = findSessionRecordById_(sessionId);
  if (!existing) {
    throwAppError_('SESSION_NOT_FOUND', 'セッションが見つかりません。');
  }
  getSessionsSheet_()
    .getRange(existing._row, 1, 1, SESSION_COLUMNS.length)
    .setValues([toRowValues_(SESSION_COLUMNS, record)]);
  return Object.assign({ _row: existing._row }, record);
}
