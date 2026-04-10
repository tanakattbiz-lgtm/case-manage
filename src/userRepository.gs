function listUserRecords_() {
  return sheetToObjects_(getUsersSheet_());
}

function findUserRecordById_(userId) {
  return listUserRecords_().find(function (record) {
    return record.ID === userId;
  }) || null;
}

function findUserRecordByEmail_(email) {
  const normalizedEmail = normalizeEmail_(email);
  return listUserRecords_().find(function (record) {
    return normalizeEmail_(record['メール']) === normalizedEmail;
  }) || null;
}

function appendUserRecord_(record) {
  getUsersSheet_().appendRow(toRowValues_(USER_COLUMNS, record));
  return findUserRecordById_(record.ID);
}

function updateUserRecordById_(userId, record) {
  const existing = findUserRecordById_(userId);
  if (!existing) {
    throwAppError_('USER_NOT_FOUND', 'ユーザーが見つかりません。');
  }
  getUsersSheet_()
    .getRange(existing._row, 1, 1, USER_COLUMNS.length)
    .setValues([toRowValues_(USER_COLUMNS, record)]);
  return Object.assign({ _row: existing._row }, record);
}
