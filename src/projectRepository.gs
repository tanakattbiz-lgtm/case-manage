function listProjectRecords_() {
  return sheetToObjects_(getProjectsSheet_());
}

function findProjectRecordById_(projectId) {
  return listProjectRecords_().find(function (record) {
    return record.ID === projectId;
  }) || null;
}

function appendProjectRecord_(record) {
  getProjectsSheet_().appendRow(toRowValues_(PROJECT_COLUMNS, record));
  return findProjectRecordById_(record.ID);
}

function updateProjectRecordById_(projectId, record) {
  const existing = findProjectRecordById_(projectId);
  if (!existing) {
    throwAppError_('PROJECT_NOT_FOUND', '案件が見つかりません。');
  }
  getProjectsSheet_()
    .getRange(existing._row, 1, 1, PROJECT_COLUMNS.length)
    .setValues([toRowValues_(PROJECT_COLUMNS, record)]);
  return Object.assign({ _row: existing._row }, record);
}

function deleteProjectRecordById_(projectId) {
  const existing = findProjectRecordById_(projectId);
  if (!existing) {
    throwAppError_('PROJECT_NOT_FOUND', '案件が見つかりません。');
  }
  getProjectsSheet_().deleteRow(existing._row);
  return true;
}
