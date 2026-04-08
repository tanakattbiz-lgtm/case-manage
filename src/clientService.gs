function getClients(sessionToken) {
  requireReadAccess_(sessionToken);
  return sheetToObjects(clientSheet());
}

function addClient(sessionToken, c) {
  requireEditAccess_(sessionToken);
  const id = genId('CLI');

  clientSheet().appendRow([
    id,
    c['クライアント名'] || '',
    c['担当者'] || '',
    c['メール'] || '',
    c['電話'] || '',
    c['備考'] || '',
    nowStr(),
  ]);

  return { success: true, id };
}

function updateClient(sessionToken, c) {
  requireEditAccess_(sessionToken);
  const row = Number(c['_row']);
  if (!row) return { success: false };

  clientSheet().getRange(row, 1, 1, CLI_COLS.length).setValues([[
    c['ID'],
    c['クライアント名'] || '',
    c['担当者'] || '',
    c['メール'] || '',
    c['電話'] || '',
    c['備考'] || '',
    c['登録日'] || nowStr(),
  ]]);

  return { success: true };
}

function deleteClient(sessionToken, row) {
  requireEditAccess_(sessionToken);
  clientSheet().deleteRow(Number(row));
  return { success: true };
}
