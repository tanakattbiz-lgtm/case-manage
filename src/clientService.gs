function getClients() {
  return sheetToObjects(clientSheet());
}

function addClient(c) {
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

function updateClient(c) {
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

function deleteClient(row) {
  clientSheet().deleteRow(Number(row));
  return { success: true };
}
