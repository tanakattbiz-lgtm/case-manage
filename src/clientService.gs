function getClients(sessionToken) {
  requireReadAccess_(sessionToken);
  return sheetToObjects(clientSheet());
}

function normalizeClientProfitRate_(value) {
  if (value === '' || value == null) return '';

  const rate = Number(value);
  if (!isFinite(rate) || rate < 0) return '';

  return Math.round(rate * 100) / 100;
}

function getClientDefaultProfitRate_(clientId, clientName) {
  if (!clientId && !clientName) return '';

  const client = sheetToObjects(clientSheet()).find(row => {
    if (clientId && row['ID'] === clientId) return true;
    return !clientId && clientName && row['クライアント名'] === clientName;
  });

  return client ? normalizeClientProfitRate_(client['既定利益率']) : '';
}

function addClient(sessionToken, c) {
  requireEditAccess_(sessionToken);
  const id = genId('CLI');
  const defaultProfitRate = normalizeClientProfitRate_(c['既定利益率']);

  clientSheet().appendRow([
    id,
    c['クライアント名'] || '',
    defaultProfitRate,
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
  const defaultProfitRate = normalizeClientProfitRate_(c['既定利益率']);

  clientSheet().getRange(row, 1, 1, CLI_COLS.length).setValues([[
    c['ID'],
    c['クライアント名'] || '',
    defaultProfitRate,
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
