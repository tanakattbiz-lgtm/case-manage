function getClients(sessionToken) {
  requireReadAccess_(sessionToken);
  return listClientRecords_()
    .sort(function (a, b) {
      return String(a['クライアント名'] || '').localeCompare(String(b['クライアント名'] || ''), 'ja');
    })
    .map(clientRecordToDto_);
}

function addClient(sessionToken, payload) {
  return runWithScriptLock_(function () {
    requireEditAccess_(sessionToken);
    const clientRecord = clientPayloadToRecord_(payload, {
      id: genId('CLI'),
      createdAt: nowDateTimeStr_(),
    });
    appendClientRecord_(clientRecord);
    return { success: true, id: clientRecord.ID };
  });
}

function updateClient(sessionToken, payload) {
  return runWithScriptLock_(function () {
    requireEditAccess_(sessionToken);
    const input = payload || {};
    const existing = findClientRecordById_(String(input.id || ''));
    if (!existing) {
      throwAppError_('CLIENT_NOT_FOUND', 'クライアントが見つかりません。');
    }
    const clientRecord = clientPayloadToRecord_(input, {
      id: existing.ID,
      createdAt: existing['登録日'] || nowDateTimeStr_(),
    });
    updateClientRecordById_(existing.ID, clientRecord);
    return { success: true };
  });
}

function deleteClient(sessionToken, clientId) {
  return runWithScriptLock_(function () {
    requireEditAccess_(sessionToken);
    deleteClientRecordById_(String(clientId || ''));
    return { success: true };
  });
}

function clientPayloadToRecord_(payload, options) {
  const input = payload || {};
  const settings = options || {};
  const name = normalizeString_(input.name);
  if (!name) {
    throwAppError_('CLIENT_NAME_REQUIRED', 'クライアント名を入力してください。');
  }

  return {
    ID: settings.id,
    'クライアント名': name,
    '既定利益率': normalizeClientProfitRate_(input.defaultProfitRate),
    '担当者': normalizeString_(input.contactName),
    'メール': normalizeEmail_(input.email),
    '電話': normalizeString_(input.phone),
    '備考': normalizeTextArea_(input.note, 1000),
    '登録日': settings.createdAt,
  };
}

function clientRecordToDto_(record) {
  return {
    id: record.ID,
    row: record._row,
    name: record['クライアント名'] || '',
    defaultProfitRate: normalizeClientProfitRate_(record['既定利益率']),
    contactName: record['担当者'] || '',
    email: record['メール'] || '',
    phone: record['電話'] || '',
    note: record['備考'] || '',
    createdAt: record['登録日'] || '',
  };
}

function normalizeClientProfitRate_(value) {
  if (value === '' || value == null) return '';
  const normalized = Number(value);
  if (!isFinite(normalized) || normalized < 0) return '';
  return Math.round(normalized * 100) / 100;
}
