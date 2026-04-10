function getProjects(sessionToken) {
  requireReadAccess_(sessionToken);
  return listProjectRecords_()
    .sort(function (a, b) {
      const left = parseDateValue_(a['完了日'] || a['登録日']);
      const right = parseDateValue_(b['完了日'] || b['登録日']);
      return (right ? right.getTime() : 0) - (left ? left.getTime() : 0);
    })
    .map(projectRecordToDto_);
}

function addProject(sessionToken, payload) {
  return runWithScriptLock_(function () {
    requireEditAccess_(sessionToken);
    const now = nowDateTimeStr_();
    const projectRecord = projectPayloadToRecord_(payload, {
      id: genId('PRJ'),
      createdAt: now,
      updatedAt: now,
    });
    appendProjectRecord_(projectRecord);
    invalidateDashboardCache_();
    return { success: true, id: projectRecord.ID };
  });
}

function updateProject(sessionToken, payload) {
  return runWithScriptLock_(function () {
    requireEditAccess_(sessionToken);
    const input = payload || {};
    const existing = findProjectRecordById_(String(input.id || ''));
    if (!existing) {
      throwAppError_('PROJECT_NOT_FOUND', '案件が見つかりません。');
    }

    const projectRecord = projectPayloadToRecord_(input, {
      id: existing.ID,
      createdAt: existing['登録日'] || nowDateTimeStr_(),
      updatedAt: nowDateTimeStr_(),
    });
    updateProjectRecordById_(existing.ID, projectRecord);
    invalidateDashboardCache_();
    return { success: true };
  });
}

function deleteProject(sessionToken, projectId) {
  return runWithScriptLock_(function () {
    requireEditAccess_(sessionToken);
    deleteProjectRecordById_(String(projectId || ''));
    invalidateDashboardCache_();
    return { success: true };
  });
}

function projectPayloadToRecord_(payload, options) {
  const input = payload || {};
  const settings = options || {};
  const name = normalizeString_(input.name);
  if (!name) {
    throwAppError_('PROJECT_NAME_REQUIRED', '案件名を入力してください。');
  }

  const clientId = normalizeString_(input.clientId);
  const clientRecord = clientId ? findClientRecordById_(clientId) : null;
  if (clientId && !clientRecord) {
    throwAppError_('CLIENT_NOT_FOUND', '選択したクライアントが見つかりません。');
  }

  const sales = Number(normalizeNonNegativeNumber_(input.sales, 0)) || 0;
  const requestedProfit = input.profit === '' || input.profit == null ? '' : normalizeNonNegativeNumber_(input.profit, 0);
  const defaultProfitRate = clientRecord ? normalizeClientProfitRate_(clientRecord['既定利益率']) : '';
  const resolvedProfit = requestedProfit === ''
    ? (defaultProfitRate === '' ? sales : Math.round(sales * Number(defaultProfitRate) / 100))
    : Number(requestedProfit || 0);

  const status = normalizeProjectStatus_(input.status);
  let completedAt = normalizeDateInput_(input.completedAt);
  if (status === PROJECT_STATUSES.completed && !completedAt) {
    completedAt = nowDateStr_();
  }

  return {
    ID: settings.id,
    '案件名': name,
    'クライアントID': clientId,
    'クライアント名': clientRecord ? clientRecord['クライアント名'] : '',
    '売上': sales,
    '利益': resolvedProfit,
    'ステータス': status,
    '完了日': completedAt,
    '備考': normalizeTextArea_(input.note, 1000),
    '登録日': settings.createdAt,
    '更新日': settings.updatedAt,
  };
}

function projectRecordToDto_(record) {
  return {
    id: record.ID,
    row: record._row,
    name: record['案件名'] || '',
    clientId: record['クライアントID'] || '',
    clientName: record['クライアント名'] || '',
    sales: Number(record['売上']) || 0,
    profit: Number(record['利益']) || 0,
    status: normalizeProjectStatus_(record['ステータス']),
    completedAt: record['完了日'] || '',
    note: record['備考'] || '',
    createdAt: record['登録日'] || '',
    updatedAt: record['更新日'] || '',
  };
}

function normalizeProjectStatus_(status) {
  const normalized = normalizeString_(status);
  return PROJECT_STATUS_LIST.indexOf(normalized) >= 0
    ? normalized
    : PROJECT_STATUSES.lead;
}
