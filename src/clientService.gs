function getClients(sessionToken, query) {
  requireReadAccess_(sessionToken);
  const normalizedQuery = normalizeClientListQuery_(query);
  const statsMap = buildClientStatsMap_(listProjectDtos_());
  const allItems = listClientDtos_(statsMap);
  const filteredItems = allItems.filter(function (client) {
    if (!normalizedQuery.query) return true;
    const joined = [
      client.name,
      client.contactName,
      client.email,
      client.phone,
      client.note,
    ].join(' ').toLowerCase();
    return joined.indexOf(normalizedQuery.query.toLowerCase()) >= 0;
  });
  const paged = paginateItems_(filteredItems, normalizedQuery.page, normalizedQuery.pageSize);

  return {
    items: paged.items,
    pagination: paged.pagination,
    filters: normalizedQuery,
    summary: buildClientListSummary_(filteredItems, allItems.length),
    fetchedAt: nowDateTimeStr_(),
  };
}

function getClientOptions(sessionToken) {
  requireReadAccess_(sessionToken);
  return listClientDtos_({})
    .slice(0, FIXED_VALUES.lists.clientOptionLimit)
    .map(function (client) {
      return {
        id: client.id,
        name: client.name,
        commissionRate: client.commissionRate,
      };
    });
}

function getClientDetail(sessionToken, clientId) {
  requireReadAccess_(sessionToken);
  const targetId = normalizeString_(clientId);
  const statsMap = buildClientStatsMap_(listProjectDtos_());
  const client = listClientDtos_(statsMap).find(function (item) {
    return item.id === targetId;
  }) || null;

  if (!client) {
    throwAppError_('CLIENT_NOT_FOUND', 'クライアントが見つかりません。');
  }

  const relatedProjects = listProjectDtos_()
    .filter(function (project) {
      return project.clientId === client.id;
    })
    .slice(0, FIXED_VALUES.lists.detailRelatedLimit);

  return {
    client: client,
    relatedProjects: relatedProjects,
  };
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

function listClientDtos_(statsMap) {
  const relatedStats = statsMap || {};
  return listClientRecords_()
    .sort(function (a, b) {
      return String(a['クライアント名'] || '').localeCompare(String(b['クライアント名'] || ''), 'ja');
    })
    .map(function (record) {
      return clientRecordToDto_(record, relatedStats[record.ID] || null);
    });
}

function normalizeClientListQuery_(query) {
  const input = query || {};
  return {
    page: normalizePositiveInteger_(input.page, 1, 1, 999),
    pageSize: normalizePositiveInteger_(
      input.pageSize,
      FIXED_VALUES.lists.defaultPageSize,
      1,
      FIXED_VALUES.lists.maxPageSize
    ),
    query: normalizeString_(input.query),
    sort: {
      field: 'name',
      direction: 'asc',
      label: 'クライアント名順',
    },
  };
}

function buildClientStatsMap_(projects) {
  const recognizedMap = buildProjectRevenueEvents_(filterRevenueProjects_(projects || [])).reduce(function (map, event) {
    if (!map[event.projectId]) map[event.projectId] = { sales: 0, profit: 0 };
    map[event.projectId].sales += Number(event.amount) || 0;
    map[event.projectId].profit += Number(event.profit) || 0;
    return map;
  }, {});

  return (projects || []).reduce(function (map, project) {
    if (!project.clientId) return map;
    if (!map[project.clientId]) {
      map[project.clientId] = {
        projectCount: 0,
        activeProjectCount: 0,
        completedProjectCount: 0,
        salesTotal: 0,
        profitTotal: 0,
        lastProjectAt: '',
      };
    }

    const bucket = map[project.clientId];
    bucket.projectCount += 1;
    if (project.status === PROJECT_STATUSES.active) bucket.activeProjectCount += 1;
    if (project.status === PROJECT_STATUSES.completed) bucket.completedProjectCount += 1;
    if (recognizedMap[project.id]) {
      bucket.salesTotal += Number(recognizedMap[project.id].sales) || 0;
      bucket.profitTotal += Number(recognizedMap[project.id].profit) || 0;
    }

    const currentDate = parseDateValue_(bucket.lastProjectAt);
    const nextDate = parseDateValue_(project.targetDate);
    if (!currentDate || (nextDate && nextDate.getTime() > currentDate.getTime())) {
      bucket.lastProjectAt = project.targetDate;
    }
    return map;
  }, {});
}

function buildClientListSummary_(filteredItems, totalItems) {
  return {
    totalItems: Number(totalItems) || 0,
    filteredItems: filteredItems.length,
    totalProjectCount: filteredItems.reduce(function (sum, client) {
      return sum + (Number(client.projectCount) || 0);
    }, 0),
    totalSales: filteredItems.reduce(function (sum, client) {
      return sum + (Number(client.salesTotal) || 0);
    }, 0),
  };
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
    '販売手数料': normalizeCommissionRate_(input.commissionRate),
    '担当者': normalizeString_(input.contactName),
    'メール': normalizeEmail_(input.email),
    '電話': normalizeString_(input.phone),
    '備考': normalizeTextArea_(input.note, 1000),
    '登録日': settings.createdAt,
  };
}

function clientRecordToDto_(record, stats) {
  const related = stats || {};
  return {
    id: record.ID,
    row: record._row,
    name: record['クライアント名'] || '',
    commissionRate: normalizeCommissionRate_(record['販売手数料']),
    contactName: record['担当者'] || '',
    email: record['メール'] || '',
    phone: record['電話'] || '',
    note: record['備考'] || '',
    createdAt: record['登録日'] || '',
    projectCount: Number(related.projectCount) || 0,
    activeProjectCount: Number(related.activeProjectCount) || 0,
    completedProjectCount: Number(related.completedProjectCount) || 0,
    salesTotal: Number(related.salesTotal) || 0,
    profitTotal: Number(related.profitTotal) || 0,
    lastProjectAt: related.lastProjectAt || '',
  };
}

function normalizeCommissionRate_(value) {
  if (value === '' || value == null) return '';
  const normalized = Number(value);
  if (!isFinite(normalized) || normalized < 0) return '';
  return Math.round(normalized * 100) / 100;
}
