function getProjects(sessionToken, query) {
  requireReadAccess_(sessionToken);
  const normalizedQuery = normalizeProjectListQuery_(query);
  const allItems = listProjectDtos_();
  const filteredItems = filterProjectDtos_(allItems, normalizedQuery);
  const paged = paginateItems_(filteredItems, normalizedQuery.page, normalizedQuery.pageSize);

  return {
    items: paged.items,
    pagination: paged.pagination,
    filters: normalizedQuery,
    summary: buildProjectListSummary_(filteredItems, allItems.length),
    fetchedAt: nowDateTimeStr_(),
  };
}

function getProjectDetail(sessionToken, projectId) {
  requireReadAccess_(sessionToken);
  const targetId = normalizeString_(projectId);
  const project = listProjectDtos_().find(function (item) {
    return item.id === targetId;
  }) || null;

  if (!project) {
    throwAppError_('PROJECT_NOT_FOUND', '案件が見つかりません。');
  }

  const clientRecord = project.clientId ? findClientRecordById_(project.clientId) : null;
  const relatedProjects = listProjectDtos_()
    .filter(function (item) {
      return item.id !== project.id && item.clientId && item.clientId === project.clientId;
    })
    .slice(0, FIXED_VALUES.lists.detailRelatedLimit);

  return {
    project: project,
    client: clientRecord ? clientRecordToDto_(clientRecord) : null,
    relatedProjects: relatedProjects,
    history: buildProjectHistory_(project),
  };
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

function listProjectDtos_() {
  return listProjectRecords_()
    .sort(function (a, b) {
      const left = parseDateValue_(a['完了日'] || a['登録日']);
      const right = parseDateValue_(b['完了日'] || b['登録日']);
      return (right ? right.getTime() : 0) - (left ? left.getTime() : 0);
    })
    .map(projectRecordToDto_);
}

function normalizeProjectListQuery_(query) {
  const input = query || {};
  return {
    page: normalizePositiveInteger_(
      input.page,
      1,
      1,
      999
    ),
    pageSize: normalizePositiveInteger_(
      input.pageSize,
      FIXED_VALUES.lists.defaultPageSize,
      1,
      FIXED_VALUES.lists.maxPageSize
    ),
    query: normalizeString_(input.query),
    status: normalizeOptionalProjectStatus_(input.status),
    clientId: normalizeString_(input.clientId),
    dateFrom: normalizeDateInput_(input.dateFrom),
    dateTo: normalizeDateInput_(input.dateTo),
    scope: normalizeOptionalProjectScope_(input.scope),
    sort: {
      field: 'targetDate',
      direction: 'desc',
      label: '対象日が新しい順',
    },
  };
}

function normalizeOptionalProjectScope_(scope) {
  const input = scope || {};
  const hasValue = Boolean(
    normalizeString_(input.mode)
    || normalizeString_(input.year)
    || normalizeString_(input.month)
    || normalizeString_(input.status)
    || normalizeString_(input.label)
  );
  if (!hasValue) return null;

  const resolvedMode = ['all', 'year', 'month'].indexOf(String(input.mode || '')) >= 0
    ? String(input.mode)
    : (input.month ? 'month' : (input.year ? 'year' : 'all'));

  return {
    mode: resolvedMode || 'all',
    year: normalizeString_(input.year),
    month: normalizeString_(input.month).padStart(2, '0'),
    status: normalizeOptionalProjectStatus_(input.status),
    label: normalizeString_(input.label),
  };
}

function filterProjectDtos_(items, query) {
  const from = parseDateValue_(query.dateFrom);
  const to = parseDateValue_(query.dateTo);
  const normalizedQuery = String(query.query || '').toLowerCase();

  return items.filter(function (project) {
    if (!isProjectInScope_(project, query.scope)) return false;
    if (query.status && project.status !== query.status) return false;
    if (query.clientId && project.clientId !== query.clientId) return false;

    if (normalizedQuery) {
      const joined = [
        project.name,
        project.clientName,
        project.note,
        project.status,
      ].join(' ').toLowerCase();
      if (joined.indexOf(normalizedQuery) === -1) return false;
    }

    const targetDate = parseDateValue_(project.targetDate);
    if (from && (!targetDate || targetDate.getTime() < from.getTime())) return false;
    if (to && (!targetDate || targetDate.getTime() > to.getTime())) return false;

    return true;
  });
}

function isProjectInScope_(project, scope) {
  if (!scope) return true;
  if (scope.status && project.status !== scope.status) return false;
  if (scope.mode === 'all') return true;

  const baseDate = parseDateValue_(project.targetDate);
  if (!baseDate) return false;

  const year = String(baseDate.getFullYear());
  const month = String(baseDate.getMonth() + 1).padStart(2, '0');
  if (scope.mode === 'year') return year === scope.year;
  return year === scope.year && month === scope.month;
}

function buildProjectListSummary_(filteredItems, totalItems) {
  const statusCount = {};
  PROJECT_STATUS_LIST.forEach(function (status) {
    statusCount[status] = 0;
  });

  const sales = filteredItems.reduce(function (sum, project) {
    const status = normalizeProjectStatus_(project.status);
    statusCount[status] = (statusCount[status] || 0) + 1;
    return sum + (Number(project.sales) || 0);
  }, 0);

  const profit = filteredItems.reduce(function (sum, project) {
    return sum + (Number(project.profit) || 0);
  }, 0);

  return {
    totalItems: Number(totalItems) || 0,
    filteredItems: filteredItems.length,
    sales: sales,
    profit: profit,
    statusCount: statusCount,
  };
}

function buildProjectHistory_(project) {
  const history = [];

  if (project.createdAt) {
    history.push({
      key: 'created',
      label: '登録',
      date: project.createdAt,
      description: '案件が登録されました。',
    });
  }
  if (project.updatedAt && project.updatedAt !== project.createdAt) {
    history.push({
      key: 'updated',
      label: '更新',
      date: project.updatedAt,
      description: '案件情報が更新されました。',
    });
  }
  if (project.completedAt) {
    history.push({
      key: 'completed',
      label: '完了',
      date: project.completedAt,
      description: '完了日が設定されています。',
    });
  }

  return history.sort(function (a, b) {
    const left = parseDateValue_(a.date);
    const right = parseDateValue_(b.date);
    return (right ? right.getTime() : 0) - (left ? left.getTime() : 0);
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
  const sales = Number(record['売上']) || 0;
  const profit = Number(record['利益']) || 0;

  return {
    id: record.ID,
    row: record._row,
    name: record['案件名'] || '',
    clientId: record['クライアントID'] || '',
    clientName: record['クライアント名'] || '',
    sales: sales,
    profit: profit,
    marginRate: sales > 0 ? Math.round((profit / sales) * 100) : 0,
    status: normalizeProjectStatus_(record['ステータス']),
    completedAt: record['完了日'] || '',
    targetDate: record['完了日'] || record['登録日'] || '',
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

function normalizeOptionalProjectStatus_(status) {
  const normalized = normalizeString_(status);
  return PROJECT_STATUS_LIST.indexOf(normalized) >= 0 ? normalized : '';
}
