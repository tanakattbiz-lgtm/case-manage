function apiSuccess_(data) {
  return { ok: true, data: data == null ? null : data };
}

function apiFailure_(error) {
  const normalized = normalizeApiError_(error);
  return {
    ok: false,
    error: {
      code: normalized.code,
      message: normalized.message,
      detail: normalized.detail,
    },
  };
}

function normalizeApiError_(error) {
  const raw = String(error && error.message ? error.message : error || 'APP_ERROR: エラーが発生しました。');
  const separator = raw.indexOf(':');
  if (separator > 0) {
    return {
      code: raw.slice(0, separator).trim() || 'APP_ERROR',
      message: raw.slice(separator + 1).trim() || 'エラーが発生しました。',
      detail: raw,
    };
  }
  return {
    code: 'APP_ERROR',
    message: raw,
    detail: raw,
  };
}

function api_getInitialData(sessionToken) {
  try {
    const sessionContext = requireReadAccess_(sessionToken);
    const dashboard = getDashboard(sessionToken, {});
    return apiSuccess_({
      user: sanitizeUser_(sessionContext.user),
      permissions: getRolePermissions_(sessionContext.user['ロール']),
      master: {
        clients: getClientOptions(sessionToken),
        projectStatuses: PROJECT_STATUS_LIST,
      },
      dashboard: dashboard,
      fetchedAt: nowDateTimeStr_(),
    });
  } catch (error) {
    return apiFailure_(error);
  }
}

function api_searchCases(sessionToken, condition) {
  try {
    const normalized = normalizeCaseSearchCondition_(condition || {});
    const result = getProjects(sessionToken, normalized);
    result.rows = (result.items || []).map(caseListDtoToApiRow_);
    result.page = result.pagination ? result.pagination.page : normalized.page;
    result.pageSize = result.pagination ? result.pagination.pageSize : normalized.pageSize;
    result.total = result.pagination ? result.pagination.totalItems : 0;
    result.hasNext = result.pagination ? result.pagination.hasNext : false;
    return apiSuccess_(result);
  } catch (error) {
    return apiFailure_(error);
  }
}

function api_getCaseDetail(sessionToken, payload) {
  try {
    const input = payload || {};
    return apiSuccess_(getProjectDetail(sessionToken, input.caseId || input.projectId || input.id));
  } catch (error) {
    return apiFailure_(error);
  }
}

function api_saveCase(sessionToken, payload) {
  try {
    const input = payload || {};
    const result = input.id ? updateProject(sessionToken, input) : addProject(sessionToken, input);
    return apiSuccess_(result);
  } catch (error) {
    return apiFailure_(error);
  }
}

function api_deleteCase(sessionToken, payload) {
  try {
    const input = payload || {};
    return apiSuccess_(deleteProject(sessionToken, input.caseId || input.projectId || input.id));
  } catch (error) {
    return apiFailure_(error);
  }
}

function api_getDashboardData(sessionToken, condition) {
  try {
    return apiSuccess_(getDashboard(sessionToken, condition || {}));
  } catch (error) {
    return apiFailure_(error);
  }
}

function api_searchClients(sessionToken, condition) {
  try {
    const input = condition || {};
    const result = getClients(sessionToken, {
      query: input.keyword || input.query || '',
      page: input.page || 1,
      pageSize: input.pageSize || 50,
    });
    result.rows = result.items || [];
    result.page = result.pagination ? result.pagination.page : 1;
    result.pageSize = result.pagination ? result.pagination.pageSize : 50;
    result.total = result.pagination ? result.pagination.totalItems : 0;
    result.hasNext = result.pagination ? result.pagination.hasNext : false;
    return apiSuccess_(result);
  } catch (error) {
    return apiFailure_(error);
  }
}

function api_getClientDetail(sessionToken, payload) {
  try {
    const input = payload || {};
    return apiSuccess_(getClientDetail(sessionToken, input.clientId || input.id));
  } catch (error) {
    return apiFailure_(error);
  }
}

function api_getReportData(sessionToken, condition) {
  try {
    requireReadAccess_(sessionToken);
    const normalizedFilter = normalizeDashboardFilter_(condition || {});
    return apiSuccess_(buildReportData_(listProjectDtos_(), normalizedFilter));
  } catch (error) {
    return apiFailure_(error);
  }
}

function buildReportData_(projects, filter) {
  const dashboard = buildDashboardSummary_(projects || [], filter);
  const statusDistribution = PROJECT_STATUS_LIST.map(function (status) {
    return {
      label: status,
      count: dashboard.statusCount[status] || 0,
      sales: dashboard.statusSales[status] || 0,
    };
  });
  const clientSales = dashboard.clientRanking.map(function (client) {
    return {
      label: client.name,
      sales: client.sales,
      profit: client.profit,
      count: client.count,
    };
  });

  const clientTable = dashboard.clientRanking.map(function (client) {
    return {
      name: client.name,
      count: client.count,
      sales: client.sales,
      profit: client.profit,
      avgDeal: client.count > 0 ? Math.round(client.sales / client.count) : 0,
      margin: client.margin,
    };
  });

  const marginTrend = (dashboard.monthly || []).map(function (m) {
    return {
      month: m.month,
      margin: m.sales > 0 ? Math.round((m.profit / m.sales) * 100) : 0,
      avgDeal: m.completedCount > 0 ? Math.round(m.sales / m.completedCount) : 0,
      completedCount: m.completedCount || 0,
    };
  });

  return {
    filter: filter,
    summary: dashboard.summary,
    monthly: dashboard.monthly,
    statusDistribution: statusDistribution,
    clientSales: clientSales,
    clientTable: clientTable,
    marginTrend: marginTrend,
    funnel: dashboard.funnel || [],
    staleItems: dashboard.staleItems || [],
    insights: dashboard.insights || [],
    ownerPerformance: clientSales,
    overdueTrend: [],
    fetchedAt: nowDateTimeStr_(),
  };
}

function normalizeCaseSearchCondition_(condition) {
  const input = condition || {};
  return {
    query: input.keyword || input.query || '',
    status: input.status || '',
    assignee: input.assignee || '',
    clientId: input.clientId || '',
    dateFrom: input.from || input.dateFrom || '',
    dateTo: input.to || input.dateTo || '',
    page: input.page || 1,
    pageSize: input.pageSize || 50,
    sortKey: input.sortKey || 'updatedAt',
    sortOrder: input.sortOrder || 'desc',
    scope: input.scope || null,
  };
}

function caseListDtoToApiRow_(project) {
  return {
    caseId: project.id,
    caseName: project.name,
    clientName: project.clientName,
    status: project.status,
    assignee: '',
    amount: project.sales,
    grossProfit: project.profit,
    parentProjectId: project.parentProjectId,
    integrationProjectId: project.integrationProjectId,
    phaseName: project.phaseName,
    depositAmount: project.depositAmount,
    dueDate: project.targetDate,
    updatedAt: project.updatedAt,
  };
}
