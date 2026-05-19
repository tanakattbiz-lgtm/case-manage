function getDashboard(sessionToken, filter) {
  requireReadAccess_(sessionToken);
  const normalizedFilter = normalizeDashboardFilter_(filter);
  const cache = CacheService.getScriptCache();
  const cacheKey = getDashboardCacheKey_(normalizedFilter);
  const cached = cache.get(cacheKey);
  if (cached) {
    return JSON.parse(cached);
  }

  const projects = listProjectDtos_();
  const dashboard = buildDashboardSummary_(projects, normalizedFilter);
  cache.put(cacheKey, JSON.stringify(dashboard), FIXED_VALUES.dashboard.cacheSeconds);
  return dashboard;
}

function normalizeDashboardFilter_(filter) {
  const input = filter || {};
  const now = new Date();
  const mode = ['all', 'year', 'month'].indexOf(String(input.mode || 'month')) >= 0
    ? String(input.mode || 'month')
    : 'month';
  const year = String(input.year || now.getFullYear());
  const month = String(input.month || Utilities.formatDate(now, APP_INFO.timezone, 'MM')).padStart(2, '0');

  if (mode === 'all') return { mode: 'all', year: '', month: '', label: '全期間' };
  if (mode === 'year') return { mode: 'year', year: year, month: '', label: year + '年' };
  return { mode: 'month', year: year, month: month, label: year + '年' + Number(month) + '月' };
}

function buildDashboardSummary_(projects, filter) {
  const filtered = projects.filter(function (project) {
    return isProjectInDashboardPeriod_(project, filter);
  });
  const revenueProjects = filterRevenueProjects_(projects);
  const filteredRevenue = revenueProjects.filter(function (project) {
    return isProjectInDashboardPeriod_(project, filter);
  });
  const statusCount = {};
  const statusSales = {};
  PROJECT_STATUS_LIST.forEach(function (status) {
    statusCount[status] = 0;
    statusSales[status] = 0;
  });

  const monthlyMap = {};
  const clientMap = {};
  const leadTimes = [];

  filtered.forEach(function (project) {
    const status = normalizeProjectStatus_(project.status);
    statusCount[status] = (statusCount[status] || 0) + 1;
  });

  filteredRevenue.forEach(function (project) {
    const sales = Number(project.sales) || 0;
    const profit = Number(project.profit) || 0;
    const status = normalizeProjectStatus_(project.status);
    statusSales[status] = (statusSales[status] || 0) + sales;

    const bucketKey = getMonthlyBucketKey_(project);
    if (!monthlyMap[bucketKey]) {
      monthlyMap[bucketKey] = {
        month: bucketKey,
        sales: 0,
        profit: 0,
        proposalSales: 0,
        activeSales: 0,
        completedCount: 0,
        activeCount: 0,
        proposalCount: 0,
      };
    }

    if (status === PROJECT_STATUSES.completed) {
      monthlyMap[bucketKey].sales += sales;
      monthlyMap[bucketKey].profit += profit;
      monthlyMap[bucketKey].completedCount += 1;

      const clientName = project.clientName || '未設定';
      if (!clientMap[clientName]) {
        clientMap[clientName] = { name: clientName, sales: 0, profit: 0, count: 0 };
      }
      clientMap[clientName].sales += sales;
      clientMap[clientName].profit += profit;
      clientMap[clientName].count += 1;

      const createdAt = parseDateValue_(project.createdAt);
      const completedAt = parseDateValue_(project.completedAt);
      if (createdAt && completedAt) {
        const diffDays = Math.ceil((completedAt.getTime() - createdAt.getTime()) / (1000 * 60 * 60 * 24));
        if (diffDays >= 0) leadTimes.push(diffDays);
      }
    }
    if (status === PROJECT_STATUSES.active) {
      monthlyMap[bucketKey].activeSales += sales;
      monthlyMap[bucketKey].activeCount += 1;
    }
    if (status === PROJECT_STATUSES.lead) {
      monthlyMap[bucketKey].proposalSales += sales;
      monthlyMap[bucketKey].proposalCount += 1;
    }
  });

  const summary = buildDashboardMetrics_(filtered, filteredRevenue, statusCount, leadTimes);
  const clientRanking = buildClientRanking_(clientMap);

  return {
    filter: filter,
    summary: summary,
    statusCount: statusCount,
    statusSales: statusSales,
    monthly: buildMonthlySeries_(monthlyMap),
    clientRanking: clientRanking,
    recent: buildRecentProjects_(filteredRevenue),
    insights: buildDashboardInsights_(summary, clientRanking, statusCount),
    staleItems: buildStaleItems_(filtered),
    funnel: buildFunnelData_(statusCount, statusSales),
    comparison: buildPeriodComparison_(revenueProjects, filter),
  };
}

function buildMonthlySeries_(monthlyMap) {
  const now = new Date();
  const items = [];
  for (let index = FIXED_VALUES.dashboard.monthWindow - 1; index >= 0; index -= 1) {
    const date = new Date(now.getFullYear(), now.getMonth() - index, 1);
    const key = date.getFullYear() + '/' + String(date.getMonth() + 1).padStart(2, '0');
    items.push(monthlyMap[key] || {
      month: key,
      sales: 0,
      profit: 0,
      proposalSales: 0,
      activeSales: 0,
      completedCount: 0,
      activeCount: 0,
      proposalCount: 0,
    });
  }
  return items;
}

function buildDashboardMetrics_(projects, revenueProjects, statusCount, leadTimes) {
  const completedProjects = revenueProjects.filter(function (project) {
    return project.status === PROJECT_STATUSES.completed;
  });
  const totalSales = completedProjects.reduce(function (sum, project) { return sum + (Number(project.sales) || 0); }, 0);
  const totalProfit = completedProjects.reduce(function (sum, project) { return sum + (Number(project.profit) || 0); }, 0);
  const activeSales = revenueProjects
    .filter(function (project) { return project.status === PROJECT_STATUSES.active; })
    .reduce(function (sum, project) { return sum + (Number(project.sales) || 0); }, 0);
  const proposalSales = revenueProjects
    .filter(function (project) { return project.status === PROJECT_STATUSES.lead; })
    .reduce(function (sum, project) { return sum + (Number(project.sales) || 0); }, 0);
  const forecastSales = totalSales + activeSales + proposalSales;
  const forecastProfit = revenueProjects
    .filter(function (project) {
      return project.status === PROJECT_STATUSES.completed
        || project.status === PROJECT_STATUSES.active
        || project.status === PROJECT_STATUSES.lead;
    })
    .reduce(function (sum, project) { return sum + (Number(project.profit) || 0); }, 0);
  const completedCount = completedProjects.length;
  const actionCount = (statusCount[PROJECT_STATUSES.lead] || 0) + (statusCount[PROJECT_STATUSES.active] || 0);

  return {
    totalCount: projects.length,
    totalSales: totalSales,
    totalProfit: totalProfit,
    activeSales: activeSales,
    proposalSales: proposalSales,
    pipelineSales: activeSales + proposalSales,
    forecastSales: forecastSales,
    forecastMargin: forecastSales > 0 ? Math.round((forecastProfit / forecastSales) * 100) : 0,
    avgMargin: totalSales > 0 ? Math.round((totalProfit / totalSales) * 100) : 0,
    avgDealSize: completedCount > 0 ? Math.round(totalSales / completedCount) : 0,
    avgLeadDays: leadTimes.length > 0 ? Math.round(leadTimes.reduce(function (sum, value) { return sum + value; }, 0) / leadTimes.length) : 0,
    completionRate: projects.length > 0 ? Math.round((completedCount / projects.length) * 100) : 0,
    actionRate: projects.length > 0 ? Math.round((actionCount / projects.length) * 100) : 0,
    completedCount: completedCount,
    activeCount: statusCount[PROJECT_STATUSES.active] || 0,
    proposalCount: statusCount[PROJECT_STATUSES.lead] || 0,
    holdCount: (statusCount[PROJECT_STATUSES.pending] || 0) + (statusCount[PROJECT_STATUSES.stopped] || 0),
  };
}

function buildClientRanking_(clientMap) {
  return Object.keys(clientMap)
    .map(function (name) {
      const row = clientMap[name];
      return {
        name: row.name,
        sales: row.sales,
        profit: row.profit,
        count: row.count,
        margin: row.sales > 0 ? Math.round((row.profit / row.sales) * 100) : 0,
      };
    })
    .sort(function (a, b) {
      return b.sales - a.sales;
    })
    .slice(0, FIXED_VALUES.dashboard.topClientsLimit);
}

function buildRecentProjects_(projects) {
  return projects
    .filter(function (project) {
      return project.status === PROJECT_STATUSES.completed;
    })
    .sort(function (a, b) {
      const left = parseDateValue_(a.completedAt || a.createdAt);
      const right = parseDateValue_(b.completedAt || b.createdAt);
      return (right ? right.getTime() : 0) - (left ? left.getTime() : 0);
    })
    .slice(0, FIXED_VALUES.dashboard.recentLimit);
}

function buildDashboardInsights_(summary, clientRanking, statusCount) {
  const items = [];

  if (summary.pipelineSales > 0) {
    items.push({ type: 'info', text: 'パイプライン合計 ' + formatCurrencyShort_(summary.pipelineSales) + ' が進行中・商談中にあります。' });
  }

  const leadCount = statusCount[PROJECT_STATUSES.lead] || 0;
  const activeCount = statusCount[PROJECT_STATUSES.active] || 0;
  if (leadCount > 0 && activeCount === 0) {
    items.push({ type: 'warning', text: '商談中案件が ' + leadCount + ' 件あります。進行中に移行できていません。' });
  }

  if (clientRanking[0]) {
    const topShare = summary.totalSales > 0 ? Math.round((clientRanking[0].sales / summary.totalSales) * 100) : 0;
    if (topShare > 50) {
      items.push({ type: 'warning', text: '「' + clientRanking[0].name + '」への売上集中度が ' + topShare + '% と高い状態です。' });
    } else {
      items.push({ type: 'success', text: '最大クライアント「' + clientRanking[0].name + '」（' + formatCurrencyShort_(clientRanking[0].sales) + '）。利益率 ' + clientRanking[0].margin + '%。' });
    }
  }

  if (summary.avgLeadDays > 0) {
    const comment = summary.avgLeadDays > 60 ? '短縮余地があります。' : '標準的な商談期間です。';
    items.push({ type: summary.avgLeadDays > 60 ? 'warning' : 'info', text: '平均リードタイムは ' + summary.avgLeadDays + ' 日です。' + comment });
  }

  if (summary.avgMargin > 0) {
    const comment = summary.avgMargin >= 30 ? '良好な利益率です。' : summary.avgMargin < 15 ? '利益率の改善余地があります。' : '安定した利益率です。';
    items.push({ type: summary.avgMargin >= 30 ? 'success' : (summary.avgMargin < 15 ? 'warning' : 'info'), text: '確定利益率は ' + summary.avgMargin + '%。' + comment });
  }

  const holdCount = (statusCount[PROJECT_STATUSES.pending] || 0) + (statusCount[PROJECT_STATUSES.stopped] || 0);
  if (holdCount > 0) {
    items.push({ type: 'warning', text: '保留・停止中の案件が ' + holdCount + ' 件あります。棚卸しを検討してください。' });
  }

  if (summary.avgDealSize > 0) {
    items.push({ type: 'info', text: '平均商談規模は ' + formatCurrencyShort_(summary.avgDealSize) + ' です。' });
  }

  if (items.length === 0) {
    items.push({ type: 'info', text: '案件データが増えると月次推移とクライアント比率が自動で可視化されます。' });
  }
  return items;
}

function formatCurrencyShort_(value) {
  const num = Math.round(Number(value) || 0);
  if (num >= 100000000) return '約 ' + (Math.round(num / 10000000) / 10) + ' 億円';
  if (num >= 10000) return '約 ' + Math.round(num / 10000) + ' 万円';
  return '¥' + String(num).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

function buildStaleItems_(projects) {
  const now = new Date();
  const STALE_DAYS = 21;
  return projects
    .filter(function (project) {
      if (project.status !== PROJECT_STATUSES.active && project.status !== PROJECT_STATUSES.lead) return false;
      const updated = parseDateValue_(project.updatedAt || project.createdAt);
      if (!updated) return false;
      return Math.floor((now.getTime() - updated.getTime()) / 86400000) >= STALE_DAYS;
    })
    .sort(function (a, b) {
      const la = parseDateValue_(a.updatedAt || a.createdAt);
      const lb = parseDateValue_(b.updatedAt || b.createdAt);
      return (la ? la.getTime() : 0) - (lb ? lb.getTime() : 0);
    })
    .slice(0, 6)
    .map(function (project) {
      const updated = parseDateValue_(project.updatedAt || project.createdAt);
      const daysSince = updated ? Math.floor((now.getTime() - updated.getTime()) / 86400000) : 0;
      return { id: project.id, name: project.name, clientName: project.clientName, status: project.status, sales: project.sales, daysSince: daysSince };
    });
}

function buildFunnelData_(statusCount, statusSales) {
  return [
    { label: '商談中', count: statusCount[PROJECT_STATUSES.lead] || 0, sales: statusSales[PROJECT_STATUSES.lead] || 0 },
    { label: '進行中', count: statusCount[PROJECT_STATUSES.active] || 0, sales: statusSales[PROJECT_STATUSES.active] || 0 },
    { label: '完了', count: statusCount[PROJECT_STATUSES.completed] || 0, sales: statusSales[PROJECT_STATUSES.completed] || 0 },
  ];
}

function buildPeriodComparison_(allProjects, filter) {
  if (filter.mode === 'all') return null;
  let prevFilter;
  if (filter.mode === 'month') {
    const d = new Date(Number(filter.year), Number(filter.month) - 2, 1);
    prevFilter = { mode: 'month', year: String(d.getFullYear()), month: String(d.getMonth() + 1).padStart(2, '0') };
  } else {
    prevFilter = { mode: 'year', year: String(Number(filter.year) - 1), month: '' };
  }
  const prev = allProjects.filter(function (p) { return isProjectInDashboardPeriod_(p, prevFilter); });
  const prevCompleted = prev.filter(function (p) { return p.status === PROJECT_STATUSES.completed; });
  const prevSales = prevCompleted.reduce(function (s, p) { return s + (Number(p.sales) || 0); }, 0);
  const prevProfit = prevCompleted.reduce(function (s, p) { return s + (Number(p.profit) || 0); }, 0);
  const prevActive = prev.filter(function (p) { return p.status === PROJECT_STATUSES.active; }).reduce(function (s, p) { return s + (Number(p.sales) || 0); }, 0);
  const prevLead = prev.filter(function (p) { return p.status === PROJECT_STATUSES.lead; }).reduce(function (s, p) { return s + (Number(p.sales) || 0); }, 0);
  return {
    totalSales: prevSales,
    totalProfit: prevProfit,
    avgMargin: prevSales > 0 ? Math.round((prevProfit / prevSales) * 100) : 0,
    pipelineSales: prevActive + prevLead,
  };
}

function isProjectInDashboardPeriod_(project, filter) {
  if (filter.mode === 'all') return true;
  const baseDate = parseDateValue_(project.completedAt || project.createdAt);
  if (!baseDate) return false;
  const year = String(baseDate.getFullYear());
  const month = String(baseDate.getMonth() + 1).padStart(2, '0');
  if (filter.mode === 'year') return year === filter.year;
  return year === filter.year && month === filter.month;
}

function getMonthlyBucketKey_(project) {
  const baseDate = parseDateValue_(project.completedAt || project.createdAt) || new Date();
  return baseDate.getFullYear() + '/' + String(baseDate.getMonth() + 1).padStart(2, '0');
}
