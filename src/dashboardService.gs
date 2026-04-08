function getDashboard(sessionToken) {
  requireReadAccess_(sessionToken);
  const projects = sheetToObjects(projSheet());
  let totalSales = 0;
  let totalProfit = 0;
  let activeSales = 0;
  const statusCount = {};
  const monthlySales = {};
  const monthlyProfit = {};
  const clientMap = {};

  projects.forEach((p) => {
    const s = Number(p['売上']) || 0;
    const pr = Number(p['利益']) || 0;
    const st = p['ステータス'] || '不明';

    statusCount[st] = (statusCount[st] || 0) + 1;

    // 完了案件のみを集計対象とする
    if (st === '完了') {
      totalSales += s;
      totalProfit += pr;

      // 手入力の「完了日」を優先して月次集計
      const ds = p['完了日'] || p['登録日'];
      if (ds) {
        const d = new Date(ds);
        if (!isNaN(d)) {
          const k = `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}`;
          monthlySales[k] = (monthlySales[k] || 0) + s;
          monthlyProfit[k] = (monthlyProfit[k] || 0) + pr;
        }
      }

      const cn = p['クライアント名'] || '未設定';
      if (!clientMap[cn]) clientMap[cn] = { sales: 0, profit: 0, count: 0 };
      clientMap[cn].sales += s;
      clientMap[cn].profit += pr;
      clientMap[cn].count += 1;
    } else if (st === '進行中') {
      activeSales += s;
    }
  });

  const allKeys = [...new Set([...Object.keys(monthlySales), ...Object.keys(monthlyProfit)])]
    .sort()
    .slice(-12);

  const monthly = allKeys.map((k) => ({
    month: k,
    sales: monthlySales[k] || 0,
    profit: monthlyProfit[k] || 0,
    margin: monthlySales[k] ? Math.round(((monthlyProfit[k] || 0) / monthlySales[k]) * 100) : 0,
  }));

  const clientRanking = Object.entries(clientMap)
    .map(([name, v]) => ({
      name,
      ...v,
      margin: v.sales > 0 ? Math.round((v.profit / v.sales) * 100) : 0,
    }))
    .sort((a, b) => b.sales - a.sales)
    .slice(0, 8);

  // ダッシュボードの「最近の案件」も手入力の日付順に
  const recent = [...projects]
    .sort((a, b) => {
      const da = new Date(a['完了日'] || a['登録日'] || 0);
      const db = new Date(b['完了日'] || b['登録日'] || 0);
      return db - da;
    })
    .slice(0, 5);

  return {
    kpi: {
      totalSales,
      totalProfit,
      activeSales,
      avgMargin: totalSales > 0 ? Math.round((totalProfit / totalSales) * 100) : 0,
      totalCount: projects.length,
      completedCount: statusCount['完了'] || 0,
    },
    statusCount,
    monthly,
    clientRanking,
    recent,
  };
}
