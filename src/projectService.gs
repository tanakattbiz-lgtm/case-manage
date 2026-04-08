function getProjects(sessionToken) {
  requireReadAccess_(sessionToken);
  const rows = sheetToObjects(projSheet());

  // 手入力の「完了日」を優先してソートし、空の場合は「登録日」でフォールバック
  rows.sort((a, b) => {
    const da = new Date(a['完了日'] || a['登録日'] || 0);
    const db = new Date(b['完了日'] || b['登録日'] || 0);
    return db - da;
  });

  return rows;
}

function addProject(sessionToken, p) {
  requireEditAccess_(sessionToken);
  const sales = Number(p['売上']) || 0;
  const profit = (p['利益'] !== '' && p['利益'] != null) ? Number(p['利益']) : sales;
  const id = genId('PRJ');
  const n = nowStr();
  const cd = p['完了日'] || '';

  projSheet().appendRow([
    id,
    p['案件名'] || '',
    p['クライアントID'] || '',
    p['クライアント名'] || '',
    sales,
    profit,
    p['ステータス'] || '商談中',
    cd,
    p['備考'] || '',
    n,
    n,
  ]);

  return { success: true, id };
}

function updateProject(sessionToken, p) {
  requireEditAccess_(sessionToken);
  const sh = projSheet();
  const row = Number(p['_row']);
  if (!row) return { success: false };

  const sales = Number(p['売上']) || 0;
  const profit = (p['利益'] !== '' && p['利益'] != null) ? Number(p['利益']) : sales;
  let cd = p['完了日'] || '';

  // ステータスが完了で日付が未入力の場合のみ、今日の日付を補完する
  if (p['ステータス'] === '完了' && !cd) cd = todayStr();

  sh.getRange(row, 1, 1, PROJ_COLS.length).setValues([[
    p['ID'],
    p['案件名'] || '',
    p['クライアントID'] || '',
    p['クライアント名'] || '',
    sales,
    profit,
    p['ステータス'] || '商談中',
    cd,
    p['備考'] || '',
    p['登録日'] || nowStr(),
    nowStr(),
  ]]);

  return { success: true };
}

function deleteProject(sessionToken, row) {
  requireEditAccess_(sessionToken);
  projSheet().deleteRow(Number(row));
  return { success: true };
}
