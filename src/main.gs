// ============================================================
//  案件管理システム - Google Apps Script バックエンド
// ============================================================

const SN = { projects: '案件データ', clients: 'クライアントマスタ' };
const PROJ_COLS = ['ID', '案件名', 'クライアントID', 'クライアント名', '売上', '利益', 'ステータス', '完了日', '備考', '登録日', '更新日'];
const CLI_COLS = ['ID', 'クライアント名', '担当者', 'メール', '電話', '備考', '登録日'];

function getSheet(name, headers) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(name);
    if (!sh) {
        sh = ss.insertSheet(name);
        sh.appendRow(headers);
        sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1f2937').setFontColor('#fff');
        sh.setFrozenRows(1);
    }
    return sh;
}
const projSheet = () => getSheet(SN.projects, PROJ_COLS);
const clientSheet = () => getSheet(SN.clients, CLI_COLS);

function doGet() {
    return HtmlService.createHtmlOutputFromFile('index')
        .setTitle('案件管理')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
        .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function genId(prefix) { return prefix + '-' + Date.now().toString(36).toUpperCase(); }
function nowStr() { return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'); }
function todayStr() { return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd'); }

function sheetToObjects(sh) {
    const data = sh.getDataRange().getValues();
    if (data.length <= 1) return [];
    const headers = data[0];
    return data.slice(1).map((row, i) => {
        const obj = { _row: i + 2 };
        headers.forEach((h, j) => { obj[h] = (row[j] instanceof Date) ? Utilities.formatDate(row[j], 'Asia/Tokyo', 'yyyy/MM/dd') : row[j]; });
        return obj;
    });
}

// ── 案件 ──────────────────────────────────────────────────
function getProjects() {
    const rows = sheetToObjects(projSheet());
    // 手入力の「完了日」を優先してソートし、空の場合は「登録日」でフォールバック
    rows.sort((a, b) => {
        const da = new Date(a['完了日'] || a['登録日'] || 0);
        const db = new Date(b['完了日'] || b['登録日'] || 0);
        return db - da;
    });
    return rows;
}

function addProject(p) {
    const sales = Number(p['売上']) || 0;
    const profit = (p['利益'] !== '' && p['利益'] != null) ? Number(p['利益']) : sales;
    const id = genId('PRJ');
    const n = nowStr();
    const cd = p['完了日'] || '';
    projSheet().appendRow([id, p['案件名'] || '', p['クライアントID'] || '', p['クライアント名'] || '', sales, profit, p['ステータス'] || '商談中', cd, p['備考'] || '', n, n]);
    return { success: true, id };
}

function updateProject(p) {
    const sh = projSheet();
    const row = Number(p['_row']);
    if (!row) return { success: false };
    const sales = Number(p['売上']) || 0;
    const profit = (p['利益'] !== '' && p['利益'] != null) ? Number(p['利益']) : sales;
    let cd = p['完了日'] || '';
    // ステータスが完了で日付が未入力の場合のみ、今日の日付を補完する
    if (p['ステータス'] === '完了' && !cd) cd = todayStr();

    sh.getRange(row, 1, 1, PROJ_COLS.length).setValues([[
        p['ID'], p['案件名'] || '', p['クライアントID'] || '', p['クライアント名'] || '',
        sales, profit, p['ステータス'] || '商談中', cd, p['備考'] || '', p['登録日'] || nowStr(), nowStr()
    ]]);
    return { success: true };
}

function deleteProject(row) { projSheet().deleteRow(Number(row)); return { success: true }; }

// ── クライアント ───────────────────────────────────────────
function getClients() { return sheetToObjects(clientSheet()); }

function addClient(c) {
    const id = genId('CLI');
    clientSheet().appendRow([id, c['クライアント名'] || '', c['担当者'] || '', c['メール'] || '', c['電話'] || '', c['備考'] || '', nowStr()]);
    return { success: true, id };
}

function updateClient(c) {
    const row = Number(c['_row']);
    if (!row) return { success: false };
    clientSheet().getRange(row, 1, 1, CLI_COLS.length).setValues([[
        c['ID'], c['クライアント名'] || '', c['担当者'] || '', c['メール'] || '', c['電話'] || '', c['備考'] || '', c['登録日'] || nowStr()
    ]]);
    return { success: true };
}

function deleteClient(row) { clientSheet().deleteRow(Number(row)); return { success: true }; }

// ── ダッシュボード集計 ─────────────────────────────────────
function getDashboard() {
    const projects = sheetToObjects(projSheet());
    let totalSales = 0, totalProfit = 0, activeSales = 0;
    const statusCount = {};
    const monthlySales = {}, monthlyProfit = {};
    const clientMap = {};

    projects.forEach(p => {
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

    const allKeys = [...new Set([...Object.keys(monthlySales), ...Object.keys(monthlyProfit)])].sort().slice(-12);
    const monthly = allKeys.map(k => ({
        month: k,
        sales: monthlySales[k] || 0,
        profit: monthlyProfit[k] || 0,
        margin: monthlySales[k] ? Math.round((monthlyProfit[k] || 0) / monthlySales[k] * 100) : 0,
    }));

    const clientRanking = Object.entries(clientMap)
        .map(([name, v]) => ({ name, ...v, margin: v.sales > 0 ? Math.round(v.profit / v.sales * 100) : 0 }))
        .sort((a, b) => b.sales - a.sales).slice(0, 8);

    // ダッシュボードの「最近の案件」も手入力の日付順に
    const recent = [...projects].sort((a, b) => {
        const da = new Date(a['完了日'] || a['登録日'] || 0);
        const db = new Date(b['完了日'] || b['登録日'] || 0);
        return db - da;
    }).slice(0, 5);

    return {
        kpi: { totalSales, totalProfit, activeSales, avgMargin: totalSales > 0 ? Math.round(totalProfit / totalSales * 100) : 0, totalCount: projects.length, completedCount: statusCount['完了'] || 0 },
        statusCount, monthly, clientRanking, recent,
    };
}