const APP_INFO = Object.freeze({
  title: '案件管理',
  timezone: 'Asia/Tokyo',
});

const SCRIPT_PROPERTY_KEYS = Object.freeze({
  spreadsheetId: 'SPREADSHEET_ID',
  authSecret: 'AUTH_SECRET',
  authPepper: 'AUTH_PEPPER',
  dashboardCacheVersion: 'DASHBOARD_CACHE_VERSION',
});

const SHEET_NAMES = Object.freeze({
  projects: '案件データ',
  clients: 'クライアントマスタ',
  users: 'ユーザー管理',
  sessions: '認証セッション',
  auditLogs: '認証監査ログ',
});

const PROJECT_COLUMNS = Object.freeze([
  'ID',
  '案件名',
  'クライアントID',
  'クライアント名',
  '売上',
  '利益',
  'ステータス',
  '完了日',
  '備考',
  '登録日',
  '更新日',
]);

const CLIENT_COLUMNS = Object.freeze([
  'ID',
  'クライアント名',
  '既定利益率',
  '担当者',
  'メール',
  '電話',
  '備考',
  '登録日',
]);

const USER_COLUMNS = Object.freeze([
  'ID',
  '氏名',
  'メール',
  'ロール',
  'パスワードハッシュ',
  'ソルト',
  'ステータス',
  '失敗回数',
  'ロック期限',
  '最終ログイン',
  '最終PW変更',
  '登録日',
  '更新日',
]);

const SESSION_COLUMNS = Object.freeze([
  'ID',
  'ユーザーID',
  'トークンハッシュ',
  '状態',
  '有効期限',
  '最終アクセス',
  'ユーザーエージェント',
  '登録日',
  '更新日',
]);

const AUDIT_COLUMNS = Object.freeze([
  'ID',
  '日時',
  'イベント',
  '結果',
  'ユーザーID',
  '氏名',
  'メール',
  '詳細',
  'ユーザーエージェント',
]);

const SHEET_DEFINITIONS = Object.freeze({
  [SHEET_NAMES.projects]: { headers: PROJECT_COLUMNS, hidden: false },
  [SHEET_NAMES.clients]: { headers: CLIENT_COLUMNS, hidden: false },
  [SHEET_NAMES.users]: { headers: USER_COLUMNS, hidden: true },
  [SHEET_NAMES.sessions]: { headers: SESSION_COLUMNS, hidden: true },
  [SHEET_NAMES.auditLogs]: { headers: AUDIT_COLUMNS, hidden: true },
});

const USER_ROLES = Object.freeze({
  admin: 'admin',
  editor: 'editor',
  viewer: 'viewer',
});

const USER_ROLE_LABELS = Object.freeze({
  admin: '管理者',
  editor: '編集者',
  viewer: '閲覧者',
});

const USER_STATUSES = Object.freeze({
  active: 'active',
  disabled: 'disabled',
});

const PROJECT_STATUSES = Object.freeze({
  lead: '商談中',
  active: '進行中',
  completed: '完了',
  stopped: '停止',
  pending: '保留',
});

const PROJECT_STATUS_LIST = Object.freeze([
  PROJECT_STATUSES.lead,
  PROJECT_STATUSES.active,
  PROJECT_STATUSES.completed,
  PROJECT_STATUSES.stopped,
  PROJECT_STATUSES.pending,
]);

const FIXED_VALUES = Object.freeze({
  auth: {
    sessionTtlMinutes: 8 * 60,
    sessionRefreshThresholdMinutes: 90,
    maxFailedAttempts: 5,
    lockoutMinutes: 15,
    passwordMinLength: 10,
    auditLogLimit: 120,
  },
  dashboard: {
    recentLimit: 5,
    topClientsLimit: 8,
    monthWindow: 12,
    cacheSeconds: 300,
  },
  cacheKeys: {
    dashboardAll: 'dashboard:all',
    dashboardYear: 'dashboard:year',
    dashboardMonth: 'dashboard:month',
  },
});
