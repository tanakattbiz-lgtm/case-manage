const SN = {
  projects: '案件データ',
  clients: 'クライアントマスタ',
  users: 'ユーザー管理',
  sessions: '認証セッション',
  auditLogs: '認証監査ログ',
};

const PROJ_COLS = [
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
];

const CLI_COLS = [
  'ID',
  'クライアント名',
  '担当者',
  'メール',
  '電話',
  '備考',
  '登録日',
];

const USER_COLS = [
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
];

const SESSION_COLS = [
  'ID',
  'ユーザーID',
  'トークンハッシュ',
  '状態',
  '有効期限',
  '最終アクセス',
  'ユーザーエージェント',
  '登録日',
  '更新日',
];

const AUDIT_COLS = [
  'ID',
  '日時',
  'イベント',
  '結果',
  'ユーザーID',
  '氏名',
  'メール',
  '詳細',
  'ユーザーエージェント',
];

const AUTH_ROLES = {
  admin: 'admin',
  editor: 'editor',
  viewer: 'viewer',
};

const AUTH_ROLE_LABELS = {
  admin: '管理者',
  editor: '編集者',
  viewer: '閲覧者',
};

const AUTH_STATUSES = {
  active: 'active',
  disabled: 'disabled',
};

const AUTH_SETTINGS = {
  sessionTtlMinutes: 8 * 60,
  sessionRefreshThresholdMinutes: 90,
  maxFailedAttempts: 5,
  lockoutMinutes: 15,
  passwordMinLength: 10,
  auditLogLimit: 120,
};
