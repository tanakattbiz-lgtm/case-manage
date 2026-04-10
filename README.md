# case-manage

Google Apps Script + HtmlService で動作する SPA 型の案件管理アプリです。  
このリファクタ版では、`index.html` への集中実装をやめて、`include()` ベースの partial 構成と `Repository / Service / API / state / ui` 分離に切り替えています。

## 変更後の構成

```text
src/
├─ app.gs
├─ constants.gs
├─ utils.gs
├─ sheets.gs
├─ projectRepository.gs
├─ clientRepository.gs
├─ userRepository.gs
├─ sessionRepository.gs
├─ auditRepository.gs
├─ authService.gs
├─ projectService.gs
├─ clientService.gs
├─ userService.gs
├─ dashboardService.gs
├─ index.html
├─ IndexShell.html
├─ Styles.html
├─ StylesCommon.html
├─ StylesLayout.html
├─ StylesAuth.html
├─ StylesDashboard.html
├─ StylesProjects.html
├─ StylesClients.html
├─ StylesUsers.html
├─ Layout.html
├─ Auth.html
├─ Dashboard.html
├─ Projects.html
├─ Clients.html
├─ Users.html
├─ AppScript.html
├─ StateScript.html
├─ ApiScript.html
├─ UiScript.html
├─ AuthScript.html
├─ ChartScript.html
├─ DashboardScript.html
├─ ProjectsScript.html
├─ ClientsScript.html
└─ UsersScript.html
```

## 責務

### サーバー側

- `app.gs`
  - `doGet()` と `include()` のみ
  - `createTemplateFromFile('index').evaluate()` で partial を束ねる
- `constants.gs`
  - シート名
  - 列定義
  - ロール
  - ステータス
  - 固定値
  - Script Properties キー
- `sheets.gs`
  - `SPREADSHEET_ID` を使った `openById()`
  - シート取得とヘッダー整備
- `projectRepository.gs` ほか各 Repository
  - シート CRUD のみ
  - DTO 変換や認可は持たない
- `authService.gs`
  - 認証、セッション、有効期限、失効、認可チェック
  - pepper / secret を Script Properties から取得
- `projectService.gs`
  - 案件 DTO 化
  - 売上/利益の補完
  - Dashboard キャッシュ無効化
- `clientService.gs`
  - クライアント DTO 化
  - 既定利益率の正規化
- `userService.gs`
  - ユーザー管理
  - ロック解除
  - 監査ログ取得
- `dashboardService.gs`
  - Dashboard 集計を server 側で実施
  - `CacheService` を利用

### フロント側

- `index.html`
  - エントリテンプレート
- `IndexShell.html`
  - ルート骨組み
  - Windows のケース非依存ファイルシステム上で `index.html` と衝突しないように分離
- `Styles*.html`
  - 共通 / layout / 画面別 CSS
- `Layout.html`
  - sidebar / topbar / modal
- `Auth.html`
  - ログイン / 初回セットアップ UI
- `Dashboard.html` `Projects.html` `Clients.html` `Users.html`
  - 各 view の静的 shell
- `StateScript.html`
  - `authState / currentView / 各画面 state` の集約
- `ApiScript.html`
  - `google.script.run` の唯一の呼び出し口
- `UiScript.html`
  - ナビゲーション
  - topbar
  - toast
  - modal
  - shell 操作
- `AuthScript.html`
  - localStorage セッション
  - 有効期限チェック
  - 失効
  - 再ログイン導線
- `ChartScript.html`
  - Dashboard 用チャート描画
- `DashboardScript.html` ほか各画面 Script
  - view 単位の取得と描画
  - inline handler を使わず `addEventListener` で統一

## 非同期設計

- 初期表示は shell のみ
- データ取得は view を開いたときだけ行う
- `loadAll` のような全件一括ロードは廃止
- Dashboard 集計は server 側でまとめて返し、client 側は描画だけ担当
- `google.script.run` は `ApiScript.html` に閉じ込める

## セキュリティ

- 認可は server 側で必須
- localStorage には token の他に
  - `status`
  - `expiresAt`
  - `reason`
  を持たせる
- セッション失効時は login へ戻し、理由を表示する
- パスワードハッシュには secret に加えて pepper を使用

## Script Properties

最低限、以下を設定してください。

- `SPREADSHEET_ID`
- `AUTH_SECRET`
- `AUTH_PEPPER`

`AUTH_SECRET` / `AUTH_PEPPER` は未設定でも初回実行時に自動生成されますが、運用では固定管理を推奨します。

## サンプル

### doGet

```javascript
function doGet() {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate()
    .setTitle(APP_INFO.title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
```

### フロント include

```html
<?!= include('Styles'); ?>
<?!= include('IndexShell'); ?>
<?!= include('AppScript'); ?>
```

### API ラッパー

```javascript
function call(method, args, options) {
  var finalArgs = (args || []).slice();
  if (!options || options.secure !== false) {
    finalArgs.unshift(App.state.auth.token);
  }

  return new Promise(function (resolve, reject) {
    google.script.run
      .withSuccessHandler(resolve)
      .withFailureHandler(reject)[method].apply(null, finalArgs);
  });
}
```

## 段階移行手順

1. `SPREADSHEET_ID` を Script Properties に設定し、`getActiveSpreadsheet()` 依存を断つ
2. `doGet()` を template + `include()` 構成へ切り替える
3. `sheets.gs` をシート取得専用へ寄せる
4. Repository を追加し、シート CRUD を移す
5. Service 側へ DTO 変換、認可、業務ロジックを集約する
6. フロントを `Index 相当の骨組み / Styles / Layout / Auth / Dashboard / Projects / Clients / Users / AppScript` に分割する
7. `google.script.run` を `ApiScript` に集約する
8. Dashboard を最優先で server 集計 + chart module 化する
9. 最後に Users / Auth / modal 系を移し、`onclick` を完全撤去する

## 確認ポイント

- `onclick` が残っていない
- `google.script.run` は `ApiScript.html` だけにある
- `getActiveSpreadsheet()` が残っていない
- `updateTopbarActions()` の重複がない
- 画面遷移時に必要なデータだけ取得している
