# 案件管理システム

Google Apps Script (GAS) と Google スプレッドシートで動く、案件・クライアント管理用の Web アプリです。  
Apps Script を Web アプリとしてデプロイし、ブラウザから案件一覧、クライアント管理、ダッシュボード分析を利用できます。

## 概要

- データ保存先は Google スプレッドシート
- フロントエンドは `src/index.html` に集約したシングルページ構成
- バックエンドは Apps Script の `.gs` ファイルを役割ごとに分割
- PC / モバイルの両方で使えるレスポンシブ UI

## 主な機能

### ダッシュボード

- 期間切り替え: `全期間` / `今年` / `今月` / 任意の年・月
- KPI 表示
  - 予測売上
  - パイプライン売上
  - 進行中売上
  - 商談中売上
  - 確定売上計
  - 確定利益計
  - 予測利益率
  - 平均利益率
  - 平均単価
  - 平均リードタイム
  - 完了率
  - 稼働率
- チャート表示
  - 月別 実績売上 / 利益
  - 月別 予測売上構成
  - 累積実績の推移
  - 利益率の推移
  - ステータス別 売上
  - 月別 案件件数
  - クライアント別売上比率
  - 案件サイズ分布
- ステータスカードや期間フィルタから、そのまま案件一覧へ絞り込み遷移

### 案件管理

- 案件の追加 / 編集 / 削除
- 5 ステータス管理: `商談中` `進行中` `完了` `停止` `保留`
- キーワード検索
  - 対象: `案件名` `クライアント名` `備考`
- 条件フィルタ
  - ステータス
  - 売上の下限 / 上限
  - 日付範囲
- 案件登録モーダルからクライアントをその場で新規登録可能
- クライアントに既定利益率が設定されていれば、案件登録時に利益へ自動反映
- 利益は従来どおり手入力で上書き可能
- 利益が空欄の場合は、クライアント既定利益率を優先し、未設定なら売上と同額として扱う
- ステータスを `完了` に変更し、完了日が空欄なら更新時に当日を自動補完

### クライアント管理

- クライアントの追加 / 編集 / 削除
- 管理項目: `クライアント名` `既定利益率` `担当者` `メール` `電話` `備考`
- クライアントごとの案件数と累計売上を一覧表示

## 技術構成

- Backend: Google Apps Script
- Frontend: HTML / CSS / Vanilla JavaScript
- Data store: Google スプレッドシート
- Charts: `Chart.js` CDN
- Fonts: Google Fonts (`Noto Sans JP`, `JetBrains Mono`)
- Build: なし

## GitHub Actions で自動デプロイ

`src/` 配下だけを Google Apps Script へ反映するように、GitHub Actions から `clasp` でデプロイできます。  
このリポジトリでは workflow 実行時に `.clasp.json` を生成し、`rootDir` を `src` に固定しているため、GAS に push されるのは `src/` 配下のみです。

### 追加したもの

- `.github/workflows/deploy-gas.yml`
  - `main` ブランチへの push、または手動実行でデプロイ
  - `src/**` の変更時だけ自動実行
- `package.json`
  - GitHub Actions で `@google/clasp` を利用
- `.gitignore`
  - `.clasp.json` と `.clasprc.json` を Git 管理外に設定

### 事前準備

1. 対象の Google スプレッドシートに紐づく Apps Script を 1 回だけ Web アプリとして手動デプロイします。
2. Apps Script 側で以下を控えます。
   - `Script ID`
   - 既存 Web アプリの `Deployment ID`
3. GitHub の `Settings > Secrets and variables > Actions` に以下の Secrets を登録します。
   - `GAS_SCRIPT_ID`
     - Apps Script プロジェクト設定の `スクリプト ID`
   - `GAS_DEPLOYMENT_ID`
     - `デプロイを管理` で確認できる既存 Web アプリの deployment ID
   - `GAS_CLASPRC_JSON`
     - ローカルで `clasp login` 済みの認証情報 JSON

### `GAS_CLASPRC_JSON` の作り方

1. ローカルで依存関係を入れます。

   ```bash
   npm install
   ```

2. Google アカウントで `clasp` にログインします。

   ```bash
   npx clasp login
   ```

3. 生成された `~/.clasprc.json` の中身を、GitHub Secret `GAS_CLASPRC_JSON` にそのまま貼り付けます。

Windows では通常 `C:\Users\<ユーザー名>\.clasprc.json`、macOS / Linux では `~/.clasprc.json` に保存されます。

### デプロイの流れ

- `main` に push
- GitHub Actions が `.clasp.json` を生成
- `rootDir: "src"` で `src/` 配下のみ `clasp push --force`
- 既存の `GAS_DEPLOYMENT_ID` に対して `clasp update-deployment` を実行

これで Web アプリの URL を変えずに更新できます。

## ディレクトリ構成

```text
case-manage/
└─ src/
   ├─ app.gs               Web アプリのエントリーポイント
   ├─ constants.gs         シート名・列定義
   ├─ projectService.gs    案件 CRUD
   ├─ clientService.gs     クライアント CRUD
   ├─ dashboardService.gs  ダッシュボード用の集計関数
   ├─ sheets.gs            シート取得・初期作成
   ├─ utils.gs             ID生成、日時整形、シート行変換
   └─ index.html           UI 全体
```

## スプレッドシート構成

### `案件データ`

- シート名: `案件データ`
- 列定義:
  - `ID`
  - `案件名`
  - `クライアントID`
  - `クライアント名`
  - `売上`
  - `利益`
  - `ステータス`
  - `完了日`
  - `備考`
  - `登録日`
  - `更新日`

### `クライアントマスタ`

- シート名: `クライアントマスタ`
- 列定義:
  - `ID`
  - `クライアント名`
  - `既定利益率`
  - `担当者`
  - `メール`
  - `電話`
  - `備考`
  - `登録日`

### 自動入力される項目

- 案件
  - `ID`
  - `登録日`
  - `更新日`
- クライアント
  - `ID`
  - `登録日`

## セットアップ

1. Google スプレッドシートを新規作成します。
2. スプレッドシートで `拡張機能 > Apps Script` を開きます。
3. `src/` 配下の `.gs` ファイルを Apps Script プロジェクトへ追加し、内容を貼り付けます。
4. HTML ファイル `index.html` を追加し、`src/index.html` の内容を貼り付けます。
5. スクリプトを保存し、初回実行時の権限承認を行います。
6. `デプロイ > 新しいデプロイ > ウェブアプリ` から公開します。
7. 初回アクセス時に `案件データ` と `クライアントマスタ` が存在しなければ自動作成されます。

## セットアップ時の注意

- この実装は `SpreadsheetApp.getActiveSpreadsheet()` を利用しているため、スプレッドシートに紐づく Apps Script として配置してください。
- `.gs` ファイル名は任意ですが、`app.gs` 内で `HtmlService.createHtmlOutputFromFile('index')` を使っているため、HTML ファイル名は `index.html` のままにしてください。
- ダッシュボードの期間判定、案件一覧の日付絞り込み、並び順は `完了日` を優先し、未入力時は `登録日` を使います。
- 外部 CDN を利用するため、`cdn.jsdelivr.net` と `fonts.googleapis.com` にアクセスできる環境が必要です。

## 補足

- 案件一覧はモバイルではカード表示、デスクトップではテーブル表示に切り替わります。
- シートのヘッダー行は初回作成時に自動整形され、1 行目を固定します。
- 既存の `クライアントマスタ` にも、不足ヘッダーがあれば自動で補完されます。
