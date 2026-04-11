# teams-weekly-summary

Neon DB のプロジェクトデータを週次で集計し、HTML レポートを生成して Microsoft Teams に自動送信する仕組みです。

## 全体フロー

```
GitHub Actions（毎週月曜 09:00 JST）
  ↓
Neon DB からプロジェクトデータを取得・集計
  ↓
Node.js で HTML レポート生成（index.html テンプレート使用）
  ↓
Puppeteer でスクリーンショット（PNG）撮影
  ↓
GitHub Pages（gh-pages ブランチ）に画像を公開
  ↓
Microsoft Teams Incoming Webhook に Adaptive Card 送信
```

## ディレクトリ構成

```
teams-weekly-summary/
├── index.html                      テンプレート HTML（{{変数}} でデータ差し込み）
├── package.json
├── scripts/
│   └── generate-report.mjs         メインスクリプト
├── dist/                           生成ファイル出力先（.gitignore 推奨）
│   ├── report.html
│   └── report-latest.png
└── .github/
    └── workflows/
        └── weekly-report.yml       GitHub Actions ワークフロー
```

---

## セットアップ手順

### 1. Neon DB の準備

このスクリプトは以下のテーブルスキーマを前提としています。  
実際のスキーマに合わせて `scripts/generate-report.mjs` 内の SQL を調整してください。

```sql
CREATE TABLE projects (
  id                  TEXT PRIMARY KEY,
  field               TEXT,
  number              TEXT,
  name                TEXT,
  office              TEXT,
  department          TEXT,
  manager             TEXT,
  start_date          DATE,
  end_date            DATE,
  revised_end_date    DATE,
  contract_amount     BIGINT,
  status              TEXT,            -- '完納' | '仮納品' | '進行中' | '未着手'
  allocation_section1 BIGINT,          -- 1課の配分額（税抜・円）
  allocation_section2 BIGINT,          -- 2課の配分額（税抜・円）
  allocation_section3 BIGINT,          -- 3課の配分額（税抜・円）
  responsible_sections TEXT[]          -- 例: '{1課,2課}'
);
```

Neon の接続文字列は、Neon ダッシュボード → プロジェクト → **Connection string** から取得できます。

```
postgresql://user:password@ep-xxxx.ap-southeast-1.aws.neon.tech/neondb?sslmode=require
```

---

### 2. GitHub Secrets の設定

GitHub リポジトリの **Settings → Secrets and variables → Actions** で以下を登録します。

#### Secrets（機密情報）

| Secret 名           | 内容                                                              |
|---------------------|-------------------------------------------------------------------|
| `DATABASE_URL`      | Neon の接続文字列（`postgresql://...@...neon.tech/...`）           |
| `TEAMS_WEBHOOK_URL` | Teams チャネルの Incoming Webhook URL                              |

**Teams Incoming Webhook の取得手順:**
1. Teams でチャネルを開く
2. チャネル名の右クリック → **コネクタ**（または「チャネルの管理」→「コネクタ」）
3. 「受信 Webhook」→ 構成 → 名前を入力 → 作成
4. 生成された URL をコピーして `TEAMS_WEBHOOK_URL` に設定

#### Variables（非機密情報）

| Variable 名    | 内容                                                             |
|----------------|------------------------------------------------------------------|
| `PAGES_BASE_URL` | GitHub Pages のベース URL（例: `https://your-org.github.io/teams-weekly-summary`） |

**`PAGES_BASE_URL` の確認方法:**  
リポジトリの **Settings → Pages** で公開 URL を確認できます。  
GitHub Pages の有効化は下記 [GitHub Pages の有効化](#4-github-pages-の有効化) を参照してください。

---

### 3. ローカル動作確認

```bash
# 依存パッケージをインストール
npm install

# 環境変数を設定してスクリプトを実行
DATABASE_URL="postgresql://..." \
TEAMS_WEBHOOK_URL="https://..." \
PAGES_BASE_URL="https://your-org.github.io/teams-weekly-summary" \
npm run report
```

生成ファイルは `dist/` に出力されます。

---

### 4. GitHub Pages の有効化

1. リポジトリの **Settings → Pages** を開く
2. **Source** を `Deploy from a branch` に設定
3. **Branch** を `gh-pages` / `/ (root)` に設定して保存
4. 初回ワークフロー実行後に `gh-pages` ブランチが作成され、Pages が有効になります

---

### 5. ワークフローの確認・手動実行

- **自動実行:** 毎週月曜 09:00 JST
- **手動実行:** GitHub の **Actions** タブ → `週次 Teams レポート送信` → `Run workflow`

---

## 集計ロジックの概要

| 指標           | 集計方法                                                              |
|----------------|-----------------------------------------------------------------------|
| 生産額（各課） | `allocation_section1 / 2 / 3` の合計（`responsible_sections` が対象課を含むプロジェクト） |
| 高リスク件数   | `status = '進行中'` かつ納期まで **14日以内** のプロジェクト数         |
| 先週比         | 先週（7〜14日前）に完納・仮納品になったプロジェクトの配分額との比較     |

集計条件は `scripts/generate-report.mjs` 内の定数で調整可能です。

```js
// 納期まで何日以内を「高リスク」とみなすか
const HIGH_RISK_DAYS = 14;

// 集計対象の担当課リスト
const SECTIONS = ["1課", "2課", "3課"];
```

---

## トラブルシューティング

| 症状                          | 確認ポイント                                                   |
|-------------------------------|----------------------------------------------------------------|
| DB 接続エラー                 | `DATABASE_URL` の接続文字列が正しいか確認                       |
| データが 0 件                 | SQL のテーブル名・カラム名を実際のスキーマに合わせて調整         |
| Teams に届かない              | `TEAMS_WEBHOOK_URL` が有効か、チャネルの Webhook 設定を確認     |
| 画像が Teams に表示されない   | `PAGES_BASE_URL` が正しいか、GitHub Pages が有効かを確認        |
| Puppeteer でエラー            | Actions ログで Chromium 依存ライブラリのインストールを確認       |
