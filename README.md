# teams-weekly-summary

Neon DB のプロジェクトデータを週次で集計し、HTML レポートを生成して Microsoft Teams に自動送信する仕組みです。

## 全体フロー

```
GitHub Actions（毎週金曜 06:05 JST）
  ↓
Neon DB からプロジェクトデータを取得・集計
  ↓
Node.js で HTML レポート生成（週次表: index.html、AI 提案: index-ai.html）
  ↓
Playwright でスクリーンショット（週次表用・AI 用で別 PNG）
  ↓
GitHub Pages にアーティファクト公開
  ↓
Microsoft Teams Incoming Webhook に送信（週次サマリー・リスク詳細・任意で AI 専用チャネル）
```

## ディレクトリ構成

```
teams-weekly-summary/
├── index.html                      週次サマリー表テンプレート（従来どおり）
├── index-ai.html                 AI 提案のみのテンプレート（別チャネル用）
├── config/
│   └── weekly-ai-prompt.md         Gemini 用プロンプト（{{CONTEXT}}）
├── package.json
├── scripts/
│   ├── generate-report.mjs         メインスクリプト
│   ├── weekly-ai-suggestions.mjs   AI 提案生成・HTML レンダリング
│   └── weekly-ai-fallback.mjs      API 未使用時のルールベース提案
├── dist/                           生成ファイル出力先（.gitignore 推奨）
│   ├── report.html / report-*.png / report-latest.png（週次表）
│   ├── report-ai.html / report-ai-*.png / report-ai-latest.png（AI）
│   └── ai-suggestions.json
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

| Secret 名                 | 内容 |
|---------------------------|-------------------------------------------------------------------|
| `DATABASE_URL`            | Neon の接続文字列（`postgresql://...@...neon.tech/...`）           |
| `TEAMS_WEBHOOK_URL`     | **週次サマリー**（画像付きカード）用の Incoming Webhook URL |
| `TEAMS_RISK_WEBHOOK_URL` | **リスク該当案件の詳細**用の Incoming Webhook URL（別チャネル向け） |
|                           | 未設定のときは `TEAMS_WEBHOOK_URL` と同じチャネルに連続投稿します。 |
| `TEAMS_AI_WEBHOOK_URL`    | **週次 AI 提案**（別 HTML・別 PNG）専用チャネルの Incoming Webhook。**未設定**のときは `report-ai.html` と画像は生成されますが、AI 用の Teams 通知は行いません。 |
| `GEMINI_API_KEY`          | Google AI（Gemini）の API キー。**未設定でも**週次レポートは動作し、AI 提案はルールベースのフォールバックで生成されます。 |

**Teams Incoming Webhook の取得手順:**
1. Teams でチャネルを開く
2. チャネル名の右クリック → **コネクタ**（または「チャネルの管理」→「コネクタ」）
3. 「受信 Webhook」→ 構成 → 名前を入力 → 作成
4. 生成された URL をコピーして `TEAMS_WEBHOOK_URL`（必要なら `TEAMS_RISK_WEBHOOK_URL`・`TEAMS_AI_WEBHOOK_URL`）に設定

#### Variables（非機密情報）

| Variable 名    | 内容                                                             |
|----------------|------------------------------------------------------------------|
| `PAGES_BASE_URL` | GitHub Pages のベース URL（例: `https://your-org.github.io/teams-weekly-summary`） |

**`PAGES_BASE_URL` の確認方法:**  
リポジトリの **Settings → Pages** で公開 URL を確認できます。  
GitHub Pages の有効化は下記 [GitHub Pages の有効化](#4-github-pages-の有効化) を参照してください。

**週次表と AI 提案の分離:**  
週次サマリー・リスク詳細は従来どおり `TEAMS_WEBHOOK_URL` / `TEAMS_RISK_WEBHOOK_URL` へ投稿します。AI 提案レポートは **`TEAMS_AI_WEBHOOK_URL` を設定した場合のみ** 別チャネルへ投稿します（同じ GitHub Pages 上に `report-ai-*.png` を載せ、その URL をカードに埋め込みます）。

**AI 提案（Gemini）:**  
`GEMINI_API_KEY` を設定すると、週次集計データをコンテキストに **Gemini で3件の提案**を生成します。未設定時や API 失敗時は、同じ3件形式で **ルールベースのフォールバック**に切り替わります。  
任意で `GEMINI_MODEL` に利用するモデル ID（例: `models/gemini-2.5-flash`）を指定できます。未指定時はスクリプト内の既定候補を順に試します。

---

### 3. ローカル動作確認

```bash
# 依存パッケージをインストール
npm install

# 環境変数を設定してスクリプトを実行
OVERALL_PROJECT_SCHEDULE_URL="https://..." \
PROGRESS_BASHBOARD_URL="postgresql://..." \
TEAMS_WEBHOOK_URL="https://..." \
TEAMS_RISK_WEBHOOK_URL="https://..." \
TEAMS_AI_WEBHOOK_URL="https://..." \
PAGES_BASE_URL="https://your-org.github.io/teams-weekly-summary" \
GEMINI_API_KEY="..." \
npm run report
```

`GEMINI_API_KEY` は省略可能です（その場合、AI 提案はフォールバックのみ）。`TEAMS_AI_WEBHOOK_URL` も省略可能です（その場合、AI 用 HTML/PNG のみ生成し、AI チャネルへの投稿はしません）。

生成ファイルは `dist/` に出力されます（週次表: `report.html`・`report-*.png`・`report-latest.png`、AI: `report-ai.html`・`report-ai-*.png`・`report-ai-latest.png`、`ai-suggestions.json`）。

---

### 4. GitHub Pages の有効化

1. リポジトリの **Settings → Pages** を開く
2. **Source** を `Deploy from a branch` に設定
3. **Branch** を `gh-pages` / `/ (root)` に設定して保存
4. 初回ワークフロー実行後に `gh-pages` ブランチが作成され、Pages が有効になります

---

### 5. ワークフローの確認・手動実行

- **自動実行:** 毎週金曜 06:05 JST
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
| Teams に届かない              | `TEAMS_WEBHOOK_URL` / `TEAMS_RISK_WEBHOOK_URL` が有効か、各チャネルの Webhook を確認 |
| 画像が Teams に表示されない   | `PAGES_BASE_URL` が正しいか、GitHub Pages が有効かを確認        |
| Puppeteer でエラー            | Actions ログで Chromium 依存ライブラリのインストールを確認       |
