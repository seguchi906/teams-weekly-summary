/**
 * generate-report.mjs
 *
 * 週次 Teams レポートを自動生成・送信するメインスクリプト。
 *
 * 処理の流れ:
 *   1. Neon DB からプロジェクトデータを取得・集計
 *   2. index.html テンプレートにデータを埋め込んで dist/report.html を生成
 *   3. Puppeteer でスクリーンショット → dist/report-latest.png
 *   4. Microsoft Teams Incoming Webhook に Adaptive Card を送信
 *
 * 必要な環境変数:
 *   DATABASE_URL    - Neon の接続文字列
 *   TEAMS_WEBHOOK_URL - Teams チャネルの Incoming Webhook URL
 *   PAGES_BASE_URL  - GitHub Pages のベース URL（例: https://user.github.io/repo）
 */

import { neon } from "@neondatabase/serverless";
import { readFile, writeFile } from "fs/promises";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import puppeteer from "puppeteer";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dirname, "..");

// ──────────────────────────────────────────────
// 定数
// ──────────────────────────────────────────────

/** 納期まで何日以内を「高リスク」とみなすか */
const HIGH_RISK_DAYS = 14;

/** 集計対象の担当課リスト */
const SECTIONS = ["1課", "2課", "3課"];

// ──────────────────────────────────────────────
// DB クエリ
// ──────────────────────────────────────────────

/**
 * 各課の集計データを取得する。
 *
 * progress-dashboard が使用する Neon DB のスキーマ:
 *
 *   CREATE TABLE app_data (
 *     key        VARCHAR(50) PRIMARY KEY,
 *     value      JSONB NOT NULL,
 *     updated_at TIMESTAMPTZ DEFAULT NOW()
 *   );
 *
 *   key='projects' の value にプロジェクト配列が JSONB で保存されている。
 *   各プロジェクトのキーは camelCase（例: allocationSection1, responsibleSections）。
 *
 * @param {import("@neondatabase/serverless").NeonQueryFunction} sql
 * @returns {Promise<SectionStat[]>}
 */
async function fetchSectionStats(sql) {
  /**
   * app_data の JSONB 配列を jsonb_array_elements で展開し、
   * さらに responsibleSections 配列を jsonb_array_elements_text で展開して課ごとに集計する。
   * 「実質の終了日」は revisedEndDate があればそちらを優先する。
   * responsibleSections が未設定のプロジェクトは COALESCE で空配列扱いとしてスキップ。
   */
  const rows = await sql`
    SELECT
      sec                                       AS section,
      SUM(
        CASE sec
          WHEN '1課' THEN COALESCE((p->>'allocationSection1')::BIGINT, 0)
          WHEN '2課' THEN COALESCE((p->>'allocationSection2')::BIGINT, 0)
          WHEN '3課' THEN COALESCE((p->>'allocationSection3')::BIGINT, 0)
          ELSE 0
        END
      )                                         AS total_allocation,
      COUNT(*)                                  AS total_count,
      SUM(
        CASE WHEN p->>'status' = '進行中'
              AND COALESCE(
                    (p->>'revisedEndDate')::DATE,
                    (p->>'endDate')::DATE
                  ) <= CURRENT_DATE + ${HIGH_RISK_DAYS}
             THEN 1 ELSE 0 END
      )                                         AS high_risk_count
    FROM app_data,
         jsonb_array_elements(value) AS p,
         jsonb_array_elements_text(
           COALESCE(p->'responsibleSections', '[]'::jsonb)
         ) AS sec
    WHERE key = 'projects'
      AND p->>'status' NOT IN ('完納', '仮納品')
    GROUP BY sec
    ORDER BY sec
  `;

  return rows;
}

/**
 * 先週の各課配分合計を取得して「先週比」を計算するためのクエリ。
 * 先週に完納 or 仮納品になった業務の配分額を「先週の生産額」とみなす。
 *
 * @param {import("@neondatabase/serverless").NeonQueryFunction} sql
 * @returns {Promise<Record<string, number>>} section → 先週の配分合計（円）
 */
async function fetchLastWeekAllocation(sql) {
  const rows = await sql`
    SELECT
      sec,
      SUM(
        CASE sec
          WHEN '1課' THEN COALESCE((p->>'allocationSection1')::BIGINT, 0)
          WHEN '2課' THEN COALESCE((p->>'allocationSection2')::BIGINT, 0)
          WHEN '3課' THEN COALESCE((p->>'allocationSection3')::BIGINT, 0)
          ELSE 0
        END
      ) AS last_week_allocation
    FROM app_data,
         jsonb_array_elements(value) AS p,
         jsonb_array_elements_text(
           COALESCE(p->'responsibleSections', '[]'::jsonb)
         ) AS sec
    WHERE key = 'projects'
      AND p->>'status' IN ('完納', '仮納品')
      AND COALESCE(
            (p->>'revisedEndDate')::DATE,
            (p->>'endDate')::DATE
          ) BETWEEN CURRENT_DATE - 14 AND CURRENT_DATE - 7
    GROUP BY sec
  `;

  /** @type {Record<string, number>} */
  const map = {};
  for (const row of rows) {
    map[row.sec] = Number(row.last_week_allocation ?? 0);
  }
  return map;
}

// ──────────────────────────────────────────────
// HTML 生成ヘルパー
// ──────────────────────────────────────────────

/**
 * @param {number} yen 円単位
 * @returns {string} 万円単位（小数点以下切り捨て）
 */
function toManYen(yen) {
  return Math.floor(yen / 10000).toLocaleString("ja-JP");
}

/**
 * 先週比の矢印と色クラスを返す。
 *
 * @param {number} current
 * @param {number} last
 * @returns {{ arrow: string; colorClass: string; label: string }}
 */
function weekOverWeekDisplay(current, last) {
  if (last === 0) {
    return { arrow: "－", colorClass: "text-slate-400", label: "先週比 N/A" };
  }
  const pct = ((current - last) / last) * 100;
  if (pct > 0) {
    return {
      arrow: "↑",
      colorClass: "text-emerald-600",
      label: `先週比 +${pct.toFixed(0)}%`,
    };
  }
  if (pct < 0) {
    return {
      arrow: "↓",
      colorClass: "text-red-500",
      label: `先週比 ${pct.toFixed(0)}%`,
    };
  }
  return { arrow: "→", colorClass: "text-slate-400", label: "先週比 ±0%" };
}

/**
 * 部署1行分の <tr> HTML を生成する。
 *
 * @param {{
 *   section: string;
 *   allocation: number;
 *   lastAllocation: number;
 *   highRiskCount: number;
 *   totalCount: number;
 * }} stat
 * @returns {string}
 */
function buildDeptRow(stat) {
  const { section, allocation, lastAllocation, highRiskCount, totalCount } =
    stat;

  const isHigh = highRiskCount > 0;
  const riskLevel = isHigh ? "high" : "low";
  const riskBg = isHigh ? "bg-red-50/60" : "bg-blue-50/60";
  const badgeColor = isHigh
    ? "bg-red-100 text-red-700"
    : "bg-blue-100 text-blue-700";
  const dotColor = isHigh ? "bg-red-500" : "bg-blue-500";
  const statusLabel = isHigh ? "危険" : "安全";
  const ariaLabel = `高リスク業務: ${statusLabel}（${totalCount}件中${highRiskCount}件）`;

  const { arrow, colorClass, label } = weekOverWeekDisplay(
    allocation,
    lastAllocation
  );

  const amountManYen = toManYen(allocation);

  return `
            <tr data-risk-level="${riskLevel}" class="transition-colors hover:bg-slate-50/50">
              <th scope="row" class="px-4 py-4 text-sm sm:text-base font-semibold text-slate-700">
                ${section}
              </th>
              <td class="px-4 py-4 border-x border-slate-100">
                <div class="text-xl sm:text-2xl font-bold text-slate-800">
                  ${amountManYen} <span class="text-sm sm:text-base font-normal text-slate-400">万円</span>
                </div>
                <div class="flex items-center justify-center gap-1 text-xs ${colorClass} font-medium mt-1">
                  <span>${arrow}</span><span>${label}</span>
                </div>
              </td>
              <td
                class="px-4 py-4 ${riskBg}"
                aria-label="${ariaLabel}"
              >
                <div class="flex flex-col items-center gap-2">
                  <span class="text-sm sm:text-base font-bold text-slate-700">
                    <span class="text-xl sm:text-2xl">${highRiskCount}</span> 件／${totalCount} 件中
                  </span>
                  <span class="inline-flex items-center gap-1.5 rounded-full ${badgeColor} px-3 py-1 text-xs sm:text-sm font-semibold">
                    <span class="w-1.5 h-1.5 rounded-full ${dotColor} flex-shrink-0" aria-hidden="true"></span>
                    ${statusLabel}
                  </span>
                </div>
              </td>
            </tr>`.trimStart();
}

// ──────────────────────────────────────────────
// スクリーンショット
// ──────────────────────────────────────────────

/**
 * Puppeteer で HTML をスクリーンショット撮影し PNG ファイルに保存する。
 *
 * @param {string} htmlPath 読み込む HTML ファイルの絶対パス（file:// URL）
 * @param {string} outputPath 出力 PNG ファイルのパス
 */
async function takeScreenshot(htmlPath, outputPath) {
  const browser = await puppeteer.launch({
    headless: true,
    args: [
      "--no-sandbox",
      "--disable-setuid-sandbox",
      "--disable-dev-shm-usage",
    ],
  });

  try {
    const page = await browser.newPage();
    await page.setViewport({ width: 720, height: 900, deviceScaleFactor: 2 });
    await page.goto(`file://${htmlPath}`, { waitUntil: "networkidle0" });

    // Tailwind の CDN スタイルが適用されるまで少し待機
    await new Promise((r) => setTimeout(r, 1500));

    const card = await page.$(".max-w-2xl");
    if (card) {
      await card.screenshot({ path: outputPath, type: "png" });
    } else {
      await page.screenshot({ path: outputPath, fullPage: true, type: "png" });
    }
  } finally {
    await browser.close();
  }
}

// ──────────────────────────────────────────────
// Teams 通知
// ──────────────────────────────────────────────

/**
 * Teams Incoming Webhook に Adaptive Card を送信する。
 *
 * @param {{
 *   webhookUrl: string;
 *   dateLabel: string;
 *   imageUrl: string;
 *   totalAmount: string;
 *   highRiskDept: string;
 * }} opts
 */
async function sendTeamsNotification(opts) {
  const { webhookUrl, dateLabel, imageUrl, totalAmount, highRiskDept } = opts;

  const payload = {
    type: "message",
    attachments: [
      {
        contentType: "application/vnd.microsoft.card.adaptive",
        contentUrl: null,
        content: {
          $schema: "http://adaptivecards.io/schemas/adaptive-card.json",
          type: "AdaptiveCard",
          version: "1.4",
          body: [
            {
              type: "TextBlock",
              text: `週次レポート　${dateLabel}`,
              weight: "Bolder",
              size: "Medium",
              wrap: true,
            },
            {
              type: "FactSet",
              facts: [
                { title: "合計生産額", value: `${totalAmount} 万円` },
                {
                  title: "高リスク該当",
                  value: highRiskDept || "なし",
                },
              ],
            },
            {
              type: "Image",
              url: imageUrl,
              altText: `週次レポート ${dateLabel}`,
              size: "Stretch",
            },
          ],
          actions: [
            {
              type: "Action.OpenUrl",
              title: "レポートを開く",
              url: imageUrl,
            },
          ],
        },
      },
    ],
  };

  const res = await fetch(webhookUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    const body = await res.text();
    throw new Error(
      `Teams webhook 送信失敗: HTTP ${res.status} ${res.statusText}\n${body}`
    );
  }

  console.log("Teams への通知を送信しました。");
}

// ──────────────────────────────────────────────
// メイン処理
// ──────────────────────────────────────────────

async function main() {
  // ── 環境変数チェック ──
  const DATABASE_URL = process.env.DATABASE_URL;
  const TEAMS_WEBHOOK_URL = process.env.TEAMS_WEBHOOK_URL;
  const PAGES_BASE_URL = process.env.PAGES_BASE_URL ?? "";

  if (!DATABASE_URL) throw new Error("環境変数 DATABASE_URL が未設定です。");
  if (!TEAMS_WEBHOOK_URL)
    throw new Error("環境変数 TEAMS_WEBHOOK_URL が未設定です。");
  if (!PAGES_BASE_URL)
    console.warn(
      "警告: PAGES_BASE_URL が未設定です。Teams の画像 URL が不正になります。"
    );

  // ── 日付ラベル ──
  const now = new Date();
  const dateLabel = now.toLocaleDateString("ja-JP", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });

  console.log(`レポート生成開始: ${dateLabel}`);

  // ── Neon DB 接続 ──
  const sql = neon(DATABASE_URL);

  // ── データ取得 ──
  console.log("Neon DB からデータを取得しています...");
  const [sectionStats, lastWeekMap] = await Promise.all([
    fetchSectionStats(sql),
    fetchLastWeekAllocation(sql),
  ]);

  // ── データが空の場合の処理 ──
  if (sectionStats.length === 0) {
    console.warn(
      "警告: 取得したデータが0件でした。DB のスキーマやクエリを確認してください。"
    );
  }

  // ── 各課データを整形 ──
  /** @type {Map<string, {allocation: number; highRiskCount: number; totalCount: number}>} */
  const sectionMap = new Map();

  for (const row of sectionStats) {
    sectionMap.set(row.section, {
      allocation: Number(row.total_allocation ?? 0),
      highRiskCount: Number(row.high_risk_count ?? 0),
      totalCount: Number(row.total_count ?? 0),
    });
  }

  // ── サマリー集計 ──
  let totalAllocationYen = 0;
  const highRiskSections = [];

  for (const sec of SECTIONS) {
    const d = sectionMap.get(sec);
    if (!d) continue;
    totalAllocationYen += d.allocation;
    if (d.highRiskCount > 0) highRiskSections.push(sec);
  }

  const totalAmountLabel = toManYen(totalAllocationYen);
  const highRiskDeptLabel =
    highRiskSections.length > 0 ? highRiskSections.join("・") : "なし";

  // ── 行 HTML 生成 ──
  const deptRows = SECTIONS.map((sec) => {
    const d = sectionMap.get(sec) ?? {
      allocation: 0,
      highRiskCount: 0,
      totalCount: 0,
    };
    return buildDeptRow({
      section: sec,
      allocation: d.allocation,
      lastAllocation: lastWeekMap[sec] ?? 0,
      highRiskCount: d.highRiskCount,
      totalCount: d.totalCount,
    });
  }).join("\n");

  // ── HTML テンプレート読み込み・置換 ──
  console.log("HTML テンプレートを生成しています...");
  const templatePath = join(ROOT, "index.html");
  let html = await readFile(templatePath, "utf-8");

  html = html
    .replace("{{DATE}}", dateLabel)
    .replace("{{TOTAL_AMOUNT}}", totalAmountLabel)
    .replace("{{HIGH_RISK_DEPT}}", highRiskDeptLabel)
    .replace("{{DEPT_ROWS}}", deptRows);

  const reportHtmlPath = join(ROOT, "dist", "report.html");
  await writeFile(reportHtmlPath, html, "utf-8");
  console.log(`HTML を生成しました: ${reportHtmlPath}`);

  // ── スクリーンショット ──
  console.log("スクリーンショットを撮影しています...");
  const reportPngPath = join(ROOT, "dist", "report-latest.png");
  await takeScreenshot(reportHtmlPath, reportPngPath);
  console.log(`スクリーンショットを保存しました: ${reportPngPath}`);

  // ── Teams 通知送信 ──
  const imageUrl = `${PAGES_BASE_URL}/report-latest.png`;
  console.log(`Teams に通知を送信しています... (画像URL: ${imageUrl})`);

  await sendTeamsNotification({
    webhookUrl: TEAMS_WEBHOOK_URL,
    dateLabel,
    imageUrl,
    totalAmount: totalAmountLabel,
    highRiskDept: highRiskDeptLabel,
  });

  console.log("完了しました。");
}

main().catch((err) => {
  console.error("エラーが発生しました:", err);
  process.exit(1);
});
