/**
 * generate-report.mjs
 *
 * 週次 Teams レポートを自動生成・送信するメインスクリプト。
 *
 * 処理の流れ:
 *   1. overall-project-schedule の REST API からプロジェクト（受注金額・担当課）を取得
 *   2. PROGRESS_BASHBOARD Neon DB から進捗率を取得
 *   3. overall.number === progress.id（業務番号）でジョインし、
 *      生産金額 = allocationSectionN × currentProgress / 100 を課ごとに集計
 *   4. index.html テンプレートにデータを埋め込んで dist/report.html を生成
 *   5. Playwright でスクリーンショット → dist/report-latest.png
 *   6. Teams Incoming Webhook に Adaptive Card を送信
 *
 * 必要な環境変数:
 *   OVERALL_PROJECT_SCHEDULE_URL - overall-project-schedule アプリの URL
 *                                  （例: https://your-app.example.com）
 *                                  /api/projects-data エンドポイントを使用する
 *   PROGRESS_BASHBOARD_URL       - PROGRESS_BASHBOARD の Neon 接続文字列
 *   TEAMS_WEBHOOK_URL            - Teams チャネルの Incoming Webhook URL
 *   PAGES_BASE_URL               - GitHub Pages のベース URL（例: https://user.github.io/repo）
 *
 * overall-project-schedule API レスポンス（GET /api/projects-data）:
 *   { projects: [ { id, number, name, status, contractAmount,
 *                   allocationSection1, allocationSection2, allocationSection3,
 *                   responsibleSections, revisedEndDate, endDate, ... } ] }
 *
 * PROGRESS_BASHBOARD の app_data スキーマ:
 *   app_data テーブル, key = 'projects', value は JSONB 配列。
 *   各要素: { id: "46-003", name: "...", weeklyProgress: [null, 30, 60, ...] }
 */

import { neon } from "@neondatabase/serverless";
import { mkdir, readFile, writeFile } from "fs/promises";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import { chromium } from "playwright";

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
 * overall-project-schedule の REST API からプロジェクト一覧を取得する。
 * エンドポイント: GET {baseUrl}/api/projects-data
 * レスポンス: { projects: Project[], title, updatedAt, fiscalPeriods }
 *
 * @param {string} baseUrl  overall-project-schedule アプリのベース URL
 * @returns {Promise<Object[]>} プロジェクトオブジェクトの配列
 */
async function fetchAllProjects(baseUrl) {
  const url = `${baseUrl.replace(/\/$/, "")}/api/projects-data`;
  const res = await fetch(url);
  if (!res.ok) {
    throw new Error(
      `overall-project-schedule API 呼び出し失敗: HTTP ${res.status} ${res.statusText}\nURL: ${url}`
    );
  }
  const data = await res.json();
  if (!Array.isArray(data.projects)) {
    throw new Error(
      `overall-project-schedule API のレスポンスに projects 配列がありません。\n受信データ: ${JSON.stringify(data).slice(0, 200)}`
    );
  }
  return data.projects;
}

/**
 * PROGRESS_BASHBOARD DB から業務番号ごとの現在進捗率（0〜100）マップを取得する。
 *
 * スキーマ想定:
 *   { id: "46-003", name: "...", weeklyProgress: [null, 30, 60, null, 75, ...] }
 *
 * `id` が OVERALL_PROJECT_SCHEDULE の `number`（業務番号）と対応する結合キー。
 * 現在進捗は weeklyProgress の末尾から遡って最初の非 null 値を使用する。
 *
 * @param {import("@neondatabase/serverless").NeonQueryFunction} sql
 * @returns {Promise<Map<string, number>>} 業務番号 → 現在進捗率（0〜100）
 */
async function fetchProgressRates(sql) {
  const rows = await sql`
    SELECT p AS project
    FROM   app_data,
           jsonb_array_elements(value) AS p
    WHERE  key = 'projects'
  `;

  const map = new Map();
  for (const { project } of rows) {
    if (project.id == null) continue;
    const weeklyProgress = Array.isArray(project.weeklyProgress)
      ? project.weeklyProgress
      : [];
    const currentProgress =
      [...weeklyProgress].reverse().find((v) => v != null) ?? 0;
    map.set(String(project.id), Number(currentProgress));
  }
  return map;
}

// ──────────────────────────────────────────────
// 生産金額の集計（JS ジョイン）
// ──────────────────────────────────────────────

/**
 * overall.number === progress.id でジョインし、課ごとの生産金額を集計する。
 *
 * 生産金額の計算:
 *   section1Production = allocationSection1 * currentProgress / 100
 *   section2Production = allocationSection2 * currentProgress / 100
 *   section3Production = allocationSection3 * currentProgress / 100
 *
 * totalCount / highRiskCount は responsibleSections を使って課ごとにカウントする。
 *
 * @param {Object[]} projects       OVERALL_PROJECT_SCHEDULE から取得した全プロジェクト
 * @param {Map<string, number>} progressMap  業務番号 → 現在進捗率（0〜100）
 * @returns {{
 *   sectionMap: Map<string, {production: number; highRiskCount: number; totalCount: number}>;
 *   matchedCount: number;
 *   unmatchedOverallCount: number;
 *   unmatchedProgressCount: number;
 * }}
 */
function computeSectionStats(projects, progressMap) {
  const sectionMap = new Map(
    SECTIONS.map((sec) => [sec, { production: 0, highRiskCount: 0, totalCount: 0 }])
  );

  let matchedCount = 0;
  let unmatchedOverallCount = 0;
  const matchedProgressIds = new Set();

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (const project of projects) {
    if (["完納", "仮納品"].includes(project.status)) continue;

    // 結合キー: overall.number === progress.id（業務番号）
    const businessNumber = String(project.number ?? "");
    const currentProgress = progressMap.get(businessNumber); // 0〜100 or undefined

    if (currentProgress == null) {
      unmatchedOverallCount++;
    } else {
      matchedCount++;
      matchedProgressIds.add(businessNumber);
    }

    const effectivePct = currentProgress ?? 0; // 未マッチは進捗 0% 扱い

    // 各課の生産金額 = 配分額 × 進捗率 / 100
    const alloc1 = Number(project.allocationSection1 ?? 0);
    const alloc2 = Number(project.allocationSection2 ?? 0);
    const alloc3 = Number(project.allocationSection3 ?? 0);
    sectionMap.get("1課").production += Math.floor(alloc1 * effectivePct / 100);
    sectionMap.get("2課").production += Math.floor(alloc2 * effectivePct / 100);
    sectionMap.get("3課").production += Math.floor(alloc3 * effectivePct / 100);

    // 高リスク判定（納期まで HIGH_RISK_DAYS 日以内かつ「進行中」）
    const rawEndDate = project.revisedEndDate ?? project.endDate ?? null;
    const isHighRisk =
      project.status === "進行中" &&
      rawEndDate != null &&
      (new Date(rawEndDate) - today) / 86400000 <= HIGH_RISK_DAYS;

    // totalCount / highRiskCount は responsibleSections 単位でカウント
    const sections = Array.isArray(project.responsibleSections)
      ? project.responsibleSections
      : [];
    for (const sec of sections) {
      if (!sectionMap.has(sec)) continue;
      const d = sectionMap.get(sec);
      d.totalCount++;
      if (isHighRisk) d.highRiskCount++;
    }
  }

  // progressMap のうち overall と結合できなかった件数
  const unmatchedProgressCount = [...progressMap.keys()].filter(
    (id) => !matchedProgressIds.has(id)
  ).length;

  return { sectionMap, matchedCount, unmatchedOverallCount, unmatchedProgressCount };
}

/**
 * 先週（CURRENT_DATE-14 〜 CURRENT_DATE-7）に完納／仮納品になった
 * プロジェクトの課ごと生産金額合計を返す（先週比の計算に使用）。
 *
 * 結合キー: overall.number === progress.id（業務番号）
 * 完納案件は進捗率マップにあればその値を使い、なければ 100% とみなす。
 *
 * @param {Object[]} projects
 * @param {Map<string, number>} progressMap  業務番号 → 現在進捗率（0〜100）
 * @returns {Record<string, number>} section → 先週の生産金額合計（円）
 */
function computeLastWeekProduction(projects, progressMap) {
  const map = {};
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const sevenDaysAgo    = new Date(today); sevenDaysAgo.setDate(today.getDate() - 7);
  const fourteenDaysAgo = new Date(today); fourteenDaysAgo.setDate(today.getDate() - 14);

  for (const project of projects) {
    if (!["完納", "仮納品"].includes(project.status)) continue;

    const rawEndDate = project.revisedEndDate ?? project.endDate ?? null;
    if (rawEndDate == null) continue;

    const endDate = new Date(rawEndDate);
    if (endDate < fourteenDaysAgo || endDate > sevenDaysAgo) continue;

    // 結合キー: overall.number === progress.id（業務番号）
    const businessNumber = String(project.number ?? "");
    const pct = progressMap.get(businessNumber) ?? 100; // 完納はデフォルト 100%

    const alloc1 = Number(project.allocationSection1 ?? 0);
    const alloc2 = Number(project.allocationSection2 ?? 0);
    const alloc3 = Number(project.allocationSection3 ?? 0);

    map["1課"] = (map["1課"] ?? 0) + Math.floor(alloc1 * pct / 100);
    map["2課"] = (map["2課"] ?? 0) + Math.floor(alloc2 * pct / 100);
    map["3課"] = (map["3課"] ?? 0) + Math.floor(alloc3 * pct / 100);
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
 *   production: number;
 *   lastProduction: number;
 *   highRiskCount: number;
 *   totalCount: number;
 * }} stat
 * @returns {string}
 */
function buildDeptRow(stat) {
  const { section, production, lastProduction, highRiskCount, totalCount } = stat;

  const isHigh = highRiskCount > 0;
  const riskLevel   = isHigh ? "high" : "low";
  const riskBg      = isHigh ? "bg-red-50/60"              : "bg-blue-50/60";
  const badgeColor  = isHigh ? "bg-red-100 text-red-700"   : "bg-blue-100 text-blue-700";
  const dotColor    = isHigh ? "bg-red-500"                 : "bg-blue-500";
  const statusLabel = isHigh ? "危険"                       : "安全";
  const ariaLabel   = `高リスク業務: ${statusLabel}（${totalCount}件中${highRiskCount}件）`;

  const { arrow, colorClass, label } = weekOverWeekDisplay(production, lastProduction);
  const amountManYen = toManYen(production);

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
 * Playwright で HTML をスクリーンショット撮影し PNG ファイルに保存する。
 *
 * @param {string} htmlPath 読み込む HTML ファイルの絶対パス（file:// URL）
 * @param {string} outputPath 出力 PNG ファイルのパス
 */
async function takeScreenshot(htmlPath, outputPath) {
  const browser = await chromium.launch({
    args: ["--no-sandbox", "--disable-setuid-sandbox"],
  });

  try {
    const page = await browser.newPage();
    await page.setViewportSize({ width: 720, height: 900 });
    await page.goto(`file://${htmlPath}`, { waitUntil: "networkidle" });

    const card = page.locator(".max-w-2xl");
    if (await card.count() > 0) {
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
                { title: "高リスク該当", value: highRiskDept || "なし" },
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
  const OVERALL_PROJECT_SCHEDULE_URL = process.env.OVERALL_PROJECT_SCHEDULE_URL;
  const PROGRESS_BASHBOARD_URL       = process.env.PROGRESS_BASHBOARD_URL;
  const TEAMS_WEBHOOK_URL            = process.env.TEAMS_WEBHOOK_URL;
  const PAGES_BASE_URL               = process.env.PAGES_BASE_URL ?? "";

  if (!OVERALL_PROJECT_SCHEDULE_URL)
    throw new Error("環境変数 OVERALL_PROJECT_SCHEDULE_URL が未設定です。");
  if (!PROGRESS_BASHBOARD_URL)
    throw new Error("環境変数 PROGRESS_BASHBOARD_URL が未設定です。");
  if (!TEAMS_WEBHOOK_URL)
    throw new Error("環境変数 TEAMS_WEBHOOK_URL が未設定です。");
  if (!PAGES_BASE_URL)
    console.warn("警告: PAGES_BASE_URL が未設定です。Teams の画像 URL が不正になります。");

  // ── 日付ラベル ──
  const now = new Date();
  const dateLabel = now.toLocaleDateString("ja-JP", {
    year: "numeric",
    month: "long",
    day: "numeric",
  });

  console.log(`レポート生成開始: ${dateLabel}`);

  // ── PROGRESS_BASHBOARD の Neon DB に接続 ──
  const sqlProgress = neon(PROGRESS_BASHBOARD_URL);

  // ── 両ソースから並列取得 ──
  console.log("データを取得しています...");
  console.log(`  overall-project-schedule API: ${OVERALL_PROJECT_SCHEDULE_URL}`);
  const [allProjects, progressMap] = await Promise.all([
    fetchAllProjects(OVERALL_PROJECT_SCHEDULE_URL),
    fetchProgressRates(sqlProgress),
  ]);

  console.log(`  overall 案件数: ${allProjects.length} 件`);
  console.log(`  progress 案件数: ${progressMap.size} 件`);

  // ── JS 側でジョイン → 生産金額を集計 ──
  // 結合キー: overall.number === progress.id（業務番号）
  const { sectionMap, matchedCount, unmatchedOverallCount, unmatchedProgressCount } =
    computeSectionStats(allProjects, progressMap);

  console.log(`  結合できた件数:                  ${matchedCount} 件`);
  console.log(`  結合できなかった overall 案件数: ${unmatchedOverallCount} 件（進捗率 0% 扱い）`);
  console.log(`  結合できなかった progress 案件数: ${unmatchedProgressCount} 件`);

  if (matchedCount === 0) {
    console.warn(
      "警告: 結合できた案件が 0 件でした。overall.number と progress.id の値を確認してください。"
    );
  }

  // ── 先週の生産金額（先週比用） ──
  const lastWeekMap = computeLastWeekProduction(allProjects, progressMap);

  // ── サマリー集計 ──
  let totalProductionYen = 0;
  const highRiskSections = [];

  for (const sec of SECTIONS) {
    const d = sectionMap.get(sec);
    if (!d) continue;
    totalProductionYen += d.production;
    if (d.highRiskCount > 0) highRiskSections.push(sec);
  }

  const totalAmountLabel  = toManYen(totalProductionYen);
  const highRiskDeptLabel =
    highRiskSections.length > 0 ? highRiskSections.join("・") : "なし";

  // ── 課ごとの生産金額をコンソールに出力 ──
  console.log("\n─── 課ごとの合計生産金額 ───────────────────");
  for (const sec of SECTIONS) {
    const d = sectionMap.get(sec);
    const prod = d?.production ?? 0;
    const total = d?.totalCount ?? 0;
    const risk  = d?.highRiskCount ?? 0;
    console.log(
      `  ${sec}: ${toManYen(prod)} 万円` +
      `  （件数: ${total}, 高リスク: ${risk}）`
    );
  }
  console.log(`  総生産金額: ${totalAmountLabel} 万円`);
  console.log("────────────────────────────────────────────\n");

  // ── 行 HTML 生成 ──
  const deptRows = SECTIONS.map((sec) => {
    const d = sectionMap.get(sec) ?? { production: 0, highRiskCount: 0, totalCount: 0 };
    return buildDeptRow({
      section:       sec,
      production:    d.production,
      lastProduction: lastWeekMap[sec] ?? 0,
      highRiskCount: d.highRiskCount,
      totalCount:    d.totalCount,
    });
  }).join("\n");

  // ── HTML テンプレート読み込み・置換 ──
  console.log("HTML テンプレートを生成しています...");
  const templatePath = join(ROOT, "index.html");
  let html = await readFile(templatePath, "utf-8");

  html = html
    .replace("{{DATE}}",          dateLabel)
    .replace("{{TOTAL_AMOUNT}}",  totalAmountLabel)
    .replace("{{HIGH_RISK_DEPT}}", highRiskDeptLabel)
    .replace("{{DEPT_ROWS}}",     deptRows);

  const distDir = join(ROOT, "dist");
  await mkdir(distDir, { recursive: true });

  const reportHtmlPath = join(distDir, "report.html");
  await writeFile(reportHtmlPath, html, "utf-8");
  console.log(`HTML を生成しました: ${reportHtmlPath}`);

  // ── スクリーンショット ──
  console.log("スクリーンショットを撮影しています...");
  const reportPngPath = join(distDir, "report-latest.png");
  await takeScreenshot(reportHtmlPath, reportPngPath);
  console.log(`スクリーンショットを保存しました: ${reportPngPath}`);

  // ── Teams 通知送信 ──
  const imageUrl = `${PAGES_BASE_URL}/report-latest.png`;
  console.log(`Teams に通知を送信しています... (画像URL: ${imageUrl})`);

  await sendTeamsNotification({
    webhookUrl:   TEAMS_WEBHOOK_URL,
    dateLabel,
    imageUrl,
    totalAmount:  totalAmountLabel,
    highRiskDept: highRiskDeptLabel,
  });

  console.log("完了しました。");
}

main().catch((err) => {
  console.error("エラーが発生しました:", err);
  process.exit(1);
});
