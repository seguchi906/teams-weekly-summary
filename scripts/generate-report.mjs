/**
 * generate-report.mjs
 *
 * 週次 Teams レポートを自動生成・送信するメインスクリプト。
 *
 * 処理の流れ:
 *   1. overall-project-schedule の REST API からプロジェクト（受注金額・担当課）を取得
 *   2. PROGRESS_BASHBOARD Neon DB から weeklyProgress（末尾から最大3つの非 null 値）を取得
 *   3. overall.number === progress.id（業務番号）でジョインし、
 *      今週生産 = allocation × (現在% - 前回%) / 100、先週分 = allocation × (前回% - その前%) / 100（増加・減少どちらも反映）を課ごとに集計
 *   4. calculateRisk で案件ごとにリスク評価し、課別に highRiskCount / cautionCount を集計
 *   5. index.html テンプレートにデータを埋め込んで dist/report.html を生成
 *   6. Playwright でスクリーンショット → dist/report-YYYYMMDD-<ms>.png を保存し、あわせて report-latest.png にも同内容をコピー
 *   7. Teams Incoming Webhook に Adaptive Card を送信
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
 *                   responsibleSections, responsibleDept, revisedEndDate, endDate,
 *                   startDate, originalStartDate, outsourcingAmount, ... } ] }
 *
 * PROGRESS_BASHBOARD の app_data スキーマ:
 *   app_data テーブル, key = 'projects', value は JSONB 配列。
 *   各要素: { id: "46-003", name: "...", weeklyProgress: [null, 30, 60, ...] }
 */

import { neon } from "@neondatabase/serverless";
import { copyFile, mkdir, readFile, writeFile } from "fs/promises";
import { fileURLToPath } from "url";
import { dirname, join } from "path";
import { chromium } from "playwright";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dirname, "..");

/**
 * @param {string} base
 */
function normalizePagesBaseUrl(base) {
  return String(base ?? "").replace(/\/+$/, "");
}

/**
 * @param {Date} d レポート基準日時（runner のローカルタイムゾーン）
 * @returns {string} 例: report-20260413-1713012345678.png
 */
function buildTimestampedReportPngFileName(d) {
  const y = d.getFullYear();
  const mo = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `report-${y}${mo}${day}-${Date.now()}.png`;
}

// ──────────────────────────────────────────────
// 定数
// ──────────────────────────────────────────────

/** 集計対象の担当課リスト */
const SECTIONS = ["1課", "2課", "3課"];

/** Max risk-detail rows to log (console) */
const TOP_RISK_LOG_LIMIT = 10;

/**
 * 現在進捗率（weeklyProgress 末尾の現在値、0〜100）がこの値以上の案件は、
 * 生産・リスク・先週比の集計から除外する（ステータスが完納前でも 100% なら対象外）。
 */
const EXCLUDE_FROM_METRICS_PROGRESS_MIN = 100;

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
 * @param {unknown} v
 * @returns {number}
 */
function clampProgressPercent(v) {
  const n = Number(v);
  if (!Number.isFinite(n)) return 0;
  return Math.min(100, Math.max(0, n));
}

/**
 * ダッシュボード / Overall Project Schedule の「進行中」フィルタと同じ定義。
 * weeklyProgress に非 null が1つも無ければ null（進捗率未記録）。
 *
 * @param {unknown} weeklyProgress
 * @returns {number | null} 0〜100、未記録は null
 */
function getCurrentProgressFromWeekly(weeklyProgress) {
  const arr = Array.isArray(weeklyProgress) ? weeklyProgress : [];
  const valid = arr.filter((v) => v !== null);
  if (!valid.length) return null;
  return clampProgressPercent(valid[valid.length - 1]);
}

/**
 * @param {string} s
 * @returns {boolean}
 */
function isResponsibleSection(s) {
  return SECTIONS.includes(s);
}

/**
 * Overall Project Schedule の getResponsibleSections と同じ（担当課の配列）。
 *
 * @param {Object} p
 * @returns {string[]}
 */
function getResponsibleSectionsForProject(p) {
  const rs = p.responsibleSections;
  if (Array.isArray(rs) && rs.length > 0) {
    return SECTIONS.filter((k) => rs.includes(k));
  }
  const legacy = p.responsibleDept?.trim();
  if (!legacy) return [];
  const found = new Set();
  for (const k of SECTIONS) {
    if (legacy.includes(k)) found.add(k);
  }
  const parts = legacy.split(/[・,\/\s]+/).map((x) => x.trim()).filter(Boolean);
  for (const part of parts) {
    if (isResponsibleSection(part)) found.add(part);
  }
  return SECTIONS.filter((k) => found.has(k));
}

/**
 * weeklyProgress 末尾側から最大3つの非 null 値（現在・1つ前・2つ前）。
 * 欠けたスロットは current/previous は 0、beforePrevious は null（その場合は先週の差分を 0 とみなします）。
 * @param {unknown} weeklyProgress
 * @returns {{ current: number; previous: number; beforePrevious: number | null }}
 */
function lastThreeWeeklyProgressValues(weeklyProgress) {
  const arr = Array.isArray(weeklyProgress) ? weeklyProgress : [];
  const found = [];
  for (let i = arr.length - 1; i >= 0 && found.length < 3; i--) {
    const v = arr[i];
    if (v != null && v !== "") found.push(clampProgressPercent(v));
  }
  return {
    current: found[0] ?? 0,
    previous: found[1] ?? 0,
    beforePrevious: found.length >= 3 ? found[2] : null,
  };
}

/**
 * @param {import("@neondatabase/serverless").NeonQueryFunction} sql
 * @returns {Promise<Map<string, { current: number; previous: number; beforePrevious: number | null; recordedProgress: number | null }>>}
 */
async function fetchProgressData(sql) {
  const rows = await sql`
    SELECT p AS project
    FROM   app_data,
           jsonb_array_elements(value) AS p
    WHERE  key = 'projects'
  `;

  const map = new Map();
  for (const { project } of rows) {
    if (project.id == null) continue;
    const { current, previous, beforePrevious } = lastThreeWeeklyProgressValues(
      project.weeklyProgress
    );
    const recordedProgress = getCurrentProgressFromWeekly(
      project.weeklyProgress
    );
    map.set(String(project.id), {
      current,
      previous,
      beforePrevious,
      recordedProgress,
    });
  }
  return map;
}

// ──────────────────────────────────────────────
// Risk scoring
// ──────────────────────────────────────────────

/**
 * @param {string | null | undefined} raw
 * @returns {Date | null}
 */
function parseProjectDate(raw) {
  if (raw == null || raw === "") return null;
  const s = typeof raw === "string" ? raw : String(raw);
  const x = new Date(s);
  if (Number.isNaN(x.getTime())) return null;
  x.setHours(0, 0, 0, 0);
  return x;
}

/**
 * @param {Date} a
 * @param {Date} b
 */
function dayDiffFloor(a, b) {
  return Math.floor((a.getTime() - b.getTime()) / 86400000);
}

/**
 * @param {Object} project
 * @param {Date} today
 */
function buildScheduleInputsForRisk(project, today) {
  const startRaw = project.originalStartDate ?? project.startDate;
  const endRaw = project.revisedEndDate ?? project.endDate ?? null;

  const start = parseProjectDate(
    startRaw == null || startRaw === "" ? null : String(startRaw)
  );
  const end = parseProjectDate(
    endRaw == null || endRaw === "" ? null : String(endRaw)
  );

  const elapsedDays = start ? Math.max(0, dayDiffFloor(today, start)) : 0;

  let totalDays = 1;
  if (start && end) {
    totalDays = Math.max(1, dayDiffFloor(end, start));
  } else if (start && !end) {
    totalDays = Math.max(1, elapsedDays + 1);
  }

  const remainingDays = end ? dayDiffFloor(end, today) : 99999;

  const contractAmount = Number(project.contractAmount ?? 0);
  const outsourceCost = Number(
    project.outsourcingAmount ?? project.outsourceCost ?? 0
  );

  return {
    elapsedDays,
    totalDays,
    remainingDays,
    contractAmount,
    outsourceCost,
  };
}

function calculateRisk({
  progress,
  elapsedDays,
  totalDays,
  remainingDays,
  outsourceCost,
  contractAmount,
}) {
  const progressClamped = clampProgressPercent(progress);
  const elapsed = Math.max(0, Number(elapsedDays) || 0);
  const total = Math.max(1, Number(totalDays) || 1);
  const remainingRaw = Number(remainingDays);
  const remaining = Number.isFinite(remainingRaw) ? remainingRaw : 99999;

  let riskScore = 0;
  const riskFactors = [];

  const expectedProgress = (elapsed / total) * 100;
  const remainingWork = 100 - progressClamped;
  const requiredSpeed = remainingWork / Math.max(remaining, 1);
  const denom = Math.max(Number(contractAmount) || 0, 1);
  const outsourceRate = (Number(outsourceCost) || 0) / denom;

  if (progressClamped < expectedProgress - 10) {
    riskScore += 2;
    riskFactors.push("工期に対して進捗が遅れている");
  }

  if (remaining < 7 && progressClamped < 80) {
    riskScore += 2;
    riskFactors.push("残り期間が短く、進捗が不足している");
  }

  if (requiredSpeed > 5) {
    riskScore += 2;
    riskFactors.push(`必要進捗スピードが高い（${requiredSpeed.toFixed(1)}%/日）`);
  } else if (requiredSpeed > 3) {
    riskScore += 1;
    riskFactors.push(`進捗スピードに注意（${requiredSpeed.toFixed(1)}%/日）`);
  }

  if (outsourceRate > 0.6) {
    if (progressClamped < 70) {
      riskScore -= 1;
      riskFactors.push("外注活用により進捗加速中");
    } else {
      riskScore += 2;
      riskFactors.push("終盤で外注依存が高い");
    }
  }

  riskScore = Math.max(riskScore, 0);

  let riskLevel = "";
  let riskColor = "";

  if (riskScore >= 4) {
    riskLevel = "高リスク";
    riskColor = "red";
  } else if (riskScore >= 2) {
    riskLevel = "注意";
    riskColor = "yellow";
  } else {
    riskLevel = "順調";
    riskColor = "green";
  }

  return {
    riskScore,
    riskLevel,
    riskColor,
    expectedProgress,
    requiredSpeed,
    outsourceRate,
    riskFactors,
  };
}


// ──────────────────────────────────────────────
// 生産金額の集計（JS ジョイン）
// ──────────────────────────────────────────────

/**
 * overall.number === progress.id でジョインし、課ごとに生産金額を集計。
 * 今週生産 = allocationSectionN × (現在% - 前回%) / 100（差分はマイナスもそのまま）
 * 現在進捗率が EXCLUDE_FROM_METRICS_PROGRESS_MIN 以上の案件は生産・リスク集計から除外。
 * 生産・リスク・件数の母集団は Overall Project Schedule の「進行中」と同じ:
 * ダッシュボードに週次進捗が1件以上記録されており（recordedProgress が non-null）、
 * かつその値が EXCLUDE_FROM_METRICS_PROGRESS_MIN 未満。status は見ない。
 * 課別 totalCount / highRiskCount / cautionCount は getResponsibleSectionsForProject（担当課）で振り分け。
 * highRiskCount / cautionCount は calculateRisk のレベルを各課に振り分ける。
 *
 * @param {Object[]} projects
 * @param {Map<string, { current: number; previous: number; beforePrevious: number | null; recordedProgress: number | null }>} progressDataMap
 */
function computeSectionStats(projects, progressDataMap) {
  const sectionMap = new Map(
    SECTIONS.map((sec) => [
      sec,
      { production: 0, highRiskCount: 0, cautionCount: 0, totalCount: 0 },
    ])
  );

  let matchedCount = 0;
  let unmatchedOverallCount = 0;
  const matchedProgressIds = new Set();

  /** @type {{ businessNumber: string; name: string; riskScore: number; riskLevel: string; riskFactors: string[] }[]} */
  const riskLogEntries = [];

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  for (const project of projects) {
    const businessNumber = String(project.number ?? "");
    const pair = progressDataMap.get(businessNumber);

    if (pair == null) {
      unmatchedOverallCount++;
      continue;
    }
    matchedCount++;
    matchedProgressIds.add(businessNumber);

    const rec = pair.recordedProgress;
    if (rec == null || rec >= EXCLUDE_FROM_METRICS_PROGRESS_MIN) continue;

    const currentProgress = pair.current;
    const previousProgress = pair.previous;
    const deltaPct = currentProgress - previousProgress;

    const alloc1 = Number(project.allocationSection1 ?? 0);
    const alloc2 = Number(project.allocationSection2 ?? 0);
    const alloc3 = Number(project.allocationSection3 ?? 0);
    sectionMap.get("1課").production += Math.floor((alloc1 * deltaPct) / 100);
    sectionMap.get("2課").production += Math.floor((alloc2 * deltaPct) / 100);
    sectionMap.get("3課").production += Math.floor((alloc3 * deltaPct) / 100);

    let riskResult = null;
    try {
      const sched = buildScheduleInputsForRisk(project, today);
      riskResult = calculateRisk({
        progress: rec,
        elapsedDays: sched.elapsedDays,
        totalDays: sched.totalDays,
        remainingDays: sched.remainingDays,
        outsourceCost: sched.outsourceCost,
        contractAmount: sched.contractAmount,
      });
    } catch {
      riskResult = calculateRisk({
        progress: rec,
        elapsedDays: 0,
        totalDays: 1,
        remainingDays: 99999,
        outsourceCost: 0,
        contractAmount: 0,
      });
    }

    if (riskResult) {
      riskLogEntries.push({
        businessNumber,
        name: String(project.name ?? ""),
        riskScore: riskResult.riskScore,
        riskLevel: riskResult.riskLevel,
        riskFactors: [...riskResult.riskFactors],
      });
    }

    const bumpRiskCounts = (d) => {
      d.totalCount++;
      if (riskResult?.riskLevel === "高リスク") d.highRiskCount++;
      else if (riskResult?.riskLevel === "注意") d.cautionCount++;
    };
    const responsibleSecs = getResponsibleSectionsForProject(project);
    for (const sec of responsibleSecs) {
      if (!sectionMap.has(sec)) continue;
      bumpRiskCounts(sectionMap.get(sec));
    }
  }

  const unmatchedProgressCount = [...progressDataMap.keys()].filter(
    (id) => !matchedProgressIds.has(id)
  ).length;

  return {
    sectionMap,
    matchedCount,
    unmatchedOverallCount,
    unmatchedProgressCount,
    riskLogEntries,
  };
}
/**
 * 先週の進捗差分 × 配分で課ごとに集計する（先週比の分母用）。
 * lastWeekDelta = previousProgress - beforePreviousProgress。
 * beforePrevious が無い案件は lastWeekDelta を 0 とする。
 * 対象案件は computeSectionStats の生産集計と同じ（recordedProgress が non-null かつ 100% 未満）。
 *
 * @param {Object[]} projects
 * @param {Map<string, { current: number; previous: number; beforePrevious: number | null; recordedProgress: number | null }>} progressDataMap
 * @returns {Record<string, number>} section → 先週分の生産金額合計（円）
 */
function computeLastWeekProduction(projects, progressDataMap) {
  const map = { "1課": 0, "2課": 0, "3課": 0 };

  for (const project of projects) {
    const businessNumber = String(project.number ?? "");
    const pair = progressDataMap.get(businessNumber);
    if (pair == null) continue;
    const rec = pair.recordedProgress;
    if (rec == null || rec >= EXCLUDE_FROM_METRICS_PROGRESS_MIN) continue;

    const lastWeekDelta =
      pair != null && pair.beforePrevious != null
        ? pair.previous - pair.beforePrevious
        : 0;

    const alloc1 = Number(project.allocationSection1 ?? 0);
    const alloc2 = Number(project.allocationSection2 ?? 0);
    const alloc3 = Number(project.allocationSection3 ?? 0);

    map["1課"] += Math.floor((alloc1 * lastWeekDelta) / 100);
    map["2課"] += Math.floor((alloc2 * lastWeekDelta) / 100);
    map["3課"] += Math.floor((alloc3 * lastWeekDelta) / 100);
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
  if (last === 0 && current === 0) {
    return { arrow: "→", colorClass: "text-slate-400", label: "先週比 ±0%" };
  }
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
 *   cautionCount: number;
 *   totalCount: number;
 * }} stat
 * @returns {string}
 */
function buildDeptRow(stat) {
  const {
    section,
    production,
    lastProduction,
    highRiskCount,
    cautionCount,
    totalCount,
  } = stat;

  const c = cautionCount ?? 0;
  const isHigh = highRiskCount > 0;
  const isCautionOnly = !isHigh && c > 0;
  const dataRisk = isHigh ? "high" : isCautionOnly ? "medium" : "low";
  const riskBg = isHigh
    ? "bg-red-50/60"
    : isCautionOnly
      ? "bg-amber-50/60"
      : "bg-blue-50/60";
  const badgeColor = isHigh
    ? "bg-red-100 text-red-700"
    : isCautionOnly
      ? "bg-amber-100 text-amber-800"
      : "bg-blue-100 text-blue-700";
  const dotColor = isHigh
    ? "bg-red-500"
    : isCautionOnly
      ? "bg-amber-500"
      : "bg-blue-500";
  const statusLabel = isHigh ? "要警戒" : isCautionOnly ? "注意あり" : "順調";
  const ariaLabel = `リスク: 高リスク${highRiskCount}件、注意${c}件、全${totalCount}件`;

  const { arrow, colorClass, label } = weekOverWeekDisplay(production, lastProduction);
  const amountManYen = toManYen(production);

  return `
            <tr data-risk-level="${dataRisk}" class="transition-colors hover:bg-slate-50/50">
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
                  <span class="text-sm sm:text-base font-bold text-slate-700 leading-tight">
                    <span class="text-lg sm:text-xl text-red-600">${highRiskCount}</span>
                    <span class="text-slate-400 font-normal"> 高</span>
                    <span class="mx-1 text-slate-300">/</span>
                    <span class="text-lg sm:text-xl text-amber-600">${c}</span>
                    <span class="text-slate-400 font-normal"> 注</span>
                    <span class="mx-1 text-slate-300">/</span>
                    <span class="text-lg sm:text-xl">${totalCount}</span>
                    <span class="text-slate-400 font-normal"> 計</span>
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
 *   cautionDept: string;
 * }} opts
 */
async function sendTeamsNotification(opts) {
  const {
    webhookUrl,
    dateLabel,
    imageUrl,
    totalAmount,
    highRiskDept,
    cautionDept,
  } = opts;

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
                { title: "注意該当", value: cautionDept || "なし" },
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

function logTopRisks(riskLogEntries) {
  const sorted = [...riskLogEntries].sort((a, b) => b.riskScore - a.riskScore);
  const top = sorted.slice(0, TOP_RISK_LOG_LIMIT);
  if (top.length === 0) return;
  console.log("\n─── リスクスコア上位案件（参考） ───────────────────");
  for (const e of top) {
    console.log(
      `  [${e.businessNumber}] ${e.name}  score=${e.riskScore}  ${e.riskLevel}`
    );
    if (e.riskFactors.length > 0) {
      console.log(`    要因: ${e.riskFactors.join(" / ")}`);
    }
  }
  console.log("────────────────────────────────────────────\n");
}

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
  const [allProjects, progressDataMap] = await Promise.all([
    fetchAllProjects(OVERALL_PROJECT_SCHEDULE_URL),
    fetchProgressData(sqlProgress),
  ]);

  console.log(`  overall 案件数: ${allProjects.length} 件`);
  console.log(`  progress 案件数: ${progressDataMap.size} 件`);

  // ── JS 側でジョイン → 生産金額を集計 ──
  // 結合キー: overall.number === progress.id（業務番号）
  const {
    sectionMap,
    matchedCount,
    unmatchedOverallCount,
    unmatchedProgressCount,
    riskLogEntries,
  } = computeSectionStats(allProjects, progressDataMap);

  logTopRisks(riskLogEntries);

  console.log(`  結合できた件数:                  ${matchedCount} 件`);
  console.log(
    `  結合できなかった overall 案件数: ${unmatchedOverallCount} 件（progress DB に業務番号なし・生産・リスク・件数の対象外）`
  );
  console.log(`  結合できなかった progress 案件数: ${unmatchedProgressCount} 件`);

  if (matchedCount === 0) {
    console.warn(
      "警告: 結合できた案件が 0 件でした。overall.number と progress.id の値を確認してください。"
    );
  }

  // ── 先週の進捗差分ベース生産金額（先週比の分母） ──
  const lastWeekMap = computeLastWeekProduction(allProjects, progressDataMap);

  // ── サマリー集計 ──
  let totalProductionYen = 0;
  const highRiskSections = [];
  const cautionSections = [];

  for (const sec of SECTIONS) {
    const d = sectionMap.get(sec);
    if (!d) continue;
    totalProductionYen += d.production;
    if (d.highRiskCount > 0) highRiskSections.push(sec);
    if (d.cautionCount > 0) cautionSections.push(sec);
  }

  const totalAmountLabel  = toManYen(totalProductionYen);
  const highRiskDeptLabel =
    highRiskSections.length > 0 ? highRiskSections.join("・") : "なし";
  const cautionDeptLabel =
    cautionSections.length > 0 ? cautionSections.join("・") : "なし";

  // ── 課ごとの生産金額をコンソールに出力 ──
  console.log("\n─── 課ごとの合計生産金額 ───────────────────");
  for (const sec of SECTIONS) {
    const d = sectionMap.get(sec);
    const prod = d?.production ?? 0;
    const total = d?.totalCount ?? 0;
    const hi = d?.highRiskCount ?? 0;
    const ca = d?.cautionCount ?? 0;
    console.log(
      `  ${sec}: ${toManYen(prod)} 万円` +
      `  （件数: ${total}, 高リスク: ${hi}, 注意: ${ca}）`
    );
  }
  console.log(`  総生産金額: ${totalAmountLabel} 万円`);
  console.log("────────────────────────────────────────────\n");

  // ── 行 HTML 生成 ──
  const deptRows = SECTIONS.map((sec) => {
    const d = sectionMap.get(sec) ?? {
      production: 0,
      highRiskCount: 0,
      cautionCount: 0,
      totalCount: 0,
    };
    return buildDeptRow({
      section:        sec,
      production:     d.production,
      lastProduction: lastWeekMap[sec] ?? 0,
      highRiskCount:  d.highRiskCount,
      cautionCount:   d.cautionCount,
      totalCount:     d.totalCount,
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
    .replace("{{CAUTION_DEPT}}",   cautionDeptLabel)
    .replace("{{DEPT_ROWS}}",     deptRows);

  const distDir = join(ROOT, "dist");
  await mkdir(distDir, { recursive: true });

  const reportHtmlPath = join(distDir, "report.html");
  await writeFile(reportHtmlPath, html, "utf-8");
  console.log(`HTML を生成しました: ${reportHtmlPath}`);

  // ── スクリーンショット（一意ファイル名 + report-latest へミラー） ──
  console.log("スクリーンショットを撮影しています...");
  const reportPngFileName = buildTimestampedReportPngFileName(now);
  const reportPngPath = join(distDir, reportPngFileName);
  await takeScreenshot(reportHtmlPath, reportPngPath);
  const reportLatestPath = join(distDir, "report-latest.png");
  await copyFile(reportPngPath, reportLatestPath);
  console.log(`スクリーンショットを保存しました: ${reportPngPath}`);
  console.log(`report-latest.png にコピーしました: ${reportLatestPath}`);

  // ── Teams 通知送信（キャッシュ回避: 一意パス + クエリ） ──
  const pagesBase = normalizePagesBaseUrl(PAGES_BASE_URL);
  const cacheBust = Date.now();
  const imageUrl = `${pagesBase}/${reportPngFileName}?v=${cacheBust}`;

  console.log("reportHtmlPath:", reportHtmlPath);
  console.log("reportPngPath:", reportPngPath);
  console.log("imageUrl:", imageUrl);
  console.log("dateLabel:", dateLabel);

  console.log(`Teams に通知を送信しています... (画像URL: ${imageUrl})`);

  await sendTeamsNotification({
    webhookUrl:   TEAMS_WEBHOOK_URL,
    dateLabel,
    imageUrl,
    totalAmount:  totalAmountLabel,
    highRiskDept: highRiskDeptLabel,
    cautionDept:  cautionDeptLabel,
  });

  console.log("完了しました。");
}

main().catch((err) => {
  console.error("エラーが発生しました:", err);
  process.exit(1);
});
