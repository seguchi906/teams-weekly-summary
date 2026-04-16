/**
 * 週次レポート用: Gemini による提案生成・HTML レンダリング。
 */

import { readFile } from "fs/promises";
import { dirname, join } from "path";
import { fileURLToPath } from "url";
import { GoogleGenerativeAI } from "@google/generative-ai";
import { generateWeeklyFallbackSuggestions } from "./weekly-ai-fallback.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const ROOT = join(__dirname, "..");

/** リスクコンテキストに含める上位件数 */
const RISK_CONTEXT_LIMIT = 15;

/**
 * @typedef {Object} AISuggestion
 * @property {'high'|'medium'|'low'} priority
 * @property {string} priorityLabel
 * @property {string} problem
 * @property {string} action
 */

/**
 * @typedef {Object} WeeklySectionStat
 * @property {string} section
 * @property {number} productionYen
 * @property {string} productionManYen
 * @property {number} lastWeekYen
 * @property {string} weekOverWeekLabel
 * @property {number} highRiskCount
 * @property {number} cautionCount
 * @property {number} totalCount
 */

/**
 * @typedef {Object} RiskLogEntry
 * @property {string} businessNumber
 * @property {string} name
 * @property {number} riskScore
 * @property {string} riskLevel
 * @property {string[]} riskFactors
 */

/**
 * @typedef {Object} WeeklyReportSnapshot
 * @property {string} dateLabel
 * @property {WeeklySectionStat[]} sections
 * @property {number} totalProductionYen
 * @property {string} totalAmountManYen
 * @property {string} highRiskDeptLabel
 * @property {string} cautionDeptLabel
 * @property {RiskLogEntry[]} riskLogEntries
 * @property {number} matchedCount
 * @property {number} unmatchedOverallCount
 * @property {number} unmatchedProgressCount
 */

/**
 * @param {WeeklyReportSnapshot} snapshot
 * @returns {string}
 */
export function buildWeeklyContextMarkdown(snapshot) {
  const lines = [];

  lines.push(`### レポート日付`);
  lines.push(`- ${snapshot.dateLabel}`);
  lines.push("");

  lines.push(`### サマリー`);
  lines.push(`- **合計生産額（表示）**: ${snapshot.totalAmountManYen} 万円（内部円: ${snapshot.totalProductionYen}）`);
  lines.push(`- **高リスク該当課**: ${snapshot.highRiskDeptLabel}`);
  lines.push(`- **注意該当課**: ${snapshot.cautionDeptLabel}`);
  lines.push(`- **結合できた案件数（overall↔progress）**: ${snapshot.matchedCount} 件`);
  lines.push(
    `- **結合不一致**: overall 側 ${snapshot.unmatchedOverallCount} 件 / progress 側 ${snapshot.unmatchedProgressCount} 件`
  );
  lines.push("");

  lines.push(`### 課別（週次生産・先週比ラベル・リスク件数）`);
  for (const s of snapshot.sections) {
    lines.push(`- **${s.section}**`);
    lines.push(
      `  - 生産: ${s.productionManYen} 万円（${s.productionYen} 円） / 先週分: ${s.lastWeekYen} 円 / 先週比: ${s.weekOverWeekLabel}`
    );
    lines.push(
      `  - リスク: 高 ${s.highRiskCount} / 注意 ${s.cautionCount} / 対象 ${s.totalCount}`
    );
  }
  lines.push("");

  const sortedRisks = [...snapshot.riskLogEntries].sort(
    (a, b) => b.riskScore - a.riskScore
  );
  const top = sortedRisks.slice(0, RISK_CONTEXT_LIMIT);

  lines.push(`### リスクスコア上位案件（最大${RISK_CONTEXT_LIMIT}件）`);
  if (top.length === 0) {
    lines.push(`- （対象案件なし、またはログなし）`);
  } else {
    for (const e of top) {
      const factors =
        e.riskFactors.length > 0 ? e.riskFactors.join(" / ") : "—";
      lines.push(
        `- [${e.businessNumber}] ${e.name} ｜ ${e.riskLevel} ｜ score ${e.riskScore} ｜ ${factors}`
      );
    }
  }

  return lines.join("\n");
}

/**
 * @param {string} text
 * @returns {unknown}
 */
export function parseAISuggestionsJson(text) {
  let jsonText = text.trim();
  jsonText = jsonText.replace(/^```json?\s*\n?/gm, "");
  jsonText = jsonText.replace(/\n?```\s*$/gm, "");
  const jsonStart = jsonText.indexOf("[");
  const jsonEnd = jsonText.lastIndexOf("]");
  if (jsonStart === -1 || jsonEnd === -1) {
    throw new Error(
      "JSONが見つかりませんでした。先頭: " + text.substring(0, 400)
    );
  }
  jsonText = jsonText.substring(jsonStart, jsonEnd + 1);
  return JSON.parse(jsonText);
}

/**
 * @param {unknown} parsed
 * @returns {parsed is AISuggestion[]}
 */
function validateSuggestions(parsed) {
  if (!Array.isArray(parsed) || parsed.length !== 3) return false;
  for (const item of parsed) {
    if (!item || typeof item !== "object") return false;
    const p = /** @type {Record<string, unknown>} */ (item);
    if (
      typeof p.priority !== "string" ||
      typeof p.priorityLabel !== "string" ||
      typeof p.problem !== "string" ||
      typeof p.action !== "string"
    ) {
      return false;
    }
  }
  return true;
}

const DEFAULT_MODELS = [
  "models/gemini-2.5-flash",
  "models/gemini-2.0-flash",
];

/**
 * @param {WeeklyReportSnapshot} snapshot
 * @returns {Promise<{ source: 'gemini' | 'fallback'; model: string | null; note?: string; suggestions: AISuggestion[] }>}
 */
export async function generateWeeklyAISuggestions(snapshot) {
  const apiKey = process.env.GEMINI_API_KEY?.trim();

  if (!apiKey) {
    const suggestions = generateWeeklyFallbackSuggestions(snapshot);
    return { source: "fallback", model: null, note: "APIキー未設定", suggestions };
  }

  const context = buildWeeklyContextMarkdown(snapshot);
  let suggestions;
  let source = "gemini";
  let modelName = null;
  let note;

  const promptPath = join(ROOT, "config", "weekly-ai-prompt.md");
  let promptTemplate;
  try {
    promptTemplate = await readFile(promptPath, "utf-8");
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    suggestions = generateWeeklyFallbackSuggestions(snapshot);
    return {
      source: "fallback",
      model: null,
      note: `プロンプト読込失敗: ${msg}`,
      suggestions,
    };
  }

  const prompt = promptTemplate.replace("{{CONTEXT}}", context);
  const genAI = new GoogleGenerativeAI(apiKey);
  const preferred = process.env.GEMINI_MODEL?.trim();
  const modelCandidates = [];
  if (preferred) modelCandidates.push(preferred);
  for (const m of DEFAULT_MODELS) {
    if (!modelCandidates.includes(m)) modelCandidates.push(m);
  }

  let lastError;
  for (const candidate of modelCandidates) {
    try {
      modelName = candidate;
      const model = genAI.getGenerativeModel({ model: modelName });
      const result = await model.generateContent(prompt);
      const response = await result.response;
      const text = response.text();
      const parsed = parseAISuggestionsJson(text);
      if (!validateSuggestions(parsed)) {
        throw new Error(
          `AI提案の形式が不正です（3件・各フィールド必須）。受信: ${JSON.stringify(parsed).slice(0, 300)}`
        );
      }
      suggestions = /** @type {AISuggestion[]} */ (parsed);
      lastError = undefined;
      break;
    } catch (e) {
      lastError = e instanceof Error ? e : new Error(String(e));
      console.warn(`Gemini モデル ${candidate} 失敗:`, lastError.message);
    }
  }

  if (!suggestions) {
    note = lastError?.message ?? "Gemini 生成失敗";
    suggestions = generateWeeklyFallbackSuggestions(snapshot);
    source = "fallback";
    modelName = null;
  }

  if (!validateSuggestions(suggestions)) {
    suggestions = generateWeeklyFallbackSuggestions(snapshot);
    source = "fallback";
    note = note || "検証失敗によりフォールバック";
    modelName = null;
  }

  return {
    source,
    model: source === "gemini" ? modelName : null,
    note,
    suggestions,
  };
}

/**
 * @param {string} s
 */
function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/**
 * @param {string} actionText
 */
function formatActionHtml(actionText) {
  const raw = String(actionText ?? "");
  const lines = raw
    .split("\n")
    .map((l) => l.replace(/^\s*[•・\-*]\s*/, "").trim())
    .filter(Boolean);
  if (lines.length === 0) {
    return `<p class="text-sm text-slate-600 whitespace-pre-wrap">${escapeHtml(raw)}</p>`;
  }
  const items = lines
    .map((line) => `<li class="text-sm text-slate-700">${escapeHtml(line)}</li>`)
    .join("");
  return `<ul class="list-disc pl-5 space-y-1 mt-2">${items}</ul>`;
}

const PRIORITY_STYLES = {
  high: {
    badge:
      "bg-red-100 text-red-800 border border-red-200",
    border: "border-l-4 border-red-400",
  },
  medium: {
    badge:
      "bg-amber-100 text-amber-900 border border-amber-200",
    border: "border-l-4 border-amber-400",
  },
  low: {
    badge:
      "bg-emerald-100 text-emerald-900 border border-emerald-200",
    border: "border-l-4 border-emerald-400",
  },
};

/**
 * @param {{
 *   source: 'gemini' | 'fallback';
 *   model?: string | null;
 *   note?: string;
 *   suggestions: AISuggestion[];
 * }} aiInfo
 * @param {{ layout?: 'embedded' | 'page' }} [options]
 * @returns {string}
 */
export function renderWeeklyAiSuggestionsHtml(aiInfo, options = {}) {
  const layout = options.layout ?? "embedded";

  const suggestions = Array.isArray(aiInfo?.suggestions)
    ? aiInfo.suggestions
    : [];

  if (suggestions.length === 0) {
    if (layout === "page") {
      return `<div class="px-6 py-8" aria-label="AIからの提案">
  <p class="text-sm text-slate-500">提案データを生成できませんでした。</p>
</div>`;
    }
    return `<section class="px-5 py-4 border-t border-slate-100" aria-label="AIからの提案">
  <h2 class="text-sm font-bold text-slate-700 mb-2">AIからの提案</h2>
  <p class="text-xs text-slate-500">提案データを生成できませんでした。</p>
</section>`;
  }

  const sourceLabel =
    aiInfo.source === "gemini"
      ? `Gemini${aiInfo.model ? `（${escapeHtml(aiInfo.model)}）` : ""} で生成`
      : "ルールベースで生成（フォールバック）";

  const noteBlock = aiInfo.note
    ? `<p class="text-xs text-amber-800 bg-amber-50 rounded px-2 py-1 mt-1">${escapeHtml(String(aiInfo.note).slice(0, 200))}${String(aiInfo.note).length > 200 ? "…" : ""}</p>`
    : "";

  const items = suggestions
    .map((s) => {
      const pr = s.priority === "medium" || s.priority === "low" ? s.priority : "high";
      const st = PRIORITY_STYLES[pr] ?? PRIORITY_STYLES.high;
      const label = escapeHtml(s.priorityLabel || "優先度");
      const problem = escapeHtml(s.problem || "");
      const actions = formatActionHtml(s.action || "");
      return `
      <article class="rounded-lg bg-slate-50/80 p-3 ${st.border}">
        <div class="flex flex-wrap items-center gap-2 mb-1">
          <span class="text-xs font-semibold rounded-full px-2 py-0.5 ${st.badge}">${label}</span>
        </div>
        <p class="text-sm font-medium text-slate-800 leading-snug">${problem}</p>
        ${actions}
      </article>`;
    })
    .join("\n");

  if (layout === "page") {
    return `<div class="px-6 py-5 bg-white" aria-label="AIからの提案">
  <p class="text-xs text-slate-500 mb-3">${escapeHtml(sourceLabel)}</p>
  ${noteBlock}
  <div class="space-y-3">${items}</div>
</div>`;
  }

  return `<section class="px-5 py-4 border-t border-slate-100 bg-white" aria-label="AIからの提案">
  <h2 class="text-sm font-bold text-slate-800 mb-1">AIからの提案</h2>
  <p class="text-xs text-slate-500 mb-3">${escapeHtml(sourceLabel)}</p>
  ${noteBlock}
  <div class="space-y-3">${items}</div>
</section>`;
}

/**
 * Teams 用: 各提案の問題文を短く1行ずつ（最大3行）
 * @param {AISuggestion[]} suggestions
 * @returns {string[]}
 */
export function buildAiDigestLines(suggestions) {
  return suggestions.slice(0, 3).map((s, i) => {
    const t = (s.problem || "").replace(/\s+/g, " ").trim();
    const max = 120;
    const short = t.length > max ? `${t.slice(0, max)}…` : t;
    return `${i + 1}. ${short}`;
  });
}
