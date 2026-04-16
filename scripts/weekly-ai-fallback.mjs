/**
 * 週次レポート用: API 未使用・失敗時のルールベース提案（必ず3件）。
 *
 * @param {import('./weekly-ai-suggestions.mjs').WeeklyReportSnapshot} snapshot
 * @returns {import('./weekly-ai-suggestions.mjs').AISuggestion[]}
 */
export function generateWeeklyFallbackSuggestions(snapshot) {
  const high = buildHighSuggestion(snapshot);
  const medium = buildMediumSuggestion(snapshot);
  const low = buildLowSuggestion(snapshot);
  return [high, medium, low];
}

/**
 * @param {string[]} lines
 */
function formatBulletLines(lines) {
  return lines
    .map((line) => line.trim())
    .filter(Boolean)
    .map((line) => `• ${line}`)
    .join("\n");
}

/**
 * @param {import('./weekly-ai-suggestions.mjs').WeeklyReportSnapshot} snapshot
 */
function buildHighSuggestion(snapshot) {
  const sorted = [...snapshot.riskLogEntries].sort(
    (a, b) => b.riskScore - a.riskScore
  );
  const high = sorted.find((e) => e.riskLevel === "高リスク");
  if (high) {
    const factors =
      high.riskFactors.length > 0 ? high.riskFactors.join(" / ") : "複合要因";
    return {
      priority: "high",
      priorityLabel: "優先度：高",
      problem: `[${high.businessNumber}] ${high.name} が高リスクと判定されており、納期・生産への影響が懸念されます。`,
      action: formatBulletLines([
        `担当窓口と本日中に状況をすり合わせ、要因（${factors}）に対する打ち手を確定する`,
        "必要に応じてリソース再配分・スコープ調整・外注追加を検討し、意思決定を記録する",
        "次回週次までの中間チェック日を設定し、進捗を可視化する",
      ]),
    };
  }

  const caution = sorted.find((e) => e.riskLevel === "注意");
  if (caution) {
    const factors =
      caution.riskFactors.length > 0
        ? caution.riskFactors.join(" / ")
        : "進捗・スケジュール面の注意";
    return {
      priority: "high",
      priorityLabel: "優先度：高",
      problem: `[${caution.businessNumber}] ${caution.name} で注意レベルのサインが出ており、悪化前の手当てが有効です。`,
      action: formatBulletLines([
        `要因（${factors}）を踏まえ、今週中のリカバリプラン（誰が何をするか）を共有する`,
        "ボトルネック作業の前倒し・並列化が可能かを確認する",
        "リスクが高リスクに昇格しないよう、日次または隔日で短時間のスタンドアップを行う",
      ]),
    };
  }

  const deptWithHigh = snapshot.sections.find((s) => s.highRiskCount > 0);
  if (deptWithHigh) {
    return {
      priority: "high",
      priorityLabel: "優先度：高",
      problem: `${deptWithHigh.section}に高リスク案件が${deptWithHigh.highRiskCount}件含まれており、課横断での注視が必要です。`,
      action: formatBulletLines([
        `${deptWithHigh.section}の高リスク案件一覧を確認し、優先順位と担当を明確にする`,
        "他課への影響（連携・依存）がないかを洗い出す",
        "経営・管理層へのエスカレーション基準をチーム内で揃える",
      ]),
    };
  }

  return {
    priority: "high",
    priorityLabel: "優先度：高",
    problem:
      "直ちに炎上している案件は見当たりませんが、週次の定点観測で取りこぼしを防ぎます。",
    action: formatBulletLines([
      "リスクスコアが注意付近の案件を3件ピックアップし、口頭で状況確認する",
      "納期が2週間以内の案件の残作業量とリソースを突き合わせる",
      "先週比で生産が落ちた課の理由（データ・現場の両面）を記録に残す",
    ]),
  };
}

/**
 * @param {import('./weekly-ai-suggestions.mjs').WeeklyReportSnapshot} snapshot
 */
function buildMediumSuggestion(snapshot) {
  const worst = pickWorstWeekOverWeekSection(snapshot.sections);
  if (worst && worst.pct != null && worst.pct < -5) {
    return {
      priority: "medium",
      priorityLabel: "優先度：中",
      problem: `${worst.section}の週次生産が先週比で悪化（${worst.weekOverWeekLabel}）しており、要因の切り分けが有効です。`,
      action: formatBulletLines([
        "進捗率の差分・配分額・対象案件の変化のどれが主因かを切り分ける",
        "一時的な遅れか構造的な遅れかを判断し、来週のリカバリ目標を設定する",
        `${worst.section}の担当者と短時間の振り返りを設定する`,
      ]),
    };
  }

  const cautionHeavy = [...snapshot.sections].sort(
    (a, b) => b.cautionCount - a.cautionCount
  )[0];
  if (cautionHeavy && cautionHeavy.cautionCount >= 2) {
    return {
      priority: "medium",
      priorityLabel: "優先度：中",
      problem: `${cautionHeavy.section}に注意レベル案件が${cautionHeavy.cautionCount}件寄っており、負荷の偏りに注意が必要です。`,
      action: formatBulletLines([
        "注意案件の共通要因（工程・外注・入力遅れ等）を洗い出す",
        "他課からのサポートやタスクの再配分が可能か検討する",
        "週内のミニレビューで優先度を再確認する",
      ]),
    };
  }

  return {
    priority: "medium",
    priorityLabel: "優先度：中",
    problem:
      "全体として大きな偏りは限定的ですが、先週比・リスク件数の変化を継続監視します。",
    action: formatBulletLines([
      "各課の先週比とリスク件数を先週レポートと並べて比較する",
      "生産が伸びた課の要因（良い実践）を他課に展開できるか検討する",
      "ダッシュボードの進捗入力の抜け漏れがないかを確認する",
    ]),
  };
}

/**
 * @param {import('./weekly-ai-suggestions.mjs').WeeklySectionStat[]} sections
 * @returns {{ section: string; weekOverWeekLabel: string; pct: number | null } | null}
 */
function pickWorstWeekOverWeekSection(sections) {
  /** @type {{ section: string; weekOverWeekLabel: string; pct: number | null }[]} */
  const scored = [];
  for (const s of sections) {
    const pct = weekOverWeekPercent(s.productionYen, s.lastWeekYen);
    if (pct != null) {
      scored.push({
        section: s.section,
        weekOverWeekLabel: s.weekOverWeekLabel,
        pct,
      });
    }
  }
  if (scored.length === 0) return null;
  scored.sort((a, b) => (a.pct ?? 0) - (b.pct ?? 0));
  return scored[0];
}

/**
 * @param {number} current
 * @param {number} last
 * @returns {number | null}
 */
function weekOverWeekPercent(current, last) {
  if (last === 0 && current === 0) return 0;
  if (last === 0) return null;
  return ((current - last) / last) * 100;
}

/**
 * @param {import('./weekly-ai-suggestions.mjs').WeeklyReportSnapshot} snapshot
 */
function buildLowSuggestion(snapshot) {
  const uo = snapshot.unmatchedOverallCount ?? 0;
  const up = snapshot.unmatchedProgressCount ?? 0;
  if (uo > 0 || up > 0) {
    return {
      priority: "low",
      priorityLabel: "優先度：低",
      problem: `データ結合の観点で、overall と progress の番号不一致が発生しています（overall側 ${uo} 件 / progress側 ${up} 件）。`,
      action: formatBulletLines([
        "業務番号の表記ゆれ（ゼロ埋め・ハイフン）をルール化し、マスタを整備する",
        "新規案件登録時のチェックリストに「両系統への反映」を追加する",
        "不一致リストを週次で減らす目標を持つ",
      ]),
    };
  }

  return {
    priority: "low",
    priorityLabel: "優先度：低",
    problem:
      "短期的な火種は限定的です。次週以降の安定運用に向けた仕組み面の改善余地があります。",
    action: formatBulletLines([
      "週次レビューのアジェンダ（数字・リスク・決定事項）をテンプレ化する",
      "リスクスコアの閾値とエスカレーションルールを文書化する",
      "レポート自動生成のログを週次で1回確認し、異常検知に使う",
    ]),
  };
}
