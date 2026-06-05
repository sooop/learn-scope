/* ================================================================
   과목명 퍼지 매칭 — 시트 간 띄어쓰기·오타 차이 흡수
   ================================================================ */

export function normalizeForMatch(s: string): string {
  return s.trim().toLowerCase().replace(/\s+/g, '').replace(/[^\w가-힣]/g, '');
}

function levenshtein(a: string, b: string): number {
  const m = a.length,
    n = b.length;
  const dp: number[][] = Array.from({ length: m + 1 }, (_, i) =>
    Array.from({ length: n + 1 }, (_, j) => (i === 0 ? j : j === 0 ? i : 0)),
  );
  for (let i = 1; i <= m; i++)
    for (let j = 1; j <= n; j++)
      dp[i][j] =
        a[i - 1] === b[j - 1]
          ? dp[i - 1][j - 1]
          : 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
  return dp[m][n];
}

// 1단계: 공백 제거 후 완전 일치 / 2단계: 포함 관계 / 3단계: Levenshtein ≤ 2
export function findCategory(subject: string, categories: Record<string, string>): string {
  const norm = normalizeForMatch(subject);
  for (const [k, v] of Object.entries(categories))
    if (normalizeForMatch(k) === norm) return v;
  for (const [k, v] of Object.entries(categories)) {
    const nk = normalizeForMatch(k);
    if (norm.includes(nk) || nk.includes(norm)) return v;
  }
  let best = '',
    bestDist = 3;
  for (const [k, v] of Object.entries(categories)) {
    const d = levenshtein(norm, normalizeForMatch(k));
    if (d < bestDist) {
      bestDist = d;
      best = v;
    }
  }
  return best;
}
