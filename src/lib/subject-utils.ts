/* ================================================================
   과목명 퍼지 매칭 — 시트 간 띄어쓰기·오타 차이 흡수
   ================================================================ */
import { fuzzyDistance, decomposeHangul } from './jamo';

// ── bigram Dice 유사도 (5단계 폴백용) ────────────────────────────────
function bigrams(s: string): Set<string> {
  const set = new Set<string>();
  for (let i = 0; i < s.length - 1; i++) set.add(s.slice(i, i + 2));
  return set;
}

function diceSimilarity(a: string, b: string): number {
  const A = bigrams(a);
  const B = bigrams(b);
  if (A.size === 0 || B.size === 0) return 0;
  let inter = 0;
  for (const g of A) if (B.has(g)) inter++;
  return (2 * inter) / (A.size + B.size);
}

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
  if (best) return best;

  // 4단계: 자모 단위 가중 OSA — 전위/인접 키 오타 흡수
  let fuzzyBest = '',
    fuzzyBestDist = Infinity;
  for (const [k, v] of Object.entries(categories)) {
    const d = fuzzyDistance(norm, normalizeForMatch(k));
    const maxLen = Math.max(
      decomposeHangul(norm).length,
      decomposeHangul(normalizeForMatch(k)).length,
    );
    if ((d < 1.5 || (maxLen > 0 && d / maxLen <= 0.30)) && d < fuzzyBestDist) {
      fuzzyBestDist = d;
      fuzzyBest = v;
    }
  }
  if (fuzzyBest) return fuzzyBest;

  // 5단계: bigram Dice — 자모 거리로 구제 불가한 단어 수준 의역·문구 변형 흡수
  // 조건: 1위 ≥ 0.4 && (1위 − 2위) ≥ 0.1 (유사한 후보가 없을 때만 채택)
  if (norm.length >= 5) {
    let diceTop = 0,
      diceSecond = 0,
      diceBestV = '';
    for (const [k, v] of Object.entries(categories)) {
      const score = diceSimilarity(norm, normalizeForMatch(k));
      if (score > diceTop) {
        diceSecond = diceTop;
        diceTop = score;
        diceBestV = v;
      } else if (score > diceSecond) {
        diceSecond = score;
      }
    }
    if (diceTop >= 0.4 && diceTop - diceSecond >= 0.1) return diceBestV;
  }

  return '';
}
