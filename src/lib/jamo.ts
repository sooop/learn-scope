/* ================================================================
   한글 자모 분해 + 가중 Damerau-Levenshtein 기반 퍼지 매칭
   ================================================================ */

// 초성 19개 (U+3131 기준 호환 자모)
const INITIAL_JAMO: readonly string[] = [
  'ㄱ', 'ㄲ', 'ㄴ', 'ㄷ', 'ㄸ', 'ㄹ', 'ㅁ', 'ㅂ', 'ㅃ',
  'ㅅ', 'ㅆ', 'ㅇ', 'ㅈ', 'ㅉ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ',
];

// 중성 21개
const VOWEL_JAMO: readonly string[] = [
  'ㅏ', 'ㅐ', 'ㅑ', 'ㅒ', 'ㅓ', 'ㅔ', 'ㅕ', 'ㅖ', 'ㅗ', 'ㅘ', 'ㅙ',
  'ㅚ', 'ㅛ', 'ㅜ', 'ㅝ', 'ㅞ', 'ㅟ', 'ㅠ', 'ㅡ', 'ㅢ', 'ㅣ',
];

// 종성 28개 (0번 인덱스 = 받침 없음 → 빈 문자열)
const FINAL_JAMO: readonly string[] = [
  '', 'ㄱ', 'ㄲ', 'ㄳ', 'ㄴ', 'ㄵ', 'ㄶ', 'ㄷ', 'ㄹ', 'ㄺ', 'ㄻ',
  'ㄼ', 'ㄽ', 'ㄾ', 'ㄿ', 'ㅀ', 'ㅁ', 'ㅂ', 'ㅄ', 'ㅅ', 'ㅆ', 'ㅇ',
  'ㅈ', 'ㅊ', 'ㅋ', 'ㅌ', 'ㅍ', 'ㅎ',
];

/**
 * 한글 문자열을 호환 자모 단위로 분해한다.
 * 비한글 문자는 그대로 통과.
 * 예: '심랑' → 'ㅅㅣㅁㄹㅏㅇ', '역사' → 'ㅇㅕㄱㅅㅏ'
 */
export function decomposeHangul(s: string): string {
  let result = '';
  for (const ch of s) {
    const code = ch.charCodeAt(0);
    if (code >= 0xAC00 && code <= 0xD7A3) {
      const offset = code - 0xAC00;
      const finalIdx = offset % 28;
      const vowelIdx = Math.floor(offset / 28) % 21;
      const initialIdx = Math.floor(offset / 28 / 21);
      result += INITIAL_JAMO[initialIdx];
      result += VOWEL_JAMO[vowelIdx];
      if (finalIdx !== 0) result += FINAL_JAMO[finalIdx];
    } else {
      result += ch;
    }
  }
  return result;
}

// 두벌식 인접 키 테이블 (자음만, 양방향)
const ADJACENT_PAIRS: readonly [string, string][] = [
  // 상하 인접
  ['ㅂ', 'ㅁ'], ['ㅈ', 'ㄴ'], ['ㄷ', 'ㅇ'], ['ㄱ', 'ㄹ'], ['ㅅ', 'ㅎ'],
  // 윗줄 좌우
  ['ㅂ', 'ㅈ'], ['ㅈ', 'ㄷ'], ['ㄷ', 'ㄱ'], ['ㄱ', 'ㅅ'],
  // 가운뎃줄 좌우
  ['ㅁ', 'ㄴ'], ['ㄴ', 'ㅇ'], ['ㅇ', 'ㄹ'], ['ㄹ', 'ㅎ'],
  // 추가
  ['ㄷ', 'ㅅ'],
];

// 양방향 Set 구성
const ADJACENT_SET = new Set<string>();
for (const [a, b] of ADJACENT_PAIRS) {
  ADJACENT_SET.add(`${a}|${b}`);
  ADJACENT_SET.add(`${b}|${a}`);
}

function isAdjacent(a: string, b: string): boolean {
  return ADJACENT_SET.has(`${a}|${b}`);
}

/**
 * 가중 OSA (Optimal String Alignment) — 자모열 간 거리 계산.
 * - 동일 자모: 0
 * - 두벌식 인접 키 치환: 0.5
 * - 그 외 치환/삽입/삭제: 1
 * - 인접 전위: 1
 */
function jamoDistance(a: string, b: string): number {
  const m = a.length;
  const n = b.length;

  const dp: number[][] = Array.from({ length: m + 1 }, (_, i) =>
    Array.from({ length: n + 1 }, (_, j) => (i === 0 ? j : j === 0 ? i : 0)),
  );

  for (let i = 1; i <= m; i++) {
    for (let j = 1; j <= n; j++) {
      const ca = a[i - 1];
      const cb = b[j - 1];

      let substCost: number;
      if (ca === cb) {
        substCost = 0;
      } else if (isAdjacent(ca, cb)) {
        substCost = 0.5;
      } else {
        substCost = 1;
      }

      dp[i][j] = Math.min(
        dp[i - 1][j] + 1,             // 삭제
        dp[i][j - 1] + 1,             // 삽입
        dp[i - 1][j - 1] + substCost, // 치환 또는 동일
      );

      // 인접 전위 (transposition)
      if (i >= 2 && j >= 2 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) {
        dp[i][j] = Math.min(dp[i][j], dp[i - 2][j - 2] + 1);
      }
    }
  }

  return dp[m][n];
}

/**
 * 두 문자열 간의 자모 단위 가중 OSA 거리.
 * 내부적으로 decomposeHangul을 적용한다.
 */
export function fuzzyDistance(s1: string, s2: string): number {
  return jamoDistance(decomposeHangul(s1), decomposeHangul(s2));
}

/**
 * 재현율 우선 퍼지 매칭 판정.
 * dist <= 1.5 OR dist / max(자모길이) <= 0.30 이면 true.
 * 길이 0인 문자열은 false.
 */
export function isFuzzyMatch(s1: string, s2: string): boolean {
  if (s1.length === 0 || s2.length === 0) return false;
  const j1 = decomposeHangul(s1);
  const j2 = decomposeHangul(s2);
  if (j1.length === 0 || j2.length === 0) return false;
  const dist = jamoDistance(j1, j2);
  const maxLen = Math.max(j1.length, j2.length);
  return dist < 1.5 || dist / maxLen <= 0.30;
}

/**
 * text와 가장 가까운 라벨을 반환한다.
 * maxDist 이하인 최소 거리 라벨 반환, 없으면 null.
 * text에는 공백/구두점 제거를 적용한다.
 */
export function fuzzyPickLabel(
  text: string,
  labels: readonly string[],
  maxDist = 1.0,
): string | null {
  const normText = text.toLowerCase().replace(/\s+/g, '').replace(/[.,!?]/g, '');
  if (normText.length === 0) return null;

  let bestLabel: string | null = null;
  let bestDist = Infinity;

  for (const label of labels) {
    const d = fuzzyDistance(normText, label);
    if (d <= maxDist && d < bestDist) {
      bestDist = d;
      bestLabel = label;
    }
  }

  return bestLabel;
}
