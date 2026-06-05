import { decomposeHangul, fuzzyDistance, isFuzzyMatch } from '../src/lib/jamo';
import { clusterSubjects } from '../src/lib/subject-cluster';

// ─────────────────────────────────────────────
// 유틸: 결과 출력
// ─────────────────────────────────────────────
function runCase(
  s1: string,
  s2: string,
  expected: boolean,
): void {
  const j1 = decomposeHangul(s1);
  const j2 = decomposeHangul(s2);
  const dist = fuzzyDistance(s1, s2);
  const maxLen = Math.max(j1.length, j2.length);
  const ratio = maxLen > 0 ? dist / maxLen : Infinity;
  const result = isFuzzyMatch(s1, s2);
  const pass = result === expected;
  const tag = pass ? 'PASS' : 'UNEXPECTED';
  console.log(
    `[${tag}] "${s1}" / "${s2}"  분해: "${j1}" / "${j2}"  dist=${dist.toFixed(2)}  maxLen=${maxLen}  ratio=${ratio.toFixed(2)}  match=${result}  (기대: ${expected})`,
  );
}

// ─────────────────────────────────────────────
// 매칭 기대 (isFuzzyMatch = true)
// ─────────────────────────────────────────────
console.log('\n=== 매칭 기대 (true) ===');
runCase('심랑', '실망', true);
// 공백 제거 후 비교
runCase('역삼치현황', '역사및현황', true);
// 한국사 / 한국어 — ratio ≤ 0.30이면 매칭
runCase('한국사', '한국어', true);

// ─────────────────────────────────────────────
// 차단 기대 (isFuzzyMatch = false)
// ─────────────────────────────────────────────
console.log('\n=== 차단 기대 (false) ===');
runCase('요가', '요리', false);
runCase('회화', '회계', false);
runCase('영어', '영화', false);
runCase('국어', '수학', false);
runCase('미술', '음악', false);

// ─────────────────────────────────────────────
// 클러스터링 테스트
// ─────────────────────────────────────────────
console.log('\n=== 클러스터링 테스트 ===');
const names = ['역사 및 현황', '역삼 치 현황', '역사 및 현황', '한글', '한굴', '한글'];
const { canonicalOf, merges } = clusterSubjects(names);

console.log('\ncanonicalOf:');
for (const [k, v] of Object.entries(canonicalOf)) {
  console.log(`  "${k}" → "${v}"`);
}

console.log('\nmerges:');
for (const m of merges) {
  console.log(`  canonical="${m.canonical}", members=[${m.members.map(s => `"${s}"`).join(', ')}]`);
}

// 검증: '역사 및 현황'이 canonical, '역삼 치 현황'이 멤버
const historyMerge = merges.find(m => m.canonical === '역사 및 현황');
const hangeulMerge = merges.find(m => m.canonical === '한글');
console.log('\n검증:');
console.log(
  `  '역사 및 현황' canonical: ${historyMerge ? 'OK' : 'FAIL'} — ${historyMerge ? JSON.stringify(historyMerge) : '없음'}`,
);
console.log(
  `  '한글' canonical (3회>1회): ${hangeulMerge ? 'OK' : 'FAIL'} — ${hangeulMerge ? JSON.stringify(hangeulMerge) : '없음'}`,
);
