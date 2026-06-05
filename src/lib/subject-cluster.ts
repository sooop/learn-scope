/* ================================================================
   과목명 퍼지 클러스터링 — union-find 기반 canonical 표기 결정
   ================================================================ */
import { isFuzzyMatch } from './jamo';
import { normalizeForMatch } from './subject-utils';
import type { SubjectMerge } from '../types/analysis';

// Union-Find 구현
function makeUnionFind(n: number): { parent: number[]; rank: number[] } {
  return {
    parent: Array.from({ length: n }, (_, i) => i),
    rank: new Array(n).fill(0),
  };
}

function find(parent: number[], x: number): number {
  if (parent[x] !== x) parent[x] = find(parent, parent[x]);
  return parent[x];
}

function union(parent: number[], rank: number[], x: number, y: number): void {
  const rx = find(parent, x);
  const ry = find(parent, y);
  if (rx === ry) return;
  if (rank[rx] < rank[ry]) {
    parent[rx] = ry;
  } else if (rank[rx] > rank[ry]) {
    parent[ry] = rx;
  } else {
    parent[ry] = rx;
    rank[rx]++;
  }
}

/**
 * 과목명 배열에서 퍼지 클러스터링을 수행하고
 * canonical 표기 맵과 병합 정보를 반환한다.
 */
export function clusterSubjects(names: string[]): {
  canonicalOf: Record<string, string>;
  merges: SubjectMerge[];
} {
  if (names.length === 0) {
    return { canonicalOf: {}, merges: [] };
  }

  // 1. 빈도 계산: 원본 표기별 등장 횟수
  const freq: Record<string, number> = {};
  for (const name of names) {
    freq[name] = (freq[name] ?? 0) + 1;
  }

  // 첫 등장 순서 기록 (동률 시 사용)
  const firstIndex: Record<string, number> = {};
  for (let i = 0; i < names.length; i++) {
    if (!(names[i] in firstIndex)) {
      firstIndex[names[i]] = i;
    }
  }

  // 2. 중복 제거한 고유 과목명 목록 (원본 표기 유지, 첫 등장 순)
  const unique: string[] = [];
  const seen = new Set<string>();
  for (const name of names) {
    if (!seen.has(name)) {
      seen.add(name);
      unique.push(name);
    }
  }

  const n = unique.length;
  const { parent, rank } = makeUnionFind(n);

  // 3. 모든 쌍 순회 → normalizeForMatch 후 isFuzzyMatch → true면 union
  for (let i = 0; i < n; i++) {
    for (let j = i + 1; j < n; j++) {
      const ni = normalizeForMatch(unique[i]);
      const nj = normalizeForMatch(unique[j]);
      if (isFuzzyMatch(ni, nj)) {
        union(parent, rank, i, j);
      }
    }
  }

  // 4. 클러스터별 멤버 수집
  const clusters: Record<number, string[]> = {};
  for (let i = 0; i < n; i++) {
    const root = find(parent, i);
    if (!clusters[root]) clusters[root] = [];
    clusters[root].push(unique[i]);
  }

  // 5. 대표 표기 결정: 빈도 최다, 동률 시 첫 등장 순
  const canonicalByRoot: Record<number, string> = {};
  for (const [rootStr, members] of Object.entries(clusters)) {
    const root = Number(rootStr);
    let canonical = members[0];
    for (const m of members) {
      const mFreq = freq[m] ?? 0;
      const cFreq = freq[canonical] ?? 0;
      if (
        mFreq > cFreq ||
        (mFreq === cFreq && (firstIndex[m] ?? Infinity) < (firstIndex[canonical] ?? Infinity))
      ) {
        canonical = m;
      }
    }
    canonicalByRoot[root] = canonical;
  }

  // 6. chaining 가드: canonical과 재검증, 불통과 멤버는 독립 클러스터로 분리
  // 분리된 멤버는 자신이 canonical
  const finalCanonicalOf: Record<string, string> = {};

  for (const [rootStr, members] of Object.entries(clusters)) {
    const root = Number(rootStr);
    const canonical = canonicalByRoot[root];
    const canonNorm = normalizeForMatch(canonical);

    const staying: string[] = [];
    const expelled: string[] = [];

    for (const m of members) {
      if (m === canonical) {
        staying.push(m);
      } else {
        const mNorm = normalizeForMatch(m);
        if (isFuzzyMatch(canonNorm, mNorm)) {
          staying.push(m);
        } else {
          expelled.push(m);
        }
      }
    }

    // 남은 멤버 → canonical 할당
    for (const m of staying) {
      finalCanonicalOf[m] = canonical;
    }
    // 분리된 멤버 → 자기 자신이 canonical
    for (const m of expelled) {
      finalCanonicalOf[m] = m;
    }
  }

  // 7. canonicalOf: names의 모든 원본 표기(중복 포함) → canonical
  const canonicalOf: Record<string, string> = {};
  for (const name of names) {
    canonicalOf[name] = finalCanonicalOf[name] ?? name;
  }

  // 8. merges: 2개 이상 묶인 클러스터만
  // finalCanonicalOf 기준으로 클러스터 재집계
  const mergeMap: Record<string, string[]> = {};
  for (const [name, canonical] of Object.entries(finalCanonicalOf)) {
    if (!mergeMap[canonical]) mergeMap[canonical] = [];
    mergeMap[canonical].push(name);
  }

  const merges: SubjectMerge[] = [];
  for (const [canonical, members] of Object.entries(mergeMap)) {
    if (members.length >= 2) {
      merges.push({ canonical, members });
    }
  }

  return { canonicalOf, merges };
}
