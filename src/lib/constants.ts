import type { QuestionDistribution } from '../types/analysis';

/* ================================================================
   공통 상수 — 분석 로직 전반에서 사용
   ================================================================ */

export const MAX_ROWS = 100000;

// 응답자 특성 연령 구간
// splitSixties=false: '60대 이상' 단일, true: '60대'/'70대 이상' 분리
export const AGE_KEYS = [
  '19세 이하',
  '20대',
  '30대',
  '40대',
  '50대',
  '60대 이상',
] as const;

export const AGE_KEYS_SPLIT = [
  '19세 이하',
  '20대',
  '30대',
  '40대',
  '50대',
  '60대',
  '70대 이상',
] as const;

// 직업 구간
export const JOB_KEYS = [
  '직장인',
  '자영업',
  '농어업축산임업',
  '주부',
  '학생',
  '기타',
] as const;

// 만족도 분포 초기값 팩토리
// ⚠️ 매 호출마다 새 객체 반환 — 이후 ++ 변이가 있으므로 공유 참조 금지
export function createEmptyQuestionDistribution(): QuestionDistribution {
  return {
    '매우 그렇다': 0,
    그렇다: 0,
    보통: 0,
    '그렇지 않다': 0,
    '매우 그렇지 않다': 0,
  };
}
