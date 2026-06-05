/* ================================================================
   출석 분석 타입  (porting-guide §4-4, §5-4)
   ================================================================ */

export interface StructuredAttendanceItem {
  연번: unknown;
  과목명: string;
  이름: string;
  성별: string;
  생년월일: string | number;
  수업일: number;
  출석일: number;
  출석률: number; // 0~1 비율 (0.85 = 85%)
  수료여부: unknown;
}

// structureAttendanceData 반환 타입 (diagnostics 포함)
export interface StructuredAttendanceResult {
  items: StructuredAttendanceItem[];
  diagnostics: Diagnostic[];
}

export interface SubjectResult {
  과목명: string;
  수강인원: number;
  평균출석률: number; // 0~1 비율 (표시 시 ×100 필요)
  수료인원: number;
  수료율: number; // 0~1 비율 (표시 시 ×100 필요)
}

export interface DistributionEntry {
  명수: number;
  비율: number; // 0~100 퍼센트값 (표시 시 추가 변환 불필요)
}

// ── 진단 타입 ─────────────────────────────────────────────────────
export type DiagnosticSeverity = 'warning' | 'info';
export type DiagnosticCategory =
  | 'skipped-row'
  | 'normalize-failed'
  | 'rate-format'
  | 'column-fallback'
  | 'integrity'
  | 'bucket-mismatch'
  | 'value-coercion';

export interface Diagnostic {
  severity: DiagnosticSeverity;
  category: DiagnosticCategory;
  code: string;        // 안정 식별자, 예: 'B1-rate-format'
  message: string;     // 한국어 사용자 메시지
  count: number;
  samples?: string[];  // 원문/값 샘플 (최대 5개)
}

export interface AttendanceAnalysis {
  subjectResults: SubjectResult[];
  // 성별 키: '여성'|'남성'|'합계' + 미상 존재 시 '기타/미상' 동적 추가
  genderDistribution: Record<string, DistributionEntry>;
  courseCountDistribution: Record<'1강좌' | '2강좌' | '3강좌 이상', DistributionEntry>;
  // splitSixties 토글에 따라 키가 '60대 이상' 또는 '60대'+'70대 이상'으로 달라짐
  ageDistribution: Record<string, DistributionEntry>;
  totalStudents: number;
  uniqueStudents: number;
  diagnostics: Diagnostic[];
}

/* ================================================================
   만족도 분석 타입  (porting-guide §6-4)
   ================================================================ */

export type SatisfactionScoreValue = string | number; // toFixed(2) 문자열 또는 숫자 0

export type SatisfactionScores = Record<string, SatisfactionScoreValue>;

export interface SubjectSatisfaction {
  subject: string;
  구분?: string;
  scores: SatisfactionScores;
}

export interface QuestionDistribution {
  '매우 그렇다': number;
  그렇다: number;
  보통: number;
  '그렇지 않다': number;
  '매우 그렇지 않다': number;
}

export interface RespondentCharacteristic {
  subject: string;
  gender: Record<'남성' | '여성', number>;
  region: Record<string, number>; // 14개 지역 + '기타지역'
  age: Record<'19세 이하' | '20대' | '30대' | '40대' | '50대' | '60대 이상', number>;
  job: Record<'직장인' | '자영업' | '농어업축산임업' | '주부' | '학생' | '기타', number>;
}

export interface SatisfactionAnalysis {
  respondentCharacteristics: Record<string, RespondentCharacteristic>;
  satisfactionDistribution: Record<
    string,
    { subject: string; questions: Record<string, QuestionDistribution> }
  >;
  satisfactionAverages: Record<string, SubjectSatisfaction>;
  subjectCategories: Record<string, string>; // 과목명 → 구분
  totalSubjects: number;
  totalResponses: number;
  diagnostics: Diagnostic[];
  subjectMerges?: SubjectMerge[];
}

/* ================================================================
   퍼지 클러스터링 타입
   ================================================================ */

export interface SubjectMerge {
  canonical: string;
  members: string[];
}

/* ================================================================
   공통 raw 데이터 타입
   ================================================================ */

export type RawRow = Record<string, unknown>;
