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

export interface AttendanceAnalysis {
  subjectResults: SubjectResult[];
  genderDistribution: Record<'여성' | '남성' | '합계', DistributionEntry>;
  courseCountDistribution: Record<'1강좌' | '2강좌' | '3강좌 이상', DistributionEntry>;
  // splitSixties 토글에 따라 키가 '60대 이상' 또는 '60대'+'70대 이상'으로 달라짐
  ageDistribution: Record<string, DistributionEntry>;
  totalStudents: number;
  uniqueStudents: number;
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
}

/* ================================================================
   공통 raw 데이터 타입
   ================================================================ */

export type RawRow = Record<string, unknown>;
