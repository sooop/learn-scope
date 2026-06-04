import type { AttendanceAnalysis, StructuredAttendanceItem } from '../types/analysis';
import { MAX_ROWS, AGE_KEYS, AGE_KEYS_SPLIT } from './constants';

/* ================================================================
   출석 분석  (analyzer.html:315-416)
   ================================================================ */

function calculateAge(birthDate: string | number | unknown): number | null {
  if (!birthDate) return null;
  const dateStr = String(birthDate).trim();
  let parts: string[];

  if (dateStr.includes('.')) parts = dateStr.split('.');
  else if (dateStr.includes('-')) parts = dateStr.split('-');
  else return null;

  const [year, month, day] = parts.map(Number);
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return null;

  const birth = new Date(year, month - 1, day);
  if (
    birth.getFullYear() !== year ||
    birth.getMonth() !== month - 1 ||
    birth.getDate() !== day
  )
    return null;

  const today = new Date();
  if (birth > today) return null;

  let age = today.getFullYear() - birth.getFullYear();
  if (
    today.getMonth() < birth.getMonth() ||
    (today.getMonth() === birth.getMonth() && today.getDate() < birth.getDate())
  )
    age--;
  return age;
}

function getAgeGroup(age: number, splitSixties = false): string {
  if (age <= 19) return '19세 이하';
  if (age < 30) return '20대';
  if (age < 40) return '30대';
  if (age < 50) return '40대';
  if (age < 60) return '50대';
  if (splitSixties) return age < 70 ? '60대' : '70대 이상';
  return '60대 이상';
}

export function analyzeData(
  data: StructuredAttendanceItem[],
  subjectCompletionRates: Record<string, number>,
  splitSixties = false,
): AttendanceAnalysis | null {
  if (!data) return null;
  if (data.length > MAX_ROWS)
    throw new Error(
      `데이터가 너무 큽니다. 최대 ${MAX_ROWS.toLocaleString()}행까지 분석 가능합니다.`,
    );

  // ── 1. 과목별 기본 분석 ──────────────────────────────────────────────
  const subjectStats: Record<string, { 수강인원: number; 출석률합계: number; 수료인원: number }> =
    {};
  data.forEach((item) => {
    if (!subjectStats[item.과목명])
      subjectStats[item.과목명] = { 수강인원: 0, 출석률합계: 0, 수료인원: 0 };
    subjectStats[item.과목명].수강인원++;
    subjectStats[item.과목명].출석률합계 += item.출석률;
    const completionRate = subjectCompletionRates[item.과목명] || 0.7;
    if (item.출석률 >= completionRate) subjectStats[item.과목명].수료인원++;
  });

  const subjectResults = Object.entries(subjectStats).map(([과목명, stats]) => ({
    과목명,
    수강인원: stats.수강인원,
    평균출석률: stats.수강인원 > 0 ? stats.출석률합계 / stats.수강인원 : 0,
    수료인원: stats.수료인원,
    수료율: stats.수강인원 > 0 ? stats.수료인원 / stats.수강인원 : 0,
  }));

  // ── 2. 성별 분포 ──────────────────────────────────────────────────────
  const genderStats = { 여성: 0, 남성: 0 };
  data.forEach((item) => {
    if (item.성별 === '여성' || item.성별 === '여') genderStats.여성++;
    else if (item.성별 === '남성' || item.성별 === '남') genderStats.남성++;
  });
  const totalStudents = data.length;
  const genderDistribution: AttendanceAnalysis['genderDistribution'] = {
    여성: {
      명수: genderStats.여성,
      비율: totalStudents > 0 ? (genderStats.여성 / totalStudents) * 100 : 0,
    },
    남성: {
      명수: genderStats.남성,
      비율: totalStudents > 0 ? (genderStats.남성 / totalStudents) * 100 : 0,
    },
    합계: {
      명수: genderStats.여성 + genderStats.남성,
      비율: totalStudents > 0 ? 100 : 0,
    },
  };

  // ── 3. 수강 강좌수별 분포 ─────────────────────────────────────────────
  const studentCourses: Record<
    string,
    { 이름: string; 성별: string; 생년월일: unknown; 강좌수: number; 강좌목록: string[] }
  > = {};
  data.forEach((item) => {
    const key = `${item.이름}_${item.성별}_${item.생년월일}`;
    if (!studentCourses[key])
      studentCourses[key] = {
        이름: item.이름,
        성별: item.성별,
        생년월일: item.생년월일,
        강좌수: 0,
        강좌목록: [],
      };
    studentCourses[key].강좌수++;
    studentCourses[key].강좌목록.push(item.과목명);
  });

  const courseCountStats: Record<string, number> = { 1: 0, 2: 0, '3이상': 0 };
  Object.values(studentCourses).forEach((s) => {
    if (s.강좌수 === 1) courseCountStats[1]++;
    else if (s.강좌수 === 2) courseCountStats[2]++;
    else courseCountStats['3이상']++;
  });
  const uniqueStudents = Object.keys(studentCourses).length;
  const courseCountDistribution: AttendanceAnalysis['courseCountDistribution'] = {
    '1강좌': {
      명수: courseCountStats[1],
      비율: uniqueStudents > 0 ? (courseCountStats[1] / uniqueStudents) * 100 : 0,
    },
    '2강좌': {
      명수: courseCountStats[2],
      비율: uniqueStudents > 0 ? (courseCountStats[2] / uniqueStudents) * 100 : 0,
    },
    '3강좌 이상': {
      명수: courseCountStats['3이상'],
      비율: uniqueStudents > 0 ? (courseCountStats['3이상'] / uniqueStudents) * 100 : 0,
    },
  };

  // ── 4. 연령 분포 (splitSixties=true 시 60대/70대 이상으로 분리) ───────────
  const ageKeys = splitSixties ? AGE_KEYS_SPLIT : AGE_KEYS;
  const ageStats: Record<string, number> = Object.fromEntries(ageKeys.map((k) => [k, 0]));
  data.forEach((item) => {
    const age = calculateAge(item.생년월일);
    if (age === null) return;
    ageStats[getAgeGroup(age, splitSixties)]++;
  });
  const ageDistribution: AttendanceAnalysis['ageDistribution'] = {};
  Object.entries(ageStats).forEach(([ageGroup, count]) => {
    ageDistribution[ageGroup] = {
      명수: count,
      비율: totalStudents > 0 ? (count / totalStudents) * 100 : 0,
    };
  });

  return {
    subjectResults,
    genderDistribution,
    courseCountDistribution,
    ageDistribution,
    totalStudents,
    uniqueStudents,
  };
}
