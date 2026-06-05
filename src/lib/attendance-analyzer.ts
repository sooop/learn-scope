import type { AttendanceAnalysis, StructuredAttendanceItem } from '../types/analysis';
import { MAX_ROWS, AGE_KEYS, AGE_KEYS_SPLIT } from './constants';
import { DiagnosticCollector } from './diagnostics';

/* ================================================================
   출석 분석  (analyzer.html:315-416)
   ================================================================ */

function ageFromParts(year: number, month: number, day: number): number | null {
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) return null;
  const birth = new Date(year, month - 1, day);
  if (birth.getFullYear() !== year || birth.getMonth() !== month - 1 || birth.getDate() !== day)
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

function calculateAge(birthDate: string | number | unknown): number | null {
  if (!birthDate) return null;

  // ── 숫자 입력: 엑셀 시퀀스 또는 yyyymmdd 정수 ──────────────────────────
  if (typeof birthDate === 'number') {
    if (birthDate >= 19000101) {
      // yyyymmdd 정수 (예: 19780703)
      return ageFromParts(
        Math.floor(birthDate / 10000),
        Math.floor((birthDate % 10000) / 100),
        birthDate % 100,
      );
    }
    if (birthDate > 0) {
      // 엑셀 날짜 시퀀스 (serial 25569 = 1970-01-01)
      const d = new Date(Math.round((birthDate - 25569) * 86400000));
      return ageFromParts(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate());
    }
    return null;
  }

  // ── 문자열 입력 ──────────────────────────────────────────────────────────
  const s = String(birthDate).trim();
  let parts: string[];
  if (s.includes('.')) parts = s.split('.');
  else if (s.includes('-')) parts = s.split('-');
  else if (/^\d{8}$/.test(s)) parts = [s.slice(0, 4), s.slice(4, 6), s.slice(6, 8)]; // yyyymmdd 문자열
  else return null;

  const [y, m, d] = parts.map(Number);
  return ageFromParts(y, m, d);
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

// C1: trim+소문자 정규화 — 표기 변형(남자/男/M) 흡수
function normalizeGender(raw: string): '남성' | '여성' | null {
  const s = raw.trim().toLowerCase().replace(/\s+/g, '');
  if (s.includes('남') || s.includes('male') || s === 'm') return '남성';
  if (s.includes('여') || s.includes('female') || s === 'f') return '여성';
  return null;
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

  const dc = new DiagnosticCollector();

  // ── 1. 과목별 기본 분석 ──────────────────────────────────────────────
  const subjectStats: Record<string, { 수강인원: number; 출석률합계: number; 수료인원: number }> =
    {};
  data.forEach((item) => {
    if (!subjectStats[item.과목명])
      subjectStats[item.과목명] = { 수강인원: 0, 출석률합계: 0, 수료인원: 0 };
    subjectStats[item.과목명].수강인원++;
    subjectStats[item.과목명].출석률합계 += item.출석률;
    const completionRate = subjectCompletionRates[item.과목명] ?? 0.7;
    // C4: epsilon 비교로 부동소수점 경계 오분류 방지
    if (item.출석률 >= completionRate - 1e-9) subjectStats[item.과목명].수료인원++;
  });

  const subjectResults = Object.entries(subjectStats).map(([과목명, stats]) => ({
    과목명,
    수강인원: stats.수강인원,
    평균출석률: stats.수강인원 > 0 ? stats.출석률합계 / stats.수강인원 : 0,
    수료인원: stats.수료인원,
    수료율: stats.수강인원 > 0 ? stats.수료인원 / stats.수강인원 : 0,
  }));

  // ── 2. 성별 분포 ──────────────────────────────────────────────────────
  // C1: 관대한 정규화 + 미매칭은 '기타/미상' 버킷으로 표면화
  const genderStats = { 여성: 0, 남성: 0, '기타/미상': 0 };
  data.forEach((item) => {
    const normalized = normalizeGender(item.성별 || '');
    if (normalized) {
      genderStats[normalized]++;
    } else {
      genderStats['기타/미상']++;
      if (item.성별) {
        dc.add(
          'C1-gender-unmatched',
          'warning',
          'normalize-failed',
          '인식할 수 없는 성별 값이 기타/미상으로 분류됨',
          String(item.성별),
        );
      }
    }
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
  };
  // '기타/미상' 버킷은 해당 인원이 있을 때만 추가
  if (genderStats['기타/미상'] > 0) {
    genderDistribution['기타/미상'] = {
      명수: genderStats['기타/미상'],
      비율: totalStudents > 0 ? (genderStats['기타/미상'] / totalStudents) * 100 : 0,
    };
  }
  const genderTotal = genderStats.여성 + genderStats.남성 + genderStats['기타/미상'];
  genderDistribution['합계'] = {
    명수: genderTotal,
    비율: totalStudents > 0 ? (genderTotal / totalStudents) * 100 : 0,
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
  let unknownAge = 0;
  data.forEach((item) => {
    const age = calculateAge(item.생년월일);
    if (age === null) {
      unknownAge++;
      return;
    }
    ageStats[getAgeGroup(age, splitSixties)]++;
  });
  const ageDistribution: AttendanceAnalysis['ageDistribution'] = {};
  Object.entries(ageStats).forEach(([ageGroup, count]) => {
    ageDistribution[ageGroup] = {
      명수: count,
      비율: totalStudents > 0 ? (count / totalStudents) * 100 : 0,
    };
  });
  if (unknownAge > 0) {
    ageDistribution['생년월일 미확인'] = {
      명수: unknownAge,
      비율: totalStudents > 0 ? (unknownAge / totalStudents) * 100 : 0,
    };
  }

  // ── 5. 버킷 합 교차검증 ───────────────────────────────────────────────
  const genderBucketSum = Object.entries(genderDistribution)
    .filter(([k]) => k !== '합계')
    .reduce((s, [, v]) => s + v.명수, 0);
  if (genderBucketSum !== totalStudents) {
    dc.add(
      'INTEGRITY-gender-sum',
      'warning',
      'bucket-mismatch',
      `성별 분포 버킷 합(${genderBucketSum})이 전체 수강인원(${totalStudents})과 불일치`,
    );
  }

  const ageBucketSum = Object.values(ageDistribution).reduce((s, v) => s + v.명수, 0);
  if (ageBucketSum !== totalStudents) {
    dc.add(
      'INTEGRITY-age-sum',
      'warning',
      'bucket-mismatch',
      `연령 분포 버킷 합(${ageBucketSum})이 전체 수강인원(${totalStudents})과 불일치`,
    );
  }

  return {
    subjectResults,
    genderDistribution,
    courseCountDistribution,
    ageDistribution,
    totalStudents,
    uniqueStudents,
    diagnostics: dc.build(),
  };
}
