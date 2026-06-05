import type { RawRow, SatisfactionAnalysis, SubjectMerge } from '../types/analysis';
import { MAX_ROWS, createEmptyQuestionDistribution } from './constants';
import { findCategory } from './subject-utils';
import { DiagnosticCollector } from './diagnostics';
import { clusterSubjects } from './subject-cluster';
import { fuzzyPickLabel } from './jamo';

/* ================================================================
   만족도 분석  (analyzer.html:421-581)
   ================================================================ */

const regionMapping: Record<string, string> = {
  통진읍: '통진읍',
  고촌읍: '고촌읍',
  양촌읍: '양촌읍',
  월곶면: '월곶면',
  하성면: '하성면',
  대곶면: '대곶면',
  김포본동: '김포본동',
  장기본동: '장기본동',
  사우동: '사우동',
  풍무동: '풍무동',
  장기동: '장기동',
  구래동: '구래동',
  마산동: '마산동',
  운양동: '운양동',
};

export { regionMapping };

const satisfactionScoreMapping: Record<string, number> = {
  '매우 그렇다': 5,
  그렇다: 4,
  보통: 3,
  '그렇지 않다': 2,
  '매우 그렇지 않다': 1,
};

type NormalizeCategory = 'gender' | 'region' | 'age' | 'job' | 'satisfaction';

// ⚠️ satisfaction 분기의 if 판정 순서 절대 변경 금지
// D3: 실패 시 '' 반환 (기존 String(text) → 호출부에서 무효응답으로 진단 처리)
function normalizeText(text: unknown, category: NormalizeCategory): string {
  if (!text) return '';
  const cleanText = text
    .toString()
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[.,!?]/g, '');

  switch (category) {
    case 'gender':
      if (cleanText.includes('남') || cleanText.includes('male')) return '남성';
      if (cleanText.includes('여') || cleanText.includes('female')) return '여성';
      {
        const genderFuzzy = fuzzyPickLabel(cleanText, ['남성', '여성'], 1.0);
        if (genderFuzzy) return genderFuzzy;
      }
      break;
    case 'region': {
      const regions = Object.keys(regionMapping);
      for (const region of regions) {
        if (cleanText.includes(region.replace(/\s+/g, ''))) return region;
      }
      {
        const regionFuzzy = fuzzyPickLabel(cleanText, regions, 1.0);
        if (regionFuzzy) return regionFuzzy;
      }
      return '기타지역';
    }
    case 'age':
      if (
        cleanText.includes('19') ||
        cleanText.includes('10대') ||
        cleanText.includes('미성년')
      )
        return '19세 이하';
      if (cleanText.includes('20')) return '20대';
      if (cleanText.includes('30')) return '30대';
      if (cleanText.includes('40')) return '40대';
      if (cleanText.includes('50')) return '50대';
      if (
        cleanText.includes('60') ||
        cleanText.includes('70') ||
        cleanText.includes('80') ||
        cleanText.includes('90')
      )
        return '60대 이상';
      break;
    case 'job':
      if (
        cleanText.includes('직장') ||
        cleanText.includes('회사') ||
        cleanText.includes('사무')
      )
        return '직장인';
      if (cleanText.includes('자영')) return '자영업';
      if (
        cleanText.includes('농업') ||
        cleanText.includes('어업') ||
        cleanText.includes('축산') ||
        cleanText.includes('임업')
      )
        return '농어업축산임업';
      if (cleanText.includes('주부') || cleanText.includes('가정')) return '주부';
      if (cleanText.includes('학생') || cleanText.includes('대학')) return '학생';
      {
        const jobFuzzy = fuzzyPickLabel(cleanText, ['직장인', '자영업', '주부', '학생'], 1.0);
        if (jobFuzzy) return jobFuzzy;
      }
      return '기타';
    case 'satisfaction':
      // ⚠️ 순서 의존: 매우그렇다 → 매우그렇지않다 → 그렇지않다 → 그렇다 → 보통
      if (cleanText.includes('매우') && cleanText.includes('그렇다')) return '매우 그렇다';
      if (
        cleanText.includes('매우') &&
        (cleanText.includes('그렇지않다') ||
          cleanText.includes('안다') ||
          cleanText.includes('않다'))
      )
        return '매우 그렇지 않다';
      if (
        cleanText.includes('그렇지않다') ||
        (cleanText.includes('그렇지') && cleanText.includes('않다'))
      )
        return '그렇지 않다';
      if (cleanText.includes('그렇다')) return '그렇다';
      if (cleanText.includes('보통')) return '보통';
      // D4: 숫자 변형 매칭 — text.toString() 사용 (cleanText는 '.' 제거로 '5.0'→'50' 오염됨)
      {
        const rawNum = String(text).trim();
        const numMatch = /^(\d+(?:\.\d+)?)/.exec(rawNum);
        if (numMatch) {
          const numVal = parseFloat(numMatch[1]);
          if (Number.isFinite(numVal)) {
            const rounded = Math.round(numVal);
            if (rounded === 5) return '매우 그렇다';
            if (rounded === 4) return '그렇다';
            if (rounded === 3) return '보통';
            if (rounded === 2) return '그렇지 않다';
            if (rounded === 1) return '매우 그렇지 않다';
          }
        }
      }
      {
        const satFuzzy = fuzzyPickLabel(cleanText, ['매우그렇다', '그렇다', '보통', '그렇지않다', '매우그렇지않다'], 1.0);
        if (satFuzzy) {
          const satMap: Record<string, string> = {
            '매우그렇다': '매우 그렇다',
            '그렇다': '그렇다',
            '보통': '보통',
            '그렇지않다': '그렇지 않다',
            '매우그렇지않다': '매우 그렇지 않다',
          };
          return satMap[satFuzzy] ?? satFuzzy;
        }
      }
      break;
  }
  // D3: 실패 시 '' 반환 (기존 String(text)에서 변경 — 호출부에서 무효응답 진단)
  return '';
}

const questionPatterns = [
  '2-1.*적극',
  '2-2.*전문',
  '2-3.*충실',
  '2-4.*시간',
  '2-5.*활용',
  '2-6.*유익',
  '2-7.*교재',
  '2-8.*(시설|시섦|교육지원)',
  '2-9.*만족',
];

// E2: export해서 excel-exporter에서 표준 순서 참조에 사용
export const satisfactionQuestions = [
  '강사 적극성',
  '강사 전문성',
  '내용 충실성',
  '시간대 적정성',
  '수업의 활용도',
  '수업 유익성',
  '교재 및 자료 충분성',
  '강의 시설 만족도',
  '체감 만족도',
];

export function analyzeSatisfactionData(
  rawData: RawRow[],
  externalCategories: Record<string, string> = {},
): SatisfactionAnalysis {
  if (!rawData || rawData.length === 0) throw new Error('분석할 데이터가 없습니다.');
  if (rawData.length > MAX_ROWS)
    throw new Error(
      `데이터가 너무 큽니다. 최대 ${MAX_ROWS.toLocaleString()}행까지 분석 가능합니다.`,
    );

  const dc = new DiagnosticCollector();

  const headers = Object.keys(rawData[0]);
  const dataRows = rawData;

  const columnIndexes = {
    serialNumber: headers.findIndex((h) => h.includes('연번')),
    subject: headers.findIndex((h) => h.includes('과목명')),
    category: headers.findIndex((h) => h.includes('구분')),
    gender: headers.findIndex((h) => /성별/i.test(h)),
    region: headers.findIndex((h) => /지역|거주|주소/i.test(h)),
    age: headers.findIndex((h) => /연령|나이/i.test(h)),
    job: headers.findIndex((h) => /직업|직종/i.test(h)),
  };

  // D5: 칼럼 자동탐지 실패 시 고정 인덱스 폴백 + 경고
  if (columnIndexes.gender === -1) {
    columnIndexes.gender = 4;
    dc.add(
      'D5-column-fallback-gender',
      'warning',
      'column-fallback',
      `성별 칼럼 자동탐지 실패 → 5번째 칼럼('${headers[4] ?? '범위 밖'}') 사용`,
    );
  }
  if (columnIndexes.region === -1) {
    columnIndexes.region = 5;
    dc.add(
      'D5-column-fallback-region',
      'warning',
      'column-fallback',
      `지역 칼럼 자동탐지 실패 → 6번째 칼럼('${headers[5] ?? '범위 밖'}') 사용`,
    );
  }
  if (columnIndexes.age === -1) {
    columnIndexes.age = 6;
    dc.add(
      'D5-column-fallback-age',
      'warning',
      'column-fallback',
      `연령 칼럼 자동탐지 실패 → 7번째 칼럼('${headers[6] ?? '범위 밖'}') 사용`,
    );
  }
  if (columnIndexes.job === -1) {
    columnIndexes.job = 7;
    dc.add(
      'D5-column-fallback-job',
      'warning',
      'column-fallback',
      `직업 칼럼 자동탐지 실패 → 8번째 칼럼('${headers[7] ?? '범위 밖'}') 사용`,
    );
  }

  const satisfactionColumnIndexes = questionPatterns.map((pattern) => {
    const regex = new RegExp(pattern, 'i');
    return headers.findIndex((header) => header && regex.test(header.toString()));
  });

  // ── 과목별 그룹핑 (병합 셀 forward-fill) ──────────────────────────────
  const subjectGroups: Record<string, RawRow[]> = {};
  const categoryForSubject: Record<string, string> = {};
  let currentSubject: string | null = null;
  let currentCategory: string | null = null;

  dataRows.forEach((row) => {
    // 구분 칼럼 forward-fill
    if (columnIndexes.category !== -1) {
      const catVal = row[headers[columnIndexes.category]];
      if (catVal && catVal.toString().trim() !== '')
        currentCategory = catVal.toString().trim();
    }

    let subject = row[headers[columnIndexes.subject]];
    if (subject && subject.toString().trim() !== '') {
      const newSubject = subject.toString().trim();

      // D2: 새 과목 행의 구분 칼럼이 공란이면 직전 구분 상속 — info 진단
      if (columnIndexes.category !== -1 && currentCategory !== null && newSubject !== (currentSubject ?? '')) {
        const rowCatVal = row[headers[columnIndexes.category]];
        const catIsBlank = !rowCatVal || rowCatVal.toString().trim() === '';
        if (catIsBlank) {
          dc.add(
            'D2-category-inherited',
            'info',
            'column-fallback',
            `구분 없는 과목에 '${currentCategory}' 상속됨`,
            newSubject,
          );
        }
      }

      currentSubject = newSubject;
      if (currentCategory) categoryForSubject[currentSubject] = currentCategory;
    } else if (currentSubject) {
      subject = currentSubject;
    }

    // D1: currentSubject === null이면 과목 귀속 불가 → 행 제외 + 경고
    if (!subject) {
      const serial =
        columnIndexes.serialNumber !== -1
          ? String(row[headers[columnIndexes.serialNumber]] ?? '')
          : '';
      dc.add(
        'D1-leading-rows-dropped',
        'warning',
        'skipped-row',
        '과목명이 없어 제외된 응답 행 (파일 최상단 과목명 누락)',
        serial || undefined,
      );
      return;
    }

    const subjectKey = subject.toString().trim();
    if (!subjectGroups[subjectKey]) subjectGroups[subjectKey] = [];
    subjectGroups[subjectKey].push(row);
  });

  // 과목명 canonical화 — 오타·띄어쓰기 변형을 하나의 키로 통합
  const rawSubjectKeys = Object.keys(subjectGroups);
  const { canonicalOf, merges: subjectMerges }: { canonicalOf: Record<string, string>; merges: SubjectMerge[] } =
    clusterSubjects(rawSubjectKeys);

  const canonicalSubjectGroups: Record<string, RawRow[]> = {};
  const canonicalCategoryForSubject: Record<string, string> = {};
  for (const [rawKey, rows] of Object.entries(subjectGroups)) {
    const canonical = canonicalOf[rawKey] ?? rawKey;
    if (!canonicalSubjectGroups[canonical]) canonicalSubjectGroups[canonical] = [];
    canonicalSubjectGroups[canonical].push(...rows);
    if (categoryForSubject[rawKey]) canonicalCategoryForSubject[canonical] = categoryForSubject[rawKey];
  }

  const respondentCharacteristics: SatisfactionAnalysis['respondentCharacteristics'] = {};
  const satisfactionDistribution: SatisfactionAnalysis['satisfactionDistribution'] = {};
  const satisfactionAverages: SatisfactionAnalysis['satisfactionAverages'] = {};

  for (const [subject, rows] of Object.entries(canonicalSubjectGroups)) {
    const rc = {
      subject,
      gender: { 남성: 0, 여성: 0 } as Record<'남성' | '여성', number>,
      region: {
        ...Object.fromEntries(Object.keys(regionMapping).map((k) => [k, 0])),
        기타지역: 0,
      } as Record<string, number>,
      age: {
        '19세 이하': 0,
        '20대': 0,
        '30대': 0,
        '40대': 0,
        '50대': 0,
        '60대 이상': 0,
      } as Record<'19세 이하' | '20대' | '30대' | '40대' | '50대' | '60대 이상', number>,
      job: {
        직장인: 0,
        자영업: 0,
        농어업축산임업: 0,
        주부: 0,
        학생: 0,
        기타: 0,
      } as Record<'직장인' | '자영업' | '농어업축산임업' | '주부' | '학생' | '기타', number>,
    };
    respondentCharacteristics[subject] = rc;
    satisfactionDistribution[subject] = { subject, questions: {} };
    const 구분 = canonicalCategoryForSubject[subject] || findCategory(subject, externalCategories);
    satisfactionAverages[subject] = { subject, 구분, scores: {} };

    rows.forEach((row) => {
      // D3: normalizeText 실패 시 '' 반환 → 무효응답 진단
      const rawGender = row[headers[columnIndexes.gender]];
      const gender = normalizeText(rawGender, 'gender');
      if (gender && rc.gender[gender as '남성' | '여성'] !== undefined) {
        rc.gender[gender as '남성' | '여성']++;
      } else if (rawGender && !gender) {
        dc.add(
          'D3-gender-invalid',
          'warning',
          'normalize-failed',
          `성별 값을 인식할 수 없어 집계에서 제외됨`,
          String(rawGender),
        );
      }

      // region: 항상 '기타지역' 폴백이 있으므로 진단 불필요
      const region = normalizeText(row[headers[columnIndexes.region]], 'region');
      if (region && rc.region[region] !== undefined) rc.region[region]++;

      const rawAge = row[headers[columnIndexes.age]];
      const age = normalizeText(rawAge, 'age');
      if (age && rc.age[age as '19세 이하'] !== undefined) {
        rc.age[age as keyof typeof rc.age]++;
      } else if (rawAge && !age) {
        dc.add(
          'D3-age-invalid',
          'warning',
          'normalize-failed',
          `연령 값을 인식할 수 없어 집계에서 제외됨`,
          String(rawAge),
        );
      }

      // job: '기타' 폴백이 있으므로 항상 집계됨
      const job = normalizeText(row[headers[columnIndexes.job]], 'job');
      if (job && rc.job[job as keyof typeof rc.job] !== undefined)
        rc.job[job as keyof typeof rc.job]++;
    });

    satisfactionQuestions.forEach((questionName, i) => {
      const columnIndex = satisfactionColumnIndexes[i];
      if (columnIndex === -1) return;
      const questionResponses = createEmptyQuestionDistribution();
      let totalScore = 0;
      let validResponses = 0;
      rows.forEach((row) => {
        const responseText = row[headers[columnIndex]];
        if (responseText) {
          const normalized = normalizeText(responseText, 'satisfaction');
          if (satisfactionScoreMapping[normalized] !== undefined) {
            questionResponses[normalized as keyof typeof questionResponses]++;
            totalScore += satisfactionScoreMapping[normalized];
            validResponses++;
          } else {
            // D3/D4: responseText 있는데 정규화 실패 → 무효 응답
            dc.add(
              'D4-satisfaction-invalid',
              'warning',
              'normalize-failed',
              `만족도 응답을 인식할 수 없어 제외됨 (과목: ${subject}, 문항: ${questionName})`,
              String(responseText),
            );
          }
        }
        // 빈 셀(무응답)은 정상 처리이므로 별도 진단 없음
      });
      satisfactionDistribution[subject].questions[questionName] = questionResponses;
      satisfactionAverages[subject].scores[questionName] =
        validResponses > 0 ? (totalScore / validResponses).toFixed(2) : 0;
    });

    // D6: 전체 평균에서 유효응답 없는 문항(score=0) 제외
    const validQScores = Object.entries(satisfactionAverages[subject].scores)
      .filter(([k]) => k !== '전체')
      .map(([, v]) => Number(v))
      .filter((n) => n > 0); // 1~5 척도이므로 0 = 유효응답 없음
    const emptyQCount = Object.entries(satisfactionAverages[subject].scores)
      .filter(([k]) => k !== '전체')
      .filter(([, v]) => Number(v) === 0).length;
    if (emptyQCount > 0) {
      dc.add(
        'D6-empty-questions-excluded',
        'info',
        'value-coercion',
        `유효응답 없는 문항 ${emptyQCount}개가 전체 평균 계산에서 제외됨`,
        subject,
      );
    }
    satisfactionAverages[subject].scores['전체'] =
      validQScores.length > 0
        ? (validQScores.reduce((a, b) => a + b, 0) / validQScores.length).toFixed(2)
        : 0;
  }

  // 원본데이터 시트의 과목명을 키로 하는 구분 맵
  // (외부 카테고리에 없는 경우 퍼지 매칭으로 보완)
  const subjectCategories: Record<string, string> = {};
  for (const subject of Object.keys(canonicalSubjectGroups)) {
    const cat = canonicalCategoryForSubject[subject] || findCategory(subject, externalCategories);
    if (cat) subjectCategories[subject] = cat;
  }

  return {
    respondentCharacteristics,
    satisfactionDistribution,
    satisfactionAverages,
    subjectCategories,
    totalSubjects: Object.keys(canonicalSubjectGroups).length,
    totalResponses: dataRows.length,
    diagnostics: dc.build(),
    subjectMerges,
  };
}
