import type { RawRow, SatisfactionAnalysis } from '../types/analysis';
import { MAX_ROWS, createEmptyQuestionDistribution } from './constants';

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
      break;
    case 'region': {
      const regions = Object.keys(regionMapping);
      for (const region of regions) {
        if (cleanText.includes(region.replace(/\s+/g, ''))) return region;
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
      if (cleanText === '5') return '매우 그렇다';
      if (cleanText === '4') return '그렇다';
      if (cleanText === '3') return '보통';
      if (cleanText === '2') return '그렇지 않다';
      if (cleanText === '1') return '매우 그렇지 않다';
      break;
  }
  return String(text);
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

const satisfactionQuestions = [
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

export function analyzeSatisfactionData(rawData: RawRow[]): SatisfactionAnalysis {
  if (!rawData || rawData.length === 0) throw new Error('분석할 데이터가 없습니다.');
  if (rawData.length > MAX_ROWS)
    throw new Error(
      `데이터가 너무 큽니다. 최대 ${MAX_ROWS.toLocaleString()}행까지 분석 가능합니다.`,
    );

  const headers = Object.keys(rawData[0]);
  const dataRows = rawData;

  const columnIndexes = {
    serialNumber: headers.findIndex((h) => h.includes('연번')),
    subject: headers.findIndex((h) => h.includes('과목명')),
    gender: headers.findIndex((h) => /성별/i.test(h)),
    region: headers.findIndex((h) => /지역|거주|주소/i.test(h)),
    age: headers.findIndex((h) => /연령|나이/i.test(h)),
    job: headers.findIndex((h) => /직업|직종/i.test(h)),
  };
  if (columnIndexes.gender === -1) columnIndexes.gender = 4;
  if (columnIndexes.region === -1) columnIndexes.region = 5;
  if (columnIndexes.age === -1) columnIndexes.age = 6;
  if (columnIndexes.job === -1) columnIndexes.job = 7;

  const satisfactionColumnIndexes = questionPatterns.map((pattern) => {
    const regex = new RegExp(pattern, 'i');
    return headers.findIndex((header) => header && regex.test(header.toString()));
  });

  // ── 과목별 그룹핑 (병합 셀 forward-fill) ──────────────────────────────
  const subjectGroups: Record<string, RawRow[]> = {};
  let currentSubject: string | null = null;
  dataRows.forEach((row) => {
    let subject = row[headers[columnIndexes.subject]];
    if (subject && subject.toString().trim() !== '') {
      currentSubject = subject.toString().trim();
    } else if (currentSubject) {
      subject = currentSubject;
    }
    if (!subject) return;
    const subjectKey = subject.toString().trim();
    if (!subjectGroups[subjectKey]) subjectGroups[subjectKey] = [];
    subjectGroups[subjectKey].push(row);
  });

  const respondentCharacteristics: SatisfactionAnalysis['respondentCharacteristics'] = {};
  const satisfactionDistribution: SatisfactionAnalysis['satisfactionDistribution'] = {};
  const satisfactionAverages: SatisfactionAnalysis['satisfactionAverages'] = {};

  for (const [subject, rows] of Object.entries(subjectGroups)) {
    respondentCharacteristics[subject] = {
      subject,
      gender: { 남성: 0, 여성: 0 },
      region: {
        ...Object.fromEntries(Object.keys(regionMapping).map((k) => [k, 0])),
        기타지역: 0,
      },
      age: { '19세 이하': 0, '20대': 0, '30대': 0, '40대': 0, '50대': 0, '60대 이상': 0 },
      job: { 직장인: 0, 자영업: 0, 농어업축산임업: 0, 주부: 0, 학생: 0, 기타: 0 },
    };
    satisfactionDistribution[subject] = { subject, questions: {} };
    satisfactionAverages[subject] = { subject, scores: {} };

    rows.forEach((row) => {
      const gender = normalizeText(row[headers[columnIndexes.gender]], 'gender');
      if (gender && respondentCharacteristics[subject].gender[gender as '남성' | '여성'] !== undefined)
        respondentCharacteristics[subject].gender[gender as '남성' | '여성']++;
      const region = normalizeText(row[headers[columnIndexes.region]], 'region');
      if (region && respondentCharacteristics[subject].region[region] !== undefined)
        respondentCharacteristics[subject].region[region]++;
      const age = normalizeText(row[headers[columnIndexes.age]], 'age');
      if (age && respondentCharacteristics[subject].age[age as '19세 이하'] !== undefined)
        respondentCharacteristics[subject].age[
          age as keyof typeof respondentCharacteristics[typeof subject]['age']
        ]++;
      const job = normalizeText(row[headers[columnIndexes.job]], 'job');
      if (job && respondentCharacteristics[subject].job[job as '직장인'] !== undefined)
        respondentCharacteristics[subject].job[
          job as keyof typeof respondentCharacteristics[typeof subject]['job']
        ]++;
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
          if (satisfactionScoreMapping[normalized]) {
            questionResponses[normalized as keyof typeof questionResponses]++;
            totalScore += satisfactionScoreMapping[normalized];
            validResponses++;
          }
        }
      });
      satisfactionDistribution[subject].questions[questionName] = questionResponses;
      satisfactionAverages[subject].scores[questionName] =
        validResponses > 0 ? (totalScore / validResponses).toFixed(2) : 0;
    });

    const questionScores = Object.values(satisfactionAverages[subject].scores).map(Number);
    satisfactionAverages[subject].scores['전체'] =
      questionScores.length > 0
        ? (questionScores.reduce((a, b) => a + b, 0) / questionScores.length).toFixed(2)
        : 0;
  }

  return {
    respondentCharacteristics,
    satisfactionDistribution,
    satisfactionAverages,
    totalSubjects: Object.keys(subjectGroups).length,
    totalResponses: dataRows.length,
  };
}
