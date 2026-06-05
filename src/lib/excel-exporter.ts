import type ExcelJS from 'exceljs';
import type { AttendanceAnalysis, SatisfactionAnalysis } from '../types/analysis';
import { regionMapping } from './satisfaction-analyzer';
import { AGE_KEYS, JOB_KEYS, createEmptyQuestionDistribution } from './constants';
import { findCategory } from './subject-utils';

/* ================================================================
   Excel 내보내기 — 4개 시트  (analyzer.html:586-709)
   ================================================================ */

export async function generateCombinedExcel(
  satisfactionResults: SatisfactionAnalysis,
  ExcelJSLib: typeof ExcelJS,
  attendanceResults: AttendanceAnalysis | null = null,
): Promise<string> {
  const workbook = new ExcelJSLib.Workbook();
  const boldRow = (row: ExcelJS.Row) => {
    row.font = { bold: true };
  };

  // ── Sheet 1: 수강생 정보 분석 (출석 결과 있을 때만) ──
  if (attendanceResults) {
    const sheet = workbook.addWorksheet('수강생 정보 분석');
    sheet.getColumn(1).width = 16;
    sheet.getColumn(2).width = 30;
    [3, 4, 5, 6].forEach((c) => {
      sheet.getColumn(c).width = c === 4 || c === 6 ? 14 : 12;
    });

    boldRow(sheet.addRow(['■ 과목별 분석']));
    boldRow(sheet.addRow(['구분', '과목명', '수강인원', '평균출석률(%)', '수료인원', '수료율(%)']));
    attendanceResults.subjectResults.forEach((item) => {
      sheet.addRow([
        findCategory(item.과목명, satisfactionResults.subjectCategories),
        item.과목명,
        item.수강인원,
        Number((item.평균출석률 * 100).toFixed(1)),
        item.수료인원,
        Number((item.수료율 * 100).toFixed(1)),
      ]);
    });

    // 구분별 집계 표
    const categoryStats: Record<
      string,
      { 과목수: number; 수강인원: number; 출석률가중합: number; 수료인원: number; 수료율가중합: number }
    > = {};
    attendanceResults.subjectResults.forEach((item) => {
      const cat = findCategory(item.과목명, satisfactionResults.subjectCategories) || '미분류';
      if (!categoryStats[cat])
        categoryStats[cat] = { 과목수: 0, 수강인원: 0, 출석률가중합: 0, 수료인원: 0, 수료율가중합: 0 };
      categoryStats[cat].과목수++;
      categoryStats[cat].수강인원 += item.수강인원;
      categoryStats[cat].출석률가중합 += item.평균출석률 * item.수강인원;
      categoryStats[cat].수료인원 += item.수료인원;
      categoryStats[cat].수료율가중합 += item.수료율 * item.수강인원;
    });
    sheet.addRow([]);
    boldRow(sheet.addRow(['■ 구분별 분석']));
    boldRow(sheet.addRow(['구분', '과목수', '수강인원', '평균출석률(%)', '수료인원', '수료율(%)']));
    Object.entries(categoryStats).forEach(([cat, s]) => {
      sheet.addRow([
        cat,
        s.과목수,
        s.수강인원,
        Number((s.수강인원 > 0 ? (s.출석률가중합 / s.수강인원) * 100 : 0).toFixed(1)),
        s.수료인원,
        Number((s.수강인원 > 0 ? (s.수료율가중합 / s.수강인원) * 100 : 0).toFixed(1)),
      ]);
    });

    sheet.addRow([]);
    boldRow(sheet.addRow(['■ 성별 분포']));
    boldRow(sheet.addRow(['구분', '명수', '비율(%)']));
    (['여성', '남성', '합계'] as const).forEach((g) => {
      const s = attendanceResults.genderDistribution[g];
      sheet.addRow([g, s.명수, Number(s.비율.toFixed(1))]);
    });

    sheet.addRow([]);
    boldRow(
      sheet.addRow([`■ 수강 강좌수별 분포 (고유 수강생 ${attendanceResults.uniqueStudents}명)`]),
    );
    boldRow(sheet.addRow(['구분', '명수', '비율(%)']));
    (['1강좌', '2강좌', '3강좌 이상'] as const).forEach((k) => {
      const s = attendanceResults.courseCountDistribution[k];
      sheet.addRow([k, s.명수, Number(s.비율.toFixed(1))]);
    });

    sheet.addRow([]);
    boldRow(sheet.addRow(['■ 연령 분포']));
    boldRow(sheet.addRow(['구분', '명수', '비율(%)']));
    Object.entries(attendanceResults.ageDistribution).forEach(([g, s]) => {
      sheet.addRow([g, s.명수, Number(s.비율.toFixed(1))]);
    });
  }

  // ── Sheet 2: 응답자 특성 (항상) ──
  {
    const sheet = workbook.addWorksheet('응답자 특성');
    const regionKeys = Object.keys(regionMapping);
    const ageKeys = AGE_KEYS;
    const jobKeys = JOB_KEYS;

    const headers = [
      '교육과목',
      '성별-남',
      '성별-여',
      '성별-합계',
      ...regionKeys,
      '기타지역',
      '지역-합계',
      ...ageKeys,
      '연령-합계',
      ...jobKeys,
      '직업-합계',
    ];
    boldRow(sheet.addRow(headers));

    Object.values(satisfactionResults.respondentCharacteristics).forEach((rc) => {
      const genderSum = rc.gender.남성 + rc.gender.여성;
      const regionVals = regionKeys.map((k) => rc.region[k] || 0);
      const etcRegion = rc.region['기타지역'] || 0;
      const regionSum = regionVals.reduce((a, b) => a + b, 0) + etcRegion;
      const ageVals = ageKeys.map((k) => rc.age[k as keyof typeof rc.age] || 0);
      const ageSum = ageVals.reduce((a, b) => a + b, 0);
      const jobVals = jobKeys.map((k) => rc.job[k as keyof typeof rc.job] || 0);
      const jobSum = jobVals.reduce((a, b) => a + b, 0);

      sheet.addRow([
        rc.subject,
        rc.gender.남성,
        rc.gender.여성,
        genderSum,
        ...regionVals,
        etcRegion,
        regionSum,
        ...ageVals,
        ageSum,
        ...jobVals,
        jobSum,
      ]);
    });
  }

  const satisfactionSubjects = Object.keys(satisfactionResults.satisfactionAverages);

  // ── Sheet 3: 객관식 응답 집계 (만족도 과목 있을 때) ──
  if (satisfactionSubjects.length > 0) {
    const sheet = workbook.addWorksheet('객관식 응답 집계');
    const questions = Object.keys(
      satisfactionResults.satisfactionAverages[satisfactionSubjects[0]].scores,
    ).filter((q) => q !== '전체');

    const headerRow = ['과목명'];
    questions.forEach((q) =>
      headerRow.push(
        `${q}-매우그렇다`,
        `${q}-그렇다`,
        `${q}-보통`,
        `${q}-그렇지않다`,
        `${q}-매우그렇지않다`,
        `${q}-합계`,
      ),
    );
    boldRow(sheet.addRow(headerRow));

    satisfactionSubjects.forEach((subject) => {
      const dist = satisfactionResults.satisfactionDistribution[subject];
      const row: (string | number)[] = [subject];
      questions.forEach((q) => {
        const qd = (dist && dist.questions[q]) || createEmptyQuestionDistribution();
        const sum =
          qd['매우 그렇다'] + qd['그렇다'] + qd['보통'] + qd['그렇지 않다'] + qd['매우 그렇지 않다'];
        row.push(
          qd['매우 그렇다'],
          qd['그렇다'],
          qd['보통'],
          qd['그렇지 않다'],
          qd['매우 그렇지 않다'],
          sum,
        );
      });
      sheet.addRow(row);
    });
  }

  // ── Sheet 4: 평균만족도 (만족도 과목 있을 때) ──
  if (satisfactionSubjects.length > 0) {
    const sheet = workbook.addWorksheet('평균만족도');
    const questions = Object.keys(
      satisfactionResults.satisfactionAverages[satisfactionSubjects[0]].scores,
    ).filter((q) => q !== '전체');
    sheet.getColumn(1).width = 28;

    boldRow(sheet.addRow(['구분', '과목명', '전체', ...questions]));
    satisfactionSubjects.forEach((subject) => {
      const avg = satisfactionResults.satisfactionAverages[subject];
      sheet.addRow([
        avg.구분 ?? '',
        subject,
        Number(avg.scores['전체']) || 0,
        ...questions.map((q) => Number(avg.scores[q]) || 0),
      ]);
    });
  }

  const excelBuffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([excelBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  return URL.createObjectURL(blob);
}
