import type ExcelJS from 'exceljs';
import type {
  AttendanceAnalysis,
  SatisfactionAnalysis,
  StructuredAttendanceItem,
  Diagnostic,
} from '../types/analysis';
import { regionMapping, satisfactionQuestions } from './satisfaction-analyzer';
import { AGE_KEYS, JOB_KEYS, createEmptyQuestionDistribution } from './constants';
import { findCategory } from './subject-utils';

/* ================================================================
   Excel 내보내기 — 5개 시트  (analyzer.html:586-709)
   ================================================================ */

export async function generateCombinedExcel(
  satisfactionResults: SatisfactionAnalysis,
  ExcelJSLib: typeof ExcelJS,
  attendanceResults: AttendanceAnalysis | null = null,
  originalAttendanceData?: StructuredAttendanceItem[] | null,
  diagnostics?: Diagnostic[],
  meta?: {
    subjectCompletionRates?: Record<string, number>;
    parsedAttendanceRows?: number;
    parsedSatisfactionRows?: number;
  },
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
      if (s) sheet.addRow([g, s.명수, Number(s.비율.toFixed(1))]);
    });
    // '기타/미상' 버킷이 있으면 추가
    if (attendanceResults.genderDistribution['기타/미상']) {
      const s = attendanceResults.genderDistribution['기타/미상'];
      sheet.addRow(['기타/미상', s.명수, Number(s.비율.toFixed(1))]);
    }

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

  // E2: 전 과목 문항 키 합집합 — 첫 과목 기준 누락 방지
  const buildQuestions = (): string[] => {
    const questionSet = new Set<string>();
    satisfactionSubjects.forEach((subject) => {
      Object.keys(satisfactionResults.satisfactionAverages[subject].scores)
        .filter((q) => q !== '전체')
        .forEach((q) => questionSet.add(q));
    });
    // 표준 순서로 정렬 (satisfactionQuestions에 없는 키는 뒤에 추가)
    const ordered = satisfactionQuestions.filter((q) => questionSet.has(q));
    questionSet.forEach((q) => {
      if (!ordered.includes(q)) ordered.push(q);
    });
    return ordered;
  };

  // ── Sheet 3: 객관식 응답 집계 (만족도 과목 있을 때) ──
  if (satisfactionSubjects.length > 0) {
    const sheet = workbook.addWorksheet('객관식 응답 집계');
    const questions = buildQuestions();

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
    const questions = buildQuestions();
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

  // ── Sheet 5: 검증 (originalAttendanceData 있을 때) ──
  if (originalAttendanceData && originalAttendanceData.length > 0) {
    const sheet = workbook.addWorksheet('검증');
    sheet.getColumn(1).width = 18;
    sheet.getColumn(2).width = 18;
    sheet.getColumn(3).width = 14;
    sheet.getColumn(4).width = 14;
    sheet.getColumn(5).width = 12;
    sheet.getColumn(6).width = 10;
    sheet.getColumn(7).width = 10;
    sheet.getColumn(8).width = 12;
    sheet.getColumn(9).width = 12;

    // (a) 검증 요약
    boldRow(sheet.addRow(['■ 검증 요약']));
    const parsedAttRows = meta?.parsedAttendanceRows;
    const parsedSatRows = meta?.parsedSatisfactionRows;
    const skipped =
      parsedAttRows != null ? parsedAttRows - originalAttendanceData.length : null;
    sheet.addRow(['파싱된 출석 원시 행 수', parsedAttRows ?? '—']);
    sheet.addRow(['구조화 출석 행 수', originalAttendanceData.length]);
    sheet.addRow(['스킵된 행 수 (과목명/이름 누락)', skipped ?? '—']);
    sheet.addRow(['파싱된 만족도 응답 수', parsedSatRows ?? '—']);
    if (attendanceResults) {
      sheet.addRow(['분석 과목 수', attendanceResults.subjectResults.length]);
    }
    sheet.addRow([]);

    // (b) 진단 목록
    boldRow(sheet.addRow(['■ 진단 결과']));
    if (!diagnostics || diagnostics.length === 0) {
      sheet.addRow(['이상 없음']);
    } else {
      boldRow(sheet.addRow(['심각도', '코드', '카테고리', '메시지', '건수', '샘플']));
      diagnostics.forEach((d) => {
        sheet.addRow([
          d.severity === 'warning' ? '경고' : '정보',
          d.code,
          d.category,
          d.message,
          d.count,
          (d.samples ?? []).join(' | '),
        ]);
      });
    }
    sheet.addRow([]);

    // (c) 출석 원본 대조
    const rates = meta?.subjectCompletionRates ?? {};
    const VERIFY_ROW_CAP = 50000;
    const dataToExport = originalAttendanceData.slice(0, VERIFY_ROW_CAP);

    boldRow(sheet.addRow(['■ 출석 원본 대조']));
    boldRow(
      sheet.addRow([
        '과목명',
        '이름',
        '성별',
        '생년월일',
        '수업일',
        '출석일',
        '출석률(%)',
        '수료여부판정',
        '연번',
      ]),
    );
    dataToExport.forEach((item) => {
      const rate = rates[item.과목명] ?? 0.7;
      // C4: epsilon 비교 (analyzeData와 동일 로직)
      const passed = item.출석률 >= rate - 1e-9;
      sheet.addRow([
        item.과목명,
        item.이름,
        item.성별,
        item.생년월일,
        item.수업일,
        item.출석일,
        Number((item.출석률 * 100).toFixed(1)),
        passed ? '수료' : '미수료',
        String(item.연번 ?? ''),
      ]);
    });

    if (originalAttendanceData.length > VERIFY_ROW_CAP) {
      boldRow(
        sheet.addRow([
          `… 이하 ${originalAttendanceData.length - VERIFY_ROW_CAP}행 생략 (총 ${originalAttendanceData.length}행)`,
        ]),
      );
    }
  }

  const excelBuffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([excelBuffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  return URL.createObjectURL(blob);
}
