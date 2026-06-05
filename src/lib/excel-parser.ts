import type ExcelJS from 'exceljs';
import type { RawRow, StructuredAttendanceItem, StructuredAttendanceResult } from '../types/analysis';
import { DiagnosticCollector } from './diagnostics';

/* ================================================================
   파싱 계층 — ExcelJS 헬퍼  (analyzer.html:224-310)
   ================================================================ */

function cellText(cell: ExcelJS.Cell): string {
  return cell.text || (cell.value != null ? String(cell.value) : '');
}

export function extractSheetData(
  workbook: ExcelJS.Workbook,
  sheetName: string,
  headerHint: string,
  dc?: DiagnosticCollector,
): RawRow[] {
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) throw new Error(`'${sheetName}' 시트를 찾을 수 없습니다.`);

  let dataStartRow = -1;
  let headers: string[] = [];

  worksheet.eachRow((row, rowNumber) => {
    if (dataStartRow !== -1) return;
    row.eachCell((cell) => {
      if (cellText(cell).includes(headerHint)) {
        dataStartRow = rowNumber;
        row.eachCell((headerCell, colNumber) => {
          headers[colNumber - 1] = cellText(headerCell);
        });
      }
    });
  });

  if (dataStartRow === -1)
    throw new Error(`'${sheetName}' 시트에서 '${headerHint}' 헤더를 찾을 수 없습니다.`);

  // 중복 헤더 → 바로 윗 행(설명 행) 값으로 대체
  if (dataStartRow > 1) {
    const seen = new Set<string>();
    const duplicates = new Set<string>();
    headers.forEach((h) => {
      if (h) {
        seen.has(h) ? duplicates.add(h) : seen.add(h);
      }
    });
    if (duplicates.size > 0) {
      const descRow = worksheet.getRow(dataStartRow - 1);
      const descHeaders: string[] = [];
      descRow.eachCell((cell, colNumber) => {
        descHeaders[colNumber - 1] = cellText(cell);
      });
      headers = headers.map((h, i) =>
        h && duplicates.has(h) && descHeaders[i] ? descHeaders[i] : h,
      );
    }
  }

  const data: RawRow[] = [];
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber <= dataStartRow) return;
    const rowData: RawRow = {};
    row.eachCell((cell, colNumber) => {
      const header = headers[colNumber - 1];
      if (header) {
        let value: unknown = cell.value;
        if (value && typeof value === 'object') {
          const v = value as Record<string, unknown>;
          if (v.richText) {
            value = (v.richText as Array<{ text: string }>).map((t) => t.text).join('');
          } else if ('result' in v) {
            // 수식 셀 + result 캐시 있음
            const result = v.result;
            if (result != null && typeof result !== 'object') {
              value = result; // 숫자·문자열 결과 그대로
            } else {
              value = cell.text || null; // {error:...} 또는 null
            }
          } else if ('formula' in v || 'sharedFormula' in v) {
            // P1: 수식 셀인데 result 캐시 없음 → null (isBlank → fallback 경로)
            // 원인: 파일이 엑셀에서 한 번도 저장되지 않았거나 계산이 비활성화된 경우
            dc?.add(
              'P1-formula-no-cache',
              'warning',
              'value-coercion',
              '수식 결과가 캐시되지 않은 셀 — 빈 값으로 처리됨 (엑셀에서 파일을 열고 저장 후 재업로드하면 해결)',
              `${sheetName} R${rowNumber}C${colNumber}`,
            );
            value = null;
          } else if (v.text) {
            value = v.text;
          } else {
            value = cell.text || null;
          }
        }
        if (value !== null) rowData[header] = value;
      }
    });
    if (Object.keys(rowData).length > 0) data.push(rowData);
  });

  return data;
}

// A6: 단위 붙은 숫자 추출 지원 ('10일'→10), 변환 불가 시 0+진단
function toCount(value: unknown, dc?: DiagnosticCollector, label?: string): number {
  const s = String(value ?? '').trim();
  if (s === '') return 0;
  const n = Number(s);
  if (Number.isFinite(n) && n >= 0) return Math.floor(n);
  // 앞부분 숫자 추출 ('10일' → 10, '8회' → 8)
  const match = /^(\d+(?:\.\d+)?)/.exec(s);
  if (match) {
    const extracted = Math.floor(parseFloat(match[1]));
    if (Number.isFinite(extracted) && extracted >= 0) {
      dc?.add(
        'A6-unit-stripped',
        'info',
        'value-coercion',
        '단위가 붙은 숫자에서 숫자만 추출됨',
        label ? `${label}: '${s}'` : `'${s}'`,
      );
      return extracted;
    }
  }
  if (s !== '') {
    dc?.add(
      'A6-nonnumeric-coerced',
      'warning',
      'value-coercion',
      '숫자로 변환할 수 없는 값을 0으로 처리함',
      label ? `${label}: '${s}'` : `'${s}'`,
    );
  }
  return 0;
}

// ── 출석률 칼럼 폴백 헬퍼 ──────────────────────────────────────────────────
// 셀 값에서 % 제거 후 숫자 변환. 유효하지 않으면 null 반환.
function parseRateCell(v: unknown): number | null {
  if (v == null) return null;
  const s = String(v).trim().replace('%', '');
  const n = Number(s);
  return Number.isFinite(n) && n >= 0 ? n : null;
}

function isBlank(v: unknown): boolean {
  return v == null || String(v).trim() === '';
}

export function structureAttendanceData(rawData: RawRow[]): StructuredAttendanceResult {
  const dc = new DiagnosticCollector();

  // B1: 출석률 형식 탐지 — 과반수/중앙값 기반 (기존 heuristic: any > 1 → 오판 가능)
  const rateNums = rawData
    .filter((r) => !isBlank(r['출석률']))
    .map((r) => parseRateCell(r['출석률']))
    .filter((n): n is number => n != null);

  let isPercentFormat = false;
  if (rateNums.length > 0) {
    const overOneCount = rateNums.filter((n) => n > 1).length;
    const frac = overOneCount / rateNums.length;
    const sortedRates = rateNums.slice().sort((a, b) => a - b);
    const median = sortedRates[Math.floor(sortedRates.length / 2)];
    isPercentFormat = frac >= 0.5 || median > 1;

    dc.add(
      'B1-rate-format',
      'info',
      'rate-format',
      `출석률 형식 자동탐지: ${isPercentFormat ? '백분율(0~100)' : '비율(0~1)'} — ${rateNums.length}개 샘플, 1초과 비중 ${(frac * 100).toFixed(0)}%, 중앙값 ${median.toFixed(2)}`,
    );

    // 판정 형식과 어긋나는 모호한 값 경고 (실수 혼재 가능성)
    const ambiguous = isPercentFormat
      ? rateNums.filter((n) => n > 0 && n <= 1) // 백분율 판정인데 0~1 사이 값 (의심)
      : rateNums.filter((n) => n > 1); // 비율 판정인데 1 초과 값 (의심)
    ambiguous.forEach((n) => {
      dc.add(
        'B1-rate-ambiguous',
        'warning',
        'rate-format',
        `판정 형식(${isPercentFormat ? '백분율' : '비율'})과 다른 범위의 출석률 값 발견`,
        n.toFixed(4),
      );
    });
  }

  const items: StructuredAttendanceItem[] = [];

  rawData.forEach((row) => {
    const rowLabel = `${String(row['과목명'] ?? '')} / ${String(row['이름'] ?? '')}`;

    const 수업일 = toCount(row['수업일'], dc, `수업일(${rowLabel})`);
    let 출석일: number;
    let 출석률: number;

    if (!isBlank(row['출석률'])) {
      // 출석률 칼럼 우선 — 엑셀에 기록된 공식 출석률을 직접 사용
      const raw = parseRateCell(row['출석률']);
      const rawRate = raw == null ? 0 : isPercentFormat ? raw / 100 : raw;
      // B3: 출석률이 100%를 초과하면 경고 후 클램프
      if (rawRate > 1) {
        dc.add(
          'B3-rate-overflow',
          'info',
          'integrity',
          `출석률이 100%를 초과하여 클램프됨`,
          `${rowLabel}: ${(rawRate * 100).toFixed(1)}%`,
        );
      }
      출석률 = Math.min(rawRate, 1);
      // 출석일: 칼럼 값이 있으면 그대로, 없으면 출석률에서 역산
      출석일 = !isBlank(row['출석일'])
        ? toCount(row['출석일'], dc, `출석일(${rowLabel})`)
        : 수업일 > 0
          ? Math.round(출석률 * 수업일)
          : 0;
    } else {
      // 출석률 칼럼 없음 → 수업일·출석일로 계산
      출석일 = toCount(row['출석일'], dc, `출석일(${rowLabel})`);
      // B4: 수업일=0인데 출석일>0이면 경고
      if (수업일 === 0 && 출석일 > 0) {
        dc.add(
          'B4-zero-class-days',
          'warning',
          'integrity',
          `수업일이 0인데 출석일(${출석일})이 있어 출석률 0으로 처리됨`,
          rowLabel,
        );
      }
      // B3: 출석일>수업일이면 경고
      if (수업일 > 0 && 출석일 > 수업일) {
        dc.add(
          'B3-attendance-exceeds-class',
          'warning',
          'integrity',
          `출석일(${출석일})이 수업일(${수업일})을 초과하여 클램프됨`,
          rowLabel,
        );
      }
      출석률 = Math.min(수업일 > 0 ? 출석일 / 수업일 : 0, 1);
    }

    const item: StructuredAttendanceItem = {
      연번: row['연번'],
      // 문자열 필드 trim — 앞뒤 공백 포함 표기(예: " 컴퓨터 기초 및 스마트폰 활용") 정규화
      과목명: String(row['과목명'] ?? '').trim(),
      이름: String(row['이름'] ?? '').trim(),
      성별: String(row['성별'] ?? '').trim(),
      생년월일: row['생년월일'] as string | number,
      수업일,
      출석일,
      출석률,
      수료여부: row['수료여부'],
    };

    // 스킵 행 보고: 과목명 또는 이름 누락 행 제외
    if (!item.과목명 || !item.이름) {
      dc.add(
        'A-skipped-row',
        'warning',
        'skipped-row',
        '과목명 또는 이름 누락으로 행 제외됨',
        String(row['연번'] ?? ''),
      );
    } else {
      items.push(item);
    }
  });

  return { items, diagnostics: dc.build() };
}

// "1. 만족도" 요약 시트에서 과목명 → 구분 매핑을 추출한다.
// 헤더 행에서 '구분'과 '과목'(띄어쓰기 무시)을 찾고 데이터 행을 읽는다.
export function extractSubjectCategories(
  workbook: ExcelJS.Workbook,
  sheetName: string,
): Record<string, string> {
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) return {};

  let headerRowNum = -1;
  let catCol = -1;
  let subCol = -1;

  worksheet.eachRow((row, rowNumber) => {
    if (headerRowNum !== -1) return;
    let foundCat = -1;
    let foundSub = -1;
    row.eachCell((cell, colNumber) => {
      const t = cellText(cell).replace(/\s+/g, '');
      if (t === '구분') foundCat = colNumber;
      if (t === '과목' || t === '과목명') foundSub = colNumber;
    });
    if (foundCat !== -1 && foundSub !== -1) {
      headerRowNum = rowNumber;
      catCol = foundCat;
      subCol = foundSub;
    }
  });

  if (headerRowNum === -1) return {};

  const result: Record<string, string> = {};
  let currentCat = '';
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber <= headerRowNum) return;
    const cat = cellText(row.getCell(catCol)).trim();
    const sub = cellText(row.getCell(subCol)).trim();
    if (cat) currentCat = cat;
    if (sub && sub !== '계' && currentCat) result[sub] = currentCat;
  });

  return result;
}

export function createDefaultCompletionRates(
  data: StructuredAttendanceItem[],
  defaultRate = 0.7,
): Record<string, number> {
  const subjects = [...new Set(data.map((item) => item.과목명))];
  const rates: Record<string, number> = {};
  subjects.forEach((subject) => {
    rates[subject] = defaultRate;
  });
  return rates;
}
