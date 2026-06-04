import type ExcelJS from 'exceljs';
import type { RawRow, StructuredAttendanceItem } from '../types/analysis';

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
          } else if (v.text) {
            value = v.text;
          } else {
            value = cell.text || String(value);
          }
        }
        rowData[header] = value;
      }
    });
    if (Object.keys(rowData).length > 0) data.push(rowData);
  });

  return data;
}

function toCount(value: unknown): number {
  const n = Number(String(value ?? '').trim());
  return Number.isFinite(n) && n >= 0 ? Math.floor(n) : 0;
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

export function structureAttendanceData(rawData: RawRow[]): StructuredAttendanceItem[] {
  // 출석률 칼럼 형식 1회 탐지 (출석일 공란 대비)
  // 1을 초과하는 값이 하나라도 있으면 백분율 형식(0~100), 없으면 비율 형식(0~1)
  const rateNums = rawData
    .map((r) => parseRateCell(r['출석률']))
    .filter((n): n is number => n != null);
  const isPercentFormat = rateNums.some((n) => n > 1);

  return rawData
    .map((row) => {
      const 수업일 = toCount(row['수업일']);
      let 출석일 = toCount(row['출석일']);
      let 출석률: number;

      if (!isBlank(row['출석일'])) {
        // 출석일 있음 → 기존 로직
        출석률 = 수업일 > 0 ? 출석일 / 수업일 : 0;
      } else {
        // 출석일 공란 → 출석률 칼럼 폴백 + 형식 휴리스틱
        const raw = parseRateCell(row['출석률']);
        출석률 = raw == null ? 0 : isPercentFormat ? raw / 100 : raw;
        // 출석일을 출석률에서 역산 (내부 일관성 유지)
        출석일 = 수업일 > 0 ? Math.round(출석률 * 수업일) : 0;
      }

      return {
        연번: row['연번'],
        과목명: row['과목명'] as string,
        이름: row['이름'] as string,
        성별: row['성별'] as string,
        생년월일: row['생년월일'] as string | number,
        수업일,
        출석일,
        출석률,
        수료여부: row['수료여부'],
      };
    })
    .filter((item) => item.과목명 && item.이름);
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
