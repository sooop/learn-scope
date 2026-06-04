# React 포팅 구현 가이드

> 이 문서는 `app-portal`(Svelte 5 + SvelteKit)의 **엑셀 파싱 → 분석 결과 생성** 로직을 React로 재구현하는 코딩 에이전트를 위한 상세 구현 가이드입니다. 모든 핵심 함수 코드를 그대로 인용하므로, 1:1 이식이 가능합니다.

---

## 목차

1. [전체 데이터 흐름](#1-전체-데이터-흐름)
2. [의존성 및 환경 설정](#2-의존성-및-환경-설정)
3. [입력 Excel 파일 구조 계약](#3-입력-excel-파일-구조-계약)
4. [파싱 계층 — ExcelJS 헬퍼](#4-파싱-계층--exceljs-헬퍼)
5. [출석 분석 — attendance-analyzer.js](#5-출석-분석--attendance-analyzerjs)
6. [만족도 분석 — satisfaction-analyzer.js](#6-만족도-분석--satisfaction-analyzerjs)
7. [결과 표시 데이터 계약](#7-결과-표시-데이터-계약)
8. [Excel 내보내기](#8-excel-내보내기)
9. [재분석 트리거 (수료 기준 변경)](#9-재분석-트리거-수료-기준-변경)
10. [React 상태 설계 & Svelte→React 매핑](#10-react-상태-설계--sveltereact-매핑)
11. [React 포팅 체크리스트 & 함정 모음](#11-react-포팅-체크리스트--함정-모음)

---

## 1. 전체 데이터 흐름

```
사용자 파일 선택 (drag-drop / input[type=file])
        │
        ▼ File 객체
[파일 크기 검증] ─── 50MB 초과 → 에러
        │
        ▼ file.arrayBuffer()
[ExcelJS 파싱]  workbook.xlsx.load(buffer)
        │
        ▼ workbook.worksheets
[시트명 탐색]   정규식 매칭 (출석 시트 / 만족도 시트)
        │
        ▼ extractSheetData(workbook, sheetName, '연번')
[데이터 추출]   헤더 동적 탐색 + Rich Text 처리 → Array<Record>
        │
        ├──── [출석 트랙] ─────────────────────────────────┐
        │     structureAttendanceData(rawData)              │
        │     → { 과목명, 이름, 성별, 생년월일,            │
        │          수업일, 출석일, 출석률, 수료여부 }[]     │
        │                                                   │
        │     createDefaultCompletionRates(data, 0.7)       │
        │     → { [과목명]: 0.7, ... }                      │
        │                                                   │
        │     analyzeData(structured, completionRates)      │
        │     → AttendanceAnalysis                          │
        │                                                   │
        └──── [만족도 트랙] ───────────────────────────────┤
              analyzeSatisfactionData(rawData)              │
              → SatisfactionAnalysis                        │
                                                            │
                            ▼ (두 결과 모두 준비됨)         │
                    [화면 표시]  AttendanceResults          │
                                SatisfactionResults         │
                                                            │
                            ▼ (다운로드 버튼 클릭 시)       │
                    [Excel 생성 (지연)]                      │
                    generateCombinedExcel(                  │
                      satisfactionAnalysis,                  │
                      ExcelJS,                              │
                      attendanceAnalysis                    │
                    ) → Blob URL                            │
                            │                               │
                            ▼                               │
                    <a download>.click() + revokeObjectURL  │
```

**핵심 원칙:**
- Excel 파싱(I/O)과 분석(순수 함수)은 완전히 분리되어 있음
- 두 분석 트랙(출석/만족도)은 동일 파일에서 추출된 서로 다른 시트를 독립 처리
- Excel 생성은 다운로드 시점에 지연 실행 (분석 완료 직후 생성하지 않음)

---

## 2. 의존성 및 환경 설정

### 패키지

```bash
npm install exceljs
# 현재 원본: ^3.4.0 (실제 3.10.0)
# ExcelJS 4.x도 API 거의 호환 — 마이그레이션 가능
# 번들 크기가 크므로 동적 import 권장
```

### ExcelJS 동적 import (React 권장 패턴)

```js
// 분석 시작 시 처음 한 번만 로드
const ExcelJS = await import('exceljs').then(m => m.default || m);
```

### 브라우저 전용 API (SSR 환경 주의)

- `File.arrayBuffer()` — 파일을 ArrayBuffer로 변환
- `Blob`, `URL.createObjectURL(blob)` — 다운로드 URL 생성
- `URL.revokeObjectURL(url)` — 메모리 정리 (다운로드 직후 호출)

### Svelte 5 Runes → React 매핑

| Svelte 5 | React |
|---|---|
| `let x = $state(val)` | `const [x, setX] = useState(val)` |
| `$derived(expr)` | `useMemo(() => expr, [deps])` |
| `$props()` | props |
| `$bindable(val)` | controlled prop + onChange 콜백 |
| `$effect(() => { ... })` | `useEffect(() => { ... }, [deps])` |
| `onDestroy(() => { ... })` | `useEffect(() => { return () => { ... } }, [])` |

---

## 3. 입력 Excel 파일 구조 계약

### 필수 시트 2개

| 시트 | 시트명 패턴 | 설명 |
|---|---|---|
| 출석 현황 | `/^(\d+\.\s*)?개인별\s?출석현황$/` | 수강생별 출석 기록 |
| 만족도 원본 | `/^(\d+\.\s*)?원본데이(?:터\|타)$/` | 만족도 설문 응답 원본 |

**패턴 설명:**
- 앞에 숫자+점 접두사 선택적 허용 (`1. 개인별 출석현황` 형태)
- 공백 0~1개 허용 (`개인별출석현황` / `개인별 출석현황` 모두 매칭)
- "데이타" 오타 허용 (실제 파일에서 발생하는 오타 대응)

```js
const attendanceSheetName = workbook.worksheets.find(ws =>
  /^(\d+\.\s*)?개인별\s?출석현황$/.test(ws.name.trim())
)?.name;

const satisfactionSheetName = workbook.worksheets.find(ws =>
  /^(\d+\.\s*)?원본데이(?:터|타)$/.test(ws.name.trim())
)?.name;

if (!attendanceSheetName || !satisfactionSheetName) {
  throw new Error(
    '필수 시트를 찾을 수 없습니다. ' +
    '파일에 "개인별 출석현황"과 "원본데이터" 시트가 있는지 확인하세요.'
  );
}
```

### 출석 시트 필요 컬럼

헤더 행은 **`'연번'`** 텍스트를 포함하는 셀이 있는 행으로 동적 탐지 (고정 행 번호 아님).

| 헤더명 | 타입 | 설명 |
|---|---|---|
| `연번` | 정수 | 행 식별 번호 (헤더 탐지용 힌트이기도 함) |
| `과목명` | 문자열 | 수강 과목 이름 |
| `이름` | 문자열 | 수강생 이름 |
| `성별` | `'남'|'여'|'남성'|'여성'` | 성별 |
| `생년월일` | `'YYYY.MM.DD'|'YYYY-MM-DD'` | 생년월일 |
| `수업일` | 정수 | 총 수업 일수 |
| `출석일` | 정수 | 출석한 일수 |
| `수료여부` | 문자열 | 수료 여부 (현재 분석에서 직접 사용 안 함, 저장만) |

### 만족도 시트 필요 컬럼

| 헤더 패턴 | 폴백 인덱스 | 설명 |
|---|---|---|
| `연번` 포함 | — | 행 식별 |
| `과목명` 포함 | — | 과목명 (병합 셀 forward-fill 필요) |
| `/성별/i` | 4 | 응답자 성별 |
| `/지역\|거주\|주소/i` | 5 | 거주 지역 |
| `/연령\|나이/i` | 6 | 연령대 |
| `/직업\|직종/i` | 7 | 직업 |
| `2-1.*적극` | — | 만족도 문항 1 |
| `2-2.*전문` | — | 만족도 문항 2 |
| `2-3.*충실` | — | 만족도 문항 3 |
| `2-4.*시간` | — | 만족도 문항 4 |
| `2-5.*활용` | — | 만족도 문항 5 |
| `2-6.*유익` | — | 만족도 문항 6 |
| `2-7.*교재` | — | 만족도 문항 7 |
| `2-8.*(시설\|시섦\|교육지원)` | — | 만족도 문항 8 (`시섦`은 오타 허용) |
| `2-9.*만족` | — | 만족도 문항 9 |

---

## 4. 파싱 계층 — ExcelJS 헬퍼

> **원본 위치:** `src/lib/components/UnifiedAnalyzer.svelte:49-168`

### 4-1. `cellText(cell)` — 셀 값을 문자열로 변환

```js
/**
 * ExcelJS 셀 객체에서 텍스트를 추출합니다.
 * cell.text는 ExcelJS가 내부적으로 richText/공식 등을 처리한 문자열을 반환합니다.
 */
function cellText(cell) {
  return cell.text || (cell.value != null ? String(cell.value) : '');
}
```

### 4-2. `extractSheetData(workbook, sheetName, headerHint)` — 핵심 파싱 함수

**전체 코드 (그대로 이식 가능):**

```js
/**
 * Excel 시트에서 헤더를 동적으로 찾아 데이터를 추출합니다.
 *
 * @param {ExcelJS.Workbook} workbook - ExcelJS 워크북 객체
 * @param {string} sheetName - 시트명
 * @param {string} headerHint - 헤더 행을 식별하는 힌트 텍스트 (예: '연번')
 * @returns {Array<Record<string, any>>} 헤더를 키로 하는 행 객체 배열
 */
function extractSheetData(workbook, sheetName, headerHint) {
  const worksheet = workbook.getWorksheet(sheetName);
  if (!worksheet) {
    throw new Error(`'${sheetName}' 시트를 찾을 수 없습니다.`);
  }

  // 헤더 행 찾기
  let dataStartRow = -1;
  let headers = [];

  worksheet.eachRow((row, rowNumber) => {
    if (dataStartRow !== -1) return; // 이미 찾았으면 종료

    row.eachCell((cell) => {
      if (cellText(cell).includes(headerHint)) {
        dataStartRow = rowNumber;
        // 헤더 행 저장 (RichText 처리)
        row.eachCell((headerCell, colNumber) => {
          headers[colNumber - 1] = cellText(headerCell); // ← 1-based → 0-based
        });
      }
    });
  });

  if (dataStartRow === -1) {
    throw new Error(`'${sheetName}' 시트에서 '${headerHint}' 헤더를 찾을 수 없습니다.`);
  }

  // 중복 헤더가 있으면 이전 행(설명 행)의 값으로 대체
  if (dataStartRow > 1) {
    const seen = new Set();
    const duplicates = new Set();
    headers.forEach(h => {
      if (h) { seen.has(h) ? duplicates.add(h) : seen.add(h); }
    });
    if (duplicates.size > 0) {
      const descRow = worksheet.getRow(dataStartRow - 1);
      const descHeaders = [];
      descRow.eachCell((cell, colNumber) => {
        descHeaders[colNumber - 1] = cellText(cell);
      });
      headers = headers.map((h, i) =>
        (h && duplicates.has(h) && descHeaders[i]) ? descHeaders[i] : h
      );
    }
  }

  // 데이터 행 추출
  const data = [];
  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber <= dataStartRow) return; // 헤더 및 이전 행 건너뛰기

    const rowData = {};
    row.eachCell((cell, colNumber) => {
      const header = headers[colNumber - 1];
      if (header) {
        // ExcelJS의 Rich Text 객체 처리
        let value = cell.value;
        if (value && typeof value === 'object') {
          if (value.richText) {
            // Rich Text: [{ text: '...' }, ...] 배열
            value = value.richText.map(t => t.text).join('');
          } else if (value.text) {
            value = value.text;
          } else {
            value = cell.text || String(value);
          }
        }
        rowData[header] = value;
      }
    });

    // 빈 행이 아닌 경우만 추가
    if (Object.keys(rowData).length > 0) {
      data.push(rowData);
    }
  });

  return data;
}
```

**⚠️ 재구현 시 반드시 주의할 점:**

1. **1-based 인덱스:** ExcelJS의 `colNumber`/`rowNumber`는 1부터 시작. `colNumber - 1`로 0-based 배열에 저장.
2. **sparse 동작:** `eachRow`/`eachCell`은 기본적으로 빈 셀/빈 행을 건너뜀. 중간에 비어있는 컬럼이 생길 수 있음.
3. **중복 헤더 보정:** 만족도 설문지처럼 여러 문항이 동일한 헤더 텍스트를 가질 경우, 바로 윗 행(설명 행)의 같은 컬럼 값으로 대체. 이 로직이 없으면 만족도 문항 매핑이 실패함.
4. **헤더 탐색은 `includes()`:** `includes(headerHint)`이므로 '연번'이 부분 문자열로 포함된 셀도 매칭됨 (정확히 일치하지 않아도 됨).
5. **Rich Text 3단계:** `richText` 배열 → `text` 속성 → `cell.text` 순으로 폴백.

### 4-3. `toCount(value)` — 안전 정수 변환

```js
/**
 * 셀 값을 안전하게 0 이상의 정수로 변환합니다.
 * parseInt 대신 Number를 사용해 "12abc" 같은 부분 파싱을 방지합니다.
 *
 * @param {unknown} value
 * @returns {number} 0 이상의 정수, 변환 불가 시 0
 */
function toCount(value) {
  const n = Number(String(value ?? '').trim());
  return Number.isFinite(n) && n >= 0 ? Math.floor(n) : 0;
}
```

> `parseInt('12abc')` → 12 (부분 파싱됨)  
> `Number('12abc')` → NaN (올바른 동작)

### 4-4. `structureAttendanceData(rawData)` — 출석 데이터 구조화

```js
/**
 * extractSheetData로 추출한 원본 행 배열을 출석 분석용 구조로 변환합니다.
 * 출석률(0~1)을 계산하고, 과목명·이름이 없는 행을 제거합니다.
 *
 * @param {Array<Record<string, any>>} rawData
 * @returns {StructuredAttendanceItem[]}
 */
function structureAttendanceData(rawData) {
  return rawData.map(row => {
    const 수업일 = toCount(row['수업일']);
    const 출석일 = toCount(row['출석일']);
    return {
      연번:     row['연번'],
      과목명:   row['과목명'],
      이름:     row['이름'],
      성별:     row['성별'],
      생년월일: row['생년월일'],
      수업일,
      출석일,
      출석률:   수업일 > 0 ? 출석일 / 수업일 : 0,  // ← 0~1 비율 (퍼센트 아님!)
      수료여부: row['수료여부']
    };
  }).filter(item => item.과목명 && item.이름);  // 과목명·이름 둘 다 있어야
}
```

**TypeScript 타입 정의:**

```typescript
interface StructuredAttendanceItem {
  연번:     unknown;
  과목명:   string;
  이름:     string;
  성별:     string;
  생년월일: string | number;
  수업일:   number;
  출석일:   number;
  출석률:   number;  // 0~1 비율 (0.85 = 85%)
  수료여부: unknown;
}
```

### 4-5. `createDefaultCompletionRates(data, defaultRate)` — 기본 수료 기준 생성

```js
/**
 * 구조화된 출석 데이터에서 과목 목록을 추출해 기본 수료 기준(0.7 = 70%)을 부여합니다.
 *
 * @param {StructuredAttendanceItem[]} data
 * @param {number} defaultRate - 기본 수료 기준율 (0~1)
 * @returns {Record<string, number>} { [과목명]: 수료기준율 }
 */
function createDefaultCompletionRates(data, defaultRate = 0.7) {
  const subjects = [...new Set(data.map(item => item.과목명))];
  const rates = {};
  subjects.forEach(subject => { rates[subject] = defaultRate; });
  return rates;
}
```

### 4-6. `processExcelFile` 오케스트레이션 순서

React에서 파일 업로드 핸들러에 해당하는 로직:

```js
async function processExcelFile(file) {
  // 1. 상태 초기화
  setLoading(true);
  setLoadingStep('파일 읽는 중...');
  setError('');
  setAttendanceResult(null);
  setSatisfactionResult(null);

  try {
    // 2. 파일 크기 검증
    const MAX_FILE_SIZE = 50 * 1024 * 1024; // 50MB
    if (file.size > MAX_FILE_SIZE) {
      throw new Error(`파일 크기가 너무 큽니다. 최대 50MB까지 업로드 가능합니다.`);
    }

    // 3. 파일 정보 저장
    setCurrentFileName(file.name);
    setCurrentFileSize(file.size);

    // 4. ArrayBuffer 변환 + 복사본 보관 (이력 저장용)
    const buffer = await file.arrayBuffer();
    const fileContentForHistory = buffer.slice(0); // ← 복사본 필수 (load()가 buffer 소비)

    // 5. ExcelJS 파싱
    setLoadingStep('Excel 파일 파싱 중...');
    const ExcelJS = await import('exceljs').then(m => m.default || m);
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    // 6. 시트 탐색
    setLoadingStep('시트 확인 중...');
    const attendanceSheetName = workbook.worksheets.find(ws =>
      /^(\d+\.\s*)?개인별\s?출석현황$/.test(ws.name.trim())
    )?.name;
    const satisfactionSheetName = workbook.worksheets.find(ws =>
      /^(\d+\.\s*)?원본데이(?:터|타)$/.test(ws.name.trim())
    )?.name;

    if (!attendanceSheetName || !satisfactionSheetName) {
      throw new Error(
        '필수 시트를 찾을 수 없습니다. ' +
        '파일에 "개인별 출석현황"과 "원본데이터" 시트가 있는지 확인하세요.'
      );
    }

    // 7. 데이터 추출
    setLoadingStep('출석 데이터 추출 중...');
    const attendanceRawData = extractSheetData(workbook, attendanceSheetName, '연번');

    setLoadingStep('만족도 데이터 추출 중...');
    const satisfactionRawData = extractSheetData(workbook, satisfactionSheetName, '연번');

    // 8. 출석 데이터 구조화
    setLoadingStep('데이터 구조화 중...');
    const structuredAttendance = structureAttendanceData(attendanceRawData);

    // 9. 수료 기준 초기화
    const defaultRates = createDefaultCompletionRates(structuredAttendance);
    setSubjectCompletionRates(defaultRates);

    // 10. 원본 데이터 보관 (수료 기준 변경 시 재분석용)
    setOriginalAttendanceData(structuredAttendance);
    setSatisfactionRawDataForReanalysis(satisfactionRawData);

    // 11. 분석 실행 (동기 함수, await 없음)
    setLoadingStep('데이터 분석 중...');
    const attendanceAnalysis = analyzeData(structuredAttendance, defaultRates);
    const satisfactionAnalysis = analyzeSatisfactionData(satisfactionRawData);

    setAttendanceResult(attendanceAnalysis);
    setSatisfactionResult(satisfactionAnalysis);
    setSatisfactionAnalysisData(satisfactionAnalysis); // Excel 생성용 별도 보관

    // 12. (선택) IndexedDB 이력 저장 — 실패해도 결과 표시 막지 않음
    // await saveToHistory({ fileName, fileSize, fileContent: fileContentForHistory, ... });

  } catch (err) {
    setError(err instanceof Error ? err.message : String(err));
    // 에러는 8초 후 자동 소멸 (아래 섹션 참고)
  } finally {
    setLoading(false);
  }
}
```

### 4-7. 에러 자동 소멸 (8초)

```js
// React 구현 예시
const [error, setErrorState] = useState('');
const errorTimerRef = useRef(null);

function setError(message) {
  if (errorTimerRef.current) clearTimeout(errorTimerRef.current);
  setErrorState(message);
  if (message) {
    errorTimerRef.current = setTimeout(() => {
      setErrorState('');
      errorTimerRef.current = null;
    }, 8000);
  }
}

// cleanup
useEffect(() => {
  return () => {
    if (errorTimerRef.current) clearTimeout(errorTimerRef.current);
  };
}, []);
```

### 4-8. 드래그 앤 드롭 / 파일 선택 UI

```jsx
// React 구현 예시
function UploadArea({ onFile }) {
  const [isDragging, setIsDragging] = useState(false);

  return (
    <div
      onDrop={(e) => {
        e.preventDefault();
        setIsDragging(false);
        const file = e.dataTransfer?.files?.[0];
        if (file) onFile(file);
      }}
      onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
      onDragLeave={() => setIsDragging(false)}
      className={isDragging ? 'border-blue-500 bg-blue-50' : 'border-gray-300'}
    >
      <label>
        <input
          type="file"
          accept=".xlsx,.xls"
          onChange={(e) => {
            const file = e.target.files?.[0];
            if (file) onFile(file);
          }}
          className="hidden"
        />
        파일을 드래그하거나 클릭하여 업로드
      </label>
    </div>
  );
}
```

---

## 5. 출석 분석 — attendance-analyzer.js

> **원본 위치:** `src/lib/attendance-analyzer.js` (195줄 전체)

### 5-1. `calculateAge(birthDate)` — 만 나이 계산

```js
/**
 * 생년월일로 만 나이를 계산합니다.
 *
 * @param {string | number} birthDate - '1990.01.01' 또는 '1990-01-01' 형식
 * @returns {number | null} 만 나이. 유효하지 않으면 null
 */
export function calculateAge(birthDate) {
  if (!birthDate) return null;

  const dateStr = birthDate.toString().trim();
  let parts;

  if (dateStr.includes('.')) {
    parts = dateStr.split('.');
  } else if (dateStr.includes('-')) {
    parts = dateStr.split('-');
  } else {
    return null;
  }

  const [year, month, day] = parts.map(Number);

  // 숫자 파싱 검증 (빈 값/문자 포함 시 NaN)
  if (!Number.isInteger(year) || !Number.isInteger(month) || !Number.isInteger(day)) {
    return null;
  }

  const birth = new Date(year, month - 1, day);

  // 존재하지 않는 날짜 자동 보정 탐지 (예: 2024.02.30 → JS가 3월로 밀어버림)
  if (birth.getFullYear() !== year || birth.getMonth() !== month - 1 || birth.getDate() !== day) {
    return null;
  }

  const today = new Date();

  // 미래 생년월일 방지
  if (birth > today) return null;

  let age = today.getFullYear() - birth.getFullYear();

  if (today.getMonth() < birth.getMonth() ||
      (today.getMonth() === birth.getMonth() && today.getDate() < birth.getDate())) {
    age--; // 생일이 아직 안 지났으면 1살 차감
  }

  return age;
}
```

### 5-2. `getAgeGroup(age)` — 연령대 분류

```js
/**
 * 만 나이를 연령대 문자열로 변환합니다.
 *
 * @param {number} age
 * @returns {string} '19세 이하' | '20대' | '30대' | '40대' | '50대' | '60대 이상'
 */
export function getAgeGroup(age) {
  if (age <= 19) return '19세 이하'; // ← <= (19세 포함)
  if (age < 30)  return '20대';      // ← < (경계값 주의)
  if (age < 40)  return '30대';
  if (age < 50)  return '40대';
  if (age < 60)  return '50대';
  return '60대 이상';                // 60, 70, 80, 90대 모두 포함
}
```

### 5-3. `analyzeData(data, subjectCompletionRates)` — 메인 출석 분석

```js
const MAX_ROWS = 100000;

/**
 * 구조화된 출석 데이터를 분석합니다.
 *
 * @param {StructuredAttendanceItem[]} data
 * @param {Record<string, number>} subjectCompletionRates - { [과목명]: 수료기준율(0~1) }
 * @returns {AttendanceAnalysis | null}
 */
export function analyzeData(data, subjectCompletionRates) {
  if (!data) return null;

  if (data.length > MAX_ROWS) {
    throw new Error(`데이터가 너무 큽니다. 최대 ${MAX_ROWS.toLocaleString()}행까지 분석 가능합니다.`);
  }

  // ── 1. 과목별 기본 분석 ──────────────────────────────────────────────────
  const subjectStats = {};

  data.forEach(item => {
    if (!subjectStats[item.과목명]) {
      subjectStats[item.과목명] = { 수강인원: 0, 출석률합계: 0, 수료인원: 0 };
    }

    subjectStats[item.과목명].수강인원++;
    subjectStats[item.과목명].출석률합계 += item.출석률;

    // 수료 기준 없으면 0.7(70%) 기본값
    const completionRate = subjectCompletionRates[item.과목명] || 0.7;
    if (item.출석률 >= completionRate) {
      subjectStats[item.과목명].수료인원++;
    }
  });

  const subjectResults = Object.entries(subjectStats).map(([과목명, stats]) => ({
    과목명,
    수강인원:   stats.수강인원,
    평균출석률: stats.수강인원 > 0 ? (stats.출석률합계 / stats.수강인원) : 0, // ← 0~1 비율
    수료인원:   stats.수료인원,
    수료율:     stats.수강인원 > 0 ? (stats.수료인원 / stats.수강인원) : 0    // ← 0~1 비율
  }));

  // ── 2. 성별 분포 ─────────────────────────────────────────────────────────
  const genderStats = { 여성: 0, 남성: 0 };
  data.forEach(item => {
    if (item.성별 === '여성' || item.성별 === '여') genderStats.여성++;
    else if (item.성별 === '남성' || item.성별 === '남') genderStats.남성++;
  });

  const totalStudents = data.length; // 중복 포함 수강 등록 건수
  const genderDistribution = {
    여성: { 명수: genderStats.여성, 비율: totalStudents > 0 ? (genderStats.여성 / totalStudents * 100) : 0 },
    남성: { 명수: genderStats.남성, 비율: totalStudents > 0 ? (genderStats.남성 / totalStudents * 100) : 0 },
    합계: { 명수: genderStats.여성 + genderStats.남성, 비율: totalStudents > 0 ? 100 : 0 }
    // ↑ 비율은 0~100 (이미 퍼센트값)
  };

  // ── 3. 수강 강좌수별 분포 ─────────────────────────────────────────────────
  // 동일인 판단: 이름 + 성별 + 생년월일 조합
  const studentCourses = {};
  data.forEach(item => {
    const key = `${item.이름}_${item.성별}_${item.생년월일}`;
    if (!studentCourses[key]) {
      studentCourses[key] = { 이름: item.이름, 성별: item.성별, 생년월일: item.생년월일, 강좌수: 0, 강좌목록: [] };
    }
    studentCourses[key].강좌수++;
    studentCourses[key].강좌목록.push(item.과목명);
  });

  const courseCountStats = { 1: 0, 2: 0, '3이상': 0 };
  Object.values(studentCourses).forEach(student => {
    if (student.강좌수 === 1)      courseCountStats[1]++;
    else if (student.강좌수 === 2) courseCountStats[2]++;
    else                           courseCountStats['3이상']++;
  });

  const uniqueStudents = Object.keys(studentCourses).length;
  const courseCountDistribution = {
    '1강좌':      { 명수: courseCountStats[1],       비율: uniqueStudents > 0 ? (courseCountStats[1]       / uniqueStudents * 100) : 0 },
    '2강좌':      { 명수: courseCountStats[2],       비율: uniqueStudents > 0 ? (courseCountStats[2]       / uniqueStudents * 100) : 0 },
    '3강좌 이상': { 명수: courseCountStats['3이상'], 비율: uniqueStudents > 0 ? (courseCountStats['3이상'] / uniqueStudents * 100) : 0 }
    // ↑ 비율 분모는 uniqueStudents (고유 수강생 수)
  };

  // ── 4. 연령 분포 ──────────────────────────────────────────────────────────
  // 강좌별로 집계 (동일인 중복 포함)
  const ageStats = {
    '19세 이하': 0, '20대': 0, '30대': 0,
    '40대': 0,      '50대': 0, '60대 이상': 0
  };

  data.forEach(item => {
    const age = calculateAge(item.생년월일);
    if (age === null) return; // 유효하지 않은 생년월일 → 연령 집계에서 제외
    ageStats[getAgeGroup(age)]++;
  });

  const ageDistribution = {};
  Object.entries(ageStats).forEach(([ageGroup, count]) => {
    ageDistribution[ageGroup] = {
      명수: count,
      비율: totalStudents > 0 ? (count / totalStudents * 100) : 0
      // ↑ 비율 분모는 totalStudents (data.length, 중복 포함)
    };
  });

  return {
    subjectResults,
    genderDistribution,
    courseCountDistribution,
    ageDistribution,
    totalStudents,   // 중복 포함 수강 등록 건수
    uniqueStudents   // 고유 수강생 수 (이름+성별+생년월일 기준)
  };
}
```

### 5-4. 반환 타입 정의

```typescript
interface SubjectResult {
  과목명:     string;
  수강인원:   number;
  평균출석률: number;  // ← 0~1 비율 (표시 시 ×100 필요)
  수료인원:   number;
  수료율:     number;  // ← 0~1 비율 (표시 시 ×100 필요)
}

interface DistributionEntry {
  명수: number;
  비율: number;  // ← 0~100 퍼센트값 (표시 시 추가 변환 불필요)
}

interface AttendanceAnalysis {
  subjectResults:          SubjectResult[];
  genderDistribution:      Record<'여성' | '남성' | '합계', DistributionEntry>;
  courseCountDistribution: Record<'1강좌' | '2강좌' | '3강좌 이상', DistributionEntry>;
  ageDistribution:         Record<'19세 이하' | '20대' | '30대' | '40대' | '50대' | '60대 이상', DistributionEntry>;
  totalStudents:           number;
  uniqueStudents:          number;
}
```

### 5-5. 출석 분석 입출력 예제

**입력 (structuredAttendance 일부):**
```json
[
  { "과목명": "생활영어", "이름": "홍길동", "성별": "남", "생년월일": "1985.03.15", "수업일": 20, "출석일": 18, "출석률": 0.9 },
  { "과목명": "생활영어", "이름": "김영희", "성별": "여성", "생년월일": "1992.07.22", "수업일": 20, "출석일": 14, "출석률": 0.7 },
  { "과목명": "컴퓨터기초", "이름": "홍길동", "성별": "남", "생년월일": "1985.03.15", "수업일": 15, "출석일": 15, "출석률": 1.0 }
]
```

**수료 기준:** `{ "생활영어": 0.8, "컴퓨터기초": 0.7 }`

**출력 (주요 부분):**
```json
{
  "subjectResults": [
    { "과목명": "생활영어", "수강인원": 2, "평균출석률": 0.8, "수료인원": 1, "수료율": 0.5 },
    { "과목명": "컴퓨터기초", "수강인원": 1, "평균출석률": 1.0, "수료인원": 1, "수료율": 1.0 }
  ],
  "genderDistribution": {
    "여성": { "명수": 1, "비율": 33.333... },
    "남성": { "명수": 2, "비율": 66.666... },
    "합계": { "명수": 3, "비율": 100 }
  },
  "courseCountDistribution": {
    "1강좌":      { "명수": 1, "비율": 50 },
    "2강좌":      { "명수": 1, "비율": 50 },
    "3강좌 이상": { "명수": 0, "비율": 0 }
  },
  "totalStudents": 3,
  "uniqueStudents": 2
}
```

---

## 6. 만족도 분석 — satisfaction-analyzer.js

> **원본 위치:** `src/lib/satisfaction-analyzer.js`

### 6-1. 상수

```js
// 14개 행정구역 매핑 (키 = 값)
const regionMapping = {
  통진읍: "통진읍", 고촌읍: "고촌읍", 양촌읍: "양촌읍",
  월곶면: "월곶면", 하성면: "하성면", 대곶면: "대곶면",
  김포본동: "김포본동", 장기본동: "장기본동", 사우동: "사우동",
  풍무동: "풍무동", 장기동: "장기동", 구래동: "구래동",
  마산동: "마산동", 운양동: "운양동",
};
// 매핑 외 지역 → "기타지역"

// 만족도 5단계 점수 매핑
const satisfactionScoreMapping = {
  "매우 그렇다":    5,
  "그렇다":         4,
  "보통":           3,
  "그렇지 않다":    2,
  "매우 그렇지 않다": 1,
};
```

### 6-2. `normalizeText(text, category)` — 텍스트 정규화

**⚠️ 이 함수는 export되지 않지만 `analyzeSatisfactionData` 내부에서 사용. 반드시 함께 구현해야 함.**

```js
/**
 * 텍스트를 정규화하여 표준 카테고리 값으로 변환합니다.
 *
 * @param {any} text
 * @param {'gender' | 'region' | 'age' | 'job' | 'satisfaction'} category
 * @returns {string} 정규화된 값, 매칭 없으면 원본 text 반환
 */
function normalizeText(text, category) {
  if (!text) return "";

  // 전처리: 소문자화, 공백·구두점 제거
  const cleanText = text
    .toString()
    .toLowerCase()
    .replace(/\s+/g, "")
    .replace(/[.,!?]/g, "");

  switch (category) {
    case "gender":
      if (cleanText.includes("남") || cleanText.includes("male")) return "남성";
      if (cleanText.includes("여") || cleanText.includes("female")) return "여성";
      break;

    case "region":
      const regions = Object.keys(regionMapping);
      for (const region of regions) {
        if (cleanText.includes(region.replace(/\s+/g, "")))
          return region;
      }
      return "기타지역"; // 매칭 없으면 기타지역 (break 없음, 항상 반환)

    case "age":
      if (cleanText.includes("19") || cleanText.includes("10대") || cleanText.includes("미성년"))
        return "19세 이하";
      if (cleanText.includes("20")) return "20대";
      if (cleanText.includes("30")) return "30대";
      if (cleanText.includes("40")) return "40대";
      if (cleanText.includes("50")) return "50대";
      if (cleanText.includes("60") || cleanText.includes("70") ||
          cleanText.includes("80") || cleanText.includes("90"))
        return "60대 이상";
      break;

    case "job":
      if (cleanText.includes("직장") || cleanText.includes("회사") || cleanText.includes("사무"))
        return "직장인";
      if (cleanText.includes("자영")) return "자영업";
      if (cleanText.includes("농업") || cleanText.includes("어업") ||
          cleanText.includes("축산") || cleanText.includes("임업"))
        return "농어업축산임업";
      if (cleanText.includes("주부") || cleanText.includes("가정")) return "주부";
      if (cleanText.includes("학생") || cleanText.includes("대학")) return "학생";
      return "기타"; // 직업은 매칭 없으면 "기타" (break 없음)

    case "satisfaction":
      // ⚠️ if 판정 순서가 중요! 절대 순서를 바꾸지 말 것
      if (cleanText.includes("매우") && cleanText.includes("그렇다"))
        return "매우 그렇다";
      if (cleanText.includes("매우") &&
          (cleanText.includes("그렇지않다") || cleanText.includes("안다") || cleanText.includes("않다")))
        return "매우 그렇지 않다";
      if (cleanText.includes("그렇지않다") ||
          (cleanText.includes("그렇지") && cleanText.includes("않다")))
        return "그렇지 않다";
      if (cleanText.includes("그렇다")) return "그렇다";
      if (cleanText.includes("보통"))   return "보통";
      // 숫자 직접 입력 지원 (전처리 후 순수 숫자만 남은 경우)
      if (cleanText === "5") return "매우 그렇다";
      if (cleanText === "4") return "그렇다";
      if (cleanText === "3") return "보통";
      if (cleanText === "2") return "그렇지 않다";
      if (cleanText === "1") return "매우 그렇지 않다";
      break;
  }

  return text; // 매칭 없으면 원본 반환
}
```

**⚠️ satisfaction 분기 순서 의존성:**
- `"매우 그렇다"` 먼저 → `"매우 그렇지 않다"` → `"그렇지 않다"` → `"그렇다"` → `"보통"`
- 예: `"매우 그렇다"` 입력 시 `includes("그렇다")`가 true이므로, 앞에서 `includes("매우") && includes("그렇다")`를 먼저 체크해야 올바르게 분류됨

### 6-3. `analyzeSatisfactionData(rawData)` — 메인 만족도 분석

```js
const MAX_ROWS = 100000;

// 9개 만족도 문항 패턴 ↔ 표시 이름 매핑
const questionPatterns = [
  "2-1.*적극", "2-2.*전문", "2-3.*충실", "2-4.*시간",
  "2-5.*활용", "2-6.*유익", "2-7.*교재",
  "2-8.*(시설|시섦|교육지원)",  // "시섦"은 오타 대응
  "2-9.*만족"
];
const satisfactionQuestions = [
  "강사 적극성", "강사 전문성", "내용 충실성", "시간대 적정성",
  "수업의 활용도", "수업 유익성", "교재 및 자료 충분성",
  "강의 시설 만족도", "체감 만족도"
];

/**
 * @param {Array<Record<string, any>>} rawData - extractSheetData로 추출한 만족도 원본 데이터
 * @returns {SatisfactionAnalysis}
 */
export function analyzeSatisfactionData(rawData) {
  if (!rawData || rawData.length === 0) {
    throw new Error("분석할 데이터가 없습니다.");
  }
  if (rawData.length > MAX_ROWS) {
    throw new Error(`데이터가 너무 큽니다. 최대 ${MAX_ROWS.toLocaleString()}행까지 분석 가능합니다.`);
  }

  const headers = Object.keys(rawData[0]); // 헤더 = 첫 번째 행의 키 목록
  const dataRows = rawData;

  // ── 컬럼 인덱스 동적 탐지 ────────────────────────────────────────────────
  const columnIndexes = {
    serialNumber: headers.findIndex(h => h.includes('연번')),
    subject:      headers.findIndex(h => h.includes('과목명')),
    gender:       headers.findIndex(h => /성별/i.test(h)),
    region:       headers.findIndex(h => /지역|거주|주소/i.test(h)),
    age:          headers.findIndex(h => /연령|나이/i.test(h)),
    job:          headers.findIndex(h => /직업|직종/i.test(h)),
  };

  // 탐지 실패 시 폴백 인덱스 (gender/region/age/job만)
  if (columnIndexes.gender === -1) columnIndexes.gender = 4;
  if (columnIndexes.region === -1) columnIndexes.region = 5;
  if (columnIndexes.age    === -1) columnIndexes.age    = 6;
  if (columnIndexes.job    === -1) columnIndexes.job    = 7;
  // serialNumber, subject는 폴백 없음 → 탐지 실패 시 그룹핑 불가

  // ── 만족도 문항 컬럼 인덱스 탐지 ────────────────────────────────────────
  const satisfactionColumnIndexes = questionPatterns.map(pattern => {
    const regex = new RegExp(pattern, "i");
    return headers.findIndex(header => header && regex.test(header.toString()));
  });

  // ── 과목별 그룹핑 (병합 셀 forward-fill) ─────────────────────────────────
  const subjectGroups = {};
  let currentSubject = null;
  dataRows.forEach(row => {
    let subject = row[headers[columnIndexes.subject]];
    if (subject && subject.toString().trim() !== "") {
      currentSubject = subject.toString().trim();
    } else if (currentSubject) {
      subject = currentSubject; // 과목명 셀이 비어있으면 직전 과목명 사용 (병합 셀 대응)
    }
    if (!subject) return; // 과목명이 없으면 행 skip

    if (!subjectGroups[subject]) subjectGroups[subject] = [];
    subjectGroups[subject].push(row);
  });

  // ── 과목별 집계 ───────────────────────────────────────────────────────────
  const respondentCharacteristics = {};
  const satisfactionDistribution = {};
  const satisfactionAverages = {};

  for (const [subject, rows] of Object.entries(subjectGroups)) {
    // 초기 구조 생성
    respondentCharacteristics[subject] = {
      subject,
      gender: { 남성: 0, 여성: 0 },
      region: {
        ...Object.fromEntries(Object.keys(regionMapping).map(k => [k, 0])),
        기타지역: 0
      },
      age: { "19세 이하": 0, "20대": 0, "30대": 0, "40대": 0, "50대": 0, "60대 이상": 0 },
      job: { "직장인": 0, "자영업": 0, "농어업축산임업": 0, "주부": 0, "학생": 0, "기타": 0 }
    };
    satisfactionDistribution[subject] = { subject, questions: {} };
    satisfactionAverages[subject]     = { subject, scores: {} };

    // 응답자 특성 카운트
    rows.forEach(row => {
      const gender = normalizeText(row[headers[columnIndexes.gender]], "gender");
      if (gender) respondentCharacteristics[subject].gender[gender]++;

      const region = normalizeText(row[headers[columnIndexes.region]], "region");
      if (region) respondentCharacteristics[subject].region[region]++;

      const age = normalizeText(row[headers[columnIndexes.age]], "age");
      if (age) respondentCharacteristics[subject].age[age]++;

      const job = normalizeText(row[headers[columnIndexes.job]], "job");
      if (job) respondentCharacteristics[subject].job[job]++;
    });

    // 만족도 문항별 집계
    satisfactionQuestions.forEach((questionName, i) => {
      const columnIndex = satisfactionColumnIndexes[i];
      if (columnIndex === -1) return; // 헤더에서 해당 문항 컬럼을 찾지 못한 경우 skip

      const questionResponses = {
        "매우 그렇다": 0, "그렇다": 0, "보통": 0,
        "그렇지 않다": 0, "매우 그렇지 않다": 0
      };
      let totalScore = 0, validResponses = 0;

      rows.forEach(row => {
        const responseText = row[headers[columnIndex]];
        if (responseText) {
          const normalized = normalizeText(responseText, "satisfaction");
          if (satisfactionScoreMapping[normalized]) {
            questionResponses[normalized]++;
            totalScore += satisfactionScoreMapping[normalized];
            validResponses++;
          }
        }
      });

      satisfactionDistribution[subject].questions[questionName] = questionResponses;
      // ↓ 평균은 소수 2자리 문자열, 응답 없으면 숫자 0
      satisfactionAverages[subject].scores[questionName] =
        validResponses > 0 ? (totalScore / validResponses).toFixed(2) : 0;
    });

    // 9개 문항의 단순 산술평균 → "전체"
    const questionScores = Object.values(satisfactionAverages[subject].scores).map(Number);
    satisfactionAverages[subject].scores["전체"] = questionScores.length > 0
      ? (questionScores.reduce((a, b) => a + b, 0) / questionScores.length).toFixed(2)
      : 0;
    // ⚠️ 응답이 없는 문항은 0으로 포함되어 전체 평균을 낮춤
  }

  return {
    respondentCharacteristics, // { [subject]: { subject, gender, region, age, job } }
    satisfactionDistribution,  // { [subject]: { subject, questions: { [질문명]: 5단계카운트 } } }
    satisfactionAverages,      // { [subject]: { subject, scores: { [질문명]: "x.xx" | 0 } } }
    totalSubjects:   Object.keys(subjectGroups).length,
    totalResponses:  dataRows.length,
  };
}
```

### 6-4. 반환 타입 정의

```typescript
type SatisfactionScores = {
  [questionName: string]: string | number;  // toFixed(2) 문자열 또는 숫자 0
  전체: string | number;
};

interface SubjectSatisfaction {
  subject: string;
  scores:  SatisfactionScores;
}

interface QuestionDistribution {
  "매우 그렇다":    number;
  "그렇다":         number;
  "보통":           number;
  "그렇지 않다":    number;
  "매우 그렇지 않다": number;
}

interface RespondentCharacteristic {
  subject: string;
  gender:  Record<'남성' | '여성', number>;
  region:  Record<string, number>;  // 14개 지역 + '기타지역'
  age:     Record<'19세 이하' | '20대' | '30대' | '40대' | '50대' | '60대 이상', number>;
  job:     Record<'직장인' | '자영업' | '농어업축산임업' | '주부' | '학생' | '기타', number>;
}

interface SatisfactionAnalysis {
  respondentCharacteristics: Record<string, RespondentCharacteristic>;
  satisfactionDistribution:  Record<string, { subject: string; questions: Record<string, QuestionDistribution> }>;
  satisfactionAverages:      Record<string, SubjectSatisfaction>;
  totalSubjects:             number;
  totalResponses:            number;
}
```

### 6-5. 만족도 분석 입출력 예제

**입력 (rawData 일부):**
```json
[
  {
    "연번": 1, "과목명": "생활영어", "성별": "여", "거주지역": "김포본동",
    "연령": "40대", "직업": "주부",
    "2-1 강사의 적극적인 수업 태도": "매우 그렇다",
    "2-2 강사의 전문성": "그렇다",
    "2-9 전반적인 만족도": "매우 그렇다"
  },
  {
    "연번": 2, "과목명": "", "성별": "남성", "거주지역": "통진읍",
    "연령": "30", "직업": "직장인",
    "2-1 강사의 적극적인 수업 태도": "그렇다",
    "2-9 전반적인 만족도": "보통"
  }
]
```

**출력 (주요 부분):**
```json
{
  "satisfactionAverages": {
    "생활영어": {
      "subject": "생활영어",
      "scores": {
        "강사 적극성": "4.50",
        "강사 전문성": "4.00",
        "체감 만족도": "4.50",
        "전체": "4.33"
      }
    }
  },
  "respondentCharacteristics": {
    "생활영어": {
      "gender":  { "남성": 1, "여성": 1 },
      "region":  { "김포본동": 1, "통진읍": 1, ... },
      "age":     { "30대": 1, "40대": 1, ... },
      "job":     { "주부": 1, "직장인": 1, ... }
    }
  }
}
```

---

## 7. 결과 표시 데이터 계약

### 7-1. AttendanceResults — 출석 분석 표시

| 데이터 필드 | 표시 포맷 | 비고 |
|---|---|---|
| `analysis.subjectResults[]` | 테이블 반복 | — |
| `item.과목명` | 그대로 | — |
| `item.수강인원` | 그대로 (정수) | — |
| `item.평균출석률` | `(x * 100).toFixed(1) + "%"` | **0~1이므로 ×100 필수** |
| `item.수료인원` | 그대로 (정수) | — |
| `item.수료율` | `(x * 100).toFixed(1) + "%"` | **0~1이므로 ×100 필수** |
| `analysis.genderDistribution` | 테이블 반복 | `비율`은 이미 0~100 |
| `entry.명수` | 그대로 | — |
| `entry.비율` | `x.toFixed(1) + "%"` | ×100 불필요 |
| `analysis.uniqueStudents` | `총 N명의 고유 수강생` | — |
| `analysis.courseCountDistribution` | 동일 | 분모=uniqueStudents |
| `analysis.ageDistribution` | 동일 | 분모=totalStudents |

### 7-2. SatisfactionResults — 만족도 결과 표시

```js
// 동적 질문 컬럼 추출 (첫 번째 과목의 scores 키 기준)
const firstSubject = Object.values(results.satisfactionAverages)[0];
const questions = Object.keys(firstSubject.scores).filter(q => q !== "전체");
// → ["강사 적극성", "강사 전문성", ..., "체감 만족도"]

// 테이블 렌더링
// 헤더: ["과목명", "전체", ...questions]
// 행: Object.values(results.satisfactionAverages)
//   → [subjectData.subject, subjectData.scores["전체"], ...questions.map(q => subjectData.scores[q])]
```

**⚠️ 주의:** 첫 번째 과목 기준으로 질문 컬럼을 동적 추출하므로, 과목마다 만족도 문항 구성이 다르면 이후 과목의 데이터가 누락될 수 있음 (원본의 잠재적 이슈).

### 7-3. CompletionRateModal — 수료 기준 설정 UI

```jsx
// 드롭다운 선택지: [30, 40, 50, 60, 70, 80, 90, 100] (퍼센트 정수)
// 내부 상태: 0~1 비율
// UI ↔ 상태 변환:
//   표시값 = Math.round(rate * 100)
//   변경 시 = parseInt(e.target.value) / 100

function CompletionRateModal({ subjectCompletionRates, onRateChange, onClose }) {
  return (
    <div>
      {Object.entries(subjectCompletionRates).map(([subject, rate]) => (
        <div key={subject}>
          <span>{subject}</span>
          <select
            value={Math.round(rate * 100)}
            onChange={(e) => onRateChange(subject, parseInt(e.target.value) / 100)}
          >
            {[30, 40, 50, 60, 70, 80, 90, 100].map(pct => (
              <option key={pct} value={pct}>{pct}%</option>
            ))}
          </select>
        </div>
      ))}
    </div>
  );
}
```

---

## 8. Excel 내보내기

> **원본 위치:** `src/lib/satisfaction-analyzer.js:270-415`

### 8-1. 함수 시그니처

```js
/**
 * 출석 + 만족도 통합 Excel 파일을 생성하고 Blob URL을 반환합니다.
 *
 * @param {SatisfactionAnalysis} satisfactionResults - analyzeSatisfactionData 반환값
 * @param {typeof import('exceljs')} ExcelJS - ExcelJS 라이브러리 객체 (의존성 주입)
 * @param {AttendanceAnalysis | null} attendanceResults - analyzeData 반환값 (선택)
 * @returns {Promise<string | null>} Blob URL (사용 후 revokeObjectURL 필요)
 */
export async function generateCombinedExcel(satisfactionResults, ExcelJS, attendanceResults = null)
```

### 8-2. 4개 시트 구성

| # | 시트명 | 헤더 | 데이터 소스 | 조건 |
|---|---|---|---|---|
| 1 | `수강생 정보 분석` | 과목별/성별/강좌수/연령대 섹션 | `attendanceResults` | `attendanceResults`가 있을 때만 |
| 2 | `응답자 특성` | 교육과목/성별(3)/지역(15)/연령(7)/직업(7) | `satisfactionResults.respondentCharacteristics` | 항상 |
| 3 | `객관식 응답 집계` | 과목명 + 각 문항의 5단계×합계 | `satisfactionResults.satisfactionDistribution` | 만족도 과목 있을 때 |
| 4 | `평균만족도` | 과목명/전체/각 질문명 | `satisfactionResults.satisfactionAverages` | 만족도 과목 있을 때 |

### 8-3. Sheet 1 — 수강생 정보 분석

```js
// 과목별 분석결과 행
sheet.addRow([
  item.과목명,
  item.수강인원,
  Number((item.평균출석률 * 100).toFixed(1)),  // 0~1 → 퍼센트 숫자로 변환
  item.수료인원,
  Number((item.수료율 * 100).toFixed(1)),       // 0~1 → 퍼센트 숫자로 변환
]);

// 성별 분포 행 (비율은 이미 0~100)
sheet.addRow([gender, s.명수, Number(s.비율.toFixed(1))]);
```

### 8-4. Sheet 2 — 응답자 특성 고정 헤더

```js
const respondentHeaders = [
  "교육과목",
  "성별-남", "성별-여", "성별-합계",
  ...Object.keys(regionMapping),  // 14개 지역
  "기타지역", "지역-합계",
  "19세 이하", "20대", "30대", "40대", "50대", "60대 이상", "연령-합계",
  "직장인", "자영업", "농어업축산임업", "주부", "학생", "기타", "직업-합계",
];
```

### 8-5. Sheet 3 — 객관식 응답 집계 헤더 패턴

```js
// 질문명 추출 (Sheet 4와 동일한 방법)
const satisfactionQuestions = Object.keys(
  satisfactionResults.satisfactionAverages[satisfactionSubjects[0]].scores
).filter(q => q !== '전체');

// 헤더: ["과목명", "강사 적극성-매우그렇다", "강사 적극성-그렇다", ..., "강사 적극성-합계", ...]
satisfactionQuestions.forEach(q =>
  headers.push(`${q}-매우그렇다`, `${q}-그렇다`, `${q}-보통`, `${q}-그렇지않다`, `${q}-매우그렇지않다`, `${q}-합계`)
);

// 행 값: qd["매우 그렇다"], qd["그렇다"], qd["보통"], qd["그렇지 않다"], qd["매우 그렇지 않다"], 합계
```

### 8-6. Blob URL 생성 및 다운로드

```js
// Excel 생성
const excelBuffer = await workbook.xlsx.writeBuffer();
const blob = new Blob([excelBuffer], {
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
});
const url = URL.createObjectURL(blob);

// 다운로드 트리거
async function handleDownload() {
  if (!satisfactionAnalysisData || !attendanceResult || isDownloading) return;
  setIsDownloading(true);
  try {
    const ExcelJS = await import('exceljs').then(m => m.default || m);
    const url = await generateCombinedExcel(satisfactionAnalysisData, ExcelJS, attendanceResult);
    if (url) {
      const link = document.createElement('a');
      link.href = url;
      link.download = '통합_분석결과.xlsx';
      link.click();
      URL.revokeObjectURL(url); // 즉시 정리 (메모리 누수 방지)
    }
  } catch (err) {
    setError('Excel 생성 실패: ' + (err instanceof Error ? err.message : String(err)));
  } finally {
    setIsDownloading(false);
  }
}
```

### 8-7. 스타일링 (최소한)

원본 코드의 스타일은 매우 단순:
- **bold 폰트:** 섹션 제목 행 + 헤더 행에만 (`row.font = { bold: true }`)
- **컬럼 너비:** Sheet 1에만 (`sheet.getColumn(1).width = 30`, 2~5는 12~14)
- fill/border/색상 등 없음

---

## 9. 재분석 트리거 (수료 기준 변경)

수료 기준(`subjectCompletionRates`)이 변경되면 **출석 분석만** 재실행합니다. 만족도 분석은 영향 없음.

```js
// React 구현
useEffect(() => {
  if (originalAttendanceData && Object.keys(subjectCompletionRates).length > 0) {
    const reanalyzed = analyzeData(originalAttendanceData, subjectCompletionRates);
    setAttendanceResult(reanalyzed);
    // ← satisfactionResult는 변경하지 않음
  }
}, [subjectCompletionRates, originalAttendanceData]);
```

**이력에서 불러오기 (파싱 없이 재분석):**
```js
function loadFromHistory(historyEntry) {
  const {
    structuredAttendanceData,  // 이미 구조화된 데이터
    satisfactionRawData,
    subjectCompletionRates: savedRates
  } = historyEntry;

  setOriginalAttendanceData(structuredAttendanceData);
  setSubjectCompletionRates(savedRates);

  // 파싱 없이 바로 분석
  const attendanceAnalysis = analyzeData(structuredAttendanceData, savedRates);
  const satisfactionAnalysis = analyzeSatisfactionData(satisfactionRawData);

  setAttendanceResult(attendanceAnalysis);
  setSatisfactionResult(satisfactionAnalysis);
}
```

---

## 10. React 상태 설계 & Svelte→React 매핑

### 10-1. 상태 목록 (메인 컴포넌트)

```typescript
// 파일 정보
const [currentFileName, setCurrentFileName]   = useState<string>('');
const [currentFileSize, setCurrentFileSize]   = useState<number>(0);
const [currentFileContent, setCurrentFileContent] = useState<ArrayBuffer | null>(null);

// 로딩
const [loading, setLoading]       = useState(false);
const [loadingStep, setLoadingStep] = useState('');

// 에러
const [error, setErrorState]      = useState('');
const errorTimerRef                = useRef<NodeJS.Timeout | null>(null);

// 분석 결과
const [attendanceResult, setAttendanceResult]       = useState<AttendanceAnalysis | null>(null);
const [satisfactionResult, setSatisfactionResult]   = useState<SatisfactionAnalysis | null>(null);
const [satisfactionAnalysisData, setSatisfactionAnalysisData] = useState<SatisfactionAnalysis | null>(null);

// 재분석용 원본 데이터
const [originalAttendanceData, setOriginalAttendanceData]   = useState<StructuredAttendanceItem[] | null>(null);
const [satisfactionRawDataRef, setSatisfactionRawDataRef]   = useState<Record<string,any>[] | null>(null);

// 수료 기준 (변경 시 출석 재분석 트리거)
const [subjectCompletionRates, setSubjectCompletionRates]   = useState<Record<string, number>>({});

// UI 상태
const [isDragging, setIsDragging]   = useState(false);
const [isDownloading, setIsDownloading] = useState(false);
const [activeTab, setActiveTab]     = useState<'attendance' | 'satisfaction'>('attendance');
const [showModal, setShowModal]     = useState(false);
```

### 10-2. 파일 구조 제안 (React)

```
src/
├── components/
│   ├── UnifiedAnalyzer.tsx       # 메인 컴포넌트 (파일 업로드 + 오케스트레이션)
│   ├── AttendanceResults.tsx     # 출석 분석 결과 표시
│   ├── SatisfactionResults.tsx   # 만족도 결과 표시
│   └── CompletionRateModal.tsx   # 수료율 설정 모달
│
├── lib/
│   ├── excel-parser.ts           # cellText, extractSheetData, toCount,
│   │                             # structureAttendanceData, createDefaultCompletionRates
│   ├── attendance-analyzer.ts    # calculateAge, getAgeGroup, analyzeData
│   ├── satisfaction-analyzer.ts  # normalizeText, analyzeSatisfactionData
│   └── excel-exporter.ts         # generateCombinedExcel
│
└── types/
    └── analysis.ts               # TypeScript 타입 정의
```

---

## 11. React 포팅 체크리스트 & 함정 모음

아래 항목은 로직 오류나 동작 불일치가 발생하기 쉬운 곳입니다. 구현 후 반드시 검증하세요.

### ✅ ExcelJS 관련

- [ ] **1-based 인덱스:** `colNumber - 1`, `rowNumber`는 1부터 시작
- [ ] **sparse 동작:** `eachRow`/`eachCell`은 빈 셀/행을 건너뜀 → 헤더 배열에 중간 빈 슬롯 가능
- [ ] **Rich Text 3단계 처리:** `richText 배열` → `text 속성` → `cell.text` 순 폴백
- [ ] **중복 헤더 보정:** 만족도 시트의 동일 헤더 → 바로 윗 행(설명 행) 값으로 대체
- [ ] **헤더 탐색은 `includes()`:** 정확히 일치하지 않아도 포함이면 매칭

### ✅ 비율 단위 비대칭 (혼동 시 화면/Excel에서 10000% 등 잘못된 값 표시)

- [ ] `subjectResults`의 `평균출석률`, `수료율` → **0~1** (표시·export 시 `×100`)
- [ ] `genderDistribution`, `courseCountDistribution`, `ageDistribution`의 `비율` → **0~100** (이미 퍼센트)
- [ ] `subjectCompletionRates` → **0~1** (모달에서 정수%로 표시 시 `×100`, 저장 시 `/100`)

### ✅ 동일인 판단 키

- [ ] 강좌수별 분포의 고유 수강생 계산: `` `${이름}_${성별}_${생년월일}` ``
- [ ] 세 필드가 완전히 일치해야 동일인으로 처리 (이름만으로는 부족)

### ✅ 만족도 정규화

- [ ] `satisfaction` 분기의 if 판정 순서 절대 변경 금지
  - `"매우 그렇다"` → `"매우 그렇지 않다"` → `"그렇지 않다"` → `"그렇다"` → `"보통"` 순서
- [ ] 전처리(공백 제거)가 먼저 적용되므로 `"그렇지않다"` (공백 없음)로 비교
- [ ] 숫자 직접 입력(`"5"`, `"4"`, ...) 지원 포함

### ✅ 컬럼 폴백 인덱스

- [ ] `gender`=4, `region`=5, `age`=6, `job`=7 (헤더 탐지 실패 시 폴백)
- [ ] `subject`(과목명) 컬럼은 폴백 없음 → 탐지 실패 시 전체 그룹핑 실패

### ✅ 과목 forward-fill

- [ ] 만족도 시트의 병합 셀 대응: 과목명 셀이 비어있으면 직전 과목명 사용
- [ ] 이 로직 없으면 병합 셀 이후 행들이 그룹핑되지 않음

### ✅ 평균 점수 타입

- [ ] `satisfactionAverages.scores[questionName]` → `toFixed(2)` 결과는 **문자열** (`"4.50"`)
- [ ] 응답 없는 문항은 숫자 **`0`** (문자열 아님)
- [ ] `전체` 평균 계산 시 `Number()` 변환 후 산술평균 (원본 코드 `map(Number)` 참조)
- [ ] `scores["전체"]`는 9개 문항 단순 산술평균이며, 미응답 문항(값=0)도 분모에 포함됨

### ✅ 동적 질문 컬럼

- [ ] SatisfactionResults의 질문 컬럼은 **첫 번째 과목의 `scores` 키**에서 `"전체"` 제외하여 추출
- [ ] 과목마다 만족도 문항 구성이 다를 경우 누락 가능 (원본의 잠재적 이슈로, 의도적으로 동일하게 구현)

### ✅ 메모리 관리

- [ ] Blob URL은 다운로드 즉시 `URL.revokeObjectURL(url)` 호출
- [ ] 에러 타이머는 컴포넌트 언마운트 시 `clearTimeout` cleanup
- [ ] `file.arrayBuffer()` 호출 후 `buffer.slice(0)`으로 복사본 보관 (원본 buffer는 load()에서 소비됨)

### ✅ 시트명 정규식

- [ ] "개인별출석현황" (공백 없음)과 "개인별 출석현황" (공백 있음) 모두 매칭
- [ ] "원본데이타" 오타도 매칭 (`(?:터|타)`)
- [ ] 앞에 `"1. "` 같은 접두사 있어도 매칭 (`(\d+\.\s*)?`)

---

*이 가이드는 `app-portal` 커밋 `c810e34` 기준으로 작성되었습니다.*
*원본 소스 참조: `src/lib/attendance-analyzer.js`, `src/lib/satisfaction-analyzer.js`, `src/lib/components/UnifiedAnalyzer.svelte`*
