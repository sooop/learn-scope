# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 개요

김포시 가까이배움터 강좌 운영결과 분석기. 엑셀 파일을 브라우저에서 파싱하여 출석·수료 분석과 만족도 집계를 수행하고, 분석 결과를 4개 시트 xlsx로 내보낸다. **최종 산출물은 단일 HTML 파일**이다.

## 명령어

```bash
npm run dev        # 개발 서버
npm run build      # tsc 타입체크 → dist/index.html 단일 파일 생성
npm run typecheck  # 타입 체크만 (빌드 없이)
npm run preview    # dist/ 빌드 결과 미리보기
```

테스트 프레임워크 없음. 기능 검증은 `npm run dev` 후 브라우저에서 직접 확인한다.

## 아키텍처

### 핵심 제약: ExcelJS는 CDN에서만 로드

`exceljs`는 `devDependencies`에 **타입 전용**으로 설치(`"exceljs": "4.4.0"`). 런타임에는 `index.html`의 `<script src="https://cdnjs.cloudflare.com/.../exceljs.min.js">` CDN 태그로 로드되며, 앱은 `window.ExcelJS`로 접근한다. **절대 값 import 금지** — `verbatimModuleSyntax: true`가 설정되어 있어 실수로 값 import 시 컴파일 에러가 발생한다. 타입 참조는 `import type ExcelJS from 'exceljs'`만 허용.

`vite-plugin-singlefile`이 React/CSS를 HTML에 인라인하지만, CDN의 절대 URL `<script>`는 건드리지 않으므로 최종 `dist/index.html`은 ExcelJS만 CDN 의존성으로 남는다.

### 데이터 흐름

```
File 업로드
  → extractSheetData() [excel-parser.ts]  — ExcelJS로 시트 파싱, 헤더 동적 탐지
  ├→ structureAttendanceData()            — 출석 시트 구조화 + 출석률(0~1) 계산
  │    analyzeData()  [attendance-analyzer.ts]  → AttendanceAnalysis
  └→ analyzeSatisfactionData() [satisfaction-analyzer.ts] → SatisfactionAnalysis
                                              ↓
                            화면 표시 (AttendanceResults / SatisfactionResults)
                                              ↓ (다운로드 클릭)
                            generateCombinedExcel() [excel-exporter.ts] → Blob URL
```

분석 함수는 모두 **순수 함수**이며 ExcelJS에 의존하지 않는다. ExcelJS 의존성은 파싱(`excel-parser.ts`)과 내보내기(`excel-exporter.ts`)에만 있다.

### 연령 분포 세분화 토글 (`splitSixties`)

`App.tsx`의 `splitSixties: boolean` 상태가 **출석·수료 분석의 연령 분포에만** 영향을 준다.
- `false`(기본): `60대 이상` 단일 구간
- `true`: `60대` / `70대 이상` 두 구간으로 분리

`analyzeData(data, rates, splitSixties)` 3번째 인자로 전달. 토글이 바뀌면 `useEffect`가 출석 분석만 재실행한다. `AttendanceAnalysis.ageDistribution`은 `Record<string, DistributionEntry>`(동적 키)이므로 표시(`DistTable`)·내보내기(`excel-exporter` Sheet1) 모두 변경 없이 자동 반영된다.

### 출석률 칼럼 폴백 (`structureAttendanceData`)

어떤 행의 `출석일`이 공란이면 `출석률` 칼럼 값으로 대체한다. 칼럼 전체를 1회 스캔하여 1 초과 값이 하나라도 있으면 **백분율 형식(0~100 → /100)**, 없으면 **비율 형식(0~1 그대로)** 으로 처리. 내부 출석률은 항상 0~1 비율로 통일된다.

### 비율 단위 비대칭 — 가장 흔한 버그

| 필드 | 단위 | 표시 시 |
|---|---|---|
| `subjectResults[].평균출석률`, `.수료율` | **0~1** | `× 100` 필요 |
| `genderDistribution`, `courseCountDistribution`, `ageDistribution`의 `.비율` | **0~100** (이미 %) | 그대로 사용 |
| `subjectCompletionRates` (수료 기준) | **0~1** | 모달에서 `× 100` 표시, 저장 시 `/ 100` |

### 만족도 분석 주의사항

- `normalizeText()` satisfaction 분기의 **if 판정 순서 절대 변경 금지** — "매우 그렇다"가 "그렇다"를 contains하므로 순서 의존적이다
- `satisfactionAverages.scores[질문명]`은 `toFixed(2)` **문자열**이지만, 미응답 문항은 숫자 `0`
- 만족도 시트의 병합 셀은 forward-fill로 처리 (`currentSubject` 변수로 직전 과목명 전파)

### 다운로드 로직 (`handleDownload`)

`document.body.appendChild(link)` → `link.click()` → **4초 후** `removeChild` + `revokeObjectURL` 패턴을 반드시 유지한다. 즉시 revoke하면 샌드박스 iframe에서 다운로드가 취소된다. `porting-guide.md §8-6`의 즉시 revoke 패턴은 버그이므로 따르지 않는다.

### 재분석 트리거

`subjectCompletionRates` 변경 시 `useEffect`가 출석 분석(`analyzeData`)만 재실행한다. 만족도 분석은 수료 기준과 무관하므로 재실행하지 않는다.

## 파일 구조 맥락

- `porting-guide.md` — 원본 Svelte 앱의 구현 명세. **타입 정의와 로직 설명은 참고**하되, §8의 다운로드 코드는 버그가 있으므로 `analyzer.html`의 구현을 따른다.
- `analyzer.html` — 빌드 프로젝트의 원본 프로토타입 (babel-standalone 기반). 동작 기준점으로 보존.
- `src/types/analysis.ts` — 모든 분석 타입의 단일 출처
- `src/types/exceljs-global.d.ts` — `window.ExcelJS` 전역 타입 선언
