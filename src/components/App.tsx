import { useState, useRef, useEffect } from 'react';
import type { AttendanceAnalysis, SatisfactionAnalysis } from '../types/analysis';
import type { StructuredAttendanceItem, RawRow } from '../types/analysis';
import {
  extractSheetData,
  structureAttendanceData,
  createDefaultCompletionRates,
  extractSubjectCategories,
} from '../lib/excel-parser';
import { analyzeData } from '../lib/attendance-analyzer';
import { analyzeSatisfactionData } from '../lib/satisfaction-analyzer';
import { generateCombinedExcel } from '../lib/excel-exporter';
import { AttendanceResults } from './AttendanceResults';
import { SatisfactionResults } from './SatisfactionResults';
import { CompletionRateModal } from './CompletionRateModal';
import { Ic } from './icons';

/* ================================================================
   메인 앱  (analyzer.html:896-1097)
   ================================================================ */

function fmtSize(b: number): string {
  return b < 1024 * 1024 ? (b / 1024).toFixed(0) + ' KB' : (b / 1024 / 1024).toFixed(1) + ' MB';
}

export function App() {
  const [fileName, setFileName] = useState('');
  const [fileSize, setFileSize] = useState(0);
  const [loading, setLoading] = useState(false);
  const [loadingStep, setLoadingStep] = useState('');
  const [error, setErrorState] = useState('');
  const errorTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);

  const [attendanceResult, setAttendanceResult] = useState<AttendanceAnalysis | null>(null);
  const [satisfactionResult, setSatisfactionResult] = useState<SatisfactionAnalysis | null>(null);
  const [originalAttendanceData, setOriginalAttendanceData] = useState<
    StructuredAttendanceItem[] | null
  >(null);
  const [subjectCompletionRates, setSubjectCompletionRates] = useState<Record<string, number>>({});

  const [isDragging, setIsDragging] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);
  const [activeTab, setActiveTab] = useState<'attendance' | 'satisfaction'>('attendance');
  const [showModal, setShowModal] = useState(false);
  // 연령 분포 60대/70대 구분 토글 (출석·수료 분석의 연령 분포에만 적용)
  const [splitSixties, setSplitSixties] = useState(true);

  const fileInputRef = useRef<HTMLInputElement>(null);

  function setError(message: string) {
    if (errorTimerRef.current) clearTimeout(errorTimerRef.current);
    setErrorState(message);
    if (message)
      errorTimerRef.current = setTimeout(() => {
        setErrorState('');
        errorTimerRef.current = null;
      }, 8000);
  }
  useEffect(
    () => () => {
      if (errorTimerRef.current) clearTimeout(errorTimerRef.current);
    },
    [],
  );

  // 수료 기준 또는 연령 세분화 토글 변경 시 출석 분석만 재실행
  useEffect(() => {
    if (originalAttendanceData && Object.keys(subjectCompletionRates).length > 0) {
      setAttendanceResult(analyzeData(originalAttendanceData, subjectCompletionRates, splitSixties));
    }
  }, [subjectCompletionRates, originalAttendanceData, splitSixties]);

  async function processExcelFile(file: File) {
    setLoading(true);
    setLoadingStep('파일 읽는 중...');
    setError('');
    setAttendanceResult(null);
    setSatisfactionResult(null);
    setOriginalAttendanceData(null);

    try {
      const MAX_FILE_SIZE = 50 * 1024 * 1024;
      if (file.size > MAX_FILE_SIZE)
        throw new Error('파일 크기가 너무 큽니다. 최대 50MB까지 업로드 가능합니다.');

      setFileName(file.name);
      setFileSize(file.size);

      const buffer = await file.arrayBuffer();

      setLoadingStep('Excel 파일 파싱 중...');
      const ExcelJS = window.ExcelJS;
      if (!ExcelJS)
        throw new Error('ExcelJS 라이브러리를 불러오지 못했습니다. 네트워크 연결을 확인하세요.');
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(buffer);

      setLoadingStep('시트 확인 중...');
      const attendanceSheetName = workbook.worksheets.find((ws) =>
        /^(\d+\.\s*)?개인별\s?출석현황$/.test(ws.name.trim()),
      )?.name;
      const satisfactionSheetName = workbook.worksheets.find((ws) =>
        /^(\d+\.\s*)?원본데이(?:터|타)$/.test(ws.name.trim()),
      )?.name;
      if (!attendanceSheetName || !satisfactionSheetName)
        throw new Error(
          '필수 시트를 찾을 수 없습니다. 파일에 "개인별 출석현황"과 "원본데이터" 시트가 있는지 확인하세요.',
        );

      setLoadingStep('출석 데이터 추출 중...');
      const attendanceRawData: RawRow[] = extractSheetData(workbook, attendanceSheetName, '연번');
      setLoadingStep('만족도 데이터 추출 중...');
      const satisfactionRawData: RawRow[] = extractSheetData(
        workbook,
        satisfactionSheetName,
        '연번',
      );

      setLoadingStep('데이터 구조화 중...');
      const structuredAttendance = structureAttendanceData(attendanceRawData);
      const defaultRates = createDefaultCompletionRates(structuredAttendance);

      setLoadingStep('데이터 분석 중...');
      const summarySheetName = workbook.worksheets.find((ws) =>
        /^(\d+\.\s*)?만족도$/.test(ws.name.trim()),
      )?.name;
      const externalCategories = summarySheetName
        ? extractSubjectCategories(workbook, summarySheetName)
        : {};
      const satisfactionAnalysis = analyzeSatisfactionData(satisfactionRawData, externalCategories);

      // 출석 분석은 setOriginalAttendanceData/setSubjectCompletionRates 설정 후
      // useEffect(66~70)가 단 1회 실행하므로 여기서 직접 호출하지 않는다.
      setOriginalAttendanceData(structuredAttendance);
      setSubjectCompletionRates(defaultRates);
      setSatisfactionResult(satisfactionAnalysis);
      setActiveTab('attendance');
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoading(false);
      setLoadingStep('');
    }
  }

  async function handleDownload() {
    if (!satisfactionResult || isDownloading) return;
    setIsDownloading(true);
    try {
      const ExcelJS = window.ExcelJS;
      if (!ExcelJS) throw new Error('ExcelJS 라이브러리를 불러오지 못했습니다.');
      const url = await generateCombinedExcel(satisfactionResult, ExcelJS, attendanceResult);
      if (url) {
        const link = document.createElement('a');
        link.href = url;
        link.download = '통합_분석결과.xlsx';
        link.rel = 'noopener';
        // ⚠️ 단순화 금지: 샌드박스 iframe에서 다운로드가 취소되지 않도록
        // document.body에 부착 후 click, 4초 후 정리해야 한다.
        // porting-guide §8-6의 즉시 revokeObjectURL 패턴은 버그이므로 사용 금지.
        document.body.appendChild(link);
        link.click();
        setTimeout(() => {
          document.body.removeChild(link);
          URL.revokeObjectURL(url);
        }, 4000);
      }
    } catch (err) {
      setError('Excel 생성 실패: ' + (err instanceof Error ? err.message : String(err)));
    } finally {
      setIsDownloading(false);
    }
  }

  function reset() {
    setFileName('');
    setFileSize(0);
    setAttendanceResult(null);
    setSatisfactionResult(null);
    setOriginalAttendanceData(null);
    setSubjectCompletionRates({});
    setError('');
    if (fileInputRef.current) fileInputRef.current.value = '';
  }

  const hasResult = attendanceResult || satisfactionResult;

  return (
    <div>
      <div className="masthead">
        <div className="kicker">평생학습 운영결과 분석</div>
        <h1 className="title">가까이배움터 강좌 운영결과 분석기</h1>
        <p className="subtitle">
          출석현황과 만족도 설문 원본이 담긴 엑셀 파일을 올리면, 출석·수료 분석과 만족도 집계를
          자동으로 수행하고 4개 시트로 구성된 결과 파일을 내려받을 수 있습니다.
        </p>
        <div className="rule" />
      </div>

      {error && (
        <div className="error">
          {Ic.alert}
          <span>{error}</span>
        </div>
      )}

      {!hasResult && !loading && (
        <div
          className={'drop' + (isDragging ? ' drag' : '')}
          onClick={() => fileInputRef.current && fileInputRef.current.click()}
          onDrop={(e) => {
            e.preventDefault();
            setIsDragging(false);
            const f = e.dataTransfer?.files?.[0];
            if (f) void processExcelFile(f);
          }}
          onDragOver={(e) => {
            e.preventDefault();
            setIsDragging(true);
          }}
          onDragLeave={() => setIsDragging(false)}
        >
          <div className="ic">{Ic.upload}</div>
          <h3>엑셀 파일을 드래그하거나 클릭하여 업로드</h3>
          <p>.xlsx · .xls 형식 · 최대 50MB</p>
          <div className="pill">개인별 출석현황 + 원본데이터 시트 필요</div>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls"
            style={{ display: 'none' }}
            onChange={(e) => {
              const f = e.target.files?.[0];
              if (f) void processExcelFile(f);
            }}
          />
        </div>
      )}

      {loading && (
        <div className="card loading">
          <div className="spinner" />
          <div>
            <div className="step">{loadingStep || '처리 중...'}</div>
            <div className="sub">{fileName || '파일을 분석하고 있습니다'}</div>
          </div>
        </div>
      )}

      {hasResult && !loading && (
        <>
          <div className="card filebar">
            <div className="fic">{Ic.file}</div>
            <div>
              <div className="fname">{fileName}</div>
              <div className="fmeta">{fmtSize(fileSize)} · 분석 완료</div>
            </div>
            <div className="grow" />
            <label className="switch-wrap">
              <span className="switch">
                <input
                  type="checkbox"
                  checked={splitSixties}
                  onChange={(e) => setSplitSixties(e.target.checked)}
                />
                <span className="switch-track" />
              </span>
              연령 60/70대 구분
            </label>
            <button className="btn ghost" onClick={() => setShowModal(true)}>
              {Ic.settings}수료 기준
            </button>
            <button className="btn ghost" onClick={reset}>
              {Ic.reset}새 파일
            </button>
            <button
              className="btn primary"
              onClick={() => void handleDownload()}
              disabled={isDownloading}
            >
              {Ic.download}
              {isDownloading ? '생성 중...' : '결과 파일 (4시트) 다운로드'}
            </button>
          </div>

          <div className="tabs">
            <button
              className={'tab' + (activeTab === 'attendance' ? ' active' : '')}
              onClick={() => setActiveTab('attendance')}
            >
              출석 · 수료 분석
              {attendanceResult && (
                <span className="cnt">{attendanceResult.subjectResults.length}</span>
              )}
            </button>
            <button
              className={'tab' + (activeTab === 'satisfaction' ? ' active' : '')}
              onClick={() => setActiveTab('satisfaction')}
            >
              만족도 분석
              {satisfactionResult && (
                <span className="cnt">{satisfactionResult.totalSubjects}</span>
              )}
            </button>
          </div>

          {activeTab === 'attendance' ? (
            <AttendanceResults
              analysis={attendanceResult}
              subjectCategories={satisfactionResult?.subjectCategories}
            />
          ) : (
            <SatisfactionResults results={satisfactionResult} />
          )}
        </>
      )}

      {showModal && (
        <CompletionRateModal
          rates={subjectCompletionRates}
          onChange={(subject, rate) =>
            setSubjectCompletionRates((prev) => ({ ...prev, [subject]: rate }))
          }
          onClose={() => setShowModal(false)}
        />
      )}

      <div className="footnote">
        모든 분석은 브라우저 내에서만 처리되며 파일은 외부로 전송되지 않습니다.
        <br />
        결과 파일 시트: 수강생 정보 분석 · 응답자 특성 · 객관식 응답 집계 · 평균만족도
      </div>
    </div>
  );
}
