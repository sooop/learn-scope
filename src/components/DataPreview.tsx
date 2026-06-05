import { useState, useMemo } from 'react';
import type { StructuredAttendanceItem, RawRow } from '../types/analysis';
import { pct } from './DistTable';

/* ================================================================
   원본 데이터 미리보기 — 파싱된 원시 데이터를 페이지네이션으로 열람
   ================================================================ */

const PAGE_SIZE = 100;

interface DataPreviewProps {
  attendance: StructuredAttendanceItem[] | null;
  satisfactionRaw: RawRow[] | null;
}

type PreviewTab = 'attendance' | 'satisfaction';

export function DataPreview({ attendance, satisfactionRaw }: DataPreviewProps) {
  const [which, setWhich] = useState<PreviewTab>('attendance');
  const [page, setPage] = useState(0);

  // 탭 전환 시 페이지 초기화
  function switchTab(tab: PreviewTab) {
    setWhich(tab);
    setPage(0);
  }

  const attData = attendance ?? [];
  const satData = satisfactionRaw ?? [];
  const currentData = which === 'attendance' ? attData : satData;
  const totalPages = Math.max(1, Math.ceil(currentData.length / PAGE_SIZE));
  const safePage = Math.min(page, totalPages - 1);

  const pageData = useMemo(
    () => currentData.slice(safePage * PAGE_SIZE, (safePage + 1) * PAGE_SIZE),
    [currentData, safePage],
  );

  // 만족도 테이블용 동적 헤더
  const satHeaders = useMemo(
    () => (satData.length > 0 ? Object.keys(satData[0]) : []),
    [satData],
  );

  return (
    <div className="panel">
      {/* 출석 / 만족도 세그먼트 전환 */}
      <div className="preview-seg-wrap">
        <div className="preview-seg">
          <button
            className={'preview-seg-btn' + (which === 'attendance' ? ' active' : '')}
            onClick={() => switchTab('attendance')}
          >
            출석 원본
            <span className="cnt">{attData.length}</span>
          </button>
          <button
            className={'preview-seg-btn' + (which === 'satisfaction' ? ' active' : '')}
            onClick={() => switchTab('satisfaction')}
          >
            만족도 원본
            <span className="cnt">{satData.length}</span>
          </button>
        </div>
        <span className="preview-total">총 {currentData.length}행</span>
      </div>

      {currentData.length === 0 ? (
        <div className="empty">데이터가 없습니다.</div>
      ) : (
        <>
          <div className="tbl-wrap">
            {which === 'attendance' ? (
              <table>
                <thead>
                  <tr>
                    <th>연번</th>
                    <th style={{ textAlign: 'left' }}>과목명</th>
                    <th style={{ textAlign: 'left' }}>이름</th>
                    <th>성별</th>
                    <th>생년월일</th>
                    <th>수업일</th>
                    <th>출석일</th>
                    <th>출석률</th>
                    <th>수료여부</th>
                  </tr>
                </thead>
                <tbody>
                  {(pageData as StructuredAttendanceItem[]).map((item, i) => (
                    <tr key={i}>
                      <td>{String(item.연번 ?? '')}</td>
                      <td style={{ textAlign: 'left' }}>{item.과목명}</td>
                      <td style={{ textAlign: 'left' }}>{item.이름}</td>
                      <td>{item.성별}</td>
                      <td>{String(item.생년월일 ?? '')}</td>
                      <td>{item.수업일}</td>
                      <td>{item.출석일}</td>
                      <td>{pct(item.출석률)}</td>
                      <td>{String(item.수료여부 ?? '')}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            ) : (
              <table>
                <thead>
                  <tr>
                    {satHeaders.map((h) => (
                      <th key={h}>{h}</th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {(pageData as RawRow[]).map((row, i) => (
                    <tr key={i}>
                      {satHeaders.map((h) => (
                        <td key={h}>{String(row[h] ?? '')}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            )}
          </div>

          {/* 페이지네이션 */}
          {totalPages > 1 && (
            <div className="pager">
              <button
                className="btn ghost pager-btn"
                onClick={() => setPage((p) => Math.max(0, p - 1))}
                disabled={safePage === 0}
              >
                ← 이전
              </button>
              <span className="pager-info">
                {safePage + 1} / {totalPages}
              </span>
              <button
                className="btn ghost pager-btn"
                onClick={() => setPage((p) => Math.min(totalPages - 1, p + 1))}
                disabled={safePage === totalPages - 1}
              >
                다음 →
              </button>
            </div>
          )}
        </>
      )}
    </div>
  );
}
