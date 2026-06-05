import type { AttendanceAnalysis, SubjectMerge } from '../types/analysis';
import { findCategory } from '../lib/subject-utils';
import { DistTable, pct } from './DistTable';

/* ================================================================
   출석 분석 결과 표시  (analyzer.html:753-791)
   ================================================================ */

interface AttendanceResultsProps {
  analysis: AttendanceAnalysis | null;
  subjectCategories?: Record<string, string>;
  merges?: SubjectMerge[];
}

export function AttendanceResults({ analysis, subjectCategories = {}, merges }: AttendanceResultsProps) {
  if (!analysis) return null;
  const sr = analysis.subjectResults;
  const maxAtt = Math.max(...sr.map((s) => s.평균출석률), 0.01);

  // 과목명 → 병합 정보 역매핑 (툴팁용)
  const mergeMap = Object.fromEntries(
    (merges ?? []).flatMap((m) => m.members.map((mb) => [mb, m])),
  );

  return (
    <div className="panel">
      {merges && merges.length > 0 && (
        <div
          className="merge-notice"
          style={{
            background: 'var(--c-surface, #f5f5f4)',
            border: '1px solid var(--c-border, #d6d3d1)',
            borderRadius: '6px',
            padding: '10px 14px',
            marginBottom: '16px',
            fontSize: '0.875rem',
          }}
        >
          <strong>유사 표기 {merges.length}건이 통합되었습니다</strong>
          <ul style={{ margin: '6px 0 0', paddingLeft: '1.25rem' }}>
            {merges.map((m) => (
              <li key={m.canonical}>
                <strong>{m.canonical}</strong> ←{' '}
                {m.members.filter((mb) => mb !== m.canonical).join(', ')}
              </li>
            ))}
          </ul>
        </div>
      )}
      <div className="stat-row">
        <div className="stat">
          <div className="n">{analysis.totalStudents}</div>
          <div className="l">총 수강 등록 건수</div>
        </div>
        <div className="stat">
          <div className="n">{analysis.uniqueStudents}</div>
          <div className="l">고유 수강생 수</div>
        </div>
        <div className="stat">
          <div className="n">{sr.length}</div>
          <div className="l">개설 과목 수</div>
        </div>
        <div className="stat">
          <div className="n">{sr.reduce((a, s) => a + s.수료인원, 0)}</div>
          <div className="l">
            총 <b>수료 인원</b>
          </div>
        </div>
      </div>

      <div className="section">
        <h4>과목별 출석 · 수료 현황</h4>
        <div className="tbl-wrap">
          <table>
            <thead>
              <tr>
                <th>과목명</th>
                <th>수강인원</th>
                <th>평균출석률</th>
                <th>수료인원</th>
                <th className="sum">수료율</th>
              </tr>
            </thead>
            <tbody>
              {sr.map((item) => {
                const cat = findCategory(item.과목명, subjectCategories);
                return (
                <tr key={item.과목명}>
                  <td
                    title={
                      mergeMap[item.과목명]?.members.filter((mb) => mb !== item.과목명).length
                        ? `통합: ${mergeMap[item.과목명].members.join(', ')}`
                        : undefined
                    }
                  >
                    {cat && <span className="category-badge">{cat}</span>}
                    {item.과목명}
                  </td>
                  <td>{item.수강인원}</td>
                  <td>
                    {pct(item.평균출석률)}
                    <span
                      className="bar"
                      style={{ width: (item.평균출석률 / maxAtt) * 50 + 'px' }}
                    />
                  </td>
                  <td>{item.수료인원}</td>
                  <td className="sum">{pct(item.수료율)}</td>
                </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>

      <DistTable
        title="성별 분포"
        dist={analysis.genderDistribution}
        expectedTotal={analysis.totalStudents}
        unclassifiedLabel="성별 미상"
        sumExcludeKeys={['합계']}
      />
      <DistTable
        title="수강 강좌수별 분포"
        dist={analysis.courseCountDistribution}
        denomLabel={`고유 수강생 ${analysis.uniqueStudents}명 기준`}
        expectedTotal={analysis.uniqueStudents}
      />
      <DistTable
        title="연령 분포"
        dist={analysis.ageDistribution}
        denomLabel={`등록 건수 ${analysis.totalStudents}건 기준`}
        expectedTotal={analysis.totalStudents}
        showTotal
      />
    </div>
  );
}
