import type { AttendanceAnalysis } from '../types/analysis';
import { findCategory } from '../lib/subject-utils';
import { DistTable, pct } from './DistTable';

/* ================================================================
   출석 분석 결과 표시  (analyzer.html:753-791)
   ================================================================ */

interface AttendanceResultsProps {
  analysis: AttendanceAnalysis | null;
  subjectCategories?: Record<string, string>;
}

export function AttendanceResults({ analysis, subjectCategories = {} }: AttendanceResultsProps) {
  if (!analysis) return null;
  const sr = analysis.subjectResults;
  const maxAtt = Math.max(...sr.map((s) => s.평균출석률), 0.01);
  return (
    <div className="panel">
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
                  <td>
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

      <DistTable title="성별 분포" dist={analysis.genderDistribution} />
      <DistTable
        title="수강 강좌수별 분포"
        dist={analysis.courseCountDistribution}
        denomLabel={`고유 수강생 ${analysis.uniqueStudents}명 기준`}
      />
      <DistTable
        title="연령 분포"
        dist={analysis.ageDistribution}
        denomLabel={`등록 건수 ${analysis.totalStudents}건 기준`}
        showTotal
      />
    </div>
  );
}
