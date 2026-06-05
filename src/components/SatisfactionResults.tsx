import type { SatisfactionAnalysis } from '../types/analysis';
import { AGE_KEYS } from '../lib/constants';

/* ================================================================
   만족도 결과 표시  (analyzer.html:793-867)
   ================================================================ */

interface SatisfactionResultsProps {
  results: SatisfactionAnalysis | null;
}

export function SatisfactionResults({ results }: SatisfactionResultsProps) {
  if (!results) return null;
  const avgList = Object.values(results.satisfactionAverages);
  if (avgList.length === 0)
    return <div className="empty">만족도 데이터가 없습니다.</div>;

  const firstSubject = avgList[0];
  const questions = Object.keys(firstSubject.scores).filter((q) => q !== '전체');

  const overallAll = avgList.map((s) => Number(s.scores['전체']) || 0);
  const grandAvg = overallAll.length
    ? overallAll.reduce((a, b) => a + b, 0) / overallAll.length
    : 0;

  return (
    <div className="panel">
      <div className="stat-row">
        <div className="stat">
          <div className="n">{results.totalSubjects}</div>
          <div className="l">설문 과목 수</div>
        </div>
        <div className="stat">
          <div className="n">{results.totalResponses}</div>
          <div className="l">총 응답 건수</div>
        </div>
        <div className="stat">
          <div className="n">{grandAvg.toFixed(2)}</div>
          <div className="l">
            전체 평균 <b>만족도</b> (5점)
          </div>
        </div>
      </div>

      <div className="section">
        <h4>과목별 평균 만족도</h4>
        <div className="tbl-wrap">
          <table>
            <thead>
              <tr>
                <th>과목명</th>
                <th className="sum">전체</th>
                {questions.map((q) => (
                  <th key={q}>{q}</th>
                ))}
              </tr>
            </thead>
            <tbody>
              {avgList.map((s) => (
                <tr key={s.subject}>
                  <td>
                    {s.구분 && <span className="category-badge">{s.구분}</span>}
                    {s.subject}
                  </td>
                  <td className="sum">{s.scores['전체']}</td>
                  {questions.map((q) => (
                    <td key={q}>{s.scores[q]}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <div className="section">
        <h4>응답자 특성 (성별 · 연령)</h4>
        <div className="tbl-wrap">
          <table>
            <thead>
              <tr>
                <th>과목명</th>
                <th>남</th>
                <th>여</th>
                <th>19↓</th>
                <th>20대</th>
                <th>30대</th>
                <th>40대</th>
                <th>50대</th>
                <th>60↑</th>
                <th className="sum">응답</th>
              </tr>
            </thead>
            <tbody>
              {Object.values(results.respondentCharacteristics).map((rc) => {
                const total = rc.gender.남성 + rc.gender.여성;
                const cat = results.subjectCategories[rc.subject];
                return (
                  <tr key={rc.subject}>
                    <td>
                      {cat && <span className="category-badge">{cat}</span>}
                      {rc.subject}
                    </td>
                    <td>{rc.gender.남성}</td>
                    <td>{rc.gender.여성}</td>
                    {AGE_KEYS.map((k) => (
                      <td key={k}>{rc.age[k] || 0}</td>
                    ))}
                    <td className="sum">{total}</td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}
