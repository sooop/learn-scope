import type { DistributionEntry } from '../types/analysis';

/* ================================================================
   공통 분포 테이블  (analyzer.html:729-751)
   ================================================================ */

export function pct(x: number): string {
  return (x * 100).toFixed(1) + '%';
}

export function pct100(x: number): string {
  return x.toFixed(1) + '%';
}

interface DistTableProps {
  title: string;
  dist: Record<string, DistributionEntry>;
  denomLabel?: string;
  showTotal?: boolean;
  /** 교차검증: 버킷 합 비교 기준 전체 인원. 불일치 시 '미분류' 행 표시 */
  expectedTotal?: number;
  /** 교차검증 미분류 행 라벨 (기본: '미분류') */
  unclassifiedLabel?: string;
  /** 합계 계산에서 제외할 키 목록 (기본: ['합계']) */
  sumExcludeKeys?: string[];
}

export function DistTable({
  title,
  dist,
  denomLabel,
  showTotal,
  expectedTotal,
  unclassifiedLabel = '미분류',
  sumExcludeKeys,
}: DistTableProps) {
  const entries = Object.entries(dist);
  const max = Math.max(...entries.map(([, v]) => v.비율), 1);
  const totalCount = entries.reduce((s, [, v]) => s + v.명수, 0);
  const totalRate = entries.reduce((s, [, v]) => s + v.비율, 0);

  // 교차검증: '합계' 키를 제외한 버킷 합 계산
  const excludeSet = new Set(sumExcludeKeys ?? ['합계']);
  const bucketCount = entries
    .filter(([k]) => !excludeSet.has(k))
    .reduce((s, [, v]) => s + v.명수, 0);
  const unclassifiedCount =
    expectedTotal != null && expectedTotal > bucketCount ? expectedTotal - bucketCount : 0;

  return (
    <div className="section">
      <h4>
        {title}
        {denomLabel ? ` · ${denomLabel}` : ''}
      </h4>
      <div className="tbl-wrap">
        <table>
          <thead>
            <tr>
              <th>구분</th>
              <th>명수</th>
              <th>비율</th>
            </tr>
          </thead>
          <tbody>
            {entries.map(([k, v]) => (
              <tr key={k}>
                <td>{k}</td>
                <td>{v.명수}</td>
                <td>
                  {pct100(v.비율)}
                  <span className="bar" style={{ width: (v.비율 / max) * 60 + 'px' }} />
                </td>
              </tr>
            ))}
            {unclassifiedCount > 0 && (
              <tr className="diag-row">
                <td>{unclassifiedLabel}</td>
                <td>{unclassifiedCount}</td>
                <td>—</td>
              </tr>
            )}
            {showTotal && (
              <tr className="total">
                <td>합계</td>
                <td>{totalCount}</td>
                <td>{pct100(totalRate)}</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
