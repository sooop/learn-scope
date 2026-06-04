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
}

export function DistTable({ title, dist, denomLabel }: DistTableProps) {
  const entries = Object.entries(dist);
  const max = Math.max(...entries.map(([, v]) => v.비율), 1);
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
          </tbody>
        </table>
      </div>
    </div>
  );
}
