import { useState, useMemo } from 'react';
import type { Diagnostic } from '../types/analysis';
import { Ic } from './icons';

/* ================================================================
   진단 패널 — 파싱/분석 중 발생한 경고·정보를 카테고리별로 표시
   ================================================================ */

const CATEGORY_LABELS: Record<string, string> = {
  'skipped-row': '행 제외',
  'normalize-failed': '정규화 실패',
  'rate-format': '출석률 형식',
  'column-fallback': '칼럼 폴백',
  'integrity': '데이터 무결성',
  'bucket-mismatch': '집계 불일치',
  'value-coercion': '값 변환',
};

interface DiagnosticsPanelProps {
  diagnostics: Diagnostic[];
}

export function DiagnosticsPanel({ diagnostics }: DiagnosticsPanelProps) {
  const [open, setOpen] = useState(false);

  const warnings = useMemo(() => diagnostics.filter((d) => d.severity === 'warning'), [diagnostics]);
  const infos = useMemo(() => diagnostics.filter((d) => d.severity === 'info'), [diagnostics]);

  // 카테고리별 그룹핑
  const groups = useMemo(() => {
    const map = new Map<string, Diagnostic[]>();
    // warning 먼저, info 나중
    [...warnings, ...infos].forEach((d) => {
      const list = map.get(d.category) ?? [];
      list.push(d);
      map.set(d.category, list);
    });
    return map;
  }, [warnings, infos]);

  if (diagnostics.length === 0) {
    return (
      <div className="diag diag-clean-wrap">
        <span className="diag-clean-ic">{Ic.check}</span>
        <span className="diag-clean-msg">검증 결과 이상 없음</span>
      </div>
    );
  }

  return (
    <div className="diag">
      <button
        className="diag-head"
        onClick={() => setOpen((v) => !v)}
        aria-expanded={open}
      >
        <span className="diag-title">데이터 검증 결과</span>
        <span className="diag-badges">
          {warnings.length > 0 && (
            <span className="diag-badge warn">경고 {warnings.length}</span>
          )}
          {infos.length > 0 && (
            <span className="diag-badge info">정보 {infos.length}</span>
          )}
        </span>
        <span className="diag-toggle">{open ? Ic.chevronUp : Ic.chevronDown}</span>
      </button>

      {open && (
        <div className="diag-body">
          {[...groups.entries()].map(([cat, items]) => (
            <div key={cat} className="diag-group">
              <h5 className="diag-group-title">{CATEGORY_LABELS[cat] ?? cat}</h5>
              {items.map((d) => (
                <div key={d.code} className={`diag-item diag-item-${d.severity}`}>
                  <span className="diag-ic">
                    {d.severity === 'warning' ? Ic.alert : Ic.info}
                  </span>
                  <div className="diag-content">
                    <span className="diag-msg">{d.message}</span>
                    {d.count > 1 && (
                      <span className="diag-count">{d.count}건</span>
                    )}
                    {d.samples && d.samples.length > 0 && (
                      <div className="diag-samples">
                        {d.samples.map((s, i) => (
                          <span key={i} className="diag-sample">{s}</span>
                        ))}
                      </div>
                    )}
                  </div>
                </div>
              ))}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}
