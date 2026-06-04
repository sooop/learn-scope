/* ================================================================
   수료 기준 설정 모달  (analyzer.html:869-891)
   ================================================================ */

interface CompletionRateModalProps {
  rates: Record<string, number>;
  onChange: (subject: string, rate: number) => void;
  onClose: () => void;
}

export function CompletionRateModal({ rates, onChange, onClose }: CompletionRateModalProps) {
  return (
    <div className="overlay" onClick={onClose}>
      <div className="modal" onClick={(e) => e.stopPropagation()}>
        <header>
          <h3>과목별 수료 기준 설정</h3>
          <p>출석률이 기준 이상이면 수료로 집계됩니다. 변경하면 출석 분석이 즉시 재계산됩니다.</p>
        </header>
        <div className="body">
          {Object.entries(rates).map(([subject, rate]) => (
            <div className="rate-row" key={subject}>
              <span className="nm">{subject}</span>
              <select
                value={Math.round(rate * 100)}
                onChange={(e) => onChange(subject, parseInt(e.target.value) / 100)}
              >
                {[30, 40, 50, 60, 70, 80, 90, 100].map((p) => (
                  <option key={p} value={p}>
                    {p}%
                  </option>
                ))}
              </select>
            </div>
          ))}
        </div>
        <footer>
          <button className="btn primary" onClick={onClose}>
            완료
          </button>
        </footer>
      </div>
    </div>
  );
}
