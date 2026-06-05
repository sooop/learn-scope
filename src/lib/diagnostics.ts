import type { Diagnostic, DiagnosticSeverity, DiagnosticCategory } from '../types/analysis';

/* ================================================================
   진단 수집기 — 파싱/분석 함수 내부에서 지역 인스턴스로만 사용
   순수 함수 유지: 외부 상태를 변경하지 않고 결과와 함께 반환됨
   ================================================================ */

const SAMPLE_LIMIT = 5;

export class DiagnosticCollector {
  private map = new Map<string, Diagnostic>();

  add(
    code: string,
    severity: DiagnosticSeverity,
    category: DiagnosticCategory,
    message: string,
    sample?: string,
  ): void {
    const existing = this.map.get(code);
    if (existing) {
      existing.count++;
      if (
        sample !== undefined &&
        sample !== '' &&
        (existing.samples?.length ?? 0) < SAMPLE_LIMIT
      ) {
        existing.samples = [...(existing.samples ?? []), sample];
      }
    } else {
      this.map.set(code, {
        code,
        severity,
        category,
        message,
        count: 1,
        samples: sample !== undefined && sample !== '' ? [sample] : undefined,
      });
    }
  }

  build(): Diagnostic[] {
    return [...this.map.values()];
  }
}
