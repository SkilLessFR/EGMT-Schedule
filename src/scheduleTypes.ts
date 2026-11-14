export type ComplianceTier = 0 | 1 | 2 | 3;

export type RelaxableRule =
  | 'NIGHT_OVER_2'
  | 'SHORT_BLOCK_2'
  | 'SHORT_REST_1'
  | 'COVERAGE_GAP';

export interface RelaxationRecord {
  tier: ComplianceTier;
  rule: RelaxableRule;
  isoDate: string;
  employee?: string;
  note: string;
}

export interface ValidationOptions {
  tier?: ComplianceTier;
  relaxations?: RelaxationRecord[];
}
