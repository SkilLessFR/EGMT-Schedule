import type { RosterData } from '../types';
import type {
  ComplianceTier,
  RelaxableRule,
  RelaxationRecord,
  ValidationOptions
} from '../scheduleTypes';
export type {
  RosterData,
  ComplianceTier,
  RelaxableRule,
  RelaxationRecord,
  ValidationOptions
};

export interface ScheduleBuilderOptions {
  seed?: number;
  targetWorkingDays?: number;
  minMorning?: number;
  minAfternoon?: number;
  minNight?: number;
  targetOffDays?: number;
  maxConsecutiveWork?: number;
  excludedEmployees?: string[];
  timeBudgetMs?: number;
  minTier?: ComplianceTier;
  maxTier?: ComplianceTier;
}

export interface DayCoverage {
  isoDate: string;
  dayOfWeek: number; // 0 = Sun, 1 = Mon ...
  isWeekend: boolean;
  M: number;
  rotM?: number;
  A: number;
  N: number;
  MID: number;
  OFF: number;
  ABS?: number;
  totalWorking: number;
  isUnderstaffed: boolean;
  warnings: string[];
}

export interface EmployeeMonthStats {
  employee: string;
  M: number;
  A: number;
  N: number;
  MID: number;
  OFF: number;
  H8: number;
  ABS?: number;
  totalWorkingDays: number;
  targetWorkingDays: number;
  workDaysDiff: number;
  maxStreak: number;
  violations: string[];
}

export interface ScheduleValidationReport {
  isFullyCompliant: boolean;
  totalViolations: number;
  understaffedDaysCount: number;
  nightOverstaffedDaysCount: number;
  streakViolationCount: number;
  nightRestViolationCount: number;
  excessiveRestViolationCount: number;
  mixedBlockViolationCount: number;
  workDaysMismatchCount: number;
  fixedStaffViolationCount?: number;
  dailyCoverage: Record<string, DayCoverage>;
  employeeStats: Record<string, EmployeeMonthStats>;
  warnings: string[];
}

export interface ReadjustmentAdjustment {
  isoDate: string;
  employee: string;
  fromShift: string;
  toShift: string;
  reason: string;
  absColleague: string;
}

export interface ReadjustmentProposal {
  hasDeficits: boolean;
  totalDeficits: number;
  adjustments: ReadjustmentAdjustment[];
  unresolvedDeficits: Array<{
    isoDate: string;
    shift: 'M' | 'A' | 'N';
    needed: number;
    available: number;
    reason: string;
  }>;
  updatedRows: Record<string, Record<string, string>>;
  validation: ScheduleValidationReport;
  tier?: ComplianceTier;
  relaxations?: RelaxationRecord[];
}

export interface GeneratedScheduleResult {
  month: number;
  year: number;
  daysCount: number;
  targetWorkingDays: number;
  dateColumns: { index: number; date: Date; isoDate: string }[];
  rows: Record<string, Record<string, string>>;
  validation: ScheduleValidationReport;
  seed: number;
  isHistorical?: boolean;
  excludedEmployees?: string[];
  tier?: ComplianceTier;
  relaxations?: RelaxationRecord[];
  needsManualReview?: boolean;
  elapsedMs?: number;
}
