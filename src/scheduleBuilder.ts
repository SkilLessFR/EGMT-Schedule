// src/scheduleBuilder.ts
import type ExcelJS from 'exceljs';
import type { RosterData } from './types';
import type {
  ComplianceTier,
  RelaxableRule,
  RelaxationRecord,
  ValidationOptions
} from './scheduleTypes';
export type {
  ComplianceTier,
  RelaxableRule,
  RelaxationRecord,
  ValidationOptions
};
import { shiftKey, ROMANIAN_HOLIDAYS } from './scheduleUtils';
import {
  getTierConfig,
  isCoverageFeasibleForTier,
  tierCovPenalty,
  canSatisfyDailyCapacityForTier
} from './relaxation';

export const FIXED_MID_EMPLOYEES = [
  'Palade Alexandru-Ionut'
];

export const FIXED_WEEKDAY_MA_EMPLOYEES = [
  'Tir George Cristian'
];

export const NON_ROTATING_EMPLOYEES = [
  ...FIXED_MID_EMPLOYEES,
  ...FIXED_WEEKDAY_MA_EMPLOYEES
];

export function getIsoWeekNumber(date: Date): number {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
}

export { ROMANIAN_HOLIDAYS };
export const ROMANIAN_HOLIDAYS_2026 = ROMANIAN_HOLIDAYS;

// Extra non-statutory dates on which fixed staff (Palade, Tir George) are scheduled OFF
export const FIXED_STAFF_EXEMPT_OFF_DATES = new Set<string>();

export const MONTHLY_WORKING_DAYS_2026: Record<number, number> = {
  0: 18, // Jan (22 weekdays minus Jan 1, 2, 6, 7 holidays)
  1: 20, // Feb
  2: 22, // Mar
  3: 20, // Apr
  4: 20, // May
  5: 21, // Jun
  6: 23, // Jul
  7: 21, // Aug
  8: 22, // Sep
  9: 22, // Oct
  10: 20, // Nov (20 in 11 Noiembrie.xlsx)
  11: 21  // Dec (21 statutory legal working days for 24/7 operations in 2026)
};

export function countLegalWorkingDaysInMonth(month: number, year: number): number {
  if (year === 2026 && MONTHLY_WORKING_DAYS_2026[month] !== undefined) {
    return MONTHLY_WORKING_DAYS_2026[month];
  }
  let count = 0;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month, day);
    const dayOfWeek = date.getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) continue;
    const isoDate = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
    if (ROMANIAN_HOLIDAYS.has(isoDate)) continue;
    count += 1;
  }
  return count;
}

function isFixedStaffOffDate(date: Date, isoDate: string): boolean {
  const dow = date.getDay();
  return dow === 0 || dow === 6 || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
}

/**
 * Checks whether any block of consecutive working shifts in a sequence contains incompatible shift types.
 * M and A (day shifts) can be combined in the same block, but Day and Night cannot be combined in the same block.
 */
export function hasMixedBlocks(sequence: string[]): boolean {
  const isDay = (s: string) => s === 'M' || s === 'A' || s === 'MID';
  let curType: string | null = null;
  for (const s of sequence) {
    const k = s === 'OFF' || s === 'H8' || s === 'HUP' || s === 'ABS' || s === 'AD' ? 'OFF' : s;
    if (k !== 'OFF') {
      if (curType === null) {
        curType = k;
      } else if (curType !== k) {
        if (isDay(curType) && isDay(k)) {
          curType = k;
        } else {
          return true;
        }
      }
    } else {
      curType = null;
    }
  }
  return false;
}

/**
 * Extracts only the active continuous work streak immediately preceding day 1 of the new month.
 * If the employee ended on OFF or break, returns empty array.
 */
export function getActiveTailBlock(history: string[]): string[] {
  const tail: string[] = [];
  for (let i = history.length - 1; i >= 0; i--) {
    const s = history[i];
    if (s === 'OFF' || s === 'H8' || s === 'HUP' || s === 'ABS' || s === 'AD') {
      break;
    }
    tail.unshift(s);
  }
  return tail;
}

/**
 * Extracts only the active continuous rest streak immediately preceding day 1 of the new month.
 * If the employee ended on a work shift, returns empty array.
 */
export function getActiveTailRest(history: string[]): string[] {
  const tail: string[] = [];
  for (let i = history.length - 1; i >= 0; i--) {
    const s = history[i];
    if (s !== 'OFF' && s !== 'H8' && s !== 'HUP' && s !== 'ABS' && s !== 'AD') {
      break;
    }
    tail.unshift('OFF');
  }
  return tail;
}

const DAG_NEXT_SHIFT_ACT: readonly number[] = [1, 2, 0]; // A(0)->M(1), M(1)->N(2), N(2)->A(0)
const DAG_LEAVE_CODES = new Set(['ABS', 'H8', 'CM']);

function encodeStage3DagState(w: number, sh: number, ml: number, sb: number): number {
  return ((w * 4 + sh) * 13 + ml) * 2 + sb;
}

function encodeGenDagState(w: number, sh: number, ml: number, offS: number, sb: number): number {
  return (((w * 4 + sh) * 13 + ml) * 4 + offS) * 2 + sb;
}

function canDoShift(dayActMasks: Uint8Array[], eIdx: number, d: number, act: number): boolean {
  return (dayActMasks[eIdx][d] & (1 << act)) !== 0;
}

function buildDagDayActMasks(
  dags: ReadonlyArray<{
    feasible: boolean;
    layers: ReadonlyArray<ReadonlyArray<ReadonlyArray<{ act: number }>>>;
  }>,
  daysCount: number
): Uint8Array[] {
  return dags.map((dag) => {
    const masks = new Uint8Array(daysCount);
    if (!dag.feasible) return masks;
    for (let d = 0; d < daysCount; d++) {
      let mask = 0;
      const layer = dag.layers[d];
      for (let u = 0; u < layer.length; u++) {
        const edges = layer[u];
        for (let ei = 0; ei < edges.length; ei++) {
          mask |= 1 << edges[ei].act;
        }
      }
      masks[d] = mask;
    }
    return masks;
  });
}

function isDagPathValid(
  dag: {
    layers: ReadonlyArray<ReadonlyArray<ReadonlyArray<{ act: number; nextNode: number }>>>;
  },
  path: number[],
  daysCount: number
): boolean {
  let u = 0;
  for (let d = 0; d < daysCount; d++) {
    const edges = dag.layers[d][u];
    const act = path[d];
    let found = -1;
    for (let ei = 0; ei < edges.length; ei++) {
      if (edges[ei].act === act) {
        found = edges[ei].nextNode;
        break;
      }
    }
    if (found === -1) return false;
    u = found;
  }
  return true;
}

function solveDagExtremeCount(
  dag: {
    maxNodes: number;
    layers: ReadonlyArray<ReadonlyArray<ReadonlyArray<{ act: number; nextNode: number }>>>;
  },
  daysCount: number,
  targetAct: number,
  sign: 1 | -1
): number {
  let dpNext = new Float64Array(Math.max(1, dag.maxNodes));
  let dpCur = new Float64Array(Math.max(1, dag.maxNodes));
  for (let d = daysCount - 1; d >= 0; d--) {
    const layer = dag.layers[d];
    const isLast = d === daysCount - 1;
    for (let u = 0; u < layer.length; u++) {
      const edges = layer[u];
      let bestVal = Infinity;
      for (let ei = 0; ei < edges.length; ei++) {
        const edge = edges[ei];
        const val = (edge.act === targetAct ? sign : 0) + (isLast ? 0 : dpNext[edge.nextNode]);
        if (val < bestVal) bestVal = val;
      }
      dpCur[u] = bestVal;
    }
    const tmp = dpNext;
    dpNext = dpCur;
    dpCur = tmp;
  }
  return sign === 1 ? dpNext[0] : -dpNext[0];
}

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


/**
 * Generate dates for a given month and year
 */
export function buildMonthDateColumns(year: number, month: number) {
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const dateColumns: { index: number; date: Date; isoDate: string }[] = [];

  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month, day);
    const dayStr = String(day).padStart(2, '0');
    const monthStr = String(month + 1).padStart(2, '0');
    const isoDate = `${year}-${monthStr}-${dayStr}`;
    dateColumns.push({ index: day - 1, date, isoDate });
  }

  return dateColumns;
}

/**
 * Inspect last shifts of previous month to ensure seamless continuity
 */
function getEmployeeHistoryTail(
  roster: RosterData | null,
  employee: string,
  year: number,
  month: number,
  count: number = 6
): string[] {
  if (!roster || !roster.dateColumns || roster.dateColumns.length === 0) return [];
  const targetStartIso = `${year}-${String(month + 1).padStart(2, '0')}-01`;
  const dayBefore = new Date(year, month, 0); // last day of preceding month
  const dayBeforeIso = `${dayBefore.getFullYear()}-${String(dayBefore.getMonth() + 1).padStart(2, '0')}-${String(dayBefore.getDate()).padStart(2, '0')}`;

  const precedingDates = roster.dateColumns
    .filter((col) => col.isoDate < targetStartIso)
    .sort((a, b) => a.isoDate.localeCompare(b.isoDate));

  if (precedingDates.length === 0) return [];

  // Check if roster contains the day immediately before targetStart
  // If there is a gap (e.g. generating future months when roster only has earlier data), return []
  const lastPreceding = precedingDates[precedingDates.length - 1];
  if (lastPreceding.isoDate !== dayBeforeIso) {
    return [];
  }

  const tailCols = precedingDates.slice(-count);
  return tailCols.map((col) => {
    const raw = roster.rows?.[employee]?.[col.isoDate];
    const k = shiftKey(raw ?? 'OFF');
    // HUP is a disciplinary break; H8 is holiday; neither is a continuous work streak
    if (k === 'HUP' || k === 'H8' || k === 'ABS' || k === 'AD') return 'OFF';
    return k;
  });
}

/**
 * Extracts raw contiguous <= 6-day prior history ending on day before target month,
 * preserving ABS/H8/CM codes to match validateScheduleCompliance boundary checks.
 */
function extractRawHistoryTail(
  roster: RosterData | null,
  employee: string,
  year: number,
  month: number,
  count: number = 6
): string[] {
  if (!roster) return [];
  const targetStartIso = `${year}-${String(month + 1).padStart(2, '0')}-01`;
  const knownDates = new Set<string>();
  for (const dc of roster.dateColumns || []) {
    if (dc?.isoDate) knownDates.add(dc.isoDate);
  }
  if (roster.rows) {
    for (const rEmp of Object.keys(roster.rows)) {
      const empRow = roster.rows[rEmp];
      if (empRow && typeof empRow === 'object') {
        for (const iso of Object.keys(empRow)) knownDates.add(iso);
      }
    }
  }
  const [y, m, d] = targetStartIso.split('-').map(Number);
  const tailIsos: string[] = [];
  for (let offset = 1; offset <= count; offset++) {
    const dt = new Date(Date.UTC(y, m - 1, d));
    dt.setUTCDate(dt.getUTCDate() - offset);
    const prevIso = dt.toISOString().slice(0, 10);
    if (knownDates.has(prevIso)) {
      tailIsos.unshift(prevIso);
    } else {
      break;
    }
  }
  return tailIsos.map((iso) => {
    const raw = roster.rows?.[employee]?.[iso];
    const s = raw ? shiftKey(String(raw)) : 'OFF';
    return s === '' ? 'OFF' : s;
  });
}

/**
 * Kuhn-Munkres (Hungarian algorithm) for optimal O(N^3) minimum cost bipartite matching.
 * Solves matching an N-length worker list to N slots in < 0.1ms.
 */
function minCostMatching(costMatrix: number[][]): number[] {
  const n = costMatrix.length;
  if (n === 0) return [];
  const u = new Array(n + 1).fill(0);
  const v = new Array(n + 1).fill(0);
  const p = new Array(n + 1).fill(0);
  const way = new Array(n + 1).fill(0);

  for (let i = 1; i <= n; i++) {
    p[0] = i;
    let j0 = 0;
    const minv = new Array(n + 1).fill(Infinity);
    const used = new Array(n + 1).fill(false);
    do {
      used[j0] = true;
      const i0 = p[j0];
      let delta = Infinity;
      let j1 = 0;
      for (let j = 1; j <= n; j++) {
        if (!used[j]) {
          const cur = costMatrix[i0 - 1][j - 1] - u[i0] - v[j];
          if (cur < minv[j]) {
            minv[j] = cur;
            way[j] = j0;
          }
          if (minv[j] < delta) {
            delta = minv[j];
            j1 = j;
          }
        }
      }
      for (let j = 0; j <= n; j++) {
        if (used[j]) {
          u[p[j]] += delta;
          v[j] -= delta;
        } else {
          minv[j] -= delta;
        }
      }
      j0 = j1;
    } while (p[j0] !== 0);

    do {
      const j1 = way[j0];
      p[j0] = p[j1];
      j0 = j1;
    } while (j0 !== 0);
  }

  const result = new Array(n);
  for (let j = 1; j <= n; j++) {
    if (p[j] > 0) {
      result[p[j] - 1] = j - 1;
    }
  }
  return result;
}

/**
 * Calculate cost for an employee starting at a given template offset
 */
export function evaluateOffsetCost(
  history: string[],
  template: string[],
  offset: number,
  employee: string,
  targetDaysCount: number,
  day1Anchor?: string | null,
  seed: number = 0,
  targetWorkingDays?: number
): number {
  let cost = 0;
  const tLen = template.length;
  const slotShifts: string[] = [];
  for (let d = 0; d < targetDaysCount; d++) {
    slotShifts.push(template[(offset + d) % tLen]);
  }

  const firstShift = slotShifts[0];
  const secondShift = slotShifts[1];

  // 1. Boundary streak check (combines tail from previous month + start of new month)
  let tailStreak = 0;
  for (let i = history.length - 1; i >= 0; i--) {
    if (history[i] !== 'OFF' && history[i] !== 'H8' && history[i] !== 'HUP') tailStreak++;
    else break;
  }

  let headStreak = 0;
  for (let i = 0; i < Math.min(5, targetDaysCount); i++) {
    if (slotShifts[i] !== 'OFF' && slotShifts[i] !== 'H8' && slotShifts[i] !== 'HUP') headStreak++;
    else break;
  }

  if (tailStreak >= 5) {
    // If employee reached 5 consecutive working days at end of previous month,
    // they MUST start the new month with at least 2 days of OFF!
    if (firstShift !== 'OFF' || (targetDaysCount > 1 && secondShift !== 'OFF')) {
      cost += 10000000;
    }
  } else if (tailStreak > 0) {
    const lastShift = history[history.length - 1];
    if (firstShift === 'OFF') {
      if (tailStreak < 3) {
        cost += 10000000; // work streak would be prematurely cut short
      } else {
        // Taking rest after >= 3 days of work is legal, but rest must be at least 2 days
        if (targetDaysCount > 1 && secondShift !== 'OFF') {
          cost += 10000000;
        }
      }
    } else {
      // Continuing the active work block into the new month: shift type must not change without rest
      if (firstShift !== lastShift) {
        cost += 10000000;
      } else {
        const combined = tailStreak + headStreak;
        if (combined < 3 || combined > 5) {
          cost += 10000000; // Block length must be 3..5 days
        }
        // Followed by at least 2 days of OFF
        if (headStreak < targetDaysCount && slotShifts[headStreak] !== 'OFF') cost += 5000000;
        if (headStreak + 1 < targetDaysCount && slotShifts[headStreak + 1] !== 'OFF') cost += 5000000;
      }
    }
  } else if (history.length > 0 && firstShift !== 'OFF') {
    // Previous month ended on rest: a new block starting Day 1 that ends in-month must be >= 3 days
    if (headStreak < 3 && headStreak < targetDaysCount) {
      cost += 10000000;
    }
  }

  // 2. Boundary rest streak check: rest periods must be at least 2 days, at most 3 days
  let tailOffStreak = 0;
  for (let i = history.length - 1; i >= 0; i--) {
    if (history[i] === 'OFF' || history[i] === 'H8' || history[i] === 'HUP') tailOffStreak++;
    else break;
  }
  let headOffStreak = 0;
  for (let i = 0; i < Math.min(5, targetDaysCount); i++) {
    if (slotShifts[i] === 'OFF' || slotShifts[i] === 'H8' || slotShifts[i] === 'HUP') headOffStreak++;
    else break;
  }

  // If employee was on day 1 of rest at end of previous month, Day 1 of new month MUST be OFF
  if (tailOffStreak === 1) {
    if (firstShift !== 'OFF') {
      cost += 10000000; // Minimum 2 days rest violation!
    }
  }

  // If employee already had >= 3 days of rest at end of previous month, adding OFF on Day 1 exceeds 3 days.
  // A 2-day tail rest + 1 OFF day on Day 1 yields a legal 3-day rest (soft preference against unless needed).
  if (tailOffStreak >= 3 && firstShift === 'OFF') {
    cost += 10000000;
  } else if (tailOffStreak === 2 && firstShift === 'OFF') {
    cost += 100;
  }

  if (tailOffStreak > 0 && headOffStreak > 0) {
    const boundaryOff = tailOffStreak + headOffStreak;
    if (boundaryOff > 3) {
      cost += 10000000 * (boundaryOff - 3); // Rest exceeded 3 days strictly forbidden
    }
  }

  // 3. Boundary night rest check:
  if (history.length > 0) {
    const lastShift = history[history.length - 1];
    if (lastShift === 'N' && (firstShift === 'M' || firstShift === 'A' || firstShift === 'MID')) {
      cost += 10000000; // STRICTLY FORBIDDEN!
    }
  }

  // 4. Anchor check if Day 01 was defined in previous month roster (never anchor to H8 or HUP)
  if (day1Anchor && day1Anchor !== 'OFF' && day1Anchor !== 'H8' && day1Anchor !== 'HUP') {
    const lastShift = history[history.length - 1];
    if (firstShift === day1Anchor && !(lastShift === 'N' && (firstShift === 'M' || firstShift === 'A'))) {
      cost -= 1500; // Reward matching anchor
    } else {
      cost += 800;
    }
  }

  // 5. Anti-repeat check & Shift Sequence continuity check:
  // The shift generation must strictly follow sequence A -> M -> N -> A
  const EXPECTED_NEXT_SHIFT: Record<string, string> = {
    A: 'M',
    M: 'N',
    N: 'A'
  };

  let lastWorkShift: string | null = null;
  for (let i = history.length - 1; i >= 0; i--) {
    const s = history[i];
    if (s && s !== 'OFF' && s !== 'H8' && s !== 'HUP' && s !== 'ABS' && s !== 'AD') {
      lastWorkShift = s;
      break;
    }
  }

  let firstWorkShift: string | null = null;
  for (let i = 0; i < slotShifts.length; i++) {
    const s = slotShifts[i];
    if (s !== 'OFF' && s !== 'H8' && s !== 'HUP') {
      firstWorkShift = s;
      break;
    }
  }

  if (lastWorkShift && firstWorkShift) {
    if (tailOffStreak > 0 || headOffStreak > 0 || slotShifts[0] === 'OFF') {
      // Starting a new block after rest: it MUST be the expected next shift in A -> M -> N -> A!
      const expectedNext = EXPECTED_NEXT_SHIFT[lastWorkShift];
      if (expectedNext) {
        if (firstWorkShift !== expectedNext) {
          cost += 10000000; // Strictly forbidden: breaks sequence A -> M -> N -> A
        }
      } else if (firstWorkShift === lastWorkShift) {
        cost += 5000000;
      }
    }
  }

  // 6. Target working days affinity: prefer offsets that naturally hit target working days
  if (targetWorkingDays !== undefined) {
    const rawWorkDays = slotShifts.filter((s) => s !== 'OFF' && s !== 'H8' && s !== 'HUP').length;
    if (rawWorkDays < targetWorkingDays) {
      cost += 500 * (targetWorkingDays - rawWorkDays);
    } else if (rawWorkDays > targetWorkingDays) {
      cost += 60 * (rawWorkDays - targetWorkingDays);
    }
  }

  // Seed variation factor
  const pseudo = ((employee.length * 37 + offset * 19 + seed * 97) % 100);
  cost += pseudo;

  return cost;
}

/**
 * Equalizes every employee so they work EXACTLY targetWorkingDays (e.g. 21 in Dec, 22 in Oct)
 */
export function equalizeExactWorkingDays(
  rows: Record<string, Record<string, string>>,
  dateColumns: { index: number; date: Date; isoDate: string }[],
  targetWorkingDays: number,
  priorRoster: RosterData | null,
  year: number,
  month: number,
  excludedEmployees: string[] = []
) {
  const excludedSet = new Set(excludedEmployees);
  const allEmployees = Object.keys(rows);
  const fixedMid = FIXED_MID_EMPLOYEES.filter((e) => allEmployees.includes(e) && !excludedSet.has(e));
  const fixedWeekdayMA = FIXED_WEEKDAY_MA_EMPLOYEES.filter((e) => allEmployees.includes(e) && !excludedSet.has(e));
  const rotating = allEmployees.filter((e) => !fixedMid.includes(e) && !fixedWeekdayMA.includes(e) && !excludedSet.has(e));

  // 1. Fixed Managers (Palade): Monday to Friday MID, Saturday and Sunday OFF, Romanian National Holidays OFF
  fixedMid.forEach((emp) => {
    dateColumns.forEach(({ date, isoDate }) => {
      const dow = date.getDay();
      const isWeekend = dow === 0 || dow === 6;
      const isOff = isWeekend || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
      if (isOff) {
        rows[emp][isoDate] = 'OFF';
      } else {
        rows[emp][isoDate] = 'MID';
      }
    });
  });

  // 2. Fixed Weekday M/A (Tir George Cristian): Monday to Friday M/A (alternating weeks), Saturday and Sunday OFF, Romanian National Holidays OFF
  fixedWeekdayMA.forEach((emp) => {
    dateColumns.forEach(({ date, isoDate }) => {
      const dow = date.getDay();
      const isWeekend = dow === 0 || dow === 6;
      const isOff = isWeekend || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
      if (isOff) {
        rows[emp][isoDate] = 'OFF';
      } else {
        const weekNum = getIsoWeekNumber(date);
        rows[emp][isoDate] = (weekNum % 2 === 1) ? 'M' : 'A';
      }
    });
  });

  const getWorkDays = (emp: string) => {
    let c = 0;
    dateColumns.forEach(({ isoDate }) => {
      const s = rows[emp][isoDate];
      if (s !== 'OFF' && s !== 'H8' && s !== 'HUP' && s !== 'ABS' && s !== 'AD') c++;
    });
    return c;
  };

  const employeeHistories: Record<string, string[]> = {};
  rotating.forEach((emp) => {
    employeeHistories[emp] = priorRoster
      ? getEmployeeHistoryTail(priorRoster, emp, year, month, 6)
      : [];
  });

  const getLastHistoryWorkShift = (emp: string): string | null => {
    const empHistory = employeeHistories[emp] ?? [];
    for (let i = empHistory.length - 1; i >= 0; i--) {
      const s = empHistory[i];
      if (s && s !== 'OFF' && s !== 'H8' && s !== 'HUP' && s !== 'ABS' && s !== 'AD') {
        return s;
      }
    }
    return null;
  };

  const getBlocks = (emp: string) => {
    const empHistory = employeeHistories[emp] ?? [];
    const empMonthShifts = dateColumns.map((col) => shiftKey(rows[emp]?.[col.isoDate] ?? 'OFF'));
    const activeTailWork = getActiveTailBlock(empHistory);
    const activeTailRest = activeTailWork.length === 0 ? getActiveTailRest(empHistory) : [];
    const activeTail = activeTailWork.length > 0 ? activeTailWork : activeTailRest;
    const full = [...activeTail, ...empMonthShifts];
    const blocks: { type: 'WORK' | 'REST'; shift: string; len: number }[] = [];
    for (const s of full) {
      const isRest = (s === 'OFF' || s === 'H8' || s === 'HUP' || s === 'ABS' || s === 'AD');
      const t = isRest ? 'REST' : 'WORK';
      const cur = blocks[blocks.length - 1];
      if (cur && cur.type === t && (t === 'REST' || cur.shift === s)) {
        cur.len++;
      } else {
        blocks.push({ type: t, shift: isRest ? 'OFF' : s, len: 1 });
      }
    }
    return { empHistory, empMonthShifts, activeTailWork, activeTailRest, full, blocks };
  };

  const countShortRests = (emp: string): number => {
    const { blocks } = getBlocks(emp);
    let shortRests = 0;
    for (let i = 1; i < blocks.length - 1; i++) {
      if (blocks[i].type === 'REST' && blocks[i].len < 2) shortRests++;
    }
    return shortRests;
  };

  const isEmpValid = (emp: string) => {
    const { empHistory, empMonthShifts, activeTailWork, activeTailRest, full, blocks } = getBlocks(emp);
    const hasHistory = empHistory.length > 0;
    const lastHistoryWorkShift = getLastHistoryWorkShift(emp);

    // 1. Mixed blocks check
    if (hasMixedBlocks([...activeTailWork, ...empMonthShifts])) return false;

    // 0. Continuous work streak across all working shifts: NEVER > 5
    let curWorkStreak = 0;
    for (const s of full) {
      if (s !== 'OFF' && s !== 'H8' && s !== 'HUP' && s !== 'ABS' && s !== 'AD') {
        curWorkStreak++;
        if (curWorkStreak > 5) return false;
      } else {
        curWorkStreak = 0;
      }
    }

    // 1. Two WORK blocks cannot be directly adjacent without rest!
    for (let i = 0; i < blocks.length - 1; i++) {
      if (blocks[i].type === 'WORK' && blocks[i + 1].type === 'WORK') {
        return false;
      }
    }

    // 2. Shift sequence check: Work blocks must follow sequence A -> M -> N -> A!
    const EXPECTED_NEXT_SHIFT: Record<string, string> = {
      A: 'M',
      M: 'N',
      N: 'A'
    };
    const workBlocks = blocks.filter((b) => b.type === 'WORK');
    if (activeTailWork.length === 0 && lastHistoryWorkShift && workBlocks.length > 0) {
      const expectedFirst = EXPECTED_NEXT_SHIFT[lastHistoryWorkShift];
      if (expectedFirst && workBlocks[0].shift !== expectedFirst) {
        return false;
      }
    }
    for (let i = 0; i < workBlocks.length - 1; i++) {
      const cur = workBlocks[i].shift;
      const nxt = workBlocks[i + 1].shift;
      if (cur === nxt) return false; // Anti-repeat check
      if (EXPECTED_NEXT_SHIFT[cur] && nxt !== EXPECTED_NEXT_SHIFT[cur]) {
        return false; // Sequence violation
      }
    }

    // 4. Work and rest streaks
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.type === 'WORK') {
        if (b.len > 5) return false;
        // Internal work blocks must be between 3 and 5 days.
        // Boundary blocks: last block at month end can be 1..5 (continues into next month),
        // and first block on Day 1 without prior history can be 1..5 (truncated from prior month).
        const isUnboundedStart = i === 0 && !hasHistory;
        if (!isUnboundedStart && i < blocks.length - 1 && b.len < 3) return false;
      } else {
        // Rest period between work blocks (and trailing month-end rest):
        // After Night shift before returning to work: strictly >= 2 days (48h mandatory rest)
        // Max rest: at most 3 days anywhere in month (including trailing OFF days)
        if (i > 0) {
          if (i < blocks.length - 1) {
            const prevWork = blocks[i - 1];
            if (prevWork && prevWork.shift === 'N' && b.len < 2) return false;
          }
          if (b.len > 3) return false;
        }
        // Rest bridging across month boundary and leading rest in standalone months
        if (i === 0 && activeTailRest.length === 0 && b.len > 3) {
          return false;
        }
        if (i === 0 && activeTailRest.length > 0 && i < blocks.length - 1) {
          if (lastHistoryWorkShift === 'N' && b.len < 2) return false;
          if (b.len > activeTailRest.length && b.len > 3) return false;
        }
      }
    }

    // 5. Night turnaround check: night shift followed by at least 2 days of rest before M, A, or MID
    const checkSeq = [...empHistory.slice(-2), ...empMonthShifts];
    const histOffset = Math.min(2, empHistory.length);
    for (let i = 0; i < checkSeq.length; i++) {
      if (checkSeq[i] === 'N') {
        if (i + 1 >= histOffset && i + 1 < checkSeq.length) {
          const next = checkSeq[i + 1];
          if (next === 'M' || next === 'A' || next === 'MID') return false;
        }
        if (i + 2 >= histOffset && i + 2 < checkSeq.length) {
          const next = checkSeq[i + 1];
          const next2 = checkSeq[i + 2];
          if (next === 'OFF' && (next2 === 'M' || next2 === 'A' || next2 === 'MID')) return false;
        }
      }
    }

    return true;
  };

  const getDailyCoverage = () => {
    const dailyCov: Record<string, { rotM: number; totalA: number; N: number }> = {};
    dateColumns.forEach(({ isoDate }) => {
      let rotM = 0;
      let totalA = 0;
      let N = 0;
      rotating.forEach((emp) => {
        const s = rows[emp][isoDate];
        if (s === 'M') rotM++;
        else if (s === 'A') totalA++;
        else if (s === 'N') N++;
      });
      fixedWeekdayMA.forEach((emp) => {
        if (rows[emp]?.[isoDate] === 'A') totalA++;
      });
      dailyCov[isoDate] = { rotM, totalA, N };
    });
    return dailyCov;
  };

  const getDayCounts = (isoDate: string) => {
    let rotM = 0;
    let totalA = 0;
    let N = 0;
    rotating.forEach((emp) => {
      const s = rows[emp][isoDate];
      if (s === 'M') rotM++;
      else if (s === 'A') totalA++;
      else if (s === 'N') N++;
    });
    fixedWeekdayMA.forEach((emp) => {
      if (rows[emp]?.[isoDate] === 'A') totalA++;
    });
    return { rotM, totalA, N };
  };

  const canDropShiftOnDay = (isoDate: string, shift: string): boolean => {
    const c = getDayCounts(isoDate);
    if (shift === 'M') return c.rotM > 1;
    if (shift === 'A') return c.totalA > 2;
    if (shift === 'N') return c.N > 2;
    return false;
  };

  const canAddShiftOnDay = (isoDate: string, shift: string): boolean => {
    if (shift !== 'N') return true;
    return getDayCounts(isoDate).N < 2;
  };

  const getPrevNeighbor = (emp: string, i: number) =>
    i > 0
      ? shiftKey(rows[emp][dateColumns[i - 1].isoDate] ?? 'OFF')
      : (employeeHistories[emp]?.[employeeHistories[emp].length - 1] ?? 'OFF');

  const getNextNeighbor = (emp: string, i: number) =>
    i < dateColumns.length - 1
      ? shiftKey(rows[emp][dateColumns[i + 1].isoDate] ?? 'OFF')
      : 'OFF';

  const getValidAddCandidates = (emp: string, i: number): ('M' | 'A' | 'N')[] => {
    const prev = getPrevNeighbor(emp, i);
    const next = getNextNeighbor(emp, i);
    const cands: ('M' | 'A' | 'N')[] = [];
    if (prev === 'M' || prev === 'A' || prev === 'N') cands.push(prev);
    if ((next === 'M' || next === 'A' || next === 'N') && !cands.includes(next)) cands.push(next);
    if (cands.length === 0) cands.push('M', 'A', 'N');
    return cands.filter((cand) =>
      !(prev !== 'OFF' && prev !== cand) && !(next !== 'OFF' && next !== cand)
    );
  };

  // Dropping strictly inside a block would split one shift block into two identical blocks separated by OFF
  const isBoundaryDrop = (emp: string, i: number, shift: string): boolean =>
    !(getPrevNeighbor(emp, i) === shift && getNextNeighbor(emp, i) === shift);

  // Single added shift must touch an existing block of `shift` or sit at an unbounded month edge
  const canAttachStandaloneAdd = (emp: string, i: number, shift: string): boolean => {
    const prev = getPrevNeighbor(emp, i);
    const next = getNextNeighbor(emp, i);
    if ((prev !== 'OFF' && prev !== shift) || (next !== 'OFF' && next !== shift)) return false;
    if (prev === shift || next === shift) return true;
    const isUnboundedStart = i === 0 && (employeeHistories[emp]?.length ?? 0) === 0;
    return isUnboundedStart || i === dateColumns.length - 1;
  };

  // In a pair (a1 < a2), an isolated a1 can only become valid if a2 = a1 + 1 bridges to an edge or existing block
  const canPlaceFirstOfPair = (emp: string, a1: number, s1: string): boolean => {
    if (canAttachStandaloneAdd(emp, a1, s1)) return true;
    const a2 = a1 + 1;
    if (a2 >= dateColumns.length || rows[emp][dateColumns[a2].isoDate] !== 'OFF') return false;
    const next2 = getNextNeighbor(emp, a2);
    return a2 === dateColumns.length - 1 || next2 === s1;
  };

  // Phase 0: Resolve boundary consecutive work streaks > 5 days by handing off excess shifts
  for (const emp of rotating) {
    const hist = employeeHistories[emp] ?? [];
    let tailStreak = 0;
    for (let i = hist.length - 1; i >= 0; i--) {
      if (hist[i] !== 'OFF' && hist[i] !== 'H8' && hist[i] !== 'HUP' && hist[i] !== 'ABS' && hist[i] !== 'AD') tailStreak++;
      else break;
    }
    if (tailStreak === 0) continue;

    let headStreak = 0;
    for (let i = 0; i < Math.min(5, dateColumns.length); i++) {
      const s = rows[emp][dateColumns[i].isoDate];
      if (s !== 'OFF' && s !== 'H8' && s !== 'HUP' && s !== 'ABS' && s !== 'AD') headStreak++;
      else break;
    }

    if (tailStreak + headStreak > 5) {
      const cutoff = Math.max(0, 5 - tailStreak);
      for (let dayIdx = cutoff; dayIdx < headStreak; dayIdx++) {
        const isoDate = dateColumns[dayIdx].isoDate;
        const curShift = rows[emp][isoDate];
        if (curShift === 'OFF') continue;

        let swapped = false;
        for (const other of rotating) {
          if (other === emp) continue;
          if (rows[other][isoDate] !== 'OFF') continue;

          rows[emp][isoDate] = 'OFF';
          rows[other][isoDate] = curShift;

          if (isEmpValid(emp) && isEmpValid(other)) {
            swapped = true;
            break;
          } else {
            rows[emp][isoDate] = curShift;
            rows[other][isoDate] = 'OFF';
          }
        }
        if (swapped) break;
      }
    }
  }

  // Phase 0b: Repair any rotating employee still failing isEmpValid (e.g. boundary <3d block or <2d N rest)
  for (const emp of rotating) {
    for (let iter = 0; iter < 5 && !isEmpValid(emp); iter++) {
      let fixed = false;
      for (const strictCov of [true, false]) {
        for (let d = 0; d < dateColumns.length; d++) {
          const dropIso = dateColumns[d].isoDate;
          const dropShift = rows[emp][dropIso];
          if (dropShift !== 'M' && dropShift !== 'A' && dropShift !== 'N') continue;
          if (strictCov && !canDropShiftOnDay(dropIso, dropShift)) continue;

          rows[emp][dropIso] = 'OFF';
          for (let a = 0; a < dateColumns.length; a++) {
            if (a === d) continue;
            const addIso = dateColumns[a].isoDate;
            if (rows[emp][addIso] !== 'OFF') continue;
            for (const cand of getValidAddCandidates(emp, a)) {
              if (!canAddShiftOnDay(addIso, cand)) continue;
              rows[emp][addIso] = cand;
              if (isEmpValid(emp)) {
                fixed = true;
                break;
              }
              rows[emp][addIso] = 'OFF';
            }
            if (fixed) break;
          }
          if (fixed) break;
          rows[emp][dropIso] = dropShift;
        }
        if (fixed) break;
      }
      if (!fixed) break;
    }
  }

  // Phase A: Shift transfers from surplus employees (> target) to deficit employees (< target)
  for (let iter = 0; iter < 30; iter++) {
    const surplus = rotating.filter((e) => getWorkDays(e) > targetWorkingDays);
    const deficit = rotating.filter((e) => getWorkDays(e) < targetWorkingDays);
    if (surplus.length === 0 || deficit.length === 0) break;

    let moved = false;
    for (const sEmp of surplus) {
      for (const dEmp of deficit) {
        for (let i = 0; i < dateColumns.length; i++) {
          const isoDate = dateColumns[i].isoDate;
          const sShift = rows[sEmp][isoDate];
          const dShift = rows[dEmp][isoDate];

          if ((sShift === 'M' || sShift === 'A' || sShift === 'N') && dShift === 'OFF') {
            const prevShift = i > 0
              ? shiftKey(rows[dEmp][dateColumns[i - 1].isoDate] ?? 'OFF')
              : (employeeHistories[dEmp]?.[employeeHistories[dEmp].length - 1] ?? 'OFF');
            const nextShift = i < dateColumns.length - 1
              ? shiftKey(rows[dEmp][dateColumns[i + 1].isoDate] ?? 'OFF')
              : 'OFF';

            // Only transfer if dEmp does not touch a different working shift type
            if ((prevShift !== 'OFF' && prevShift !== sShift) || (nextShift !== 'OFF' && nextShift !== sShift)) {
              continue;
            }

            rows[sEmp][isoDate] = 'OFF';
            rows[dEmp][isoDate] = sShift;

            if (isEmpValid(sEmp) && isEmpValid(dEmp)) {
              moved = true;
              break;
            } else {
              rows[sEmp][isoDate] = sShift;
              rows[dEmp][isoDate] = 'OFF';
            }
          }
        }
        if (moved) break;
      }
      if (moved) break;
    }
    if (!moved) break;
  }

  const runPhaseB = () => {
    // Phase B: If any deficit remains, add shifts (preferring coverage deficits and preserving >=2d rests)
    for (let iter = 0; iter < 25; iter++) {
      const deficit = rotating.filter((e) => getWorkDays(e) < targetWorkingDays);
      if (deficit.length === 0) break;

      let added = false;
      for (const dEmp of deficit) {
        let bestMove: { isoDate: string; cand: 'M' | 'A' | 'N'; score: number } | null = null;
        const prevShortRests = countShortRests(dEmp);

        const counts: Record<string, number> = { M: 0, A: 0, N: 0 };
        dateColumns.forEach(({ isoDate: dIso }) => {
          const s = rows[dEmp][dIso];
          if (s === 'M' || s === 'A' || s === 'N') counts[s]++;
        });

        for (let i = 0; i < dateColumns.length; i++) {
          const isoDate = dateColumns[i].isoDate;
          if (rows[dEmp][isoDate] !== 'OFF') continue;

          const prevShift = getPrevNeighbor(dEmp, i);
          const nextShift = getNextNeighbor(dEmp, i);

          const candidates: ('M' | 'A' | 'N')[] = [];
          if (prevShift === 'M' || prevShift === 'A' || prevShift === 'N') candidates.push(prevShift);
          if ((nextShift === 'M' || nextShift === 'A' || nextShift === 'N') && !candidates.includes(nextShift)) {
            candidates.push(nextShift);
          }
          if (candidates.length === 0) {
            candidates.push('M', 'A');
            if (getDayCounts(isoDate).N < 2) candidates.push('N');
          }

          const dayCov = getDayCounts(isoDate);
          for (const cand of candidates) {
            if ((prevShift !== 'OFF' && prevShift !== cand) || (nextShift !== 'OFF' && nextShift !== cand)) {
              continue;
            }
            if (cand === 'N' && dayCov.N >= 2) continue;

            rows[dEmp][isoDate] = cand;
            if (isEmpValid(dEmp)) {
              const newCounts = { ...counts, [cand]: (counts[cand] ?? 0) + 1 };
              const spread = Math.max(newCounts.M, newCounts.A) - Math.min(newCounts.M, newCounts.A);
              const fixesDeficit =
                (cand === 'N' && dayCov.N < 2) ||
                (cand === 'A' && dayCov.totalA < 2) ||
                (cand === 'M' && dayCov.rotM < 1);
              const newShortRests = countShortRests(dEmp);
              const shortRestPenalty = Math.max(0, newShortRests - prevShortRests) * 100;
              const score = (fixesDeficit ? -1000 : 0) + shortRestPenalty + spread;
              if (!bestMove || score < bestMove.score) {
                bestMove = { isoDate, cand, score };
              }
            }
            rows[dEmp][isoDate] = 'OFF';
          }
        }

        if (bestMove) {
          rows[dEmp][bestMove.isoDate] = bestMove.cand;
          added = true;
          break;
        }
      }
      if (!added) break;
    }
  };

  const runPhaseB2 = () => {
    // Phase B2: Multi-step & 2-employee repair for remaining deficit employees (net +1 work day)
    for (let iter = 0; iter < 20; iter++) {
      const deficit = rotating.filter((e) => getWorkDays(e) < targetWorkingDays);
      if (deficit.length === 0) break;

      let repaired = false;
      for (const dEmp of deficit) {
        // B2a: 1-employee (1 drop + 2 adds = net +1) sliding block boundaries while preserving coverage
        const tryOneEmpSlideAdd = (): boolean => {
          let fallbackMove: { dropIso: string; dropShift: string; iso1: string; s1: string; iso2: string; s2: string } | null = null;

          for (let d = 0; d < dateColumns.length; d++) {
            const dropIso = dateColumns[d].isoDate;
            const dropShift = rows[dEmp][dropIso];
            if (dropShift !== 'M' && dropShift !== 'A' && dropShift !== 'N') continue;
            if (!isBoundaryDrop(dEmp, d, dropShift) || !canDropShiftOnDay(dropIso, dropShift)) continue;

            rows[dEmp][dropIso] = 'OFF';
            for (let a1 = 0; a1 < dateColumns.length; a1++) {
              if (a1 === d) continue;
              const iso1 = dateColumns[a1].isoDate;
              if (rows[dEmp][iso1] !== 'OFF') continue;

              const cands1 = getValidAddCandidates(dEmp, a1);
              for (const s1 of cands1) {
                if (!canAddShiftOnDay(iso1, s1) || !canPlaceFirstOfPair(dEmp, a1, s1)) continue;
                rows[dEmp][iso1] = s1;

                for (let a2 = a1 + 1; a2 < dateColumns.length; a2++) {
                  if (a2 === d) continue;
                  const iso2 = dateColumns[a2].isoDate;
                  if (rows[dEmp][iso2] !== 'OFF') continue;

                  const cands2 = getValidAddCandidates(dEmp, a2);
                  for (const s2 of cands2) {
                    if (!canAddShiftOnDay(iso2, s2) || !canAttachStandaloneAdd(dEmp, a2, s2)) continue;
                    rows[dEmp][iso2] = s2;

                    if (isEmpValid(dEmp)) {
                      if (countShortRests(dEmp) === 0) {
                        return true;
                      }
                      if (!fallbackMove) {
                        fallbackMove = { dropIso, dropShift, iso1, s1, iso2, s2 };
                      }
                    }
                    rows[dEmp][iso2] = 'OFF';
                  }
                }
                rows[dEmp][iso1] = 'OFF';
              }
            }
            rows[dEmp][dropIso] = dropShift;
          }

          if (fallbackMove) {
            rows[dEmp][fallbackMove.dropIso] = 'OFF';
            rows[dEmp][fallbackMove.iso1] = fallbackMove.s1;
            rows[dEmp][fallbackMove.iso2] = fallbackMove.s2;
            return true;
          }
          return false;
        };

        if (tryOneEmpSlideAdd()) {
          repaired = true;
          break;
        }

        // B2b-1: 2-employee transfer + recovery (helper transfers iso1 to dEmp, helper adds iso2)
        for (const helper of rotating) {
          if (helper === dEmp) continue;
          for (let i = 0; i < dateColumns.length; i++) {
            const iso1 = dateColumns[i].isoDate;
            const hShift = rows[helper][iso1];
            if ((hShift !== 'M' && hShift !== 'A' && hShift !== 'N') || rows[dEmp][iso1] !== 'OFF') {
              continue;
            }
            const prevD = getPrevNeighbor(dEmp, i);
            const nextD = getNextNeighbor(dEmp, i);
            if ((prevD !== 'OFF' && prevD !== hShift) || (nextD !== 'OFF' && nextD !== hShift)) {
              continue;
            }
            if (!isBoundaryDrop(helper, i, hShift)) continue;

            rows[helper][iso1] = 'OFF';
            rows[dEmp][iso1] = hShift;
            if (isEmpValid(dEmp)) {
              for (let j = 0; j < dateColumns.length; j++) {
                if (j === i) continue;
                const iso2 = dateColumns[j].isoDate;
                if (rows[helper][iso2] !== 'OFF') continue;
                for (const s2 of getValidAddCandidates(helper, j)) {
                  if (!canAddShiftOnDay(iso2, s2) || !canAttachStandaloneAdd(helper, j, s2)) continue;
                  rows[helper][iso2] = s2;
                  if (isEmpValid(helper)) {
                    repaired = true;
                    break;
                  }
                  rows[helper][iso2] = 'OFF';
                }
                if (repaired) break;
              }
            }
            if (repaired) break;
            rows[helper][iso1] = hShift;
            rows[dEmp][iso1] = 'OFF';
          }
          if (repaired) break;
        }
        if (repaired) break;

        // B2b-2: dEmp drops iso1 (at min cov) + adds iso2 & iso3, helper covers iso1 + drops surplus iso4
        for (let d = 0; d < dateColumns.length; d++) {
          const iso1 = dateColumns[d].isoDate;
          const dropShift = rows[dEmp][iso1];
          if (dropShift !== 'M' && dropShift !== 'A' && dropShift !== 'N') continue;
          if (!isBoundaryDrop(dEmp, d, dropShift)) continue;
          const validHelpers = rotating.filter(
            (helper) => helper !== dEmp && rows[helper][iso1] === 'OFF' && canAttachStandaloneAdd(helper, d, dropShift)
          );
          if (validHelpers.length === 0) continue;

          rows[dEmp][iso1] = 'OFF';
          for (let a1 = 0; a1 < dateColumns.length; a1++) {
            if (a1 === d) continue;
            const iso2 = dateColumns[a1].isoDate;
            if (rows[dEmp][iso2] !== 'OFF') continue;
            for (const s1 of getValidAddCandidates(dEmp, a1)) {
              if (!canAddShiftOnDay(iso2, s1) || !canPlaceFirstOfPair(dEmp, a1, s1)) continue;
              rows[dEmp][iso2] = s1;
              for (let a2 = a1 + 1; a2 < dateColumns.length; a2++) {
                if (a2 === d) continue;
                const iso3 = dateColumns[a2].isoDate;
                if (rows[dEmp][iso3] !== 'OFF') continue;
                for (const s2 of getValidAddCandidates(dEmp, a2)) {
                  if (!canAddShiftOnDay(iso3, s2) || !canAttachStandaloneAdd(dEmp, a2, s2)) continue;
                  rows[dEmp][iso3] = s2;
                  if (isEmpValid(dEmp)) {
                    for (const helper of validHelpers) {
                      rows[helper][iso1] = dropShift;
                      for (let k = 0; k < dateColumns.length; k++) {
                        if (k === d) continue;
                        const iso4 = dateColumns[k].isoDate;
                        const s4 = rows[helper][iso4];
                        if (s4 !== 'M' && s4 !== 'A' && s4 !== 'N') continue;
                        if (!isBoundaryDrop(helper, k, s4) || !canDropShiftOnDay(iso4, s4)) continue;
                        rows[helper][iso4] = 'OFF';
                        if (isEmpValid(helper)) {
                          repaired = true;
                          break;
                        }
                        rows[helper][iso4] = s4;
                      }
                      if (repaired) break;
                      rows[helper][iso1] = 'OFF';
                    }
                  }
                  if (repaired) break;
                  rows[dEmp][iso3] = 'OFF';
                }
                if (repaired) break;
              }
              if (repaired) break;
              rows[dEmp][iso2] = 'OFF';
            }
            if (repaired) break;
          }
          if (repaired) break;
          rows[dEmp][iso1] = dropShift;
        }
        if (repaired) break;
      }
      if (!repaired) break;
    }
  };

  runPhaseB();
  runPhaseB2();

  // Phase C: If any surplus remains, drop shift to OFF on days with adequate coverage (including fixedWeekdayMA for A)
  for (let iter = 0; iter < 25; iter++) {
    const surplus = rotating.filter((e) => getWorkDays(e) > targetWorkingDays);
    if (surplus.length === 0) break;

    let dropped = false;
    for (const sEmp of surplus) {
      let bestDropIso: string | null = null;
      let bestDropScore = Infinity;
      const prevShortRests = countShortRests(sEmp);

      for (let i = 0; i < dateColumns.length; i++) {
        const isoDate = dateColumns[i].isoDate;
        const cur = rows[sEmp][isoDate];
        if (cur !== 'M' && cur !== 'A' && cur !== 'N') continue;

        const c = getDayCounts(isoDate);
        const minAllowed = cur === 'M' ? 1 : 2;
        const count = cur === 'M' ? c.rotM : cur === 'A' ? c.totalA : c.N;
        if (count > minAllowed) {
          rows[sEmp][isoDate] = 'OFF';
          if (isEmpValid(sEmp)) {
            const shortRestDelta = Math.max(0, countShortRests(sEmp) - prevShortRests);
            const surplusMargin = count - minAllowed;
            const score = shortRestDelta * 100 - surplusMargin;
            if (score < bestDropScore) {
              bestDropScore = score;
              bestDropIso = isoDate;
            }
          }
          rows[sEmp][isoDate] = cur;
        }
      }

      if (bestDropIso) {
        rows[sEmp][bestDropIso] = 'OFF';
        dropped = true;
        break;
      }
    }
    if (!dropped) break;
  }

  // Phase C2: Multi-step & 2-employee repair for remaining surplus employees (net -1 work day)
  for (let iter = 0; iter < 20; iter++) {
    const surplus = rotating.filter((e) => getWorkDays(e) > targetWorkingDays);
    if (surplus.length === 0) break;

    let repaired = false;
    for (const sEmp of surplus) {
      // C2a: 1-employee (1 add + 2 drops = net -1)
      for (let a = 0; a < dateColumns.length; a++) {
        const isoA = dateColumns[a].isoDate;
        if (rows[sEmp][isoA] !== 'OFF') continue;
        for (const sA of getValidAddCandidates(sEmp, a)) {
          if (!canAddShiftOnDay(isoA, sA) || !canAttachStandaloneAdd(sEmp, a, sA)) continue;
          rows[sEmp][isoA] = sA;

          for (let d1 = 0; d1 < dateColumns.length; d1++) {
            if (d1 === a) continue;
            const iso1 = dateColumns[d1].isoDate;
            const sh1 = rows[sEmp][iso1];
            if (sh1 !== 'M' && sh1 !== 'A' && sh1 !== 'N') continue;
            if (!isBoundaryDrop(sEmp, d1, sh1) || !canDropShiftOnDay(iso1, sh1)) continue;
            rows[sEmp][iso1] = 'OFF';

            for (let d2 = d1 + 1; d2 < dateColumns.length; d2++) {
              if (d2 === a) continue;
              const iso2 = dateColumns[d2].isoDate;
              const sh2 = rows[sEmp][iso2];
              if (sh2 !== 'M' && sh2 !== 'A' && sh2 !== 'N') continue;
              if (!isBoundaryDrop(sEmp, d2, sh2) || !canDropShiftOnDay(iso2, sh2)) continue;
              rows[sEmp][iso2] = 'OFF';

              if (isEmpValid(sEmp)) {
                repaired = true;
                break;
              }
              rows[sEmp][iso2] = sh2;
            }
            if (repaired) break;
            rows[sEmp][iso1] = sh1;
          }
          if (repaired) break;
          rows[sEmp][isoA] = 'OFF';
        }
        if (repaired) break;
      }
      if (repaired) break;

      // C2b: 2-employee handoff + drop (sEmp hands off iso1 to helper, helper drops surplus iso2)
      for (let i = 0; i < dateColumns.length; i++) {
        const iso1 = dateColumns[i].isoDate;
        const sh1 = rows[sEmp][iso1];
        if (sh1 !== 'M' && sh1 !== 'A' && sh1 !== 'N') continue;
        if (!isBoundaryDrop(sEmp, i, sh1)) continue;

        rows[sEmp][iso1] = 'OFF';
        if (isEmpValid(sEmp)) {
          for (const helper of rotating) {
            if (helper === sEmp || rows[helper][iso1] !== 'OFF') continue;
            if (!canAttachStandaloneAdd(helper, i, sh1)) continue;
            rows[helper][iso1] = sh1;
            for (let j = 0; j < dateColumns.length; j++) {
              if (j === i) continue;
              const iso2 = dateColumns[j].isoDate;
              const sh2 = rows[helper][iso2];
              if (sh2 !== 'M' && sh2 !== 'A' && sh2 !== 'N') continue;
              if (!isBoundaryDrop(helper, j, sh2) || !canDropShiftOnDay(iso2, sh2)) continue;
              rows[helper][iso2] = 'OFF';
              if (isEmpValid(helper)) {
                repaired = true;
                break;
              }
              rows[helper][iso2] = sh2;
            }
            if (repaired) break;
            rows[helper][iso1] = 'OFF';
          }
        }
        if (repaired) break;
        rows[sEmp][iso1] = sh1;
      }
      if (repaired) break;
    }
    if (!repaired) break;
  }

  // Phase NightCap: Strict hard ceiling of maximum 2 people on Night shift (N = 2) on all days
  let nightCapDropped = false;
  for (let iter = 0; iter < 25; iter++) {
    const overDays = dateColumns.filter(({ isoDate }) => getDayCounts(isoDate).N > 2);
    if (overDays.length === 0) break;

    let reduced = false;
    for (const { index, isoDate } of overDays) {
      const empsOnN = rotating
        .filter((e) => rows[e][isoDate] === 'N')
        .sort((a, b) => getWorkDays(b) - getWorkDays(a));
      for (const emp of empsOnN) {
        rows[emp][isoDate] = 'OFF';
        if (isEmpValid(emp)) {
          reduced = true;
          nightCapDropped = true;
          break;
        }
        rows[emp][isoDate] = 'N';
      }
      if (reduced) break;

      // Fallback: slide N from isoDate to an adjacent day where N < 2
      for (const emp of empsOnN) {
        rows[emp][isoDate] = 'OFF';
        for (const adjIdx of [index - 1, index + 1]) {
          if (adjIdx < 0 || adjIdx >= dateColumns.length) continue;
          const adjIso = dateColumns[adjIdx].isoDate;
          if (rows[emp][adjIso] !== 'OFF' || getDayCounts(adjIso).N >= 2) continue;
          rows[emp][adjIso] = 'N';
          if (isEmpValid(emp)) {
            reduced = true;
            break;
          }
          rows[emp][adjIso] = 'OFF';
        }
        if (reduced) break;
        rows[emp][isoDate] = 'N';
      }
      if (reduced) break;
    }
    if (!reduced) break;
  }

  // Phase NightCap Compensation: Restore any employee pushed below targetWorkingDays
  if (nightCapDropped) {
    runPhaseB();
    runPhaseB2();
  }

  // Phase D: Coverage Equalization - Resolves understaffed days (N < 2, totalA < 2, rotM < 1) by shifting shifts within valid work blocks
  for (let iter = 0; iter < 40; iter++) {
    const cov = getDailyCoverage();
    let moved = false;

    for (let i = 0; i < dateColumns.length; i++) {
      const isoDate = dateColumns[i].isoDate;
      const c = cov[isoDate];

      const deficits: ('N' | 'A' | 'M')[] = [];
      if (c.N < 2) deficits.push('N');
      if (c.totalA < 2) deficits.push('A');
      if (c.rotM < 1) deficits.push('M');

      for (const targetShift of deficits) {
        for (const emp of rotating) {
          if (rows[emp][isoDate] !== 'OFF') continue;

          // Target shift must touch an existing block of targetShift or sit at an unbounded month edge
          if (!canAttachStandaloneAdd(emp, i, targetShift)) continue;

          rows[emp][isoDate] = targetShift;

          let bestDropIso: string | null = null;
          let bestShortRests = Infinity;
          for (let j = 0; j < dateColumns.length; j++) {
            if (i === j) continue;
            const dropIso = dateColumns[j].isoDate;
            const dropShift = rows[emp][dropIso];
            if (dropShift !== 'M' && dropShift !== 'A' && dropShift !== 'N') continue;

            const dropCov = cov[dropIso];
            const hasSurplus =
              (dropShift === 'M' && dropCov.rotM > 1) ||
              (dropShift === 'A' && dropCov.totalA > 2) ||
              (dropShift === 'N' && dropCov.N > 2);
            if (!hasSurplus) continue;

            rows[emp][dropIso] = 'OFF';
            if (isEmpValid(emp)) {
              const sr = countShortRests(emp);
              if (sr < bestShortRests) {
                bestShortRests = sr;
                bestDropIso = dropIso;
              }
            }
            rows[emp][dropIso] = dropShift;
          }

          if (bestDropIso) {
            rows[emp][bestDropIso] = 'OFF';
            moved = true;
            break;
          }
          rows[emp][isoDate] = 'OFF';
        }
        if (moved) break;
      }
      if (moved) break;
    }
    if (!moved) break;
  }

  // Phase D2: Multi-step & 2-employee coverage repair for remaining understaffed days
  for (let iter = 0; iter < 30; iter++) {
    const cov = getDailyCoverage();
    let repaired = false;

    for (let i = 0; i < dateColumns.length; i++) {
      const isoDate = dateColumns[i].isoDate;
      const c = cov[isoDate];

      const deficits: ('N' | 'A' | 'M')[] = [];
      if (c.N < 2) deficits.push('N');
      if (c.totalA < 2) deficits.push('A');
      if (c.rotM < 1) deficits.push('M');
      if (deficits.length === 0) continue;

      for (const targetShift of deficits) {
        // D2a: 1-employee block-slide (emp adds targetShift on isoDate, drops blocking d1 with surplus, slides d1->a2, drops surplus d2)
        for (const emp of rotating) {
          if (rows[emp][isoDate] !== 'OFF') continue;
          if (!canAttachStandaloneAdd(emp, i, targetShift)) continue;

          rows[emp][isoDate] = targetShift;
          const localStart = Math.max(0, i - 5);
          const localEnd = Math.min(dateColumns.length - 1, i + 5);

          for (let d1 = localStart; d1 <= localEnd; d1++) {
            if (d1 === i) continue;
            const iso1 = dateColumns[d1].isoDate;
            const sh1 = rows[emp][iso1];
            if (sh1 !== 'M' && sh1 !== 'A' && sh1 !== 'N') continue;
            if (!isBoundaryDrop(emp, d1, sh1) || !canDropShiftOnDay(iso1, sh1)) continue;

            rows[emp][iso1] = 'OFF';
            for (let a2 = Math.max(0, d1 - 6); a2 <= Math.min(dateColumns.length - 1, d1 + 6); a2++) {
              if (a2 === i || a2 === d1) continue;
              const iso2 = dateColumns[a2].isoDate;
              if (rows[emp][iso2] !== 'OFF') continue;
              if (!canAddShiftOnDay(iso2, sh1) || !canAttachStandaloneAdd(emp, a2, sh1)) continue;

              rows[emp][iso2] = sh1;
              for (let d2 = 0; d2 < dateColumns.length; d2++) {
                if (d2 === i || d2 === d1 || d2 === a2) continue;
                const iso3 = dateColumns[d2].isoDate;
                const sh3 = rows[emp][iso3];
                if (sh3 !== 'M' && sh3 !== 'A' && sh3 !== 'N') continue;
                if (!isBoundaryDrop(emp, d2, sh3) || !canDropShiftOnDay(iso3, sh3)) continue;

                rows[emp][iso3] = 'OFF';
                if (isEmpValid(emp)) {
                  repaired = true;
                  break;
                }
                rows[emp][iso3] = sh3;
              }
              if (repaired) break;
              rows[emp][iso2] = 'OFF';
            }
            if (repaired) break;
            rows[emp][iso1] = sh1;
          }
          if (repaired) break;
          rows[emp][isoDate] = 'OFF';
        }
        if (repaired) break;

        // D2b: 2-employee relay swap
        // emp1 adds targetShift on isoDate and drops midIso; emp2 adds midShift on midIso and drops surplusIso
        for (const emp1 of rotating) {
          if (rows[emp1][isoDate] !== 'OFF') continue;
          if (!canAttachStandaloneAdd(emp1, i, targetShift)) continue;

          rows[emp1][isoDate] = targetShift;
          for (let j = 0; j < dateColumns.length; j++) {
            if (j === i) continue;
            const midIso = dateColumns[j].isoDate;
            const midShift = rows[emp1][midIso];
            if (midShift !== 'M' && midShift !== 'A' && midShift !== 'N') continue;
            if (!isBoundaryDrop(emp1, j, midShift)) continue;

            rows[emp1][midIso] = 'OFF';
            if (isEmpValid(emp1)) {
              for (const emp2 of rotating) {
                if (emp2 === emp1 || rows[emp2][midIso] !== 'OFF') continue;
                if (!canAttachStandaloneAdd(emp2, j, midShift)) continue;
                rows[emp2][midIso] = midShift;
                for (let k = 0; k < dateColumns.length; k++) {
                  if (k === i || k === j) continue;
                  const surplusIso = dateColumns[k].isoDate;
                  const surShift = rows[emp2][surplusIso];
                  if (surShift !== 'M' && surShift !== 'A' && surShift !== 'N') continue;
                  if (!isBoundaryDrop(emp2, k, surShift) || !canDropShiftOnDay(surplusIso, surShift)) continue;

                  rows[emp2][surplusIso] = 'OFF';
                  if (isEmpValid(emp2)) {
                    repaired = true;
                    break;
                  }
                  rows[emp2][surplusIso] = surShift;
                }
                if (repaired) break;
                rows[emp2][midIso] = 'OFF';
              }
            }
            if (repaired) break;
            rows[emp1][midIso] = midShift;
          }
          if (repaired) break;
          rows[emp1][isoDate] = 'OFF';
        }
        if (repaired) break;

        // D2c: 2-employee relay where emp1 also slides an adjacent block (adds targetShift on isoDate, drops midIso, slides d1->a2; emp2 covers midIso and drops surplusIso)
        for (const emp1 of rotating) {
          if (rows[emp1][isoDate] !== 'OFF') continue;
          if (!canAttachStandaloneAdd(emp1, i, targetShift)) continue;

          rows[emp1][isoDate] = targetShift;
          const localStart = Math.max(0, i - 5);
          const localEnd = Math.min(dateColumns.length - 1, i + 5);

          for (let j = localStart; j <= localEnd; j++) {
            if (j === i) continue;
            const midIso = dateColumns[j].isoDate;
            const midShift = rows[emp1][midIso];
            if (midShift !== 'M' && midShift !== 'A' && midShift !== 'N') continue;
            if (!isBoundaryDrop(emp1, j, midShift)) continue;
            const validEmp2s = rotating.filter(
              (emp2) => emp2 !== emp1 && rows[emp2][midIso] === 'OFF' && canAttachStandaloneAdd(emp2, j, midShift)
            );
            if (validEmp2s.length === 0) continue;

            rows[emp1][midIso] = 'OFF';
            for (let a2 = Math.max(0, j - 6); a2 <= Math.min(dateColumns.length - 1, j + 6); a2++) {
              if (a2 === i || a2 === j) continue;
              const iso2 = dateColumns[a2].isoDate;
              if (rows[emp1][iso2] !== 'OFF') continue;
              if (!canAddShiftOnDay(iso2, midShift) || !canAttachStandaloneAdd(emp1, a2, midShift)) continue;

              rows[emp1][iso2] = midShift;
              for (let d2 = 0; d2 < dateColumns.length; d2++) {
                if (d2 === i || d2 === j || d2 === a2) continue;
                const dropIso2 = dateColumns[d2].isoDate;
                const dropSh2 = rows[emp1][dropIso2];
                if (dropSh2 !== 'M' && dropSh2 !== 'A' && dropSh2 !== 'N') continue;
                if (!isBoundaryDrop(emp1, d2, dropSh2) || !canDropShiftOnDay(dropIso2, dropSh2)) continue;

                rows[emp1][dropIso2] = 'OFF';
                if (isEmpValid(emp1)) {
                  for (const emp2 of validEmp2s) {
                    rows[emp2][midIso] = midShift;
                    for (let k = 0; k < dateColumns.length; k++) {
                      if (k === i || k === j || k === a2 || k === d2) continue;
                      const surplusIso = dateColumns[k].isoDate;
                      const surShift = rows[emp2][surplusIso];
                      if (surShift !== 'M' && surShift !== 'A' && surShift !== 'N') continue;
                      if (!isBoundaryDrop(emp2, k, surShift) || !canDropShiftOnDay(surplusIso, surShift)) continue;

                      rows[emp2][surplusIso] = 'OFF';
                      if (isEmpValid(emp2)) {
                        repaired = true;
                        break;
                      }
                      rows[emp2][surplusIso] = surShift;
                    }
                    if (repaired) break;
                    rows[emp2][midIso] = 'OFF';
                  }
                }
                if (repaired) break;
                rows[emp1][dropIso2] = dropSh2;
              }
              if (repaired) break;
              rows[emp1][iso2] = 'OFF';
            }
            if (repaired) break;
            rows[emp1][midIso] = midShift;
          }
          if (repaired) break;
          rows[emp1][isoDate] = 'OFF';
        }
        if (repaired) break;
      }
      if (repaired) break;
    }
    if (!repaired) break;
  }
}

/**
 * Core Schedule Generation Function
 */
export function generateBalancedSchedule(
  roster: RosterData | null,
  targetYear: number,
  targetMonth: number,
  options: ScheduleBuilderOptions = {}
): GeneratedScheduleResult {
  const startTimeMs = performance.now();
  const timeBudgetMs = options.timeBudgetMs ?? 8000;
  const minTier: ComplianceTier = options.minTier ?? 0;
  const maxTier: ComplianceTier = options.maxTier ?? 3;
  const activeTiers: ComplianceTier[] = ([0, 1, 2, 3] as ComplianceTier[]).filter(
    (t) => t >= minTier && t <= maxTier
  );
  if (activeTiers.length === 0) activeTiers.push(maxTier);
  const safetyMarginMs = Math.min(150, Math.floor(timeBudgetMs * 0.1));
  const totalDeadlineMs = startTimeMs + Math.max(50, timeBudgetMs - safetyMarginMs);

  const seed = options.seed ?? Math.floor(Math.random() * 1000000);
  const targetWorkingDays =
    options.targetWorkingDays ?? countLegalWorkingDaysInMonth(targetMonth, targetYear);
  const dateColumns = buildMonthDateColumns(targetYear, targetMonth);
  const daysCount = dateColumns.length;
  const targetStartIso = `${targetYear}-${String(targetMonth + 1).padStart(2, '0')}-01`;

  // Employees list
  const allEmployees: string[] = (roster?.employees && roster.employees.length > 0)
    ? [...roster.employees]
    : [
        'Palade Alexandru-Ionut',
        'Tir George Cristian',
        'Urzica Ion-Alexandru',
        'Debona Andrei-Laurentiu',
        'Nica Bogdan-Alexandru',
        'Corneci Stefan-Cristian',
        'Toma Mihai',
        'Gheorghe Adrian Marinescu',
        'Olteanu Ionut-Lucian',
        'Stoian Alexandru-Gabriel',
        'Prisecaru Alexandru-Ionut'
      ];

  const excludedSet = new Set(options.excludedEmployees ?? []);
  const activeAllEmployees = allEmployees.filter((emp) => !excludedSet.has(emp));

  const fixedMidEmployees = FIXED_MID_EMPLOYEES.filter((name) => activeAllEmployees.includes(name));
  const fixedWeekdayMAEmployees = FIXED_WEEKDAY_MA_EMPLOYEES.filter((name) => activeAllEmployees.includes(name));
  const rotatingEmployees = activeAllEmployees.filter((name) => !fixedMidEmployees.includes(name) && !fixedWeekdayMAEmployees.includes(name));

  // Balanced 1:1:1 templates with 4-day blocks and 2-day rests across 18-day coprime cycles
  // Strict cyclic direction: A -> M -> N -> A (Afternoon -> Morning -> Night -> Afternoon)
  // Standard density (12/18 = 66.7% work -> ~20-21 work days in 30-31d)
  const STANDARD_TEMPLATES = [
    // Template 0: A -> M -> N -> A
    [
      'A','A','A','A',
      'OFF','OFF',
      'M','M','M','M',
      'OFF','OFF',
      'N','N','N','N',
      'OFF','OFF'
    ],
    // Template 1: M -> N -> A -> M
    [
      'M','M','M','M',
      'OFF','OFF',
      'N','N','N','N',
      'OFF','OFF',
      'A','A','A','A',
      'OFF','OFF'
    ],
    // Template 2: N -> A -> M -> N
    [
      'N','N','N','N',
      'OFF','OFF',
      'A','A','A','A',
      'OFF','OFF',
      'M','M','M','M',
      'OFF','OFF'
    ]
  ];

  // Lower-density 11/18 (61.1% work -> ~18-19 work days in 30-31d) templates preserving A/N=2 and M>=1 coverage
  const LOW_DENSITY_TEMPLATES = [
    [
      'A','A','A','A',
      'OFF','OFF',
      'M','M','M',
      'OFF','OFF','OFF',
      'N','N','N','N',
      'OFF','OFF'
    ],
    [
      'M','M','M',
      'OFF','OFF','OFF',
      'N','N','N','N',
      'OFF','OFF',
      'A','A','A','A',
      'OFF','OFF'
    ],
    [
      'N','N','N','N',
      'OFF','OFF',
      'A','A','A','A',
      'OFF','OFF',
      'M','M','M',
      'OFF','OFF','OFF'
    ]
  ];

  // Larger rotating teams (>= 10) need 3-day N blocks per 18d cycle to avoid 3/2 Night overstaffing
  const LARGE_TEAM_3N_TEMPLATES = [
    [
      'A','A','A','A',
      'OFF','OFF',
      'M','M','M','M',
      'OFF','OFF',
      'N','N','N',
      'OFF','OFF','OFF'
    ],
    [
      'M','M','M','M',
      'OFF','OFF',
      'N','N','N',
      'OFF','OFF','OFF',
      'A','A','A','A',
      'OFF','OFF'
    ],
    [
      'N','N','N',
      'OFF','OFF','OFF',
      'A','A','A','A',
      'OFF','OFF',
      'M','M','M','M',
      'OFF','OFF'
    ]
  ];

  const isLowWorkRatio = (targetWorkingDays / daysCount) < 0.644;
  const BALANCED_TEMPLATES = rotatingEmployees.length >= 10
    ? [...LARGE_TEAM_3N_TEMPLATES, ...(isLowWorkRatio ? LOW_DENSITY_TEMPLATES : []), ...STANDARD_TEMPLATES]
    : isLowWorkRatio
      ? [...LOW_DENSITY_TEMPLATES, ...STANDARD_TEMPLATES]
      : STANDARD_TEMPLATES;

  // Pre-configured slot sets ensuring optimal coverage across variable team sizes
  const VALID_SLOT_SETS_6 = [
    [0, 1, 4, 6, 10, 14],
    [0, 1, 4, 7, 10, 14],
    [0, 1, 4, 7, 11, 14],
    [0, 1, 4, 7, 11, 15],
    [0, 1, 4, 8, 10, 14],
    [0, 1, 4, 8, 11, 14],
    [0, 1, 4, 8, 11, 15],
    [0, 1, 4, 8, 12, 14],
    [0, 1, 4, 8, 12, 15],
    [0, 1, 5, 6, 10, 14]
  ];

  const VALID_SLOT_SETS_7 = [
    [0, 1, 4, 5, 8, 10, 14],
    [0, 1, 4, 5, 8, 11, 14],
    [0, 1, 4, 5, 8, 11, 15],
    [0, 1, 4, 5, 8, 12, 14],
    [0, 1, 4, 5, 8, 12, 15],
    [0, 1, 4, 5, 9, 10, 14],
    [0, 1, 4, 5, 9, 11, 14],
    [0, 1, 4, 5, 9, 11, 15],
    [0, 1, 4, 5, 9, 12, 14],
    [0, 1, 4, 5, 9, 12, 15]
  ];

  const VALID_SLOT_SETS_8 = [
    [0, 1, 4, 5, 8, 9, 12, 14],
    [0, 1, 4, 5, 8, 9, 12, 15],
    [0, 1, 4, 5, 8, 9, 13, 14],
    [0, 1, 4, 5, 8, 9, 13, 15],
    [0, 1, 4, 5, 8, 10, 12, 14],
    [0, 1, 4, 5, 8, 10, 12, 15],
    [0, 1, 4, 5, 8, 10, 13, 14],
    [0, 1, 4, 5, 8, 10, 13, 15],
    [0, 1, 4, 5, 8, 10, 14, 15],
    [0, 1, 4, 5, 8, 11, 12, 15]
  ];

  const BASE_SLOT_SETS_9 = [
    [0, 2, 4, 6, 8, 10, 12, 14, 16],
    [1, 3, 5, 7, 9, 11, 13, 15, 17],
    [0, 2, 4, 5, 8, 9, 12, 14, 16],
    [0, 2, 4, 5, 8, 10, 12, 14, 16],
    [1, 2, 4, 5, 8, 10, 13, 14, 16],
    [1, 2, 4, 5, 8, 10, 13, 15, 16],
    [1, 2, 4, 5, 9, 10, 13, 14, 16],
    [1, 2, 4, 5, 9, 10, 14, 15, 16],
    [1, 2, 4, 5, 8, 9, 13, 14, 16],
    [1, 2, 4, 5, 8, 9, 14, 15, 16],
    [0, 1, 3, 4, 7, 9, 12, 13, 15],
    [0, 1, 3, 4, 8, 9, 12, 13, 15]
  ];

  const VALID_SLOT_SETS_9: number[][] = [];
  for (const s of BASE_SLOT_SETS_9) {
    for (let shift = 0; shift < 18; shift += 1) {
      const rotated = s.map((x) => (x + shift) % 18).sort((a, b) => a - b);
      const key = rotated.join(',');
      if (!VALID_SLOT_SETS_9.some((existing) => existing.join(',') === key)) {
        VALID_SLOT_SETS_9.push(rotated);
      }
    }
  }

  const VALID_SLOT_SETS_10 = [
    [0, 1, 2, 4, 6, 7, 10, 11, 14, 15],
    [0, 1, 2, 4, 6, 8, 10, 11, 14, 15],
    [0, 1, 2, 4, 6, 8, 10, 12, 14, 15],
    [0, 1, 2, 4, 6, 8, 10, 12, 14, 16],
    [0, 1, 2, 5, 6, 7, 10, 11, 14, 15],
    [0, 1, 2, 5, 6, 8, 10, 11, 14, 15],
    [0, 1, 2, 5, 6, 8, 10, 12, 14, 15],
    [0, 1, 2, 5, 6, 8, 10, 12, 14, 16],
    [0, 1, 2, 5, 6, 9, 10, 11, 14, 15],
    [0, 1, 2, 5, 6, 9, 10, 12, 14, 15]
  ];

  const VALID_SLOT_SETS_11 = [
    [0, 1, 2, 3, 4, 6, 7, 10, 11, 14, 15],
    [0, 1, 2, 4, 6, 8, 10, 11, 13, 14, 17],
    [0, 1, 3, 4, 6, 7, 10, 11, 12, 14, 15],
    [0, 1, 3, 5, 6, 9, 10, 11, 14, 15, 16],
    [0, 1, 4, 5, 6, 9, 10, 11, 13, 15, 17],
    [0, 2, 3, 4, 6, 7, 9, 10, 13, 14, 17],
    [0, 2, 3, 5, 7, 8, 10, 12, 13, 14, 16],
    [0, 2, 4, 5, 6, 8, 10, 11, 14, 15, 17],
    [0, 2, 4, 6, 7, 9, 10, 11, 12, 15, 16],
    [0, 3, 4, 5, 8, 9, 10, 12, 14, 15, 17],
    [1, 2, 3, 4, 6, 8, 10, 12, 14, 15, 17],
    [1, 2, 3, 6, 7, 10, 11, 14, 15, 16, 17],
    [1, 2, 4, 6, 7, 10, 11, 13, 14, 15, 16],
    [1, 2, 5, 6, 9, 10, 13, 14, 15, 16, 17],
    [1, 3, 4, 7, 8, 11, 12, 13, 15, 16, 17],
    [2, 3, 4, 6, 7, 8, 11, 12, 13, 16, 17]
  ];

  const VALID_SLOT_SETS_12 = [
    [0, 1, 2, 4, 5, 6, 8, 9, 10, 12, 14, 15],
    [0, 1, 2, 4, 5, 6, 8, 9, 10, 12, 14, 16],
    [0, 1, 2, 4, 5, 6, 8, 9, 10, 13, 14, 15],
    [0, 1, 2, 4, 5, 6, 8, 9, 10, 13, 14, 16],
    [0, 1, 2, 4, 5, 6, 8, 9, 11, 12, 14, 15],
    [0, 1, 2, 4, 5, 6, 8, 9, 11, 12, 14, 16],
    [0, 1, 2, 4, 5, 6, 8, 9, 11, 12, 15, 16],
    [0, 1, 2, 4, 5, 6, 8, 9, 11, 13, 14, 15],
    [0, 1, 2, 4, 5, 6, 8, 9, 11, 13, 14, 16],
    [0, 1, 2, 4, 5, 6, 8, 9, 11, 13, 15, 16]
  ];

  function getValidSlotSets(count: number): number[][] {
    if (count <= 6) return VALID_SLOT_SETS_6;
    if (count === 7) return VALID_SLOT_SETS_7;
    if (count === 8) return VALID_SLOT_SETS_8;
    if (count === 9) return VALID_SLOT_SETS_9;
    if (count === 10) return VALID_SLOT_SETS_10;
    if (count === 11) return VALID_SLOT_SETS_11;
    if (count === 12) return VALID_SLOT_SETS_12;
    const step = 18 / count;
    return [
      Array.from({ length: count }, (_, i) => Math.floor(i * step) % 18),
      Array.from({ length: count }, (_, i) => (Math.floor(i * step) + 1) % 18)
    ];
  }

  const VALID_SLOT_SETS = getValidSlotSets(rotatingEmployees.length);

  let bestResultRows: Record<string, Record<string, string>> = {};
  let bestReport: ScheduleValidationReport | null = null;
  let bestRelaxations: RelaxationRecord[] = [];
  let bestScoreValue = Infinity;
  let bestHandoffSafe = false;
  let finalSeed = seed;

  const nextMonthMinRotA = [1, 2, 3].map((dayOfMonth) => {
    const d = new Date(targetYear, targetMonth + 1, dayOfMonth);
    const y = d.getFullYear();
    const m = d.getMonth() + 1;
    const iso = `${y}-${String(m).padStart(2, '0')}-${String(dayOfMonth).padStart(2, '0')}`;
    const dow = d.getDay();
    const isOff = dow === 0 || dow === 6 || ROMANIAN_HOLIDAYS.has(iso) || FIXED_STAFF_EXEMPT_OFF_DATES.has(iso);
    const fixedA = (!isOff && fixedWeekdayMAEmployees.length > 0 && getIsoWeekNumber(d) % 2 === 0)
      ? fixedWeekdayMAEmployees.length
      : 0;
    return Math.max(0, 2 - fixedA);
  });

  const isMonthEndHandoffSafe = (rows: Record<string, Record<string, string>>): boolean => {
    if (rotatingEmployees.length < 6) return true;
    const isRest = (s: string) => s === 'OFF' || s === 'H8' || s === 'HUP' || s === 'ABS' || s === 'AD';
    const tails = rotatingEmployees.map((emp) => {
      const len = Math.min(6, dateColumns.length);
      const tail: string[] = [];
      for (let idx = dateColumns.length - len; idx < dateColumns.length; idx++) {
        const s = shiftKey(rows[emp]?.[dateColumns[idx].isoDate] ?? 'OFF');
        tail.push(isRest(s) ? 'OFF' : s);
      }
      let tailWorkStreak = 0;
      let tailOffStreak = 0;
      let tailWorkShift: string | null = null;
      let lastWorkShift: string | null = null;
      for (let i = tail.length - 1; i >= 0; i--) {
        if (tail[i] !== 'OFF') {
          if (!lastWorkShift) lastWorkShift = tail[i];
          if (tailOffStreak === 0) {
            tailWorkStreak++;
            tailWorkShift = tail[i];
          } else {
            break;
          }
        } else {
          if (tailWorkStreak > 0) break;
          tailOffStreak++;
        }
      }
      return { tailWorkShift, tailWorkStreak, tailOffStreak, lastWorkShift };
    });

    for (const k of [1, 2, 3]) {
      let eligN = 0;
      let eligA = 0;
      let eligM = 0;
      for (const t of tails) {
        const canCover = (target: 'M' | 'A' | 'N', prevShift: 'M' | 'A' | 'N') => {
          if (t.tailWorkShift === target && t.tailWorkStreak <= 5 - k) return true;
          if (t.tailWorkShift === null && t.lastWorkShift === prevShift && t.tailOffStreak >= (k === 1 ? 2 : 1)) return true;
          if (k === 3 && t.tailWorkShift === prevShift && t.tailWorkStreak >= 3) return true;
          return false;
        };
        if (canCover('N', 'M')) eligN++;
        if (canCover('A', 'N')) eligA++;
        if (canCover('M', 'A')) eligM++;
      }
      if (eligN < 2 || eligA < nextMonthMinRotA[k - 1] || eligM < 1) {
        return false;
      }
    }
    return true;
  };

  interface CandidateConfig {
    template: string[];
    slots: number[];
    matching: number[];
    cost: number;
    attemptSeed: number;
    signature: string;
  }

  const templateRotations = BALANCED_TEMPLATES.map((tmpl) =>
    tmpl.map((_, off) => tmpl.map((_, i) => tmpl[(off + i) % tmpl.length][0]).join(''))
  );
  const seenAssignments = new Set<string>();

  const rotatingHistories: Record<string, string[]> = {};
  const rotatingRawHistories: Record<string, string[]> = {};
  const rotatingAnchors: Record<string, string | null> = {};
  rotatingEmployees.forEach((emp) => {
    const hist = getEmployeeHistoryTail(roster, emp, targetYear, targetMonth, 6);
    rotatingHistories[emp] = hist;
    rotatingRawHistories[emp] = extractRawHistoryTail(roster, emp, targetYear, targetMonth, 6);
    rotatingAnchors[emp] = hist.length > 0 ? roster?.rows?.[emp]?.[targetStartIso] ?? null : null;
  });

  const SEED_OFFSETS = [0, 1, 3, 5, 7, 10, 100];

  const fixedAfternoonCoverage = dateColumns.map(({ date, isoDate }) => {
    const dow = date.getDay();
    const isOff = dow === 0 || dow === 6 || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
    if (isOff || fixedWeekdayMAEmployees.length === 0) return 0;
    return getIsoWeekNumber(date) % 2 === 0 ? fixedWeekdayMAEmployees.length : 0;
  });

  const getSlotImbalancePenalty = (template: string[], slots: number[]): number => {
    const tLen = template.length;
    let penalty = 0;
    for (let d = 0; d < daysCount; d++) {
      let m = 0;
      let a = fixedAfternoonCoverage[d];
      let n = 0;
      for (const offset of slots) {
        const s = template[(offset + d) % tLen];
        if (s === 'M') m++;
        else if (s === 'A') a++;
        else if (s === 'N') n++;
      }
      if (m < 1) penalty += 1 - m;
      if (a < 2) penalty += 2 - a;
      if (n !== 2) penalty += Math.abs(n - 2);
    }
    return penalty * 40;
  };

  const scoreReport = (rep: ScheduleValidationReport, handoffSafe: boolean) => {
    const legalHard =
      rep.streakViolationCount +
      rep.nightRestViolationCount +
      rep.excessiveRestViolationCount +
      rep.mixedBlockViolationCount +
      rep.workDaysMismatchCount +
      (rep.fixedStaffViolationCount ?? 0);
    return (
      legalHard * 10000 +
      rep.nightOverstaffedDaysCount * 200 +
      rep.understaffedDaysCount * 100 +
      (handoffSafe ? 0 : 500)
    );
  };

  const stage2DeadlineMs = startTimeMs + Math.min(250, Math.floor(timeBudgetMs * 0.15));
  let stage2FoundTier2Safe = false;

  for (const sOffset of SEED_OFFSETS) {
    if (bestReport && (stage2FoundTier2Safe || performance.now() >= stage2DeadlineMs || performance.now() >= totalDeadlineMs)) {
      break;
    }
    const effSeed = seed + sOffset;
    const allCandidates: CandidateConfig[] = [];

    // Stage 1: Fast boundary screening across all template and slot configurations (< 10ms)
    for (let t = 0; t < BALANCED_TEMPLATES.length; t++) {
      for (let s = 0; s < VALID_SLOT_SETS.length; s++) {
        const templateIdx = (effSeed + t) % BALANCED_TEMPLATES.length;
        const slotSetIdx = (Math.floor(effSeed / BALANCED_TEMPLATES.length) + s) % VALID_SLOT_SETS.length;
        const template = BALANCED_TEMPLATES[templateIdx];
        const slots = VALID_SLOT_SETS[slotSetIdx];
        const currentAttemptSeed = effSeed + (t * VALID_SLOT_SETS.length + s) * 101;

        const costMatrix: number[][] = [];
        for (let eIdx = 0; eIdx < rotatingEmployees.length; eIdx++) {
          const emp = rotatingEmployees[eIdx];
          const history = rotatingHistories[emp];
          const anchor = rotatingAnchors[emp];
          costMatrix[eIdx] = [];
          for (let sIdx = 0; sIdx < slots.length; sIdx++) {
            costMatrix[eIdx][sIdx] = evaluateOffsetCost(
              history,
              template,
              slots[sIdx],
              emp,
              daysCount,
              anchor,
              currentAttemptSeed,
              targetWorkingDays
            );
          }
        }

        const matching = minCostMatching(costMatrix);
        let bestCost = getSlotImbalancePenalty(template, slots);
        for (let eIdx = 0; eIdx < rotatingEmployees.length; eIdx++) {
          bestCost += costMatrix[eIdx][matching[eIdx]];
        }

        const signature = matching
          .map((sIdx) => templateRotations[templateIdx][slots[sIdx]])
          .join('|');

        allCandidates.push({
          template,
          slots,
          matching,
          cost: bestCost,
          attemptSeed: currentAttemptSeed,
          signature
        });
      }
    }

    // Sort candidates by boundary cost ascending (strictly legal combinations first)
    allCandidates.sort((a, b) => a.cost - b.cost);

    // Stage 2: Equalize and validate distinct candidate configurations (early exit on minTier compliant + handoffSafe)
    const configLimit = sOffset === 0 ? 48 : 24;
    let tested = 0;

    for (const config of allCandidates) {
      if (bestReport && (stage2FoundTier2Safe || performance.now() >= stage2DeadlineMs || performance.now() >= totalDeadlineMs)) {
        break;
      }
      if (seenAssignments.has(config.signature)) continue;
      seenAssignments.add(config.signature);
      if (++tested > configLimit) break;
      const candidateRows: Record<string, Record<string, string>> = {};
      allEmployees.forEach((emp) => {
        candidateRows[emp] = {};
        if (excludedSet.has(emp)) {
          dateColumns.forEach(({ isoDate }) => {
            candidateRows[emp][isoDate] = 'OFF';
          });
        }
      });

      fixedMidEmployees.forEach((emp) => {
        dateColumns.forEach(({ date, isoDate }) => {
          const dow = date.getDay();
          const isWeekend = dow === 0 || dow === 6;
          const isOff = isWeekend || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
          if (isOff) {
            candidateRows[emp][isoDate] = 'OFF';
          } else {
            candidateRows[emp][isoDate] = 'MID';
          }
        });
      });

      fixedWeekdayMAEmployees.forEach((emp) => {
        dateColumns.forEach(({ date, isoDate }) => {
          const dow = date.getDay();
          const isWeekend = dow === 0 || dow === 6;
          const isOff = isWeekend || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
          if (isOff) {
            candidateRows[emp][isoDate] = 'OFF';
          } else {
            const weekNum = getIsoWeekNumber(date);
            candidateRows[emp][isoDate] = (weekNum % 2 === 1) ? 'M' : 'A';
          }
        });
      });

      const tLen = config.template.length;
      rotatingEmployees.forEach((emp, eIdx) => {
        const slotOffset = config.slots[config.matching[eIdx]];
        dateColumns.forEach(({ index, isoDate }) => {
          candidateRows[emp][isoDate] = config.template[(slotOffset + index) % tLen];
        });
      });

      equalizeExactWorkingDays(
        candidateRows,
        dateColumns,
        targetWorkingDays,
        roster,
        targetYear,
        targetMonth,
        options.excludedEmployees ?? []
      );
      const candRelax: RelaxationRecord[] = [];
      const testReport = validateScheduleCompliance(
        candidateRows,
        dateColumns,
        roster,
        targetYear,
        targetMonth,
        targetWorkingDays,
        options.excludedEmployees ?? [],
        { tier: minTier, relaxations: candRelax }
      );

      const curHandoffSafe = isMonthEndHandoffSafe(candidateRows);
      const tier2Rep =
        minTier === 2
          ? testReport
          : validateScheduleCompliance(
              candidateRows,
              dateColumns,
              roster,
              targetYear,
              targetMonth,
              targetWorkingDays,
              options.excludedEmployees ?? [],
              { tier: 2 }
            );
      let hardCheckViol = 0;
      if (!tier2Rep.isFullyCompliant && minTier < 3) {
        const hardRep =
          minTier === 3
            ? testReport
            : validateScheduleCompliance(
                candidateRows,
                dateColumns,
                roster,
                targetYear,
                targetMonth,
                targetWorkingDays,
                options.excludedEmployees ?? [],
                { tier: 3 }
              );
        hardCheckViol = hardRep.totalViolations;
      }
      const curScore =
        hardCheckViol * 10000000 +
        (tier2Rep.isFullyCompliant ? 0 : 100000) +
        scoreReport(tier2Rep, curHandoffSafe) * 10 +
        scoreReport(testReport, curHandoffSafe);

      if (!bestReport || curScore < bestScoreValue) {
        bestResultRows = candidateRows;
        bestReport = testReport;
        bestRelaxations = candRelax;
        bestScoreValue = curScore;
        bestHandoffSafe = curHandoffSafe;
        finalSeed = config.attemptSeed;
        if (testReport.isFullyCompliant && curHandoffSafe) {
          break;
        }
        if (minTier < 2 && tier2Rep.isFullyCompliant && curHandoffSafe) {
          stage2FoundTier2Safe = true;
          break;
        }
      }
    }

    if ((bestReport?.isFullyCompliant && bestHandoffSafe) || stage2FoundTier2Safe) {
      break;
    }
  }

  if (bestReport?.isFullyCompliant && bestHandoffSafe) {
    return {
      month: targetMonth,
      year: targetYear,
      daysCount,
      targetWorkingDays,
      dateColumns,
      rows: bestResultRows,
      validation: bestReport,
      seed: finalSeed,
      excludedEmployees: options.excludedEmployees ?? [],
      tier: minTier,
      relaxations: bestRelaxations,
      needsManualReview: false,
      elapsedMs: Math.round(performance.now() - startTimeMs)
    };
  }

  const stage2BestRows = bestReport ? bestResultRows : null;
  const stage2BestSeed = finalSeed;

  const tryStage2AtTier = (t: ComplianceTier): GeneratedScheduleResult | null => {
    if (!stage2BestRows) return null;
    const s2Relax: RelaxationRecord[] = [];
    const s2Rep = validateScheduleCompliance(
      stage2BestRows,
      dateColumns,
      roster,
      targetYear,
      targetMonth,
      targetWorkingDays,
      options.excludedEmployees ?? [],
      { tier: t, relaxations: s2Relax }
    );
    if (s2Rep.isFullyCompliant) {
      return {
        month: targetMonth,
        year: targetYear,
        daysCount,
        targetWorkingDays,
        dateColumns,
        rows: stage2BestRows,
        validation: s2Rep,
        seed: stage2BestSeed,
        excludedEmployees: options.excludedEmployees ?? [],
        tier: t,
        relaxations: s2Relax,
        needsManualReview: false,
        elapsedMs: Math.round(performance.now() - startTimeMs)
      };
    }
    return null;
  };

  interface FallbackAttempt {
    rows: Record<string, Record<string, string>>;
    seed: number;
    handoffSafe: boolean;
  }
  const fallbackAttempts: FallbackAttempt[] = [];
  if (stage2BestRows) {
    fallbackAttempts.push({
      rows: stage2BestRows,
      seed: stage2BestSeed,
      handoffSafe: bestHandoffSafe
    });
  }

  // Stage 3: Exact Legal-Trajectory DP + Large Neighborhood Search across active tiers
  if (rotatingEmployees.length > 0 && targetWorkingDays > 0 && targetWorkingDays <= daysCount) {
    const ACT_SHIFT = ['A', 'M', 'N', 'OFF'] as const;
    const NEXT_SH = DAG_NEXT_SHIFT_ACT;
    const shToIdx = (s: string) => (s === 'A' ? 0 : s === 'M' ? 1 : s === 'N' ? 2 : 3);
    const numStatesPerDay = (targetWorkingDays + 1) * 4 * 13 * 2;
    const encState = encodeStage3DagState;

    interface DagEdge {
      act: number;
      nextNode: number;
      pen: number;
    }
    interface EmpDag {
      layers: DagEdge[][][]; // layers[d][nodeIdx] -> outgoing edges to layers[d+1]
      feasible: boolean;
      maxNodes: number;
      cAct: Int8Array[];
      cNext: Int32Array[];
    }

    const buildEmpDag = (
      history: string[],
      minNonNightRest: number,
      allowShortBlock2: boolean,
      adjLeaveBefore: boolean
    ): EmpDag => {
      const hasHistory = history.length > 0;
      const activeTailWork = getActiveTailBlock(history);
      const activeTailRest = activeTailWork.length === 0 ? getActiveTailRest(history) : [];
      let lastHistWork: string | null = null;
      for (let i = history.length - 1; i >= 0; i--) {
        if (history[i] === 'A' || history[i] === 'M' || history[i] === 'N') {
          lastHistWork = history[i];
          break;
        }
      }

      // Compute outgoing transitions from virtual start node (at d = 0 -> state at d = 1)
      const startEdges: Array<{ act: number; nextState: number; pen: number }> = [];
      if (activeTailWork.length > 0) {
        const curShift = activeTailWork[activeTailWork.length - 1];
        const curIdx = shToIdx(curShift);
        const totalWorkTail = activeTailWork.length;
        let sameLen = 0;
        for (let i = activeTailWork.length - 1; i >= 0 && activeTailWork[i] === curShift; i--) sameLen++;
        if (totalWorkTail < 5 && targetWorkingDays >= 1) {
          const nextMl = adjLeaveBefore ? 5 + totalWorkTail : totalWorkTail;
          startEdges.push({ act: curIdx, nextState: encState(1, curIdx, nextMl, 0), pen: 0 });
        }
        if (sameLen >= 3 || adjLeaveBefore) {
          startEdges.push({ act: 3, nextState: encState(0, curIdx, 10, 0), pen: 0 });
        }
      } else if (hasHistory) {
        const restLen = activeTailRest.length;
        const lastIdx = lastHistWork ? shToIdx(lastHistWork) : 3;
        if (restLen < 3) {
          startEdges.push({ act: 3, nextState: encState(0, lastIdx, 10 + restLen, 0), pen: 0 });
        }
        const reqRest = lastIdx === 2 ? 2 : (adjLeaveBefore ? 1 : minNonNightRest);
        if (targetWorkingDays >= 1 && (lastIdx === 3 || restLen >= reqRest)) {
          const pen = lastIdx !== 3 && lastIdx !== 2 && restLen === 1 && !adjLeaveBefore ? 15 : 0;
          const firstMl = adjLeaveBefore ? 5 : 0;
          if (lastIdx === 3) {
            for (const ns of [0, 1, 2]) {
              startEdges.push({ act: ns, nextState: encState(1, ns, firstMl, 0), pen: 0 });
            }
          } else {
            const ns = NEXT_SH[lastIdx];
            startEdges.push({ act: ns, nextState: encState(1, ns, firstMl, 0), pen });
            if (adjLeaveBefore) {
              startEdges.push({ act: lastIdx, nextState: encState(1, lastIdx, firstMl, 0), pen: 0 });
            }
          }
        }
      } else {
        startEdges.push({ act: 3, nextState: encState(0, 3, 10, 0), pen: 0 });
        if (targetWorkingDays >= 1) {
          for (const ns of [0, 1, 2]) {
            startEdges.push({ act: ns, nextState: encState(1, ns, 5, 0), pen: 0 });
          }
        }
      }

      if (startEdges.length === 0) {
        return { layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
      }

      // Forward reachability sets: activeStates[d] for d = 1..daysCount
      const activeStates: number[][] = Array.from({ length: daysCount + 1 }, () => []);
      const seen = new Uint8Array(numStatesPerDay);
      for (const e of startEdges) {
        if (!seen[e.nextState]) {
          seen[e.nextState] = 1;
          activeStates[1].push(e.nextState);
        }
      }

      const stepTransitions = (state: number): Array<{ act: number; nextState: number; pen: number }> => {
        const sb = state & 1;
        const t0 = state >> 1;
        const ml = t0 % 13;
        const tmp = (t0 - ml) / 13;
        const sh = tmp % 4;
        const w = (tmp - sh) / 4;
        const out: Array<{ act: number; nextState: number; pen: number }> = [];

        if (ml < 10) {
          const isUnb = ml >= 5;
          const len = isUnb ? ml - 4 : ml + 1;
          if (len < 5 && w + 1 <= targetWorkingDays) {
            out.push({ act: sh, nextState: encState(w + 1, sh, ml + 1, sb), pen: 0 });
          }
          if (len >= 3 || isUnb) {
            out.push({ act: 3, nextState: encState(w, sh, 10, sb), pen: 0 });
          } else if (allowShortBlock2 && len === 2 && sb === 0) {
            out.push({ act: 3, nextState: encState(w, sh, 10, 1), pen: 25 });
          }
        } else {
          const len = ml - 9;
          if (len < 3) {
            out.push({ act: 3, nextState: encState(w, sh, ml + 1, sb), pen: 0 });
          }
          if (w + 1 <= targetWorkingDays) {
            if (sh === 3) {
              for (const ns of [0, 1, 2]) {
                out.push({ act: ns, nextState: encState(w + 1, ns, 0, sb), pen: 0 });
              }
            } else if (len >= (sh === 2 ? 2 : minNonNightRest)) {
              const ns = NEXT_SH[sh];
              const pen = sh !== 2 && len === 1 ? 15 : 0;
              out.push({ act: ns, nextState: encState(w + 1, ns, 0, sb), pen });
            }
          }
        }
        return out;
      };

      for (let d = 1; d < daysCount; d++) {
        seen.fill(0);
        for (const st of activeStates[d]) {
          for (const tr of stepTransitions(st)) {
            if (!seen[tr.nextState]) {
              seen[tr.nextState] = 1;
              activeStates[d + 1].push(tr.nextState);
            }
          }
        }
      }

      // Backward reachability filter from terminal states (w === targetWorkingDays at d = daysCount)
      const reachNext = new Int32Array(numStatesPerDay).fill(-1);
      let termIdx = 0;
      for (const st of activeStates[daysCount]) {
        const sb = st & 1;
        const t0 = st >> 1;
        const ml = t0 % 13;
        const tmp = (t0 - ml) / 13;
        const sh = tmp % 4;
        const w = (tmp - sh) / 4;
        if (w === targetWorkingDays && (sb === 0 || allowShortBlock2)) {
          reachNext[st] = termIdx++;
        }
      }
      if (termIdx === 0) {
        return { layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
      }

      const layers: DagEdge[][][] = new Array(daysCount);
      const reachCur = new Int32Array(numStatesPerDay);
      for (let d = daysCount - 1; d >= 1; d--) {
        reachCur.fill(-1);
        const dayLayer: DagEdge[][] = [];
        for (const st of activeStates[d]) {
          const nodeEdges: DagEdge[] = [];
          for (const tr of stepTransitions(st)) {
            const nxtIdx = reachNext[tr.nextState];
            if (nxtIdx !== -1) {
              nodeEdges.push({ act: tr.act, nextNode: nxtIdx, pen: tr.pen });
            }
          }
          if (nodeEdges.length > 0) {
            reachCur[st] = dayLayer.length;
            dayLayer.push(nodeEdges);
          }
        }
        if (dayLayer.length === 0) {
          return { layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
        }
        layers[d] = dayLayer;
        reachNext.set(reachCur);
      }

      const rootEdges: DagEdge[] = [];
      for (const e of startEdges) {
        const nxtIdx = reachNext[e.nextState];
        if (nxtIdx !== -1) {
          rootEdges.push({ act: e.act, nextNode: nxtIdx, pen: e.pen });
        }
      }
      if (rootEdges.length === 0) {
        return { layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
      }
      layers[0] = [rootEdges];
      let maxNodes = 1;
      const cAct: Int8Array[] = new Array(daysCount);
      const cNext: Int32Array[] = new Array(daysCount);
      for (let d = 0; d < daysCount; d++) {
        const len = layers[d].length;
        if (len > maxNodes) maxNodes = len;
        cAct[d] = new Int8Array(len);
        cNext[d] = new Int32Array(len);
      }
      return { layers, feasible: true, maxNodes, cAct, cNext };
    };

    const dayCostScratch = new Float64Array(daysCount * 4);
    let dpNextBuf = new Float64Array(512);
    let dpCurBuf = new Float64Array(512);

    const solveBestPath = (
      dag: EmpDag,
      otherM: Int32Array,
      otherA: Int32Array,
      otherN: Int32Array,
      dayWeights: Float64Array | null,
      noise: Float64Array | null,
      tier: ComplianceTier
    ): number[] => {
      if (dag.maxNodes > dpNextBuf.length) {
        dpNextBuf = new Float64Array(dag.maxNodes + 64);
        dpCurBuf = new Float64Array(dag.maxNodes + 64);
      }
      for (let d = 0; d < daysCount; d++) {
        const m0 = otherM[d];
        const a0 = otherA[d] + fixedAfternoonCoverage[d];
        const n0 = otherN[d];
        const dw = dayWeights ? dayWeights[d] : 1;
        const base = d * 4;
        dayCostScratch[base + 0] = tierCovPenalty(m0, a0 + 1, n0, dw, tier) + (noise ? noise[base + 0] : 0);
        dayCostScratch[base + 1] = tierCovPenalty(m0 + 1, a0, n0, dw, tier) + (noise ? noise[base + 1] : 0);
        dayCostScratch[base + 2] = tierCovPenalty(m0, a0, n0 + 1, dw, tier) + (noise ? noise[base + 2] : 0);
        dayCostScratch[base + 3] = tierCovPenalty(m0, a0, n0, dw, tier) + (noise ? noise[base + 3] : 0);
      }

      const { layers, cAct: dpChoicesAct, cNext: dpChoicesNext } = dag;

      for (let d = daysCount - 1; d >= 0; d--) {
        const layer = layers[d];
        const nNodes = layer.length;
        const cAct = dpChoicesAct[d];
        const cNext = dpChoicesNext[d];
        const base = d * 4;
        const isLast = d === daysCount - 1;
        const c0 = dayCostScratch[base + 0];
        const c1 = dayCostScratch[base + 1];
        const c2 = dayCostScratch[base + 2];
        const c3 = dayCostScratch[base + 3];

        for (let u = 0; u < nNodes; u++) {
          const edges = layer[u];
          const elen = edges.length;
          let bestVal = Infinity;
          let bestA = 3;
          let bestV = 0;
          for (let ei = 0; ei < elen; ei++) {
            const edge = edges[ei];
            const act = edge.act;
            const actCost = act === 0 ? c0 : act === 1 ? c1 : act === 2 ? c2 : c3;
            const vCost = isLast ? 0 : dpNextBuf[edge.nextNode];
            const val = actCost + edge.pen + vCost;
            if (val < bestVal) {
              bestVal = val;
              bestA = act;
              bestV = edge.nextNode;
            }
          }
          dpCurBuf[u] = bestVal;
          cAct[u] = bestA;
          cNext[u] = bestV;
        }
        const tmp = dpNextBuf;
        dpNextBuf = dpCurBuf;
        dpCurBuf = tmp;
      }

      const path = new Array<number>(daysCount);
      let curNode = 0;
      for (let d = 0; d < daysCount; d++) {
        path[d] = dpChoicesAct[d][curNode];
        curNode = dpChoicesNext[d][curNode];
      }
      return path;
    };

    const buildRowsFromPaths = (paths: number[][]): Record<string, Record<string, string>> => {
      const candidateRows: Record<string, Record<string, string>> = {};
      allEmployees.forEach((emp) => {
        candidateRows[emp] = {};
        if (excludedSet.has(emp)) {
          dateColumns.forEach(({ isoDate }) => {
            candidateRows[emp][isoDate] = 'OFF';
          });
        }
      });
      fixedMidEmployees.forEach((emp) => {
        dateColumns.forEach(({ date, isoDate }) => {
          candidateRows[emp][isoDate] = isFixedStaffOffDate(date, isoDate) ? 'OFF' : 'MID';
        });
      });
      fixedWeekdayMAEmployees.forEach((emp) => {
        dateColumns.forEach(({ date, isoDate }) => {
          candidateRows[emp][isoDate] = isFixedStaffOffDate(date, isoDate)
            ? 'OFF'
            : getIsoWeekNumber(date) % 2 === 1
              ? 'M'
              : 'A';
        });
      });
      rotatingEmployees.forEach((emp, eIdx) => {
        const path = paths[eIdx];
        dateColumns.forEach(({ index, isoDate }) => {
          candidateRows[emp][isoDate] = ACT_SHIFT[path[index]];
        });
      });
      return candidateRows;
    };

    let rngState = (seed ^ 0x9e3779b9) >>> 0;
    const nextRand = (): number => {
      rngState = (rngState + 0x6d2b79f5) >>> 0;
      let t = Math.imul(rngState ^ (rngState >>> 15), 1 | rngState);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
    const shuffleStage3 = (arr: number[]): void => {
      for (let i = arr.length - 1; i > 0; i--) {
        const j = Math.floor(nextRand() * (i + 1));
        const tmp = arr[i];
        arr[i] = arr[j];
        arr[j] = tmp;
      }
    };
    const isStage3PathValid = (dag: EmpDag, path: number[]): boolean =>
      isDagPathValid(dag, path, daysCount);
    const solveStage3ExtremeCount = (dag: EmpDag, targetAct: number, sign: 1 | -1): number =>
      solveDagExtremeCount(dag, daysCount, targetAct, sign);

    const evalCovWeighted = (
      mArr: Int32Array,
      aArr: Int32Array,
      nArr: Int32Array,
      dayWeights: Float64Array | null,
      tier: ComplianceTier
    ) => {
      let hard = 0;
      let total = 0;
      let weightedTotal = 0;
      for (let d = 0; d < daysCount; d++) {
        const m = mArr[d];
        const a = aArr[d] + fixedAfternoonCoverage[d];
        const n = nArr[d];
        if (!isCoverageFeasibleForTier(m, a, n, tier)) hard++;
        const baseP = tierCovPenalty(m, a, n, 1, tier);
        total += baseP;
        weightedTotal += dayWeights ? tierCovPenalty(m, a, n, dayWeights[d], tier) : baseP;
      }
      return { hard, total, weightedTotal };
    };

    const addPath = (path: number[], sign: number, mArr: Int32Array, aArr: Int32Array, nArr: Int32Array) => {
      for (let d = 0; d < daysCount; d++) {
        const act = path[d];
        if (act === 0) aArr[d] += sign;
        else if (act === 1) mArr[d] += sign;
        else if (act === 2) nArr[d] += sign;
      }
    };

    const buildDayActMasks = (dags: EmpDag[]): Uint8Array[] =>
      buildDagDayActMasks(dags, daysCount);

    const hasSufficientDagCapacity = (
      dags: EmpDag[],
      dayActMasks: Uint8Array[],
      tier: ComplianceTier
    ): boolean => {
      const availM = new Int32Array(daysCount);
      const availTotalA = new Int32Array(daysCount);
      const availN = new Int32Array(daysCount);
      for (let d = 0; d < daysCount; d++) {
        let mCnt = 0;
        let aCnt = fixedAfternoonCoverage[d];
        let nCnt = 0;
        for (let i = 0; i < dayActMasks.length; i++) {
          const mk = dayActMasks[i][d];
          if (mk & 1) aCnt++;
          if (mk & 2) mCnt++;
          if (mk & 4) nCnt++;
        }
        availM[d] = mCnt;
        availTotalA[d] = aCnt;
        availN[d] = nCnt;
      }
      let sumMinA = 0, sumMaxA = 0, sumMinM = 0, sumMaxM = 0, sumMinN = 0, sumMaxN = 0;
      for (let i = 0; i < dags.length; i++) {
        sumMinA += solveStage3ExtremeCount(dags[i], 0, 1);
        sumMaxA += solveStage3ExtremeCount(dags[i], 0, -1);
        sumMinM += solveStage3ExtremeCount(dags[i], 1, 1);
        sumMaxM += solveStage3ExtremeCount(dags[i], 1, -1);
        sumMinN += solveStage3ExtremeCount(dags[i], 2, 1);
        sumMaxN += solveStage3ExtremeCount(dags[i], 2, -1);
      }
      return canSatisfyDailyCapacityForTier({
        tier,
        daysCount,
        totalRotSlots: dags.length * targetWorkingDays,
        availM,
        availTotalA,
        availN,
        fixedMaxA: fixedAfternoonCoverage,
        sumMinRotA: sumMinA,
        sumMaxRotA: sumMaxA,
        sumMinM,
        sumMaxM,
        sumMinN,
        sumMaxN
      });
    };

    const selectStage3RuinGroup = (
      focusDay: number,
      k: number,
      curM: Int32Array,
      curA: Int32Array,
      curN: Int32Array,
      paths: number[][],
      dayActMasks: Uint8Array[],
      tier: ComplianceTier
    ): number[] => {
      shuffleStage3(empPool);
      if (nextRand() < 0.2) {
        return empPool.slice(0, k);
      }
      const cfg = getTierConfig(tier);
      const minReqM = cfg.allowCoverageGapMA ? 0 : 1;
      const minReqA = cfg.allowCoverageGapMA ? 1 : 2;
      const maxReqN = cfg.allowNightOver2 ? 3 : 2;
      const needAct =
        curN[focusDay] < 2
          ? 2
          : curM[focusDay] < minReqM
            ? 1
            : curA[focusDay] + fixedAfternoonCoverage[focusDay] < minReqA
              ? 0
              : curN[focusDay] > maxReqN
                ? -1
                : -2;
      const helpers: number[] = [];
      const capable: number[] = [];
      const others: number[] = [];
      for (let idx = 0; idx < empPool.length; idx++) {
        const e = empPool[idx];
        if (needAct >= 0) {
          if (canDoShift(dayActMasks, e, focusDay, needAct)) {
            if (paths[e][focusDay] !== needAct) helpers.push(e);
            else capable.push(e);
          } else {
            others.push(e);
          }
        } else if (needAct === -1 && paths[e][focusDay] === 2 && (dayActMasks[e][focusDay] & ~4) !== 0) {
          helpers.push(e);
        } else {
          others.push(e);
        }
      }
      if (helpers.length > 0) {
        const partners = [...helpers.slice(1), ...capable, ...others];
        shuffleStage3(partners);
        return [helpers[0], ...partners].slice(0, k);
      }
      if (capable.length > 0) {
        return [...capable, ...others].slice(0, k);
      }
      return empPool.slice(0, k);
    };

    const R = rotatingEmployees.length;
    const LEAVE_CODES = DAG_LEAVE_CODES;
    const stage3EmpData = rotatingEmployees.map((emp) => {
      const rawHist = rotatingRawHistories[emp] ?? [];
      const sanitizedHist = rawHist.map((s) =>
        s === 'HUP' || s === 'H8' || s === 'ABS' || s === 'AD' || s === 'CM' ? 'OFF' : s
      );
      const activeTailWork = getActiveTailBlock(sanitizedHist);
      const idxBefore = rawHist.length - 1 - activeTailWork.length;
      const adjLeaveBefore = idxBefore >= 0 && LEAVE_CODES.has(rawHist[idxBefore]);
      return { sanitizedHist, adjLeaveBefore };
    });

    const priorBestPaths = stage2BestRows
      ? rotatingEmployees.map((emp) =>
          dateColumns.map(({ isoDate }) => shToIdx(shiftKey(stage2BestRows[emp]?.[isoDate] ?? 'OFF')))
        )
      : [];
    const dayWeights = new Float64Array(daysCount);
    const noise = new Float64Array(daysCount * 4);
    const empPool = Array.from({ length: R }, (_, i) => i);

    let hasDagFallback = false;
    const addQuickDagFallback = (dags: EmpDag[], tier: ComplianceTier, fallbackSeed: number) => {
      if (hasDagFallback || dags.some((g) => !g.feasible)) return;
      const curM = new Int32Array(daysCount);
      const curA = new Int32Array(daysCount);
      const curN = new Int32Array(daysCount);
      const paths: number[][] = new Array(R);
      const order = Array.from({ length: R }, (_, i) => i);
      shuffleStage3(order);
      for (let oi = 0; oi < R; oi++) {
        const eIdx = order[oi];
        for (let k = 0; k < noise.length; k++) noise[k] = nextRand() * 6;
        const p = solveBestPath(dags[eIdx], curM, curA, curN, null, noise, tier);
        paths[eIdx] = p;
        addPath(p, 1, curM, curA, curN);
      }
      for (let sweep = 0; sweep < 2; sweep++) {
        let changed = false;
        for (let oi = 0; oi < R; oi++) {
          const eIdx = order[(oi + sweep) % R];
          const oldP = paths[eIdx];
          const before = evalCovWeighted(curM, curA, curN, null, tier).weightedTotal;
          addPath(oldP, -1, curM, curA, curN);
          const newP = solveBestPath(dags[eIdx], curM, curA, curN, null, null, tier);
          addPath(newP, 1, curM, curA, curN);
          const after = evalCovWeighted(curM, curA, curN, null, tier).weightedTotal;
          if (after < before) {
            paths[eIdx] = newP;
            changed = true;
          } else if (after > before) {
            addPath(newP, -1, curM, curA, curN);
            addPath(oldP, 1, curM, curA, curN);
          } else {
            paths[eIdx] = newP;
          }
        }
        if (!changed) break;
      }
      const candidateRows = buildRowsFromPaths(paths);
      fallbackAttempts.push({
        rows: candidateRows,
        seed: fallbackSeed,
        handoffSafe: isMonthEndHandoffSafe(candidateRows)
      });
      hasDagFallback = true;
    };

    for (let ti = 0; ti < activeTiers.length; ti++) {
      const tier = activeTiers[ti];
      const cfg = getTierConfig(tier);
      const dags = stage3EmpData.map(({ sanitizedHist, adjLeaveBefore }) =>
        buildEmpDag(sanitizedHist, cfg.minRest, cfg.allowShortBlock2, cfg.allowShortBlock2 && adjLeaveBefore)
      );
      if (dags.some((g) => !g.feasible)) {
        const s2Hit = tryStage2AtTier(tier);
        if (s2Hit) return s2Hit;
        continue;
      }

      const nowMs = performance.now();
      if (nowMs >= totalDeadlineMs) {
        addQuickDagFallback(dags, tier, seed);
        for (let tj = ti; tj < activeTiers.length; tj++) {
          const s2Hit = tryStage2AtTier(activeTiers[tj]);
          if (s2Hit) return s2Hit;
        }
        break;
      }

      const remBudgetMs = Math.max(0, totalDeadlineMs - nowMs);
      const remShares = activeTiers
        .slice(ti)
        .reduce((sum, t) => sum + getTierConfig(t).budgetShare, 0);
      const tierAllocMs = remShares > 0 ? (remBudgetMs * cfg.budgetShare) / remShares : remBudgetMs;
      const tierDeadlineMs = Math.min(totalDeadlineMs, nowMs + tierAllocMs);

      const dayActMasks = buildDayActMasks(dags);
      if (!hasSufficientDagCapacity(dags, dayActMasks, tier)) {
        if (ti === activeTiers.length - 1 || !hasDagFallback) {
          addQuickDagFallback(dags, tier, seed);
        }
        const s2Hit = tryStage2AtTier(tier);
        if (s2Hit) return s2Hit;
        continue;
      }

      const canSeedStage2 =
        priorBestPaths.length === R &&
        priorBestPaths.every((p, idx) => isStage3PathValid(dags[idx], p));
      const numRestarts = tier === 0 ? 6 : tier === 1 ? 3 : 10;
      let compliantRestarts = 0;
      let tierBestWinner: {
        rows: Record<string, Record<string, string>>;
        report: ScheduleValidationReport;
        relaxations: RelaxationRecord[];
        handoffSafe: boolean;
        seed: number;
        score: number;
      } | null = null;

      if (stage2BestRows) {
        const s2Relax: RelaxationRecord[] = [];
        const s2Rep = validateScheduleCompliance(
          stage2BestRows,
          dateColumns,
          roster,
          targetYear,
          targetMonth,
          targetWorkingDays,
          options.excludedEmployees ?? [],
          { tier, relaxations: s2Relax }
        );
        if (s2Rep.isFullyCompliant) {
          const s2Safe = isMonthEndHandoffSafe(stage2BestRows);
          tierBestWinner = {
            rows: stage2BestRows,
            report: s2Rep,
            relaxations: s2Relax,
            handoffSafe: s2Safe,
            seed: stage2BestSeed,
            score: (s2Safe ? 0 : 10000) + s2Relax.length * 10
          };
          if (s2Safe) {
            return {
              month: targetMonth,
              year: targetYear,
              daysCount,
              targetWorkingDays,
              dateColumns,
              rows: tierBestWinner.rows,
              validation: tierBestWinner.report,
              seed: tierBestWinner.seed,
              excludedEmployees: options.excludedEmployees ?? [],
              tier,
              relaxations: tierBestWinner.relaxations,
              needsManualReview: false,
              elapsedMs: Math.round(performance.now() - startTimeMs)
            };
          }
        }
      }

      for (let restart = 0; restart < numRestarts; restart++) {
        if (performance.now() >= tierDeadlineMs) break;
        if (tierBestWinner?.handoffSafe) break;
        if (tierBestWinner && (compliantRestarts >= 3 || restart >= 4)) break;

        dayWeights.fill(1);
        let paths: number[][] = new Array(R);
        const curM = new Int32Array(daysCount);
        const curA = new Int32Array(daysCount);
        const curN = new Int32Array(daysCount);
        const order = Array.from({ length: R }, (_, i) => i);
        shuffleStage3(order);

        if (restart === 0 && canSeedStage2) {
          for (let eIdx = 0; eIdx < R; eIdx++) {
            paths[eIdx] = priorBestPaths[eIdx];
            addPath(paths[eIdx], 1, curM, curA, curN);
          }
        } else {
          // Greedy randomized initialization
          for (let oi = 0; oi < R; oi++) {
            const eIdx = order[oi];
            if (restart > 0 && oi < 3) {
              for (let k = 0; k < noise.length; k++) noise[k] = nextRand() * 250;
            } else {
              for (let k = 0; k < noise.length; k++) noise[k] = nextRand() * (restart === 0 ? 10 : 8);
            }
            const p = solveBestPath(dags[eIdx], curM, curA, curN, null, noise, tier);
            paths[eIdx] = p;
            addPath(p, 1, curM, curA, curN);
          }
        }

        // Coordinate-descent best-response sweeps
        const runSweeps = (maxSweeps: number, useWeights: boolean) => {
          const wArr = useWeights ? dayWeights : null;
          for (let sweep = 0; sweep < maxSweeps; sweep++) {
            let changed = false;
            for (let oi = 0; oi < R; oi++) {
              const eIdx = order[(oi + sweep) % R];
              const oldP = paths[eIdx];
              const before = evalCovWeighted(curM, curA, curN, wArr, tier).weightedTotal;
              addPath(oldP, -1, curM, curA, curN);
              const newP = solveBestPath(dags[eIdx], curM, curA, curN, wArr, null, tier);
              addPath(newP, 1, curM, curA, curN);
              const after = evalCovWeighted(curM, curA, curN, wArr, tier).weightedTotal;
              if (after < before) {
                paths[eIdx] = newP;
                changed = true;
              } else if (after > before) {
                addPath(newP, -1, curM, curA, curN);
                addPath(oldP, 1, curM, curA, curN);
              } else {
                paths[eIdx] = newP;
              }
            }
            if (!changed) break;
          }
        };

        if (!(restart === 0 && canSeedStage2)) {
          runSweeps(3, false);
        }

        // Multi-employee Ruin-and-Recreate Simulated Annealing + Guided Local Search dayWeights
        let curEval = evalCovWeighted(curM, curA, curN, dayWeights, tier);
        let bestRestartHard = curEval.hard;
        let bestRestartTotal = curEval.total;
        let bestRestartPaths = paths.map((p) => p);

        let temp = 1600;
        const cool = 0.994;
        const saSteps = curEval.hard > 0 ? 1500 : 0;
        for (let step = 0; step < saSteps && bestRestartHard > 0; step++) {
          if ((step & 15) === 0 && performance.now() >= tierDeadlineMs) break;

          const badDays: number[] = [];
          for (let d = 0; d < daysCount; d++) {
            if (!isCoverageFeasibleForTier(curM[d], curA[d] + fixedAfternoonCoverage[d], curN[d], tier)) {
              badDays.push(d);
            }
          }
          if (badDays.length === 0) break;

          // Guided Local Search: boost weights on stubborn bad days every 35 steps
          if (step > 0 && step % 35 === 0) {
            for (const bd of badDays) {
              dayWeights[bd] = Math.min(8, dayWeights[bd] + 0.5);
            }
            if (step % 210 === 0) {
              for (let d = 0; d < daysCount; d++) dayWeights[d] = 1 + (dayWeights[d] - 1) * 0.5;
            }
            curEval = evalCovWeighted(curM, curA, curN, dayWeights, tier);
          }

          const focusDay = badDays[Math.floor(nextRand() * badDays.length)];
          const k = Math.min(R, step % 3 === 0 ? 3 : 2);
          const ruinGroup = selectStage3RuinGroup(focusDay, k, curM, curA, curN, paths, dayActMasks, tier);

          const savedSub: number[][] = ruinGroup.map((idx) => paths[idx]);

          for (let gi = 0; gi < ruinGroup.length; gi++) {
            addPath(paths[ruinGroup[gi]], -1, curM, curA, curN);
          }

          const noiseAmp = Math.min(350, temp * 0.32);
          for (let gi = 0; gi < ruinGroup.length; gi++) {
            const eIdx = ruinGroup[gi];
            if (gi === 0 && nextRand() < 0.5) {
              for (let nk = 0; nk < noise.length; nk++) noise[nk] = nextRand() * 6;
              for (let d = Math.max(0, focusDay - 5); d <= Math.min(daysCount - 1, focusDay + 5); d++) {
                for (let a = 0; a < 4; a++) noise[d * 4 + a] += (nextRand() - 0.5) * noiseAmp * 3;
              }
            } else {
              for (let nk = 0; nk < noise.length; nk++) noise[nk] = nextRand() * (4 + noiseAmp * 0.25);
            }
            const newP = solveBestPath(dags[eIdx], curM, curA, curN, dayWeights, noise, tier);
            paths[eIdx] = newP;
            addPath(newP, 1, curM, curA, curN);
          }

          let nextEval = evalCovWeighted(curM, curA, curN, dayWeights, tier);
          const improvedRecord = nextEval.hard < bestRestartHard;
          if (improvedRecord) {
            runSweeps(2, true);
            nextEval = evalCovWeighted(curM, curA, curN, dayWeights, tier);
          }

          const delta = nextEval.weightedTotal - curEval.weightedTotal;
          const accept =
            improvedRecord ||
            delta <= 0 ||
            nextRand() < Math.exp(-delta / Math.max(1, temp));

          if (accept) {
            curEval = nextEval;
            if (
              curEval.hard < bestRestartHard ||
              (curEval.hard === bestRestartHard && curEval.total < bestRestartTotal)
            ) {
              bestRestartHard = curEval.hard;
              bestRestartTotal = curEval.total;
              bestRestartPaths = paths.map((p) => p);
              if (bestRestartHard === 0) break;
            }
          } else {
            for (let gi = 0; gi < ruinGroup.length; gi++) {
              const eIdx = ruinGroup[gi];
              addPath(paths[eIdx], -1, curM, curA, curN);
              paths[eIdx] = savedSub[gi];
              addPath(savedSub[gi], 1, curM, curA, curN);
            }
          }

          temp *= cool;
        }

        paths = bestRestartPaths;
        curM.fill(0);
        curA.fill(0);
        curN.fill(0);
        for (let idx = 0; idx < R; idx++) addPath(paths[idx], 1, curM, curA, curN);
        runSweeps(2, false);

        const candidateRows = buildRowsFromPaths(paths);
        const candRelax: RelaxationRecord[] = [];
        const testReport = validateScheduleCompliance(
          candidateRows,
          dateColumns,
          roster,
          targetYear,
          targetMonth,
          targetWorkingDays,
          options.excludedEmployees ?? [],
          { tier, relaxations: candRelax }
        );
        const curHandoffSafe = isMonthEndHandoffSafe(candidateRows);
        fallbackAttempts.push({
          rows: candidateRows,
          seed: seed + restart,
          handoffSafe: curHandoffSafe
        });
        hasDagFallback = true;

        if (testReport.isFullyCompliant) {
          compliantRestarts++;
          const wScore = (curHandoffSafe ? 0 : 10000) + candRelax.length * 10;
          if (!tierBestWinner || wScore < tierBestWinner.score) {
            tierBestWinner = {
              rows: candidateRows,
              report: testReport,
              relaxations: candRelax,
              handoffSafe: curHandoffSafe,
              seed: seed + restart,
              score: wScore
            };
            if (curHandoffSafe) break;
          }
        }
      }

      if (!hasDagFallback) {
        addQuickDagFallback(dags, tier, seed);
      }

      if (tierBestWinner) {
        return {
          month: targetMonth,
          year: targetYear,
          daysCount,
          targetWorkingDays,
          dateColumns,
          rows: tierBestWinner.rows,
          validation: tierBestWinner.report,
          seed: tierBestWinner.seed,
          excludedEmployees: options.excludedEmployees ?? [],
          tier,
          relaxations: tierBestWinner.relaxations,
          needsManualReview: false,
          elapsedMs: Math.round(performance.now() - startTimeMs)
        };
      }

      const s2Hit = tryStage2AtTier(tier);
      if (s2Hit) return s2Hit;
    }
  }

  // No active tier achieved totalViolations === 0 (or Stage 3 skipped):
  // pick best attempt prioritizing 0 hard H1-H8 violations, then lowest totalViolations at maxTier.
  let chosenFallback: FallbackAttempt | null = null;
  let chosenScore = Infinity;
  for (const att of fallbackAttempts) {
    const hardRep = validateScheduleCompliance(
      att.rows,
      dateColumns,
      roster,
      targetYear,
      targetMonth,
      targetWorkingDays,
      options.excludedEmployees ?? [],
      { tier: 3 }
    );
    const maxRep = maxTier === 3
      ? hardRep
      : validateScheduleCompliance(
          att.rows,
          dateColumns,
          roster,
          targetYear,
          targetMonth,
          targetWorkingDays,
          options.excludedEmployees ?? [],
          { tier: maxTier }
        );
    const hardLegal =
      hardRep.streakViolationCount +
      hardRep.nightRestViolationCount +
      hardRep.excessiveRestViolationCount +
      hardRep.mixedBlockViolationCount +
      hardRep.workDaysMismatchCount +
      (hardRep.fixedStaffViolationCount ?? 0);
    const covTier3 = hardRep.understaffedDaysCount + hardRep.nightOverstaffedDaysCount;
    const sc =
      hardLegal * 1e9 +
      covTier3 * 1e5 +
      maxRep.totalViolations * 100 +
      (att.handoffSafe ? 0 : 1);
    if (!chosenFallback || sc < chosenScore) {
      chosenFallback = att;
      chosenScore = sc;
    }
  }

  const finalRows = chosenFallback?.rows ?? bestResultRows;
  const finalOutSeed = chosenFallback?.seed ?? finalSeed;
  const finalRelax: RelaxationRecord[] = [];
  const validation = validateScheduleCompliance(
    finalRows,
    dateColumns,
    roster,
    targetYear,
    targetMonth,
    targetWorkingDays,
    options.excludedEmployees ?? [],
    { tier: maxTier, relaxations: finalRelax }
  );

  return {
    month: targetMonth,
    year: targetYear,
    daysCount,
    targetWorkingDays,
    dateColumns,
    rows: finalRows,
    validation,
    seed: finalOutSeed,
    excludedEmployees: options.excludedEmployees ?? [],
    tier: maxTier,
    relaxations: finalRelax,
    needsManualReview: !validation.isFullyCompliant,
    elapsedMs: Math.round(performance.now() - startTimeMs)
  };
}

/**
 * Validates all legal and staffing constraints of a schedule
 */
export function validateScheduleCompliance(
  rows: Record<string, Record<string, string>>,
  dateColumns: { index: number; date: Date; isoDate: string }[],
  priorRoster: RosterData | null = null,
  targetYear?: number,
  targetMonth?: number,
  targetWorkingDays?: number,
  excludedEmployees: string[] = [],
  options?: ValidationOptions
): ScheduleValidationReport {
  const tier: ComplianceTier = options?.tier ?? 0;
  const tierConfig = getTierConfig(tier);
  const excludedSet = new Set(excludedEmployees);
  const dailyCoverage: Record<string, DayCoverage> = {};
  const employeeStats: Record<string, EmployeeMonthStats> = {};
  const warnings: string[] = [];
  let understaffedDaysCount = 0;
  let nightOverstaffedDaysCount = 0;
  let streakViolationCount = 0;
  let nightRestViolationCount = 0;
  let excessiveRestViolationCount = 0;
  let mixedBlockViolationCount = 0;
  let workDaysMismatchCount = 0;
  let fixedStaffViolationCount = 0;

  const WORK_SHIFTS = new Set(['M', 'A', 'N', 'MID']);
  const ROTATING_SHIFTS = new Set(['M', 'A', 'N']);
  const LEAVE_OR_SPLIT_CODES = new Set(['ABS', 'H8', 'CM']);
  const VALID_CODES = new Set(['M', 'A', 'N', 'MID', 'OFF', 'ABS', 'H8']);
  const EXPECTED_NEXT_SHIFT: Record<string, string> = { A: 'M', M: 'N', N: 'A' };

  const normalizeShiftCode = (raw: string | undefined | null): string => {
    if (raw === undefined || raw === null) return 'OFF';
    const s = shiftKey(String(raw));
    return s === '' ? 'OFF' : s;
  };

  const parseIsoUtc = (isoDate: string): Date => {
    const [y, m, d] = isoDate.split('-').map(Number);
    return new Date(Date.UTC(y, m - 1, d));
  };

  const addDaysIso = (isoDate: string, delta: number): string => {
    const dt = parseIsoUtc(isoDate);
    dt.setUTCDate(dt.getUTCDate() + delta);
    return dt.toISOString().slice(0, 10);
  };

  const cols = [...(dateColumns || [])].sort((a, b) => a.isoDate.localeCompare(b.isoDate));
  const monthLen = cols.length;
  const monthIsos = cols.map((c) => c.isoDate);

  const calMeta = monthIsos.map((isoDate) => {
    const utcDt = parseIsoUtc(isoDate);
    const dow = utcDt.getUTCDay();
    const isWeekend = dow === 0 || dow === 6;
    const isFixedOff =
      isWeekend || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
    const isoWeek = getIsoWeekNumber(
      new Date(utcDt.getUTCFullYear(), utcDt.getUTCMonth(), utcDt.getUTCDate())
    );
    return { isoDate, dow, isWeekend, isFixedOff, isoWeek };
  });

  const fixedMonthWorkdays = calMeta.filter((m) => !m.isFixedOff).length;
  const employees = Object.keys(rows);

  // Extract contiguous <= 6-day prior history ending at monthIsos[0] - 1, preserving H8/ABS/CM codes.
  const knownHistDates = new Set<string>();
  if (priorRoster) {
    for (const dc of priorRoster.dateColumns || []) {
      if (dc?.isoDate) knownHistDates.add(dc.isoDate);
    }
    if (priorRoster.rows) {
      for (const rEmp of Object.keys(priorRoster.rows)) {
        const empRow = priorRoster.rows[rEmp];
        if (empRow && typeof empRow === 'object') {
          for (const iso of Object.keys(empRow)) knownHistDates.add(iso);
        }
      }
    }
  }
  const histTailIsos: string[] = [];
  if (monthLen > 0 && priorRoster) {
    const firstIso = monthIsos[0];
    for (let offset = 1; offset <= 6; offset++) {
      const prevIso = addDaysIso(firstIso, -offset);
      if (knownHistDates.has(prevIso)) {
        histTailIsos.unshift(prevIso);
      } else {
        break;
      }
    }
  }
  const hasHistory = histTailIsos.length > 0;
  const histLen = histTailIsos.length;

  // 1. Check Daily Coverage
  calMeta.forEach(({ isoDate, dow, isWeekend }) => {
    const c = { M: 0, A: 0, N: 0, MID: 0, OFF: 0, H8: 0, ABS: 0, other: 0 };
    const rotC = { M: 0, A: 0, N: 0 };
    const dayWarnings: string[] = [];

    employees.forEach((emp) => {
      if (excludedSet.has(emp)) return;
      const s = normalizeShiftCode(rows[emp]?.[isoDate]);
      if (s === 'M') c.M++;
      else if (s === 'A') c.A++;
      else if (s === 'N') c.N++;
      else if (s === 'MID') c.MID++;
      else if (s === 'H8') c.H8++;
      else if (s === 'ABS') c.ABS++;
      else if (s === 'OFF') c.OFF++;
      else c.other++;

      const isNonRotating =
        FIXED_MID_EMPLOYEES.includes(emp) || FIXED_WEEKDAY_MA_EMPLOYEES.includes(emp);
      if (!isNonRotating) {
        if (s === 'M') rotC.M++;
        else if (s === 'A') rotC.A++;
        else if (s === 'N') rotC.N++;
      }
    });

    let isUnderstaffed = false;
    let isOverstaffedNight = false;

    if (rotC.M < 1) {
      if (tierConfig.allowCoverageGapMA) {
        options?.relaxations?.push({
          tier: 3,
          rule: 'COVERAGE_GAP',
          isoDate,
          note: `Rotating Morning understaffed on ${isoDate}: rotM=${rotC.M}/1`,
        });
      } else {
        dayWarnings.push(`Morning understaffed (${rotC.M}/1 minimum)`);
        isUnderstaffed = true;
      }
    }

    if (c.A < 2) {
      if (c.A === 1 && tierConfig.allowCoverageGapMA) {
        options?.relaxations?.push({
          tier: 3,
          rule: 'COVERAGE_GAP',
          isoDate,
          note: `Afternoon understaffed on ${isoDate}: totalA=${c.A}/2`,
        });
      } else {
        dayWarnings.push(`Afternoon understaffed (${c.A}/2 minimum)`);
        isUnderstaffed = true;
      }
    }

    if (c.N < 2) {
      dayWarnings.push(`Night understaffed (${c.N}/2 minimum)`);
      isUnderstaffed = true;
    } else if (c.N === 3) {
      if (tierConfig.allowNightOver2) {
        options?.relaxations?.push({
          tier: 1,
          rule: 'NIGHT_OVER_2',
          isoDate,
          note: `Night overstaffed on ${isoDate}: N=3 (Tier 0 requires 2)`,
        });
      } else {
        dayWarnings.push(`Night overstaffed (${c.N}/2 maximum)`);
        isOverstaffedNight = true;
      }
    } else if (c.N > 3) {
      dayWarnings.push(`Night overstaffed (${c.N}/2 maximum)`);
      isOverstaffedNight = true;
    }

    const coverageFeasible = isCoverageFeasibleForTier(rotC.M, c.A, c.N, tier);
    if (isUnderstaffed) understaffedDaysCount++;
    if (isOverstaffedNight) nightOverstaffedDaysCount++;
    if (dayWarnings.length > 0) {
      warnings.push(...dayWarnings.map((w) => `${isoDate}: ${w}`));
    }

    dailyCoverage[isoDate] = {
      isoDate,
      dayOfWeek: dow,
      isWeekend,
      M: c.M,
      rotM: rotC.M,
      A: c.A,
      N: c.N,
      MID: c.MID,
      OFF: c.OFF + c.H8 + c.ABS + c.other,
      ABS: c.ABS,
      totalWorking: c.M + c.A + c.N + c.MID,
      isUnderstaffed: !coverageFeasible,
      warnings: dayWarnings,
    };
  });

  // 2. Check Employee Legal Constraints, Streaks, Rest Durations & Tier-Aware Blocks
  employees.forEach((emp) => {
    const isExcluded = excludedSet.has(emp);
    const isFixedMid = FIXED_MID_EMPLOYEES.includes(emp);
    const isFixedWeekdayMA = FIXED_WEEKDAY_MA_EMPLOYEES.includes(emp);
    const isNonRotating = isFixedMid || isFixedWeekdayMA;
    const empViolations: string[] = [];
    const counts = { M: 0, A: 0, N: 0, MID: 0, OFF: 0, H8: 0, ABS: 0 };

    const monthShifts = monthIsos.map((isoDate) => normalizeShiftCode(rows[emp]?.[isoDate]));
    let fixedLeaveOnWorkdays = 0;

    monthShifts.forEach((s, idx) => {
      if (s === 'M') counts.M++;
      else if (s === 'A') counts.A++;
      else if (s === 'N') counts.N++;
      else if (s === 'MID') counts.MID++;
      else if (s === 'H8') counts.H8++;
      else if (s === 'ABS') counts.ABS++;
      else counts.OFF++;

      const meta = calMeta[idx];

      // H6: Excluded employees must be strictly OFF
      if (isExcluded) {
        if (s !== 'OFF') {
          empViolations.push(
            `Excluded employee ${emp} has non-OFF shift ${s} on ${meta.isoDate}`
          );
          fixedStaffViolationCount++;
        }
        return;
      }

      // H6: Valid shift codes
      if (!VALID_CODES.has(s)) {
        empViolations.push(
          `${emp} has invalid/unexpected shift code "${s}" on ${meta.isoDate}`
        );
        fixedStaffViolationCount++;
      }

      // H6: Fixed-staff rules & rotating MID ban
      if (isFixedMid) {
        if (s === 'ABS' || s === 'H8') {
          if (!meta.isFixedOff) fixedLeaveOnWorkdays++;
          return;
        }
        if (!VALID_CODES.has(s)) return;
        const expected = meta.isFixedOff ? 'OFF' : 'MID';
        if (s !== expected) {
          if (meta.isFixedOff) {
            empViolations.push(
              `Fixed Manager ${emp} scheduled on ${s} instead of OFF on weekend/holiday (${meta.isoDate})`
            );
          } else {
            empViolations.push(
              `Fixed Manager ${emp} scheduled on ${s} instead of MID on ${meta.isoDate}`
            );
          }
          fixedStaffViolationCount++;
        }
      } else if (isFixedWeekdayMA) {
        if (s === 'ABS' || s === 'H8') {
          if (!meta.isFixedOff) fixedLeaveOnWorkdays++;
          return;
        }
        if (!VALID_CODES.has(s)) return;
        const expected = meta.isFixedOff ? 'OFF' : meta.isoWeek % 2 === 1 ? 'M' : 'A';
        if (s !== expected) {
          if (meta.isFixedOff) {
            empViolations.push(
              `Weekday staff ${emp} scheduled on ${s} instead of OFF on weekend/holiday (${meta.isoDate})`
            );
          } else {
            empViolations.push(
              `Weekday staff ${emp} scheduled on ${s} instead of ${expected} on ${meta.isoDate} (ISO wk ${meta.isoWeek})`
            );
          }
          fixedStaffViolationCount++;
        }
      } else if (s === 'MID') {
        empViolations.push(`Rotating employee ${emp} assigned MID on ${meta.isoDate}`);
        fixedStaffViolationCount++;
      }
    });

    const totalWorkingDays = counts.M + counts.A + counts.N + counts.MID;

    if (isExcluded) {
      employeeStats[emp] = {
        employee: emp,
        M: counts.M,
        A: counts.A,
        N: counts.N,
        MID: counts.MID,
        OFF: counts.OFF,
        H8: counts.H8,
        ABS: counts.ABS,
        totalWorkingDays,
        targetWorkingDays: 0,
        workDaysDiff: 0,
        maxStreak: 0,
        violations: empViolations,
      };
      if (empViolations.length > 0) {
        warnings.push(...empViolations);
      }
      return;
    }

    // Build combined sequence (historyLast6 + month) for both rotating & fixed active staff (H7)
    const empHistory = histTailIsos.map((iso) =>
      normalizeShiftCode(priorRoster?.rows?.[emp]?.[iso])
    );
    const seq = [...empHistory, ...monthShifts];
    const totalLen = seq.length;
    const isoAt = (idx: number): string =>
      idx >= histLen ? monthIsos[idx - histLen] : `history[${idx - histLen}]`;

    // H1: Max 5 consecutive work days (A/M/N/MID count as work; OFF/ABS/H8/HUP/AD/CM break streak)
    let maxStreak = 0;
    {
      let runStart = -1;
      for (let i = 0; i <= totalLen; i++) {
        const w = i < totalLen && WORK_SHIFTS.has(seq[i]);
        if (w && runStart === -1) runStart = i;
        if (!w && runStart !== -1) {
          const runEnd = i - 1;
          const runLen = runEnd - runStart + 1;
          if (runEnd >= histLen) {
            maxStreak = Math.max(maxStreak, runLen);
            if (runLen > 5) {
              empViolations.push(
                `${emp} has ${runLen} consecutive work days (${isoAt(runStart)}..${isoAt(runEnd)}, max 5)`
              );
              streakViolationCount++;
            }
          }
          runStart = -1;
        }
      }
    }

    // H2 & H3: Post-Night rest >= 2 non-work days and no N directly followed by M/A/MID
    for (let i = 0; i < totalLen - 1; i++) {
      if (seq[i] !== 'N') continue;
      const next1 = seq[i + 1];
      if (next1 === 'N') continue;

      if ((next1 === 'M' || next1 === 'A' || next1 === 'MID') && i + 1 >= histLen) {
        empViolations.push(
          `Night shift on ${isoAt(i)} followed directly by ${next1} on ${isoAt(i + 1)} without rest`
        );
        nightRestViolationCount++;
      }

      const prematureIndices: number[] = [];
      for (const k of [1, 2]) {
        if (i + k < totalLen && WORK_SHIFTS.has(seq[i + k]) && i + k >= histLen) {
          prematureIndices.push(i + k);
        }
      }
      if (prematureIndices.length > 0) {
        const firstPremature = prematureIndices[0];
        empViolations.push(
          `Night shift on ${isoAt(i)} followed by < 2 rest days before ${seq[firstPremature]} on ${isoAt(firstPremature)} (minimum 2 rest days required)`
        );
        nightRestViolationCount++;
      }
    }

    // Extract maximal contiguous work blocks across seq
    interface WorkBlock {
      start: number;
      end: number;
      len: number;
      shifts: string[];
      primaryShift: string;
      isMixed: boolean;
    }
    const workBlocks: WorkBlock[] = [];
    {
      let i = 0;
      while (i < totalLen) {
        if (WORK_SHIFTS.has(seq[i])) {
          const start = i;
          const shiftSet = new Set<string>();
          while (i < totalLen && WORK_SHIFTS.has(seq[i])) {
            shiftSet.add(seq[i]);
            i++;
          }
          const end = i - 1;
          workBlocks.push({
            start,
            end,
            len: end - start + 1,
            shifts: Array.from(shiftSet),
            primaryShift: seq[start],
            isMixed: shiftSet.size > 1,
          });
        } else {
          i++;
        }
      }
    }

    // H4 part 1: No mixed shift types inside one block (for all active staff)
    for (const blk of workBlocks) {
      if (blk.end < histLen) continue;
      if (blk.isMixed) {
        empViolations.push(
          `Mixed shift block on ${isoAt(Math.max(histLen, blk.start))}: cannot combine [${blk.shifts.join(',')}] in single work block ${isoAt(blk.start)}..${isoAt(blk.end)}`
        );
        mixedBlockViolationCount++;
      }
    }

    // Rotating-only checks: H4 cycle order, H8 rest <= 3d, Tier 0/2 block lengths & rests
    if (!isNonRotating) {
      // H4 part 2: Forward cycle A -> M -> N -> A (with ABS/H8/CM split rules)
      for (let b = 1; b < workBlocks.length; b++) {
        const prev = workBlocks[b - 1];
        const cur = workBlocks[b];
        if (cur.start < histLen) continue;
        if (prev.isMixed || cur.isMixed) continue;
        if (!ROTATING_SHIFTS.has(prev.primaryShift) || !ROTATING_SHIFTS.has(cur.primaryShift)) {
          continue;
        }

        const gap = seq.slice(prev.end + 1, cur.start);
        const absH8Count = gap.filter((s) => LEAVE_OR_SPLIT_CODES.has(s)).length;
        const expectedNext = EXPECTED_NEXT_SHIFT[prev.primaryShift];

        if (absH8Count === 0) {
          if (cur.primaryShift !== expectedNext) {
            empViolations.push(
              `Shift sequence violation on ${isoAt(cur.start)}: expected ${expectedNext} after ${prev.primaryShift}, got ${cur.primaryShift}`
            );
            mixedBlockViolationCount++;
          }
        } else {
          const isSame = cur.primaryShift === prev.primaryShift;
          const isNext = cur.primaryShift === expectedNext;
          if (!isSame && !isNext && absH8Count < 2) {
            empViolations.push(
              `Shift sequence violation on ${isoAt(cur.start)} across single-day leave split: ${prev.primaryShift} --[${gap.join(',')}]--> ${cur.primaryShift}`
            );
            mixedBlockViolationCount++;
          }
        }
      }

      // H8: Pure-OFF rest > 3 consecutive days (rest containing ABS/H8/CM exempt)
      {
        let rStart = -1;
        for (let i = 0; i <= totalLen; i++) {
          const nonWork = i < totalLen && !WORK_SHIFTS.has(seq[i]);
          if (nonWork && rStart === -1) rStart = i;
          if (!nonWork && rStart !== -1) {
            const rEnd = i - 1;
            const rLen = rEnd - rStart + 1;
            if (rEnd >= histLen && rLen > 3) {
              const hasLeave = seq
                .slice(rStart, rEnd + 1)
                .some((s) => LEAVE_OR_SPLIT_CODES.has(s));
              if (!hasLeave) {
                empViolations.push(
                  `Rest period exceeded 3 days on ${isoAt(Math.max(histLen, rStart))} (${rLen} days: ${isoAt(rStart)}..${isoAt(rEnd)})`
                );
                excessiveRestViolationCount++;
              }
            }
            rStart = -1;
          }
        }
      }

      // Tier 0 & Tier 2: Rotating work block lengths (3..5; exempt tail, !hasHistory Day-1, adjacent to ABS/H8/CM)
      let twoDayBlocksCount = 0;
      for (const blk of workBlocks) {
        if (blk.end < histLen) continue;
        const isTrailingAtMonthEnd = blk.end === totalLen - 1;
        const isDay1NoHistory = !hasHistory && blk.start === histLen;
        const adjBefore = blk.start > 0 && LEAVE_OR_SPLIT_CODES.has(seq[blk.start - 1]);
        const adjAfter = blk.end + 1 < totalLen && LEAVE_OR_SPLIT_CODES.has(seq[blk.end + 1]);
        const isExempt = isTrailingAtMonthEnd || isDay1NoHistory || adjBefore || adjAfter;

        if (!isExempt && blk.len < 3) {
          const blockIso = isoAt(Math.max(histLen, blk.start));
          if (blk.len === 1) {
            empViolations.push(
              `Short work block (${blk.primaryShift}, 1 day) starting ${blockIso}: 1-day blocks are not allowed`
            );
            streakViolationCount++;
          } else if (blk.len === 2) {
            twoDayBlocksCount++;
            if (twoDayBlocksCount <= 1) {
              const note = `${emp} has 2-day work block (${blk.primaryShift}) at ${isoAt(blk.start)}..${isoAt(blk.end)}`;
              if (tierConfig.allowShortBlock2) {
                options?.relaxations?.push({
                  tier: 2,
                  rule: 'SHORT_BLOCK_2',
                  isoDate: blockIso,
                  employee: emp,
                  note,
                });
              } else {
                empViolations.push(
                  `Short work block (${blk.primaryShift}, 2 days) starting ${blockIso}: internal blocks require 3-5 consecutive days`
                );
                streakViolationCount++;
              }
            } else {
              empViolations.push(
                `${emp} has ${twoDayBlocksCount} two-day work blocks in month (max 1 in Tier 2) at ${isoAt(blk.start)}..${isoAt(blk.end)}`
              );
              streakViolationCount++;
            }
          }
        }
      }

      // Tier 0 & Tier 2: 1-day pure-OFF rest between non-Night blocks (SHORT_REST_1)
      for (let b = 1; b < workBlocks.length; b++) {
        const prev = workBlocks[b - 1];
        const cur = workBlocks[b];
        if (cur.start < histLen) continue;
        const gap = seq.slice(prev.end + 1, cur.start);
        if (gap.length === 1 && !LEAVE_OR_SPLIT_CODES.has(gap[0]) && prev.primaryShift !== 'N') {
          const restIso = isoAt(Math.max(histLen, prev.end + 1));
          const note = `${emp} has 1-day OFF rest on ${isoAt(prev.end + 1)} between ${prev.primaryShift} and ${cur.primaryShift}`;
          if (tierConfig.minRest <= 1) {
            options?.relaxations?.push({
              tier: 2,
              rule: 'SHORT_REST_1',
              isoDate: restIso,
              employee: emp,
              note,
            });
          } else {
            empViolations.push(note);
            nightRestViolationCount++;
          }
        }
      }
    }

    // H5: Exact monthly norm
    let effectiveTarget: number;
    if (isNonRotating) {
      effectiveTarget = Math.max(0, fixedMonthWorkdays - fixedLeaveOnWorkdays);
    } else {
      const baseTarget =
        targetWorkingDays ?? countLegalWorkingDaysInMonth(targetMonth ?? 11, targetYear ?? 2026);
      effectiveTarget = Math.max(0, baseTarget - counts.ABS - counts.H8);
    }
    const workDaysDiff = totalWorkingDays - effectiveTarget;

    if (isNonRotating) {
      if (workDaysDiff !== 0) {
        empViolations.push(
          `Fixed employee ${emp}: worked ${totalWorkingDays} days (target is ${effectiveTarget} weekdays${fixedLeaveOnWorkdays > 0 ? ` after ${fixedLeaveOnWorkdays}d ABS/H8` : ''})`
        );
        workDaysMismatchCount++;
      }
    } else if (targetWorkingDays !== undefined && workDaysDiff !== 0) {
      const leaveDays = counts.ABS + counts.H8;
      empViolations.push(
        `${emp}: worked ${totalWorkingDays} days (target is ${effectiveTarget} days${leaveDays > 0 ? ` after ${counts.ABS}d ABS / ${counts.H8}d H8` : ''})`
      );
      workDaysMismatchCount++;
    }

    employeeStats[emp] = {
      employee: emp,
      M: counts.M,
      A: counts.A,
      N: counts.N,
      MID: counts.MID,
      OFF: counts.OFF,
      H8: counts.H8,
      ABS: counts.ABS,
      totalWorkingDays,
      targetWorkingDays: effectiveTarget,
      workDaysDiff,
      maxStreak,
      violations: empViolations,
    };

    if (empViolations.length > 0) {
      warnings.push(...empViolations);
    }
  });

  const totalViolations =
    understaffedDaysCount +
    nightOverstaffedDaysCount +
    streakViolationCount +
    nightRestViolationCount +
    excessiveRestViolationCount +
    mixedBlockViolationCount +
    workDaysMismatchCount +
    fixedStaffViolationCount;

  return {
    isFullyCompliant: totalViolations === 0,
    totalViolations,
    understaffedDaysCount,
    nightOverstaffedDaysCount,
    streakViolationCount,
    nightRestViolationCount,
    excessiveRestViolationCount,
    mixedBlockViolationCount,
    workDaysMismatchCount,
    fixedStaffViolationCount,
    dailyCoverage,
    employeeStats,
    warnings,
  };
}

/**
 * Swaps shifts between two employees across one or multiple dates
 */
export function swapEmployeeShifts(
  currentRows: Record<string, Record<string, string>>,
  employeeA: string,
  employeeB: string,
  dates: string[]
): Record<string, Record<string, string>> {
  const updated = Object.fromEntries(
    Object.entries(currentRows).map(([emp, shifts]) => [emp, { ...shifts }])
  );

  dates.forEach((date) => {
    const shiftA = updated[employeeA]?.[date] ?? 'OFF';
    const shiftB = updated[employeeB]?.[date] ?? 'OFF';
    if (!updated[employeeA]) updated[employeeA] = {};
    if (!updated[employeeB]) updated[employeeB] = {};
    updated[employeeA][date] = shiftB;
    updated[employeeB][date] = shiftA;
  });

  return updated;
}

/**
 * Checks if a given month already exists with full date columns in the roster
 */
export function isHistoricalMonthInRoster(
  roster: RosterData | null | undefined,
  year: number,
  month: number
): boolean {
  if (!roster || !roster.dateColumns || roster.dateColumns.length === 0) return false;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  const monthPrefix = `${year}-${String(month + 1).padStart(2, '0')}`;
  const datesInRosterForMonth = roster.dateColumns.filter((d) => d.isoDate.startsWith(monthPrefix));
  return datesInRosterForMonth.length >= daysInMonth;
}

/**
 * Loads an existing historical month directly from the roster without regenerating
 */
export function loadHistoricalSchedule(
  roster: RosterData,
  year: number,
  month: number
): GeneratedScheduleResult {
  const monthDateCols = buildMonthDateColumns(year, month);
  const rows: Record<string, Record<string, string>> = {};
  const employees = roster.employees && roster.employees.length > 0 ? roster.employees : Object.keys(roster.rows);

  employees.forEach((emp) => {
    rows[emp] = {};
    monthDateCols.forEach((col) => {
      rows[emp][col.isoDate] = roster.rows[emp]?.[col.isoDate] ?? 'OFF';
    });
  });

  const targetWorkingDays = countLegalWorkingDaysInMonth(month, year);
  const validation = validateScheduleCompliance(
    rows,
    monthDateCols,
    roster,
    year,
    month,
    targetWorkingDays
  );

  return {
    month,
    year,
    daysCount: monthDateCols.length,
    targetWorkingDays,
    dateColumns: monthDateCols,
    rows,
    validation,
    seed: 0,
    isHistorical: true
  };
}

/**
 * Exports a roster to an Excel (.xlsx) workbook file replicating 1/1 layout of 11 Noiembrie.xlsx
 */
export async function exportRosterToExcelFile(
  roster: {
    fileName: string;
    month: number;
    year: number;
    employees: string[];
    rows: Record<string, Record<string, string>>;
    dateColumns: { index: number; isoDate: string }[];
    targetWorkingDays?: number;
  },
  customFileName?: string
) {
  const ExcelJSModule = await import('exceljs');
  const ExcelJS = (ExcelJSModule.default || ExcelJSModule) as unknown as typeof import('exceljs');
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Sheet1', {
    views: [{ showGridLines: true }]
  });

  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'thin', color: { argb: 'FFD9D9D9' } },
    bottom: { style: 'thin', color: { argb: 'FFD9D9D9' } },
    left: { style: 'thin', color: { argb: 'FFD9D9D9' } },
    right: { style: 'thin', color: { argb: 'FFD9D9D9' } }
  };

  const mediumBorder: Partial<ExcelJS.Borders> = {
    top: { style: 'medium', color: { argb: 'FF000000' } },
    bottom: { style: 'medium', color: { argb: 'FF000000' } },
    left: { style: 'medium', color: { argb: 'FF000000' } },
    right: { style: 'medium', color: { argb: 'FF000000' } }
  };

  // 1. Column Widths
  ws.getColumn(1).width = 53.33;
  const numDays = roster.dateColumns.length;
  for (let c = 2; c <= numDays + 1; c++) {
    ws.getColumn(c).width = 7.5; // width 7.5 ensures dates like 25-Dec fit without ###
  }
  const totalColIdx = numDays + 2;
  ws.getColumn(totalColIdx).width = 8.5;

  // 2. Row Heights
  ws.getRow(1).height = 18.6;
  ws.getRow(2).height = 15.6;

  // 3. Header Cells A1 and A2
  const cellA1 = ws.getCell(1, 1);
  cellA1.value = 'EQUIPMENT';
  cellA1.font = { name: 'Calibri', size: 14, bold: true, color: { argb: 'FF000000' } };
  cellA1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
  cellA1.alignment = { horizontal: 'center', vertical: 'middle' };
  cellA1.border = mediumBorder;

  const cellA2 = ws.getCell(2, 1);
  cellA2.value = 'Employee Name';
  cellA2.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  cellA2.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF70AD47' } };
  cellA2.alignment = { horizontal: 'center', vertical: 'middle' };
  cellA2.border = thinBorder;

  const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

  // 4. Header Day Columns
  roster.dateColumns.forEach((col, idx) => {
    const colIdx = 2 + idx;
    const [y, m, dNum] = col.isoDate.split('-').map(Number);
    const dateObj = new Date(Date.UTC(y, m - 1, dNum, 12, 0, 0));
    const dayOfWeek = new Date(y, m - 1, dNum).getDay();
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    const isHoliday = ROMANIAN_HOLIDAYS.has(col.isoDate);
    const isSpecial = isWeekend || isHoliday;

    // Row 1: Day of week name
    const r1Cell = ws.getCell(1, colIdx);
    r1Cell.value = dayNames[dayOfWeek];
    r1Cell.font = {
      name: 'Calibri',
      size: 12,
      bold: true,
      color: { argb: isSpecial ? 'FFFFFFFF' : 'FF000000' }
    };
    if (isSpecial) {
      r1Cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC00000' } };
    }
    r1Cell.alignment = { horizontal: 'center', vertical: 'middle' };
    r1Cell.border = thinBorder;

    // Row 2: Date formatted as d-mmm (e.g. 1-Nov, 2-Nov...)
    const r2Cell = ws.getCell(2, colIdx);
    r2Cell.value = dateObj;
    r2Cell.numFmt = 'd-mmm';
    r2Cell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FF000000' } };
    r2Cell.alignment = { horizontal: 'center', vertical: 'middle' };
    r2Cell.border = thinBorder;
  });

  // Target working days for month
  const targetWorkingDays =
    roster.targetWorkingDays ?? countLegalWorkingDaysInMonth(roster.month, roster.year);

  // 5. Employee Rows
  const employees = roster.employees && roster.employees.length > 0 ? roster.employees : Object.keys(roster.rows);
  const firstDayColLetter = ws.getColumn(2).letter;
  const lastDayColLetter = ws.getColumn(1 + numDays).letter;
  const totalColLetter = ws.getColumn(totalColIdx).letter;

  employees.forEach((emp, eIdx) => {
    const rowIdx = 3 + eIdx;
    const row = ws.getRow(rowIdx);
    row.height = 15.6;

    const isLightGreen = eIdx % 2 === 1;

    // Col A: Employee Name
    const nameCell = ws.getCell(rowIdx, 1);
    nameCell.value = emp;
    nameCell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FF000000' } };
    nameCell.alignment = { horizontal: 'center', vertical: 'middle' };
    nameCell.border = thinBorder;
    if (isLightGreen) {
      nameCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2F0D9' } };
    }

    let empWorkedDaysCount = 0;

    // Days shifts
    roster.dateColumns.forEach((col, dIdx) => {
      const colIdx = 2 + dIdx;
      const rawShift = roster.rows[emp]?.[col.isoDate] ?? 'OFF';
      const shiftCell = ws.getCell(rowIdx, colIdx);
      shiftCell.alignment = { horizontal: 'center', vertical: 'middle' };
      shiftCell.border = thinBorder;

      if (rawShift === 'OFF' || rawShift === '') {
        shiftCell.value = null;
        // In 11 Noiembrie.xlsx, empty / OFF days are filled with lime green
        shiftCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF99FF66' } };
      } else {
        empWorkedDaysCount++;
        shiftCell.value = rawShift;

        if (rawShift === 'N') {
          // Night: Bold Red font
          shiftCell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFF0000' } };
        } else if (rawShift === 'M') {
          // Morning: Bold Dark Green font
          shiftCell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FF375623' } };
        } else if (rawShift === 'A') {
          // Afternoon: Bold Dark Amber font
          shiftCell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFD09A00' } };
        } else if (rawShift === 'MID') {
          // MID: Bold Black font
          shiftCell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FF000000' } };
        } else if (rawShift === 'H8') {
          // Legal Holiday: Red background, white text
          shiftCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC00000' } };
          shiftCell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
        } else if (rawShift === 'HUP') {
          // Paid Leave: Dark green background, white text
          shiftCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF375623' } };
          shiftCell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
        } else if (rawShift === 'ABS') {
          // Absence: Solid black background, white text
          shiftCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } };
          shiftCell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
        } else if (rawShift === 'HE') {
          shiftCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0070C0' } };
          shiftCell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
        } else if (rawShift === 'CM') {
          shiftCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF9BC2E6' } };
          shiftCell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FF000000' } };
        } else {
          shiftCell.font = { name: 'Calibri', size: 12, color: { argb: 'FF000000' } };
        }
      }
    });

    // Total Col: formula =COUNTA(B{r}:AE{r})
    const totalCell = ws.getCell(rowIdx, totalColIdx);
    totalCell.value = {
      formula: `COUNTA(${firstDayColLetter}${rowIdx}:${lastDayColLetter}${rowIdx})`,
      result: empWorkedDaysCount
    };
    const isTargetMet = empWorkedDaysCount === targetWorkingDays;
    totalCell.font = {
      name: 'Calibri',
      size: 12,
      bold: true,
      color: { argb: isTargetMet ? 'FF006100' : 'FF9C0006' }
    };
    totalCell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: isTargetMet ? 'FFC6EFCE' : 'FFFFC7CE' }
    };
    totalCell.alignment = { horizontal: 'center', vertical: 'middle' };
    totalCell.border = thinBorder;
  });

  const lastEmpRow = 2 + employees.length;

  // 6. Coverage summary rows
  const summaryDefs = [
    { label: 'MORNING 06-14', code: 'M', offset: 3, targetMin: 1 },
    { label: 'AFTER 14-22', code: 'A', offset: 4, targetMin: 2 },
    { label: 'NIGHT 22-06', code: 'N', offset: 5, targetMin: 2 },
    { label: 'MID', code: 'MID', offset: 6, targetMin: 0 }
  ];

  summaryDefs.forEach(({ label, code, offset, targetMin }) => {
    const sRowIdx = lastEmpRow + offset;
    const row = ws.getRow(sRowIdx);
    row.height = 15.6;

    const labelCell = ws.getCell(sRowIdx, 1);
    labelCell.value = label;
    labelCell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF000000' } };
    labelCell.alignment = { horizontal: 'center', vertical: 'middle' };
    labelCell.border = thinBorder;

    roster.dateColumns.forEach((col, dIdx) => {
      const colIdx = 2 + dIdx;
      const colLetter = ws.getColumn(colIdx).letter;
      const countCell = ws.getCell(sRowIdx, colIdx);

      // Precalculate count for instant accurate rendering
      let count = 0;
      employees.forEach((emp) => {
        const val = roster.rows[emp]?.[col.isoDate];
        if (code === 'MID') {
          if (val === 'MID') count++;
        } else if (val && val.includes(code)) {
          count++;
        }
      });

      countCell.value = {
        formula: `COUNTIF(${colLetter}3:${colLetter}${lastEmpRow},"${code}")`,
        result: count
      };

      const [yNum, mNum, dNum] = col.isoDate.split('-').map(Number);
      const dayOfWeek = new Date(yNum, mNum - 1, dNum).getDay();
      const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
      const isHoliday = ROMANIAN_HOLIDAYS.has(col.isoDate);
      const isSpecial = isWeekend || isHoliday;
      const expectedMin = code === 'MID' ? (isSpecial ? 0 : 1) : targetMin;

      if (count === 0 && expectedMin > 0) {
        countCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFC7CE' } };
        countCell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF9C0006' } };
      } else if (count === 1 && expectedMin >= 2) {
        countCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEB9C' } };
        countCell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF9C5700' } };
      } else {
        countCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFC6EFCE' } };
        countCell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FF006100' } };
      }

      countCell.alignment = { horizontal: 'center', vertical: 'middle' };
      countCell.border = thinBorder;
    });
  });

  // 7. Dynamic conditional formatting rules
  const shiftsRange = `${firstDayColLetter}3:${lastDayColLetter}${lastEmpRow}`;
  ws.addConditionalFormatting({
    ref: shiftsRange,
    rules: [
      {
        type: 'expression',
        priority: 1,
        formulae: [`LEN(TRIM(${firstDayColLetter}3))=0`],
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF99FF66' } }
        }
      },
      {
        type: 'cellIs',
        operator: 'equal',
        priority: 2,
        formulae: ['"N"'],
        style: {
          font: { bold: true, color: { argb: 'FFFF0000' } }
        }
      },
      {
        type: 'cellIs',
        operator: 'equal',
        priority: 3,
        formulae: ['"M"'],
        style: {
          font: { bold: true, color: { argb: 'FF375623' } }
        }
      },
      {
        type: 'cellIs',
        operator: 'equal',
        priority: 4,
        formulae: ['"A"'],
        style: {
          font: { bold: true, color: { argb: 'FFD09A00' } }
        }
      },
      {
        type: 'cellIs',
        operator: 'equal',
        priority: 5,
        formulae: ['"H8"'],
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFC00000' } },
          font: { color: { argb: 'FFFFFFFF' } }
        }
      },
      {
        type: 'cellIs',
        operator: 'equal',
        priority: 6,
        formulae: ['"HUP"'],
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF375623' } },
          font: { color: { argb: 'FFFFFFFF' } }
        }
      },
      {
        type: 'cellIs',
        operator: 'equal',
        priority: 7,
        formulae: ['"ABS"'],
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF000000' } },
          font: { color: { argb: 'FFFFFFFF' } }
        }
      }
    ]
  });

  // Total Column conditional formatting
  const totalRange = `${totalColLetter}3:${totalColLetter}${lastEmpRow}`;
  ws.addConditionalFormatting({
    ref: totalRange,
    rules: [
      {
        type: 'cellIs',
        operator: 'equal',
        priority: 10,
        formulae: [String(targetWorkingDays)],
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFC6EFCE' } },
          font: { color: { argb: 'FF006100' } }
        }
      },
      {
        type: 'cellIs',
        operator: 'lessThan',
        priority: 11,
        formulae: [String(targetWorkingDays)],
        style: {
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FFFFC7CE' } },
          font: { color: { argb: 'FF9C0006' } }
        }
      }
    ]
  });

  // 8. Write to buffer and trigger browser download
  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  const fileName = customFileName || `Schedule_${roster.year}_${String(roster.month + 1).padStart(2, '0')}.xlsx`;
  a.download = fileName;
  a.click();
  URL.revokeObjectURL(url);
}

/**
 * Intelligent, minimally-invasive shift readjustment algorithm to makeup for colleagues calling in sick (ABS).
 *
 * Rules:
 * 1. Minimum daily staffing: rotating M >= 1, total A >= 2, N = 2 (Tier 0 strict; relaxed in Tiers 1-3).
 * 2. Minimum invasiveness: If staffing on a shift with an ABS colleague is still >= 2, DO NOT change anyone.
 * 3. Morning deficit (M < 2):
 *    - Priority 1: Afternoon surplus on the same day (A > 2 -> move candidate to M).
 *      Candidate must NOT have worked Night on the previous day.
 *    - Priority 2: Eligible OFF colleague on that day.
 * 4. Afternoon deficit (A < 2):
 *    - Priority 1: Morning surplus on the same day (M > 2 -> move candidate to A).
 *      Candidate should not work Morning on the following day (to avoid short 8h turnaround).
 *    - Priority 2: Eligible OFF colleague on that day.
 * 5. Night deficit (N < 2):
 *    - Priority 1: Colleague currently on Afternoon (A) on the last day of their Afternoon block (next day is OFF).
 *      Convert A -> N. Night shift requires rest on subsequent day(s).
 *    - STRICT PROHIBITION: NEVER move someone on Morning shifts to Night (M -> N).
 *    - Priority 2: Colleague scheduled OFF who is about to start Night or has sufficient rest.
 *    - Priority 3: Other eligible Afternoon colleague if Afternoon has surplus.
 * 6. Non-rotating employees (Palade Alexandru-Ionut, Tir George Cristian) cannot be assigned to Night.
 * 7. Multi-variant support: variantSeed rotates candidate selection when multiple valid options exist.
 */
export function calculateAbsReadjustments(
  rows: Record<string, Record<string, string>>,
  dateColumns: { index: number; date: Date; isoDate: string }[],
  priorRoster: RosterData | null,
  targetYear: number,
  targetMonth: number,
  targetWorkingDays: number,
  variantSeed: number = 0
): ReadjustmentProposal {
  // Deep clone helper
  const cloneRows = (src: Record<string, Record<string, string>>) => {
    const c: Record<string, Record<string, string>> = {};
    for (const k of Object.keys(src)) {
      c[k] = { ...src[k] };
    }
    return c;
  };

  const workingRows = cloneRows(rows);
  const allEmployees = Object.keys(workingRows);
  const rotatingEmployees = allEmployees.filter(
    (e) => !FIXED_MID_EMPLOYEES.includes(e) && !FIXED_WEEKDAY_MA_EMPLOYEES.includes(e)
  );

  const adjustments: ReadjustmentAdjustment[] = [];
  const unresolvedDeficits: Array<{
    isoDate: string;
    shift: 'M' | 'A' | 'N';
    needed: number;
    available: number;
    reason: string;
  }> = [];

  const getShift = (src: Record<string, Record<string, string>>, emp: string, iso: string | null): string => {
    if (!iso) return 'OFF';
    return shiftKey(src[emp]?.[iso] ?? 'OFF');
  };

  // Penalty evaluation function using validateScheduleCompliance (strict Tier 0 in Sections 1-4)
  const evaluatePenalty = (testRows: Record<string, Record<string, string>>) => {
    const rep = validateScheduleCompliance(
      testRows,
      dateColumns,
      priorRoster,
      targetYear,
      targetMonth,
      targetWorkingDays,
      [],
      { tier: 0 }
    );
    const penalty =
      rep.totalViolations * 1000 +
      rep.understaffedDaysCount * 500 +
      rep.nightRestViolationCount * 400 +
      rep.streakViolationCount * 300 +
      rep.workDaysMismatchCount * 150 +
      rep.excessiveRestViolationCount * 50 +
      rep.mixedBlockViolationCount * 50;
    return { penalty, rep };
  };

  // Helper to find an offsetting day off for an employee called in from rest
  const findOffsetDayForEmp = (
    baseRows: Record<string, Record<string, string>>,
    emp: string,
    excludeIsoDate: string,
    pendingChanges: Array<{ isoDate: string; employee?: string; toShift: string }> = []
  ): { offsetIso: string; oldShift: string } | null => {
    const preRows = cloneRows(baseRows);
    for (const pc of pendingChanges) {
      const targetEmp = pc.employee ?? emp;
      if (!preRows[targetEmp]) preRows[targetEmp] = {};
      preRows[targetEmp][pc.isoDate] = pc.toShift;
    }

    let bestCandidate: { offsetIso: string; oldShift: string; violations: number; penalty: number } | null = null;
    for (let k = 0; k < dateColumns.length; k++) {
      const kIso = dateColumns[k].isoDate;
      if (kIso === excludeIsoDate) continue;
      if (pendingChanges.some((pc) => (pc.employee ?? emp) === emp && pc.isoDate === kIso)) continue;
      const curS = shiftKey(preRows[emp]?.[kIso] ?? 'OFF');
      if (curS !== 'M' && curS !== 'A') continue;

      let count = 0;
      if (curS === 'M') {
        for (const re of rotatingEmployees) {
          if (shiftKey(preRows[re]?.[kIso] ?? 'OFF') === 'M') count++;
        }
      } else {
        for (const ae of allEmployees) {
          if (shiftKey(preRows[ae]?.[kIso] ?? 'OFF') === 'A') count++;
        }
      }
      const minStaff = curS === 'M' ? 2 : 3;
      if (count >= minStaff) {
        const tempRows = cloneRows(preRows);
        tempRows[emp][kIso] = 'OFF';
        const { penalty, rep } = evaluatePenalty(tempRows);
        if (!rep.warnings.some((w) => w.startsWith(`${kIso}:`))) {
          if (
            !bestCandidate ||
            rep.totalViolations < bestCandidate.violations ||
            (rep.totalViolations === bestCandidate.violations && penalty < bestCandidate.penalty)
          ) {
            bestCandidate = {
              offsetIso: kIso,
              oldShift: curS,
              violations: rep.totalViolations,
              penalty
            };
          }
        }
      }
    }
    return bestCandidate ? { offsetIso: bestCandidate.offsetIso, oldShift: bestCandidate.oldShift } : null;
  };

  interface CandidatePlan {
    changes: Array<{
      isoDate: string;
      employee: string;
      fromShift: string;
      toShift: string;
      reason: string;
      absColleague: string;
    }>;
    penalty: number;
    violations: number;
  }

  const absBlockDayIndex: Record<string, number> = {};

  // 1. Process day by day in chronological order
  for (let d = 0; d < dateColumns.length; d++) {
    const col = dateColumns[d];
    const { isoDate } = col;
    const prevCol = d > 0 ? dateColumns[d - 1] : null;
    const nextCol = d < dateColumns.length - 1 ? dateColumns[d + 1] : null;
    const nextNextCol = d < dateColumns.length - 2 ? dateColumns[d + 2] : null;

    const prevIso = prevCol?.isoDate ?? null;
    const nextIso = nextCol?.isoDate ?? null;
    const nextNextIso = nextNextCol?.isoDate ?? null;

    const absColleagues = allEmployees.filter(
      (emp) => (workingRows[emp]?.[isoDate] ?? '').toUpperCase() === 'ABS'
    );
    const absNames = absColleagues.join(', ') || 'Colleague on sick leave';

    // Track absence block day indices
    for (const abs of absColleagues) {
      const isContinuation = prevIso && (workingRows[abs]?.[prevIso] ?? '').toUpperCase() === 'ABS';
      if (!isContinuation) {
        absBlockDayIndex[abs] = 0;
      } else {
        absBlockDayIndex[abs] = (absBlockDayIndex[abs] ?? 0) + 1;
      }
    }

    // Total staff counts on duty (including George Tir on weekdays)
    const totalM = allEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'M').length;
    const totalA = allEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'A').length;
    const totalN = allEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'N').length;

    // Staffing checks
    let mStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'M');
    let aStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'A');
    let nStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'N');
    const offStaff = rotatingEmployees.filter((emp) => {
      const s = getShift(workingRows, emp, isoDate);
      return s === 'OFF' || s === 'H8' || s === 'AD';
    });

    const neededN = Math.max(0, 2 - totalN);
    const neededM = Math.max(0, 1 - mStaff.length);
    const neededA = Math.max(0, 2 - totalA);

    // -------------------------------------------------------------------------
    // 1. NIGHT DEFICIT (N < 2)
    // -------------------------------------------------------------------------
    for (let i = 0; i < neededN; i++) {
      const plans: CandidatePlan[] = [];

      // Category 1: Afternoon staff on last day of block (next day is OFF)
      for (const emp of aStaff) {
        const nextS = getShift(workingRows, emp, nextIso);
        const isLastDay = nextS === 'OFF' || nextS === 'H8' || nextIso === null;
        if (!isLastDay) continue;

        const nextNextS = getShift(workingRows, emp, nextNextIso);
        const hasNaturalTwoDaysRest = nextNextS === 'OFF' || nextNextS === 'H8' || nextNextIso === null;

        // Plan 1A: Simple transfer A -> N
        const changes1: CandidatePlan['changes'] = [
          {
            isoDate,
            employee: emp,
            fromShift: 'A',
            toShift: 'N',
            reason: hasNaturalTwoDaysRest
              ? `Converted last day of Afternoon block to Night (A ➔ N). Natural 2-day post-night rest already scheduled (0 extra rest days needed).`
              : `Converted last day of Afternoon block to Night (A ➔ N) to cover sick colleague.`,
            absColleague: absNames
          }
        ];

        // If nextNext day is working, set it to OFF for mandatory 2-day night rest
        if (!hasNaturalTwoDaysRest && nextNextIso) {
          changes1.push({
            isoDate: nextNextIso,
            employee: emp,
            fromShift: nextNextS,
            toShift: 'OFF',
            reason: `Mandatory post-night rest day 2 granted after covering Night shift.`,
            absColleague: absNames
          });
        }

        // Domino backfill: If Afternoon drops below 2, attempt to backfill from Morning surplus (rotating M > 1)
        if (totalA <= 2 && mStaff.length > 1) {
          const mSurplusEmp = mStaff.find((me) => getShift(workingRows, me, nextIso) !== 'M');
          if (mSurplusEmp) {
            changes1.push({
              isoDate,
              employee: mSurplusEmp,
              fromShift: 'M',
              toShift: 'A',
              reason: `Domino rebalance: Moved from Morning surplus (${totalM} ➔ ${totalM - 1}) to Afternoon to maintain 2 Afternoon staff.`,
              absColleague: absNames
            });
          }
        }

        const test1 = cloneRows(workingRows);
        for (const c of changes1) {
          test1[c.employee][c.isoDate] = c.toShift;
        }
        const { penalty: pen1, rep: rep1 } = evaluatePenalty(test1);
        plans.push({ changes: changes1, penalty: pen1, violations: rep1.totalViolations });
      }

      // Category 2: Afternoon staff with surplus (totalA > 2)
      if (totalA > 2) {
        for (const emp of aStaff) {
          const changes2: CandidatePlan['changes'] = [
            {
              isoDate,
              employee: emp,
              fromShift: 'A',
              toShift: 'N',
              reason: `Moved from Afternoon surplus (${totalA} ➔ ${totalA - 1}) to Night (A ➔ N) to cover sick colleague.`,
              absColleague: absNames
            }
          ];
          if (nextIso && getShift(workingRows, emp, nextIso) !== 'OFF') {
            changes2.push({
              isoDate: nextIso,
              employee: emp,
              fromShift: getShift(workingRows, emp, nextIso),
              toShift: 'OFF',
              reason: `Mandatory post-night rest day 1.`,
              absColleague: absNames
            });
          }
          if (nextNextIso && getShift(workingRows, emp, nextNextIso) !== 'OFF') {
            changes2.push({
              isoDate: nextNextIso,
              employee: emp,
              fromShift: getShift(workingRows, emp, nextNextIso),
              toShift: 'OFF',
              reason: `Mandatory post-night rest day 2.`,
              absColleague: absNames
            });
          }

          const test2 = cloneRows(workingRows);
          for (const c of changes2) {
            test2[c.employee][c.isoDate] = c.toShift;
          }
          const { penalty: pen2, rep: rep2 } = evaluatePenalty(test2);
          plans.push({ changes: changes2, penalty: pen2, violations: rep2.totalViolations });
        }
      }

      // Category 3: OFF staff who didn't work Night on prev day
      const eligibleOff = offStaff.filter((emp) => getShift(workingRows, emp, prevIso) !== 'N');
      for (const emp of eligibleOff) {
        const changes3: CandidatePlan['changes'] = [
          {
            isoDate,
            employee: emp,
            fromShift: 'OFF',
            toShift: 'N',
            reason: `Assigned from rest to Night shift (OFF ➔ N) to cover sick colleague.`,
            absColleague: absNames
          }
        ];
        let clearedWorkingDays = 0;
        if (nextIso && getShift(workingRows, emp, nextIso) !== 'OFF') {
          const nextOrig = getShift(workingRows, emp, nextIso);
          if (nextOrig === 'M' || nextOrig === 'A' || nextOrig === 'N' || nextOrig === 'MID') {
            clearedWorkingDays++;
          }
          changes3.push({
            isoDate: nextIso,
            employee: emp,
            fromShift: nextOrig,
            toShift: 'OFF',
            reason: `Mandatory post-night rest day 1.`,
            absColleague: absNames
          });
        }
        if (nextNextIso && getShift(workingRows, emp, nextNextIso) !== 'OFF') {
          const nextNextOrig = getShift(workingRows, emp, nextNextIso);
          if (nextNextOrig === 'M' || nextNextOrig === 'A' || nextNextOrig === 'N' || nextNextOrig === 'MID') {
            clearedWorkingDays++;
          }
          changes3.push({
            isoDate: nextNextIso,
            employee: emp,
            fromShift: nextNextOrig,
            toShift: 'OFF',
            reason: `Mandatory post-night rest day 2.`,
            absColleague: absNames
          });
        }
        const offset = clearedWorkingDays === 0 ? findOffsetDayForEmp(workingRows, emp, isoDate, changes3) : null;
        if (offset) {
          changes3.push({
            isoDate: offset.offsetIso,
            employee: emp,
            fromShift: offset.oldShift,
            toShift: 'OFF',
            reason: `Offsetting day off to preserve legal ${targetWorkingDays}-day monthly norm (balances extra Night shift on ${isoDate}).`,
            absColleague: absNames
          });
        }

        const test3 = cloneRows(workingRows);
        for (const c of changes3) {
          test3[c.employee][c.isoDate] = c.toShift;
        }
        const { penalty: pen3, rep: rep3 } = evaluatePenalty(test3);
        plans.push({ changes: changes3, penalty: pen3, violations: rep3.totalViolations });
      }

      // Choose plan that minimizes penalties & warnings
      if (plans.length > 0) {
        plans.sort((a, b) => a.penalty - b.penalty);
        const bestPenalty = plans[0].penalty;
        const topPlans = plans.filter((p) => p.penalty === bestPenalty);
        const chosen = topPlans[(variantSeed + d + i) % topPlans.length];

        for (const c of chosen.changes) {
          workingRows[c.employee][c.isoDate] = c.toShift;
          adjustments.push(c);
        }

        mStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'M');
        aStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'A');
        nStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'N');
      } else {
        unresolvedDeficits.push({
          isoDate,
          shift: 'N',
          needed: 1,
          available: nStaff.length,
          reason: `No eligible colleague found for Night shift without violating mandatory rest (M ➔ N is strictly prohibited).`
        });
      }
    }

    // -------------------------------------------------------------------------
    // 2. MORNING DEFICIT (rotM < 1)
    // -------------------------------------------------------------------------
    for (let i = 0; i < neededM; i++) {
      const plans: CandidatePlan[] = [];
      const primaryAbs = absColleagues[0] || 'colleague';
      const dayInBlock = absBlockDayIndex[primaryAbs] ?? 0;

      // Category 1: Afternoon staff with surplus (totalA > 2)
      if (totalA > 2) {
        const candidates = aStaff.filter((emp) => getShift(workingRows, emp, prevIso) !== 'N');
        // Sort so Prisecaru is first, then Olteanu (Lucian), then others
        candidates.sort((a, b) => {
          if (a.includes('Prisecaru')) return -1;
          if (b.includes('Prisecaru')) return 1;
          if (a.includes('Olteanu')) return -1;
          if (b.includes('Olteanu')) return 1;
          return a.localeCompare(b);
        });

        for (const emp of candidates) {
          const changes: CandidatePlan['changes'] = [
            {
              isoDate,
              employee: emp,
              fromShift: 'A',
              toShift: 'M',
              reason: `Moved from Afternoon surplus (${totalA} ➔ ${totalA - 1}) to Morning (A ➔ M) to cover sick colleague (0 norm change).`,
              absColleague: absNames
            }
          ];
          const test = cloneRows(workingRows);
          test[emp][isoDate] = 'M';
          const { penalty, rep } = evaluatePenalty(test);
          plans.push({ changes, penalty, violations: rep.totalViolations });
        }
      }

      // Category 2: OFF staff
      const eligibleOff = offStaff.filter((emp) => getShift(workingRows, emp, prevIso) !== 'N');
      eligibleOff.sort((a, b) => {
        if (a.includes('Urzica')) return -1;
        if (b.includes('Urzica')) return 1;
        return a.localeCompare(b);
      });

      for (const emp of eligibleOff) {
        const preChanges: CandidatePlan['changes'] = [
          {
            isoDate,
            employee: emp,
            fromShift: 'OFF',
            toShift: 'M',
            reason: `Assigned from rest to Morning shift (OFF ➔ M) to ensure minimum of 2 Morning staff.`,
            absColleague: absNames
          }
        ];
        const offset = findOffsetDayForEmp(workingRows, emp, isoDate, preChanges);
        const changes = [...preChanges];
        if (offset) {
          changes.push({
            isoDate: offset.offsetIso,
            employee: emp,
            fromShift: offset.oldShift,
            toShift: 'OFF',
            reason: `Offsetting day off granted to preserve legal ${targetWorkingDays}-day monthly norm (balances extra Morning shift on ${isoDate}).`,
            absColleague: absNames
          });
        }
        const test = cloneRows(workingRows);
        for (const c of changes) {
          test[c.employee][c.isoDate] = c.toShift;
        }
        const { penalty, rep } = evaluatePenalty(test);
        plans.push({ changes, penalty, violations: rep.totalViolations });
      }

      if (plans.length > 0) {
        plans.sort((a, b) => a.violations - b.violations || a.penalty - b.penalty);
        const minViolations = plans[0].violations;
        const minViolPlans = plans.filter((p) => p.violations === minViolations);
        let chosen: CandidatePlan;

        const prisecaruPlan = minViolPlans.find((p) => p.changes.some((c) => c.employee.includes('Prisecaru')));
        const lucianPlan = minViolPlans.find((p) => p.changes.some((c) => c.employee.includes('Olteanu')));
        const urzicaPlan = minViolPlans.find((p) => p.changes.some((c) => c.employee.includes('Urzica')));

        // Strategic variant assignment (only among minimum-violation plans):
        if (variantSeed === 0) {
          // Variant 1: Prefer Prisecaru when equally compliant
          chosen = prisecaruPlan ?? lucianPlan ?? minViolPlans[0];
        } else if (variantSeed === 1) {
          // Variant 2: Urzica 1 morning and Lucian 1 morning when equally compliant
          if (dayInBlock === 0) {
            chosen = urzicaPlan ?? prisecaruPlan ?? minViolPlans[0];
          } else {
            chosen = lucianPlan ?? prisecaruPlan ?? minViolPlans[0];
          }
        } else if (variantSeed === 2) {
          // Variant 3: Prefer Lucian when equally compliant
          chosen = lucianPlan ?? prisecaruPlan ?? minViolPlans[0];
        } else {
          // Generic cycling across top variants
          const bestPenalty = minViolPlans[0].penalty;
          const topPlans = minViolPlans.filter((p) => p.penalty === bestPenalty);
          chosen = topPlans[(variantSeed + dayInBlock + i) % topPlans.length];
        }

        for (const c of chosen.changes) {
          workingRows[c.employee][c.isoDate] = c.toShift;
          adjustments.push(c);
        }

        mStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'M');
        aStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'A');
        nStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'N');
      } else {
        unresolvedDeficits.push({
          isoDate,
          shift: 'M',
          needed: 1,
          available: mStaff.length,
          reason: `No eligible colleague from Afternoon surplus or scheduled rest could cover Morning.`
        });
      }
    }

    // -------------------------------------------------------------------------
    // 3. AFTERNOON DEFICIT (A < 2)
    // -------------------------------------------------------------------------
    for (let i = 0; i < neededA; i++) {
      const plans: CandidatePlan[] = [];

      // Category 1: Morning staff with surplus (totalM > 2)
      if (totalM > 2) {
        const candidates = mStaff.filter((emp) => getShift(workingRows, emp, nextIso) !== 'M');
        for (const emp of candidates) {
          const changes: CandidatePlan['changes'] = [
            {
              isoDate,
              employee: emp,
              fromShift: 'M',
              toShift: 'A',
              reason: `Moved from Morning surplus (${totalM} ➔ ${totalM - 1}) to Afternoon (M ➔ A) to cover sick colleague (0 norm change).`,
              absColleague: absNames
            }
          ];
          const test = cloneRows(workingRows);
          test[emp][isoDate] = 'A';
          const { penalty, rep } = evaluatePenalty(test);
          plans.push({ changes, penalty, violations: rep.totalViolations });
        }
      }

      // Category 2: OFF staff (keep fixed weekday M/A parity intact)
      const eligibleOff = offStaff.filter((emp) => getShift(workingRows, emp, prevIso) !== 'N');
      for (const emp of eligibleOff) {
        const preChanges: CandidatePlan['changes'] = [
          {
            isoDate,
            employee: emp,
            fromShift: 'OFF',
            toShift: 'A',
            reason: `Assigned from rest to Afternoon shift (OFF ➔ A) to maintain minimum of 2 Afternoon staff.`,
            absColleague: absNames
          }
        ];
        const offset = findOffsetDayForEmp(workingRows, emp, isoDate, preChanges);
        const changes = [...preChanges];
        if (offset) {
          changes.push({
            isoDate: offset.offsetIso,
            employee: emp,
            fromShift: offset.oldShift,
            toShift: 'OFF',
            reason: `Offsetting day off granted to preserve legal ${targetWorkingDays}-day monthly norm (balances extra Afternoon shift on ${isoDate}).`,
            absColleague: absNames
          });
        }
        const test = cloneRows(workingRows);
        for (const c of changes) {
          test[c.employee][c.isoDate] = c.toShift;
        }
        const { penalty, rep } = evaluatePenalty(test);
        plans.push({ changes, penalty, violations: rep.totalViolations });
      }

      if (plans.length > 0) {
        plans.sort((a, b) => a.penalty - b.penalty);
        const bestPenalty = plans[0].penalty;
        const topPlans = plans.filter((p) => p.penalty === bestPenalty);
        const chosen = topPlans[(variantSeed + d + i) % topPlans.length];

        for (const c of chosen.changes) {
          workingRows[c.employee][c.isoDate] = c.toShift;
          adjustments.push(c);
        }

        mStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'M');
        aStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'A');
        nStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'N');
      } else {
        unresolvedDeficits.push({
          isoDate,
          shift: 'A',
          needed: 1,
          available: aStaff.length,
          reason: `No eligible colleague from Morning surplus or scheduled rest could cover Afternoon.`
        });
      }
    }
  }

  // ---------------------------------------------------------------------------
  // 4. GLOBAL REBALANCING POST-PASS (Eliminate any remaining warnings at Tier 0)
  // ---------------------------------------------------------------------------
  let finalTier: ComplianceTier = 0;
  let finalRelaxations: RelaxationRecord[] = [];
  let currentValidation = validateScheduleCompliance(
    workingRows,
    dateColumns,
    priorRoster,
    targetYear,
    targetMonth,
    targetWorkingDays,
    [],
    { tier: 0 }
  );

  let improved = true;
  let maxPasses = 5;
  while (improved && maxPasses > 0 && currentValidation.totalViolations > 0) {
    improved = false;
    maxPasses--;

    for (const emp of rotatingEmployees) {
      const stats = currentValidation.employeeStats[emp];
      if (!stats) continue;

      // Case A: Employee worked more days than target (e.g. called in from OFF)
      if (stats.totalWorkingDays > stats.targetWorkingDays) {
        let bestOffset: { kIso: string; curS: string; rep: typeof currentValidation } | null = null;
        for (let k = 0; k < dateColumns.length; k++) {
          const kIso = dateColumns[k].isoDate;
          const curS = shiftKey(workingRows[emp]?.[kIso] ?? 'OFF');
          if (curS !== 'M' && curS !== 'A') continue;

          // Morning requires >= 1 rotating M after removal; Afternoon requires >= 2 total A (incl. weekday M/A) after removal
          const curCount = curS === 'M'
            ? rotatingEmployees.filter((re) => shiftKey(workingRows[re]?.[kIso] ?? 'OFF') === 'M').length
            : allEmployees.filter((ae) => shiftKey(workingRows[ae]?.[kIso] ?? 'OFF') === 'A').length;

          if (curCount >= (curS === 'M' ? 2 : 3)) {
            const testRows = cloneRows(workingRows);
            testRows[emp][kIso] = 'OFF';
            const testRep = validateScheduleCompliance(
              testRows,
              dateColumns,
              priorRoster,
              targetYear,
              targetMonth,
              targetWorkingDays,
              [],
              { tier: 0 }
            );

            if (testRep.totalViolations < currentValidation.totalViolations) {
              if (!bestOffset || testRep.totalViolations < bestOffset.rep.totalViolations) {
                bestOffset = { kIso, curS, rep: testRep };
              }
            }
          }
        }
        if (bestOffset) {
          workingRows[emp][bestOffset.kIso] = 'OFF';
          currentValidation = bestOffset.rep;
          adjustments.push({
            isoDate: bestOffset.kIso,
            employee: emp,
            fromShift: bestOffset.curS,
            toShift: 'OFF',
            reason: `Rebalanced: Added offsetting day off to preserve legal monthly norm of ${targetWorkingDays} working days.`,
            absColleague: 'Rebalancing'
          });
          improved = true;
        }
      }
    }
  }

  // ---------------------------------------------------------------------------
  // 5. EXACT MINIMAL-DIFF DP+SA FALLBACK REPAIR ACROSS TIERS [0, 1, 2, 3]
  // ---------------------------------------------------------------------------
  if (currentValidation.totalViolations > 0 && rotatingEmployees.length > 0) {
    const s5StartMs = performance.now();
    const s5BudgetMs = 3500;
    const s5DeadlineMs = s5StartMs + s5BudgetMs - 100;
    const absTiers: ComplianceTier[] = [0, 1, 2, 3];
    const daysCount = dateColumns.length;
    const GEN_NEXT = DAG_NEXT_SHIFT_ACT;
    const shToGen = (s: string) => (s === 'A' ? 0 : s === 'M' ? 1 : s === 'N' ? 2 : s === 'ABS' ? 4 : (s === 'H8' || s === 'HUP') ? 5 : 3);
    const fixedMidSet = new Set(FIXED_MID_EMPLOYEES);
    const fixedMaEmps = allEmployees.filter((e) => FIXED_WEEKDAY_MA_EMPLOYEES.includes(e));

    let firstGlobalAbsIdx = daysCount;
    const empFirstAbsIndices: number[] = [];
    for (const emp of rotatingEmployees) {
      let eFirst = daysCount;
      for (let d = 0; d < daysCount; d++) {
        if (shiftKey(rows[emp]?.[dateColumns[d].isoDate] ?? 'OFF') === 'ABS') {
          eFirst = d;
          if (d < firstGlobalAbsIdx) firstGlobalAbsIdx = d;
          break;
        }
      }
      if (eFirst < daysCount) empFirstAbsIndices.push(eFirst);
    }
    for (const emp of fixedMaEmps) {
      for (let d = 0; d < daysCount; d++) {
        if (shiftKey(rows[emp]?.[dateColumns[d].isoDate] ?? 'OFF') === 'ABS') {
          if (d < firstGlobalAbsIdx) firstGlobalAbsIdx = d;
          break;
        }
      }
    }
    if (firstGlobalAbsIdx === daysCount) firstGlobalAbsIdx = 0;
    const hasDistinctEmpAbsStarts = empFirstAbsIndices.some((idx) => idx > firstGlobalAbsIdx);

    // Extract raw 6-day contiguous history from priorRoster preserving ABS/H8/CM/OFF/work sequence
    const empHistories: Record<string, string[]> = {};
    for (const emp of rotatingEmployees) {
      empHistories[emp] = extractRawHistoryTail(priorRoster, emp, targetYear, targetMonth, 6);
    }

    interface GenEdge {
      act: number; // 0:A, 1:M, 2:N, 3:OFF(or Tir M/OFF/ABS/H8), 4:ABS, 5:H8/HUP
      nextNode: number;
      penA: number;
      penB: number;
    }
    interface GenDag {
      emp: string;
      isFixedTir: boolean;
      origCodes: string[];
      origActs: number[];
      layers: GenEdge[][][];
      feasible: boolean;
      maxNodes: number;
      cAct: Int8Array[];
      cNext: Int32Array[];
    }

    const buildTirGenDag = (emp: string): GenDag => {
      const origCodes = dateColumns.map(({ isoDate }) => shiftKey(rows[emp]?.[isoDate] ?? 'OFF'));
      const origActs = origCodes.map((c) => (c === 'A' ? 0 : 3));
      const layers: GenEdge[][][] = new Array(daysCount);
      for (let d = 0; d < daysCount; d++) {
        // Lock Tir strictly to origActs[d] to preserve H6 ISO-week M/A parity and leave/off days
        layers[d] = [[{ act: origActs[d], nextNode: 0, penA: 0, penB: 0 }]];
      }
      return {
        emp,
        isFixedTir: true,
        origCodes,
        origActs,
        layers,
        feasible: true,
        maxNodes: 1,
        cAct: Array.from({ length: daysCount }, () => new Int8Array(1)),
        cNext: Array.from({ length: daysCount }, () => new Int32Array(1))
      };
    };

    const LEAVE_CODES = DAG_LEAVE_CODES;

    const buildGeneralDag = (
      emp: string,
      empIdx: number,
      minNonNightRest: number,
      allowShortBlock2: boolean,
      lockMode: 'globalPre' | 'empPre' | 'none'
    ): GenDag => {
      const origCodes = dateColumns.map(({ isoDate }) => shiftKey(rows[emp]?.[isoDate] ?? 'OFF'));
      const origActs = origCodes.map((c) => shToGen(c));
      let empFirstAbsIdx = Infinity;
      let fixedNonWorkDrops = 0;
      let empHasAbsOrH8 = false;
      for (let d = 0; d < daysCount; d++) {
        if (origActs[d] === 4) {
          if (d < empFirstAbsIdx) empFirstAbsIdx = d;
          fixedNonWorkDrops++;
          empHasAbsOrH8 = true;
        } else if (origActs[d] === 5) {
          fixedNonWorkDrops++;
          empHasAbsOrH8 = true;
        }
      }
      const effectiveTarget = Math.max(0, targetWorkingDays - fixedNonWorkDrops);
      const editableDays = origActs.filter((a) => a !== 4 && a !== 5).length;
      if (effectiveTarget > editableDays) {
        return { emp, isFixedTir: false, origCodes, origActs, layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
      }

      const dayAllowed: number[][] = new Array(daysCount);
      for (let d = 0; d < daysCount; d++) {
        const oa = origActs[d];
        if (oa === 4 || oa === 5) {
          dayAllowed[d] = [oa];
        } else if (lockMode === 'globalPre' && d < firstGlobalAbsIdx) {
          dayAllowed[d] = [oa];
        } else if (lockMode === 'empPre' && empFirstAbsIdx < Infinity && d < empFirstAbsIdx) {
          dayAllowed[d] = [oa];
        } else {
          dayAllowed[d] = [0, 1, 2, 3];
        }
      }

      const edgePens = (d: number, act: number, restPen: number): { penA: number; penB: number } => {
        if (act === origActs[d]) return { penA: restPen, penB: restPen * 2 };
        const tie = variantSeed > 0 ? (((empIdx * 31 + d * 17 + act * 7 + variantSeed * 13) % 7) * 0.1) : 0;
        if (d < empFirstAbsIdx && empFirstAbsIdx < Infinity) {
          return { penA: restPen + 120 + tie, penB: restPen * 2 + 1200 + tie };
        }
        if (d < firstGlobalAbsIdx) {
          return { penA: restPen + 40 + tie, penB: restPen * 2 + 400 + tie };
        }
        if (empHasAbsOrH8) {
          return { penA: restPen + 8 + tie, penB: restPen * 2 + 120 + tie };
        }
        return { penA: restPen + 3 + tie, penB: restPen * 2 + 45 + tie };
      };

      const numStates = (effectiveTarget + 1) * 4 * 13 * 4 * 2;
      const enc = encodeGenDagState;

      const rawHistory = empHistories[emp] ?? [];
      const history = rawHistory.map((s) =>
        s === 'HUP' || s === 'H8' || s === 'ABS' || s === 'AD' || s === 'CM' ? 'OFF' : s
      );
      const activeTailWork = getActiveTailBlock(history);
      const activeTailRest = activeTailWork.length === 0 ? getActiveTailRest(history) : [];
      const idxBefore = rawHistory.length - 1 - activeTailWork.length;
      const adjLeaveBefore = idxBefore >= 0 && LEAVE_CODES.has(rawHistory[idxBefore]);
      let lastHistWork: string | null = null;
      for (let i = history.length - 1; i >= 0; i--) {
        if (history[i] === 'A' || history[i] === 'M' || history[i] === 'N') {
          lastHistWork = history[i];
          break;
        }
      }

      const allowed0 = dayAllowed[0];
      const startEdges: Array<{ act: number; nextState: number; penA: number; penB: number }> = [];
      if (activeTailWork.length > 0) {
        const curShift = activeTailWork[activeTailWork.length - 1];
        const curIdx = shToGen(curShift);
        const totalWorkTail = activeTailWork.length;
        let sameLen = 0;
        for (let i = activeTailWork.length - 1; i >= 0 && activeTailWork[i] === curShift; i--) sameLen++;
        if (curIdx <= 2) {
          if (totalWorkTail < 5 && 1 <= effectiveTarget && allowed0.includes(curIdx)) {
            const p = edgePens(0, curIdx, 0);
            const nextMl = adjLeaveBefore ? 5 + totalWorkTail : totalWorkTail;
            startEdges.push({ act: curIdx, nextState: enc(1, curIdx, nextMl, 0, 0), ...p });
          }
          if (sameLen >= 3 || adjLeaveBefore) {
            if (allowed0.includes(3)) {
              startEdges.push({ act: 3, nextState: enc(0, curIdx, 10, 1, 0), ...edgePens(0, 3, 0) });
            }
          }
          if (allowed0.includes(4)) {
            startEdges.push({ act: 4, nextState: enc(0, curIdx, 10, 0, 0), ...edgePens(0, 4, 0) });
          }
          if (allowed0.includes(5)) {
            startEdges.push({ act: 5, nextState: enc(0, curIdx, 10, 0, 0), ...edgePens(0, 5, 0) });
          }
        }
      } else if (history.length > 0) {
        const restLen = activeTailRest.length;
        const lastIdx = lastHistWork ? shToGen(lastHistWork) : 3;
        const initMl = Math.min(12, 10 + restLen);
        if (restLen < 3 && allowed0.includes(3)) {
          startEdges.push({ act: 3, nextState: enc(0, lastIdx, initMl, restLen + 1, 0), ...edgePens(0, 3, 0) });
        }
        if (allowed0.includes(4)) {
          startEdges.push({ act: 4, nextState: enc(0, lastIdx, initMl, 0, 0), ...edgePens(0, 4, 0) });
        }
        if (allowed0.includes(5)) {
          startEdges.push({ act: 5, nextState: enc(0, lastIdx, initMl, 0, 0), ...edgePens(0, 5, 0) });
        }
        const reqRest = lastIdx === 2 ? 2 : (adjLeaveBefore ? 1 : minNonNightRest);
        if (1 <= effectiveTarget && (lastIdx === 3 || restLen >= reqRest)) {
          const rPen = lastIdx !== 3 && lastIdx !== 2 && restLen === 1 && !adjLeaveBefore ? 15 : 0;
          const firstMl = adjLeaveBefore ? 5 : 0;
          if (lastIdx === 3) {
            for (const ns of [0, 1, 2]) {
              if (allowed0.includes(ns)) {
                startEdges.push({ act: ns, nextState: enc(1, ns, firstMl, 0, 0), ...edgePens(0, ns, 0) });
              }
            }
          } else {
            const ns = GEN_NEXT[lastIdx];
            if (allowed0.includes(ns)) {
              startEdges.push({ act: ns, nextState: enc(1, ns, firstMl, 0, 0), ...edgePens(0, ns, rPen) });
            }
            if (adjLeaveBefore && allowed0.includes(lastIdx)) {
              startEdges.push({ act: lastIdx, nextState: enc(1, lastIdx, firstMl, 0, 0), ...edgePens(0, lastIdx, 0) });
            }
          }
        }
      } else {
        if (allowed0.includes(3)) {
          startEdges.push({ act: 3, nextState: enc(0, 3, 10, 1, 0), ...edgePens(0, 3, 0) });
        }
        if (allowed0.includes(4)) {
          startEdges.push({ act: 4, nextState: enc(0, 3, 10, 0, 0), ...edgePens(0, 4, 0) });
        }
        if (allowed0.includes(5)) {
          startEdges.push({ act: 5, nextState: enc(0, 3, 10, 0, 0), ...edgePens(0, 5, 0) });
        }
        if (1 <= effectiveTarget) {
          for (const ns of [0, 1, 2]) {
            if (allowed0.includes(ns)) {
              startEdges.push({ act: ns, nextState: enc(1, ns, 5, 0, 0), ...edgePens(0, ns, 0) });
            }
          }
        }
      }

      if (startEdges.length === 0) {
        return { emp, isFixedTir: false, origCodes, origActs, layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
      }

      const activeStates: number[][] = Array.from({ length: daysCount + 1 }, () => []);
      const seen = new Uint8Array(numStates);
      for (const e of startEdges) {
        if (!seen[e.nextState]) {
          seen[e.nextState] = 1;
          activeStates[1].push(e.nextState);
        }
      }

      const stepTransitions = (d: number, st: number) => {
        const sb = st & 1;
        const t0 = st >> 1;
        const offS = t0 % 4;
        const t1 = (t0 - offS) / 4;
        const ml = t1 % 13;
        const t2 = (t1 - ml) / 13;
        const sh = t2 % 4;
        const w = (t2 - sh) / 4;
        const allowed = dayAllowed[d];
        const prevWasLeave = origActs[d - 1] === 4 || origActs[d - 1] === 5;
        const out: Array<{ act: number; nextState: number; penA: number; penB: number }> = [];

        if (ml < 10) {
          const isUnb = ml >= 5;
          const len = isUnb ? ml - 4 : ml + 1;
          if (len < 5 && w + 1 <= effectiveTarget && allowed.includes(sh)) {
            out.push({ act: sh, nextState: enc(w + 1, sh, ml + 1, 0, sb), ...edgePens(d, sh, 0) });
          }
          if (len >= 3 || isUnb) {
            if (allowed.includes(3)) {
              out.push({ act: 3, nextState: enc(w, sh, 10, 1, sb), ...edgePens(d, 3, 0) });
            }
          } else if (allowShortBlock2 && len === 2 && sb === 0 && allowed.includes(3)) {
            out.push({ act: 3, nextState: enc(w, sh, 10, 1, 1), ...edgePens(d, 3, 25) });
          }
          // Ending work block immediately before ABS/H8 is always exempt (adjAfter)
          if (allowed.includes(4)) {
            out.push({ act: 4, nextState: enc(w, sh, 10, 0, sb), ...edgePens(d, 4, 0) });
          }
          if (allowed.includes(5)) {
            out.push({ act: 5, nextState: enc(w, sh, 10, 0, sb), ...edgePens(d, 5, 0) });
          }
        } else {
          const restLen = ml - 9;
          const nextMl = Math.min(12, ml + 1);
          if (offS < 3 && allowed.includes(3)) {
            out.push({ act: 3, nextState: enc(w, sh, nextMl, offS + 1, sb), ...edgePens(d, 3, 0) });
          }
          if (allowed.includes(4)) {
            out.push({ act: 4, nextState: enc(w, sh, nextMl, 0, sb), ...edgePens(d, 4, 0) });
          }
          if (allowed.includes(5)) {
            out.push({ act: 5, nextState: enc(w, sh, nextMl, 0, sb), ...edgePens(d, 5, 0) });
          }
          if (w + 1 <= effectiveTarget) {
            const nextMlStart = prevWasLeave ? 5 : 0;
            if (sh === 3) {
              for (const ns of [0, 1, 2]) {
                if (allowed.includes(ns)) {
                  out.push({ act: ns, nextState: enc(w + 1, ns, nextMlStart, 0, sb), ...edgePens(d, ns, 0) });
                }
              }
            } else if (restLen >= (sh === 2 ? 2 : (prevWasLeave ? 1 : minNonNightRest))) {
              const ns = GEN_NEXT[sh];
              if (allowed.includes(ns)) {
                const rPen = sh !== 2 && restLen === 1 && !prevWasLeave ? 15 : 0;
                out.push({ act: ns, nextState: enc(w + 1, ns, nextMlStart, 0, sb), ...edgePens(d, ns, rPen) });
              }
              if (prevWasLeave && allowed.includes(sh)) {
                out.push({ act: sh, nextState: enc(w + 1, sh, nextMlStart, 0, sb), ...edgePens(d, sh, 0) });
              }
            }
          }
        }
        return out;
      };

      for (let d = 1; d < daysCount; d++) {
        seen.fill(0);
        for (const st of activeStates[d]) {
          for (const tr of stepTransitions(d, st)) {
            if (!seen[tr.nextState]) {
              seen[tr.nextState] = 1;
              activeStates[d + 1].push(tr.nextState);
            }
          }
        }
      }

      const reachNext = new Int32Array(numStates).fill(-1);
      let termIdx = 0;
      for (const st of activeStates[daysCount]) {
        const sb = st & 1;
        const t0 = st >> 1;
        const offS = t0 % 4;
        const t1 = (t0 - offS) / 4;
        const ml = t1 % 13;
        const t2 = (t1 - ml) / 13;
        const sh = t2 % 4;
        const w = (t2 - sh) / 4;
        if (w === effectiveTarget && (sb === 0 || allowShortBlock2)) {
          reachNext[st] = termIdx++;
        }
      }
      if (termIdx === 0) {
        return { emp, isFixedTir: false, origCodes, origActs, layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
      }

      const layers: GenEdge[][][] = new Array(daysCount);
      const reachCur = new Int32Array(numStates);
      for (let d = daysCount - 1; d >= 1; d--) {
        reachCur.fill(-1);
        const dayLayer: GenEdge[][] = [];
        for (const st of activeStates[d]) {
          const nodeEdges: GenEdge[] = [];
          for (const tr of stepTransitions(d, st)) {
            const nxtIdx = reachNext[tr.nextState];
            if (nxtIdx !== -1) {
              nodeEdges.push({ act: tr.act, nextNode: nxtIdx, penA: tr.penA, penB: tr.penB });
            }
          }
          if (nodeEdges.length > 0) {
            reachCur[st] = dayLayer.length;
            dayLayer.push(nodeEdges);
          }
        }
        if (dayLayer.length === 0) {
          return { emp, isFixedTir: false, origCodes, origActs, layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
        }
        layers[d] = dayLayer;
        reachNext.set(reachCur);
      }

      const rootEdges: GenEdge[] = [];
      for (const e of startEdges) {
        const nxtIdx = reachNext[e.nextState];
        if (nxtIdx !== -1) {
          rootEdges.push({ act: e.act, nextNode: nxtIdx, penA: e.penA, penB: e.penB });
        }
      }
      if (rootEdges.length === 0) {
        return { emp, isFixedTir: false, origCodes, origActs, layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
      }
      layers[0] = [rootEdges];

      let maxNodes = 1;
      const cAct: Int8Array[] = new Array(daysCount);
      const cNext: Int32Array[] = new Array(daysCount);
      for (let d = 0; d < daysCount; d++) {
        const len = layers[d].length;
        if (len > maxNodes) maxNodes = len;
        cAct[d] = new Int8Array(len);
        cNext[d] = new Int32Array(len);
      }
      return { emp, isFixedTir: false, origCodes, origActs, layers, feasible: true, maxNodes, cAct, cNext };
    };

    let dpNextGen = new Float64Array(512);
    let dpCurGen = new Float64Array(512);
    const dayCostGen = new Float64Array(daysCount * 4);

    const covPenaltyA = (m: number, a: number, n: number, dw: number = 1, _useSoft: boolean = false, tier: ComplianceTier = 0) =>
      tierCovPenalty(m, a, n, dw, tier);

    const covPenaltyB = (m: number, a: number, n: number, tier: ComplianceTier = 0) => {
      if (!isCoverageFeasibleForTier(m, a, n, tier)) {
        return tierCovPenalty(m, a, n, 100, tier);
      }
      return (tier >= 1 && n === 3 ? 5000 : 0) + (tier >= 3 && (m === 0 || a === 1) ? 15000 : 0);
    };

    // mode: 0 = gentle minimal-edit Phase A1, 1 = deep Stage-3 FastSA Phase A2, 2 = strict minimal-diff Phase B
    const solveGenPath = (
      dag: GenDag,
      otherM: Int32Array,
      otherA: Int32Array,
      otherN: Int32Array,
      dayWeights: Float64Array | null,
      noise: Float64Array | null,
      mode: 0 | 1 | 2,
      tier: ComplianceTier
    ): number[] => {
      if (dag.maxNodes > dpNextGen.length) {
        dpNextGen = new Float64Array(dag.maxNodes + 64);
        dpCurGen = new Float64Array(dag.maxNodes + 64);
      }
      const useSoft = mode === 1;
      if (dag.isFixedTir) {
        for (let d = 0; d < daysCount; d++) {
          const m0 = otherM[d];
          const a0 = otherA[d];
          const n0 = otherN[d];
          const dw = dayWeights ? dayWeights[d] : 1;
          const base = d * 4;
          dayCostGen[base + 0] = (mode === 2 ? covPenaltyB(m0, a0 + 1, n0, tier) : covPenaltyA(m0, a0 + 1, n0, dw, useSoft, tier)) + (noise ? noise[base + 0] : 0);
          dayCostGen[base + 3] = (mode === 2 ? covPenaltyB(m0, a0, n0, tier) : covPenaltyA(m0, a0, n0, dw, useSoft, tier)) + (noise ? noise[base + 3] : 0);
        }
      } else {
        for (let d = 0; d < daysCount; d++) {
          const m0 = otherM[d];
          const a0 = otherA[d];
          const n0 = otherN[d];
          const dw = dayWeights ? dayWeights[d] : 1;
          const base = d * 4;
          if (mode === 2) {
            dayCostGen[base + 0] = covPenaltyB(m0, a0 + 1, n0, tier);
            dayCostGen[base + 1] = covPenaltyB(m0 + 1, a0, n0, tier);
            dayCostGen[base + 2] = covPenaltyB(m0, a0, n0 + 1, tier);
            dayCostGen[base + 3] = covPenaltyB(m0, a0, n0, tier);
          } else {
            dayCostGen[base + 0] = covPenaltyA(m0, a0 + 1, n0, dw, useSoft, tier) + (noise ? noise[base + 0] : 0);
            dayCostGen[base + 1] = covPenaltyA(m0 + 1, a0, n0, dw, useSoft, tier) + (noise ? noise[base + 1] : 0);
            dayCostGen[base + 2] = covPenaltyA(m0, a0, n0 + 1, dw, useSoft, tier) + (noise ? noise[base + 2] : 0);
            dayCostGen[base + 3] = covPenaltyA(m0, a0, n0, dw, useSoft, tier) + (noise ? noise[base + 3] : 0);
          }
        }
      }

      const { layers, cAct, cNext } = dag;
      for (let d = daysCount - 1; d >= 0; d--) {
        const layer = layers[d];
        const isLast = d === daysCount - 1;
        const cA = cAct[d];
        const cN = cNext[d];
        const base = d * 4;
        const c0 = dayCostGen[base + 0];
        const c1 = dayCostGen[base + 1];
        const c2 = dayCostGen[base + 2];
        const c3 = dayCostGen[base + 3];

        for (let u = 0; u < layer.length; u++) {
          const edges = layer[u];
          let bestVal = Infinity;
          let bestA = 3;
          let bestV = 0;
          for (let ei = 0; ei < edges.length; ei++) {
            const edge = edges[ei];
            const act = edge.act;
            const actCost = act === 0 ? c0 : act === 1 ? c1 : act === 2 ? c2 : c3;
            const edgePen = mode === 2 ? edge.penB : (mode === 0 ? edge.penA : 0);
            const vCost = isLast ? 0 : dpNextGen[edge.nextNode];
            const val = actCost + edgePen + vCost;
            if (val < bestVal) {
              bestVal = val;
              bestA = act;
              bestV = edge.nextNode;
            }
          }
          dpCurGen[u] = bestVal;
          cA[u] = bestA;
          cN[u] = bestV;
        }
        const tmp = dpNextGen;
        dpNextGen = dpCurGen;
        dpCurGen = tmp;
      }

      const path = new Array<number>(daysCount);
      let curNode = 0;
      for (let d = 0; d < daysCount; d++) {
        path[d] = cAct[d][curNode];
        curNode = cNext[d][curNode];
      }
      return path;
    };

    const addGenPath = (dag: GenDag, path: number[], sign: number, mArr: Int32Array, aArr: Int32Array, nArr: Int32Array) => {
      if (dag.isFixedTir) {
        for (let d = 0; d < daysCount; d++) {
          if (path[d] === 0) aArr[d] += sign;
        }
      } else {
        for (let d = 0; d < daysCount; d++) {
          const act = path[d];
          if (act === 0) aArr[d] += sign;
          else if (act === 1) mArr[d] += sign;
          else if (act === 2) nArr[d] += sign;
        }
      }
    };

    const evalGenCov = (
      dags: GenDag[],
      paths: number[][],
      mArr: Int32Array,
      aArr: Int32Array,
      nArr: Int32Array,
      dayWeights: Float64Array | null,
      mode: 0 | 1 | 2,
      tier: ComplianceTier
    ) => {
      let hard = 0;
      let weightedTotal = 0;
      let total = 0;
      let editPenB = 0;
      let diffCount = 0;
      const useSoft = mode === 1;
      for (let d = 0; d < daysCount; d++) {
        const m = mArr[d];
        const a = aArr[d];
        const n = nArr[d];
        if (!isCoverageFeasibleForTier(m, a, n, tier)) hard++;
        const baseP = mode === 2 ? covPenaltyB(m, a, n, tier) : covPenaltyA(m, a, n, 1, useSoft, tier);
        total += baseP;
        weightedTotal += mode === 2 ? baseP : (dayWeights ? covPenaltyA(m, a, n, dayWeights[d], useSoft, tier) : baseP);
      }
      for (let i = 0; i < dags.length; i++) {
        const dg = dags[i];
        const p = paths[i];
        let u = 0;
        for (let d = 0; d < daysCount; d++) {
          const act = p[d];
          if (act !== dg.origActs[d]) diffCount++;
          const edges = dg.layers[d][u];
          for (let ei = 0; ei < edges.length; ei++) {
            if (edges[ei].act === act) {
              const ep = mode === 2 ? edges[ei].penB : (mode === 0 ? edges[ei].penA : 0);
              weightedTotal += ep;
              total += ep;
              editPenB += edges[ei].penB;
              u = edges[ei].nextNode;
              break;
            }
          }
        }
      }
      return { hard, weightedTotal, total, editPenB, diffCount };
    };

    const buildRowsFromGenPaths = (dags: GenDag[], paths: number[][]): Record<string, Record<string, string>> => {
      const candidate = cloneRows(rows);
      for (let i = 0; i < dags.length; i++) {
        const dg = dags[i];
        const p = paths[i];
        for (let d = 0; d < daysCount; d++) {
          const iso = dateColumns[d].isoDate;
          const act = p[d];
          if (dg.isFixedTir) {
            candidate[dg.emp][iso] = dg.origCodes[d];
          } else {
            if (act === 0) candidate[dg.emp][iso] = 'A';
            else if (act === 1) candidate[dg.emp][iso] = 'M';
            else if (act === 2) candidate[dg.emp][iso] = 'N';
            else if (act === 4) candidate[dg.emp][iso] = 'ABS';
            else if (act === 5) candidate[dg.emp][iso] = dg.origCodes[d];
            else candidate[dg.emp][iso] = 'OFF';
          }
        }
      }
      for (const emp of allEmployees) {
        if (fixedMidSet.has(emp)) {
          // Keep original Palade shifts unchanged
          candidate[emp] = { ...rows[emp] };
        }
      }
      return candidate;
    };

    let genRngState = ((((variantSeed + 1) * 0x9e3779b9) ^ 0x517cc1b7) >>> 0);
    const genRand = (): number => {
      genRngState = (genRngState + 0x6d2b79f5) >>> 0;
      let t = Math.imul(genRngState ^ (genRngState >>> 15), 1 | genRngState);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
    const shuffleGen = (arr: number[]): void => {
      for (let i = arr.length - 1; i > 0; i--) {
        const j = Math.floor(genRand() * (i + 1));
        const tmp = arr[i];
        arr[i] = arr[j];
        arr[j] = tmp;
      }
    };
    const isGenPathValid = (dag: GenDag, path: number[]): boolean =>
      isDagPathValid(dag, path, daysCount);
    const solveGenExtremeCount = (dag: GenDag, targetAct: number, sign: 1 | -1): number =>
      solveDagExtremeCount(dag, daysCount, targetAct, sign);

    let bestRepairedRows: Record<string, Record<string, string>> | null = null;
    let bestRepairedValidation = currentValidation;
    let bestRepairedTier: ComplianceTier = 0;
    let bestRepairedRelaxations: RelaxationRecord[] = [];
    let bestRepairedHardTotal = Infinity;
    let bestRepairedDiffCost = Infinity;
    let foundCompliantTier = false;

    const solverEmps = [...rotatingEmployees, ...fixedMaEmps];
    const G = solverEmps.length;

    const hasSufficientGenCapacity = (dags: GenDag[], tier: ComplianceTier): boolean => {
      const fixedMaxA = new Int32Array(daysCount);
      const availM = new Int32Array(daysCount);
      const availA = new Int32Array(daysCount);
      const availN = new Int32Array(daysCount);
      let sumMinRotA = 0, sumMaxRotA = 0, sumMinM = 0, sumMaxM = 0, sumMinN = 0, sumMaxN = 0;
      let sumMaxFixedA = 0;
      let totalRotSlots = 0;

      for (const dg of dags) {
        for (let d = 0; d < daysCount; d++) {
          let hasA = false, hasM = false, hasN = false;
          for (const layerEdges of dg.layers[d]) {
            for (let ei = 0; ei < layerEdges.length; ei++) {
              const a = layerEdges[ei].act;
              if (a === 0) hasA = true;
              else if (!dg.isFixedTir && a === 1) hasM = true;
              else if (!dg.isFixedTir && a === 2) hasN = true;
            }
          }
          if (hasA) {
            availA[d]++;
            if (dg.isFixedTir) fixedMaxA[d]++;
          }
          if (hasM) availM[d]++;
          if (hasN) availN[d]++;
        }

        if (dg.isFixedTir) {
          sumMaxFixedA += solveGenExtremeCount(dg, 0, -1);
        } else {
          const drops = dg.origActs.filter((a) => a === 4 || a === 5).length;
          totalRotSlots += Math.max(0, targetWorkingDays - drops);
          sumMinRotA += solveGenExtremeCount(dg, 0, 1);
          sumMaxRotA += solveGenExtremeCount(dg, 0, -1);
          sumMinM += solveGenExtremeCount(dg, 1, 1);
          sumMaxM += solveGenExtremeCount(dg, 1, -1);
          sumMinN += solveGenExtremeCount(dg, 2, 1);
          sumMaxN += solveGenExtremeCount(dg, 2, -1);
        }
      }

      return canSatisfyDailyCapacityForTier({
        tier,
        daysCount,
        totalRotSlots,
        availM,
        availTotalA: availA,
        availN,
        fixedMaxA,
        sumMinRotA,
        sumMaxRotA,
        sumMaxFixedA,
        sumMinM,
        sumMaxM,
        sumMinN,
        sumMaxN
      });
    };

    const gNoise = new Float64Array(daysCount * 4);
    const gDayWeights = new Float64Array(daysCount);
    const gPool = Array.from({ length: G }, (_, i) => i);
    const s5Helpers: number[] = [];
    const s5Capable: number[] = [];
    const s5Unlocked: number[] = [];
    const s5Locked: number[] = [];

    const buildGenDayActMasks = (dags: GenDag[]): Uint8Array[] =>
      buildDagDayActMasks(dags, daysCount);

    const appendRandomFromBucket = (bucket: number[], out: number[], limit: number): void => {
      let rem = bucket.length;
      while (out.length < limit && rem > 0) {
        const r = Math.floor(genRand() * rem);
        out.push(bucket[r]);
        bucket[r] = bucket[rem - 1];
        rem--;
      }
    };

    const selectSection5RuinGroup = (
      dags: GenDag[],
      genDayActMasks: Uint8Array[],
      paths: number[][],
      focusDay: number,
      k: number,
      curM: Int32Array,
      curA: Int32Array,
      curN: Int32Array,
      tier: ComplianceTier
    ): number[] => {
      if (genRand() < 0.2) {
        shuffleGen(gPool);
        return gPool.slice(0, k);
      }
      s5Helpers.length = 0;
      s5Capable.length = 0;
      s5Unlocked.length = 0;
      s5Locked.length = 0;

      const cfg = getTierConfig(tier);
      const minReqM = cfg.allowCoverageGapMA ? 0 : 1;
      const minReqA = cfg.allowCoverageGapMA ? 1 : 2;
      const maxReqN = cfg.allowNightOver2 ? 3 : 2;
      const nVal = curN[focusDay];
      const mVal = curM[focusDay];
      const aVal = curA[focusDay];
      const needAct = nVal < 2 ? 2 : mVal < minReqM ? 1 : aVal < minReqA ? 0 : (nVal > maxReqN ? -1 : -2);

      for (let e = 0; e < G; e++) {
        const dg = dags[e];
        const mask = genDayActMasks[e][focusDay];
        // Defer rotating employees locked to ABS/H8 (acts 4/5) and Tir when focusDay needs M/N or Tir cannot hold/toggle A
        const isAct45Locked = !dg.isFixedTir && (mask & ~((1 << 4) | (1 << 5))) === 0;
        const isTirExcluded = dg.isFixedTir && (needAct !== 0 || (mask & 1) === 0 || mask === 1);
        if (isAct45Locked || isTirExcluded) {
          s5Locked.push(e);
          continue;
        }
        if (needAct >= 0) {
          if (canDoShift(genDayActMasks, e, focusDay, needAct)) {
            if (paths[e][focusDay] !== needAct) s5Helpers.push(e);
            else s5Capable.push(e);
          } else {
            s5Unlocked.push(e);
          }
        } else if (needAct === -1) {
          if (paths[e][focusDay] === 2 && (mask & ~((1 << 2) | (1 << 4) | (1 << 5))) !== 0) {
            s5Helpers.push(e);
          } else {
            s5Unlocked.push(e);
          }
        } else {
          s5Unlocked.push(e);
        }
      }

      const group: number[] = [];
      appendRandomFromBucket(s5Helpers, group, k);
      appendRandomFromBucket(s5Capable, group, k);
      appendRandomFromBucket(s5Unlocked, group, k);
      appendRandomFromBucket(s5Locked, group, k);
      return group;
    };
    const lockModes: Array<'globalPre' | 'empPre' | 'none'> = hasDistinctEmpAbsStarts
      ? ['globalPre', 'empPre', 'none']
      : ['globalPre', 'none'];

    for (let ti = 0; ti < absTiers.length; ti++) {
      if (foundCompliantTier) break;
      const nowMs = performance.now();
      if (nowMs >= s5DeadlineMs) break;

      const tier = absTiers[ti];
      const cfg = getTierConfig(tier);
      const remBudgetMs = Math.max(0, s5DeadlineMs - nowMs);
      const s5Share = tier === 0 ? 0.25 : tier === 1 ? 0.15 : tier === 2 ? 0.40 : 0.20;
      const remShares = absTiers
        .slice(ti)
        .reduce((sum, t) => sum + (t === 0 ? 0.25 : t === 1 ? 0.15 : t === 2 ? 0.40 : 0.20), 0);
      const tierAllocMs = remShares > 0 ? (remBudgetMs * s5Share) / remShares : remBudgetMs;
      let tierDeadlineMs = Math.min(s5DeadlineMs, nowMs + tierAllocMs);

      let tierBestRows: Record<string, Record<string, string>> | null = null;
      let tierBestValidation: ScheduleValidationReport | null = null;
      let tierBestRelaxations: RelaxationRecord[] = [];
      let tierBestDiffCost = Infinity;

      for (const lockMode of lockModes) {
        if (performance.now() >= tierDeadlineMs) break;
        if (tierBestValidation?.totalViolations === 0) break;

        let fallbackToNoneCount = 0;
        const dags: GenDag[] = solverEmps.map((emp, eIdx) => {
          if (fixedMaEmps.includes(emp)) return buildTirGenDag(emp);
          let dg = buildGeneralDag(emp, eIdx, cfg.minRest, cfg.allowShortBlock2, lockMode);
          if (!dg.feasible && lockMode === 'globalPre') {
            dg = buildGeneralDag(emp, eIdx, cfg.minRest, cfg.allowShortBlock2, 'empPre');
            if (!dg.feasible) {
              dg = buildGeneralDag(emp, eIdx, cfg.minRest, cfg.allowShortBlock2, 'none');
              fallbackToNoneCount++;
            }
          } else if (!dg.feasible && lockMode === 'empPre') {
            dg = buildGeneralDag(emp, eIdx, cfg.minRest, cfg.allowShortBlock2, 'none');
            fallbackToNoneCount++;
          }
          return dg;
        });

        if (dags.some((g) => !g.feasible) || !hasSufficientGenCapacity(dags, tier)) continue;
        if (lockMode !== 'none' && fallbackToNoneCount > Math.floor(G / 2)) continue;

        const genDayActMasks = buildGenDayActMasks(dags);
        const origPathsValid = dags.map((dg) => isGenPathValid(dg, dg.origActs));
        const validOrigCount = origPathsValid.filter(Boolean).length;
        if (cfg.minRest === 2 && validOrigCount < Math.max(2, G - 3)) {
          tierDeadlineMs = Math.min(tierDeadlineMs, performance.now() + 250);
        }
        const numRestarts =
          lockMode === 'none'
            ? cfg.minRest === 2
              ? validOrigCount < Math.max(2, G - 3)
                ? 1
                : 4
              : 8
            : 2;

        for (let restart = 0; restart < numRestarts; restart++) {
          if (performance.now() >= tierDeadlineMs) break;

          gDayWeights.fill(1);
          const curM = new Int32Array(daysCount);
          const curA = new Int32Array(daysCount);
          const curN = new Int32Array(daysCount);
          let paths: number[][] = new Array(G);

          const order = Array.from({ length: G }, (_, i) => i);
          shuffleGen(order);

          const deepOnlyRestart = lockMode === 'none' && restart >= 2;
          const initMode: 0 | 1 = deepOnlyRestart ? 1 : 0;

          if (restart === 0) {
            // Always warm-start from original schedule for all legal employees, solve remaining greedily
            for (let i = 0; i < G; i++) {
              if (origPathsValid[i]) {
                paths[i] = [...dags[i].origActs];
                addGenPath(dags[i], paths[i], 1, curM, curA, curN);
              }
            }
            for (let oi = 0; oi < G; oi++) {
              const i = order[oi];
              if (!origPathsValid[i]) {
                paths[i] = solveGenPath(dags[i], curM, curA, curN, null, null, 0, tier);
                addGenPath(dags[i], paths[i], 1, curM, curA, curN);
              }
            }
          } else if (deepOnlyRestart) {
            // Full fresh Stage-3-style randomized initialization to escape deep origActs basins
            for (let oi = 0; oi < G; oi++) {
              const i = order[oi];
              if (oi < 3) {
                for (let k = 0; k < gNoise.length; k++) gNoise[k] = genRand() * 250;
              } else {
                for (let k = 0; k < gNoise.length; k++) gNoise[k] = genRand() * 8;
              }
              paths[i] = solveGenPath(dags[i], curM, curA, curN, null, gNoise, 1, tier);
              addGenPath(dags[i], paths[i], 1, curM, curA, curN);
            }
          } else {
            // Drop random subset from original warm-start instead of scrambling the whole roster
            for (let oi = 0; oi < G; oi++) {
              const i = order[oi];
              if (origPathsValid[i] && oi >= 2 + (restart % 3)) {
                paths[i] = [...dags[i].origActs];
                addGenPath(dags[i], paths[i], 1, curM, curA, curN);
              } else {
                paths[i] = [];
              }
            }
            for (let oi = 0; oi < G; oi++) {
              const i = order[oi];
              if (paths[i].length === 0) {
                for (let k = 0; k < gNoise.length; k++) gNoise[k] = genRand() * (oi < 3 ? 120 : 12);
                paths[i] = solveGenPath(dags[i], curM, curA, curN, null, gNoise, 0, tier);
                addGenPath(dags[i], paths[i], 1, curM, curA, curN);
              }
            }
          }

          const runGenSweeps = (maxSweeps: number, useWeights: boolean, mode: 0 | 1 | 2) => {
            const wArr = useWeights && mode !== 2 ? gDayWeights : null;
            for (let sweep = 0; sweep < maxSweeps; sweep++) {
              let changed = false;
              for (let oi = 0; oi < G; oi++) {
                const eIdx = order[(oi + sweep) % G];
                const oldP = paths[eIdx];
                const before = evalGenCov(dags, paths, curM, curA, curN, wArr, mode, tier).weightedTotal;
                addGenPath(dags[eIdx], oldP, -1, curM, curA, curN);
                const newP = solveGenPath(dags[eIdx], curM, curA, curN, wArr, null, mode, tier);
                paths[eIdx] = newP;
                addGenPath(dags[eIdx], newP, 1, curM, curA, curN);
                const after = evalGenCov(dags, paths, curM, curA, curN, wArr, mode, tier).weightedTotal;
                if (after < before) {
                  changed = true;
                } else if (after > before) {
                  addGenPath(dags[eIdx], newP, -1, curM, curA, curN);
                  paths[eIdx] = oldP;
                  addGenPath(dags[eIdx], oldP, 1, curM, curA, curN);
                }
              }
              if (!changed) break;
            }
          };

          runGenSweeps(deepOnlyRestart ? 3 : 4, false, initMode);

          let curEval = evalGenCov(dags, paths, curM, curA, curN, gDayWeights, initMode, tier);
          let bestRH = curEval.hard;
          let bestRT = curEval.total;
          let bestRPaths = paths.map((p) => p);

          if (bestRH > 0) {
            const gentleSteps = deepOnlyRestart ? 0 : 350;
            const saSteps = deepOnlyRestart ? 2000 : 1850;
            let temp = deepOnlyRestart ? 1600 : 300;
            let cool = deepOnlyRestart ? 0.994 : 0.985;
            for (let step = 0; step < saSteps && bestRH > 0; step++) {
              if ((step & 15) === 0 && performance.now() >= tierDeadlineMs) break;

              if (!deepOnlyRestart && step === gentleSteps) {
                paths = bestRPaths.map((p) => p);
                curM.fill(0);
                curA.fill(0);
                curN.fill(0);
                for (let i = 0; i < G; i++) addGenPath(dags[i], paths[i], 1, curM, curA, curN);
                gDayWeights.fill(1);
                curEval = evalGenCov(dags, paths, curM, curA, curN, gDayWeights, 1, tier);
                temp = 1600;
                cool = 0.994;
              } else if (step === (deepOnlyRestart ? 1000 : 1150) || (deepOnlyRestart && step === 1550)) {
                temp = Math.max(temp, 750);
              }
              const stepMode: 0 | 1 = step < gentleSteps ? 0 : 1;

              const badDays: number[] = [];
              for (let d = 0; d < daysCount; d++) {
                if (!isCoverageFeasibleForTier(curM[d], curA[d], curN[d], tier)) badDays.push(d);
              }
              if (badDays.length === 0) break;

              if (step > 0 && step % 35 === 0 && step !== gentleSteps) {
                for (const bd of badDays) gDayWeights[bd] = Math.min(8, gDayWeights[bd] + 0.5);
                if (step % 210 === 0) {
                  for (let d = 0; d < daysCount; d++) gDayWeights[d] = 1 + (gDayWeights[d] - 1) * 0.5;
                }
                curEval = evalGenCov(dags, paths, curM, curA, curN, gDayWeights, stepMode, tier);
              }

              const focusDay = badDays[Math.floor(genRand() * badDays.length)];
              const k = stepMode === 0 ? Math.min(G, 2) : Math.min(G, step % 2 === 0 ? 3 : 2);
              const ruinGroup = selectSection5RuinGroup(dags, genDayActMasks, paths, focusDay, k, curM, curA, curN, tier);

              const savedSub = ruinGroup.map((idx) => paths[idx]);
              for (let gi = 0; gi < ruinGroup.length; gi++) {
                addGenPath(dags[ruinGroup[gi]], paths[ruinGroup[gi]], -1, curM, curA, curN);
              }

              const noiseAmp = stepMode === 0 ? Math.min(40, temp * 0.15) : Math.min(350, temp * 0.3);
              for (let gi = 0; gi < ruinGroup.length; gi++) {
                const eIdx = ruinGroup[gi];
                if (gi === 0 && genRand() < 0.5) {
                  for (let nk = 0; nk < gNoise.length; nk++) gNoise[nk] = genRand() * (stepMode === 0 ? 2 : 6);
                  for (let d = Math.max(0, focusDay - 5); d <= Math.min(daysCount - 1, focusDay + 5); d++) {
                    for (let a = 0; a < 4; a++) gNoise[d * 4 + a] += (genRand() - 0.5) * noiseAmp * 3;
                  }
                } else {
                  for (let nk = 0; nk < gNoise.length; nk++) gNoise[nk] = genRand() * (2 + noiseAmp * 0.25);
                }
                const newP = solveGenPath(dags[eIdx], curM, curA, curN, gDayWeights, gNoise, stepMode, tier);
                paths[eIdx] = newP;
                addGenPath(dags[eIdx], newP, 1, curM, curA, curN);
              }

              let nextEval = evalGenCov(dags, paths, curM, curA, curN, gDayWeights, stepMode, tier);
              const improvedRec = nextEval.hard < bestRH;
              if (improvedRec) {
                runGenSweeps(2, true, stepMode);
                nextEval = evalGenCov(dags, paths, curM, curA, curN, gDayWeights, stepMode, tier);
              }
              const delta = nextEval.weightedTotal - curEval.weightedTotal;
              const accept = improvedRec || delta <= 0 || genRand() < Math.exp(-delta / Math.max(1, temp));
              if (accept) {
                curEval = nextEval;
                if (curEval.hard < bestRH || (curEval.hard === bestRH && curEval.total < bestRT)) {
                  bestRH = curEval.hard;
                  bestRT = curEval.total;
                  bestRPaths = paths.map((p) => p);
                  if (bestRH === 0) break;
                }
              } else {
                for (let gi = 0; gi < ruinGroup.length; gi++) {
                  const eIdx = ruinGroup[gi];
                  addGenPath(dags[eIdx], paths[eIdx], -1, curM, curA, curN);
                  paths[eIdx] = savedSub[gi];
                  addGenPath(dags[eIdx], savedSub[gi], 1, curM, curA, curN);
                }
              }
              temp *= cool;
            }
          }

          paths = bestRPaths;
          curM.fill(0);
          curA.fill(0);
          curN.fill(0);
          for (let i = 0; i < G; i++) addGenPath(dags[i], paths[i], 1, curM, curA, curN);

          // Phase B: Strict feasibility-preserving deterministic minimal-diff polishing
          runGenSweeps(6, false, 2);

          // Pairwise 2-employee zero-noise minimization across modified employees
          for (let pass = 0; pass < 2; pass++) {
            if (performance.now() >= tierDeadlineMs) break;
            let pairImproved = false;
            const changedEmps: number[] = [];
            for (let i = 0; i < G; i++) {
              for (let d = 0; d < daysCount; d++) {
                if (paths[i][d] !== dags[i].origActs[d]) {
                  changedEmps.push(i);
                  break;
                }
              }
            }
            for (let aIdx = 0; aIdx < changedEmps.length; aIdx++) {
              if (performance.now() >= tierDeadlineMs) break;
              for (let bIdx = aIdx + 1; bIdx < changedEmps.length; bIdx++) {
                if (performance.now() >= tierDeadlineMs) break;
                const i1 = changedEmps[aIdx];
                const i2 = changedEmps[bIdx];
                const baseScore = evalGenCov(dags, paths, curM, curA, curN, null, 2, tier);
                const old1 = paths[i1];
                const old2 = paths[i2];

                for (const [first, second] of [[i1, i2], [i2, i1]]) {
                  const s1 = paths[first];
                  const s2 = paths[second];
                  addGenPath(dags[first], s1, -1, curM, curA, curN);
                  addGenPath(dags[second], s2, -1, curM, curA, curN);
                  const n1 = solveGenPath(dags[first], curM, curA, curN, null, null, 2, tier);
                  paths[first] = n1;
                  addGenPath(dags[first], n1, 1, curM, curA, curN);
                  const n2 = solveGenPath(dags[second], curM, curA, curN, null, null, 2, tier);
                  paths[second] = n2;
                  addGenPath(dags[second], n2, 1, curM, curA, curN);

                  const nextScore = evalGenCov(dags, paths, curM, curA, curN, null, 2, tier);
                  if (nextScore.hard <= baseScore.hard && nextScore.weightedTotal < baseScore.weightedTotal) {
                    pairImproved = true;
                    break;
                  } else {
                    addGenPath(dags[first], n1, -1, curM, curA, curN);
                    addGenPath(dags[second], n2, -1, curM, curA, curN);
                    paths[i1] = old1;
                    paths[i2] = old2;
                    addGenPath(dags[i1], old1, 1, curM, curA, curN);
                    addGenPath(dags[i2], old2, 1, curM, curA, curN);
                  }
                }
              }
            }
            if (!pairImproved) break;
            runGenSweeps(4, false, 2);
          }

          const finalEval = evalGenCov(dags, paths, curM, curA, curN, null, 2, tier);
          const candidateRows = buildRowsFromGenPaths(dags, paths);
          const candRelax: RelaxationRecord[] = [];
          const candidateRep = validateScheduleCompliance(
            candidateRows,
            dateColumns,
            priorRoster,
            targetYear,
            targetMonth,
            targetWorkingDays,
            [],
            { tier, relaxations: candRelax }
          );

          const isBetterInTier =
            !tierBestValidation ||
            candidateRep.totalViolations < tierBestValidation.totalViolations ||
            (candidateRep.totalViolations === tierBestValidation.totalViolations &&
              finalEval.editPenB < tierBestDiffCost);

          if (isBetterInTier) {
            tierBestRows = candidateRows;
            tierBestValidation = candidateRep;
            tierBestRelaxations = candRelax;
            tierBestDiffCost = finalEval.editPenB;
          }

          const hardRep = tier === 3
            ? candidateRep
            : validateScheduleCompliance(
                candidateRows,
                dateColumns,
                priorRoster,
                targetYear,
                targetMonth,
                targetWorkingDays,
                [],
                { tier: 3 }
              );
          const isBetterFallback =
            !bestRepairedRows ||
            hardRep.totalViolations < bestRepairedHardTotal ||
            (hardRep.totalViolations === bestRepairedHardTotal &&
              candidateRep.totalViolations < bestRepairedValidation.totalViolations) ||
            (hardRep.totalViolations === bestRepairedHardTotal &&
              candidateRep.totalViolations === bestRepairedValidation.totalViolations &&
              finalEval.editPenB < bestRepairedDiffCost);

          if (isBetterFallback) {
            bestRepairedRows = candidateRows;
            bestRepairedValidation = candidateRep;
            bestRepairedTier = tier;
            bestRepairedRelaxations = candRelax;
            bestRepairedHardTotal = hardRep.totalViolations;
            bestRepairedDiffCost = finalEval.editPenB;
          }

          if (candidateRep.totalViolations === 0) break;
        }
      }

      if (tierBestRows && tierBestValidation && tierBestValidation.totalViolations === 0) {
        bestRepairedRows = tierBestRows;
        bestRepairedValidation = tierBestValidation;
        bestRepairedTier = tier;
        bestRepairedRelaxations = tierBestRelaxations;
        foundCompliantTier = true;
        break;
      }
    }

    if (bestRepairedRows && !foundCompliantTier) {
      let resolved = false;
      for (const t of absTiers) {
        const tRelax: RelaxationRecord[] = [];
        const tRep = validateScheduleCompliance(
          bestRepairedRows,
          dateColumns,
          priorRoster,
          targetYear,
          targetMonth,
          targetWorkingDays,
          [],
          { tier: t, relaxations: tRelax }
        );
        if (tRep.totalViolations === 0) {
          bestRepairedTier = t;
          bestRepairedRelaxations = tRelax;
          bestRepairedValidation = tRep;
          foundCompliantTier = true;
          resolved = true;
          break;
        }
      }
      if (!resolved) {
        const t3Relax: RelaxationRecord[] = [];
        const t3Rep = validateScheduleCompliance(
          bestRepairedRows,
          dateColumns,
          priorRoster,
          targetYear,
          targetMonth,
          targetWorkingDays,
          [],
          { tier: 3, relaxations: t3Relax }
        );
        bestRepairedTier = 3;
        bestRepairedRelaxations = t3Relax;
        bestRepairedValidation = t3Rep;
      }
    }

    if (bestRepairedRows && (foundCompliantTier || bestRepairedValidation.totalViolations < currentValidation.totalViolations)) {
      for (const emp of allEmployees) {
        workingRows[emp] = { ...bestRepairedRows[emp] };
      }
      finalTier = bestRepairedTier;
      finalRelaxations = bestRepairedRelaxations;
      currentValidation = bestRepairedValidation;
      adjustments.length = 0;
      unresolvedDeficits.length = 0;

      for (const col of dateColumns) {
        const iso = col.isoDate;
        const absOnDay = allEmployees.filter((e) => shiftKey(rows[e]?.[iso] ?? 'OFF') === 'ABS').join(', ');
        for (const emp of allEmployees) {
          const origC = shiftKey(rows[emp]?.[iso] ?? 'OFF');
          const newC = shiftKey(workingRows[emp]?.[iso] ?? 'OFF');
          if (origC !== newC) {
            adjustments.push({
              isoDate: iso,
              employee: emp,
              fromShift: origC,
              toShift: newC,
              reason: absOnDay
                ? `Rebalanced schedule to legally cover sick colleague (${absOnDay}) while maintaining ${targetWorkingDays}-day monthly norm and rest rules.`
                : `Rebalanced shift to preserve legal ${targetWorkingDays}-day monthly norm, shift sequence, and rest rules.`,
              absColleague: absOnDay || 'Rebalancing'
            });
          }
        }

        // Recompute unresolvedDeficits cleanly if any coverage gaps remain
        const rotM = rotatingEmployees.filter((e) => shiftKey(workingRows[e]?.[iso] ?? 'OFF') === 'M').length;
        const totalA = allEmployees.filter((e) => shiftKey(workingRows[e]?.[iso] ?? 'OFF') === 'A').length;
        const totalN = allEmployees.filter((e) => shiftKey(workingRows[e]?.[iso] ?? 'OFF') === 'N').length;
        if (rotM < 1) {
          unresolvedDeficits.push({
            isoDate: iso,
            shift: 'M',
            needed: 1 - rotM,
            available: rotM,
            reason: `Rotating Morning below 1 on ${iso} after Tier ${finalTier} repair.`
          });
        }
        if (totalA < 2) {
          unresolvedDeficits.push({
            isoDate: iso,
            shift: 'A',
            needed: 2 - totalA,
            available: totalA,
            reason: `Afternoon below 2 on ${iso} after Tier ${finalTier} repair.`
          });
        }
        if (totalN < 2) {
          unresolvedDeficits.push({
            isoDate: iso,
            shift: 'N',
            needed: 2 - totalN,
            available: totalN,
            reason: `Night below 2 on ${iso} after Tier ${finalTier} repair.`
          });
        }
      }
    }
  }

  return {
    hasDeficits: adjustments.length > 0 || unresolvedDeficits.length > 0,
    totalDeficits: adjustments.length + unresolvedDeficits.length,
    adjustments,
    unresolvedDeficits,
    updatedRows: workingRows,
    validation: currentValidation,
    tier: finalTier,
    relaxations: finalRelaxations
  };
}

