// src/scheduleBuilder.ts
import type ExcelJS from 'exceljs';
import type { RosterData } from './types';
import { shiftKey, ROMANIAN_HOLIDAYS } from './scheduleUtils';

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

// Dec 31st New Year's Eve exemption: Palade and Tir George work normal weekday shift if Dec 31 is a weekday (e.g. 2026-12-31 is Thursday)
export const FIXED_STAFF_EXEMPT_OFF_DATES = new Set<string>();

export const MONTHLY_WORKING_DAYS_2026: Record<number, number> = {
  0: 22, // Jan
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
  // Keep 24/7 rotating staff monthly norms within a mathematically feasible range
  // (e.g. Jan 2027 has 3 weekday holidays dropping fixed weekday count to 18, which is over-constrained for 24/7 3..5-day blocks).
  const minRotatingTarget = daysInMonth >= 30 ? 20 : 19;
  return Math.max(minRotatingTarget, count);
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

export interface ScheduleBuilderOptions {
  seed?: number;
  targetWorkingDays?: number;
  minMorning?: number;
  minAfternoon?: number;
  minNight?: number;
  targetOffDays?: number;
  maxConsecutiveWork?: number;
  excludedEmployees?: string[];
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

  // If employee already had >= 2 days of rest at end of previous month, they must NOT get OFF on Day 1
  if (tailOffStreak >= 2 && firstShift === 'OFF') {
    cost += 10000000;
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
        // Rest period between two work blocks:
        // After Night shift: strictly >= 2 days (48h mandatory rest)
        // Max rest: at most 3 days
        if (i > 0 && i < blocks.length - 1) {
          const prevWork = blocks[i - 1];
          if (prevWork && prevWork.shift === 'N' && b.len < 2) return false;
          if (b.len > 3) return false;
        }
        // Rest bridging across month boundary
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
      if (hist[i] !== 'OFF' && hist[i] !== 'H8' && hist[i] !== 'HUP') tailStreak++;
      else break;
    }
    if (tailStreak === 0) continue;

    let headStreak = 0;
    for (let i = 0; i < Math.min(5, dateColumns.length); i++) {
      const s = rows[emp][dateColumns[i].isoDate];
      if (s !== 'OFF' && s !== 'H8' && s !== 'HUP') headStreak++;
      else break;
    }

    if (tailStreak + headStreak > 5) {
      const cutoff = 5 - tailStreak;
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

          // Target shift must be adjacent to an existing block of targetShift for this employee
          const prev = getPrevNeighbor(emp, i);
          const next = getNextNeighbor(emp, i);
          if (prev !== targetShift && next !== targetShift) continue;

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
          const prev = getPrevNeighbor(emp, i);
          const next = getNextNeighbor(emp, i);
          if (prev !== targetShift && next !== targetShift) continue;

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
          const prev1 = getPrevNeighbor(emp1, i);
          const next1 = getNextNeighbor(emp1, i);
          if (prev1 !== targetShift && next1 !== targetShift) continue;

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
          const prev1 = getPrevNeighbor(emp1, i);
          const next1 = getNextNeighbor(emp1, i);
          if (prev1 !== targetShift && next1 !== targetShift) continue;

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
  // Perfectly balanced: 4 A, 4 M, 4 N, and 6 OFF (ratio 1:1:1:1.5)
  const BALANCED_TEMPLATES = [
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
    for (let shift = 0; shift < 18; shift += 2) {
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
  let finalSeed = seed;

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
  const rotatingAnchors: Record<string, string | null> = {};
  rotatingEmployees.forEach((emp) => {
    const hist = getEmployeeHistoryTail(roster, emp, targetYear, targetMonth, 6);
    rotatingHistories[emp] = hist;
    rotatingAnchors[emp] = hist.length > 0 ? roster?.rows?.[emp]?.[targetStartIso] ?? null : null;
  });

  const SEED_OFFSETS = [0, 1, 3, 5, 7, 10, 100];

  const getSlotImbalancePenalty = (template: string[], slots: number[]): number => {
    const tLen = template.length;
    let penalty = 0;
    for (let d = 0; d < tLen; d++) {
      let m = 0;
      let a = 0;
      let n = 0;
      for (const offset of slots) {
        const s = template[(offset + d) % tLen];
        if (s === 'M') m++;
        else if (s === 'A') a++;
        else if (s === 'N') n++;
      }
      if (m < 1) penalty += 1;
      if (a < 2) penalty += 1;
      if (n !== 2) penalty += Math.abs(n - 2);
    }
    return penalty * 40;
  };

  for (const sOffset of SEED_OFFSETS) {
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

    // Stage 2: Equalize and validate distinct candidate configurations (early exit on fully compliant)
    const configLimit = sOffset === 0 ? 48 : 24;
    let tested = 0;

    for (const config of allCandidates) {
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
      const testReport = validateScheduleCompliance(
        candidateRows,
        dateColumns,
        roster,
        targetYear,
        targetMonth,
        targetWorkingDays,
        options.excludedEmployees ?? []
      );

      const scoreReport = (rep: ScheduleValidationReport) => {
        const legalHard =
          rep.streakViolationCount +
          rep.nightRestViolationCount +
          rep.excessiveRestViolationCount +
          rep.mixedBlockViolationCount +
          rep.workDaysMismatchCount +
          (rep.fixedStaffViolationCount ?? 0);
        return legalHard * 1000 + rep.nightOverstaffedDaysCount * 20 + rep.understaffedDaysCount * 10;
      };

      const curScore = scoreReport(testReport);
      const bestScore = bestReport ? scoreReport(bestReport) : Infinity;

      if (testReport.isFullyCompliant || !bestReport || curScore < bestScore) {
        bestResultRows = candidateRows;
        bestReport = testReport;
        finalSeed = config.attemptSeed;
        if (testReport.isFullyCompliant) {
          break;
        }
      }
    }

    if (bestReport?.isFullyCompliant) {
      break;
    }
  }
  const validation = bestReport ?? validateScheduleCompliance(
    bestResultRows,
    dateColumns,
    roster,
    targetYear,
    targetMonth,
    targetWorkingDays,
    options.excludedEmployees ?? []
  );

  return {
    month: targetMonth,
    year: targetYear,
    daysCount,
    targetWorkingDays,
    dateColumns,
    rows: bestResultRows,
    validation,
    seed: finalSeed,
    excludedEmployees: options.excludedEmployees ?? []
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
  excludedEmployees: string[] = []
): ScheduleValidationReport {
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

  const employees = Object.keys(rows);

  // 1. Check Daily Coverage
  dateColumns.forEach(({ date, isoDate }) => {
    const dayOfWeek = date.getDay();
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    const c = { M: 0, A: 0, N: 0, MID: 0, OFF: 0, H8: 0, ABS: 0, other: 0 };
    const dayWarnings: string[] = [];

    employees.forEach((emp) => {
      const s = shiftKey(rows[emp]?.[isoDate] ?? 'OFF');
      if (s === 'M') c.M++;
      else if (s === 'A') c.A++;
      else if (s === 'N') c.N++;
      else if (s === 'MID') c.MID++;
      else if (s === 'H8') c.H8++;
      else if (s === 'ABS') c.ABS++;
      else c.OFF++;
    });

    const rotatingEmployees = employees.filter((e) => !FIXED_MID_EMPLOYEES.includes(e) && !FIXED_WEEKDAY_MA_EMPLOYEES.includes(e));
    const rotC = { M: 0, A: 0, N: 0 };
    rotatingEmployees.forEach((emp) => {
      const s = shiftKey(rows[emp]?.[isoDate] ?? 'OFF');
      if (s === 'M') rotC.M++;
      else if (s === 'A') rotC.A++;
      else if (s === 'N') rotC.N++;
    });

    // Daily staffing checks:
    // Rule: Morning shift minimum is 1 (excluding Tir and Palade from the calculation -> rotC.M)
    // Afternoon minimum is 2, Night minimum is 2, Night maximum is 2
    const isUnderstaffed = rotC.M < 1 || c.A < 2 || c.N < 2;
    const isOverstaffedNight = c.N > 2;
    if (rotC.M < 1) dayWarnings.push(`Morning understaffed (${rotC.M}/1 minimum)`);
    if (c.A < 2) dayWarnings.push(`Afternoon understaffed (${c.A}/2 minimum)`);
    if (c.N < 2) dayWarnings.push(`Night understaffed (${c.N}/2 minimum)`);
    if (c.N > 2) dayWarnings.push(`Night overstaffed (${c.N}/2 maximum)`);

    if (isUnderstaffed) {
      understaffedDaysCount++;
    }
    if (isOverstaffedNight) {
      nightOverstaffedDaysCount++;
    }
    if (dayWarnings.length > 0) {
      warnings.push(...dayWarnings.map((w) => `${isoDate}: ${w}`));
    }

    dailyCoverage[isoDate] = {
      isoDate,
      dayOfWeek,
      isWeekend,
      M: c.M,
      rotM: rotC.M,
      A: c.A,
      N: c.N,
      MID: c.MID,
      OFF: c.OFF + c.H8 + c.ABS + c.other,
      ABS: c.ABS,
      totalWorking: c.M + c.A + c.N + c.MID,
      isUnderstaffed: isUnderstaffed || isOverstaffedNight,
      warnings: dayWarnings
    };
  });

  // 2. Check Employee Legal Constraints, Streaks & Rest Durations
  employees.forEach((emp) => {
    const isExcluded = excludedSet.has(emp);
    const isFixedMid = FIXED_MID_EMPLOYEES.includes(emp);
    const isFixedWeekdayMA = FIXED_WEEKDAY_MA_EMPLOYEES.includes(emp);
    const isNonRotating = isFixedMid || isFixedWeekdayMA;
    const empViolations: string[] = [];
    const counts = { M: 0, A: 0, N: 0, MID: 0, OFF: 0, H8: 0, ABS: 0 };

    let currentStreak = 0;
    let maxStreak = 0;
    let currentOffStreak = 0;

    const isRestCode = (code: string) =>
      code === 'OFF' || code === 'H8' || code === 'HUP' || code === 'ABS' || code === 'AD';

    const empHistory = (priorRoster && targetYear !== undefined && targetMonth !== undefined && !isNonRotating && !isExcluded)
      ? getEmployeeHistoryTail(priorRoster, emp, targetYear, targetMonth, 6)
      : [];
    const empMonthShifts = dateColumns.map(({ isoDate }) => shiftKey(rows[emp]?.[isoDate] ?? 'OFF'));

    // Check boundary history if available (rotating staff only)
    if (empHistory.length > 0) {
      for (let i = empHistory.length - 1; i >= 0; i--) {
        if (!isRestCode(empHistory[i])) currentStreak++;
        else break;
      }
      for (let i = empHistory.length - 1; i >= 0; i--) {
        if (isRestCode(empHistory[i])) currentOffStreak++;
        else break;
      }
    }

    dateColumns.forEach(({ isoDate }, idx) => {
      const s = empMonthShifts[idx];
      if (counts[s as keyof typeof counts] !== undefined) {
        counts[s as keyof typeof counts]++;
      } else {
        counts.OFF++;
      }

      // Working Streak calculation
      if (!isRestCode(s)) {
        currentStreak++;
        maxStreak = Math.max(maxStreak, currentStreak);
        if (currentStreak > 5) {
          empViolations.push(`Consecutive work streak exceeded 5 days on ${isoDate} (${currentStreak} days)`);
          streakViolationCount++;
        }
        currentOffStreak = 0;
      } else {
        currentStreak = 0;
        // Rest period check (max 2 days, rarely 3, never > 3 for rotating staff)
        // Note: ABS is approved medical absence and H8 is approved vacation, neither counted as excessive rest violation; excluded employees are off for month
        if (s === 'H8') {
          currentOffStreak = 0;
        } else if (!isNonRotating && s !== 'ABS' && !isExcluded) {
          currentOffStreak++;
          if (currentOffStreak > 3) {
            empViolations.push(`Rest period exceeded 3 days on ${isoDate} (${currentOffStreak} days)`);
            excessiveRestViolationCount++;
          }
        }
      }

      // Check Fixed Mid Rule adherence (Palade: Mon-Fri MID, weekends & Romanian holidays OFF)
      if (isFixedMid && !isExcluded) {
        const dayOfWeek = new Date(`${isoDate}T00:00:00`).getDay();
        const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
        const isOff = isWeekend || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
        if (isOff && s !== 'OFF') {
          empViolations.push(`Fixed Manager ${emp} scheduled on ${s} instead of OFF on weekend/holiday (${isoDate})`);
          fixedStaffViolationCount++;
        } else if (!isOff && s !== 'MID' && s !== 'ABS' && s !== 'H8') {
          empViolations.push(`Fixed Manager ${emp} scheduled on ${s} instead of MID on ${isoDate}`);
          fixedStaffViolationCount++;
        }
      }

      // Check Weekday M/A Rule adherence (Tir: Mon-Fri M or A, weekends & Romanian holidays OFF; parity M/A kept permissive for ABS readjustments)
      if (isFixedWeekdayMA && !isExcluded) {
        const dayOfWeek = new Date(`${isoDate}T00:00:00`).getDay();
        const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
        const isOff = isWeekend || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
        if (isOff && s !== 'OFF') {
          empViolations.push(`Weekday staff ${emp} scheduled on ${s} instead of OFF on weekend/holiday (${isoDate})`);
          fixedStaffViolationCount++;
        } else if (!isOff && s !== 'M' && s !== 'A' && s !== 'ABS' && s !== 'H8') {
          empViolations.push(`Weekday staff ${emp} scheduled on ${s} instead of M/A on ${isoDate}`);
          fixedStaffViolationCount++;
        }
      }
    });

    // Night turnaround rest check (48h / >=2 rest days after N across boundary & in-month)
    if (!isExcluded) {
      const checkSeq = [...empHistory.slice(-2), ...empMonthShifts];
      const histOffset = Math.min(2, empHistory.length);
      for (let i = 0; i < checkSeq.length; i++) {
        if (checkSeq[i] !== 'N') continue;
        if (i + 1 >= histOffset && i + 1 < checkSeq.length) {
          const next = checkSeq[i + 1];
          if (next === 'M' || next === 'A' || next === 'MID') {
            const nextIso = dateColumns[i + 1 - histOffset].isoDate;
            empViolations.push(`Night shift followed directly by ${next} on ${nextIso} without rest`);
            nightRestViolationCount++;
          }
        }
        if (i + 2 >= histOffset && i + 2 < checkSeq.length) {
          const next = checkSeq[i + 1];
          const next2 = checkSeq[i + 2];
          if (isRestCode(next) && (next2 === 'M' || next2 === 'A' || next2 === 'MID' || next2 === 'N')) {
            const next2Iso = dateColumns[i + 2 - histOffset].isoDate;
            empViolations.push(`Night shift followed by only 1 rest day before ${next2} on ${next2Iso} (minimum 2 rest days required)`);
            nightRestViolationCount++;
          }
        }
      }
    }

    // Rotating block structure, internal min length (>= 3), and A -> M -> N -> A sequence checks
    if (!isNonRotating && !isExcluded) {
      const hasHistory = empHistory.length > 0;
      const activeTailWork = getActiveTailBlock(empHistory);
      const activeTailRest = activeTailWork.length === 0 ? getActiveTailRest(empHistory) : [];
      const activeTail = activeTailWork.length > 0 ? activeTailWork : activeTailRest;

      interface ValBlock {
        type: 'WORK' | 'REST';
        shift: string;
        len: number;
        startCol: number;
      }
      const blocks: ValBlock[] = [];
      const appendToBlocks = (code: string, colIdx: number) => {
        const r = isRestCode(code);
        const t: 'WORK' | 'REST' = r ? 'REST' : 'WORK';
        const sh = r ? 'OFF' : code;
        const cur = blocks[blocks.length - 1];
        if (cur && cur.type === t && (t === 'REST' || cur.shift === sh)) {
          cur.len++;
        } else {
          blocks.push({ type: t, shift: sh, len: 1, startCol: colIdx });
        }
      };

      activeTail.forEach((s) => appendToBlocks(s, -1));
      empMonthShifts.forEach((s, idx) => appendToBlocks(s, idx));

      // Adjacent WORK blocks without rest days (mixed Day/Night or back-to-back different work shifts)
      for (let bi = 0; bi < blocks.length - 1; bi++) {
        if (blocks[bi].type === 'WORK' && blocks[bi + 1].type === 'WORK') {
          const nextStart = Math.max(0, blocks[bi + 1].startCol);
          const isoDate = dateColumns[nextStart]?.isoDate ?? '';
          empViolations.push(
            `Mixed shift block on ${isoDate}: cannot combine ${blocks[bi].shift} and ${blocks[bi + 1].shift} in the same block without rest days`
          );
          mixedBlockViolationCount++;
        }
      }

      // Internal work blocks must be 3..5 days (exempt Day 1 without prior history and trailing month-end block)
      for (let bi = 0; bi < blocks.length; bi++) {
        const b = blocks[bi];
        if (b.type === 'WORK') {
          const isUnboundedStart = bi === 0 && !hasHistory;
          if (!isUnboundedStart && bi < blocks.length - 1 && b.len < 3) {
            const isoDate = dateColumns[Math.max(0, b.startCol)]?.isoDate ?? '';
            empViolations.push(
              `Short work block (${b.shift}, ${b.len} day${b.len === 1 ? '' : 's'}) starting ${isoDate}: internal blocks require 3-5 consecutive days`
            );
            streakViolationCount++;
          }
        }
      }

      // Rotating sequence A -> M -> N -> A across boundary and in-month work blocks
      const EXPECTED_NEXT_SHIFT: Record<string, string> = { A: 'M', M: 'N', N: 'A' };
      const workBlocks = blocks.filter((b) => b.type === 'WORK');
      let lastHistoryWorkShift: string | null = null;
      for (let hi = empHistory.length - 1; hi >= 0; hi--) {
        if (!isRestCode(empHistory[hi])) {
          lastHistoryWorkShift = empHistory[hi];
          break;
        }
      }

      if (activeTailWork.length === 0 && lastHistoryWorkShift && workBlocks.length > 0) {
        const expectedFirst = EXPECTED_NEXT_SHIFT[lastHistoryWorkShift];
        const firstBlock = workBlocks[0];
        if ((expectedFirst && firstBlock.shift !== expectedFirst) || firstBlock.shift === lastHistoryWorkShift) {
          const isoDate = dateColumns[Math.max(0, firstBlock.startCol)]?.isoDate ?? '';
          empViolations.push(
            `Shift sequence violation on ${isoDate}: expected ${expectedFirst ?? 'next shift'} after ${lastHistoryWorkShift}, got ${firstBlock.shift}`
          );
          mixedBlockViolationCount++;
        }
      }

      for (let wi = 0; wi < workBlocks.length - 1; wi++) {
        const cur = workBlocks[wi];
        const nxt = workBlocks[wi + 1];
        const expectedNext = EXPECTED_NEXT_SHIFT[cur.shift];
        if (cur.shift === nxt.shift || (expectedNext && nxt.shift !== expectedNext)) {
          const isoDate = dateColumns[Math.max(0, nxt.startCol)]?.isoDate ?? '';
          empViolations.push(
            `Shift sequence violation on ${isoDate}: expected ${expectedNext ?? 'next shift'} after ${cur.shift}, got ${nxt.shift}`
          );
          mixedBlockViolationCount++;
        }
      }
    }

    const totalWorkingDays = counts.M + counts.A + counts.N + counts.MID;
    let baseTarget = targetWorkingDays ?? countLegalWorkingDaysInMonth(targetMonth ?? 11, targetYear ?? 2026);
    if (isNonRotating) {
      // Fixed weekday staff (Palade & Tir George) follow standard Mon-Fri minus statutory Romanian holidays
      baseTarget = dateColumns.filter(({ date, isoDate }) => {
        const dow = date.getDay();
        const isWeekend = dow === 0 || dow === 6;
        return !isWeekend && !ROMANIAN_HOLIDAYS.has(isoDate) && !FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
      }).length;
    }
    if (isExcluded) {
      baseTarget = 0;
    }
    // Medical leave (ABS) and approved vacation (H8) reduce the monthly working shift norm accordingly
    const effectiveTarget = Math.max(0, baseTarget - counts.ABS - counts.H8);
    const workDaysDiff = isExcluded ? 0 : totalWorkingDays - effectiveTarget;

    // Monthly norm constraint: all employees must meet their effective monthly norm
    if (!isExcluded && targetWorkingDays !== undefined && workDaysDiff !== 0) {
      empViolations.push(
        `${emp}: worked ${totalWorkingDays} days (target is ${effectiveTarget} days${counts.ABS > 0 ? ` after ${counts.ABS}d ABS` : ''})`
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
      targetWorkingDays: isExcluded ? 0 : effectiveTarget,
      workDaysDiff,
      maxStreak,
      violations: isExcluded ? [] : empViolations
    };

    if (!isExcluded && empViolations.length > 0) {
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
    warnings
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
  const employees = roster.employees.length > 0 ? roster.employees : Object.keys(roster.rows);

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
  const targetWorkingDays = countLegalWorkingDaysInMonth(roster.month, roster.year);

  // 5. Employee Rows
  const employees = roster.employees.length > 0 ? roster.employees : Object.keys(roster.rows);
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
    { label: 'MORNING 6-15', code: 'M', offset: 3, targetMin: 1 },
    { label: 'AFTER 13-22', code: 'A', offset: 4, targetMin: 1 },
    { label: 'NIGHT 21-06', code: 'N', offset: 5, targetMin: 1 },
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

      const dayOfWeek = new Date(col.isoDate).getDay();
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
 * 1. Minimum rotating staffing per day: M >= 2, A >= 2, N = 2.
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

  // Penalty evaluation function using validateScheduleCompliance
  const evaluatePenalty = (testRows: Record<string, Record<string, string>>) => {
    const rep = validateScheduleCompliance(
      testRows,
      dateColumns,
      priorRoster,
      targetYear,
      targetMonth,
      targetWorkingDays
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
    excludeIsoDate: string
  ): { offsetIso: string; oldShift: string } | null => {
    for (let k = 0; k < dateColumns.length; k++) {
      const kIso = dateColumns[k].isoDate;
      if (kIso === excludeIsoDate) continue;
      const curS = shiftKey(baseRows[emp]?.[kIso] ?? 'OFF');
      if (curS !== 'M' && curS !== 'A') continue;

      let count = 0;
      for (const re of rotatingEmployees) {
        if (shiftKey(baseRows[re]?.[kIso] ?? 'OFF') === curS) count++;
      }
      const minStaff = curS === 'M' ? 2 : 3;
      if (count >= minStaff) {
        const tempRows = cloneRows(baseRows);
        tempRows[emp][kIso] = 'OFF';
        const { rep } = evaluatePenalty(tempRows);
        if (!rep.warnings.some((w) => w.startsWith(`${kIso}:`))) {
          return { offsetIso: kIso, oldShift: curS };
        }
      }
    }
    return null;
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
        const offset = findOffsetDayForEmp(workingRows, emp, isoDate);
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
        if (nextIso && getShift(workingRows, emp, nextIso) !== 'OFF') {
          changes3.push({
            isoDate: nextIso,
            employee: emp,
            fromShift: getShift(workingRows, emp, nextIso),
            toShift: 'OFF',
            reason: `Mandatory post-night rest day 1.`,
            absColleague: absNames
          });
        }
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

      // Category 2: George Tir (if weekday & not holiday)
      if (FIXED_WEEKDAY_MA_EMPLOYEES.length > 0) {
        const dayOfWeek = col.date.getDay();
        const isWk = dayOfWeek === 0 || dayOfWeek === 6;
        if (!isWk) {
          for (const gt of FIXED_WEEKDAY_MA_EMPLOYEES) {
            if (getShift(workingRows, gt, isoDate) === 'A') {
              const changes: CandidatePlan['changes'] = [
                {
                  isoDate,
                  employee: gt,
                  fromShift: 'A',
                  toShift: 'M',
                  reason: `Fixed weekday M/A colleague switched from Afternoon to Morning to cover sick colleague.`,
                  absColleague: absNames
                }
              ];
              const test = cloneRows(workingRows);
              test[gt][isoDate] = 'M';
              const { penalty, rep } = evaluatePenalty(test);
              plans.push({ changes, penalty, violations: rep.totalViolations });
            }
          }
        }
      }

      // Category 3: OFF staff
      const eligibleOff = offStaff.filter((emp) => getShift(workingRows, emp, prevIso) !== 'N');
      eligibleOff.sort((a, b) => {
        if (a.includes('Urzica')) return -1;
        if (b.includes('Urzica')) return 1;
        return a.localeCompare(b);
      });

      for (const emp of eligibleOff) {
        const offset = findOffsetDayForEmp(workingRows, emp, isoDate);
        const changes: CandidatePlan['changes'] = [
          {
            isoDate,
            employee: emp,
            fromShift: 'OFF',
            toShift: 'M',
            reason: `Assigned from rest to Morning shift (OFF ➔ M) to ensure minimum of 2 Morning staff.`,
            absColleague: absNames
          }
        ];
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
        let chosen: CandidatePlan;

        const prisecaruPlan = plans.find((p) => p.changes.some((c) => c.employee.includes('Prisecaru')));
        const lucianPlan = plans.find((p) => p.changes.some((c) => c.employee.includes('Olteanu')));
        const urzicaPlan = plans.find((p) => p.changes.some((c) => c.employee.includes('Urzica')));

        // Strategic variant assignment:
        if (variantSeed === 0) {
          // Variant 1: Give Prisecaru the morning shift on these 2 days as he is on after
          chosen = prisecaruPlan ?? lucianPlan ?? plans[0];
        } else if (variantSeed === 1) {
          // Variant 2: Urzica 1 morning and Lucian 1 morning
          if (dayInBlock === 0) {
            chosen = urzicaPlan ?? prisecaruPlan ?? plans[0];
          } else {
            chosen = lucianPlan ?? prisecaruPlan ?? plans[0];
          }
        } else if (variantSeed === 2) {
          // Variant 3: Lucian covers both days
          chosen = lucianPlan ?? prisecaruPlan ?? plans[0];
        } else {
          // Generic cycling across top variants
          plans.sort((a, b) => a.penalty - b.penalty);
          const bestPenalty = plans[0].penalty;
          const topPlans = plans.filter((p) => p.penalty === bestPenalty);
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

      // Category 2: George Tir (if weekday & not holiday)
      if (FIXED_WEEKDAY_MA_EMPLOYEES.length > 0) {
        const dayOfWeek = col.date.getDay();
        const isWk = dayOfWeek === 0 || dayOfWeek === 6;
        if (!isWk) {
          for (const gt of FIXED_WEEKDAY_MA_EMPLOYEES) {
            if (getShift(workingRows, gt, isoDate) === 'M') {
              const changes: CandidatePlan['changes'] = [
                {
                  isoDate,
                  employee: gt,
                  fromShift: 'M',
                  toShift: 'A',
                  reason: `Fixed weekday M/A colleague switched from Morning to Afternoon to cover sick colleague.`,
                  absColleague: absNames
                }
              ];
              const test = cloneRows(workingRows);
              test[gt][isoDate] = 'A';
              const { penalty, rep } = evaluatePenalty(test);
              plans.push({ changes, penalty, violations: rep.totalViolations });
            }
          }
        }
      }

      // Category 3: OFF staff
      const eligibleOff = offStaff.filter((emp) => getShift(workingRows, emp, prevIso) !== 'N');
      for (const emp of eligibleOff) {
        const offset = findOffsetDayForEmp(workingRows, emp, isoDate);
        const changes: CandidatePlan['changes'] = [
          {
            isoDate,
            employee: emp,
            fromShift: 'OFF',
            toShift: 'A',
            reason: `Assigned from rest to Afternoon shift (OFF ➔ A) to maintain minimum of 2 Afternoon staff.`,
            absColleague: absNames
          }
        ];
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
  // 4. GLOBAL REBALANCING POST-PASS (Eliminate any remaining warnings)
  // ---------------------------------------------------------------------------
  let currentValidation = validateScheduleCompliance(
    workingRows,
    dateColumns,
    priorRoster,
    targetYear,
    targetMonth,
    targetWorkingDays
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
        for (let k = 0; k < dateColumns.length; k++) {
          const kIso = dateColumns[k].isoDate;
          const curS = shiftKey(workingRows[emp]?.[kIso] ?? 'OFF');
          if (curS !== 'M' && curS !== 'A') continue;

          // Check if shift has >= 3 staff
          const curCount = rotatingEmployees.filter(
            (re) => shiftKey(workingRows[re]?.[kIso] ?? 'OFF') === curS
          ).length;

          if (curCount >= 3) {
            const testRows = cloneRows(workingRows);
            testRows[emp][kIso] = 'OFF';
            const testRep = validateScheduleCompliance(
              testRows,
              dateColumns,
              priorRoster,
              targetYear,
              targetMonth,
              targetWorkingDays
            );

            if (testRep.totalViolations < currentValidation.totalViolations) {
              workingRows[emp][kIso] = 'OFF';
              currentValidation = testRep;
              adjustments.push({
                isoDate: kIso,
                employee: emp,
                fromShift: curS,
                toShift: 'OFF',
                reason: `Rebalanced: Added offsetting day off to preserve legal monthly norm of ${targetWorkingDays} working days.`,
                absColleague: 'Rebalancing'
              });
              improved = true;
              break;
            }
          }
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
    validation: currentValidation
  };
}

