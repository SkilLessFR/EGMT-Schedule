// src/scheduleBuilder.ts
import ExcelJS from 'exceljs';
import type { RosterData } from './types';
import { shiftKey } from './scheduleUtils';

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

export const ROMANIAN_HOLIDAYS_2026 = new Set<string>([
  '2026-01-01', '2026-01-02', '2026-01-06', '2026-01-07', '2026-01-24',
  '2026-04-10', '2026-04-12', '2026-04-13', '2026-05-01', '2026-05-31',
  '2026-06-01', '2026-08-15', '2026-11-30', '2026-12-01', '2026-12-25', '2026-12-26'
]);

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
  11: 21  // Dec (21 working days in 2026: 23 weekdays minus 2 statutory public holidays: Dec 1 & Dec 25)
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
    count += 1;
  }
  return count;
}

/**
 * Checks whether any block of consecutive working shifts in a sequence contains multiple shift types.
 * Every continuous block between OFF days must consist strictly of a single shift type.
 */
export function hasMixedBlocks(sequence: string[]): boolean {
  let curType: string | null = null;
  for (const s of sequence) {
    const k = s === 'OFF' || s === 'H8' || s === 'HUP' || s === 'ABS' || s === 'AD' ? 'OFF' : s;
    if (k !== 'OFF') {
      if (curType === null) curType = k;
      else if (curType !== k) return true;
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
}

export interface DayCoverage {
  isoDate: string;
  dayOfWeek: number; // 0 = Sun, 1 = Mon ...
  isWeekend: boolean;
  M: number;
  A: number;
  N: number;
  MID: number;
  OFF: number;
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
  dailyCoverage: Record<string, DayCoverage>;
  employeeStats: Record<string, EmployeeMonthStats>;
  warnings: string[];
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
function evaluateOffsetCost(
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

  // 5. Anti-repeat check: An employee must NOT repeat the same shift block in a row after rest!
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
      if (firstWorkShift === lastWorkShift) {
        cost += 5000000; // STRICTLY FORBIDDEN: 2 blocks of the same shift in a row!
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
function equalizeExactWorkingDays(
  rows: Record<string, Record<string, string>>,
  dateColumns: { index: number; date: Date; isoDate: string }[],
  targetWorkingDays: number,
  priorRoster: RosterData | null,
  year: number,
  month: number
) {
  const allEmployees = Object.keys(rows);
  const fixedMid = FIXED_MID_EMPLOYEES.filter((e) => allEmployees.includes(e));
  const fixedWeekdayMA = FIXED_WEEKDAY_MA_EMPLOYEES.filter((e) => allEmployees.includes(e));
  const rotating = allEmployees.filter((e) => !fixedMid.includes(e) && !fixedWeekdayMA.includes(e));

  // 1. Fixed Managers (Palade): Monday to Friday MID, Saturday and Sunday OFF, Romanian National Holidays OFF
  fixedMid.forEach((emp) => {
    dateColumns.forEach(({ date, isoDate }) => {
      const dow = date.getDay();
      const isWeekend = dow === 0 || dow === 6;
      const isHoliday = ROMANIAN_HOLIDAYS_2026.has(isoDate);
      if (isWeekend || isHoliday) {
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
      const isHoliday = ROMANIAN_HOLIDAYS_2026.has(isoDate);
      if (isWeekend || isHoliday) {
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
      ? getEmployeeHistoryTail(priorRoster, emp, year, month, 5)
      : [];
  });

  const isEmpValid = (emp: string) => {
    const empHistory = employeeHistories[emp] ?? [];
    const empMonthShifts = dateColumns.map((col) => shiftKey(rows[emp]?.[col.isoDate] ?? 'OFF'));

    // 1. Mixed blocks check
    const activeTailWork = getActiveTailBlock(empHistory);
    const activeTailRest = activeTailWork.length === 0 ? getActiveTailRest(empHistory) : [];
    const activeTail = activeTailWork.length > 0 ? activeTailWork : activeTailRest;
    if (hasMixedBlocks([...activeTailWork, ...empMonthShifts])) return false;

    // 2. Group into continuous blocks of work and rest
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

    // 3. Anti-repeat check: NO consecutive same-shift work blocks!
    const workBlocks = blocks.filter((b) => b.type === 'WORK');
    for (let i = 0; i < workBlocks.length - 1; i++) {
      if (workBlocks[i].shift === workBlocks[i + 1].shift) {
        return false;
      }
    }

    // 4. Work and rest streaks
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.type === 'WORK') {
        if (b.len > 5) return false;
      } else {
        // Rest period between two work blocks:
        // After Night shift: strictly >= 2 days (48h mandatory rest)
        // Max rest: at most 3 days
        if (i > 0 && i < blocks.length - 1) {
          const prevWork = blocks[i - 1];
          if (prevWork && prevWork.shift === 'N' && b.len < 2) return false;
          if (b.len > 3) return false;
        }
        // Rest bridging across month boundary must not exceed 3 days
        if (i === 0 && activeTailRest.length > 0 && b.len > activeTailRest.length && b.len > 3) {
          return false;
        }
      }
    }

    // 5. Night turnaround check: night shift followed by at least 2 days of rest before M, A, or MID
    for (let i = 0; i < empMonthShifts.length; i++) {
      if (empMonthShifts[i] === 'N') {
        if (i + 1 < empMonthShifts.length) {
          const next = empMonthShifts[i + 1];
          if (next === 'M' || next === 'A' || next === 'MID') return false;
          if (next === 'OFF' && i + 2 < empMonthShifts.length) {
            const next2 = empMonthShifts[i + 2];
            if (next2 === 'M' || next2 === 'A' || next2 === 'MID') return false;
          }
        }
      }
    }

    return true;
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

  // Phase B: If any deficit remains, add shifts on days where staffing is low (matching adjacent block type)
  // Prioritizes additions that maintain balanced 1:1:1 ratio across M, A, and N shifts
  for (let iter = 0; iter < 25; iter++) {
    const deficit = rotating.filter((e) => getWorkDays(e) < targetWorkingDays);
    if (deficit.length === 0) break;

    let added = false;
    for (const dEmp of deficit) {
      let bestMove: { isoDate: string; cand: 'M' | 'A' | 'N'; spread: number } | null = null;

      const counts: Record<string, number> = { M: 0, A: 0, N: 0 };
      dateColumns.forEach(({ isoDate: dIso }) => {
        const s = rows[dEmp][dIso];
        if (s === 'M' || s === 'A' || s === 'N') counts[s]++;
      });

      for (let i = 0; i < dateColumns.length; i++) {
        const isoDate = dateColumns[i].isoDate;
        if (rows[dEmp][isoDate] === 'OFF') {
          const prevShift = i > 0
            ? shiftKey(rows[dEmp][dateColumns[i - 1].isoDate] ?? 'OFF')
            : (employeeHistories[dEmp]?.[employeeHistories[dEmp].length - 1] ?? 'OFF');
          const nextShift = i < dateColumns.length - 1
            ? shiftKey(rows[dEmp][dateColumns[i + 1].isoDate] ?? 'OFF')
            : 'OFF';

          const candidates: ('M' | 'A' | 'N')[] = [];
          if (prevShift === 'M' || prevShift === 'A' || prevShift === 'N') candidates.push(prevShift);
          if (nextShift === 'M' || nextShift === 'A' || nextShift === 'N') candidates.push(nextShift);
          if (candidates.length === 0) candidates.push('M', 'A');

          for (const cand of candidates) {
            if ((prevShift !== 'OFF' && prevShift !== cand) || (nextShift !== 'OFF' && nextShift !== cand)) {
              continue;
            }
            if (cand === 'N') {
              let currentN = 0;
              rotating.forEach((e) => { if (rows[e][isoDate] === 'N') currentN++; });
              if (currentN >= 2) continue;
            }
            rows[dEmp][isoDate] = cand;
            if (isEmpValid(dEmp)) {
              const newCounts = { ...counts, [cand]: (counts[cand] ?? 0) + 1 };
              const spread = Math.max(newCounts.M, newCounts.A) - Math.min(newCounts.M, newCounts.A);
              if (!bestMove || spread < bestMove.spread) {
                bestMove = { isoDate, cand, spread };
              }
            }
            rows[dEmp][isoDate] = 'OFF';
          }
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

  // Phase C: If any surplus remains, drop shift to OFF on days with adequate coverage
  for (let iter = 0; iter < 25; iter++) {
    const surplus = rotating.filter((e) => getWorkDays(e) > targetWorkingDays);
    if (surplus.length === 0) break;

    let dropped = false;
    for (const sEmp of surplus) {
      for (let i = 0; i < dateColumns.length; i++) {
        const isoDate = dateColumns[i].isoDate;
        const cur = rows[sEmp][isoDate];
        if (cur === 'M' || cur === 'A' || cur === 'N') {
          let count = 0;
          rotating.forEach((e) => {
            if (rows[e][isoDate] === cur) count++;
          });
          if (count > 2) {
            rows[sEmp][isoDate] = 'OFF';
            if (isEmpValid(sEmp)) {
              dropped = true;
              break;
            } else {
              rows[sEmp][isoDate] = cur;
            }
          }
        }
      }
      if (dropped) break;
    }
    if (!dropped) break;
  }

  // Phase NightCap: Strict hard ceiling of maximum 2 people on Night shift (N = 2) on all days
  for (let iter = 0; iter < 25; iter++) {
    const overDays = dateColumns.filter(({ isoDate }) => {
      let n = 0;
      rotating.forEach((e) => { if (rows[e][isoDate] === 'N') n++; });
      return n > 2;
    });

    if (overDays.length === 0) break;

    let reduced = false;
    for (const { isoDate } of overDays) {
      const empsOnN = rotating
        .filter((e) => rows[e][isoDate] === 'N')
        .sort((a, b) => getWorkDays(b) - getWorkDays(a));
      for (const emp of empsOnN) {
        rows[emp][isoDate] = 'OFF';
        if (isEmpValid(emp)) {
          reduced = true;
          break;
        }
        rows[emp][isoDate] = 'N';
      }
      if (reduced) break;
    }
    if (!reduced) break;
  }

  // Phase NightCap Compensation: For any employee who has workDays < targetWorkingDays, add M or A
  for (let iter = 0; iter < 25; iter++) {
    const deficit = rotating.filter((e) => getWorkDays(e) < targetWorkingDays);
    if (deficit.length === 0) break;

    let added = false;
    for (const dEmp of deficit) {
      for (let i = 0; i < dateColumns.length; i++) {
        const isoDate = dateColumns[i].isoDate;
        if (rows[dEmp][isoDate] === 'OFF') {
          const prev = i > 0 ? rows[dEmp][dateColumns[i - 1].isoDate] : 'OFF';
          const next = i < dateColumns.length - 1 ? rows[dEmp][dateColumns[i + 1].isoDate] : 'OFF';

          const cands: ('M' | 'A' | 'N')[] = [];
          if (prev === 'M' || prev === 'A' || prev === 'N') cands.push(prev);
          if (next === 'M' || next === 'A' || next === 'N') cands.push(next);
          if (cands.length === 0) cands.push('M', 'A');

          for (const cand of cands) {
            if (cand === 'N') {
              let currentN = 0;
              rotating.forEach((e) => { if (rows[e][isoDate] === 'N') currentN++; });
              if (currentN >= 2) continue;
            }
            rows[dEmp][isoDate] = cand;
            if (isEmpValid(dEmp)) {
              added = true;
              break;
            }
            rows[dEmp][isoDate] = 'OFF';
          }
          if (added) break;
        }
      }
      if (added) break;
    }
    if (!added) break;
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

  const fixedMidEmployees = FIXED_MID_EMPLOYEES.filter((name) => allEmployees.includes(name));
  const fixedWeekdayMAEmployees = FIXED_WEEKDAY_MA_EMPLOYEES.filter((name) => allEmployees.includes(name));
  const rotatingEmployees = allEmployees.filter((name) => !fixedMidEmployees.includes(name) && !fixedWeekdayMAEmployees.includes(name));

  // Balanced 1:1:1 templates with 4-day blocks and 2-day rests across 18-day coprime cycles
  // Templates rotate in all 6 possible cyclic directions (M->A->N, A->N->M, N->M->A, M->N->A, A->M->N, N->A->M)
  // Perfectly balanced: 4 M, 4 A, 4 N, and 6 OFF (ratio 1:1:1:1.5)
  const BALANCED_TEMPLATES = [
    // Template 0: M -> A -> N
    [
      'M','M','M','M',
      'OFF','OFF',
      'A','A','A','A',
      'OFF','OFF',
      'N','N','N','N',
      'OFF','OFF'
    ],
    // Template 1: A -> N -> M
    [
      'A','A','A','A',
      'OFF','OFF',
      'N','N','N','N',
      'OFF','OFF',
      'M','M','M','M',
      'OFF','OFF'
    ],
    // Template 2: N -> M -> A
    [
      'N','N','N','N',
      'OFF','OFF',
      'M','M','M','M',
      'OFF','OFF',
      'A','A','A','A',
      'OFF','OFF'
    ],
    // Template 3: M -> N -> A
    [
      'M','M','M','M',
      'OFF','OFF',
      'N','N','N','N',
      'OFF','OFF',
      'A','A','A','A',
      'OFF','OFF'
    ],
    // Template 4: A -> M -> N
    [
      'A','A','A','A',
      'OFF','OFF',
      'M','M','M','M',
      'OFF','OFF',
      'N','N','N','N',
      'OFF','OFF'
    ],
    // Template 5: N -> A -> M
    [
      'N','N','N','N',
      'OFF','OFF',
      'A','A','A','A',
      'OFF','OFF',
      'M','M','M','M',
      'OFF','OFF'
    ]
  ];

  // Pre-configured slot sets ensuring 100% daily coverage (M >= 2, A >= 2, N = 2 every day)
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

  const VALID_SLOT_SETS_9 = [
    [0, 2, 4, 6, 8, 10, 12, 14, 16],
    [2, 4, 6, 8, 10, 12, 14, 16, 0],
    [4, 6, 8, 10, 12, 14, 16, 0, 2],
    [6, 8, 10, 12, 14, 16, 0, 2, 4],
    [8, 10, 12, 14, 16, 0, 2, 4, 6],
    [10, 12, 14, 16, 0, 2, 4, 6, 8],
    [12, 14, 16, 0, 2, 4, 6, 8, 10],
    [14, 16, 0, 2, 4, 6, 8, 10, 12],
    [16, 0, 2, 4, 6, 8, 10, 12, 14],
    [1, 3, 5, 7, 9, 11, 13, 15, 17]
  ];

  const VALID_SLOT_SETS = rotatingEmployees.length === 9
    ? VALID_SLOT_SETS_9
    : VALID_SLOT_SETS_11;

  let bestResultRows: Record<string, Record<string, string>> = {};
  let bestReport: ScheduleValidationReport | null = null;
  let finalSeed = seed;

  interface CandidateConfig {
    template: string[];
    slots: number[];
    matching: number[];
    cost: number;
    attemptSeed: number;
  }

  const rotatingHistories: Record<string, string[]> = {};
  const rotatingAnchors: Record<string, string | null> = {};
  rotatingEmployees.forEach((emp) => {
    const hist = getEmployeeHistoryTail(roster, emp, targetYear, targetMonth, 6);
    rotatingHistories[emp] = hist;
    rotatingAnchors[emp] = hist.length > 0 ? roster?.rows?.[emp]?.[targetStartIso] ?? null : null;
  });

  const SEED_OFFSETS = [0, 1, 3, 5, 7, 10, 100];

  for (const sOffset of SEED_OFFSETS) {
    const effSeed = seed + sOffset;
    const allCandidates: CandidateConfig[] = [];

    // Stage 1: Fast boundary screening across all 96 template and slot configurations (< 10ms)
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
        let bestCost = 0;
        for (let eIdx = 0; eIdx < rotatingEmployees.length; eIdx++) {
          bestCost += costMatrix[eIdx][matching[eIdx]];
        }

        allCandidates.push({
          template,
          slots,
          matching,
          cost: bestCost,
          attemptSeed: currentAttemptSeed
        });
      }
    }

    // Sort candidates by boundary cost ascending (strictly legal combinations first)
    allCandidates.sort((a, b) => a.cost - b.cost);

    // Stage 2: Equalize and validate candidate configurations (early exit on fully compliant)
    const topConfigs = allCandidates.slice(0, 96);

    for (const config of topConfigs) {
      const candidateRows: Record<string, Record<string, string>> = {};
      allEmployees.forEach((emp) => { candidateRows[emp] = {}; });

      fixedMidEmployees.forEach((emp) => {
        dateColumns.forEach(({ date, isoDate }) => {
          const dow = date.getDay();
          const isWeekend = dow === 0 || dow === 6;
          const isHoliday = ROMANIAN_HOLIDAYS_2026.has(isoDate);
          if (isWeekend || isHoliday) {
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
          const isHoliday = ROMANIAN_HOLIDAYS_2026.has(isoDate);
          if (isWeekend || isHoliday) {
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

      equalizeExactWorkingDays(candidateRows, dateColumns, targetWorkingDays, roster, targetYear, targetMonth);
      const testReport = validateScheduleCompliance(
        candidateRows,
        dateColumns,
        roster,
        targetYear,
        targetMonth,
        targetWorkingDays
      );

      if (testReport.isFullyCompliant || !bestReport || testReport.totalViolations < bestReport.totalViolations) {
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
    targetWorkingDays
  );

  return {
    month: targetMonth,
    year: targetYear,
    daysCount,
    targetWorkingDays,
    dateColumns,
    rows: bestResultRows,
    validation,
    seed: finalSeed
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
  targetWorkingDays?: number
): ScheduleValidationReport {
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

  const employees = Object.keys(rows);

  // 1. Check Daily Coverage
  dateColumns.forEach(({ date, isoDate }) => {
    const dayOfWeek = date.getDay();
    const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
    const c = { M: 0, A: 0, N: 0, MID: 0, OFF: 0, H8: 0, other: 0 };
    const dayWarnings: string[] = [];

    employees.forEach((emp) => {
      const s = shiftKey(rows[emp]?.[isoDate] ?? 'OFF');
      if (s === 'M') c.M++;
      else if (s === 'A') c.A++;
      else if (s === 'N') c.N++;
      else if (s === 'MID') c.MID++;
      else if (s === 'H8') c.H8++;
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

    const isUnderstaffed = rotC.M < 2 || rotC.A < 2 || rotC.N < 2;
    const isOverstaffedNight = c.N > 2;
    if (rotC.M < 2) dayWarnings.push(`Morning understaffed (${rotC.M}/2 minimum from rotating staff)`);
    if (rotC.A < 2) dayWarnings.push(`Afternoon understaffed (${rotC.A}/2 minimum from rotating staff)`);
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
      A: c.A,
      N: c.N,
      MID: c.MID,
      OFF: c.OFF + c.H8 + c.other,
      totalWorking: c.M + c.A + c.N + c.MID,
      isUnderstaffed: isUnderstaffed || isOverstaffedNight,
      warnings: dayWarnings
    };
  });

  // 2. Check Employee Legal Constraints, Streaks & Rest Durations
  employees.forEach((emp) => {
    const isFixedMid = FIXED_MID_EMPLOYEES.includes(emp);
    const isFixedWeekdayMA = FIXED_WEEKDAY_MA_EMPLOYEES.includes(emp);
    const isNonRotating = isFixedMid || isFixedWeekdayMA;
    const empViolations: string[] = [];
    const counts = { M: 0, A: 0, N: 0, MID: 0, OFF: 0, H8: 0 };

    let currentStreak = 0;
    let maxStreak = 0;
    let currentOffStreak = 0;

    // Check boundary history if available (rotating staff only)
    if (priorRoster && targetYear !== undefined && targetMonth !== undefined && !isNonRotating) {
      const history = getEmployeeHistoryTail(priorRoster, emp, targetYear, targetMonth, 5);
      for (let i = history.length - 1; i >= 0; i--) {
        if (history[i] !== 'OFF' && history[i] !== 'H8' && history[i] !== 'HUP') currentStreak++;
        else break;
      }
      for (let i = history.length - 1; i >= 0; i--) {
        if (history[i] === 'OFF' || history[i] === 'H8' || history[i] === 'HUP') currentOffStreak++;
        else break;
      }
    }

    dateColumns.forEach(({ isoDate }, idx) => {
      const s = shiftKey(rows[emp]?.[isoDate] ?? 'OFF');
      if (counts[s as keyof typeof counts] !== undefined) {
        counts[s as keyof typeof counts]++;
      } else {
        counts.OFF++;
      }

      // Working Streak calculation
      if (s !== 'OFF' && s !== 'H8' && s !== 'HUP') {
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
        if (!isNonRotating) {
          currentOffStreak++;
          if (currentOffStreak > 3) {
            empViolations.push(`Rest period exceeded 3 days on ${isoDate} (${currentOffStreak} days)`);
            excessiveRestViolationCount++;
          }
        }
      }

      // Night rest constraint check
      if (idx < dateColumns.length - 1) {
        const nextIso = dateColumns[idx + 1].isoDate;
        const nextShift = shiftKey(rows[emp]?.[nextIso] ?? 'OFF');
        if (s === 'N' && (nextShift === 'M' || nextShift === 'A' || nextShift === 'MID')) {
          empViolations.push(`Night shift on ${isoDate} followed directly by ${nextShift} on ${nextIso} without rest`);
          nightRestViolationCount++;
        }
      }

      // Check Fixed Mid Rule adherence (Palade: Mon-Fri MID, weekends & Romanian holidays OFF)
      if (isFixedMid) {
        const dayOfWeek = new Date(`${isoDate}T00:00:00`).getDay();
        const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
        const isHoliday = ROMANIAN_HOLIDAYS_2026.has(isoDate);
        if ((isWeekend || isHoliday) && s !== 'OFF') {
          empViolations.push(`Fixed Manager ${emp} scheduled on ${s} instead of OFF on weekend/holiday (${isoDate})`);
        } else if (!isWeekend && !isHoliday && s !== 'MID') {
          empViolations.push(`Fixed Manager ${emp} scheduled on ${s} instead of MID on ${isoDate}`);
        }
      }

      // Check Weekday M/A Rule adherence (Tir: Mon-Fri M or A, weekends & Romanian holidays OFF)
      if (isFixedWeekdayMA) {
        const dayOfWeek = new Date(`${isoDate}T00:00:00`).getDay();
        const isWeekend = dayOfWeek === 0 || dayOfWeek === 6;
        const isHoliday = ROMANIAN_HOLIDAYS_2026.has(isoDate);
        if ((isWeekend || isHoliday) && s !== 'OFF') {
          empViolations.push(`Weekday staff ${emp} scheduled on ${s} instead of OFF on weekend/holiday (${isoDate})`);
        } else if (!isWeekend && !isHoliday && s !== 'M' && s !== 'A') {
          empViolations.push(`Weekday staff ${emp} scheduled on ${s} instead of M/A on ${isoDate}`);
        }
      }
    });

    // Check Pure Blocks (no mixed shift types in any block of consecutive work days for rotating staff)
    if (!isNonRotating) {
      const empHistory = (priorRoster && targetYear !== undefined && targetMonth !== undefined)
        ? getEmployeeHistoryTail(priorRoster, emp, targetYear, targetMonth, 5)
        : [];
      const activeTail = getActiveTailBlock(empHistory);
      const tailType = activeTail.length > 0 ? activeTail[0] : null;

      let currentBlockType: string | null = tailType;
      dateColumns.forEach(({ isoDate }) => {
        const shift = shiftKey(rows[emp]?.[isoDate] ?? 'OFF');
        if (shift !== 'OFF' && shift !== 'H8' && shift !== 'HUP' && shift !== 'ABS' && shift !== 'AD') {
          if (currentBlockType === null) {
            currentBlockType = shift;
          } else if (currentBlockType !== shift) {
            empViolations.push(
              `Mixed shift block on ${isoDate}: cannot combine ${currentBlockType} and ${shift} in the same block without rest days`
            );
            mixedBlockViolationCount++;
          }
        } else {
          currentBlockType = null;
        }
      });
    }

    const totalWorkingDays = counts.M + counts.A + counts.N + counts.MID;
    const effectiveTarget = targetWorkingDays ?? countLegalWorkingDaysInMonth(targetMonth ?? 11, targetYear ?? 2026);
    const workDaysDiff = totalWorkingDays - effectiveTarget;

    // Monthly norm constraint: all employees must have the exact same number of working days
    if (targetWorkingDays !== undefined && workDaysDiff !== 0) {
      empViolations.push(
        `${emp}: worked ${totalWorkingDays} days (monthly norm is ${targetWorkingDays} days)`
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
      totalWorkingDays,
      targetWorkingDays: effectiveTarget,
      workDaysDiff,
      maxStreak,
      violations: empViolations
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
    workDaysMismatchCount;

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
    const isHoliday = ROMANIAN_HOLIDAYS_2026.has(col.isoDate);
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
          // Absence: Purple background, white text
          shiftCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF7030A0' } };
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
      const isHoliday = ROMANIAN_HOLIDAYS_2026.has(col.isoDate);
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
          fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: 'FF7030A0' } },
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
