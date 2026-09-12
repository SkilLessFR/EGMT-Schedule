// scheduleUtils.ts
import type { RosterData, ShiftColor, ShiftEvent } from './types';

// ---------------------------------------------------------------------------
// Shift colors & labels — shared by the calendar grid and the day details modal
// ---------------------------------------------------------------------------

export const shiftColors: Record<string, ShiftColor> = {
  MID: { bg: 'bg-blue-500/15 dark:bg-blue-400/20', text: 'text-blue-600 dark:text-blue-300', border: 'border-transparent' },
  A: { bg: 'bg-orange-500/15 dark:bg-orange-400/20', text: 'text-orange-600 dark:text-orange-300', border: 'border-transparent' },
  M: { bg: 'bg-green-500/15 dark:bg-green-400/20', text: 'text-green-600 dark:text-green-300', border: 'border-transparent' },
  N: { bg: 'bg-purple-500/15 dark:bg-purple-400/20', text: 'text-purple-600 dark:text-purple-300', border: 'border-transparent' },
  OFF: { bg: 'bg-transparent', text: 'text-zinc-400 dark:text-zinc-500', border: 'border-transparent' },
  H8: { bg: 'bg-pink-500/15 dark:bg-pink-400/20', text: 'text-pink-600 dark:text-pink-300', border: 'border-transparent' },
};
export const fallbackColor: ShiftColor = { bg: 'bg-zinc-500/15 dark:bg-zinc-400/20', text: 'text-zinc-600 dark:text-zinc-300', border: 'border-transparent' };

export const shiftSolid: Record<string, string> = {
  MID: 'bg-blue-500 text-white',
  A: 'bg-orange-500 text-white',
  M: 'bg-green-500 text-white',
  N: 'bg-purple-500 text-white',
  OFF: 'bg-zinc-400 text-white',
  H8: 'bg-pink-500 text-white',
};
export const fallbackSolid = 'bg-zinc-500 text-white';

export const hours: Record<string, string> = { MID: '09:00–17:00', M: '06:00–14:00', A: '14:00–22:00', N: '22:00–06:00', H8: 'Holiday', OFF: 'AD' };
export const shiftLabels: Record<string, string> = { M: 'Morning', A: 'Afternoon', N: 'Night', MID: 'Mid', OFF: 'Off', H8: 'H8' };
export const shiftHourValues: Record<string, number> = { M: 8, A: 8, N: 8, MID: 8, OFF: 0, H8: 0 };

const OFF_SHIFT_ALIASES = new Set(['OFF', 'ABS', 'AD']);

export function colorFor(shift: string) { return shiftColors[shift.toUpperCase()] ?? fallbackColor; }
export function solidColorFor(shift: string) { return shiftSolid[shift.toUpperCase()] ?? fallbackSolid; }
export function shiftKey(shift: string) {
  const code = shift.trim().toUpperCase();
  return OFF_SHIFT_ALIASES.has(code) ? 'OFF' : code;
}
export function shiftHours(shift: string) { return hours[shiftKey(shift)] ?? 'Not provided'; }
export function shiftLabel(shift: string) { return shiftLabels[shiftKey(shift)] ?? shift; }
export function initials(name: string) { return name.split(/\s+/).filter(Boolean).slice(0, 2).map((part) => part[0]?.toUpperCase()).join(''); }
export function byName(a: { name: string }, b: { name: string }) { return a.name.localeCompare(b.name); }

// ---------------------------------------------------------------------------
// Daily roster index — who's on which shift, per day (built once per roster load)
// ---------------------------------------------------------------------------

export type DailyRoster = { morning: string[]; afternoon: string[]; night: string[]; mid: string[]; off: string[] };
export type WorkingGroup = { title: string; employees: { name: string; suffix?: string }[] };

export function emptyDailyRoster(): DailyRoster { return { morning: [], afternoon: [], night: [], mid: [], off: [] }; }

export function buildRosterIndex(roster: RosterData | null) {
  if (!roster) return {} as Record<string, DailyRoster>;

  return roster.dateColumns.reduce<Record<string, DailyRoster>>((index, { isoDate }) => {
    const daily = emptyDailyRoster();

    roster.employees.forEach((employee) => {
      const code = shiftKey(roster.rows[employee]?.[isoDate] ?? 'OFF');
      if (code === 'M') daily.morning.push(employee);
      if (code === 'A') daily.afternoon.push(employee);
      if (code === 'N') daily.night.push(employee);
      if (code === 'MID') daily.mid.push(employee);
      if (code === 'OFF' || code === 'H8') daily.off.push(employee);
    });

    index[isoDate] = daily;
    return index;
  }, {});
}

function removeEmployee(names: string[], selectedEmployee: string) {
  return names.filter((name) => name !== selectedEmployee);
}

/**
 * Builds the "also working" group for the bottom sheet.
 * - Morning: everyone on M, plus every MID employee (tagged "Mid").
 * - Afternoon: everyone on A, plus every MID employee (tagged "Mid").
 * - Night: only N employees.
 * - Mid: Morning employees (tagged "Morning"), Afternoon employees (tagged "Afternoon"),
 *        and any other MID employees (tagged "Mid").
 */
export function groupForShift(shift: string, daily: DailyRoster | undefined, selectedEmployee: string): WorkingGroup | null {
  if (!daily) return null;
  const code = shiftKey(shift);
  if (code === 'OFF') return null;

  let employees: { name: string; suffix?: string }[] = [];

  if (code === 'M') {
    employees = [
      ...removeEmployee(daily.morning, selectedEmployee).map((name) => ({ name })),
      ...removeEmployee(daily.mid, selectedEmployee).map((name) => ({ name, suffix: 'Mid' })),
    ];
  } else if (code === 'A') {
    employees = [
      ...removeEmployee(daily.afternoon, selectedEmployee).map((name) => ({ name })),
      ...removeEmployee(daily.mid, selectedEmployee).map((name) => ({ name, suffix: 'Mid' })),
    ];
  } else if (code === 'N') {
    employees = removeEmployee(daily.night, selectedEmployee).map((name) => ({ name }));
  } else if (code === 'MID') {
    employees = [
      ...removeEmployee(daily.morning, selectedEmployee).map((name) => ({ name, suffix: 'Morning' })),
      ...removeEmployee(daily.afternoon, selectedEmployee).map((name) => ({ name, suffix: 'Afternoon' })),
      ...removeEmployee(daily.mid, selectedEmployee).map((name) => ({ name, suffix: 'Mid' })),
    ];
  }

  return { title: shiftLabel(shift), employees: employees.sort(byName) };
}

// ---------------------------------------------------------------------------
// "Show all shifts" — every employee on this day, grouped by shift
// ---------------------------------------------------------------------------

export type AllShiftsGroup = { code: string; label: string; employees: string[] };

const ALL_SHIFTS_ORDER: { code: string; field: keyof DailyRoster }[] = [
  { code: 'M', field: 'morning' },
  { code: 'MID', field: 'mid' },
  { code: 'A', field: 'afternoon' },
  { code: 'N', field: 'night' },
  { code: 'OFF', field: 'off' },
];

export function buildAllShiftsGroups(daily: DailyRoster | undefined): AllShiftsGroup[] {
  if (!daily) return [];
  return ALL_SHIFTS_ORDER
    .map(({ code, field }) => ({ code, label: shiftLabel(code), employees: [...daily[field]].sort((a, b) => a.localeCompare(b)) }))
    .filter((group) => group.employees.length > 0);
}

// ---------------------------------------------------------------------------
// Shift swap compatibility — evaluates whether two employees can trade shifts
// safely, scoring the trade and explaining why. Mirrors the phased design:
// shift-type compatibility, rest-time checks, monthly hours balance, and
// minimum staffing coverage.
// ---------------------------------------------------------------------------

export type SwapSeverity = 'good' | 'warning' | 'blocked';

export type SwapCompatibilityResult = {
  compatible: boolean;
  score: number;
  severity: SwapSeverity;
  reasons: string[];
};

export type SwapCandidateResult = SwapCompatibilityResult & {
  employee: string;
  shift: string;
};

// A normal 8h break is fine for roster coverage (for example, Night ends at
// 06:00 and Morning starts at 06:00 the next day = 0h, which is the only true
// problem). We allow standard 8h adjacent gaps, but block any trade that would
// force back-to-back continuous work with no pause between shifts.
const PREFERRED_MIN_REST_HOURS = 0;
const HARD_MIN_REST_HOURS = 0;

// Which shift codes are reasonable to trade with which. Night is kept strict
// since it affects rest the most; MID is allowed to flex into M/A since it
// already overlaps both in groupForShift().
const SWAP_COMPATIBLE_SHIFTS: Record<string, string[]> = {
  M: ['M', 'MID', 'A'],
  A: ['A', 'MID', 'M'],
  MID: ['MID', 'M', 'A'],
  N: ['N'],
};

const SHIFT_TIME_RANGES: Record<string, { startHour: number; endHour: number; overnight?: boolean }> = {
  M: { startHour: 6, endHour: 14 },
  A: { startHour: 14, endHour: 22 },
  N: { startHour: 22, endHour: 6, overnight: true },
  MID: { startHour: 9, endHour: 17 },
};

function addDaysIso(isoDate: string, delta: number): string {
  const [y, m, d] = isoDate.split('-').map(Number);
  const date = new Date(y, m - 1, d + delta);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function isSameMonth(isoDate: string, month: number, year: number) {
  const date = new Date(`${isoDate}T00:00:00`);
  return date.getMonth() === month && date.getFullYear() === year;
}

function isBlockBoundary(roster: RosterData, employee: string, isoDate: string): boolean {
  const shift = shiftKey(roster.rows[employee]?.[isoDate] ?? 'OFF');
  if (shift === 'OFF' || shift === 'H8') return false;
  const block = detectShiftBlock(roster, employee, isoDate);
  return block.shift === shift && (block.startDate === isoDate || block.endDate === isoDate);
}

/**
 * Whole block trade: both colleagues' shift blocks have the exact same date span
 * and differing shift types.
 */
export function isWholeBlockOverlapCandidate(
  roster: RosterData,
  requester: string,
  requesterDate: string,
  candidate: string,
  candidateDate: string,
): boolean {
  const requesterShift = shiftKey(roster.rows[requester]?.[requesterDate] ?? 'OFF');
  const candidateShift = shiftKey(roster.rows[candidate]?.[candidateDate] ?? 'OFF');
  if (
    requesterShift === 'OFF' || candidateShift === 'OFF' ||
    requesterShift === 'H8' || candidateShift === 'H8' ||
    requesterShift === 'MID' || candidateShift === 'MID'
  ) {
    return false;
  }

  const requesterBlock = detectShiftBlock(roster, requester, requesterDate);
  const candidateBlock = detectShiftBlock(roster, candidate, candidateDate);
  const requesterDates = blockDateRange(requesterBlock);
  const candidateDates = blockDateRange(candidateBlock);
  const sameBoundary = requesterBlock.startDate === candidateBlock.startDate
    && requesterBlock.endDate === candidateBlock.endDate;
  const blockLenMatch = requesterDates.length === candidateDates.length;
  const sameExactBlock = sameBoundary && requesterShift === candidateShift;

  return sameBoundary && blockLenMatch && !sameExactBlock;
}

/**
 * Partial block trade: the two shift blocks overlap partially (different boundaries or lengths).
 */
export function isPartialBlockOverlapCandidate(
  roster: RosterData,
  requester: string,
  requesterDate: string,
  candidate: string,
  candidateDate: string,
): boolean {
  const requesterShift = shiftKey(roster.rows[requester]?.[requesterDate] ?? 'OFF');
  const candidateShift = shiftKey(roster.rows[candidate]?.[candidateDate] ?? 'OFF');
  if (
    requesterShift === 'OFF' || candidateShift === 'OFF' ||
    requesterShift === 'H8' || candidateShift === 'H8' ||
    requesterShift === 'MID' || candidateShift === 'MID'
  ) {
    return false;
  }

  // Same-shift is pointless
  if (requesterShift === candidateShift) return false;

  const requesterBlock = detectShiftBlock(roster, requester, requesterDate);
  const candidateBlock = detectShiftBlock(roster, candidate, candidateDate);
  const requesterDates = blockDateRange(requesterBlock);
  const candidateDates = new Set(blockDateRange(candidateBlock));
  const overlapDates = requesterDates.filter((date) => candidateDates.has(date));
  const sameBoundary = requesterBlock.startDate === candidateBlock.startDate
    && requesterBlock.endDate === candidateBlock.endDate;
  const blockLenMatch = requesterDates.length === candidateDates.size;
  const sameExactBlock = sameBoundary && requesterShift === candidateShift;

  return overlapDates.length > 1 && !sameExactBlock && !(sameBoundary && blockLenMatch);
}

/**
 * A day-only swap should only be suggested when the colleague's block starts or
 * ends on the selected date. This prevents middle-of-block swaps from being
 * treated as valid one-day changes and keeps the "no two linked shifts in a row"
 * rule intact unless the request is on the block's edge.
 */
export function isDayOnlyBoundaryCandidate(roster: RosterData, employee: string, isoDate: string): boolean {
  const shift = shiftKey(roster.rows[employee]?.[isoDate] ?? 'OFF');
  if (shift === 'OFF' || shift === 'H8' || shift === 'MID') return false;
  const block = detectShiftBlock(roster, employee, isoDate);
  return block.startDate === isoDate || block.endDate === isoDate;
}

export function remainderOfBlockDates(roster: RosterData, employee: string, isoDate: string): string[] {
  const shift = shiftKey(roster.rows[employee]?.[isoDate] ?? 'OFF');
  if (shift === 'OFF' || shift === 'H8') return [];

  const block = detectShiftBlock(roster, employee, isoDate);
  if (isoDate < block.startDate || isoDate > block.endDate) return [];

  const dates: string[] = [];
  let currentDate = isoDate;
  while (currentDate <= block.endDate) {
    dates.push(currentDate);
    currentDate = addDaysIso(currentDate, 1);
  }

  return dates;
}

export function isSwapRemainingOfBlockCandidate(
  roster: RosterData,
  requester: string,
  requesterDate: string,
  candidate: string,
  candidateDate: string,
): boolean {
  const requesterShift = shiftKey(roster.rows[requester]?.[requesterDate] ?? 'OFF');
  const candidateShift = shiftKey(roster.rows[candidate]?.[candidateDate] ?? 'OFF');
  if (
    requesterShift === 'OFF' || candidateShift === 'OFF' ||
    requesterShift === 'H8' || candidateShift === 'H8' ||
    requesterShift === 'MID' || candidateShift === 'MID' ||
    requesterShift === candidateShift
  ) {
    return false;
  }

  const requesterBlock = detectShiftBlock(roster, requester, requesterDate);
  // Must be partway through the block (not on the first day)
  if (requesterDate <= requesterBlock.startDate) return false;

  const candidateBlock = detectShiftBlock(roster, candidate, candidateDate);
  const candidateDates = new Set(blockDateRange(candidateBlock));
  const remainderDates = remainderOfBlockDates(roster, requester, requesterDate);
  if (remainderDates.length <= 1) return false;

  const overlapDates = remainderDates.filter((date) => candidateDates.has(date));
  return overlapDates.length > 1;
}

/** Wall-clock start/end for a shift on a given date. Null for OFF/H8, which have no time span. */
export function getShiftInterval(isoDate: string, shift: string): { start: Date; end: Date } | null {
  const range = SHIFT_TIME_RANGES[shiftKey(shift)];
  if (!range) return null;
  const [y, m, d] = isoDate.split('-').map(Number);
  const start = new Date(y, m - 1, d, range.startHour, 0, 0);
  const end = new Date(y, m - 1, d, range.endHour, 0, 0);
  if (range.overnight) end.setDate(end.getDate() + 1);
  return { start, end };
}

/**
 * Smallest rest gap (in hours) adjacent to `isoDate` for `employee`, assuming
 * `newShift` replaces whatever they currently have that day. Infinity if the
 * new shift has no time span (OFF/H8) or there's nothing adjacent to compare.
 */
function minAdjacentRestHours(roster: RosterData, employee: string, isoDate: string, newShift: string): number {
  const newInterval = getShiftInterval(isoDate, newShift);
  if (!newInterval) return Infinity;

  const prevIso = addDaysIso(isoDate, -1);
  const nextIso = addDaysIso(isoDate, 1);
  const prevInterval = getShiftInterval(prevIso, roster.rows[employee]?.[prevIso] ?? 'OFF');
  const nextInterval = getShiftInterval(nextIso, roster.rows[employee]?.[nextIso] ?? 'OFF');

  const gaps: number[] = [];
  if (prevInterval) gaps.push((newInterval.start.getTime() - prevInterval.end.getTime()) / 3_600_000);
  if (nextInterval) gaps.push((nextInterval.start.getTime() - newInterval.end.getTime()) / 3_600_000);

  return gaps.length ? Math.min(...gaps) : Infinity;
}

// Month-state balancing is a final-schedule constraint for the solver. It must
// not be enforced at pairwise swap level, otherwise any valid whole-block trade
// is rejected before the optimizer even gets a chance to evaluate the final month.
function countRegularOffDaysForEmployee(
  rows: Record<string, Record<string, string>>,
  employee: string,
  monthDates: { isoDate: string }[],
): number {
  let count = 0;
  for (const { isoDate } of monthDates) {
    const rawShift = (rows[employee]?.[isoDate] ?? 'OFF').trim().toUpperCase();
    const normalized = shiftKey(rawShift);
    if (normalized === 'OFF' && rawShift !== 'ABS' && rawShift !== 'AD' && rawShift !== 'H8') count += 1;
  }
  return count;
}

function isMonthStateOffDayBalanced(
  rows: Record<string, Record<string, string>>,
  employee: string,
  monthDates: { isoDate: string }[],
): boolean {
  return countRegularOffDaysForEmployee(rows, employee, monthDates) === 8;
}

function getBaselineMaxStreak(roster: RosterData, employee: string, monthDates: { isoDate: string }[]): number {
  const monthDateSet = new Set(monthDates.map((m) => m.isoDate));
  let streak = 0;
  let maxStreak = 0;
  let streakTouchesMonth = false;

  for (const { isoDate } of roster.dateColumns) {
    const shift = shiftKey(roster.rows[employee]?.[isoDate] ?? 'OFF');
    if (isWorkingShift(shift)) {
      streak += 1;
      if (monthDateSet.has(isoDate)) streakTouchesMonth = true;
      if (streakTouchesMonth) {
        maxStreak = Math.max(maxStreak, streak);
      }
    } else {
      streak = 0;
      streakTouchesMonth = false;
    }
  }

  return maxStreak;
}

function isMonthStateLegallyBalanced(
  roster: RosterData,
  rows: Record<string, Record<string, string>>,
  monthDates: { isoDate: string }[],
  employees: string[] = roster.employees,
  options: { requireOffDayBalance?: boolean; allowNightToDay?: boolean } = {},
): boolean {
  if (!employees.length) return false;

  const monthDateSet = new Set(monthDates.map((m) => m.isoDate));

  for (const employee of employees) {
    if (options.requireOffDayBalance && !isMonthStateOffDayBalanced(rows, employee, monthDates)) return false;

    let streak = 0;
    let maxStreak = 0;
    let streakTouchesMonth = false;
    for (const { isoDate } of roster.dateColumns) {
      const shift = shiftKey(rows[employee]?.[isoDate] ?? 'OFF');
      if (isWorkingShift(shift)) {
        streak += 1;
        if (monthDateSet.has(isoDate)) streakTouchesMonth = true;
        if (streakTouchesMonth) {
          maxStreak = Math.max(maxStreak, streak);
        }
      } else {
        streak = 0;
        streakTouchesMonth = false;
      }
    }
    if (maxStreak > 5) {
      const baselineMax = getBaselineMaxStreak(roster, employee, monthDates);
      if (maxStreak > baselineMax || baselineMax <= 5) {
        return false;
      }
    }

    for (let index = 0; index < roster.dateColumns.length - 1; index += 1) {
      const a = roster.dateColumns[index].isoDate;
      const b = roster.dateColumns[index + 1].isoDate;
      if (!monthDateSet.has(a) && !monthDateSet.has(b)) continue;
      const currentShift = shiftKey(rows[employee]?.[a] ?? 'OFF');
      const nextShift = shiftKey(rows[employee]?.[b] ?? 'OFF');
      if (!options.allowNightToDay && currentShift === 'N' && (nextShift === 'M' || nextShift === 'A' || nextShift === 'MID')) return false;
    }
  }

  return true;
}

function wouldCreateConsecutiveShiftForSwap(
  roster: RosterData,
  employee: string,
  isoDate: string,
  newShift: string,
): boolean {
  const targetShift = shiftKey(newShift);
  if (!isWorkingShift(targetShift)) return false;

  const prevIso = addDaysIso(isoDate, -1);
  const prevShift = shiftKey(roster.rows[employee]?.[prevIso] ?? 'OFF');
  const nextIso = addDaysIso(isoDate, 1);
  const nextShift = shiftKey(roster.rows[employee]?.[nextIso] ?? 'OFF');

  if (prevShift === 'N' && (targetShift === 'M' || targetShift === 'A' || targetShift === 'MID')) {
    return true;
  }

  if (targetShift === 'N' && (nextShift === 'M' || nextShift === 'A' || nextShift === 'MID')) {
    return true;
  }

  return false;
}

function countOnShift(roster: RosterData, isoDate: string, shiftCode: string, exclude: string): number {
  let count = 0;
  roster.employees.forEach((employee) => {
    if (employee === exclude) return;
    if (shiftKey(roster.rows[employee]?.[isoDate] ?? 'OFF') === shiftCode) count += 1;
  });
  return count;
}

function isWorkingShift(shift: string): boolean {
  const code = shiftKey(shift);
  return code !== 'OFF' && code !== 'H8';
}

function getMaxConsecutiveWorkStreak(
  roster: RosterData,
  employee: string,
  overrides: Record<string, string> = {},
  focusDate?: string,
  radius: number = 7,
): { max: number; startDate: string | null; endDate: string | null } {
  let max = 0;
  let streak = 0;
  let streakStart: string | null = null;
  let streakEnd: string | null = null;

  for (const { isoDate } of roster.dateColumns) {
    if (focusDate) {
      const focus = new Date(`${focusDate}T00:00:00`);
      const current = new Date(`${isoDate}T00:00:00`);
      const diffDays = Math.abs((current.getTime() - focus.getTime()) / 86_400_000);
      if (diffDays > radius) continue;
    }

    const nextShift = overrides[isoDate] ?? shiftKey(roster.rows[employee]?.[isoDate] ?? 'OFF');
    if (isWorkingShift(nextShift)) {
      streak += 1;
      if (streak === 1) streakStart = isoDate;
      streakEnd = isoDate;
      max = Math.max(max, streak);
    } else {
      streak = 0;
      streakStart = null;
      streakEnd = null;
    }
  }

  return { max, startDate: streakStart, endDate: streakEnd };
}

export function simulateWholeBlockTradeStreaks(
  roster: RosterData,
  requester: string,
  requesterDate: string,
  candidate: string,
  candidateDate: string,
): { requester: { max: number; startDate: string | null; endDate: string | null }; candidate: { max: number; startDate: string | null; endDate: string | null } } {
  const requesterBlock = detectShiftBlock(roster, requester, requesterDate);
  const candidateBlock = detectShiftBlock(roster, candidate, candidateDate);

  const requesterOverrides: Record<string, string> = {};
  const candidateOverrides: Record<string, string> = {};

  for (const { isoDate } of roster.dateColumns) {
    requesterOverrides[isoDate] = shiftKey(roster.rows[requester]?.[isoDate] ?? 'OFF');
    candidateOverrides[isoDate] = shiftKey(roster.rows[candidate]?.[isoDate] ?? 'OFF');
  }

  for (const date of blockDateRange(candidateBlock)) {
    requesterOverrides[date] = shiftKey(roster.rows[candidate]?.[date] ?? 'OFF');
  }
  for (const date of blockDateRange(requesterBlock)) {
    candidateOverrides[date] = shiftKey(roster.rows[requester]?.[date] ?? 'OFF');
  }

  return {
    requester: getMaxConsecutiveWorkStreak(roster, requester, requesterOverrides, requesterDate, 14),
    candidate: getMaxConsecutiveWorkStreak(roster, candidate, candidateOverrides, candidateDate, 14),
  };
}

function simulatePartialBlockTradeStreaks(
  roster: RosterData,
  requester: string,
  candidate: string,
  overlapDates: string[],
  requesterNewShift: string,
  candidateNewShift: string,
  focusDate: string,
): { requester: { max: number; startDate: string | null; endDate: string | null }; candidate: { max: number; startDate: string | null; endDate: string | null } } {
  const requesterOverrides: Record<string, string> = {};
  const candidateOverrides: Record<string, string> = {};

  for (const { isoDate } of roster.dateColumns) {
    requesterOverrides[isoDate] = shiftKey(roster.rows[requester]?.[isoDate] ?? 'OFF');
    candidateOverrides[isoDate] = shiftKey(roster.rows[candidate]?.[isoDate] ?? 'OFF');
  }

  for (const date of overlapDates) {
    requesterOverrides[date] = requesterNewShift;
    candidateOverrides[date] = candidateNewShift;
  }

  return {
    requester: getMaxConsecutiveWorkStreak(roster, requester, requesterOverrides, focusDate, 14),
    candidate: getMaxConsecutiveWorkStreak(roster, candidate, candidateOverrides, focusDate, 14),
  };
}

function isPartialTradeRestValid(
  roster: RosterData,
  employee: string,
  overlapDates: string[],
  newShift: string,
): boolean {
  if (!isWorkingShift(newShift)) return true;
  const sortedDates = [...overlapDates].sort();
  const firstDate = sortedDates[0];
  const lastDate = sortedDates[sortedDates.length - 1];

  const prevIso = addDaysIso(firstDate, -1);
  const prevShift = shiftKey(roster.rows[employee]?.[prevIso] ?? 'OFF');
  const nextIso = addDaysIso(lastDate, 1);
  const nextShift = shiftKey(roster.rows[employee]?.[nextIso] ?? 'OFF');

  // Night to Day without rest
  if (prevShift === 'N' && (newShift === 'M' || newShift === 'A' || newShift === 'MID')) {
    return false;
  }
  if (newShift === 'N' && (nextShift === 'M' || nextShift === 'A' || nextShift === 'MID')) {
    return false;
  }

  const firstInterval = getShiftInterval(firstDate, newShift);
  const lastInterval = getShiftInterval(lastDate, newShift);

  const prevInterval = getShiftInterval(prevIso, prevShift);
  if (firstInterval && prevInterval) {
    const gap = (firstInterval.start.getTime() - prevInterval.end.getTime()) / 3_600_000;
    if (gap < 0) return false;
  }

  const nextInterval = getShiftInterval(nextIso, nextShift);
  if (lastInterval && nextInterval) {
    const gap = (nextInterval.start.getTime() - lastInterval.end.getTime()) / 3_600_000;
    if (gap < 0) return false;
  }

  return true;
}

/**
 * Reports how each side's *old* shift is left staffed after they move away
 * from it — informational by default (`warnings`), only escalated to
 * `blockers` when a shift would be left with nobody on it at all. We don't
 * assume a specific headcount policy; we just surface the real numbers.
 */
function checkCoverageAfterSwap(
  roster: RosterData,
  requester: string, requesterDate: string, requesterOldShift: string,
  candidate: string, candidateDate: string, candidateOldShift: string,
): { warnings: string[]; blockers: string[] } {
  const warnings: string[] = [];
  const blockers: string[] = [];

  // When swapping on the exact same date, two working employees exchange shifts.
  // The shift counts on that date are unchanged.
  if (requesterDate === candidateDate) {
    if (requesterOldShift === 'OFF' || candidateOldShift === 'OFF') {
      const workingShift = requesterOldShift !== 'OFF' ? requesterOldShift : candidateOldShift;
      const workingPerson = requesterOldShift !== 'OFF' ? requester : candidate;
      const remaining = countOnShift(roster, requesterDate, workingShift, workingPerson);
      if (remaining === 0) {
        warnings.push(`No one else would be covering ${shiftLabel(workingShift)} on ${requesterDate} after this swap.`);
      } else if (remaining === 1) {
        warnings.push(`${shiftLabel(workingShift)} on ${requesterDate} would go down to a single person.`);
      }
    }
    return { warnings, blockers };
  }

  const checkDateShift = (date: string, shiftLost: string, person: string) => {
    if (shiftLost !== 'OFF') {
      const remaining = countOnShift(roster, date, shiftLost, person);
      if (remaining === 0) {
        warnings.push(`No one else would be covering ${shiftLabel(shiftLost)} on ${date} after this swap.`);
      } else if (remaining === 1) {
        warnings.push(`${shiftLabel(shiftLost)} on ${date} would go down to a single person.`);
      }
    }
  };

  checkDateShift(requesterDate, requesterOldShift, requester);
  checkDateShift(candidateDate, candidateOldShift, candidate);

  return { warnings, blockers };
}

export type SwapLegalityResult = {
  valid: boolean;
  blockers: string[];
};

export type SwapTradeStructureType = 'WHOLE_BLOCK' | 'PARTIAL_BLOCK' | 'DAY_ONLY' | 'OFF_DAY';

export type SwapStructureResult = {
  type: SwapTradeStructureType;
  applicableDates: string[];
  blockedDates: string[];
};

export type SwapBalanceResult = {
  hoursGiven: number;
  hoursReceived: number;
  debtDays: number;
  debtOwner?: string;
  owedTo?: string;
  possible: boolean;
  suggestedDates: string[];
};

export type SwapEvaluationResult = {
  legality: SwapLegalityResult;
  structure: SwapStructureResult;
  balance: SwapBalanceResult;
  quality: SwapCompatibilityResult;
};

export type ShiftTransformationSwap = {
  employee: string;
  otherEmployee: string;
  isoDate: string;
  candidateIsoDate?: string;
  from: string;
  to: string;
  type?: 'WHOLE_BLOCK' | 'PARTIAL_BLOCK' | 'DAY_ONLY';
  status?: 'good' | 'review';
};

export type ShiftTransformationResult = {
  possible: boolean;
  targetShift: string;
  swaps: ShiftTransformationSwap[];
  finalRoster: Record<string, Record<string, string>>;
  summary: string;
  firstAttempt?: string;
};

type TransformationMove = {
  employee: string;
  otherEmployee: string;
  dates: string[];
  employeeAfter: Record<string, string>;
  otherAfter: Record<string, string>;
  type: 'WHOLE_BLOCK' | 'PARTIAL_BLOCK' | 'DAY_ONLY';
  label: string;
};

export function buildTransformationMoves(
  roster: RosterData,
  rows: Record<string, Record<string, string>>,
  employee: string,
  targetShift: string,
): TransformationMove[] {
  const requestedShift = shiftKey(targetShift);
  if (requestedShift === 'MID') return [];
  const unresolvedDates = roster.dateColumns
    .map((entry) => entry.isoDate)
    .filter((isoDate) => {
      const shift = shiftKey(rows[employee]?.[isoDate] ?? 'OFF');
      return shift !== requestedShift && shift !== 'MID' && shift !== 'H8';
    });

  const moves: TransformationMove[] = [];

  for (const isoDate of unresolvedDates) {
    const employeeBlock = detectShiftBlock(roster, employee, isoDate, rows);
    const employeeDates = new Set(blockDateRange(employeeBlock));

    for (const candidate of roster.employees) {
      if (candidate === employee) continue;

      for (const candidateDate of roster.dateColumns.map((entry) => entry.isoDate)) {
        const candidateShift = shiftKey(rows[candidate]?.[candidateDate] ?? 'OFF');
        if (candidateShift !== requestedShift || candidateShift === 'MID') continue;

        const candidateBlock = detectShiftBlock(roster, candidate, candidateDate, rows);
        const candidateDates = new Set(blockDateRange(candidateBlock));
        const overlapDates = [...employeeDates].filter((date) => candidateDates.has(date));

        const tradeDates = overlapDates.length > 0
          ? Array.from(new Set([...employeeDates, ...candidateDates]))
          : [isoDate, candidateDate];

        const currentRoster = { ...roster, rows };
        const compatibility = evaluateShiftSwapCompatibility(
          currentRoster,
          employee,
          isoDate,
          candidate,
          candidateDate,
        );

        if (!compatibility.compatible) continue;

        const employeeAfter: Record<string, string> = {};
        const otherAfter: Record<string, string> = {};
        for (const tradeDate of tradeDates) {
          employeeAfter[tradeDate] = shiftKey(rows[candidate]?.[tradeDate] ?? 'OFF');
          otherAfter[tradeDate] = shiftKey(rows[employee]?.[tradeDate] ?? 'OFF');
        }

        const type = overlapDates.length > 0
          ? (employeeDates.size === candidateDates.size && overlapDates.length === employeeDates.size ? 'WHOLE_BLOCK' : 'PARTIAL_BLOCK')
          : 'DAY_ONLY';

        moves.push({
          employee,
          otherEmployee: candidate,
          dates: tradeDates,
          employeeAfter,
          otherAfter,
          type,
          label: `${employee} ${isoDate} ↔ ${candidate} ${candidateDate} (${type})`,
        });
      }
    }
  }

  return moves.sort((a, b) => b.dates.length - a.dates.length);
}

export function applyTransformationMove(
  rows: Record<string, Record<string, string>>,
  move: TransformationMove,
): Record<string, Record<string, string>> {
  const nextRows = Object.fromEntries(
    Object.entries(rows).map(([person, dayMap]) => [person, { ...dayMap }]),
  ) as Record<string, Record<string, string>>;

  for (const date of move.dates) {
    nextRows[move.employee][date] = move.otherAfter[date] ?? 'OFF';
    nextRows[move.otherEmployee][date] = move.employeeAfter[date] ?? 'OFF';
  }

  return nextRows;
}

export function findShiftTransformationPreview(
  roster: RosterData,
  employee: string,
  targetShift: string,
  month: number,
  year: number,
  simulationVariant: number = 0,
): ShiftTransformationResult {
  const requestedShift = shiftKey(targetShift);
  if (requestedShift === 'MID') {
    return {
      possible: false,
      targetShift: 'MID',
      swaps: [],
      finalRoster: roster.rows,
      summary: 'MID shifts cannot swap with anyone.',
    };
  }

  const monthDates = roster.dateColumns.filter((entry) => {
    const date = new Date(`${entry.isoDate}T00:00:00`);
    return date.getMonth() === month && date.getFullYear() === year;
  });

  const cloneRows = (rows: Record<string, Record<string, string>>) => Object.fromEntries(
    roster.employees.map((person) => [person, { ...(rows[person] ?? {}) }]),
  ) as Record<string, Record<string, string>>;

  const currentRows = cloneRows(roster.rows);
  const seen = new Set<string>();
  const suggested: Array<ShiftTransformationSwap & { type: 'WHOLE_BLOCK' | 'PARTIAL_BLOCK' | 'DAY_ONLY'; status: 'good' | 'review' | 'blocked' }> = [];

  for (const isoDate of monthDates.map(({ isoDate }) => isoDate)) {
    const currentShift = shiftKey(currentRows[employee]?.[isoDate] ?? 'OFF');
    if (currentShift === requestedShift || currentShift === 'OFF' || currentShift === 'H8' || currentShift === 'MID') continue;

    for (const candidate of roster.employees) {
      if (candidate === employee) continue;

      for (const candidateDate of monthDates.map(({ isoDate }) => isoDate)) {
        const candidateShift = shiftKey(currentRows[candidate]?.[candidateDate] ?? 'OFF');
        if (candidateShift !== requestedShift || candidateShift === 'MID') continue;

        const currentRoster = { ...roster, rows: currentRows };
        const requesterBlock = detectShiftBlock(currentRoster, employee, isoDate, currentRows);
        const candidateBlock = detectShiftBlock(currentRoster, candidate, candidateDate, currentRows);
        const requesterDates = blockDateRange(requesterBlock);
        const candidateDates = new Set(blockDateRange(candidateBlock));
        const overlapDates = requesterDates.filter((date) => candidateDates.has(date));
        if (overlapDates.length === 0 && candidateDate !== isoDate) continue;
        const candidateHasRequestedShiftOnDate = candidateShift === requestedShift;
        const sameDateRange = requesterBlock.startDate === candidateBlock.startDate
          && requesterBlock.endDate === candidateBlock.endDate;
        const structureType: 'WHOLE_BLOCK' | 'PARTIAL_BLOCK' | 'DAY_ONLY' = sameDateRange
          ? 'WHOLE_BLOCK'
          : overlapDates.length > 1
            ? 'PARTIAL_BLOCK'
            : candidateDate === isoDate
              ? (overlapDates.length === 1 && (requesterDates.length > 1 || candidateDates.size > 1) ? 'PARTIAL_BLOCK' : 'DAY_ONLY')
              : overlapDates.length > 0
                ? 'PARTIAL_BLOCK'
                : candidateHasRequestedShiftOnDate
                  ? 'PARTIAL_BLOCK'
                  : 'DAY_ONLY';
        const normalizedType: 'WHOLE_BLOCK' | 'PARTIAL_BLOCK' | 'DAY_ONLY' = structureType;

        const operationDates = normalizedType === 'DAY_ONLY'
          ? [isoDate]
          : overlapDates.length > 0
            ? overlapDates
            : requesterDates;
        const sourceDates = normalizedType === 'DAY_ONLY' ? [candidateDate] : operationDates;
        const sourceShifts = sourceDates.map((sourceDate) => shiftKey(currentRows[candidate]?.[sourceDate] ?? 'OFF'));
        if (sourceShifts.some((sourceShift) => sourceShift !== requestedShift || sourceShift === 'H8' || sourceShift === 'MID')) continue;
        const requesterShifts = operationDates.map((opDate) => shiftKey(currentRows[employee]?.[opDate] ?? 'OFF'));
        if (requesterShifts.some((s) => s === 'MID' || s === 'H8')) continue;
        const operationRows = cloneRows(currentRows);
        if (normalizedType === 'DAY_ONLY' && isoDate !== candidateDate) {
          operationRows[employee][isoDate] = shiftKey(currentRows[candidate]?.[candidateDate] ?? 'OFF');
          operationRows[candidate][candidateDate] = shiftKey(currentRows[employee]?.[isoDate] ?? 'OFF');
        } else {
          for (const operationDate of operationDates) {
            operationRows[employee][operationDate] = shiftKey(currentRows[candidate]?.[operationDate] ?? 'OFF');
            operationRows[candidate][operationDate] = shiftKey(currentRows[employee]?.[operationDate] ?? 'OFF');
          }
        }
        const partialNightStartsAtBlockStart = normalizedType === 'PARTIAL_BLOCK'
          && candidateShift === 'N'
          && operationDates[0] === candidateBlock.startDate;
        const operationIsLegal = isMonthStateLegallyBalanced(
          roster,
          operationRows,
          monthDates,
          [employee, candidate],
          { requireOffDayBalance: false, allowNightToDay: partialNightStartsAtBlockStart },
        );
        if (!operationIsLegal) continue;
        if (normalizedType === 'DAY_ONLY') {
          const requesterDayShift = shiftKey(currentRows[employee]?.[isoDate] ?? 'OFF');
          const candidateDayShift = shiftKey(currentRows[candidate]?.[candidateDate] ?? 'OFF');
          if (requesterDayShift === 'MID' || candidateDayShift === 'MID'
            || requesterDayShift === 'H8' || candidateDayShift === 'H8'
            || (requesterDayShift === 'OFF' && candidateDayShift === 'OFF')
            || requesterDayShift === candidateDayShift) continue;
        }
        const status: 'good' | 'review' = normalizedType === 'WHOLE_BLOCK'
          ? 'good'
          : 'review';
        const key = `${employee}|${candidate}|${normalizedType}|${operationDates.join(',')}|${candidateDate}`;
        if (seen.has(key)) continue;
        seen.add(key);

        suggested.push({
          employee,
          otherEmployee: candidate,
          isoDate: normalizedType === 'DAY_ONLY' && isoDate !== candidateDate
            ? `${isoDate} ↔ ${candidateDate}`
            : operationDates.join(', '),
          candidateIsoDate: candidateDate,
          from: currentShift,
          to: requestedShift,
          type: normalizedType,
          status,
        });
      }
    }
  }

  if (suggested.length > 0) {
    const operationDatesForSuggestion = (suggestion: ShiftTransformationSwap): string[] => {
      const crossDateParts = suggestion.isoDate.split(' ↔ ');
      return crossDateParts.length === 2 ? [crossDateParts[0]] : suggestion.isoDate.split(', ');
    };
    const candidateDatesForSuggestion = (suggestion: ShiftTransformationSwap): string[] => {
      const requesterDates = operationDatesForSuggestion(suggestion);
      if (requesterDates.length === 1 && suggestion.candidateIsoDate) return [suggestion.candidateIsoDate];
      return requesterDates;
    };
    const qualityRank = (suggestion: ShiftTransformationSwap) => suggestion.type === 'WHOLE_BLOCK'
      ? 3
      : suggestion.type === 'PARTIAL_BLOCK'
        ? 2
        : 1;
    const reservedDates = new Set<string>();
    const reservedTargetDates = new Set(
      monthDates
        .map(({ isoDate }) => isoDate)
        .filter((date) => shiftKey(currentRows[employee]?.[date] ?? 'OFF') === requestedShift),
    );
    let cumulativeRows = cloneRows(currentRows);

    const selectedSuggestions = [...suggested]
      .sort((a, b) => {
        const qualityDifference = qualityRank(b) - qualityRank(a);
        if (qualityDifference !== 0) return qualityDifference;
        const aLength = operationDatesForSuggestion(a).length;
        const bLength = operationDatesForSuggestion(b).length;
        if (aLength !== bLength) return bLength - aLength;
        const employeeDifference = a.otherEmployee.localeCompare(b.otherEmployee);
        return simulationVariant % 2 === 0 ? employeeDifference : -employeeDifference;
      })
      .filter((suggestion) => {
        const requesterDates = operationDatesForSuggestion(suggestion);
        const candidateDates = candidateDatesForSuggestion(suggestion);
        if (requesterDates.some((date) => reservedTargetDates.has(date))) return false;
        if (candidateDates.some((date) => reservedTargetDates.has(date))) return false;
        const requesterKeys = requesterDates.map((date) => `${employee}|${date}`);
        const candidateKeys = candidateDates.map((date) => `${suggestion.otherEmployee}|${date}`);
        const operationKeys = [...requesterKeys, ...candidateKeys];
        if (operationKeys.some((key) => reservedDates.has(key))) return false;

        const testRows = cloneRows(cumulativeRows);
        const crossDateParts = suggestion.isoDate.split(' ↔ ');
        const previewDates = crossDateParts.length === 2 ? [crossDateParts[0]] : suggestion.isoDate.split(', ');
        for (const date of previewDates) {
          const reqShift = shiftKey(testRows[employee]?.[date] ?? 'OFF');
          const candDate = crossDateParts.length === 2 ? crossDateParts[1] : date;
          const candShift = shiftKey(testRows[suggestion.otherEmployee]?.[candDate] ?? 'OFF');
          testRows[employee][date] = candShift;
          testRows[suggestion.otherEmployee][candDate] = reqShift;
        }

        const isLegal = isMonthStateLegallyBalanced(
          roster,
          testRows,
          monthDates,
          [employee, suggestion.otherEmployee],
          { requireOffDayBalance: false, allowNightToDay: suggestion.type === 'PARTIAL_BLOCK' && suggestion.to === 'N' },
        );
        if (!isLegal) return false;

        cumulativeRows = testRows;
        operationKeys.forEach((key) => reservedDates.add(key));
        requesterDates.forEach((date) => reservedTargetDates.add(date));
        return true;
      });
    const bestSuggestion = selectedSuggestions[0];
    if (!bestSuggestion) {
      return {
        possible: false,
        targetShift: requestedShift,
        swaps: [],
        finalRoster: currentRows,
        summary: `No non-overlapping eligible ${shiftLabel(requestedShift)} swap was found for ${employee}.`,
      };
    }
    const previewRows = cumulativeRows;
    const selectedWholeBlockGood = selectedSuggestions.filter(({ type }) => type === 'WHOLE_BLOCK').length;
    const selectedPartialBlockReview = selectedSuggestions.filter(({ type }) => type === 'PARTIAL_BLOCK').length;

    return {
      possible: true,
      targetShift: requestedShift,
      swaps: selectedSuggestions.map(({ employee, otherEmployee, isoDate, candidateIsoDate, from, to, type, status }) => ({ employee, otherEmployee, isoDate, candidateIsoDate, from, to, type, status })),
      finalRoster: previewRows,
      summary: `Found ${selectedSuggestions.length} non-overlapping eligible ${shiftLabel(requestedShift)} swap${selectedSuggestions.length === 1 ? '' : 's'} for ${employee}: ${selectedWholeBlockGood} whole-block good, ${selectedPartialBlockReview} partial-block under review. Other candidates were blocked by the normal swap rules.`,
      firstAttempt: `${bestSuggestion.isoDate} with ${bestSuggestion.otherEmployee} (${bestSuggestion.type})`,
    };
  }

  if (!monthDates.length) {
    const emptyState = cloneRows(roster.rows);
    return {
      possible: false,
      targetShift: requestedShift,
      swaps: [],
      finalRoster: emptyState,
      summary: `No dates were available for ${month + 1}/${year}.`,
    };
  }

  return {
    possible: false,
    targetShift: requestedShift,
    swaps: [],
    finalRoster: currentRows,
    firstAttempt: undefined,
    summary: `No eligible ${shiftLabel(requestedShift)} swaps were found for ${employee} in this month under the normal swap rules.`,
  };
}

export function evaluateSwapLegality(
  roster: RosterData,
  requester: string,
  requesterDate: string,
  candidate: string,
  candidateDate: string,
): SwapLegalityResult {
  // 1. Both employees must exist; cannot swap with yourself; both dates must be in roster
  if (!roster.employees.includes(requester) || !roster.employees.includes(candidate)) {
    return { valid: false, blockers: ['One of the employees is not part of this roster.'] };
  }
  if (requester === candidate) {
    return { valid: false, blockers: ['Cannot swap a shift with yourself.'] };
  }
  const knownDates = new Set(roster.dateColumns.map((c) => c.isoDate));
  if (!knownDates.has(requesterDate) || !knownDates.has(candidateDate)) {
    return { valid: false, blockers: ['One of the selected dates is outside this roster.'] };
  }

  const requesterShift = shiftKey(roster.rows[requester]?.[requesterDate] ?? 'OFF');
  const candidateShift = shiftKey(roster.rows[candidate]?.[candidateDate] ?? 'OFF');

  // 2. MID exclusion — if either shift is MID, hard block. MID is never swappable.
  if (requesterShift === 'MID' || candidateShift === 'MID') {
    return { valid: false, blockers: ['MID shifts are excluded from swap possibilities.'] };
  }

  // Both-OFF check
  if (requesterShift === 'OFF' && candidateShift === 'OFF') {
    return { valid: false, blockers: ['Both employees are off — there is nothing to swap.'] };
  }

  // Same-shift swaps on the same date are pointless
  if (requesterDate === candidateDate && requesterShift === candidateShift) {
    return { valid: false, blockers: [`Same-shift swaps like ${shiftLabel(requesterShift)} ↔ ${shiftLabel(candidateShift)} are pointless and should not be suggested.`] };
  }

  // 3. Determine work blocks (shift is not OFF/H8)
  const requesterIsWorkBlock = requesterShift !== 'OFF' && requesterShift !== 'H8';
  const candidateIsWorkBlock = candidateShift !== 'OFF' && candidateShift !== 'H8';

  // 4. Compute blocks, overlap, and isBlockOverlapTrade (ONLY true when BOTH sides are real work blocks and overlap > 1 day)
  const requesterBlock = detectShiftBlock(roster, requester, requesterDate);
  const candidateBlock = detectShiftBlock(roster, candidate, candidateDate);
  const requesterDates = new Set(blockDateRange(requesterBlock));
  const candidateDates = new Set(blockDateRange(candidateBlock));
  const overlapDates = requesterIsWorkBlock && candidateIsWorkBlock
    ? [...requesterDates].filter((date) => candidateDates.has(date))
    : [];
  const sameBlockDifferentShift = requesterIsWorkBlock && candidateIsWorkBlock
    && requesterBlock.startDate === candidateBlock.startDate
    && requesterBlock.endDate === candidateBlock.endDate
    && requesterShift !== candidateShift;
  const isBlockOverlapTrade = requesterIsWorkBlock && candidateIsWorkBlock
    && (overlapDates.length > 1 || (requesterBlock.startDate <= candidateBlock.endDate && requesterBlock.endDate >= candidateBlock.startDate && overlapDates.length > 1));

  // 5 & 6. If sameBlockDifferentShift || isBlockOverlapTrade: treated as a block trade.
  if (sameBlockDifferentShift || isBlockOverlapTrade) {
    const sameShiftSameBoundary = requesterBlock.startDate === candidateBlock.startDate
      && requesterBlock.endDate === candidateBlock.endDate
      && requesterShift === candidateShift;
    if (sameShiftSameBoundary) {
      return { valid: false, blockers: [`Same-shift swaps like ${shiftLabel(requesterShift)} ↔ ${shiftLabel(candidateShift)} are pointless and should not be suggested.`] };
    }

    const sameBoundary = requesterBlock.startDate === candidateBlock.startDate
      && requesterBlock.endDate === candidateBlock.endDate;
    const tradeDates = sameBoundary ? blockDateRange(requesterBlock) : (overlapDates.length > 0 ? overlapDates : [requesterDate]);

    // Rest check across the block trade boundaries
    if (!isPartialTradeRestValid(roster, requester, tradeDates, candidateShift) ||
        !isPartialTradeRestValid(roster, candidate, tradeDates, requesterShift)) {
      return { valid: false, blockers: ['This swap would force either person into back-to-back shifts with no break, which is not allowed.'] };
    }

    // Streak check across the trade
    const simulated = simulatePartialBlockTradeStreaks(
      roster,
      requester,
      candidate,
      tradeDates,
      candidateShift,
      requesterShift,
      requesterDate,
    );
    if (simulated.requester.max > 5) {
      return { valid: false, blockers: [`${requester} would be on ${simulated.requester.max} consecutive working days`, `(${simulated.requester.startDate ?? requesterDate} → ${simulated.requester.endDate ?? requesterDate}), which exceeds the 5-day maximum.`] };
    }
    if (simulated.candidate.max > 5) {
      return { valid: false, blockers: [`${candidate} would be on ${simulated.candidate.max} consecutive working days`, `(${simulated.candidate.startDate ?? requesterDate} → ${simulated.candidate.endDate ?? requesterDate}), which exceeds the 5-day maximum.`] };
    }

    return { valid: true, blockers: [] };
  }

  // 7 & 8. Fallthrough path (no block overlap):
  if (requesterShift === 'OFF' && candidateShift === 'OFF') {
    return { valid: false, blockers: ['Both employees are off — there is nothing to swap.'] };
  }
  if (requesterShift === 'H8' || candidateShift === 'H8') {
    return { valid: false, blockers: ['Holiday shifts cannot be swapped here.'] };
  }
  if (requesterDate === candidateDate && requesterShift === candidateShift) {
    return { valid: false, blockers: [`Same-shift swaps like ${shiftLabel(requesterShift)} ↔ ${shiftLabel(candidateShift)} are pointless and should not be suggested.`] };
  }

  // Single-day swaps between working shifts must occur on the boundary of the colleague's block
  if (requesterShift !== 'OFF' && candidateShift !== 'OFF' && !isDayOnlyBoundaryCandidate(roster, candidate, candidateDate)) {
    return { valid: false, blockers: ['Single-day swaps are only allowed at the start or end of a colleague\'s shift block, not in the middle.'] };
  }

  const requesterOverrides: Record<string, string> = { [requesterDate]: candidateShift };
  const candidateOverrides: Record<string, string> = { [candidateDate]: requesterShift };

  const requesterStreak = getMaxConsecutiveWorkStreak(roster, requester, requesterOverrides, requesterDate, 7);
  const candidateStreak = getMaxConsecutiveWorkStreak(roster, candidate, candidateOverrides, candidateDate, 7);
  if (requesterStreak.max > 5) {
    return { valid: false, blockers: [`${requester} would be on ${requesterStreak.max} consecutive working days`, `(${requesterStreak.startDate ?? requesterDate} → ${requesterStreak.endDate ?? requesterDate}), which exceeds the 5-day maximum.`] };
  }
  if (candidateStreak.max > 5) {
    return { valid: false, blockers: [`${candidate} would be on ${candidateStreak.max} consecutive working days`, `(${candidateStreak.startDate ?? candidateDate} → ${candidateStreak.endDate ?? candidateDate}), which exceeds the 5-day maximum.`] };
  }

  // 9. Redundant same-day-adjacency check
  if (wouldCreateConsecutiveShiftForSwap(roster, requester, requesterDate, candidateShift) ||
    wouldCreateConsecutiveShiftForSwap(roster, candidate, candidateDate, requesterShift)) {
    return { valid: false, blockers: ['This swap would force either person into back-to-back shifts with no break, which is not allowed.'] };
  }

  // 10. Same-shift check
  if (requesterShift === candidateShift) {
    return { valid: false, blockers: [`Same-shift swaps like ${shiftLabel(requesterShift)} ↔ ${shiftLabel(candidateShift)} are pointless and should not be suggested.`] };
  }

  return { valid: true, blockers: [] };
}

export function classifySwapStructure(
  roster: RosterData,
  requester: string,
  requesterDate: string,
  candidate: string,
  candidateDate: string,
): SwapStructureResult {
  const requesterShift = shiftKey(roster.rows[requester]?.[requesterDate] ?? 'OFF');
  const candidateShift = shiftKey(roster.rows[candidate]?.[candidateDate] ?? 'OFF');

  if (requesterShift === 'OFF' || candidateShift === 'OFF') {
    return { type: 'OFF_DAY', applicableDates: [requesterDate], blockedDates: [] };
  }

  const requesterBlock = detectShiftBlock(roster, requester, requesterDate);
  const candidateBlock = detectShiftBlock(roster, candidate, candidateDate);
  const requesterDates = blockDateRange(requesterBlock);
  const candidateDates = new Set(blockDateRange(candidateBlock));
  const overlapDates = requesterDates.filter((date) => candidateDates.has(date));
  const sameExactBlock = requesterBlock.startDate === candidateBlock.startDate
    && requesterBlock.endDate === candidateBlock.endDate
    && requesterShift === candidateShift;

  if (sameExactBlock) {
    return { type: 'WHOLE_BLOCK', applicableDates: overlapDates, blockedDates: [] };
  }

  if (overlapDates.length > 1 && (overlapDates.length < requesterDates.length || overlapDates.length < candidateDates.size)) {
    return {
      type: 'PARTIAL_BLOCK',
      applicableDates: overlapDates,
      blockedDates: requesterDates.filter((date) => !candidateDates.has(date)),
    };
  }

  if (overlapDates.length > 1) {
    return { type: 'WHOLE_BLOCK', applicableDates: overlapDates, blockedDates: [] };
  }

  return { type: 'DAY_ONLY', applicableDates: [requesterDate], blockedDates: [requesterDate] };
}

export function calculateSwapBalance(
  roster: RosterData,
  requester: string,
  requesterDate: string,
  candidate: string,
  candidateDate: string,
): SwapBalanceResult {
  const requesterShift = shiftKey(roster.rows[requester]?.[requesterDate] ?? 'OFF');
  const candidateShift = shiftKey(roster.rows[candidate]?.[candidateDate] ?? 'OFF');
  const requesterBlock = detectShiftBlock(roster, requester, requesterDate);
  const candidateBlock = detectShiftBlock(roster, candidate, candidateDate);
  const requesterDates = blockDateRange(requesterBlock);
  const candidateDates = blockDateRange(candidateBlock);
  const overlapDates = requesterDates.filter((date) => candidateDates.includes(date));

  const debt = buildBlockDebt(roster, requester, candidate, requesterDate);

  return {
    hoursGiven: (shiftHourValues[requesterShift] ?? 0) * Math.max(1, overlapDates.length || 1),
    hoursReceived: (shiftHourValues[candidateShift] ?? 0) * Math.max(1, overlapDates.length || 1),
    debtDays: debt?.debtDays ?? 0,
    debtOwner: debt?.whoOwes,
    owedTo: debt?.owedTo,
    possible: debt !== null,
    suggestedDates: debt?.suggestedOffDates ?? [],
  };
}

export function scoreSwapResult(
  roster: RosterData,
  requester: string,
  requesterDate: string,
  candidate: string,
  candidateDate: string,
): SwapCompatibilityResult {
  const legality = evaluateSwapLegality(roster, requester, requesterDate, candidate, candidateDate);
  if (!legality.valid) {
    return { compatible: false, score: 0, severity: 'blocked', reasons: legality.blockers };
  }

  const requesterShift = shiftKey(roster.rows[requester]?.[requesterDate] ?? 'OFF');
  const candidateShift = shiftKey(roster.rows[candidate]?.[candidateDate] ?? 'OFF');
  const requesterIsWorkBlock = requesterShift !== 'OFF' && requesterShift !== 'H8';
  const candidateIsWorkBlock = candidateShift !== 'OFF' && candidateShift !== 'H8';

  const requesterBlock = detectShiftBlock(roster, requester, requesterDate);
  const candidateBlock = detectShiftBlock(roster, candidate, candidateDate);
  const requesterDates = new Set(blockDateRange(requesterBlock));
  const candidateDates = new Set(blockDateRange(candidateBlock));
  const overlapDates = requesterIsWorkBlock && candidateIsWorkBlock
    ? [...requesterDates].filter((date) => candidateDates.has(date))
    : [];
  const sameExactBlock = requesterIsWorkBlock && candidateIsWorkBlock
    && requesterBlock.startDate === candidateBlock.startDate
    && requesterBlock.endDate === candidateBlock.endDate
    && requesterShift === candidateShift;
  const sameBlockDifferentShift = requesterIsWorkBlock && candidateIsWorkBlock
    && requesterBlock.startDate === candidateBlock.startDate
    && requesterBlock.endDate === candidateBlock.endDate
    && requesterShift !== candidateShift;

  if (sameBlockDifferentShift) {
    return {
      compatible: true,
      score: 100,
      severity: 'good',
      reasons: [
        `${shiftLabel(requesterShift)} block ↔ ${shiftLabel(candidateShift)} block on the same date range is a valid full-block swap.`,
      ],
    };
  }

  if (requesterShift === candidateShift) {
    return {
      compatible: false,
      score: 0,
      severity: 'blocked',
      reasons: [`Same-shift swaps like ${shiftLabel(requesterShift)} ↔ ${shiftLabel(candidateShift)} are pointless and should not be suggested.`],
    };
  }

  if (overlapDates.length > 1 && !sameExactBlock) {
    const blockLenMatch = blockDateRange(requesterBlock).length === blockDateRange(candidateBlock).length;
    const sameBoundary = requesterBlock.startDate === candidateBlock.startDate && requesterBlock.endDate === candidateBlock.endDate;
    const isSameShift = requesterShift === candidateShift;
    const blockTradeLabel = `${shiftLabel(requesterShift)} ${requesterBlock.startDate.slice(5)} → ${requesterBlock.endDate.slice(5)} vs ${shiftLabel(candidateShift)} ${candidateBlock.startDate.slice(5)} → ${candidateBlock.endDate.slice(5)}`;

    const tradeDates = sameBoundary ? blockDateRange(requesterBlock) : overlapDates;
    const simulated = simulatePartialBlockTradeStreaks(
      roster,
      requester,
      candidate,
      tradeDates,
      candidateShift,
      requesterShift,
      requesterDate,
    );

    if (simulated.requester.max > 5 || simulated.candidate.max > 5) {
      return {
        compatible: false,
        score: 0,
        severity: 'blocked',
        reasons: [
          `${isSameShift ? 'Same-shift' : 'Different-shift'} block trade overlaps here: ${blockTradeLabel}.`,
          `This block swap would create more than 5 consecutive working days for one employee, so it is not allowed.`,
        ],
      };
    }

    if (!isPartialTradeRestValid(roster, requester, tradeDates, candidateShift)
      || !isPartialTradeRestValid(roster, candidate, tradeDates, requesterShift)) {
      return {
        compatible: false,
        score: 0,
        severity: 'blocked',
        reasons: [
          `${isSameShift ? 'Same-shift' : 'Different-shift'} block trade overlaps here: ${blockTradeLabel}.`,
          `This swap would force either person into back-to-back shifts with no break, which is not allowed.`,
        ],
      };
    }

    if (sameBoundary && blockLenMatch) {
      return {
        compatible: true,
        score: 100,
        severity: 'good',
        reasons: [
          `${isSameShift ? 'Same-shift' : 'Different-shift'} block trade is valid here: ${blockTradeLabel}.`,
          `This is a clean whole-block swap with no debt and no date-range mismatch.`,
        ],
      };
    }

    return {
      compatible: true,
      score: 82,
      severity: 'warning',
      reasons: [
        `${isSameShift ? 'Same-shift' : 'Different-shift'} block trade overlaps here: ${blockTradeLabel}.`,
        `This has a different block boundary or implied debt, so it should be reviewed before approval.`,
      ],
    };
  }

  if (requesterShift === candidateShift) {
    return {
      compatible: false,
      score: 0,
      severity: 'blocked',
      reasons: [`Same-shift swaps like ${shiftLabel(requesterShift)} ↔ ${shiftLabel(candidateShift)} are pointless and should not be suggested.`],
    };
  }

  const reasons: string[] = [];
  let score = 100;
  let blocked = false;

  if (requesterShift !== 'OFF' && candidateShift !== 'OFF') {
    const allowed = SWAP_COMPATIBLE_SHIFTS[requesterShift];
    if (!allowed || !allowed.includes(candidateShift)) {
      score -= 12;
      reasons.push(`${shiftLabel(requesterShift)} is not a common pairing with ${shiftLabel(candidateShift)}, but it can still be considered if both people agree.`);
    } else {
      score -= 8;
      reasons.push(`Different shift types (${shiftLabel(requesterShift)} ↔ ${shiftLabel(candidateShift)}) are allowed as a personal preference trade.`);
    }
    if (requesterShift === 'MID' || candidateShift === 'MID') {
      score -= 8;
      reasons.push('MID overlaps with both morning and afternoon coverage — double-check staffing.');
    }
  } else {
    score -= 10;
    reasons.push('One side is off — this is a one-way coverage swap, not an equal trade.');
  }

  const hoursDelta = (shiftHourValues[candidateShift] ?? 0) - (shiftHourValues[requesterShift] ?? 0);
  if (hoursDelta !== 0) {
    score -= 15;
    reasons.push(`Changes monthly hours by ${hoursDelta > 0 ? '+' : ''}${hoursDelta}h for ${requester} (and the opposite for ${candidate}).`);
  }

  const worstRest = Math.min(
    minAdjacentRestHours(roster, requester, requesterDate, candidateShift),
    minAdjacentRestHours(roster, candidate, candidateDate, requesterShift),
  );
  if (worstRest <= HARD_MIN_REST_HOURS) {
    blocked = true;
    reasons.push(`Shifts would overlap with only ${Math.max(0, Math.round(worstRest))}h between them — no real rest. This includes impossible 16h consecutive work patterns like ${shiftLabel(requesterShift)} → ${shiftLabel(candidateShift)}.`);
  } else if (worstRest < PREFERRED_MIN_REST_HOURS) {
    score -= 15;
    reasons.push(`Rest gap of ${Math.round(worstRest)}h is under the preferred ${PREFERRED_MIN_REST_HOURS}h minimum.`);
  }

  const coverage = checkCoverageAfterSwap(roster, requester, requesterDate, requesterShift, candidate, candidateDate, candidateShift);
  if (coverage.warnings.length > 0) {
    score -= 8 * coverage.warnings.length;
    reasons.push(...coverage.warnings);
  }

  score = Math.max(0, Math.min(100, score));
  const severity: SwapSeverity = blocked ? 'blocked' : score >= 85 ? 'good' : score >= 60 ? 'warning' : 'warning';

  if (reasons.length === 0) reasons.push('No conflicts found.');

  return { compatible: severity !== 'blocked', score: blocked ? 0 : score, severity, reasons };
}

/**
 * Evaluates trading `requester`'s shift on `requesterDate` with `candidate`'s
 * shift on `candidateDate` (the same date for a same-day swap, different
 * dates for a cross-day trade). Returns a score/severity plus human-readable
 * reasons so the UI can explain the verdict.
 */
export function evaluateShiftSwapCompatibility(
  roster: RosterData,
  requester: string,
  requesterDate: string,
  candidate: string,
  candidateDate: string,
): SwapCompatibilityResult {
  return scoreSwapResult(roster, requester, requesterDate, candidate, candidateDate);
}

/** Ranks every other employee as a same-day swap partner for `requester` on `isoDate`, best first. */
export function findSwapCandidates(roster: RosterData, requester: string, isoDate: string): SwapCandidateResult[] {
  const requesterShift = shiftKey(roster.rows[requester]?.[isoDate] ?? 'OFF');

  return roster.employees
    .filter((employee) => employee !== requester)
    .filter((employee) => shiftKey(roster.rows[employee]?.[isoDate] ?? 'OFF') !== 'MID')
    .filter(() => requesterShift !== 'MID')
    .map((employee) => ({
      employee,
      shift: shiftKey(roster.rows[employee]?.[isoDate] ?? 'OFF'),
      ...evaluateShiftSwapCompatibility(roster, requester, isoDate, employee, isoDate),
    }))
    .sort((a, b) => b.score - a.score || a.employee.localeCompare(b.employee));
}

/** Plain-text swap request the requester can copy and send to the candidate. */
export function buildSwapRequestMessage(
  requester: string,
  requesterShift: string,
  requesterDate: string,
  candidate: string,
  candidateShift: string,
  candidateDate: string,
  result: SwapCompatibilityResult,
): string {
  const severityLabel = result.severity === 'good' ? 'Good' : result.severity === 'warning' ? 'Needs review' : 'Blocked';
  return [
    `Hi ${candidate}, can you swap with me?`,
    `${requester} has ${shiftLabel(requesterShift)} on ${requesterDate}.`,
    `You have ${shiftLabel(candidateShift)} on ${candidateDate}.`,
    `Compatibility: ${severityLabel}.`,
    ...result.reasons.map((reason) => `- ${reason}`),
  ].join('\n');
}

// ---------------------------------------------------------------------------
// Day navigation — builds a ShiftEvent for any date the roster covers, even if
// the employee has no explicit entry for it (treated as a day off).
// ---------------------------------------------------------------------------

export function eventForIso(roster: RosterData, employee: string, isoDate: string): ShiftEvent {
  const shift = shiftKey(roster.rows[employee]?.[isoDate] ?? 'OFF');
  return { id: `${employee}-${isoDate}`, isoDate, shift, date: new Date(`${isoDate}T00:00:00`) };
}

// ---------------------------------------------------------------------------
// Debt-aware rebalancing — calculates imbalance between two employees and
// suggests specific days to trade for rebalancing.
//
// Scenario: You work 2-3-4, they work 1-2-3-4. You take their whole block
// (1-2-3-4) and they take yours (2-3-4). Net: you received 4 days, gave 3.
// You owe them 1 day. The system should suggest which day you could give them
// from your roster to settle the debt.
// ---------------------------------------------------------------------------

export type BlockTrade = {
  giver: string;      // employee who gives away these dates
  receiver: string;   // employee who receives these dates
  dates: string[];    // ISO dates
};

export type TradeBalance = {
  employeeA: string;
  employeeB: string;
  aNetDays: number;     // + = A received more (owes B), - = A gave more (B owes A)
  aNetHours: number;    // + = A received more hours, - = A gave more hours
  imbalanceDays: number; // absolute value of net days (always >= 0)
  whoOwes: string;      // employee who owes (received more than they gave)
  owedTo: string;       // employee who is owed (gave more than they received)
};

export type RebalancingSuggestion = {
  date: string;
  from: string;         // current holder of shift (would give it away)
  to: string;           // would receive the shift
  fromShift: string;    // shift the giver currently has
  toShift: string;      // shift the receiver currently has
  compatibility: SwapCompatibilityResult;
  settlesDays: number;  // usually 1
  settlesHours: number; // hour value of the shift being given
};

/**
 * Calculates net trade balance from a list of proposed block trades.
 * Positive aNetDays/aNetHours means employeeA is up (received more → owes B).
 */
export function calculateTradeBalance(
  roster: RosterData,
  trades: BlockTrade[],
): TradeBalance | null {
  if (!roster || trades.length === 0) return null;

  // Gather the two unique employees involved
  const involved = new Set<string>();
  trades.forEach((t) => { involved.add(t.giver); involved.add(t.receiver); });
  const employees = Array.from(involved);
  if (employees.length !== 2) return null;

  const [employeeA, employeeB] = employees;

  let aReceivedDays = 0;
  let aReceivedHours = 0;
  let aGaveDays = 0;
  let aGaveHours = 0;

  trades.forEach((trade) => {
    trade.dates.forEach((isoDate) => {
      const shift = shiftKey(roster.rows[trade.giver]?.[isoDate] ?? 'OFF');
      const hours = shiftHourValues[shift] ?? 0;
      const isWorkDay = shift !== 'OFF' && shift !== 'H8';

      if (trade.giver === employeeA && trade.receiver === employeeB) {
        if (isWorkDay) aGaveDays += 1;
        aGaveHours += hours;
      } else if (trade.giver === employeeB && trade.receiver === employeeA) {
        if (isWorkDay) aReceivedDays += 1;
        aReceivedHours += hours;
      }
    });
  });

  const aNetDays = aReceivedDays - aGaveDays;
  const aNetHours = aReceivedHours - aGaveHours;
  const imbalanceDays = Math.abs(aNetDays);
  const whoOwes = aNetDays > 0 || (aNetDays === 0 && aNetHours > 0) ? employeeA : employeeB;
  const owedTo = whoOwes === employeeA ? employeeB : employeeA;

  return {
    employeeA,
    employeeB,
    aNetDays,
    aNetHours,
    imbalanceDays,
    whoOwes,
    owedTo,
  };
}

/**
 * Suggests same-day rebalancing trades where `owes` gives shifts to `owedTo`.
 * Only returns trades that are compatible (not blocked), sorted best first.
 *
 * @param targetDays  — how many days of debt to settle (0 = ignore day count)
 * @param targetHours — how many hours of debt to settle (0 = ignore hours)
 * @param excludeDates — dates already in the proposed trade (optional)
 */
export function suggestRebalancingTrades(
  roster: RosterData,
  owes: string,
  owedTo: string,
  targetDays: number = 0,
  targetHours: number = 0,
  excludeDates: string[] = [],
  monthContext?: { month: number; year: number },
  preferredShift?: string,
): RebalancingSuggestion[] {
  if (!roster || !roster.employees.includes(owes) || !roster.employees.includes(owedTo)) return [];

  const excludeSet = new Set(excludeDates);
  const suggestions: RebalancingSuggestion[] = [];

  for (const { isoDate } of roster.dateColumns) {
    if (excludeSet.has(isoDate)) continue;
    if (monthContext && !isSameMonth(isoDate, monthContext.month, monthContext.year)) continue;

    const fromShift = shiftKey(roster.rows[owes]?.[isoDate] ?? 'OFF');
    if (fromShift === 'OFF' || fromShift === 'H8') continue;

    const toShift = shiftKey(roster.rows[owedTo]?.[isoDate] ?? 'OFF');

    const compatibility = evaluateShiftSwapCompatibility(roster, owedTo, isoDate, owes, isoDate);
    if (!compatibility.compatible) continue;

    const settlesDays = 1;
    const settlesHours = shiftHourValues[fromShift] ?? 0;

    if (targetDays > 0 && settlesDays > targetDays) continue;
    if (targetHours > 0 && settlesHours > targetHours) continue;

    suggestions.push({
      date: isoDate,
      from: owes,
      to: owedTo,
      fromShift,
      toShift,
      compatibility,
      settlesDays,
      settlesHours,
    });
  }

  const preferred = preferredShift ? shiftKey(preferredShift) : null;
  return suggestions.sort((a, b) => {
    const aSameType = preferred ? (a.fromShift === preferred ? 2 : 0) : 0;
    const bSameType = preferred ? (b.fromShift === preferred ? 2 : 0) : 0;
    const sameTypeDiff = bSameType - aSameType;
    if (sameTypeDiff !== 0) return sameTypeDiff;

    const aEdge = isBlockBoundary(roster, a.from, a.date) ? 1 : 0;
    const bEdge = isBlockBoundary(roster, b.from, b.date) ? 1 : 0;
    const edgeDiff = bEdge - aEdge;
    if (edgeDiff !== 0) return edgeDiff;

    const scoreDiff = b.compatibility.score - a.compatibility.score;
    if (scoreDiff !== 0) return scoreDiff;
    return a.date.localeCompare(b.date);
  });
}

/**
 * Full roster scan between two employees. Finds every same-day trade
 * opportunity in both directions, scored by compatibility.
 */
export function findAllRebalancingOpportunities(
  roster: RosterData,
  employeeA: string,
  employeeB: string,
  context?: { month?: number; year?: number; preferredShift?: string },
): RebalancingSuggestion[] {
  if (!roster || !roster.employees.includes(employeeA) || !roster.employees.includes(employeeB)) return [];

  const monthContext = context && typeof context.month === 'number' && typeof context.year === 'number'
    ? { month: context.month, year: context.year }
    : undefined;

  const aToB = suggestRebalancingTrades(
    roster,
    employeeA,
    employeeB,
    0,
    0,
    [],
    monthContext,
    context?.preferredShift,
  );
  const bToA = suggestRebalancingTrades(
    roster,
    employeeB,
    employeeA,
    0,
    0,
    [],
    monthContext,
    context?.preferredShift,
  );

  return [...aToB, ...bToA].sort((a, b) => {
    const aScore = (a.fromShift === context?.preferredShift ? 3 : 0) + (isBlockBoundary(roster, a.from, a.date) ? 2 : 0) + a.compatibility.score;
    const bScore = (b.fromShift === context?.preferredShift ? 3 : 0) + (isBlockBoundary(roster, b.from, b.date) ? 2 : 0) + b.compatibility.score;
    const scoreDiff = bScore - aScore;
    if (scoreDiff !== 0) return scoreDiff;
    return a.date.localeCompare(b.date);
  });
}

/** Plain-text rebalancing request the giver can copy and send. */
export function buildRebalancingMessage(
  to: string,
  date: string,
  shift: string,
  result: SwapCompatibilityResult,
): string {
  const severityLabel = result.severity === 'good' ? 'Good' : result.severity === 'warning' ? 'Needs review' : 'Blocked';
  return [
    `Hi ${to}, to rebalance our shifts:`,
    `I can give you my ${shiftLabel(shift)} on ${date}.`,
    `Compatibility: ${severityLabel}.`,
    ...result.reasons.map((reason) => `- ${reason}`),
  ].join('\n');
}

// ---------------------------------------------------------------------------
// "Liquid Glass" surface tokens — shared material used by App.tsx (nav bar,
// cards) and DayDetailsModal.tsx (sheet) so the frosted-glass look stays
// consistent instead of drifting between hand-written Tailwind strings.
// ---------------------------------------------------------------------------

export const GLASS_CARD =
  'rounded-[28px] bg-white/70 shadow-[0_1px_0_0_rgba(255,255,255,0.6)_inset,0_10px_30px_-12px_rgba(0,0,0,0.14)] ring-1 ring-black/[0.04] backdrop-blur-xl dark:bg-white/[0.05] dark:shadow-[0_1px_0_0_rgba(255,255,255,0.06)_inset,0_10px_30px_-12px_rgba(0,0,0,0.55)] dark:ring-white/[0.08]';

export const GLASS_NAV =
  'rounded-[28px] bg-white/50 shadow-[0_1px_0_0_rgba(255,255,255,0.6)_inset,0_14px_38px_-10px_rgba(0,0,0,0.2)] ring-1 ring-black/[0.05] backdrop-blur-2xl dark:bg-white/[0.06] dark:shadow-[0_1px_0_0_rgba(255,255,255,0.07)_inset,0_14px_38px_-12px_rgba(0,0,0,0.6)] dark:ring-white/[0.1]';

export const GLASS_SHEET =
  'rounded-t-[32px] bg-white/90 shadow-[0_1px_0_0_rgba(255,255,255,0.7)_inset,0_-24px_60px_-18px_rgba(0,0,0,0.32)] ring-1 ring-black/[0.05] backdrop-blur-2xl dark:bg-zinc-900/90 dark:shadow-[0_1px_0_0_rgba(255,255,255,0.05)_inset,0_-24px_60px_-18px_rgba(0,0,0,0.65)] dark:ring-white/[0.08]';

// ---------------------------------------------------------------------------
// Enhanced swap candidates with block trading & rebalancing suggestions
// ---------------------------------------------------------------------------

export type BlockDebtInfo = {
  debtDays: number;
  whoOwes: string;
  owedTo: string;
  suggestedOffDates: string[];
  debtReason: string;
};

export type RebalancingOpportunity = {
  date: string;
  direction: 'take' | 'give';
  shift: string;
  compatibility: SwapCompatibilityResult;
  reason: string;
  debt?: BlockDebtInfo;
};

export type ShiftBlock = {
  shift: string;
  startDate: string;
  endDate: string;
};

export type EnhancedSwapCandidate = SwapCandidateResult & {
  requesterBlock: ShiftBlock;
  candidateBlock: ShiftBlock;
  overlapDates: string[];
  blockResult: SwapCompatibilityResult;
  rebalancing: RebalancingOpportunity[];
  blockDebt?: BlockDebtInfo;
  partialBlockNote?: string;
};

/**
 * Detects the shift block containing the given date, bounded by OFF/H8 days.
 * A block is a contiguous sequence of the same shift type, separated by OFF or H8 days.
 */
export function detectShiftBlock(
  roster: RosterData,
  employee: string,
  targetIsoDate: string,
  rows: Record<string, Record<string, string>> = roster.rows,
): ShiftBlock {
  const targetShift = shiftKey(rows[employee]?.[targetIsoDate] ?? 'OFF');

  // If the target day is OFF or H8, it's not part of a work block
  if (targetShift === 'OFF' || targetShift === 'H8') {
    return { shift: targetShift, startDate: targetIsoDate, endDate: targetIsoDate };
  }

  // Find the start of the block by going backward
  let startDate = targetIsoDate;
  let currentDate = addDaysIso(targetIsoDate, -1);
  while (true) {
    const shift = shiftKey(rows[employee]?.[currentDate] ?? 'OFF');
    // Stop when we hit an OFF/H8 day or run out of dates
    if (shift === 'OFF' || shift === 'H8' || !roster.dateColumns.some(c => c.isoDate === currentDate)) {
      break;
    }
    // Stop if shift type changes
    if (shift !== targetShift) {
      break;
    }
    startDate = currentDate;
    currentDate = addDaysIso(currentDate, -1);
  }

  // Find the end of the block by going forward
  let endDate = targetIsoDate;
  currentDate = addDaysIso(targetIsoDate, 1);
  while (true) {
    const shift = shiftKey(rows[employee]?.[currentDate] ?? 'OFF');
    // Stop when we hit an OFF/H8 day or run out of dates
    if (shift === 'OFF' || shift === 'H8' || !roster.dateColumns.some(c => c.isoDate === currentDate)) {
      break;
    }
    // Stop if shift type changes
    if (shift !== targetShift) {
      break;
    }
    endDate = currentDate;
    currentDate = addDaysIso(currentDate, 1);
  }

  return { shift: targetShift, startDate, endDate };
}

export function blockDateRange(block: ShiftBlock): string[] {
  const dates: string[] = [];
  const start = new Date(`${block.startDate}T00:00:00`);
  const end = new Date(`${block.endDate}T00:00:00`);
  const cursor = new Date(start);
  while (cursor <= end) {
    const iso = `${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, '0')}-${String(cursor.getDate()).padStart(2, '0')}`;
    dates.push(iso);
    cursor.setDate(cursor.getDate() + 1);
  }
  return dates;
}

function buildPartialBlockNote(
  roster: RosterData,
  requester: string,
  requesterDate: string,
  candidate: string,
  candidateDate: string,
): string | undefined {
  const requesterShift = shiftKey(roster.rows[requester]?.[requesterDate] ?? 'OFF');
  const candidateShift = shiftKey(roster.rows[candidate]?.[candidateDate] ?? 'OFF');
  if (requesterShift === 'OFF' || candidateShift === 'OFF' || requesterShift === 'H8' || candidateShift === 'H8') return undefined;

  const requesterBlock = detectShiftBlock(roster, requester, requesterDate);
  const candidateBlock = detectShiftBlock(roster, candidate, candidateDate);
  const requesterDates = blockDateRange(requesterBlock);
  const candidateDates = new Set(blockDateRange(candidateBlock));
  const overlapDates = requesterDates.filter((date) => candidateDates.has(date));

  if (overlapDates.length === 0) return undefined;

  const sameExactBlock = requesterBlock.startDate === candidateBlock.startDate
    && requesterBlock.endDate === candidateBlock.endDate
    && requesterShift === candidateShift;
  if (sameExactBlock) return undefined;

  const overlapStart = overlapDates[0];
  const overlapEnd = overlapDates[overlapDates.length - 1];
  if (overlapDates.length >= requesterDates.length) return undefined;

  return `${overlapStart.slice(5)} → ${overlapEnd.slice(5)} is applicable, rest is blocked.`;
}

function isDebtSettlementDay(roster: RosterData, employee: string, isoDate: string): boolean {
  const shift = shiftKey(roster.rows[employee]?.[isoDate] ?? 'OFF');
  if (shift !== 'OFF') return false;

  const targetDate = new Date(`${isoDate}T00:00:00`);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  return targetDate >= today;
}

function buildBlockDebt(
  roster: RosterData,
  receiver: string,
  giver: string,
  isoDate: string,
): BlockDebtInfo | null {
  const receiverShift = shiftKey(roster.rows[receiver]?.[isoDate] ?? 'OFF');
  const giverShift = shiftKey(roster.rows[giver]?.[isoDate] ?? 'OFF');
  if (receiverShift === 'OFF' || receiverShift === 'H8' || giverShift === 'OFF' || giverShift === 'H8') {
    return null;
  }

  const receiverBlock = detectShiftBlock(roster, receiver, isoDate);
  const giverBlock = detectShiftBlock(roster, giver, isoDate);
  const receiverDates = new Set(blockDateRange(receiverBlock));
  const giverDates = new Set(blockDateRange(giverBlock));

  const debtDays = Math.abs(receiverDates.size - giverDates.size);
  if (debtDays === 0) return null;

  // Receiver takes giver's block, Giver takes receiver's block.
  // If giver's block is longer, receiver receives more working days than given -> receiver owes giver days off.
  // If receiver's block is longer, giver receives more working days than given -> giver owes receiver days off.
  const whoOwes = giverDates.size > receiverDates.size ? receiver : giver;
  const owedTo = whoOwes === receiver ? giver : receiver;

  const monthStart = new Date(`${isoDate}T00:00:00`);
  const monthDates = roster.dateColumns
    .filter(({ isoDate: date }) => {
      const dateObj = new Date(`${date}T00:00:00`);
      return dateObj.getMonth() === monthStart.getMonth() && dateObj.getFullYear() === monthStart.getFullYear();
    })
    .map(({ isoDate: date }) => date)
    .filter((date) => isDebtSettlementDay(roster, whoOwes, date));

  const validDebtSettlementDates = monthDates.filter((date) => {
    // H8 is a holiday flag, never a real off-day settlement option for debt.
    if (shiftKey(roster.rows[whoOwes]?.[date] ?? 'OFF') === 'H8') return false;
    const compatibility = evaluateShiftSwapCompatibility(roster, whoOwes, date, owedTo, date);
    return compatibility.compatible && compatibility.severity !== 'blocked';
  });

  const preferred = validDebtSettlementDates.filter((date) =>
    date === receiverBlock.startDate || date === receiverBlock.endDate || date === giverBlock.startDate || date === giverBlock.endDate,
  );

  const suggestedOffDates = (preferred.length > 0 ? preferred : validDebtSettlementDates).slice(0, 2);
  const suggestionText = suggestedOffDates.length > 0
    ? ` Suggested off day${suggestedOffDates.length > 1 ? 's' : ''}: ${suggestedOffDates.join(', ')}.`
    : ' No off-day settlement is currently available in this month.';

  return {
    debtDays,
    whoOwes,
    owedTo,
    suggestedOffDates,
    debtReason: `${whoOwes} owes ${debtDays} day${debtDays === 1 ? '' : 's'} off to ${owedTo}.${suggestionText}`,
  };
}

/**
 * Finds all same-day swap candidates for an employee on a specific date,
 * enriched with block trading context and rebalancing suggestions.
 */
export function findEnhancedSwapCandidates(
  roster: RosterData,
  employee: string,
  isoDate: string,
  _monthContext?: { month: number; year: number },
): EnhancedSwapCandidate[] {
  const baseCandidates = findSwapCandidates(roster, employee, isoDate);
  const requesterBlock = detectShiftBlock(roster, employee, isoDate);

  return baseCandidates.map((candidate) => {
    const rebalancing: RebalancingOpportunity[] = [];

    // Detect the candidate's shift block
    const candidateBlock = detectShiftBlock(roster, candidate.employee, isoDate);

    // Find overlap dates in the two blocks
    const overlapDates: string[] = [];
    const knownDates = new Set(roster.dateColumns.map(c => c.isoDate));
    let currentDate = requesterBlock.startDate;
    while (currentDate <= requesterBlock.endDate && knownDates.has(currentDate)) {
      if (currentDate >= candidateBlock.startDate && currentDate <= candidateBlock.endDate) {
        overlapDates.push(currentDate);
      }
      currentDate = addDaysIso(currentDate, 1);
    }

    const blockDebt = buildBlockDebt(roster, employee, candidate.employee, isoDate) ?? undefined;
    const partialBlockNote = buildPartialBlockNote(roster, employee, isoDate, candidate.employee, isoDate);

    return {
      ...candidate,
      requesterBlock,
      candidateBlock,
      overlapDates,
      blockResult: candidate,
      rebalancing,
      blockDebt,
      partialBlockNote,
    };
  });
}

/**
 * Builds an enhanced swap request message with the whole-block outcome and debt summary.
 */
export function buildEnhancedSwapRequestMessage(
  requester: string,
  requesterShift: string,
  requesterDate: string,
  candidate: string,
  candidateShift: string,
  candidateDate: string,
  result: SwapCompatibilityResult,
  context?: {
    requesterBlock?: { shift: string; startDate: string; endDate: string };
    candidateBlock?: { shift: string; startDate: string; endDate: string };
    overlapDates?: string[];
    blockDebt?: BlockDebtInfo;
  },
): string {
  const severityLabel = result.severity === 'good' ? 'Good' : result.severity === 'warning' ? 'Needs review' : 'Blocked';
  const lines = [
    `Hi ${candidate}, can you swap with me?`,
    `${requester} has ${shiftLabel(requesterShift)} on ${requesterDate}.`,
    `You have ${shiftLabel(candidateShift)} on ${candidateDate}.`,
    `Compatibility: ${severityLabel}.`,
    ...result.reasons.map((reason) => `- ${reason}`),
  ];

  if (context?.requesterBlock && context?.candidateBlock) {
    lines.push('');
    lines.push('Whole-block outcome:');
    lines.push(`- If I take your whole block (${shiftLabel(context.candidateBlock.shift)} ${context.candidateBlock.startDate.slice(5)} → ${context.candidateBlock.endDate.slice(5)}), the trade is based on the full block difference.`);
    if (context.blockDebt) {
      lines.push(`- ${context.blockDebt.debtReason}`);
    }
  }

  return lines.join('\n');
}