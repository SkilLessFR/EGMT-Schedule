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

export function colorFor(shift: string) { return shiftColors[shift.toUpperCase()] ?? fallbackColor; }
export function solidColorFor(shift: string) { return shiftSolid[shift.toUpperCase()] ?? fallbackSolid; }
export function shiftKey(shift: string) { return shift.trim().toUpperCase(); }
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
      const code = shiftKey(roster.rows[employee]?.[isoDate] ?? '');
      if (code === 'M') daily.morning.push(employee);
      if (code === 'A') daily.afternoon.push(employee);
      if (code === 'N') daily.night.push(employee);
      if (code === 'MID') daily.mid.push(employee);
      if (code === 'OFF') daily.off.push(employee);
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
// Day navigation — builds a ShiftEvent for any date the roster covers, even if
// the employee has no explicit entry for it (treated as a day off).
// ---------------------------------------------------------------------------

export function eventForIso(roster: RosterData, employee: string, isoDate: string): ShiftEvent {
  const shift = roster.rows[employee]?.[isoDate] ?? 'OFF';
  return { id: `${employee}-${isoDate}`, isoDate, shift, date: new Date(isoDate) };
}