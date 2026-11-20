import type { RosterData, GeneratedScheduleResult } from './types';
import { shiftKey } from './constants';
import { buildMonthDateColumns, countLegalWorkingDaysInMonth } from './dateUtils';
import { validateScheduleCompliance } from './validator';

const WORKING_SHIFTS = new Set(['M', 'A', 'N', 'MID']);

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
  if (datesInRosterForMonth.length < daysInMonth) return false;

  // A month with only H8/leave/OFF entries (no M/A/N/MID working shifts) is a pre-planned
  // blank/leave month, not a completed historical schedule.
  if (roster.rows) {
    for (const emp of Object.keys(roster.rows)) {
      const empRow = roster.rows[emp];
      if (!empRow) continue;
      for (const col of datesInRosterForMonth) {
        const s = shiftKey(empRow[col.isoDate] ?? '');
        if (WORKING_SHIFTS.has(s)) {
          return true;
        }
      }
    }
  }
  return false;
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
