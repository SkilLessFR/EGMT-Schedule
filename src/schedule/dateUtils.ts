import {
  MONTHLY_WORKING_DAYS_2026,
  ROMANIAN_HOLIDAYS,
  FIXED_STAFF_EXEMPT_OFF_DATES
} from './constants';

export function getIsoWeekNumber(date: Date): number {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d.getTime() - yearStart.getTime()) / 86400000) + 1) / 7);
}

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

export function isFixedStaffOffDate(date: Date, isoDate: string): boolean {
  const dow = date.getDay();
  return dow === 0 || dow === 6 || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
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
