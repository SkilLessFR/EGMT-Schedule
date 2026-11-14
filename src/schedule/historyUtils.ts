import type { RosterData } from './types';
import { shiftKey } from './constants';

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

/**
 * Inspect last shifts of previous month to ensure seamless continuity
 */
export function getEmployeeHistoryTail(
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
export function extractRawHistoryTail(
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
