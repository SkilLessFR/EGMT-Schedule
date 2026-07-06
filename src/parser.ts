import * as XLSX from 'xlsx';
import type { RosterData, ShiftEvent } from './types';

const monthNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
const dateHeaderPattern = /^\s*(\d{1,2})\s*[-\s/]\s*([A-Za-z]{3,})\b/;
const knownShiftCodes = new Set(['MID', 'OFF', 'M', 'A', 'N', 'H8']);

function toIso(date: Date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function parseHeaderDate(text: string, year: number): Date | null {
  const match = text.match(dateHeaderPattern);
  if (!match) return null;

  const day = Number(match[1]);
  const monthIndex = monthNames.findIndex((month) => match[2].toLowerCase().startsWith(month));
  if (monthIndex < 0 || day < 1 || day > 31) return null;

  const date = new Date(year, monthIndex, day);
  if (date.getMonth() !== monthIndex || date.getDate() !== day) return null;

  return date;
}

function cellText(sheet: XLSX.WorkSheet, rowIndex: number, columnIndex: number) {
  const cell = sheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
  return String(cell?.w ?? cell?.v ?? '').trim();
}

function normalizedCode(value: string) {
  return value.trim().toUpperCase();
}

function isKnownShiftCode(value: string) {
  return knownShiftCodes.has(normalizedCode(value));
}

function likelyName(value: unknown) {
  const text = String(value ?? '').trim();
  return Boolean(text) && !isKnownShiftCode(text) && /^[\p{L} .'-]{3,}$/u.test(text) && !/employee|equipment|morning|night|after|shift/i.test(text);
}

function rowHasShift(sheet: XLSX.WorkSheet, rowIndex: number, dateColumns: RosterData['dateColumns']) {
  return dateColumns.some(({ index }) => isKnownShiftCode(cellText(sheet, rowIndex, index)));
}

function inferEmployeeColumn(sheet: XLSX.WorkSheet, range: XLSX.Range, headerRow: number, dateColumns: RosterData['dateColumns']) {
  for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
    if (/employee|name/i.test(cellText(sheet, headerRow, columnIndex))) return columnIndex;
  }

  const firstDateColumn = Math.min(...dateColumns.map(({ index }) => index));
  let bestColumn = range.s.c;
  let bestScore = -1;

  for (let columnIndex = range.s.c; columnIndex < firstDateColumn; columnIndex += 1) {
    let score = 0;

    for (let rowIndex = headerRow + 1; rowIndex <= range.e.r; rowIndex += 1) {
      const value = cellText(sheet, rowIndex, columnIndex);
      if (likelyName(value) && rowHasShift(sheet, rowIndex, dateColumns)) score += 1;
    }

    if (score > bestScore) {
      bestScore = score;
      bestColumn = columnIndex;
    }
  }

  return bestColumn;
}

export async function parseRoster(file: File): Promise<RosterData> {
  const workbook = XLSX.read(await file.arrayBuffer(), { type: 'array', cellDates: false });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1:A1');
  const currentYear = new Date().getFullYear();

  let headerRow = -1;
  let dateColumns: RosterData['dateColumns'] = [];

  for (let rowIndex = range.s.r; rowIndex <= Math.min(range.e.r, range.s.r + 11); rowIndex += 1) {
    const detected: RosterData['dateColumns'] = [];

    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      const date = parseHeaderDate(cellText(sheet, rowIndex, columnIndex), currentYear);
      if (date) detected.push({ index: columnIndex, date, isoDate: toIso(date) });
    }

    if (detected.length >= 3) {
      headerRow = rowIndex;
      dateColumns = detected;
      break;
    }
  }

  if (headerRow < 0 || dateColumns.length === 0) throw new Error('Could not find the date header row. Make sure the top row contains dates like 1-Jul, 2-Jul, etc.');

  const employeeCol = inferEmployeeColumn(sheet, range, headerRow, dateColumns);

  const employees: string[] = [];
  const parsedRows: RosterData['rows'] = {};

  for (let rowIndex = headerRow + 1; rowIndex <= range.e.r; rowIndex += 1) {
    const employee = cellText(sheet, rowIndex, employeeCol);
    if (!likelyName(employee)) continue;

    const shifts: Record<string, string> = {};
    dateColumns.forEach(({ index, isoDate }) => {
      const value = cellText(sheet, rowIndex, index);
      if (value) shifts[isoDate] = value;
    });

    if (Object.keys(shifts).length) {
      employees.push(employee);
      parsedRows[employee] = shifts;
    }
  }

  if (!employees.length) throw new Error('No employee rows were found below the date header.');

  const monthCounts = dateColumns.reduce<Record<number, number>>((acc, column) => {
    acc[column.date.getMonth()] = (acc[column.date.getMonth()] ?? 0) + 1;
    return acc;
  }, {});
  const month = Number(Object.entries(monthCounts).sort((a, b) => b[1] - a[1])[0][0]);
  const year = dateColumns.find((column) => column.date.getMonth() === month)?.date.getFullYear() ?? currentYear;
  return { month, year, employees, rows: parsedRows, dateColumns, fileName: file.name };
}

export function eventsForEmployee(roster: RosterData, employee: string): ShiftEvent[] {
  const shifts = roster.rows[employee] ?? {};
  return roster.dateColumns
    .filter(({ isoDate }) => shifts[isoDate])
    .map(({ date, isoDate }) => ({ id: `${employee}-${isoDate}`, date, isoDate, shift: shifts[isoDate] }));
}
