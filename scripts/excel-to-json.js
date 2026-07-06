#!/usr/bin/env node
const fs = require('node:fs');
const path = require('node:path');
const XLSX = require('xlsx');

const monthNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
const dateHeaderPattern = /^\s*(\d{1,2})\s*[-\s/]\s*([A-Za-z]{3,})\b/;
const knownShiftCodes = new Set(['MID', 'OFF', 'M', 'A', 'N', 'H8']);

function toIso(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}
function parseHeaderDate(text, year) {
  const match = text.match(dateHeaderPattern);
  if (!match) return null;

  const day = Number(match[1]);
  const monthIndex = monthNames.findIndex((month) => match[2].toLowerCase().startsWith(month));
  if (monthIndex < 0 || day < 1 || day > 31) return null;

  const date = new Date(year, monthIndex, day);
  if (date.getMonth() !== monthIndex || date.getDate() !== day) return null;

  return date;
}
function cellText(sheet, rowIndex, columnIndex) {
  const cell = sheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })];
  return String(cell?.w ?? cell?.v ?? '').trim();
}
function normalizedCode(value) {
  return value.trim().toUpperCase();
}
function isKnownShiftCode(value) {
  return knownShiftCodes.has(normalizedCode(value));
}
function likelyName(value) {
  const text = String(value ?? '').trim();
  return Boolean(text) && !isKnownShiftCode(text) && /^[\p{L} .'-]{3,}$/u.test(text) && !/employee|equipment|morning|night|after|shift/i.test(text);
}
function rowHasShift(sheet, rowIndex, dateColumns) {
  return dateColumns.some(({ index }) => isKnownShiftCode(cellText(sheet, rowIndex, index)));
}
function inferEmployeeColumn(sheet, range, headerRow, dateColumns) {
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
function parseWorkbook(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1:A1');
  const currentYear = new Date().getFullYear();

  let headerRow = -1;
  let dateColumns = [];

  for (let rowIndex = range.s.r; rowIndex <= Math.min(range.e.r, range.s.r + 11); rowIndex += 1) {
    const detected = [];

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
  const employees = [];
  const rows = {};

  for (let rowIndex = headerRow + 1; rowIndex <= range.e.r; rowIndex += 1) {
    const employee = cellText(sheet, rowIndex, employeeCol);
    if (!likelyName(employee)) continue;

    const shifts = {};
    dateColumns.forEach(({ index, isoDate }) => {
      const value = cellText(sheet, rowIndex, index);
      if (value) shifts[isoDate] = value;
    });

    if (Object.keys(shifts).length) {
      employees.push(employee);
      rows[employee] = shifts;
    }
  }

  if (!employees.length) throw new Error('No employee rows were found below the date header.');

  const monthCounts = dateColumns.reduce((acc, column) => {
    acc[column.date.getMonth()] = (acc[column.date.getMonth()] ?? 0) + 1;
    return acc;
  }, {});
  const month = Number(Object.entries(monthCounts).sort((a, b) => b[1] - a[1])[0][0]);
  const year = dateColumns.find((column) => column.date.getMonth() === month)?.date.getFullYear() ?? currentYear;

  return {
    month,
    year,
    employees,
    rows,
    dateColumns: dateColumns.map(({ index, date, isoDate }) => ({ index, date: toIso(date), isoDate })),
  };
}

const input = process.argv[2];
if (!input) {
  console.error('Usage: npm run import -- path/to/roster.xlsx');
  process.exit(1);
}

const workbook = XLSX.readFile(input, { cellDates: false });
const schedule = parseWorkbook(workbook);
schedule.fileName = path.basename(input);
const output = path.resolve('public/schedule.json');
fs.mkdirSync(path.dirname(output), { recursive: true });
fs.writeFileSync(output, `${JSON.stringify(schedule, null, 2)}\n`);
console.log(`Wrote ${output}`);