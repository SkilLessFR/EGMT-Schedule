import type ExcelJS from 'exceljs';
import { ROMANIAN_HOLIDAYS } from './constants';
import { countLegalWorkingDaysInMonth } from './dateUtils';

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
