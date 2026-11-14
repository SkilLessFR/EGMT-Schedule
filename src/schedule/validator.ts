import type {
  RosterData,
  ComplianceTier,
  ValidationOptions,
  DayCoverage,
  EmployeeMonthStats,
  ScheduleValidationReport
} from './types';
import {
  FIXED_MID_EMPLOYEES,
  FIXED_WEEKDAY_MA_EMPLOYEES,
  FIXED_STAFF_EXEMPT_OFF_DATES,
  ROMANIAN_HOLIDAYS,
  shiftKey
} from './constants';
import { getIsoWeekNumber, countLegalWorkingDaysInMonth } from './dateUtils';
import { getTierConfig, isCoverageFeasibleForTier } from '../relaxation';

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
  excludedEmployees: string[] = [],
  options?: ValidationOptions
): ScheduleValidationReport {
  const tier: ComplianceTier = options?.tier ?? 0;
  const tierConfig = getTierConfig(tier);
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

  const WORK_SHIFTS = new Set(['M', 'A', 'N', 'MID']);
  const ROTATING_SHIFTS = new Set(['M', 'A', 'N']);
  const LEAVE_OR_SPLIT_CODES = new Set(['ABS', 'H8', 'CM']);
  const VALID_CODES = new Set(['M', 'A', 'N', 'MID', 'OFF', 'ABS', 'H8']);
  const EXPECTED_NEXT_SHIFT: Record<string, string> = { A: 'M', M: 'N', N: 'A' };

  const normalizeShiftCode = (raw: string | undefined | null): string => {
    if (raw === undefined || raw === null) return 'OFF';
    const s = shiftKey(String(raw));
    return s === '' ? 'OFF' : s;
  };

  const parseIsoUtc = (isoDate: string): Date => {
    const [y, m, d] = isoDate.split('-').map(Number);
    return new Date(Date.UTC(y, m - 1, d));
  };

  const addDaysIso = (isoDate: string, delta: number): string => {
    const dt = parseIsoUtc(isoDate);
    dt.setUTCDate(dt.getUTCDate() + delta);
    return dt.toISOString().slice(0, 10);
  };

  const cols = [...(dateColumns || [])].sort((a, b) => a.isoDate.localeCompare(b.isoDate));
  const monthLen = cols.length;
  const monthIsos = cols.map((c) => c.isoDate);

  const calMeta = monthIsos.map((isoDate) => {
    const utcDt = parseIsoUtc(isoDate);
    const dow = utcDt.getUTCDay();
    const isWeekend = dow === 0 || dow === 6;
    const isFixedOff =
      isWeekend || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
    const isoWeek = getIsoWeekNumber(
      new Date(utcDt.getUTCFullYear(), utcDt.getUTCMonth(), utcDt.getUTCDate())
    );
    return { isoDate, dow, isWeekend, isFixedOff, isoWeek };
  });

  const fixedMonthWorkdays = calMeta.filter((m) => !m.isFixedOff).length;
  const employees = Object.keys(rows);

  // Extract contiguous <= 6-day prior history ending at monthIsos[0] - 1, preserving H8/ABS/CM codes.
  const knownHistDates = new Set<string>();
  if (priorRoster) {
    for (const dc of priorRoster.dateColumns || []) {
      if (dc?.isoDate) knownHistDates.add(dc.isoDate);
    }
    if (priorRoster.rows) {
      for (const rEmp of Object.keys(priorRoster.rows)) {
        const empRow = priorRoster.rows[rEmp];
        if (empRow && typeof empRow === 'object') {
          for (const iso of Object.keys(empRow)) knownHistDates.add(iso);
        }
      }
    }
  }
  const histTailIsos: string[] = [];
  if (monthLen > 0 && priorRoster) {
    const firstIso = monthIsos[0];
    for (let offset = 1; offset <= 6; offset++) {
      const prevIso = addDaysIso(firstIso, -offset);
      if (knownHistDates.has(prevIso)) {
        histTailIsos.unshift(prevIso);
      } else {
        break;
      }
    }
  }
  const hasHistory = histTailIsos.length > 0;
  const histLen = histTailIsos.length;

  // 1. Check Daily Coverage
  calMeta.forEach(({ isoDate, dow, isWeekend }) => {
    const c = { M: 0, A: 0, N: 0, MID: 0, OFF: 0, H8: 0, ABS: 0, other: 0 };
    const rotC = { M: 0, A: 0, N: 0 };
    const dayWarnings: string[] = [];

    employees.forEach((emp) => {
      if (excludedSet.has(emp)) return;
      const s = normalizeShiftCode(rows[emp]?.[isoDate]);
      if (s === 'M') c.M++;
      else if (s === 'A') c.A++;
      else if (s === 'N') c.N++;
      else if (s === 'MID') c.MID++;
      else if (s === 'H8') c.H8++;
      else if (s === 'ABS') c.ABS++;
      else if (s === 'OFF') c.OFF++;
      else c.other++;

      const isNonRotating =
        FIXED_MID_EMPLOYEES.includes(emp) || FIXED_WEEKDAY_MA_EMPLOYEES.includes(emp);
      if (!isNonRotating) {
        if (s === 'M') rotC.M++;
        else if (s === 'A') rotC.A++;
        else if (s === 'N') rotC.N++;
      }
    });

    let isUnderstaffed = false;
    let isOverstaffedNight = false;

    if (rotC.M < 1) {
      if (tierConfig.allowCoverageGapMA) {
        options?.relaxations?.push({
          tier: 3,
          rule: 'COVERAGE_GAP',
          isoDate,
          note: `Rotating Morning understaffed on ${isoDate}: rotM=${rotC.M}/1`,
        });
      } else {
        dayWarnings.push(`Morning understaffed (${rotC.M}/1 minimum)`);
        isUnderstaffed = true;
      }
    }

    if (c.A < 2) {
      if (c.A === 1 && tierConfig.allowCoverageGapMA) {
        options?.relaxations?.push({
          tier: 3,
          rule: 'COVERAGE_GAP',
          isoDate,
          note: `Afternoon understaffed on ${isoDate}: totalA=${c.A}/2`,
        });
      } else {
        dayWarnings.push(`Afternoon understaffed (${c.A}/2 minimum)`);
        isUnderstaffed = true;
      }
    }

    if (c.N < 2) {
      dayWarnings.push(`Night understaffed (${c.N}/2 minimum)`);
      isUnderstaffed = true;
    } else if (c.N === 3) {
      if (tierConfig.allowNightOver2) {
        options?.relaxations?.push({
          tier: 1,
          rule: 'NIGHT_OVER_2',
          isoDate,
          note: `Night overstaffed on ${isoDate}: N=3 (Tier 0 requires 2)`,
        });
      } else {
        dayWarnings.push(`Night overstaffed (${c.N}/2 maximum)`);
        isOverstaffedNight = true;
      }
    } else if (c.N > 3) {
      dayWarnings.push(`Night overstaffed (${c.N}/2 maximum)`);
      isOverstaffedNight = true;
    }

    const coverageFeasible = isCoverageFeasibleForTier(rotC.M, c.A, c.N, tier);
    if (isUnderstaffed) understaffedDaysCount++;
    if (isOverstaffedNight) nightOverstaffedDaysCount++;
    if (dayWarnings.length > 0) {
      warnings.push(...dayWarnings.map((w) => `${isoDate}: ${w}`));
    }

    dailyCoverage[isoDate] = {
      isoDate,
      dayOfWeek: dow,
      isWeekend,
      M: c.M,
      rotM: rotC.M,
      A: c.A,
      N: c.N,
      MID: c.MID,
      OFF: c.OFF + c.H8 + c.ABS + c.other,
      ABS: c.ABS,
      totalWorking: c.M + c.A + c.N + c.MID,
      isUnderstaffed: !coverageFeasible,
      warnings: dayWarnings,
    };
  });

  // 2. Check Employee Legal Constraints, Streaks, Rest Durations & Tier-Aware Blocks
  employees.forEach((emp) => {
    const isExcluded = excludedSet.has(emp);
    const isFixedMid = FIXED_MID_EMPLOYEES.includes(emp);
    const isFixedWeekdayMA = FIXED_WEEKDAY_MA_EMPLOYEES.includes(emp);
    const isNonRotating = isFixedMid || isFixedWeekdayMA;
    const empViolations: string[] = [];
    const counts = { M: 0, A: 0, N: 0, MID: 0, OFF: 0, H8: 0, ABS: 0 };

    const monthShifts = monthIsos.map((isoDate) => normalizeShiftCode(rows[emp]?.[isoDate]));
    let fixedLeaveOnWorkdays = 0;

    monthShifts.forEach((s, idx) => {
      if (s === 'M') counts.M++;
      else if (s === 'A') counts.A++;
      else if (s === 'N') counts.N++;
      else if (s === 'MID') counts.MID++;
      else if (s === 'H8') counts.H8++;
      else if (s === 'ABS') counts.ABS++;
      else counts.OFF++;

      const meta = calMeta[idx];

      // H6: Excluded employees must be strictly OFF
      if (isExcluded) {
        if (s !== 'OFF') {
          empViolations.push(
            `Excluded employee ${emp} has non-OFF shift ${s} on ${meta.isoDate}`
          );
          fixedStaffViolationCount++;
        }
        return;
      }

      // H6: Valid shift codes
      if (!VALID_CODES.has(s)) {
        empViolations.push(
          `${emp} has invalid/unexpected shift code "${s}" on ${meta.isoDate}`
        );
        fixedStaffViolationCount++;
      }

      // H6: Fixed-staff rules & rotating MID ban
      if (isFixedMid) {
        if (s === 'ABS' || s === 'H8') {
          if (!meta.isFixedOff) fixedLeaveOnWorkdays++;
          return;
        }
        if (!VALID_CODES.has(s)) return;
        const expected = meta.isFixedOff ? 'OFF' : 'MID';
        if (s !== expected) {
          if (meta.isFixedOff) {
            empViolations.push(
              `Fixed Manager ${emp} scheduled on ${s} instead of OFF on weekend/holiday (${meta.isoDate})`
            );
          } else {
            empViolations.push(
              `Fixed Manager ${emp} scheduled on ${s} instead of MID on ${meta.isoDate}`
            );
          }
          fixedStaffViolationCount++;
        }
      } else if (isFixedWeekdayMA) {
        if (s === 'ABS' || s === 'H8') {
          if (!meta.isFixedOff) fixedLeaveOnWorkdays++;
          return;
        }
        if (!VALID_CODES.has(s)) return;
        const expected = meta.isFixedOff ? 'OFF' : meta.isoWeek % 2 === 1 ? 'M' : 'A';
        if (s !== expected) {
          if (meta.isFixedOff) {
            empViolations.push(
              `Weekday staff ${emp} scheduled on ${s} instead of OFF on weekend/holiday (${meta.isoDate})`
            );
          } else {
            empViolations.push(
              `Weekday staff ${emp} scheduled on ${s} instead of ${expected} on ${meta.isoDate} (ISO wk ${meta.isoWeek})`
            );
          }
          fixedStaffViolationCount++;
        }
      } else if (s === 'MID') {
        empViolations.push(`Rotating employee ${emp} assigned MID on ${meta.isoDate}`);
        fixedStaffViolationCount++;
      }
    });

    const totalWorkingDays = counts.M + counts.A + counts.N + counts.MID;

    if (isExcluded) {
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
        targetWorkingDays: 0,
        workDaysDiff: 0,
        maxStreak: 0,
        violations: empViolations,
      };
      if (empViolations.length > 0) {
        warnings.push(...empViolations);
      }
      return;
    }

    // Build combined sequence (historyLast6 + month) for both rotating & fixed active staff (H7)
    const empHistory = histTailIsos.map((iso) =>
      normalizeShiftCode(priorRoster?.rows?.[emp]?.[iso])
    );
    const seq = [...empHistory, ...monthShifts];
    const totalLen = seq.length;
    const isoAt = (idx: number): string =>
      idx >= histLen ? monthIsos[idx - histLen] : `history[${idx - histLen}]`;

    // H1: Max 5 consecutive work days (A/M/N/MID count as work; OFF/ABS/H8/HUP/AD/CM break streak)
    let maxStreak = 0;
    {
      let runStart = -1;
      for (let i = 0; i <= totalLen; i++) {
        const w = i < totalLen && WORK_SHIFTS.has(seq[i]);
        if (w && runStart === -1) runStart = i;
        if (!w && runStart !== -1) {
          const runEnd = i - 1;
          const runLen = runEnd - runStart + 1;
          if (runEnd >= histLen) {
            maxStreak = Math.max(maxStreak, runLen);
            if (runLen > 5) {
              empViolations.push(
                `${emp} has ${runLen} consecutive work days (${isoAt(runStart)}..${isoAt(runEnd)}, max 5)`
              );
              streakViolationCount++;
            }
          }
          runStart = -1;
        }
      }
    }

    // H2 & H3: Post-Night rest >= 2 non-work days and no N directly followed by M/A/MID
    for (let i = 0; i < totalLen - 1; i++) {
      if (seq[i] !== 'N') continue;
      const next1 = seq[i + 1];
      if (next1 === 'N') continue;

      if ((next1 === 'M' || next1 === 'A' || next1 === 'MID') && i + 1 >= histLen) {
        empViolations.push(
          `Night shift on ${isoAt(i)} followed directly by ${next1} on ${isoAt(i + 1)} without rest`
        );
        nightRestViolationCount++;
      }

      const prematureIndices: number[] = [];
      for (const k of [1, 2]) {
        if (i + k < totalLen && WORK_SHIFTS.has(seq[i + k]) && i + k >= histLen) {
          prematureIndices.push(i + k);
        }
      }
      if (prematureIndices.length > 0) {
        const firstPremature = prematureIndices[0];
        empViolations.push(
          `Night shift on ${isoAt(i)} followed by < 2 rest days before ${seq[firstPremature]} on ${isoAt(firstPremature)} (minimum 2 rest days required)`
        );
        nightRestViolationCount++;
      }
    }

    // Extract maximal contiguous work blocks across seq
    interface WorkBlock {
      start: number;
      end: number;
      len: number;
      shifts: string[];
      primaryShift: string;
      isMixed: boolean;
    }
    const workBlocks: WorkBlock[] = [];
    {
      let i = 0;
      while (i < totalLen) {
        if (WORK_SHIFTS.has(seq[i])) {
          const start = i;
          const shiftSet = new Set<string>();
          while (i < totalLen && WORK_SHIFTS.has(seq[i])) {
            shiftSet.add(seq[i]);
            i++;
          }
          const end = i - 1;
          workBlocks.push({
            start,
            end,
            len: end - start + 1,
            shifts: Array.from(shiftSet),
            primaryShift: seq[start],
            isMixed: shiftSet.size > 1,
          });
        } else {
          i++;
        }
      }
    }

    // H4 part 1: No mixed shift types inside one block (for all active staff)
    for (const blk of workBlocks) {
      if (blk.end < histLen) continue;
      if (blk.isMixed) {
        empViolations.push(
          `Mixed shift block on ${isoAt(Math.max(histLen, blk.start))}: cannot combine [${blk.shifts.join(',')}] in single work block ${isoAt(blk.start)}..${isoAt(blk.end)}`
        );
        mixedBlockViolationCount++;
      }
    }

    // Rotating-only checks: H4 cycle order, H8 rest <= 3d, Tier 0/2 block lengths & rests
    if (!isNonRotating) {
      // H4 part 2: Forward cycle A -> M -> N -> A (with ABS/H8/CM split rules)
      for (let b = 1; b < workBlocks.length; b++) {
        const prev = workBlocks[b - 1];
        const cur = workBlocks[b];
        if (cur.start < histLen) continue;
        if (prev.isMixed || cur.isMixed) continue;
        if (!ROTATING_SHIFTS.has(prev.primaryShift) || !ROTATING_SHIFTS.has(cur.primaryShift)) {
          continue;
        }

        const gap = seq.slice(prev.end + 1, cur.start);
        const absH8Count = gap.filter((s) => LEAVE_OR_SPLIT_CODES.has(s)).length;
        const expectedNext = EXPECTED_NEXT_SHIFT[prev.primaryShift];

        if (absH8Count === 0) {
          if (cur.primaryShift !== expectedNext) {
            empViolations.push(
              `Shift sequence violation on ${isoAt(cur.start)}: expected ${expectedNext} after ${prev.primaryShift}, got ${cur.primaryShift}`
            );
            mixedBlockViolationCount++;
          }
        } else {
          const isSame = cur.primaryShift === prev.primaryShift;
          const isNext = cur.primaryShift === expectedNext;
          if (!isSame && !isNext && absH8Count < 2) {
            empViolations.push(
              `Shift sequence violation on ${isoAt(cur.start)} across single-day leave split: ${prev.primaryShift} --[${gap.join(',')}]--> ${cur.primaryShift}`
            );
            mixedBlockViolationCount++;
          }
        }
      }

      // H8: Pure-OFF rest > 3 consecutive days (rest containing ABS/H8/CM exempt)
      {
        let rStart = -1;
        for (let i = 0; i <= totalLen; i++) {
          const nonWork = i < totalLen && !WORK_SHIFTS.has(seq[i]);
          if (nonWork && rStart === -1) rStart = i;
          if (!nonWork && rStart !== -1) {
            const rEnd = i - 1;
            const rLen = rEnd - rStart + 1;
            if (rEnd >= histLen && rLen > 3) {
              const hasLeave = seq
                .slice(rStart, rEnd + 1)
                .some((s) => LEAVE_OR_SPLIT_CODES.has(s));
              if (!hasLeave) {
                empViolations.push(
                  `Rest period exceeded 3 days on ${isoAt(Math.max(histLen, rStart))} (${rLen} days: ${isoAt(rStart)}..${isoAt(rEnd)})`
                );
                excessiveRestViolationCount++;
              }
            }
            rStart = -1;
          }
        }
      }

      // Tier 0 & Tier 2: Rotating work block lengths (3..5; exempt tail, !hasHistory Day-1, adjacent to ABS/H8/CM)
      let twoDayBlocksCount = 0;
      for (const blk of workBlocks) {
        if (blk.end < histLen) continue;
        const isTrailingAtMonthEnd = blk.end === totalLen - 1;
        const isDay1NoHistory = !hasHistory && blk.start === histLen;
        const adjBefore = blk.start > 0 && LEAVE_OR_SPLIT_CODES.has(seq[blk.start - 1]);
        const adjAfter = blk.end + 1 < totalLen && LEAVE_OR_SPLIT_CODES.has(seq[blk.end + 1]);
        const isExempt = isTrailingAtMonthEnd || isDay1NoHistory || adjBefore || adjAfter;

        if (!isExempt && blk.len < 3) {
          const blockIso = isoAt(Math.max(histLen, blk.start));
          if (blk.len === 1) {
            empViolations.push(
              `Short work block (${blk.primaryShift}, 1 day) starting ${blockIso}: 1-day blocks are not allowed`
            );
            streakViolationCount++;
          } else if (blk.len === 2) {
            twoDayBlocksCount++;
            if (twoDayBlocksCount <= 1) {
              const note = `${emp} has 2-day work block (${blk.primaryShift}) at ${isoAt(blk.start)}..${isoAt(blk.end)}`;
              if (tierConfig.allowShortBlock2) {
                options?.relaxations?.push({
                  tier: 2,
                  rule: 'SHORT_BLOCK_2',
                  isoDate: blockIso,
                  employee: emp,
                  note,
                });
              } else {
                empViolations.push(
                  `Short work block (${blk.primaryShift}, 2 days) starting ${blockIso}: internal blocks require 3-5 consecutive days`
                );
                streakViolationCount++;
              }
            } else {
              empViolations.push(
                `${emp} has ${twoDayBlocksCount} two-day work blocks in month (max 1 in Tier 2) at ${isoAt(blk.start)}..${isoAt(blk.end)}`
              );
              streakViolationCount++;
            }
          }
        }
      }

      // Tier 0 & Tier 2: 1-day pure-OFF rest between non-Night blocks (SHORT_REST_1)
      for (let b = 1; b < workBlocks.length; b++) {
        const prev = workBlocks[b - 1];
        const cur = workBlocks[b];
        if (cur.start < histLen) continue;
        const gap = seq.slice(prev.end + 1, cur.start);
        if (gap.length === 1 && !LEAVE_OR_SPLIT_CODES.has(gap[0]) && prev.primaryShift !== 'N') {
          const restIso = isoAt(Math.max(histLen, prev.end + 1));
          const note = `${emp} has 1-day OFF rest on ${isoAt(prev.end + 1)} between ${prev.primaryShift} and ${cur.primaryShift}`;
          if (tierConfig.minRest <= 1) {
            options?.relaxations?.push({
              tier: 2,
              rule: 'SHORT_REST_1',
              isoDate: restIso,
              employee: emp,
              note,
            });
          } else {
            empViolations.push(note);
            nightRestViolationCount++;
          }
        }
      }
    }

    // H5: Exact monthly norm
    let effectiveTarget: number;
    if (isNonRotating) {
      effectiveTarget = Math.max(0, fixedMonthWorkdays - fixedLeaveOnWorkdays);
    } else {
      const baseTarget =
        targetWorkingDays ?? countLegalWorkingDaysInMonth(targetMonth ?? 11, targetYear ?? 2026);
      effectiveTarget = Math.max(0, baseTarget - counts.ABS - counts.H8);
    }
    const workDaysDiff = totalWorkingDays - effectiveTarget;

    if (isNonRotating) {
      if (workDaysDiff !== 0) {
        empViolations.push(
          `Fixed employee ${emp}: worked ${totalWorkingDays} days (target is ${effectiveTarget} weekdays${fixedLeaveOnWorkdays > 0 ? ` after ${fixedLeaveOnWorkdays}d ABS/H8` : ''})`
        );
        workDaysMismatchCount++;
      }
    } else if (targetWorkingDays !== undefined && workDaysDiff !== 0) {
      const leaveDays = counts.ABS + counts.H8;
      empViolations.push(
        `${emp}: worked ${totalWorkingDays} days (target is ${effectiveTarget} days${leaveDays > 0 ? ` after ${counts.ABS}d ABS / ${counts.H8}d H8` : ''})`
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
      targetWorkingDays: effectiveTarget,
      workDaysDiff,
      maxStreak,
      violations: empViolations,
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
    warnings,
  };
}
