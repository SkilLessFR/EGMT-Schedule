import type { RosterData } from './types';
import {
  FIXED_MID_EMPLOYEES,
  FIXED_WEEKDAY_MA_EMPLOYEES,
  FIXED_STAFF_EXEMPT_OFF_DATES,
  ROMANIAN_HOLIDAYS,
  shiftKey
} from './constants';
import { getIsoWeekNumber } from './dateUtils';
import {
  hasMixedBlocks,
  getActiveTailBlock,
  getActiveTailRest,
  getEmployeeHistoryTail
} from './historyUtils';

/**
 * Equalizes every employee so they work EXACTLY targetWorkingDays (e.g. 21 in Dec, 22 in Oct)
 */
export function equalizeExactWorkingDays(
  rows: Record<string, Record<string, string>>,
  dateColumns: { index: number; date: Date; isoDate: string }[],
  targetWorkingDays: number,
  priorRoster: RosterData | null,
  year: number,
  month: number,
  excludedEmployees: string[] = []
) {
  const excludedSet = new Set(excludedEmployees);
  const allEmployees = Object.keys(rows);
  const fixedMid = FIXED_MID_EMPLOYEES.filter((e) => allEmployees.includes(e) && !excludedSet.has(e));
  const fixedWeekdayMA = FIXED_WEEKDAY_MA_EMPLOYEES.filter((e) => allEmployees.includes(e) && !excludedSet.has(e));
  const rotating = allEmployees.filter((e) => !fixedMid.includes(e) && !fixedWeekdayMA.includes(e) && !excludedSet.has(e));

  // 1. Fixed Managers (Palade): Monday to Friday MID, Saturday and Sunday OFF, Romanian National Holidays OFF
  fixedMid.forEach((emp) => {
    dateColumns.forEach(({ date, isoDate }) => {
      const dow = date.getDay();
      const isWeekend = dow === 0 || dow === 6;
      const isOff = isWeekend || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
      if (isOff) {
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
      const isOff = isWeekend || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
      if (isOff) {
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
      ? getEmployeeHistoryTail(priorRoster, emp, year, month, 6)
      : [];
  });

  const getLastHistoryWorkShift = (emp: string): string | null => {
    const empHistory = employeeHistories[emp] ?? [];
    for (let i = empHistory.length - 1; i >= 0; i--) {
      const s = empHistory[i];
      if (s && s !== 'OFF' && s !== 'H8' && s !== 'HUP' && s !== 'ABS' && s !== 'AD') {
        return s;
      }
    }
    return null;
  };

  const getBlocks = (emp: string) => {
    const empHistory = employeeHistories[emp] ?? [];
    const empMonthShifts = dateColumns.map((col) => shiftKey(rows[emp]?.[col.isoDate] ?? 'OFF'));
    const activeTailWork = getActiveTailBlock(empHistory);
    const activeTailRest = activeTailWork.length === 0 ? getActiveTailRest(empHistory) : [];
    const activeTail = activeTailWork.length > 0 ? activeTailWork : activeTailRest;
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
    return { empHistory, empMonthShifts, activeTailWork, activeTailRest, full, blocks };
  };

  const countShortRests = (emp: string): number => {
    const { blocks } = getBlocks(emp);
    let shortRests = 0;
    for (let i = 1; i < blocks.length - 1; i++) {
      if (blocks[i].type === 'REST' && blocks[i].len < 2) shortRests++;
    }
    return shortRests;
  };

  const isEmpValid = (emp: string) => {
    const { empHistory, empMonthShifts, activeTailWork, activeTailRest, full, blocks } = getBlocks(emp);
    const hasHistory = empHistory.length > 0;
    const lastHistoryWorkShift = getLastHistoryWorkShift(emp);

    // 1. Mixed blocks check
    if (hasMixedBlocks([...activeTailWork, ...empMonthShifts])) return false;

    // 0. Continuous work streak across all working shifts: NEVER > 5
    let curWorkStreak = 0;
    for (const s of full) {
      if (s !== 'OFF' && s !== 'H8' && s !== 'HUP' && s !== 'ABS' && s !== 'AD') {
        curWorkStreak++;
        if (curWorkStreak > 5) return false;
      } else {
        curWorkStreak = 0;
      }
    }

    // 1. Two WORK blocks cannot be directly adjacent without rest!
    for (let i = 0; i < blocks.length - 1; i++) {
      if (blocks[i].type === 'WORK' && blocks[i + 1].type === 'WORK') {
        return false;
      }
    }

    // 2. Shift sequence check: Work blocks must follow sequence A -> M -> N -> A!
    const EXPECTED_NEXT_SHIFT: Record<string, string> = {
      A: 'M',
      M: 'N',
      N: 'A'
    };
    const workBlocks = blocks.filter((b) => b.type === 'WORK');
    if (activeTailWork.length === 0 && lastHistoryWorkShift && workBlocks.length > 0) {
      const expectedFirst = EXPECTED_NEXT_SHIFT[lastHistoryWorkShift];
      if (expectedFirst && workBlocks[0].shift !== expectedFirst) {
        return false;
      }
    }
    for (let i = 0; i < workBlocks.length - 1; i++) {
      const cur = workBlocks[i].shift;
      const nxt = workBlocks[i + 1].shift;
      if (cur === nxt) return false; // Anti-repeat check
      if (EXPECTED_NEXT_SHIFT[cur] && nxt !== EXPECTED_NEXT_SHIFT[cur]) {
        return false; // Sequence violation
      }
    }

    // 4. Work and rest streaks
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.type === 'WORK') {
        if (b.len > 5) return false;
        // Internal work blocks must be between 3 and 5 days.
        // Boundary blocks: last block at month end can be 1..5 (continues into next month),
        // and first block on Day 1 without prior history can be 1..5 (truncated from prior month).
        const isUnboundedStart = i === 0 && !hasHistory;
        if (!isUnboundedStart && i < blocks.length - 1 && b.len < 3) return false;
      } else {
        // Rest period between work blocks (and trailing month-end rest):
        // After Night shift before returning to work: strictly >= 2 days (48h mandatory rest)
        // Max rest: at most 3 days anywhere in month (including trailing OFF days)
        if (i > 0) {
          if (i < blocks.length - 1) {
            const prevWork = blocks[i - 1];
            if (prevWork && prevWork.shift === 'N' && b.len < 2) return false;
          }
          if (b.len > 3) return false;
        }
        // Rest bridging across month boundary and leading rest in standalone months
        if (i === 0 && activeTailRest.length === 0 && b.len > 3) {
          return false;
        }
        if (i === 0 && activeTailRest.length > 0 && i < blocks.length - 1) {
          if (lastHistoryWorkShift === 'N' && b.len < 2) return false;
          if (b.len > activeTailRest.length && b.len > 3) return false;
        }
      }
    }

    // 5. Night turnaround check: night shift followed by at least 2 days of rest before M, A, or MID
    const checkSeq = [...empHistory.slice(-2), ...empMonthShifts];
    const histOffset = Math.min(2, empHistory.length);
    for (let i = 0; i < checkSeq.length; i++) {
      if (checkSeq[i] === 'N') {
        if (i + 1 >= histOffset && i + 1 < checkSeq.length) {
          const next = checkSeq[i + 1];
          if (next === 'M' || next === 'A' || next === 'MID') return false;
        }
        if (i + 2 >= histOffset && i + 2 < checkSeq.length) {
          const next = checkSeq[i + 1];
          const next2 = checkSeq[i + 2];
          if (next === 'OFF' && (next2 === 'M' || next2 === 'A' || next2 === 'MID')) return false;
        }
      }
    }

    return true;
  };

  const getDailyCoverage = () => {
    const dailyCov: Record<string, { rotM: number; totalA: number; N: number }> = {};
    dateColumns.forEach(({ isoDate }) => {
      let rotM = 0;
      let totalA = 0;
      let N = 0;
      rotating.forEach((emp) => {
        const s = rows[emp][isoDate];
        if (s === 'M') rotM++;
        else if (s === 'A') totalA++;
        else if (s === 'N') N++;
      });
      fixedWeekdayMA.forEach((emp) => {
        if (rows[emp]?.[isoDate] === 'A') totalA++;
      });
      dailyCov[isoDate] = { rotM, totalA, N };
    });
    return dailyCov;
  };

  const getDayCounts = (isoDate: string) => {
    let rotM = 0;
    let totalA = 0;
    let N = 0;
    rotating.forEach((emp) => {
      const s = rows[emp][isoDate];
      if (s === 'M') rotM++;
      else if (s === 'A') totalA++;
      else if (s === 'N') N++;
    });
    fixedWeekdayMA.forEach((emp) => {
      if (rows[emp]?.[isoDate] === 'A') totalA++;
    });
    return { rotM, totalA, N };
  };

  const canDropShiftOnDay = (isoDate: string, shift: string): boolean => {
    const c = getDayCounts(isoDate);
    if (shift === 'M') return c.rotM > 1;
    if (shift === 'A') return c.totalA > 2;
    if (shift === 'N') return c.N > 2;
    return false;
  };

  const canAddShiftOnDay = (isoDate: string, shift: string): boolean => {
    if (shift !== 'N') return true;
    return getDayCounts(isoDate).N < 2;
  };

  const getPrevNeighbor = (emp: string, i: number) =>
    i > 0
      ? shiftKey(rows[emp][dateColumns[i - 1].isoDate] ?? 'OFF')
      : (employeeHistories[emp]?.[employeeHistories[emp].length - 1] ?? 'OFF');

  const getNextNeighbor = (emp: string, i: number) =>
    i < dateColumns.length - 1
      ? shiftKey(rows[emp][dateColumns[i + 1].isoDate] ?? 'OFF')
      : 'OFF';

  const getValidAddCandidates = (emp: string, i: number): ('M' | 'A' | 'N')[] => {
    const prev = getPrevNeighbor(emp, i);
    const next = getNextNeighbor(emp, i);
    const cands: ('M' | 'A' | 'N')[] = [];
    if (prev === 'M' || prev === 'A' || prev === 'N') cands.push(prev);
    if ((next === 'M' || next === 'A' || next === 'N') && !cands.includes(next)) cands.push(next);
    if (cands.length === 0) cands.push('M', 'A', 'N');
    return cands.filter((cand) =>
      !(prev !== 'OFF' && prev !== cand) && !(next !== 'OFF' && next !== cand)
    );
  };

  // Dropping strictly inside a block would split one shift block into two identical blocks separated by OFF
  const isBoundaryDrop = (emp: string, i: number, shift: string): boolean =>
    !(getPrevNeighbor(emp, i) === shift && getNextNeighbor(emp, i) === shift);

  // Single added shift must touch an existing block of `shift` or sit at an unbounded month edge
  const canAttachStandaloneAdd = (emp: string, i: number, shift: string): boolean => {
    const prev = getPrevNeighbor(emp, i);
    const next = getNextNeighbor(emp, i);
    if ((prev !== 'OFF' && prev !== shift) || (next !== 'OFF' && next !== shift)) return false;
    if (prev === shift || next === shift) return true;
    const isUnboundedStart = i === 0 && (employeeHistories[emp]?.length ?? 0) === 0;
    return isUnboundedStart || i === dateColumns.length - 1;
  };

  // In a pair (a1 < a2), an isolated a1 can only become valid if a2 = a1 + 1 bridges to an edge or existing block
  const canPlaceFirstOfPair = (emp: string, a1: number, s1: string): boolean => {
    if (canAttachStandaloneAdd(emp, a1, s1)) return true;
    const a2 = a1 + 1;
    if (a2 >= dateColumns.length || rows[emp][dateColumns[a2].isoDate] !== 'OFF') return false;
    const next2 = getNextNeighbor(emp, a2);
    return a2 === dateColumns.length - 1 || next2 === s1;
  };

  // Phase 0: Resolve boundary consecutive work streaks > 5 days by handing off excess shifts
  for (const emp of rotating) {
    const hist = employeeHistories[emp] ?? [];
    let tailStreak = 0;
    for (let i = hist.length - 1; i >= 0; i--) {
      if (hist[i] !== 'OFF' && hist[i] !== 'H8' && hist[i] !== 'HUP' && hist[i] !== 'ABS' && hist[i] !== 'AD') tailStreak++;
      else break;
    }
    if (tailStreak === 0) continue;

    let headStreak = 0;
    for (let i = 0; i < Math.min(5, dateColumns.length); i++) {
      const s = rows[emp][dateColumns[i].isoDate];
      if (s !== 'OFF' && s !== 'H8' && s !== 'HUP' && s !== 'ABS' && s !== 'AD') headStreak++;
      else break;
    }

    if (tailStreak + headStreak > 5) {
      const cutoff = Math.max(0, 5 - tailStreak);
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

  // Phase 0b: Repair any rotating employee still failing isEmpValid (e.g. boundary <3d block or <2d N rest)
  for (const emp of rotating) {
    for (let iter = 0; iter < 5 && !isEmpValid(emp); iter++) {
      let fixed = false;
      for (const strictCov of [true, false]) {
        for (let d = 0; d < dateColumns.length; d++) {
          const dropIso = dateColumns[d].isoDate;
          const dropShift = rows[emp][dropIso];
          if (dropShift !== 'M' && dropShift !== 'A' && dropShift !== 'N') continue;
          if (strictCov && !canDropShiftOnDay(dropIso, dropShift)) continue;

          rows[emp][dropIso] = 'OFF';
          for (let a = 0; a < dateColumns.length; a++) {
            if (a === d) continue;
            const addIso = dateColumns[a].isoDate;
            if (rows[emp][addIso] !== 'OFF') continue;
            for (const cand of getValidAddCandidates(emp, a)) {
              if (!canAddShiftOnDay(addIso, cand)) continue;
              rows[emp][addIso] = cand;
              if (isEmpValid(emp)) {
                fixed = true;
                break;
              }
              rows[emp][addIso] = 'OFF';
            }
            if (fixed) break;
          }
          if (fixed) break;
          rows[emp][dropIso] = dropShift;
        }
        if (fixed) break;
      }
      if (!fixed) break;
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

  const runPhaseB = () => {
    // Phase B: If any deficit remains, add shifts (preferring coverage deficits and preserving >=2d rests)
    for (let iter = 0; iter < 25; iter++) {
      const deficit = rotating.filter((e) => getWorkDays(e) < targetWorkingDays);
      if (deficit.length === 0) break;

      let added = false;
      for (const dEmp of deficit) {
        let bestMove: { isoDate: string; cand: 'M' | 'A' | 'N'; score: number } | null = null;
        const prevShortRests = countShortRests(dEmp);

        const counts: Record<string, number> = { M: 0, A: 0, N: 0 };
        dateColumns.forEach(({ isoDate: dIso }) => {
          const s = rows[dEmp][dIso];
          if (s === 'M' || s === 'A' || s === 'N') counts[s]++;
        });

        for (let i = 0; i < dateColumns.length; i++) {
          const isoDate = dateColumns[i].isoDate;
          if (rows[dEmp][isoDate] !== 'OFF') continue;

          const prevShift = getPrevNeighbor(dEmp, i);
          const nextShift = getNextNeighbor(dEmp, i);

          const candidates: ('M' | 'A' | 'N')[] = [];
          if (prevShift === 'M' || prevShift === 'A' || prevShift === 'N') candidates.push(prevShift);
          if ((nextShift === 'M' || nextShift === 'A' || nextShift === 'N') && !candidates.includes(nextShift)) {
            candidates.push(nextShift);
          }
          if (candidates.length === 0) {
            candidates.push('M', 'A');
            if (getDayCounts(isoDate).N < 2) candidates.push('N');
          }

          const dayCov = getDayCounts(isoDate);
          for (const cand of candidates) {
            if ((prevShift !== 'OFF' && prevShift !== cand) || (nextShift !== 'OFF' && nextShift !== cand)) {
              continue;
            }
            if (cand === 'N' && dayCov.N >= 2) continue;

            rows[dEmp][isoDate] = cand;
            if (isEmpValid(dEmp)) {
              const newCounts = { ...counts, [cand]: (counts[cand] ?? 0) + 1 };
              const spread = Math.max(newCounts.M, newCounts.A) - Math.min(newCounts.M, newCounts.A);
              const fixesDeficit =
                (cand === 'N' && dayCov.N < 2) ||
                (cand === 'A' && dayCov.totalA < 2) ||
                (cand === 'M' && dayCov.rotM < 1);
              const newShortRests = countShortRests(dEmp);
              const shortRestPenalty = Math.max(0, newShortRests - prevShortRests) * 100;
              const score = (fixesDeficit ? -1000 : 0) + shortRestPenalty + spread;
              if (!bestMove || score < bestMove.score) {
                bestMove = { isoDate, cand, score };
              }
            }
            rows[dEmp][isoDate] = 'OFF';
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
  };

  const runPhaseB2 = () => {
    // Phase B2: Multi-step & 2-employee repair for remaining deficit employees (net +1 work day)
    for (let iter = 0; iter < 20; iter++) {
      const deficit = rotating.filter((e) => getWorkDays(e) < targetWorkingDays);
      if (deficit.length === 0) break;

      let repaired = false;
      for (const dEmp of deficit) {
        // B2a: 1-employee (1 drop + 2 adds = net +1) sliding block boundaries while preserving coverage
        const tryOneEmpSlideAdd = (): boolean => {
          let fallbackMove: { dropIso: string; dropShift: string; iso1: string; s1: string; iso2: string; s2: string } | null = null;

          for (let d = 0; d < dateColumns.length; d++) {
            const dropIso = dateColumns[d].isoDate;
            const dropShift = rows[dEmp][dropIso];
            if (dropShift !== 'M' && dropShift !== 'A' && dropShift !== 'N') continue;
            if (!isBoundaryDrop(dEmp, d, dropShift) || !canDropShiftOnDay(dropIso, dropShift)) continue;

            rows[dEmp][dropIso] = 'OFF';
            for (let a1 = 0; a1 < dateColumns.length; a1++) {
              if (a1 === d) continue;
              const iso1 = dateColumns[a1].isoDate;
              if (rows[dEmp][iso1] !== 'OFF') continue;

              const cands1 = getValidAddCandidates(dEmp, a1);
              for (const s1 of cands1) {
                if (!canAddShiftOnDay(iso1, s1) || !canPlaceFirstOfPair(dEmp, a1, s1)) continue;
                rows[dEmp][iso1] = s1;

                for (let a2 = a1 + 1; a2 < dateColumns.length; a2++) {
                  if (a2 === d) continue;
                  const iso2 = dateColumns[a2].isoDate;
                  if (rows[dEmp][iso2] !== 'OFF') continue;

                  const cands2 = getValidAddCandidates(dEmp, a2);
                  for (const s2 of cands2) {
                    if (!canAddShiftOnDay(iso2, s2) || !canAttachStandaloneAdd(dEmp, a2, s2)) continue;
                    rows[dEmp][iso2] = s2;

                    if (isEmpValid(dEmp)) {
                      if (countShortRests(dEmp) === 0) {
                        return true;
                      }
                      if (!fallbackMove) {
                        fallbackMove = { dropIso, dropShift, iso1, s1, iso2, s2 };
                      }
                    }
                    rows[dEmp][iso2] = 'OFF';
                  }
                }
                rows[dEmp][iso1] = 'OFF';
              }
            }
            rows[dEmp][dropIso] = dropShift;
          }

          if (fallbackMove) {
            rows[dEmp][fallbackMove.dropIso] = 'OFF';
            rows[dEmp][fallbackMove.iso1] = fallbackMove.s1;
            rows[dEmp][fallbackMove.iso2] = fallbackMove.s2;
            return true;
          }
          return false;
        };

        if (tryOneEmpSlideAdd()) {
          repaired = true;
          break;
        }

        // B2b-1: 2-employee transfer + recovery (helper transfers iso1 to dEmp, helper adds iso2)
        for (const helper of rotating) {
          if (helper === dEmp) continue;
          for (let i = 0; i < dateColumns.length; i++) {
            const iso1 = dateColumns[i].isoDate;
            const hShift = rows[helper][iso1];
            if ((hShift !== 'M' && hShift !== 'A' && hShift !== 'N') || rows[dEmp][iso1] !== 'OFF') {
              continue;
            }
            const prevD = getPrevNeighbor(dEmp, i);
            const nextD = getNextNeighbor(dEmp, i);
            if ((prevD !== 'OFF' && prevD !== hShift) || (nextD !== 'OFF' && nextD !== hShift)) {
              continue;
            }
            if (!isBoundaryDrop(helper, i, hShift)) continue;

            rows[helper][iso1] = 'OFF';
            rows[dEmp][iso1] = hShift;
            if (isEmpValid(dEmp)) {
              for (let j = 0; j < dateColumns.length; j++) {
                if (j === i) continue;
                const iso2 = dateColumns[j].isoDate;
                if (rows[helper][iso2] !== 'OFF') continue;
                for (const s2 of getValidAddCandidates(helper, j)) {
                  if (!canAddShiftOnDay(iso2, s2) || !canAttachStandaloneAdd(helper, j, s2)) continue;
                  rows[helper][iso2] = s2;
                  if (isEmpValid(helper)) {
                    repaired = true;
                    break;
                  }
                  rows[helper][iso2] = 'OFF';
                }
                if (repaired) break;
              }
            }
            if (repaired) break;
            rows[helper][iso1] = hShift;
            rows[dEmp][iso1] = 'OFF';
          }
          if (repaired) break;
        }
        if (repaired) break;

        // B2b-2: dEmp drops iso1 (at min cov) + adds iso2 & iso3, helper covers iso1 + drops surplus iso4
        for (let d = 0; d < dateColumns.length; d++) {
          const iso1 = dateColumns[d].isoDate;
          const dropShift = rows[dEmp][iso1];
          if (dropShift !== 'M' && dropShift !== 'A' && dropShift !== 'N') continue;
          if (!isBoundaryDrop(dEmp, d, dropShift)) continue;
          const validHelpers = rotating.filter(
            (helper) => helper !== dEmp && rows[helper][iso1] === 'OFF' && canAttachStandaloneAdd(helper, d, dropShift)
          );
          if (validHelpers.length === 0) continue;

          rows[dEmp][iso1] = 'OFF';
          for (let a1 = 0; a1 < dateColumns.length; a1++) {
            if (a1 === d) continue;
            const iso2 = dateColumns[a1].isoDate;
            if (rows[dEmp][iso2] !== 'OFF') continue;
            for (const s1 of getValidAddCandidates(dEmp, a1)) {
              if (!canAddShiftOnDay(iso2, s1) || !canPlaceFirstOfPair(dEmp, a1, s1)) continue;
              rows[dEmp][iso2] = s1;
              for (let a2 = a1 + 1; a2 < dateColumns.length; a2++) {
                if (a2 === d) continue;
                const iso3 = dateColumns[a2].isoDate;
                if (rows[dEmp][iso3] !== 'OFF') continue;
                for (const s2 of getValidAddCandidates(dEmp, a2)) {
                  if (!canAddShiftOnDay(iso3, s2) || !canAttachStandaloneAdd(dEmp, a2, s2)) continue;
                  rows[dEmp][iso3] = s2;
                  if (isEmpValid(dEmp)) {
                    for (const helper of validHelpers) {
                      rows[helper][iso1] = dropShift;
                      for (let k = 0; k < dateColumns.length; k++) {
                        if (k === d) continue;
                        const iso4 = dateColumns[k].isoDate;
                        const s4 = rows[helper][iso4];
                        if (s4 !== 'M' && s4 !== 'A' && s4 !== 'N') continue;
                        if (!isBoundaryDrop(helper, k, s4) || !canDropShiftOnDay(iso4, s4)) continue;
                        rows[helper][iso4] = 'OFF';
                        if (isEmpValid(helper)) {
                          repaired = true;
                          break;
                        }
                        rows[helper][iso4] = s4;
                      }
                      if (repaired) break;
                      rows[helper][iso1] = 'OFF';
                    }
                  }
                  if (repaired) break;
                  rows[dEmp][iso3] = 'OFF';
                }
                if (repaired) break;
              }
              if (repaired) break;
              rows[dEmp][iso2] = 'OFF';
            }
            if (repaired) break;
          }
          if (repaired) break;
          rows[dEmp][iso1] = dropShift;
        }
        if (repaired) break;
      }
      if (!repaired) break;
    }
  };

  runPhaseB();
  runPhaseB2();

  // Phase C: If any surplus remains, drop shift to OFF on days with adequate coverage (including fixedWeekdayMA for A)
  for (let iter = 0; iter < 25; iter++) {
    const surplus = rotating.filter((e) => getWorkDays(e) > targetWorkingDays);
    if (surplus.length === 0) break;

    let dropped = false;
    for (const sEmp of surplus) {
      let bestDropIso: string | null = null;
      let bestDropScore = Infinity;
      const prevShortRests = countShortRests(sEmp);

      for (let i = 0; i < dateColumns.length; i++) {
        const isoDate = dateColumns[i].isoDate;
        const cur = rows[sEmp][isoDate];
        if (cur !== 'M' && cur !== 'A' && cur !== 'N') continue;

        const c = getDayCounts(isoDate);
        const minAllowed = cur === 'M' ? 1 : 2;
        const count = cur === 'M' ? c.rotM : cur === 'A' ? c.totalA : c.N;
        if (count > minAllowed) {
          rows[sEmp][isoDate] = 'OFF';
          if (isEmpValid(sEmp)) {
            const shortRestDelta = Math.max(0, countShortRests(sEmp) - prevShortRests);
            const surplusMargin = count - minAllowed;
            const score = shortRestDelta * 100 - surplusMargin;
            if (score < bestDropScore) {
              bestDropScore = score;
              bestDropIso = isoDate;
            }
          }
          rows[sEmp][isoDate] = cur;
        }
      }

      if (bestDropIso) {
        rows[sEmp][bestDropIso] = 'OFF';
        dropped = true;
        break;
      }
    }
    if (!dropped) break;
  }

  // Phase C2: Multi-step & 2-employee repair for remaining surplus employees (net -1 work day)
  for (let iter = 0; iter < 20; iter++) {
    const surplus = rotating.filter((e) => getWorkDays(e) > targetWorkingDays);
    if (surplus.length === 0) break;

    let repaired = false;
    for (const sEmp of surplus) {
      // C2a: 1-employee (1 add + 2 drops = net -1)
      for (let a = 0; a < dateColumns.length; a++) {
        const isoA = dateColumns[a].isoDate;
        if (rows[sEmp][isoA] !== 'OFF') continue;
        for (const sA of getValidAddCandidates(sEmp, a)) {
          if (!canAddShiftOnDay(isoA, sA) || !canAttachStandaloneAdd(sEmp, a, sA)) continue;
          rows[sEmp][isoA] = sA;

          for (let d1 = 0; d1 < dateColumns.length; d1++) {
            if (d1 === a) continue;
            const iso1 = dateColumns[d1].isoDate;
            const sh1 = rows[sEmp][iso1];
            if (sh1 !== 'M' && sh1 !== 'A' && sh1 !== 'N') continue;
            if (!isBoundaryDrop(sEmp, d1, sh1) || !canDropShiftOnDay(iso1, sh1)) continue;
            rows[sEmp][iso1] = 'OFF';

            for (let d2 = d1 + 1; d2 < dateColumns.length; d2++) {
              if (d2 === a) continue;
              const iso2 = dateColumns[d2].isoDate;
              const sh2 = rows[sEmp][iso2];
              if (sh2 !== 'M' && sh2 !== 'A' && sh2 !== 'N') continue;
              if (!isBoundaryDrop(sEmp, d2, sh2) || !canDropShiftOnDay(iso2, sh2)) continue;
              rows[sEmp][iso2] = 'OFF';

              if (isEmpValid(sEmp)) {
                repaired = true;
                break;
              }
              rows[sEmp][iso2] = sh2;
            }
            if (repaired) break;
            rows[sEmp][iso1] = sh1;
          }
          if (repaired) break;
          rows[sEmp][isoA] = 'OFF';
        }
        if (repaired) break;
      }
      if (repaired) break;

      // C2b: 2-employee handoff + drop (sEmp hands off iso1 to helper, helper drops surplus iso2)
      for (let i = 0; i < dateColumns.length; i++) {
        const iso1 = dateColumns[i].isoDate;
        const sh1 = rows[sEmp][iso1];
        if (sh1 !== 'M' && sh1 !== 'A' && sh1 !== 'N') continue;
        if (!isBoundaryDrop(sEmp, i, sh1)) continue;

        rows[sEmp][iso1] = 'OFF';
        if (isEmpValid(sEmp)) {
          for (const helper of rotating) {
            if (helper === sEmp || rows[helper][iso1] !== 'OFF') continue;
            if (!canAttachStandaloneAdd(helper, i, sh1)) continue;
            rows[helper][iso1] = sh1;
            for (let j = 0; j < dateColumns.length; j++) {
              if (j === i) continue;
              const iso2 = dateColumns[j].isoDate;
              const sh2 = rows[helper][iso2];
              if (sh2 !== 'M' && sh2 !== 'A' && sh2 !== 'N') continue;
              if (!isBoundaryDrop(helper, j, sh2) || !canDropShiftOnDay(iso2, sh2)) continue;
              rows[helper][iso2] = 'OFF';
              if (isEmpValid(helper)) {
                repaired = true;
                break;
              }
              rows[helper][iso2] = sh2;
            }
            if (repaired) break;
            rows[helper][iso1] = 'OFF';
          }
        }
        if (repaired) break;
        rows[sEmp][iso1] = sh1;
      }
      if (repaired) break;
    }
    if (!repaired) break;
  }

  // Phase NightCap: Strict hard ceiling of maximum 2 people on Night shift (N = 2) on all days
  let nightCapDropped = false;
  for (let iter = 0; iter < 25; iter++) {
    const overDays = dateColumns.filter(({ isoDate }) => getDayCounts(isoDate).N > 2);
    if (overDays.length === 0) break;

    let reduced = false;
    for (const { index, isoDate } of overDays) {
      const empsOnN = rotating
        .filter((e) => rows[e][isoDate] === 'N')
        .sort((a, b) => getWorkDays(b) - getWorkDays(a));
      for (const emp of empsOnN) {
        rows[emp][isoDate] = 'OFF';
        if (isEmpValid(emp)) {
          reduced = true;
          nightCapDropped = true;
          break;
        }
        rows[emp][isoDate] = 'N';
      }
      if (reduced) break;

      // Fallback: slide N from isoDate to an adjacent day where N < 2
      for (const emp of empsOnN) {
        rows[emp][isoDate] = 'OFF';
        for (const adjIdx of [index - 1, index + 1]) {
          if (adjIdx < 0 || adjIdx >= dateColumns.length) continue;
          const adjIso = dateColumns[adjIdx].isoDate;
          if (rows[emp][adjIso] !== 'OFF' || getDayCounts(adjIso).N >= 2) continue;
          rows[emp][adjIso] = 'N';
          if (isEmpValid(emp)) {
            reduced = true;
            break;
          }
          rows[emp][adjIso] = 'OFF';
        }
        if (reduced) break;
        rows[emp][isoDate] = 'N';
      }
      if (reduced) break;
    }
    if (!reduced) break;
  }

  // Phase NightCap Compensation: Restore any employee pushed below targetWorkingDays
  if (nightCapDropped) {
    runPhaseB();
    runPhaseB2();
  }

  // Phase D: Coverage Equalization - Resolves understaffed days (N < 2, totalA < 2, rotM < 1) by shifting shifts within valid work blocks
  for (let iter = 0; iter < 40; iter++) {
    const cov = getDailyCoverage();
    let moved = false;

    for (let i = 0; i < dateColumns.length; i++) {
      const isoDate = dateColumns[i].isoDate;
      const c = cov[isoDate];

      const deficits: ('N' | 'A' | 'M')[] = [];
      if (c.N < 2) deficits.push('N');
      if (c.totalA < 2) deficits.push('A');
      if (c.rotM < 1) deficits.push('M');

      for (const targetShift of deficits) {
        for (const emp of rotating) {
          if (rows[emp][isoDate] !== 'OFF') continue;

          // Target shift must touch an existing block of targetShift or sit at an unbounded month edge
          if (!canAttachStandaloneAdd(emp, i, targetShift)) continue;

          rows[emp][isoDate] = targetShift;

          let bestDropIso: string | null = null;
          let bestShortRests = Infinity;
          for (let j = 0; j < dateColumns.length; j++) {
            if (i === j) continue;
            const dropIso = dateColumns[j].isoDate;
            const dropShift = rows[emp][dropIso];
            if (dropShift !== 'M' && dropShift !== 'A' && dropShift !== 'N') continue;

            const dropCov = cov[dropIso];
            const hasSurplus =
              (dropShift === 'M' && dropCov.rotM > 1) ||
              (dropShift === 'A' && dropCov.totalA > 2) ||
              (dropShift === 'N' && dropCov.N > 2);
            if (!hasSurplus) continue;

            rows[emp][dropIso] = 'OFF';
            if (isEmpValid(emp)) {
              const sr = countShortRests(emp);
              if (sr < bestShortRests) {
                bestShortRests = sr;
                bestDropIso = dropIso;
              }
            }
            rows[emp][dropIso] = dropShift;
          }

          if (bestDropIso) {
            rows[emp][bestDropIso] = 'OFF';
            moved = true;
            break;
          }
          rows[emp][isoDate] = 'OFF';
        }
        if (moved) break;
      }
      if (moved) break;
    }
    if (!moved) break;
  }

  // Phase D2: Multi-step & 2-employee coverage repair for remaining understaffed days
  for (let iter = 0; iter < 30; iter++) {
    const cov = getDailyCoverage();
    let repaired = false;

    for (let i = 0; i < dateColumns.length; i++) {
      const isoDate = dateColumns[i].isoDate;
      const c = cov[isoDate];

      const deficits: ('N' | 'A' | 'M')[] = [];
      if (c.N < 2) deficits.push('N');
      if (c.totalA < 2) deficits.push('A');
      if (c.rotM < 1) deficits.push('M');
      if (deficits.length === 0) continue;

      for (const targetShift of deficits) {
        // D2a: 1-employee block-slide (emp adds targetShift on isoDate, drops blocking d1 with surplus, slides d1->a2, drops surplus d2)
        for (const emp of rotating) {
          if (rows[emp][isoDate] !== 'OFF') continue;
          if (!canAttachStandaloneAdd(emp, i, targetShift)) continue;

          rows[emp][isoDate] = targetShift;
          const localStart = Math.max(0, i - 5);
          const localEnd = Math.min(dateColumns.length - 1, i + 5);

          for (let d1 = localStart; d1 <= localEnd; d1++) {
            if (d1 === i) continue;
            const iso1 = dateColumns[d1].isoDate;
            const sh1 = rows[emp][iso1];
            if (sh1 !== 'M' && sh1 !== 'A' && sh1 !== 'N') continue;
            if (!isBoundaryDrop(emp, d1, sh1) || !canDropShiftOnDay(iso1, sh1)) continue;

            rows[emp][iso1] = 'OFF';
            for (let a2 = Math.max(0, d1 - 6); a2 <= Math.min(dateColumns.length - 1, d1 + 6); a2++) {
              if (a2 === i || a2 === d1) continue;
              const iso2 = dateColumns[a2].isoDate;
              if (rows[emp][iso2] !== 'OFF') continue;
              if (!canAddShiftOnDay(iso2, sh1) || !canAttachStandaloneAdd(emp, a2, sh1)) continue;

              rows[emp][iso2] = sh1;
              for (let d2 = 0; d2 < dateColumns.length; d2++) {
                if (d2 === i || d2 === d1 || d2 === a2) continue;
                const iso3 = dateColumns[d2].isoDate;
                const sh3 = rows[emp][iso3];
                if (sh3 !== 'M' && sh3 !== 'A' && sh3 !== 'N') continue;
                if (!isBoundaryDrop(emp, d2, sh3) || !canDropShiftOnDay(iso3, sh3)) continue;

                rows[emp][iso3] = 'OFF';
                if (isEmpValid(emp)) {
                  repaired = true;
                  break;
                }
                rows[emp][iso3] = sh3;
              }
              if (repaired) break;
              rows[emp][iso2] = 'OFF';
            }
            if (repaired) break;
            rows[emp][iso1] = sh1;
          }
          if (repaired) break;
          rows[emp][isoDate] = 'OFF';
        }
        if (repaired) break;

        // D2b: 2-employee relay swap
        // emp1 adds targetShift on isoDate and drops midIso; emp2 adds midShift on midIso and drops surplusIso
        for (const emp1 of rotating) {
          if (rows[emp1][isoDate] !== 'OFF') continue;
          if (!canAttachStandaloneAdd(emp1, i, targetShift)) continue;

          rows[emp1][isoDate] = targetShift;
          for (let j = 0; j < dateColumns.length; j++) {
            if (j === i) continue;
            const midIso = dateColumns[j].isoDate;
            const midShift = rows[emp1][midIso];
            if (midShift !== 'M' && midShift !== 'A' && midShift !== 'N') continue;
            if (!isBoundaryDrop(emp1, j, midShift)) continue;

            rows[emp1][midIso] = 'OFF';
            if (isEmpValid(emp1)) {
              for (const emp2 of rotating) {
                if (emp2 === emp1 || rows[emp2][midIso] !== 'OFF') continue;
                if (!canAttachStandaloneAdd(emp2, j, midShift)) continue;
                rows[emp2][midIso] = midShift;
                for (let k = 0; k < dateColumns.length; k++) {
                  if (k === i || k === j) continue;
                  const surplusIso = dateColumns[k].isoDate;
                  const surShift = rows[emp2][surplusIso];
                  if (surShift !== 'M' && surShift !== 'A' && surShift !== 'N') continue;
                  if (!isBoundaryDrop(emp2, k, surShift) || !canDropShiftOnDay(surplusIso, surShift)) continue;

                  rows[emp2][surplusIso] = 'OFF';
                  if (isEmpValid(emp2)) {
                    repaired = true;
                    break;
                  }
                  rows[emp2][surplusIso] = surShift;
                }
                if (repaired) break;
                rows[emp2][midIso] = 'OFF';
              }
            }
            if (repaired) break;
            rows[emp1][midIso] = midShift;
          }
          if (repaired) break;
          rows[emp1][isoDate] = 'OFF';
        }
        if (repaired) break;

        // D2c: 2-employee relay where emp1 also slides an adjacent block (adds targetShift on isoDate, drops midIso, slides d1->a2; emp2 covers midIso and drops surplusIso)
        for (const emp1 of rotating) {
          if (rows[emp1][isoDate] !== 'OFF') continue;
          if (!canAttachStandaloneAdd(emp1, i, targetShift)) continue;

          rows[emp1][isoDate] = targetShift;
          const localStart = Math.max(0, i - 5);
          const localEnd = Math.min(dateColumns.length - 1, i + 5);

          for (let j = localStart; j <= localEnd; j++) {
            if (j === i) continue;
            const midIso = dateColumns[j].isoDate;
            const midShift = rows[emp1][midIso];
            if (midShift !== 'M' && midShift !== 'A' && midShift !== 'N') continue;
            if (!isBoundaryDrop(emp1, j, midShift)) continue;
            const validEmp2s = rotating.filter(
              (emp2) => emp2 !== emp1 && rows[emp2][midIso] === 'OFF' && canAttachStandaloneAdd(emp2, j, midShift)
            );
            if (validEmp2s.length === 0) continue;

            rows[emp1][midIso] = 'OFF';
            for (let a2 = Math.max(0, j - 6); a2 <= Math.min(dateColumns.length - 1, j + 6); a2++) {
              if (a2 === i || a2 === j) continue;
              const iso2 = dateColumns[a2].isoDate;
              if (rows[emp1][iso2] !== 'OFF') continue;
              if (!canAddShiftOnDay(iso2, midShift) || !canAttachStandaloneAdd(emp1, a2, midShift)) continue;

              rows[emp1][iso2] = midShift;
              for (let d2 = 0; d2 < dateColumns.length; d2++) {
                if (d2 === i || d2 === j || d2 === a2) continue;
                const dropIso2 = dateColumns[d2].isoDate;
                const dropSh2 = rows[emp1][dropIso2];
                if (dropSh2 !== 'M' && dropSh2 !== 'A' && dropSh2 !== 'N') continue;
                if (!isBoundaryDrop(emp1, d2, dropSh2) || !canDropShiftOnDay(dropIso2, dropSh2)) continue;

                rows[emp1][dropIso2] = 'OFF';
                if (isEmpValid(emp1)) {
                  for (const emp2 of validEmp2s) {
                    rows[emp2][midIso] = midShift;
                    for (let k = 0; k < dateColumns.length; k++) {
                      if (k === i || k === j || k === a2 || k === d2) continue;
                      const surplusIso = dateColumns[k].isoDate;
                      const surShift = rows[emp2][surplusIso];
                      if (surShift !== 'M' && surShift !== 'A' && surShift !== 'N') continue;
                      if (!isBoundaryDrop(emp2, k, surShift) || !canDropShiftOnDay(surplusIso, surShift)) continue;

                      rows[emp2][surplusIso] = 'OFF';
                      if (isEmpValid(emp2)) {
                        repaired = true;
                        break;
                      }
                      rows[emp2][surplusIso] = surShift;
                    }
                    if (repaired) break;
                    rows[emp2][midIso] = 'OFF';
                  }
                }
                if (repaired) break;
                rows[emp1][dropIso2] = dropSh2;
              }
              if (repaired) break;
              rows[emp1][iso2] = 'OFF';
            }
            if (repaired) break;
            rows[emp1][midIso] = midShift;
          }
          if (repaired) break;
          rows[emp1][isoDate] = 'OFF';
        }
        if (repaired) break;
      }
      if (repaired) break;
    }
    if (!repaired) break;
  }
}
