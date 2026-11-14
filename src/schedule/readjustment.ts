import type {
  RosterData,
  ComplianceTier,
  RelaxationRecord,
  ScheduleValidationReport,
  ReadjustmentAdjustment,
  ReadjustmentProposal
} from './types';
import {
  FIXED_MID_EMPLOYEES,
  FIXED_WEEKDAY_MA_EMPLOYEES,
  DAG_NEXT_SHIFT_ACT,
  DAG_LEAVE_CODES,
  shiftKey
} from './constants';
import {
  getActiveTailBlock,
  getActiveTailRest,
  extractRawHistoryTail
} from './historyUtils';
import {
  encodeGenDagState,
  canDoShift,
  buildDagDayActMasks,
  isDagPathValid,
  solveDagExtremeCount
} from './dagUtils';
import { validateScheduleCompliance } from './validator';
import {
  getTierConfig,
  isCoverageFeasibleForTier,
  tierCovPenalty,
  canSatisfyDailyCapacityForTier
} from '../relaxation';

/**
 * Intelligent, minimally-invasive shift readjustment algorithm to makeup for colleagues calling in sick (ABS).
 *
 * Rules:
 * 1. Minimum daily staffing: rotating M >= 1, total A >= 2, N = 2 (Tier 0 strict; relaxed in Tiers 1-3).
 * 2. Minimum invasiveness: If staffing on a shift with an ABS colleague is still >= 2, DO NOT change anyone.
 * 3. Morning deficit (M < 2):
 *    - Priority 1: Afternoon surplus on the same day (A > 2 -> move candidate to M).
 *      Candidate must NOT have worked Night on the previous day.
 *    - Priority 2: Eligible OFF colleague on that day.
 * 4. Afternoon deficit (A < 2):
 *    - Priority 1: Morning surplus on the same day (M > 2 -> move candidate to A).
 *      Candidate should not work Morning on the following day (to avoid short 8h turnaround).
 *    - Priority 2: Eligible OFF colleague on that day.
 * 5. Night deficit (N < 2):
 *    - Priority 1: Colleague currently on Afternoon (A) on the last day of their Afternoon block (next day is OFF).
 *      Convert A -> N. Night shift requires rest on subsequent day(s).
 *    - STRICT PROHIBITION: NEVER move someone on Morning shifts to Night (M -> N).
 *    - Priority 2: Colleague scheduled OFF who is about to start Night or has sufficient rest.
 *    - Priority 3: Other eligible Afternoon colleague if Afternoon has surplus.
 * 6. Non-rotating employees (Palade Alexandru-Ionut, Tir George Cristian) cannot be assigned to Night.
 * 7. Multi-variant support: variantSeed rotates candidate selection when multiple valid options exist.
 */
export function calculateAbsReadjustments(
  rows: Record<string, Record<string, string>>,
  dateColumns: { index: number; date: Date; isoDate: string }[],
  priorRoster: RosterData | null,
  targetYear: number,
  targetMonth: number,
  targetWorkingDays: number,
  variantSeed: number = 0
): ReadjustmentProposal {
  // Deep clone helper
  const cloneRows = (src: Record<string, Record<string, string>>) => {
    const c: Record<string, Record<string, string>> = {};
    for (const k of Object.keys(src)) {
      c[k] = { ...src[k] };
    }
    return c;
  };

  const workingRows = cloneRows(rows);
  const allEmployees = Object.keys(workingRows);
  const rotatingEmployees = allEmployees.filter(
    (e) => !FIXED_MID_EMPLOYEES.includes(e) && !FIXED_WEEKDAY_MA_EMPLOYEES.includes(e)
  );

  const adjustments: ReadjustmentAdjustment[] = [];
  const unresolvedDeficits: Array<{
    isoDate: string;
    shift: 'M' | 'A' | 'N';
    needed: number;
    available: number;
    reason: string;
  }> = [];

  const getShift = (src: Record<string, Record<string, string>>, emp: string, iso: string | null): string => {
    if (!iso) return 'OFF';
    return shiftKey(src[emp]?.[iso] ?? 'OFF');
  };

  // Penalty evaluation function using validateScheduleCompliance (strict Tier 0 in Sections 1-4)
  const evaluatePenalty = (testRows: Record<string, Record<string, string>>) => {
    const rep = validateScheduleCompliance(
      testRows,
      dateColumns,
      priorRoster,
      targetYear,
      targetMonth,
      targetWorkingDays,
      [],
      { tier: 0 }
    );
    const penalty =
      rep.totalViolations * 1000 +
      rep.understaffedDaysCount * 500 +
      rep.nightRestViolationCount * 400 +
      rep.streakViolationCount * 300 +
      rep.workDaysMismatchCount * 150 +
      rep.excessiveRestViolationCount * 50 +
      rep.mixedBlockViolationCount * 50;
    return { penalty, rep };
  };

  // Helper to find an offsetting day off for an employee called in from rest
  const findOffsetDayForEmp = (
    baseRows: Record<string, Record<string, string>>,
    emp: string,
    excludeIsoDate: string,
    pendingChanges: Array<{ isoDate: string; employee?: string; toShift: string }> = []
  ): { offsetIso: string; oldShift: string } | null => {
    const preRows = cloneRows(baseRows);
    for (const pc of pendingChanges) {
      const targetEmp = pc.employee ?? emp;
      if (!preRows[targetEmp]) preRows[targetEmp] = {};
      preRows[targetEmp][pc.isoDate] = pc.toShift;
    }

    let bestCandidate: { offsetIso: string; oldShift: string; violations: number; penalty: number } | null = null;
    for (let k = 0; k < dateColumns.length; k++) {
      const kIso = dateColumns[k].isoDate;
      if (kIso === excludeIsoDate) continue;
      if (pendingChanges.some((pc) => (pc.employee ?? emp) === emp && pc.isoDate === kIso)) continue;
      const curS = shiftKey(preRows[emp]?.[kIso] ?? 'OFF');
      if (curS !== 'M' && curS !== 'A') continue;

      let count = 0;
      if (curS === 'M') {
        for (const re of rotatingEmployees) {
          if (shiftKey(preRows[re]?.[kIso] ?? 'OFF') === 'M') count++;
        }
      } else {
        for (const ae of allEmployees) {
          if (shiftKey(preRows[ae]?.[kIso] ?? 'OFF') === 'A') count++;
        }
      }
      const minStaff = curS === 'M' ? 2 : 3;
      if (count >= minStaff) {
        const tempRows = cloneRows(preRows);
        tempRows[emp][kIso] = 'OFF';
        const { penalty, rep } = evaluatePenalty(tempRows);
        if (!rep.warnings.some((w) => w.startsWith(`${kIso}:`))) {
          if (
            !bestCandidate ||
            rep.totalViolations < bestCandidate.violations ||
            (rep.totalViolations === bestCandidate.violations && penalty < bestCandidate.penalty)
          ) {
            bestCandidate = {
              offsetIso: kIso,
              oldShift: curS,
              violations: rep.totalViolations,
              penalty
            };
          }
        }
      }
    }
    return bestCandidate ? { offsetIso: bestCandidate.offsetIso, oldShift: bestCandidate.oldShift } : null;
  };

  interface CandidatePlan {
    changes: Array<{
      isoDate: string;
      employee: string;
      fromShift: string;
      toShift: string;
      reason: string;
      absColleague: string;
    }>;
    penalty: number;
    violations: number;
  }

  const absBlockDayIndex: Record<string, number> = {};

  // 1. Process day by day in chronological order
  for (let d = 0; d < dateColumns.length; d++) {
    const col = dateColumns[d];
    const { isoDate } = col;
    const prevCol = d > 0 ? dateColumns[d - 1] : null;
    const nextCol = d < dateColumns.length - 1 ? dateColumns[d + 1] : null;
    const nextNextCol = d < dateColumns.length - 2 ? dateColumns[d + 2] : null;

    const prevIso = prevCol?.isoDate ?? null;
    const nextIso = nextCol?.isoDate ?? null;
    const nextNextIso = nextNextCol?.isoDate ?? null;

    const absColleagues = allEmployees.filter(
      (emp) => (workingRows[emp]?.[isoDate] ?? '').toUpperCase() === 'ABS'
    );
    const absNames = absColleagues.join(', ') || 'Colleague on sick leave';

    // Track absence block day indices
    for (const abs of absColleagues) {
      const isContinuation = prevIso && (workingRows[abs]?.[prevIso] ?? '').toUpperCase() === 'ABS';
      if (!isContinuation) {
        absBlockDayIndex[abs] = 0;
      } else {
        absBlockDayIndex[abs] = (absBlockDayIndex[abs] ?? 0) + 1;
      }
    }

    // Total staff counts on duty (including George Tir on weekdays)
    const totalM = allEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'M').length;
    const totalA = allEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'A').length;
    const totalN = allEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'N').length;

    // Staffing checks
    let mStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'M');
    let aStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'A');
    let nStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'N');
    const offStaff = rotatingEmployees.filter((emp) => {
      const s = getShift(workingRows, emp, isoDate);
      return s === 'OFF' || s === 'H8' || s === 'AD';
    });

    const neededN = Math.max(0, 2 - totalN);
    const neededM = Math.max(0, 1 - mStaff.length);
    const neededA = Math.max(0, 2 - totalA);

    // -------------------------------------------------------------------------
    // 1. NIGHT DEFICIT (N < 2)
    // -------------------------------------------------------------------------
    for (let i = 0; i < neededN; i++) {
      const plans: CandidatePlan[] = [];

      // Category 1: Afternoon staff on last day of block (next day is OFF)
      for (const emp of aStaff) {
        const nextS = getShift(workingRows, emp, nextIso);
        const isLastDay = nextS === 'OFF' || nextS === 'H8' || nextIso === null;
        if (!isLastDay) continue;

        const nextNextS = getShift(workingRows, emp, nextNextIso);
        const hasNaturalTwoDaysRest = nextNextS === 'OFF' || nextNextS === 'H8' || nextNextIso === null;

        // Plan 1A: Simple transfer A -> N
        const changes1: CandidatePlan['changes'] = [
          {
            isoDate,
            employee: emp,
            fromShift: 'A',
            toShift: 'N',
            reason: hasNaturalTwoDaysRest
              ? `Converted last day of Afternoon block to Night (A ➔ N). Natural 2-day post-night rest already scheduled (0 extra rest days needed).`
              : `Converted last day of Afternoon block to Night (A ➔ N) to cover sick colleague.`,
            absColleague: absNames
          }
        ];

        // If nextNext day is working, set it to OFF for mandatory 2-day night rest
        if (!hasNaturalTwoDaysRest && nextNextIso) {
          changes1.push({
            isoDate: nextNextIso,
            employee: emp,
            fromShift: nextNextS,
            toShift: 'OFF',
            reason: `Mandatory post-night rest day 2 granted after covering Night shift.`,
            absColleague: absNames
          });
        }

        // Domino backfill: If Afternoon drops below 2, attempt to backfill from Morning surplus (rotating M > 1)
        if (totalA <= 2 && mStaff.length > 1) {
          const mSurplusEmp = mStaff.find((me) => getShift(workingRows, me, nextIso) !== 'M');
          if (mSurplusEmp) {
            changes1.push({
              isoDate,
              employee: mSurplusEmp,
              fromShift: 'M',
              toShift: 'A',
              reason: `Domino rebalance: Moved from Morning surplus (${totalM} ➔ ${totalM - 1}) to Afternoon to maintain 2 Afternoon staff.`,
              absColleague: absNames
            });
          }
        }

        const test1 = cloneRows(workingRows);
        for (const c of changes1) {
          test1[c.employee][c.isoDate] = c.toShift;
        }
        const { penalty: pen1, rep: rep1 } = evaluatePenalty(test1);
        plans.push({ changes: changes1, penalty: pen1, violations: rep1.totalViolations });
      }

      // Category 2: Afternoon staff with surplus (totalA > 2)
      if (totalA > 2) {
        for (const emp of aStaff) {
          const changes2: CandidatePlan['changes'] = [
            {
              isoDate,
              employee: emp,
              fromShift: 'A',
              toShift: 'N',
              reason: `Moved from Afternoon surplus (${totalA} ➔ ${totalA - 1}) to Night (A ➔ N) to cover sick colleague.`,
              absColleague: absNames
            }
          ];
          if (nextIso && getShift(workingRows, emp, nextIso) !== 'OFF') {
            changes2.push({
              isoDate: nextIso,
              employee: emp,
              fromShift: getShift(workingRows, emp, nextIso),
              toShift: 'OFF',
              reason: `Mandatory post-night rest day 1.`,
              absColleague: absNames
            });
          }
          if (nextNextIso && getShift(workingRows, emp, nextNextIso) !== 'OFF') {
            changes2.push({
              isoDate: nextNextIso,
              employee: emp,
              fromShift: getShift(workingRows, emp, nextNextIso),
              toShift: 'OFF',
              reason: `Mandatory post-night rest day 2.`,
              absColleague: absNames
            });
          }

          const test2 = cloneRows(workingRows);
          for (const c of changes2) {
            test2[c.employee][c.isoDate] = c.toShift;
          }
          const { penalty: pen2, rep: rep2 } = evaluatePenalty(test2);
          plans.push({ changes: changes2, penalty: pen2, violations: rep2.totalViolations });
        }
      }

      // Category 3: OFF staff who didn't work Night on prev day
      const eligibleOff = offStaff.filter((emp) => getShift(workingRows, emp, prevIso) !== 'N');
      for (const emp of eligibleOff) {
        const changes3: CandidatePlan['changes'] = [
          {
            isoDate,
            employee: emp,
            fromShift: 'OFF',
            toShift: 'N',
            reason: `Assigned from rest to Night shift (OFF ➔ N) to cover sick colleague.`,
            absColleague: absNames
          }
        ];
        let clearedWorkingDays = 0;
        if (nextIso && getShift(workingRows, emp, nextIso) !== 'OFF') {
          const nextOrig = getShift(workingRows, emp, nextIso);
          if (nextOrig === 'M' || nextOrig === 'A' || nextOrig === 'N' || nextOrig === 'MID') {
            clearedWorkingDays++;
          }
          changes3.push({
            isoDate: nextIso,
            employee: emp,
            fromShift: nextOrig,
            toShift: 'OFF',
            reason: `Mandatory post-night rest day 1.`,
            absColleague: absNames
          });
        }
        if (nextNextIso && getShift(workingRows, emp, nextNextIso) !== 'OFF') {
          const nextNextOrig = getShift(workingRows, emp, nextNextIso);
          if (nextNextOrig === 'M' || nextNextOrig === 'A' || nextNextOrig === 'N' || nextNextOrig === 'MID') {
            clearedWorkingDays++;
          }
          changes3.push({
            isoDate: nextNextIso,
            employee: emp,
            fromShift: nextNextOrig,
            toShift: 'OFF',
            reason: `Mandatory post-night rest day 2.`,
            absColleague: absNames
          });
        }
        const offset = clearedWorkingDays === 0 ? findOffsetDayForEmp(workingRows, emp, isoDate, changes3) : null;
        if (offset) {
          changes3.push({
            isoDate: offset.offsetIso,
            employee: emp,
            fromShift: offset.oldShift,
            toShift: 'OFF',
            reason: `Offsetting day off to preserve legal ${targetWorkingDays}-day monthly norm (balances extra Night shift on ${isoDate}).`,
            absColleague: absNames
          });
        }

        const test3 = cloneRows(workingRows);
        for (const c of changes3) {
          test3[c.employee][c.isoDate] = c.toShift;
        }
        const { penalty: pen3, rep: rep3 } = evaluatePenalty(test3);
        plans.push({ changes: changes3, penalty: pen3, violations: rep3.totalViolations });
      }

      // Choose plan that minimizes penalties & warnings
      if (plans.length > 0) {
        plans.sort((a, b) => a.penalty - b.penalty);
        const bestPenalty = plans[0].penalty;
        const topPlans = plans.filter((p) => p.penalty === bestPenalty);
        const chosen = topPlans[(variantSeed + d + i) % topPlans.length];

        for (const c of chosen.changes) {
          workingRows[c.employee][c.isoDate] = c.toShift;
          adjustments.push(c);
        }

        mStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'M');
        aStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'A');
        nStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'N');
      } else {
        unresolvedDeficits.push({
          isoDate,
          shift: 'N',
          needed: 1,
          available: nStaff.length,
          reason: `No eligible colleague found for Night shift without violating mandatory rest (M ➔ N is strictly prohibited).`
        });
      }
    }

    // -------------------------------------------------------------------------
    // 2. MORNING DEFICIT (rotM < 1)
    // -------------------------------------------------------------------------
    for (let i = 0; i < neededM; i++) {
      const plans: CandidatePlan[] = [];
      const primaryAbs = absColleagues[0] || 'colleague';
      const dayInBlock = absBlockDayIndex[primaryAbs] ?? 0;

      // Category 1: Afternoon staff with surplus (totalA > 2)
      if (totalA > 2) {
        const candidates = aStaff.filter((emp) => getShift(workingRows, emp, prevIso) !== 'N');
        // Sort so Prisecaru is first, then Olteanu (Lucian), then others
        candidates.sort((a, b) => {
          if (a.includes('Prisecaru')) return -1;
          if (b.includes('Prisecaru')) return 1;
          if (a.includes('Olteanu')) return -1;
          if (b.includes('Olteanu')) return 1;
          return a.localeCompare(b);
        });

        for (const emp of candidates) {
          const changes: CandidatePlan['changes'] = [
            {
              isoDate,
              employee: emp,
              fromShift: 'A',
              toShift: 'M',
              reason: `Moved from Afternoon surplus (${totalA} ➔ ${totalA - 1}) to Morning (A ➔ M) to cover sick colleague (0 norm change).`,
              absColleague: absNames
            }
          ];
          const test = cloneRows(workingRows);
          test[emp][isoDate] = 'M';
          const { penalty, rep } = evaluatePenalty(test);
          plans.push({ changes, penalty, violations: rep.totalViolations });
        }
      }

      // Category 2: OFF staff
      const eligibleOff = offStaff.filter((emp) => getShift(workingRows, emp, prevIso) !== 'N');
      eligibleOff.sort((a, b) => {
        if (a.includes('Urzica')) return -1;
        if (b.includes('Urzica')) return 1;
        return a.localeCompare(b);
      });

      for (const emp of eligibleOff) {
        const preChanges: CandidatePlan['changes'] = [
          {
            isoDate,
            employee: emp,
            fromShift: 'OFF',
            toShift: 'M',
            reason: `Assigned from rest to Morning shift (OFF ➔ M) to ensure minimum of 2 Morning staff.`,
            absColleague: absNames
          }
        ];
        const offset = findOffsetDayForEmp(workingRows, emp, isoDate, preChanges);
        const changes = [...preChanges];
        if (offset) {
          changes.push({
            isoDate: offset.offsetIso,
            employee: emp,
            fromShift: offset.oldShift,
            toShift: 'OFF',
            reason: `Offsetting day off granted to preserve legal ${targetWorkingDays}-day monthly norm (balances extra Morning shift on ${isoDate}).`,
            absColleague: absNames
          });
        }
        const test = cloneRows(workingRows);
        for (const c of changes) {
          test[c.employee][c.isoDate] = c.toShift;
        }
        const { penalty, rep } = evaluatePenalty(test);
        plans.push({ changes, penalty, violations: rep.totalViolations });
      }

      if (plans.length > 0) {
        plans.sort((a, b) => a.violations - b.violations || a.penalty - b.penalty);
        const minViolations = plans[0].violations;
        const minViolPlans = plans.filter((p) => p.violations === minViolations);
        let chosen: CandidatePlan;

        const prisecaruPlan = minViolPlans.find((p) => p.changes.some((c) => c.employee.includes('Prisecaru')));
        const lucianPlan = minViolPlans.find((p) => p.changes.some((c) => c.employee.includes('Olteanu')));
        const urzicaPlan = minViolPlans.find((p) => p.changes.some((c) => c.employee.includes('Urzica')));

        // Strategic variant assignment (only among minimum-violation plans):
        if (variantSeed === 0) {
          // Variant 1: Prefer Prisecaru when equally compliant
          chosen = prisecaruPlan ?? lucianPlan ?? minViolPlans[0];
        } else if (variantSeed === 1) {
          // Variant 2: Urzica 1 morning and Lucian 1 morning when equally compliant
          if (dayInBlock === 0) {
            chosen = urzicaPlan ?? prisecaruPlan ?? minViolPlans[0];
          } else {
            chosen = lucianPlan ?? prisecaruPlan ?? minViolPlans[0];
          }
        } else if (variantSeed === 2) {
          // Variant 3: Prefer Lucian when equally compliant
          chosen = lucianPlan ?? prisecaruPlan ?? minViolPlans[0];
        } else {
          // Generic cycling across top variants
          const bestPenalty = minViolPlans[0].penalty;
          const topPlans = minViolPlans.filter((p) => p.penalty === bestPenalty);
          chosen = topPlans[(variantSeed + dayInBlock + i) % topPlans.length];
        }

        for (const c of chosen.changes) {
          workingRows[c.employee][c.isoDate] = c.toShift;
          adjustments.push(c);
        }

        mStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'M');
        aStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'A');
        nStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'N');
      } else {
        unresolvedDeficits.push({
          isoDate,
          shift: 'M',
          needed: 1,
          available: mStaff.length,
          reason: `No eligible colleague from Afternoon surplus or scheduled rest could cover Morning.`
        });
      }
    }

    // -------------------------------------------------------------------------
    // 3. AFTERNOON DEFICIT (A < 2)
    // -------------------------------------------------------------------------
    for (let i = 0; i < neededA; i++) {
      const plans: CandidatePlan[] = [];

      // Category 1: Morning staff with surplus (totalM > 2)
      if (totalM > 2) {
        const candidates = mStaff.filter((emp) => getShift(workingRows, emp, nextIso) !== 'M');
        for (const emp of candidates) {
          const changes: CandidatePlan['changes'] = [
            {
              isoDate,
              employee: emp,
              fromShift: 'M',
              toShift: 'A',
              reason: `Moved from Morning surplus (${totalM} ➔ ${totalM - 1}) to Afternoon (M ➔ A) to cover sick colleague (0 norm change).`,
              absColleague: absNames
            }
          ];
          const test = cloneRows(workingRows);
          test[emp][isoDate] = 'A';
          const { penalty, rep } = evaluatePenalty(test);
          plans.push({ changes, penalty, violations: rep.totalViolations });
        }
      }

      // Category 2: OFF staff (keep fixed weekday M/A parity intact)
      const eligibleOff = offStaff.filter((emp) => getShift(workingRows, emp, prevIso) !== 'N');
      for (const emp of eligibleOff) {
        const preChanges: CandidatePlan['changes'] = [
          {
            isoDate,
            employee: emp,
            fromShift: 'OFF',
            toShift: 'A',
            reason: `Assigned from rest to Afternoon shift (OFF ➔ A) to maintain minimum of 2 Afternoon staff.`,
            absColleague: absNames
          }
        ];
        const offset = findOffsetDayForEmp(workingRows, emp, isoDate, preChanges);
        const changes = [...preChanges];
        if (offset) {
          changes.push({
            isoDate: offset.offsetIso,
            employee: emp,
            fromShift: offset.oldShift,
            toShift: 'OFF',
            reason: `Offsetting day off granted to preserve legal ${targetWorkingDays}-day monthly norm (balances extra Afternoon shift on ${isoDate}).`,
            absColleague: absNames
          });
        }
        const test = cloneRows(workingRows);
        for (const c of changes) {
          test[c.employee][c.isoDate] = c.toShift;
        }
        const { penalty, rep } = evaluatePenalty(test);
        plans.push({ changes, penalty, violations: rep.totalViolations });
      }

      if (plans.length > 0) {
        plans.sort((a, b) => a.penalty - b.penalty);
        const bestPenalty = plans[0].penalty;
        const topPlans = plans.filter((p) => p.penalty === bestPenalty);
        const chosen = topPlans[(variantSeed + d + i) % topPlans.length];

        for (const c of chosen.changes) {
          workingRows[c.employee][c.isoDate] = c.toShift;
          adjustments.push(c);
        }

        mStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'M');
        aStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'A');
        nStaff = rotatingEmployees.filter((emp) => getShift(workingRows, emp, isoDate) === 'N');
      } else {
        unresolvedDeficits.push({
          isoDate,
          shift: 'A',
          needed: 1,
          available: aStaff.length,
          reason: `No eligible colleague from Morning surplus or scheduled rest could cover Afternoon.`
        });
      }
    }
  }

  // ---------------------------------------------------------------------------
  // 4. GLOBAL REBALANCING POST-PASS (Eliminate any remaining warnings at Tier 0)
  // ---------------------------------------------------------------------------
  let finalTier: ComplianceTier = 0;
  let finalRelaxations: RelaxationRecord[] = [];
  let currentValidation = validateScheduleCompliance(
    workingRows,
    dateColumns,
    priorRoster,
    targetYear,
    targetMonth,
    targetWorkingDays,
    [],
    { tier: 0 }
  );

  let improved = true;
  let maxPasses = 5;
  while (improved && maxPasses > 0 && currentValidation.totalViolations > 0) {
    improved = false;
    maxPasses--;

    for (const emp of rotatingEmployees) {
      const stats = currentValidation.employeeStats[emp];
      if (!stats) continue;

      // Case A: Employee worked more days than target (e.g. called in from OFF)
      if (stats.totalWorkingDays > stats.targetWorkingDays) {
        let bestOffset: { kIso: string; curS: string; rep: typeof currentValidation } | null = null;
        for (let k = 0; k < dateColumns.length; k++) {
          const kIso = dateColumns[k].isoDate;
          const curS = shiftKey(workingRows[emp]?.[kIso] ?? 'OFF');
          if (curS !== 'M' && curS !== 'A') continue;

          // Morning requires >= 1 rotating M after removal; Afternoon requires >= 2 total A (incl. weekday M/A) after removal
          const curCount = curS === 'M'
            ? rotatingEmployees.filter((re) => shiftKey(workingRows[re]?.[kIso] ?? 'OFF') === 'M').length
            : allEmployees.filter((ae) => shiftKey(workingRows[ae]?.[kIso] ?? 'OFF') === 'A').length;

          if (curCount >= (curS === 'M' ? 2 : 3)) {
            const testRows = cloneRows(workingRows);
            testRows[emp][kIso] = 'OFF';
            const testRep = validateScheduleCompliance(
              testRows,
              dateColumns,
              priorRoster,
              targetYear,
              targetMonth,
              targetWorkingDays,
              [],
              { tier: 0 }
            );

            if (testRep.totalViolations < currentValidation.totalViolations) {
              if (!bestOffset || testRep.totalViolations < bestOffset.rep.totalViolations) {
                bestOffset = { kIso, curS, rep: testRep };
              }
            }
          }
        }
        if (bestOffset) {
          workingRows[emp][bestOffset.kIso] = 'OFF';
          currentValidation = bestOffset.rep;
          adjustments.push({
            isoDate: bestOffset.kIso,
            employee: emp,
            fromShift: bestOffset.curS,
            toShift: 'OFF',
            reason: `Rebalanced: Added offsetting day off to preserve legal monthly norm of ${targetWorkingDays} working days.`,
            absColleague: 'Rebalancing'
          });
          improved = true;
        }
      }
    }
  }

  // ---------------------------------------------------------------------------
  // 5. EXACT MINIMAL-DIFF DP+SA FALLBACK REPAIR ACROSS TIERS [0, 1, 2, 3]
  // ---------------------------------------------------------------------------
  if (currentValidation.totalViolations > 0 && rotatingEmployees.length > 0) {
    const s5StartMs = performance.now();
    const s5BudgetMs = 3500;
    const s5DeadlineMs = s5StartMs + s5BudgetMs - 100;
    const absTiers: ComplianceTier[] = [0, 1, 2, 3];
    const daysCount = dateColumns.length;
    const GEN_NEXT = DAG_NEXT_SHIFT_ACT;
    const shToGen = (s: string) => (s === 'A' ? 0 : s === 'M' ? 1 : s === 'N' ? 2 : s === 'ABS' ? 4 : (s === 'H8' || s === 'HUP') ? 5 : 3);
    const fixedMidSet = new Set(FIXED_MID_EMPLOYEES);
    const fixedMaEmps = allEmployees.filter((e) => FIXED_WEEKDAY_MA_EMPLOYEES.includes(e));

    let firstGlobalAbsIdx = daysCount;
    const empFirstAbsIndices: number[] = [];
    for (const emp of rotatingEmployees) {
      let eFirst = daysCount;
      for (let d = 0; d < daysCount; d++) {
        if (shiftKey(rows[emp]?.[dateColumns[d].isoDate] ?? 'OFF') === 'ABS') {
          eFirst = d;
          if (d < firstGlobalAbsIdx) firstGlobalAbsIdx = d;
          break;
        }
      }
      if (eFirst < daysCount) empFirstAbsIndices.push(eFirst);
    }
    for (const emp of fixedMaEmps) {
      for (let d = 0; d < daysCount; d++) {
        if (shiftKey(rows[emp]?.[dateColumns[d].isoDate] ?? 'OFF') === 'ABS') {
          if (d < firstGlobalAbsIdx) firstGlobalAbsIdx = d;
          break;
        }
      }
    }
    if (firstGlobalAbsIdx === daysCount) firstGlobalAbsIdx = 0;
    const hasDistinctEmpAbsStarts = empFirstAbsIndices.some((idx) => idx > firstGlobalAbsIdx);

    // Extract raw 6-day contiguous history from priorRoster preserving ABS/H8/CM/OFF/work sequence
    const empHistories: Record<string, string[]> = {};
    for (const emp of rotatingEmployees) {
      empHistories[emp] = extractRawHistoryTail(priorRoster, emp, targetYear, targetMonth, 6);
    }

    interface GenEdge {
      act: number; // 0:A, 1:M, 2:N, 3:OFF(or Tir M/OFF/ABS/H8), 4:ABS, 5:H8/HUP
      nextNode: number;
      penA: number;
      penB: number;
    }
    interface GenDag {
      emp: string;
      isFixedTir: boolean;
      origCodes: string[];
      origActs: number[];
      layers: GenEdge[][][];
      feasible: boolean;
      maxNodes: number;
      cAct: Int8Array[];
      cNext: Int32Array[];
    }

    const buildTirGenDag = (emp: string): GenDag => {
      const origCodes = dateColumns.map(({ isoDate }) => shiftKey(rows[emp]?.[isoDate] ?? 'OFF'));
      const origActs = origCodes.map((c) => (c === 'A' ? 0 : 3));
      const layers: GenEdge[][][] = new Array(daysCount);
      for (let d = 0; d < daysCount; d++) {
        // Lock Tir strictly to origActs[d] to preserve H6 ISO-week M/A parity and leave/off days
        layers[d] = [[{ act: origActs[d], nextNode: 0, penA: 0, penB: 0 }]];
      }
      return {
        emp,
        isFixedTir: true,
        origCodes,
        origActs,
        layers,
        feasible: true,
        maxNodes: 1,
        cAct: Array.from({ length: daysCount }, () => new Int8Array(1)),
        cNext: Array.from({ length: daysCount }, () => new Int32Array(1))
      };
    };

    const LEAVE_CODES = DAG_LEAVE_CODES;

    const buildGeneralDag = (
      emp: string,
      empIdx: number,
      minNonNightRest: number,
      allowShortBlock2: boolean,
      lockMode: 'globalPre' | 'empPre' | 'none'
    ): GenDag => {
      const origCodes = dateColumns.map(({ isoDate }) => shiftKey(rows[emp]?.[isoDate] ?? 'OFF'));
      const origActs = origCodes.map((c) => shToGen(c));
      let empFirstAbsIdx = Infinity;
      let fixedNonWorkDrops = 0;
      let empHasAbsOrH8 = false;
      for (let d = 0; d < daysCount; d++) {
        if (origActs[d] === 4) {
          if (d < empFirstAbsIdx) empFirstAbsIdx = d;
          fixedNonWorkDrops++;
          empHasAbsOrH8 = true;
        } else if (origActs[d] === 5) {
          fixedNonWorkDrops++;
          empHasAbsOrH8 = true;
        }
      }
      const effectiveTarget = Math.max(0, targetWorkingDays - fixedNonWorkDrops);
      const editableDays = origActs.filter((a) => a !== 4 && a !== 5).length;
      if (effectiveTarget > editableDays) {
        return { emp, isFixedTir: false, origCodes, origActs, layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
      }

      const dayAllowed: number[][] = new Array(daysCount);
      for (let d = 0; d < daysCount; d++) {
        const oa = origActs[d];
        if (oa === 4 || oa === 5) {
          dayAllowed[d] = [oa];
        } else if (lockMode === 'globalPre' && d < firstGlobalAbsIdx) {
          dayAllowed[d] = [oa];
        } else if (lockMode === 'empPre' && empFirstAbsIdx < Infinity && d < empFirstAbsIdx) {
          dayAllowed[d] = [oa];
        } else {
          dayAllowed[d] = [0, 1, 2, 3];
        }
      }

      const edgePens = (d: number, act: number, restPen: number): { penA: number; penB: number } => {
        if (act === origActs[d]) return { penA: restPen, penB: restPen * 2 };
        const tie = variantSeed > 0 ? (((empIdx * 31 + d * 17 + act * 7 + variantSeed * 13) % 7) * 0.1) : 0;
        if (d < empFirstAbsIdx && empFirstAbsIdx < Infinity) {
          return { penA: restPen + 120 + tie, penB: restPen * 2 + 1200 + tie };
        }
        if (d < firstGlobalAbsIdx) {
          return { penA: restPen + 40 + tie, penB: restPen * 2 + 400 + tie };
        }
        if (empHasAbsOrH8) {
          return { penA: restPen + 8 + tie, penB: restPen * 2 + 120 + tie };
        }
        return { penA: restPen + 3 + tie, penB: restPen * 2 + 45 + tie };
      };

      const numStates = (effectiveTarget + 1) * 4 * 13 * 4 * 2;
      const enc = encodeGenDagState;

      const rawHistory = empHistories[emp] ?? [];
      const history = rawHistory.map((s) =>
        s === 'HUP' || s === 'H8' || s === 'ABS' || s === 'AD' || s === 'CM' ? 'OFF' : s
      );
      const activeTailWork = getActiveTailBlock(history);
      const activeTailRest = activeTailWork.length === 0 ? getActiveTailRest(history) : [];
      const idxBefore = rawHistory.length - 1 - activeTailWork.length;
      const adjLeaveBefore = idxBefore >= 0 && LEAVE_CODES.has(rawHistory[idxBefore]);
      let lastHistWork: string | null = null;
      for (let i = history.length - 1; i >= 0; i--) {
        if (history[i] === 'A' || history[i] === 'M' || history[i] === 'N') {
          lastHistWork = history[i];
          break;
        }
      }

      const allowed0 = dayAllowed[0];
      const startEdges: Array<{ act: number; nextState: number; penA: number; penB: number }> = [];
      if (activeTailWork.length > 0) {
        const curShift = activeTailWork[activeTailWork.length - 1];
        const curIdx = shToGen(curShift);
        const totalWorkTail = activeTailWork.length;
        let sameLen = 0;
        for (let i = activeTailWork.length - 1; i >= 0 && activeTailWork[i] === curShift; i--) sameLen++;
        if (curIdx <= 2) {
          if (totalWorkTail < 5 && 1 <= effectiveTarget && allowed0.includes(curIdx)) {
            const p = edgePens(0, curIdx, 0);
            const nextMl = adjLeaveBefore ? 5 + totalWorkTail : totalWorkTail;
            startEdges.push({ act: curIdx, nextState: enc(1, curIdx, nextMl, 0, 0), ...p });
          }
          if (sameLen >= 3 || adjLeaveBefore) {
            if (allowed0.includes(3)) {
              startEdges.push({ act: 3, nextState: enc(0, curIdx, 10, 1, 0), ...edgePens(0, 3, 0) });
            }
          }
          if (allowed0.includes(4)) {
            startEdges.push({ act: 4, nextState: enc(0, curIdx, 10, 0, 0), ...edgePens(0, 4, 0) });
          }
          if (allowed0.includes(5)) {
            startEdges.push({ act: 5, nextState: enc(0, curIdx, 10, 0, 0), ...edgePens(0, 5, 0) });
          }
        }
      } else if (history.length > 0) {
        const restLen = activeTailRest.length;
        const lastIdx = lastHistWork ? shToGen(lastHistWork) : 3;
        const initMl = Math.min(12, 10 + restLen);
        if (restLen < 3 && allowed0.includes(3)) {
          startEdges.push({ act: 3, nextState: enc(0, lastIdx, initMl, restLen + 1, 0), ...edgePens(0, 3, 0) });
        }
        if (allowed0.includes(4)) {
          startEdges.push({ act: 4, nextState: enc(0, lastIdx, initMl, 0, 0), ...edgePens(0, 4, 0) });
        }
        if (allowed0.includes(5)) {
          startEdges.push({ act: 5, nextState: enc(0, lastIdx, initMl, 0, 0), ...edgePens(0, 5, 0) });
        }
        const reqRest = lastIdx === 2 ? 2 : (adjLeaveBefore ? 1 : minNonNightRest);
        if (1 <= effectiveTarget && (lastIdx === 3 || restLen >= reqRest)) {
          const rPen = lastIdx !== 3 && lastIdx !== 2 && restLen === 1 && !adjLeaveBefore ? 15 : 0;
          const firstMl = adjLeaveBefore ? 5 : 0;
          if (lastIdx === 3) {
            for (const ns of [0, 1, 2]) {
              if (allowed0.includes(ns)) {
                startEdges.push({ act: ns, nextState: enc(1, ns, firstMl, 0, 0), ...edgePens(0, ns, 0) });
              }
            }
          } else {
            const ns = GEN_NEXT[lastIdx];
            if (allowed0.includes(ns)) {
              startEdges.push({ act: ns, nextState: enc(1, ns, firstMl, 0, 0), ...edgePens(0, ns, rPen) });
            }
            if (adjLeaveBefore && allowed0.includes(lastIdx)) {
              startEdges.push({ act: lastIdx, nextState: enc(1, lastIdx, firstMl, 0, 0), ...edgePens(0, lastIdx, 0) });
            }
          }
        }
      } else {
        if (allowed0.includes(3)) {
          startEdges.push({ act: 3, nextState: enc(0, 3, 10, 1, 0), ...edgePens(0, 3, 0) });
        }
        if (allowed0.includes(4)) {
          startEdges.push({ act: 4, nextState: enc(0, 3, 10, 0, 0), ...edgePens(0, 4, 0) });
        }
        if (allowed0.includes(5)) {
          startEdges.push({ act: 5, nextState: enc(0, 3, 10, 0, 0), ...edgePens(0, 5, 0) });
        }
        if (1 <= effectiveTarget) {
          for (const ns of [0, 1, 2]) {
            if (allowed0.includes(ns)) {
              startEdges.push({ act: ns, nextState: enc(1, ns, 5, 0, 0), ...edgePens(0, ns, 0) });
            }
          }
        }
      }

      if (startEdges.length === 0) {
        return { emp, isFixedTir: false, origCodes, origActs, layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
      }

      const activeStates: number[][] = Array.from({ length: daysCount + 1 }, () => []);
      const seen = new Uint8Array(numStates);
      for (const e of startEdges) {
        if (!seen[e.nextState]) {
          seen[e.nextState] = 1;
          activeStates[1].push(e.nextState);
        }
      }

      const stepTransitions = (d: number, st: number) => {
        const sb = st & 1;
        const t0 = st >> 1;
        const offS = t0 % 4;
        const t1 = (t0 - offS) / 4;
        const ml = t1 % 13;
        const t2 = (t1 - ml) / 13;
        const sh = t2 % 4;
        const w = (t2 - sh) / 4;
        const allowed = dayAllowed[d];
        const prevWasLeave = origActs[d - 1] === 4 || origActs[d - 1] === 5;
        const out: Array<{ act: number; nextState: number; penA: number; penB: number }> = [];

        if (ml < 10) {
          const isUnb = ml >= 5;
          const len = isUnb ? ml - 4 : ml + 1;
          if (len < 5 && w + 1 <= effectiveTarget && allowed.includes(sh)) {
            out.push({ act: sh, nextState: enc(w + 1, sh, ml + 1, 0, sb), ...edgePens(d, sh, 0) });
          }
          if (len >= 3 || isUnb) {
            if (allowed.includes(3)) {
              out.push({ act: 3, nextState: enc(w, sh, 10, 1, sb), ...edgePens(d, 3, 0) });
            }
          } else if (allowShortBlock2 && len === 2 && sb === 0 && allowed.includes(3)) {
            out.push({ act: 3, nextState: enc(w, sh, 10, 1, 1), ...edgePens(d, 3, 25) });
          }
          // Ending work block immediately before ABS/H8 is always exempt (adjAfter)
          if (allowed.includes(4)) {
            out.push({ act: 4, nextState: enc(w, sh, 10, 0, sb), ...edgePens(d, 4, 0) });
          }
          if (allowed.includes(5)) {
            out.push({ act: 5, nextState: enc(w, sh, 10, 0, sb), ...edgePens(d, 5, 0) });
          }
        } else {
          const restLen = ml - 9;
          const nextMl = Math.min(12, ml + 1);
          if (offS < 3 && allowed.includes(3)) {
            out.push({ act: 3, nextState: enc(w, sh, nextMl, offS + 1, sb), ...edgePens(d, 3, 0) });
          }
          if (allowed.includes(4)) {
            out.push({ act: 4, nextState: enc(w, sh, nextMl, 0, sb), ...edgePens(d, 4, 0) });
          }
          if (allowed.includes(5)) {
            out.push({ act: 5, nextState: enc(w, sh, nextMl, 0, sb), ...edgePens(d, 5, 0) });
          }
          if (w + 1 <= effectiveTarget) {
            const nextMlStart = prevWasLeave ? 5 : 0;
            if (sh === 3) {
              for (const ns of [0, 1, 2]) {
                if (allowed.includes(ns)) {
                  out.push({ act: ns, nextState: enc(w + 1, ns, nextMlStart, 0, sb), ...edgePens(d, ns, 0) });
                }
              }
            } else if (restLen >= (sh === 2 ? 2 : (prevWasLeave ? 1 : minNonNightRest))) {
              const ns = GEN_NEXT[sh];
              if (allowed.includes(ns)) {
                const rPen = sh !== 2 && restLen === 1 && !prevWasLeave ? 15 : 0;
                out.push({ act: ns, nextState: enc(w + 1, ns, nextMlStart, 0, sb), ...edgePens(d, ns, rPen) });
              }
              if (prevWasLeave && allowed.includes(sh)) {
                out.push({ act: sh, nextState: enc(w + 1, sh, nextMlStart, 0, sb), ...edgePens(d, sh, 0) });
              }
            }
          }
        }
        return out;
      };

      for (let d = 1; d < daysCount; d++) {
        seen.fill(0);
        for (const st of activeStates[d]) {
          for (const tr of stepTransitions(d, st)) {
            if (!seen[tr.nextState]) {
              seen[tr.nextState] = 1;
              activeStates[d + 1].push(tr.nextState);
            }
          }
        }
      }

      const reachNext = new Int32Array(numStates).fill(-1);
      let termIdx = 0;
      for (const st of activeStates[daysCount]) {
        const sb = st & 1;
        const t0 = st >> 1;
        const offS = t0 % 4;
        const t1 = (t0 - offS) / 4;
        const ml = t1 % 13;
        const t2 = (t1 - ml) / 13;
        const sh = t2 % 4;
        const w = (t2 - sh) / 4;
        if (w === effectiveTarget && (sb === 0 || allowShortBlock2)) {
          reachNext[st] = termIdx++;
        }
      }
      if (termIdx === 0) {
        return { emp, isFixedTir: false, origCodes, origActs, layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
      }

      const layers: GenEdge[][][] = new Array(daysCount);
      const reachCur = new Int32Array(numStates);
      for (let d = daysCount - 1; d >= 1; d--) {
        reachCur.fill(-1);
        const dayLayer: GenEdge[][] = [];
        for (const st of activeStates[d]) {
          const nodeEdges: GenEdge[] = [];
          for (const tr of stepTransitions(d, st)) {
            const nxtIdx = reachNext[tr.nextState];
            if (nxtIdx !== -1) {
              nodeEdges.push({ act: tr.act, nextNode: nxtIdx, penA: tr.penA, penB: tr.penB });
            }
          }
          if (nodeEdges.length > 0) {
            reachCur[st] = dayLayer.length;
            dayLayer.push(nodeEdges);
          }
        }
        if (dayLayer.length === 0) {
          return { emp, isFixedTir: false, origCodes, origActs, layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
        }
        layers[d] = dayLayer;
        reachNext.set(reachCur);
      }

      const rootEdges: GenEdge[] = [];
      for (const e of startEdges) {
        const nxtIdx = reachNext[e.nextState];
        if (nxtIdx !== -1) {
          rootEdges.push({ act: e.act, nextNode: nxtIdx, penA: e.penA, penB: e.penB });
        }
      }
      if (rootEdges.length === 0) {
        return { emp, isFixedTir: false, origCodes, origActs, layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
      }
      layers[0] = [rootEdges];

      let maxNodes = 1;
      const cAct: Int8Array[] = new Array(daysCount);
      const cNext: Int32Array[] = new Array(daysCount);
      for (let d = 0; d < daysCount; d++) {
        const len = layers[d].length;
        if (len > maxNodes) maxNodes = len;
        cAct[d] = new Int8Array(len);
        cNext[d] = new Int32Array(len);
      }
      return { emp, isFixedTir: false, origCodes, origActs, layers, feasible: true, maxNodes, cAct, cNext };
    };

    let dpNextGen = new Float64Array(512);
    let dpCurGen = new Float64Array(512);
    const dayCostGen = new Float64Array(daysCount * 4);

    const covPenaltyA = (m: number, a: number, n: number, dw: number = 1, _useSoft: boolean = false, tier: ComplianceTier = 0) =>
      tierCovPenalty(m, a, n, dw, tier);

    const covPenaltyB = (m: number, a: number, n: number, tier: ComplianceTier = 0) => {
      if (!isCoverageFeasibleForTier(m, a, n, tier)) {
        return tierCovPenalty(m, a, n, 100, tier);
      }
      return (tier >= 1 && n === 3 ? 5000 : 0) + (tier >= 3 && (m === 0 || a === 1) ? 15000 : 0);
    };

    // mode: 0 = gentle minimal-edit Phase A1, 1 = deep Stage-3 FastSA Phase A2, 2 = strict minimal-diff Phase B
    const solveGenPath = (
      dag: GenDag,
      otherM: Int32Array,
      otherA: Int32Array,
      otherN: Int32Array,
      dayWeights: Float64Array | null,
      noise: Float64Array | null,
      mode: 0 | 1 | 2,
      tier: ComplianceTier
    ): number[] => {
      if (dag.maxNodes > dpNextGen.length) {
        dpNextGen = new Float64Array(dag.maxNodes + 64);
        dpCurGen = new Float64Array(dag.maxNodes + 64);
      }
      const useSoft = mode === 1;
      if (dag.isFixedTir) {
        for (let d = 0; d < daysCount; d++) {
          const m0 = otherM[d];
          const a0 = otherA[d];
          const n0 = otherN[d];
          const dw = dayWeights ? dayWeights[d] : 1;
          const base = d * 4;
          dayCostGen[base + 0] = (mode === 2 ? covPenaltyB(m0, a0 + 1, n0, tier) : covPenaltyA(m0, a0 + 1, n0, dw, useSoft, tier)) + (noise ? noise[base + 0] : 0);
          dayCostGen[base + 3] = (mode === 2 ? covPenaltyB(m0, a0, n0, tier) : covPenaltyA(m0, a0, n0, dw, useSoft, tier)) + (noise ? noise[base + 3] : 0);
        }
      } else {
        for (let d = 0; d < daysCount; d++) {
          const m0 = otherM[d];
          const a0 = otherA[d];
          const n0 = otherN[d];
          const dw = dayWeights ? dayWeights[d] : 1;
          const base = d * 4;
          if (mode === 2) {
            dayCostGen[base + 0] = covPenaltyB(m0, a0 + 1, n0, tier);
            dayCostGen[base + 1] = covPenaltyB(m0 + 1, a0, n0, tier);
            dayCostGen[base + 2] = covPenaltyB(m0, a0, n0 + 1, tier);
            dayCostGen[base + 3] = covPenaltyB(m0, a0, n0, tier);
          } else {
            dayCostGen[base + 0] = covPenaltyA(m0, a0 + 1, n0, dw, useSoft, tier) + (noise ? noise[base + 0] : 0);
            dayCostGen[base + 1] = covPenaltyA(m0 + 1, a0, n0, dw, useSoft, tier) + (noise ? noise[base + 1] : 0);
            dayCostGen[base + 2] = covPenaltyA(m0, a0, n0 + 1, dw, useSoft, tier) + (noise ? noise[base + 2] : 0);
            dayCostGen[base + 3] = covPenaltyA(m0, a0, n0, dw, useSoft, tier) + (noise ? noise[base + 3] : 0);
          }
        }
      }

      const { layers, cAct, cNext } = dag;
      for (let d = daysCount - 1; d >= 0; d--) {
        const layer = layers[d];
        const isLast = d === daysCount - 1;
        const cA = cAct[d];
        const cN = cNext[d];
        const base = d * 4;
        const c0 = dayCostGen[base + 0];
        const c1 = dayCostGen[base + 1];
        const c2 = dayCostGen[base + 2];
        const c3 = dayCostGen[base + 3];

        for (let u = 0; u < layer.length; u++) {
          const edges = layer[u];
          let bestVal = Infinity;
          let bestA = 3;
          let bestV = 0;
          for (let ei = 0; ei < edges.length; ei++) {
            const edge = edges[ei];
            const act = edge.act;
            const actCost = act === 0 ? c0 : act === 1 ? c1 : act === 2 ? c2 : c3;
            const edgePen = mode === 2 ? edge.penB : (mode === 0 ? edge.penA : 0);
            const vCost = isLast ? 0 : dpNextGen[edge.nextNode];
            const val = actCost + edgePen + vCost;
            if (val < bestVal) {
              bestVal = val;
              bestA = act;
              bestV = edge.nextNode;
            }
          }
          dpCurGen[u] = bestVal;
          cA[u] = bestA;
          cN[u] = bestV;
        }
        const tmp = dpNextGen;
        dpNextGen = dpCurGen;
        dpCurGen = tmp;
      }

      const path = new Array<number>(daysCount);
      let curNode = 0;
      for (let d = 0; d < daysCount; d++) {
        path[d] = cAct[d][curNode];
        curNode = cNext[d][curNode];
      }
      return path;
    };

    const addGenPath = (dag: GenDag, path: number[], sign: number, mArr: Int32Array, aArr: Int32Array, nArr: Int32Array) => {
      if (dag.isFixedTir) {
        for (let d = 0; d < daysCount; d++) {
          if (path[d] === 0) aArr[d] += sign;
        }
      } else {
        for (let d = 0; d < daysCount; d++) {
          const act = path[d];
          if (act === 0) aArr[d] += sign;
          else if (act === 1) mArr[d] += sign;
          else if (act === 2) nArr[d] += sign;
        }
      }
    };

    const evalGenCov = (
      dags: GenDag[],
      paths: number[][],
      mArr: Int32Array,
      aArr: Int32Array,
      nArr: Int32Array,
      dayWeights: Float64Array | null,
      mode: 0 | 1 | 2,
      tier: ComplianceTier
    ) => {
      let hard = 0;
      let weightedTotal = 0;
      let total = 0;
      let editPenB = 0;
      let diffCount = 0;
      const useSoft = mode === 1;
      for (let d = 0; d < daysCount; d++) {
        const m = mArr[d];
        const a = aArr[d];
        const n = nArr[d];
        if (!isCoverageFeasibleForTier(m, a, n, tier)) hard++;
        const baseP = mode === 2 ? covPenaltyB(m, a, n, tier) : covPenaltyA(m, a, n, 1, useSoft, tier);
        total += baseP;
        weightedTotal += mode === 2 ? baseP : (dayWeights ? covPenaltyA(m, a, n, dayWeights[d], useSoft, tier) : baseP);
      }
      for (let i = 0; i < dags.length; i++) {
        const dg = dags[i];
        const p = paths[i];
        let u = 0;
        for (let d = 0; d < daysCount; d++) {
          const act = p[d];
          if (act !== dg.origActs[d]) diffCount++;
          const edges = dg.layers[d][u];
          for (let ei = 0; ei < edges.length; ei++) {
            if (edges[ei].act === act) {
              const ep = mode === 2 ? edges[ei].penB : (mode === 0 ? edges[ei].penA : 0);
              weightedTotal += ep;
              total += ep;
              editPenB += edges[ei].penB;
              u = edges[ei].nextNode;
              break;
            }
          }
        }
      }
      return { hard, weightedTotal, total, editPenB, diffCount };
    };

    const buildRowsFromGenPaths = (dags: GenDag[], paths: number[][]): Record<string, Record<string, string>> => {
      const candidate = cloneRows(rows);
      for (let i = 0; i < dags.length; i++) {
        const dg = dags[i];
        const p = paths[i];
        for (let d = 0; d < daysCount; d++) {
          const iso = dateColumns[d].isoDate;
          const act = p[d];
          if (dg.isFixedTir) {
            candidate[dg.emp][iso] = dg.origCodes[d];
          } else {
            if (act === 0) candidate[dg.emp][iso] = 'A';
            else if (act === 1) candidate[dg.emp][iso] = 'M';
            else if (act === 2) candidate[dg.emp][iso] = 'N';
            else if (act === 4) candidate[dg.emp][iso] = 'ABS';
            else if (act === 5) candidate[dg.emp][iso] = dg.origCodes[d];
            else candidate[dg.emp][iso] = 'OFF';
          }
        }
      }
      for (const emp of allEmployees) {
        if (fixedMidSet.has(emp)) {
          // Keep original Palade shifts unchanged
          candidate[emp] = { ...rows[emp] };
        }
      }
      return candidate;
    };

    let genRngState = ((((variantSeed + 1) * 0x9e3779b9) ^ 0x517cc1b7) >>> 0);
    const genRand = (): number => {
      genRngState = (genRngState + 0x6d2b79f5) >>> 0;
      let t = Math.imul(genRngState ^ (genRngState >>> 15), 1 | genRngState);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
    const shuffleGen = (arr: number[]): void => {
      for (let i = arr.length - 1; i > 0; i--) {
        const j = Math.floor(genRand() * (i + 1));
        const tmp = arr[i];
        arr[i] = arr[j];
        arr[j] = tmp;
      }
    };
    const isGenPathValid = (dag: GenDag, path: number[]): boolean =>
      isDagPathValid(dag, path, daysCount);
    const solveGenExtremeCount = (dag: GenDag, targetAct: number, sign: 1 | -1): number =>
      solveDagExtremeCount(dag, daysCount, targetAct, sign);

    let bestRepairedRows: Record<string, Record<string, string>> | null = null;
    let bestRepairedValidation = currentValidation;
    let bestRepairedTier: ComplianceTier = 0;
    let bestRepairedRelaxations: RelaxationRecord[] = [];
    let bestRepairedHardTotal = Infinity;
    let bestRepairedDiffCost = Infinity;
    let foundCompliantTier = false;

    const solverEmps = [...rotatingEmployees, ...fixedMaEmps];
    const G = solverEmps.length;

    const hasSufficientGenCapacity = (dags: GenDag[], tier: ComplianceTier): boolean => {
      const fixedMaxA = new Int32Array(daysCount);
      const availM = new Int32Array(daysCount);
      const availA = new Int32Array(daysCount);
      const availN = new Int32Array(daysCount);
      let sumMinRotA = 0, sumMaxRotA = 0, sumMinM = 0, sumMaxM = 0, sumMinN = 0, sumMaxN = 0;
      let sumMaxFixedA = 0;
      let totalRotSlots = 0;

      for (const dg of dags) {
        for (let d = 0; d < daysCount; d++) {
          let hasA = false, hasM = false, hasN = false;
          for (const layerEdges of dg.layers[d]) {
            for (let ei = 0; ei < layerEdges.length; ei++) {
              const a = layerEdges[ei].act;
              if (a === 0) hasA = true;
              else if (!dg.isFixedTir && a === 1) hasM = true;
              else if (!dg.isFixedTir && a === 2) hasN = true;
            }
          }
          if (hasA) {
            availA[d]++;
            if (dg.isFixedTir) fixedMaxA[d]++;
          }
          if (hasM) availM[d]++;
          if (hasN) availN[d]++;
        }

        if (dg.isFixedTir) {
          sumMaxFixedA += solveGenExtremeCount(dg, 0, -1);
        } else {
          const drops = dg.origActs.filter((a) => a === 4 || a === 5).length;
          totalRotSlots += Math.max(0, targetWorkingDays - drops);
          sumMinRotA += solveGenExtremeCount(dg, 0, 1);
          sumMaxRotA += solveGenExtremeCount(dg, 0, -1);
          sumMinM += solveGenExtremeCount(dg, 1, 1);
          sumMaxM += solveGenExtremeCount(dg, 1, -1);
          sumMinN += solveGenExtremeCount(dg, 2, 1);
          sumMaxN += solveGenExtremeCount(dg, 2, -1);
        }
      }

      return canSatisfyDailyCapacityForTier({
        tier,
        daysCount,
        totalRotSlots,
        availM,
        availTotalA: availA,
        availN,
        fixedMaxA,
        sumMinRotA,
        sumMaxRotA,
        sumMaxFixedA,
        sumMinM,
        sumMaxM,
        sumMinN,
        sumMaxN
      });
    };

    const gNoise = new Float64Array(daysCount * 4);
    const gDayWeights = new Float64Array(daysCount);
    const gPool = Array.from({ length: G }, (_, i) => i);
    const s5Helpers: number[] = [];
    const s5Capable: number[] = [];
    const s5Unlocked: number[] = [];
    const s5Locked: number[] = [];

    const buildGenDayActMasks = (dags: GenDag[]): Uint8Array[] =>
      buildDagDayActMasks(dags, daysCount);

    const appendRandomFromBucket = (bucket: number[], out: number[], limit: number): void => {
      let rem = bucket.length;
      while (out.length < limit && rem > 0) {
        const r = Math.floor(genRand() * rem);
        out.push(bucket[r]);
        bucket[r] = bucket[rem - 1];
        rem--;
      }
    };

    const selectSection5RuinGroup = (
      dags: GenDag[],
      genDayActMasks: Uint8Array[],
      paths: number[][],
      focusDay: number,
      k: number,
      curM: Int32Array,
      curA: Int32Array,
      curN: Int32Array,
      tier: ComplianceTier
    ): number[] => {
      if (genRand() < 0.2) {
        shuffleGen(gPool);
        return gPool.slice(0, k);
      }
      s5Helpers.length = 0;
      s5Capable.length = 0;
      s5Unlocked.length = 0;
      s5Locked.length = 0;

      const cfg = getTierConfig(tier);
      const minReqM = cfg.allowCoverageGapMA ? 0 : 1;
      const minReqA = cfg.allowCoverageGapMA ? 1 : 2;
      const maxReqN = cfg.allowNightOver2 ? 3 : 2;
      const nVal = curN[focusDay];
      const mVal = curM[focusDay];
      const aVal = curA[focusDay];
      const needAct = nVal < 2 ? 2 : mVal < minReqM ? 1 : aVal < minReqA ? 0 : (nVal > maxReqN ? -1 : -2);

      for (let e = 0; e < G; e++) {
        const dg = dags[e];
        const mask = genDayActMasks[e][focusDay];
        // Defer rotating employees locked to ABS/H8 (acts 4/5) and Tir when focusDay needs M/N or Tir cannot hold/toggle A
        const isAct45Locked = !dg.isFixedTir && (mask & ~((1 << 4) | (1 << 5))) === 0;
        const isTirExcluded = dg.isFixedTir && (needAct !== 0 || (mask & 1) === 0 || mask === 1);
        if (isAct45Locked || isTirExcluded) {
          s5Locked.push(e);
          continue;
        }
        if (needAct >= 0) {
          if (canDoShift(genDayActMasks, e, focusDay, needAct)) {
            if (paths[e][focusDay] !== needAct) s5Helpers.push(e);
            else s5Capable.push(e);
          } else {
            s5Unlocked.push(e);
          }
        } else if (needAct === -1) {
          if (paths[e][focusDay] === 2 && (mask & ~((1 << 2) | (1 << 4) | (1 << 5))) !== 0) {
            s5Helpers.push(e);
          } else {
            s5Unlocked.push(e);
          }
        } else {
          s5Unlocked.push(e);
        }
      }

      const group: number[] = [];
      appendRandomFromBucket(s5Helpers, group, k);
      appendRandomFromBucket(s5Capable, group, k);
      appendRandomFromBucket(s5Unlocked, group, k);
      appendRandomFromBucket(s5Locked, group, k);
      return group;
    };
    const lockModes: Array<'globalPre' | 'empPre' | 'none'> = hasDistinctEmpAbsStarts
      ? ['globalPre', 'empPre', 'none']
      : ['globalPre', 'none'];

    for (let ti = 0; ti < absTiers.length; ti++) {
      if (foundCompliantTier) break;
      const nowMs = performance.now();
      if (nowMs >= s5DeadlineMs) break;

      const tier = absTiers[ti];
      const cfg = getTierConfig(tier);
      const remBudgetMs = Math.max(0, s5DeadlineMs - nowMs);
      const s5Share = tier === 0 ? 0.25 : tier === 1 ? 0.15 : tier === 2 ? 0.40 : 0.20;
      const remShares = absTiers
        .slice(ti)
        .reduce((sum, t) => sum + (t === 0 ? 0.25 : t === 1 ? 0.15 : t === 2 ? 0.40 : 0.20), 0);
      const tierAllocMs = remShares > 0 ? (remBudgetMs * s5Share) / remShares : remBudgetMs;
      let tierDeadlineMs = Math.min(s5DeadlineMs, nowMs + tierAllocMs);

      let tierBestRows: Record<string, Record<string, string>> | null = null;
      let tierBestValidation: ScheduleValidationReport | null = null;
      let tierBestRelaxations: RelaxationRecord[] = [];
      let tierBestDiffCost = Infinity;

      for (const lockMode of lockModes) {
        if (performance.now() >= tierDeadlineMs) break;
        if (tierBestValidation?.totalViolations === 0) break;

        let fallbackToNoneCount = 0;
        const dags: GenDag[] = solverEmps.map((emp, eIdx) => {
          if (fixedMaEmps.includes(emp)) return buildTirGenDag(emp);
          let dg = buildGeneralDag(emp, eIdx, cfg.minRest, cfg.allowShortBlock2, lockMode);
          if (!dg.feasible && lockMode === 'globalPre') {
            dg = buildGeneralDag(emp, eIdx, cfg.minRest, cfg.allowShortBlock2, 'empPre');
            if (!dg.feasible) {
              dg = buildGeneralDag(emp, eIdx, cfg.minRest, cfg.allowShortBlock2, 'none');
              fallbackToNoneCount++;
            }
          } else if (!dg.feasible && lockMode === 'empPre') {
            dg = buildGeneralDag(emp, eIdx, cfg.minRest, cfg.allowShortBlock2, 'none');
            fallbackToNoneCount++;
          }
          return dg;
        });

        if (dags.some((g) => !g.feasible) || !hasSufficientGenCapacity(dags, tier)) continue;
        if (lockMode !== 'none' && fallbackToNoneCount > Math.floor(G / 2)) continue;

        const genDayActMasks = buildGenDayActMasks(dags);
        const origPathsValid = dags.map((dg) => isGenPathValid(dg, dg.origActs));
        const validOrigCount = origPathsValid.filter(Boolean).length;
        if (cfg.minRest === 2 && validOrigCount < Math.max(2, G - 3)) {
          tierDeadlineMs = Math.min(tierDeadlineMs, performance.now() + 250);
        }
        const numRestarts =
          lockMode === 'none'
            ? cfg.minRest === 2
              ? validOrigCount < Math.max(2, G - 3)
                ? 1
                : 4
              : 8
            : 2;

        for (let restart = 0; restart < numRestarts; restart++) {
          if (performance.now() >= tierDeadlineMs) break;

          gDayWeights.fill(1);
          const curM = new Int32Array(daysCount);
          const curA = new Int32Array(daysCount);
          const curN = new Int32Array(daysCount);
          let paths: number[][] = new Array(G);

          const order = Array.from({ length: G }, (_, i) => i);
          shuffleGen(order);

          const deepOnlyRestart = lockMode === 'none' && restart >= 2;
          const initMode: 0 | 1 = deepOnlyRestart ? 1 : 0;

          if (restart === 0) {
            // Always warm-start from original schedule for all legal employees, solve remaining greedily
            for (let i = 0; i < G; i++) {
              if (origPathsValid[i]) {
                paths[i] = [...dags[i].origActs];
                addGenPath(dags[i], paths[i], 1, curM, curA, curN);
              }
            }
            for (let oi = 0; oi < G; oi++) {
              const i = order[oi];
              if (!origPathsValid[i]) {
                paths[i] = solveGenPath(dags[i], curM, curA, curN, null, null, 0, tier);
                addGenPath(dags[i], paths[i], 1, curM, curA, curN);
              }
            }
          } else if (deepOnlyRestart) {
            // Full fresh Stage-3-style randomized initialization to escape deep origActs basins
            for (let oi = 0; oi < G; oi++) {
              const i = order[oi];
              if (oi < 3) {
                for (let k = 0; k < gNoise.length; k++) gNoise[k] = genRand() * 250;
              } else {
                for (let k = 0; k < gNoise.length; k++) gNoise[k] = genRand() * 8;
              }
              paths[i] = solveGenPath(dags[i], curM, curA, curN, null, gNoise, 1, tier);
              addGenPath(dags[i], paths[i], 1, curM, curA, curN);
            }
          } else {
            // Drop random subset from original warm-start instead of scrambling the whole roster
            for (let oi = 0; oi < G; oi++) {
              const i = order[oi];
              if (origPathsValid[i] && oi >= 2 + (restart % 3)) {
                paths[i] = [...dags[i].origActs];
                addGenPath(dags[i], paths[i], 1, curM, curA, curN);
              } else {
                paths[i] = [];
              }
            }
            for (let oi = 0; oi < G; oi++) {
              const i = order[oi];
              if (paths[i].length === 0) {
                for (let k = 0; k < gNoise.length; k++) gNoise[k] = genRand() * (oi < 3 ? 120 : 12);
                paths[i] = solveGenPath(dags[i], curM, curA, curN, null, gNoise, 0, tier);
                addGenPath(dags[i], paths[i], 1, curM, curA, curN);
              }
            }
          }

          const runGenSweeps = (maxSweeps: number, useWeights: boolean, mode: 0 | 1 | 2) => {
            const wArr = useWeights && mode !== 2 ? gDayWeights : null;
            for (let sweep = 0; sweep < maxSweeps; sweep++) {
              let changed = false;
              for (let oi = 0; oi < G; oi++) {
                const eIdx = order[(oi + sweep) % G];
                const oldP = paths[eIdx];
                const before = evalGenCov(dags, paths, curM, curA, curN, wArr, mode, tier).weightedTotal;
                addGenPath(dags[eIdx], oldP, -1, curM, curA, curN);
                const newP = solveGenPath(dags[eIdx], curM, curA, curN, wArr, null, mode, tier);
                paths[eIdx] = newP;
                addGenPath(dags[eIdx], newP, 1, curM, curA, curN);
                const after = evalGenCov(dags, paths, curM, curA, curN, wArr, mode, tier).weightedTotal;
                if (after < before) {
                  changed = true;
                } else if (after > before) {
                  addGenPath(dags[eIdx], newP, -1, curM, curA, curN);
                  paths[eIdx] = oldP;
                  addGenPath(dags[eIdx], oldP, 1, curM, curA, curN);
                }
              }
              if (!changed) break;
            }
          };

          runGenSweeps(deepOnlyRestart ? 3 : 4, false, initMode);

          let curEval = evalGenCov(dags, paths, curM, curA, curN, gDayWeights, initMode, tier);
          let bestRH = curEval.hard;
          let bestRT = curEval.total;
          let bestRPaths = paths.map((p) => p);

          if (bestRH > 0) {
            const gentleSteps = deepOnlyRestart ? 0 : 350;
            const saSteps = deepOnlyRestart ? 2000 : 1850;
            let temp = deepOnlyRestart ? 1600 : 300;
            let cool = deepOnlyRestart ? 0.994 : 0.985;
            for (let step = 0; step < saSteps && bestRH > 0; step++) {
              if ((step & 15) === 0 && performance.now() >= tierDeadlineMs) break;

              if (!deepOnlyRestart && step === gentleSteps) {
                paths = bestRPaths.map((p) => p);
                curM.fill(0);
                curA.fill(0);
                curN.fill(0);
                for (let i = 0; i < G; i++) addGenPath(dags[i], paths[i], 1, curM, curA, curN);
                gDayWeights.fill(1);
                curEval = evalGenCov(dags, paths, curM, curA, curN, gDayWeights, 1, tier);
                temp = 1600;
                cool = 0.994;
              } else if (step === (deepOnlyRestart ? 1000 : 1150) || (deepOnlyRestart && step === 1550)) {
                temp = Math.max(temp, 750);
              }
              const stepMode: 0 | 1 = step < gentleSteps ? 0 : 1;

              const badDays: number[] = [];
              for (let d = 0; d < daysCount; d++) {
                if (!isCoverageFeasibleForTier(curM[d], curA[d], curN[d], tier)) badDays.push(d);
              }
              if (badDays.length === 0) break;

              if (step > 0 && step % 35 === 0 && step !== gentleSteps) {
                for (const bd of badDays) gDayWeights[bd] = Math.min(8, gDayWeights[bd] + 0.5);
                if (step % 210 === 0) {
                  for (let d = 0; d < daysCount; d++) gDayWeights[d] = 1 + (gDayWeights[d] - 1) * 0.5;
                }
                curEval = evalGenCov(dags, paths, curM, curA, curN, gDayWeights, stepMode, tier);
              }

              const focusDay = badDays[Math.floor(genRand() * badDays.length)];
              const k = stepMode === 0 ? Math.min(G, 2) : Math.min(G, step % 2 === 0 ? 3 : 2);
              const ruinGroup = selectSection5RuinGroup(dags, genDayActMasks, paths, focusDay, k, curM, curA, curN, tier);

              const savedSub = ruinGroup.map((idx) => paths[idx]);
              for (let gi = 0; gi < ruinGroup.length; gi++) {
                addGenPath(dags[ruinGroup[gi]], paths[ruinGroup[gi]], -1, curM, curA, curN);
              }

              const noiseAmp = stepMode === 0 ? Math.min(40, temp * 0.15) : Math.min(350, temp * 0.3);
              for (let gi = 0; gi < ruinGroup.length; gi++) {
                const eIdx = ruinGroup[gi];
                if (gi === 0 && genRand() < 0.5) {
                  for (let nk = 0; nk < gNoise.length; nk++) gNoise[nk] = genRand() * (stepMode === 0 ? 2 : 6);
                  for (let d = Math.max(0, focusDay - 5); d <= Math.min(daysCount - 1, focusDay + 5); d++) {
                    for (let a = 0; a < 4; a++) gNoise[d * 4 + a] += (genRand() - 0.5) * noiseAmp * 3;
                  }
                } else {
                  for (let nk = 0; nk < gNoise.length; nk++) gNoise[nk] = genRand() * (2 + noiseAmp * 0.25);
                }
                const newP = solveGenPath(dags[eIdx], curM, curA, curN, gDayWeights, gNoise, stepMode, tier);
                paths[eIdx] = newP;
                addGenPath(dags[eIdx], newP, 1, curM, curA, curN);
              }

              let nextEval = evalGenCov(dags, paths, curM, curA, curN, gDayWeights, stepMode, tier);
              const improvedRec = nextEval.hard < bestRH;
              if (improvedRec) {
                runGenSweeps(2, true, stepMode);
                nextEval = evalGenCov(dags, paths, curM, curA, curN, gDayWeights, stepMode, tier);
              }
              const delta = nextEval.weightedTotal - curEval.weightedTotal;
              const accept = improvedRec || delta <= 0 || genRand() < Math.exp(-delta / Math.max(1, temp));
              if (accept) {
                curEval = nextEval;
                if (curEval.hard < bestRH || (curEval.hard === bestRH && curEval.total < bestRT)) {
                  bestRH = curEval.hard;
                  bestRT = curEval.total;
                  bestRPaths = paths.map((p) => p);
                  if (bestRH === 0) break;
                }
              } else {
                for (let gi = 0; gi < ruinGroup.length; gi++) {
                  const eIdx = ruinGroup[gi];
                  addGenPath(dags[eIdx], paths[eIdx], -1, curM, curA, curN);
                  paths[eIdx] = savedSub[gi];
                  addGenPath(dags[eIdx], savedSub[gi], 1, curM, curA, curN);
                }
              }
              temp *= cool;
            }
          }

          paths = bestRPaths;
          curM.fill(0);
          curA.fill(0);
          curN.fill(0);
          for (let i = 0; i < G; i++) addGenPath(dags[i], paths[i], 1, curM, curA, curN);

          // Phase B: Strict feasibility-preserving deterministic minimal-diff polishing
          runGenSweeps(6, false, 2);

          // Pairwise 2-employee zero-noise minimization across modified employees
          for (let pass = 0; pass < 2; pass++) {
            if (performance.now() >= tierDeadlineMs) break;
            let pairImproved = false;
            const changedEmps: number[] = [];
            for (let i = 0; i < G; i++) {
              for (let d = 0; d < daysCount; d++) {
                if (paths[i][d] !== dags[i].origActs[d]) {
                  changedEmps.push(i);
                  break;
                }
              }
            }
            for (let aIdx = 0; aIdx < changedEmps.length; aIdx++) {
              if (performance.now() >= tierDeadlineMs) break;
              for (let bIdx = aIdx + 1; bIdx < changedEmps.length; bIdx++) {
                if (performance.now() >= tierDeadlineMs) break;
                const i1 = changedEmps[aIdx];
                const i2 = changedEmps[bIdx];
                const baseScore = evalGenCov(dags, paths, curM, curA, curN, null, 2, tier);
                const old1 = paths[i1];
                const old2 = paths[i2];

                for (const [first, second] of [[i1, i2], [i2, i1]]) {
                  const s1 = paths[first];
                  const s2 = paths[second];
                  addGenPath(dags[first], s1, -1, curM, curA, curN);
                  addGenPath(dags[second], s2, -1, curM, curA, curN);
                  const n1 = solveGenPath(dags[first], curM, curA, curN, null, null, 2, tier);
                  paths[first] = n1;
                  addGenPath(dags[first], n1, 1, curM, curA, curN);
                  const n2 = solveGenPath(dags[second], curM, curA, curN, null, null, 2, tier);
                  paths[second] = n2;
                  addGenPath(dags[second], n2, 1, curM, curA, curN);

                  const nextScore = evalGenCov(dags, paths, curM, curA, curN, null, 2, tier);
                  if (nextScore.hard <= baseScore.hard && nextScore.weightedTotal < baseScore.weightedTotal) {
                    pairImproved = true;
                    break;
                  } else {
                    addGenPath(dags[first], n1, -1, curM, curA, curN);
                    addGenPath(dags[second], n2, -1, curM, curA, curN);
                    paths[i1] = old1;
                    paths[i2] = old2;
                    addGenPath(dags[i1], old1, 1, curM, curA, curN);
                    addGenPath(dags[i2], old2, 1, curM, curA, curN);
                  }
                }
              }
            }
            if (!pairImproved) break;
            runGenSweeps(4, false, 2);
          }

          const finalEval = evalGenCov(dags, paths, curM, curA, curN, null, 2, tier);
          const candidateRows = buildRowsFromGenPaths(dags, paths);
          const candRelax: RelaxationRecord[] = [];
          const candidateRep = validateScheduleCompliance(
            candidateRows,
            dateColumns,
            priorRoster,
            targetYear,
            targetMonth,
            targetWorkingDays,
            [],
            { tier, relaxations: candRelax }
          );

          const isBetterInTier =
            !tierBestValidation ||
            candidateRep.totalViolations < tierBestValidation.totalViolations ||
            (candidateRep.totalViolations === tierBestValidation.totalViolations &&
              finalEval.editPenB < tierBestDiffCost);

          if (isBetterInTier) {
            tierBestRows = candidateRows;
            tierBestValidation = candidateRep;
            tierBestRelaxations = candRelax;
            tierBestDiffCost = finalEval.editPenB;
          }

          const hardRep = tier === 3
            ? candidateRep
            : validateScheduleCompliance(
                candidateRows,
                dateColumns,
                priorRoster,
                targetYear,
                targetMonth,
                targetWorkingDays,
                [],
                { tier: 3 }
              );
          const isBetterFallback =
            !bestRepairedRows ||
            hardRep.totalViolations < bestRepairedHardTotal ||
            (hardRep.totalViolations === bestRepairedHardTotal &&
              candidateRep.totalViolations < bestRepairedValidation.totalViolations) ||
            (hardRep.totalViolations === bestRepairedHardTotal &&
              candidateRep.totalViolations === bestRepairedValidation.totalViolations &&
              finalEval.editPenB < bestRepairedDiffCost);

          if (isBetterFallback) {
            bestRepairedRows = candidateRows;
            bestRepairedValidation = candidateRep;
            bestRepairedTier = tier;
            bestRepairedRelaxations = candRelax;
            bestRepairedHardTotal = hardRep.totalViolations;
            bestRepairedDiffCost = finalEval.editPenB;
          }

          if (candidateRep.totalViolations === 0) break;
        }
      }

      if (tierBestRows && tierBestValidation && tierBestValidation.totalViolations === 0) {
        bestRepairedRows = tierBestRows;
        bestRepairedValidation = tierBestValidation;
        bestRepairedTier = tier;
        bestRepairedRelaxations = tierBestRelaxations;
        foundCompliantTier = true;
        break;
      }
    }

    if (bestRepairedRows && !foundCompliantTier) {
      let resolved = false;
      for (const t of absTiers) {
        const tRelax: RelaxationRecord[] = [];
        const tRep = validateScheduleCompliance(
          bestRepairedRows,
          dateColumns,
          priorRoster,
          targetYear,
          targetMonth,
          targetWorkingDays,
          [],
          { tier: t, relaxations: tRelax }
        );
        if (tRep.totalViolations === 0) {
          bestRepairedTier = t;
          bestRepairedRelaxations = tRelax;
          bestRepairedValidation = tRep;
          foundCompliantTier = true;
          resolved = true;
          break;
        }
      }
      if (!resolved) {
        const t3Relax: RelaxationRecord[] = [];
        const t3Rep = validateScheduleCompliance(
          bestRepairedRows,
          dateColumns,
          priorRoster,
          targetYear,
          targetMonth,
          targetWorkingDays,
          [],
          { tier: 3, relaxations: t3Relax }
        );
        bestRepairedTier = 3;
        bestRepairedRelaxations = t3Relax;
        bestRepairedValidation = t3Rep;
      }
    }

    if (bestRepairedRows && (foundCompliantTier || bestRepairedValidation.totalViolations < currentValidation.totalViolations)) {
      for (const emp of allEmployees) {
        workingRows[emp] = { ...bestRepairedRows[emp] };
      }
      finalTier = bestRepairedTier;
      finalRelaxations = bestRepairedRelaxations;
      currentValidation = bestRepairedValidation;
      adjustments.length = 0;
      unresolvedDeficits.length = 0;

      for (const col of dateColumns) {
        const iso = col.isoDate;
        const absOnDay = allEmployees.filter((e) => shiftKey(rows[e]?.[iso] ?? 'OFF') === 'ABS').join(', ');
        for (const emp of allEmployees) {
          const origC = shiftKey(rows[emp]?.[iso] ?? 'OFF');
          const newC = shiftKey(workingRows[emp]?.[iso] ?? 'OFF');
          if (origC !== newC) {
            adjustments.push({
              isoDate: iso,
              employee: emp,
              fromShift: origC,
              toShift: newC,
              reason: absOnDay
                ? `Rebalanced schedule to legally cover sick colleague (${absOnDay}) while maintaining ${targetWorkingDays}-day monthly norm and rest rules.`
                : `Rebalanced shift to preserve legal ${targetWorkingDays}-day monthly norm, shift sequence, and rest rules.`,
              absColleague: absOnDay || 'Rebalancing'
            });
          }
        }

        // Recompute unresolvedDeficits cleanly if any coverage gaps remain
        const rotM = rotatingEmployees.filter((e) => shiftKey(workingRows[e]?.[iso] ?? 'OFF') === 'M').length;
        const totalA = allEmployees.filter((e) => shiftKey(workingRows[e]?.[iso] ?? 'OFF') === 'A').length;
        const totalN = allEmployees.filter((e) => shiftKey(workingRows[e]?.[iso] ?? 'OFF') === 'N').length;
        if (rotM < 1) {
          unresolvedDeficits.push({
            isoDate: iso,
            shift: 'M',
            needed: 1 - rotM,
            available: rotM,
            reason: `Rotating Morning below 1 on ${iso} after Tier ${finalTier} repair.`
          });
        }
        if (totalA < 2) {
          unresolvedDeficits.push({
            isoDate: iso,
            shift: 'A',
            needed: 2 - totalA,
            available: totalA,
            reason: `Afternoon below 2 on ${iso} after Tier ${finalTier} repair.`
          });
        }
        if (totalN < 2) {
          unresolvedDeficits.push({
            isoDate: iso,
            shift: 'N',
            needed: 2 - totalN,
            available: totalN,
            reason: `Night below 2 on ${iso} after Tier ${finalTier} repair.`
          });
        }
      }
    }
  }

  return {
    hasDeficits: adjustments.length > 0 || unresolvedDeficits.length > 0,
    totalDeficits: adjustments.length + unresolvedDeficits.length,
    adjustments,
    unresolvedDeficits,
    updatedRows: workingRows,
    validation: currentValidation,
    tier: finalTier,
    relaxations: finalRelaxations
  };
}
