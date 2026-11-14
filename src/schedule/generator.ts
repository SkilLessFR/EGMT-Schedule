import type {
  RosterData,
  ComplianceTier,
  RelaxationRecord,
  ScheduleBuilderOptions,
  ScheduleValidationReport,
  GeneratedScheduleResult
} from './types';
import {
  FIXED_MID_EMPLOYEES,
  FIXED_WEEKDAY_MA_EMPLOYEES,
  FIXED_STAFF_EXEMPT_OFF_DATES,
  ROMANIAN_HOLIDAYS,
  DAG_NEXT_SHIFT_ACT,
  DAG_LEAVE_CODES,
  shiftKey
} from './constants';
import {
  getIsoWeekNumber,
  countLegalWorkingDaysInMonth,
  isFixedStaffOffDate,
  buildMonthDateColumns
} from './dateUtils';
import {
  getActiveTailBlock,
  getActiveTailRest,
  getEmployeeHistoryTail,
  extractRawHistoryTail
} from './historyUtils';
import { minCostMatching } from './hungarian';
import {
  encodeStage3DagState,
  canDoShift,
  buildDagDayActMasks,
  isDagPathValid,
  solveDagExtremeCount
} from './dagUtils';
import { evaluateOffsetCost } from './costEvaluator';
import { equalizeExactWorkingDays } from './equalizer';
import { validateScheduleCompliance } from './validator';
import {
  getTierConfig,
  isCoverageFeasibleForTier,
  tierCovPenalty,
  canSatisfyDailyCapacityForTier
} from '../relaxation';

/**
 * Core Schedule Generation Function
 */
export function generateBalancedSchedule(
  roster: RosterData | null,
  targetYear: number,
  targetMonth: number,
  options: ScheduleBuilderOptions = {}
): GeneratedScheduleResult {
  const startTimeMs = performance.now();
  const timeBudgetMs = options.timeBudgetMs ?? 8000;
  const minTier: ComplianceTier = options.minTier ?? 0;
  const maxTier: ComplianceTier = options.maxTier ?? 3;
  const activeTiers: ComplianceTier[] = ([0, 1, 2, 3] as ComplianceTier[]).filter(
    (t) => t >= minTier && t <= maxTier
  );
  if (activeTiers.length === 0) activeTiers.push(maxTier);
  const safetyMarginMs = Math.min(150, Math.floor(timeBudgetMs * 0.1));
  const totalDeadlineMs = startTimeMs + Math.max(50, timeBudgetMs - safetyMarginMs);

  const seed = options.seed ?? Math.floor(Math.random() * 1000000);
  const targetWorkingDays =
    options.targetWorkingDays ?? countLegalWorkingDaysInMonth(targetMonth, targetYear);
  const dateColumns = buildMonthDateColumns(targetYear, targetMonth);
  const daysCount = dateColumns.length;
  const targetStartIso = `${targetYear}-${String(targetMonth + 1).padStart(2, '0')}-01`;

  // Employees list
  const allEmployees: string[] = (roster?.employees && roster.employees.length > 0)
    ? [...roster.employees]
    : [
        'Palade Alexandru-Ionut',
        'Tir George Cristian',
        'Urzica Ion-Alexandru',
        'Debona Andrei-Laurentiu',
        'Nica Bogdan-Alexandru',
        'Corneci Stefan-Cristian',
        'Toma Mihai',
        'Gheorghe Adrian Marinescu',
        'Olteanu Ionut-Lucian',
        'Stoian Alexandru-Gabriel',
        'Prisecaru Alexandru-Ionut'
      ];

  const excludedSet = new Set(options.excludedEmployees ?? []);
  const activeAllEmployees = allEmployees.filter((emp) => !excludedSet.has(emp));

  const fixedMidEmployees = FIXED_MID_EMPLOYEES.filter((name) => activeAllEmployees.includes(name));
  const fixedWeekdayMAEmployees = FIXED_WEEKDAY_MA_EMPLOYEES.filter((name) => activeAllEmployees.includes(name));
  const rotatingEmployees = activeAllEmployees.filter((name) => !fixedMidEmployees.includes(name) && !fixedWeekdayMAEmployees.includes(name));

  // Balanced 1:1:1 templates with 4-day blocks and 2-day rests across 18-day coprime cycles
  // Strict cyclic direction: A -> M -> N -> A (Afternoon -> Morning -> Night -> Afternoon)
  // Standard density (12/18 = 66.7% work -> ~20-21 work days in 30-31d)
  const STANDARD_TEMPLATES = [
    // Template 0: A -> M -> N -> A
    [
      'A','A','A','A',
      'OFF','OFF',
      'M','M','M','M',
      'OFF','OFF',
      'N','N','N','N',
      'OFF','OFF'
    ],
    // Template 1: M -> N -> A -> M
    [
      'M','M','M','M',
      'OFF','OFF',
      'N','N','N','N',
      'OFF','OFF',
      'A','A','A','A',
      'OFF','OFF'
    ],
    // Template 2: N -> A -> M -> N
    [
      'N','N','N','N',
      'OFF','OFF',
      'A','A','A','A',
      'OFF','OFF',
      'M','M','M','M',
      'OFF','OFF'
    ]
  ];

  // Lower-density 11/18 (61.1% work -> ~18-19 work days in 30-31d) templates preserving A/N=2 and M>=1 coverage
  const LOW_DENSITY_TEMPLATES = [
    [
      'A','A','A','A',
      'OFF','OFF',
      'M','M','M',
      'OFF','OFF','OFF',
      'N','N','N','N',
      'OFF','OFF'
    ],
    [
      'M','M','M',
      'OFF','OFF','OFF',
      'N','N','N','N',
      'OFF','OFF',
      'A','A','A','A',
      'OFF','OFF'
    ],
    [
      'N','N','N','N',
      'OFF','OFF',
      'A','A','A','A',
      'OFF','OFF',
      'M','M','M',
      'OFF','OFF','OFF'
    ]
  ];

  // Larger rotating teams (>= 10) need 3-day N blocks per 18d cycle to avoid 3/2 Night overstaffing
  const LARGE_TEAM_3N_TEMPLATES = [
    [
      'A','A','A','A',
      'OFF','OFF',
      'M','M','M','M',
      'OFF','OFF',
      'N','N','N',
      'OFF','OFF','OFF'
    ],
    [
      'M','M','M','M',
      'OFF','OFF',
      'N','N','N',
      'OFF','OFF','OFF',
      'A','A','A','A',
      'OFF','OFF'
    ],
    [
      'N','N','N',
      'OFF','OFF','OFF',
      'A','A','A','A',
      'OFF','OFF',
      'M','M','M','M',
      'OFF','OFF'
    ]
  ];

  const isLowWorkRatio = (targetWorkingDays / daysCount) < 0.644;
  const BALANCED_TEMPLATES = rotatingEmployees.length >= 10
    ? [...LARGE_TEAM_3N_TEMPLATES, ...(isLowWorkRatio ? LOW_DENSITY_TEMPLATES : []), ...STANDARD_TEMPLATES]
    : isLowWorkRatio
      ? [...LOW_DENSITY_TEMPLATES, ...STANDARD_TEMPLATES]
      : STANDARD_TEMPLATES;

  // Pre-configured slot sets ensuring optimal coverage across variable team sizes
  const VALID_SLOT_SETS_6 = [
    [0, 1, 4, 6, 10, 14],
    [0, 1, 4, 7, 10, 14],
    [0, 1, 4, 7, 11, 14],
    [0, 1, 4, 7, 11, 15],
    [0, 1, 4, 8, 10, 14],
    [0, 1, 4, 8, 11, 14],
    [0, 1, 4, 8, 11, 15],
    [0, 1, 4, 8, 12, 14],
    [0, 1, 4, 8, 12, 15],
    [0, 1, 5, 6, 10, 14]
  ];

  const VALID_SLOT_SETS_7 = [
    [0, 1, 4, 5, 8, 10, 14],
    [0, 1, 4, 5, 8, 11, 14],
    [0, 1, 4, 5, 8, 11, 15],
    [0, 1, 4, 5, 8, 12, 14],
    [0, 1, 4, 5, 8, 12, 15],
    [0, 1, 4, 5, 9, 10, 14],
    [0, 1, 4, 5, 9, 11, 14],
    [0, 1, 4, 5, 9, 11, 15],
    [0, 1, 4, 5, 9, 12, 14],
    [0, 1, 4, 5, 9, 12, 15]
  ];

  const VALID_SLOT_SETS_8 = [
    [0, 1, 4, 5, 8, 9, 12, 14],
    [0, 1, 4, 5, 8, 9, 12, 15],
    [0, 1, 4, 5, 8, 9, 13, 14],
    [0, 1, 4, 5, 8, 9, 13, 15],
    [0, 1, 4, 5, 8, 10, 12, 14],
    [0, 1, 4, 5, 8, 10, 12, 15],
    [0, 1, 4, 5, 8, 10, 13, 14],
    [0, 1, 4, 5, 8, 10, 13, 15],
    [0, 1, 4, 5, 8, 10, 14, 15],
    [0, 1, 4, 5, 8, 11, 12, 15]
  ];

  const BASE_SLOT_SETS_9 = [
    [0, 2, 4, 6, 8, 10, 12, 14, 16],
    [1, 3, 5, 7, 9, 11, 13, 15, 17],
    [0, 2, 4, 5, 8, 9, 12, 14, 16],
    [0, 2, 4, 5, 8, 10, 12, 14, 16],
    [1, 2, 4, 5, 8, 10, 13, 14, 16],
    [1, 2, 4, 5, 8, 10, 13, 15, 16],
    [1, 2, 4, 5, 9, 10, 13, 14, 16],
    [1, 2, 4, 5, 9, 10, 14, 15, 16],
    [1, 2, 4, 5, 8, 9, 13, 14, 16],
    [1, 2, 4, 5, 8, 9, 14, 15, 16],
    [0, 1, 3, 4, 7, 9, 12, 13, 15],
    [0, 1, 3, 4, 8, 9, 12, 13, 15]
  ];

  const VALID_SLOT_SETS_9: number[][] = [];
  for (const s of BASE_SLOT_SETS_9) {
    for (let shift = 0; shift < 18; shift += 1) {
      const rotated = s.map((x) => (x + shift) % 18).sort((a, b) => a - b);
      const key = rotated.join(',');
      if (!VALID_SLOT_SETS_9.some((existing) => existing.join(',') === key)) {
        VALID_SLOT_SETS_9.push(rotated);
      }
    }
  }

  const VALID_SLOT_SETS_10 = [
    [0, 1, 2, 4, 6, 7, 10, 11, 14, 15],
    [0, 1, 2, 4, 6, 8, 10, 11, 14, 15],
    [0, 1, 2, 4, 6, 8, 10, 12, 14, 15],
    [0, 1, 2, 4, 6, 8, 10, 12, 14, 16],
    [0, 1, 2, 5, 6, 7, 10, 11, 14, 15],
    [0, 1, 2, 5, 6, 8, 10, 11, 14, 15],
    [0, 1, 2, 5, 6, 8, 10, 12, 14, 15],
    [0, 1, 2, 5, 6, 8, 10, 12, 14, 16],
    [0, 1, 2, 5, 6, 9, 10, 11, 14, 15],
    [0, 1, 2, 5, 6, 9, 10, 12, 14, 15]
  ];

  const VALID_SLOT_SETS_11 = [
    [0, 1, 2, 3, 4, 6, 7, 10, 11, 14, 15],
    [0, 1, 2, 4, 6, 8, 10, 11, 13, 14, 17],
    [0, 1, 3, 4, 6, 7, 10, 11, 12, 14, 15],
    [0, 1, 3, 5, 6, 9, 10, 11, 14, 15, 16],
    [0, 1, 4, 5, 6, 9, 10, 11, 13, 15, 17],
    [0, 2, 3, 4, 6, 7, 9, 10, 13, 14, 17],
    [0, 2, 3, 5, 7, 8, 10, 12, 13, 14, 16],
    [0, 2, 4, 5, 6, 8, 10, 11, 14, 15, 17],
    [0, 2, 4, 6, 7, 9, 10, 11, 12, 15, 16],
    [0, 3, 4, 5, 8, 9, 10, 12, 14, 15, 17],
    [1, 2, 3, 4, 6, 8, 10, 12, 14, 15, 17],
    [1, 2, 3, 6, 7, 10, 11, 14, 15, 16, 17],
    [1, 2, 4, 6, 7, 10, 11, 13, 14, 15, 16],
    [1, 2, 5, 6, 9, 10, 13, 14, 15, 16, 17],
    [1, 3, 4, 7, 8, 11, 12, 13, 15, 16, 17],
    [2, 3, 4, 6, 7, 8, 11, 12, 13, 16, 17]
  ];

  const VALID_SLOT_SETS_12 = [
    [0, 1, 2, 4, 5, 6, 8, 9, 10, 12, 14, 15],
    [0, 1, 2, 4, 5, 6, 8, 9, 10, 12, 14, 16],
    [0, 1, 2, 4, 5, 6, 8, 9, 10, 13, 14, 15],
    [0, 1, 2, 4, 5, 6, 8, 9, 10, 13, 14, 16],
    [0, 1, 2, 4, 5, 6, 8, 9, 11, 12, 14, 15],
    [0, 1, 2, 4, 5, 6, 8, 9, 11, 12, 14, 16],
    [0, 1, 2, 4, 5, 6, 8, 9, 11, 12, 15, 16],
    [0, 1, 2, 4, 5, 6, 8, 9, 11, 13, 14, 15],
    [0, 1, 2, 4, 5, 6, 8, 9, 11, 13, 14, 16],
    [0, 1, 2, 4, 5, 6, 8, 9, 11, 13, 15, 16]
  ];

  function getValidSlotSets(count: number): number[][] {
    if (count <= 6) return VALID_SLOT_SETS_6;
    if (count === 7) return VALID_SLOT_SETS_7;
    if (count === 8) return VALID_SLOT_SETS_8;
    if (count === 9) return VALID_SLOT_SETS_9;
    if (count === 10) return VALID_SLOT_SETS_10;
    if (count === 11) return VALID_SLOT_SETS_11;
    if (count === 12) return VALID_SLOT_SETS_12;
    const step = 18 / count;
    return [
      Array.from({ length: count }, (_, i) => Math.floor(i * step) % 18),
      Array.from({ length: count }, (_, i) => (Math.floor(i * step) + 1) % 18)
    ];
  }

  const VALID_SLOT_SETS = getValidSlotSets(rotatingEmployees.length);

  let bestResultRows: Record<string, Record<string, string>> = {};
  let bestReport: ScheduleValidationReport | null = null;
  let bestRelaxations: RelaxationRecord[] = [];
  let bestScoreValue = Infinity;
  let bestHandoffSafe = false;
  let finalSeed = seed;

  const nextMonthMinRotA = [1, 2, 3].map((dayOfMonth) => {
    const d = new Date(targetYear, targetMonth + 1, dayOfMonth);
    const y = d.getFullYear();
    const m = d.getMonth() + 1;
    const iso = `${y}-${String(m).padStart(2, '0')}-${String(dayOfMonth).padStart(2, '0')}`;
    const dow = d.getDay();
    const isOff = dow === 0 || dow === 6 || ROMANIAN_HOLIDAYS.has(iso) || FIXED_STAFF_EXEMPT_OFF_DATES.has(iso);
    const fixedA = (!isOff && fixedWeekdayMAEmployees.length > 0 && getIsoWeekNumber(d) % 2 === 0)
      ? fixedWeekdayMAEmployees.length
      : 0;
    return Math.max(0, 2 - fixedA);
  });

  const isMonthEndHandoffSafe = (rows: Record<string, Record<string, string>>): boolean => {
    if (rotatingEmployees.length < 6) return true;
    const isRest = (s: string) => s === 'OFF' || s === 'H8' || s === 'HUP' || s === 'ABS' || s === 'AD';
    const tails = rotatingEmployees.map((emp) => {
      const len = Math.min(6, dateColumns.length);
      const tail: string[] = [];
      for (let idx = dateColumns.length - len; idx < dateColumns.length; idx++) {
        const s = shiftKey(rows[emp]?.[dateColumns[idx].isoDate] ?? 'OFF');
        tail.push(isRest(s) ? 'OFF' : s);
      }
      let tailWorkStreak = 0;
      let tailOffStreak = 0;
      let tailWorkShift: string | null = null;
      let lastWorkShift: string | null = null;
      for (let i = tail.length - 1; i >= 0; i--) {
        if (tail[i] !== 'OFF') {
          if (!lastWorkShift) lastWorkShift = tail[i];
          if (tailOffStreak === 0) {
            tailWorkStreak++;
            tailWorkShift = tail[i];
          } else {
            break;
          }
        } else {
          if (tailWorkStreak > 0) break;
          tailOffStreak++;
        }
      }
      return { tailWorkShift, tailWorkStreak, tailOffStreak, lastWorkShift };
    });

    for (const k of [1, 2, 3]) {
      let eligN = 0;
      let eligA = 0;
      let eligM = 0;
      for (const t of tails) {
        const canCover = (target: 'M' | 'A' | 'N', prevShift: 'M' | 'A' | 'N') => {
          if (t.tailWorkShift === target && t.tailWorkStreak <= 5 - k) return true;
          if (t.tailWorkShift === null && t.lastWorkShift === prevShift && t.tailOffStreak >= (k === 1 ? 2 : 1)) return true;
          if (k === 3 && t.tailWorkShift === prevShift && t.tailWorkStreak >= 3) return true;
          return false;
        };
        if (canCover('N', 'M')) eligN++;
        if (canCover('A', 'N')) eligA++;
        if (canCover('M', 'A')) eligM++;
      }
      if (eligN < 2 || eligA < nextMonthMinRotA[k - 1] || eligM < 1) {
        return false;
      }
    }
    return true;
  };

  interface CandidateConfig {
    template: string[];
    slots: number[];
    matching: number[];
    cost: number;
    attemptSeed: number;
    signature: string;
  }

  const templateRotations = BALANCED_TEMPLATES.map((tmpl) =>
    tmpl.map((_, off) => tmpl.map((_, i) => tmpl[(off + i) % tmpl.length][0]).join(''))
  );
  const seenAssignments = new Set<string>();

  const rotatingHistories: Record<string, string[]> = {};
  const rotatingRawHistories: Record<string, string[]> = {};
  const rotatingAnchors: Record<string, string | null> = {};
  rotatingEmployees.forEach((emp) => {
    const hist = getEmployeeHistoryTail(roster, emp, targetYear, targetMonth, 6);
    rotatingHistories[emp] = hist;
    rotatingRawHistories[emp] = extractRawHistoryTail(roster, emp, targetYear, targetMonth, 6);
    rotatingAnchors[emp] = hist.length > 0 ? roster?.rows?.[emp]?.[targetStartIso] ?? null : null;
  });

  const SEED_OFFSETS = [0, 1, 3, 5, 7, 10, 100];

  const fixedAfternoonCoverage = dateColumns.map(({ date, isoDate }) => {
    const dow = date.getDay();
    const isOff = dow === 0 || dow === 6 || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
    if (isOff || fixedWeekdayMAEmployees.length === 0) return 0;
    return getIsoWeekNumber(date) % 2 === 0 ? fixedWeekdayMAEmployees.length : 0;
  });

  const getSlotImbalancePenalty = (template: string[], slots: number[]): number => {
    const tLen = template.length;
    let penalty = 0;
    for (let d = 0; d < daysCount; d++) {
      let m = 0;
      let a = fixedAfternoonCoverage[d];
      let n = 0;
      for (const offset of slots) {
        const s = template[(offset + d) % tLen];
        if (s === 'M') m++;
        else if (s === 'A') a++;
        else if (s === 'N') n++;
      }
      if (m < 1) penalty += 1 - m;
      if (a < 2) penalty += 2 - a;
      if (n !== 2) penalty += Math.abs(n - 2);
    }
    return penalty * 40;
  };

  const scoreReport = (rep: ScheduleValidationReport, handoffSafe: boolean) => {
    const legalHard =
      rep.streakViolationCount +
      rep.nightRestViolationCount +
      rep.excessiveRestViolationCount +
      rep.mixedBlockViolationCount +
      rep.workDaysMismatchCount +
      (rep.fixedStaffViolationCount ?? 0);
    return (
      legalHard * 10000 +
      rep.nightOverstaffedDaysCount * 200 +
      rep.understaffedDaysCount * 100 +
      (handoffSafe ? 0 : 500)
    );
  };

  const stage2DeadlineMs = startTimeMs + Math.min(250, Math.floor(timeBudgetMs * 0.15));
  let stage2FoundTier2Safe = false;

  for (const sOffset of SEED_OFFSETS) {
    if (bestReport && (stage2FoundTier2Safe || performance.now() >= stage2DeadlineMs || performance.now() >= totalDeadlineMs)) {
      break;
    }
    const effSeed = seed + sOffset;
    const allCandidates: CandidateConfig[] = [];

    // Stage 1: Fast boundary screening across all template and slot configurations (< 10ms)
    for (let t = 0; t < BALANCED_TEMPLATES.length; t++) {
      for (let s = 0; s < VALID_SLOT_SETS.length; s++) {
        const templateIdx = (effSeed + t) % BALANCED_TEMPLATES.length;
        const slotSetIdx = (Math.floor(effSeed / BALANCED_TEMPLATES.length) + s) % VALID_SLOT_SETS.length;
        const template = BALANCED_TEMPLATES[templateIdx];
        const slots = VALID_SLOT_SETS[slotSetIdx];
        const currentAttemptSeed = effSeed + (t * VALID_SLOT_SETS.length + s) * 101;

        const costMatrix: number[][] = [];
        for (let eIdx = 0; eIdx < rotatingEmployees.length; eIdx++) {
          const emp = rotatingEmployees[eIdx];
          const history = rotatingHistories[emp];
          const anchor = rotatingAnchors[emp];
          costMatrix[eIdx] = [];
          for (let sIdx = 0; sIdx < slots.length; sIdx++) {
            costMatrix[eIdx][sIdx] = evaluateOffsetCost(
              history,
              template,
              slots[sIdx],
              emp,
              daysCount,
              anchor,
              currentAttemptSeed,
              targetWorkingDays
            );
          }
        }

        const matching = minCostMatching(costMatrix);
        let bestCost = getSlotImbalancePenalty(template, slots);
        for (let eIdx = 0; eIdx < rotatingEmployees.length; eIdx++) {
          bestCost += costMatrix[eIdx][matching[eIdx]];
        }

        const signature = matching
          .map((sIdx) => templateRotations[templateIdx][slots[sIdx]])
          .join('|');

        allCandidates.push({
          template,
          slots,
          matching,
          cost: bestCost,
          attemptSeed: currentAttemptSeed,
          signature
        });
      }
    }

    // Sort candidates by boundary cost ascending (strictly legal combinations first)
    allCandidates.sort((a, b) => a.cost - b.cost);

    // Stage 2: Equalize and validate distinct candidate configurations (early exit on minTier compliant + handoffSafe)
    const configLimit = sOffset === 0 ? 48 : 24;
    let tested = 0;

    for (const config of allCandidates) {
      if (bestReport && (stage2FoundTier2Safe || performance.now() >= stage2DeadlineMs || performance.now() >= totalDeadlineMs)) {
        break;
      }
      if (seenAssignments.has(config.signature)) continue;
      seenAssignments.add(config.signature);
      if (++tested > configLimit) break;
      const candidateRows: Record<string, Record<string, string>> = {};
      allEmployees.forEach((emp) => {
        candidateRows[emp] = {};
        if (excludedSet.has(emp)) {
          dateColumns.forEach(({ isoDate }) => {
            candidateRows[emp][isoDate] = 'OFF';
          });
        }
      });

      fixedMidEmployees.forEach((emp) => {
        dateColumns.forEach(({ date, isoDate }) => {
          const dow = date.getDay();
          const isWeekend = dow === 0 || dow === 6;
          const isOff = isWeekend || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
          if (isOff) {
            candidateRows[emp][isoDate] = 'OFF';
          } else {
            candidateRows[emp][isoDate] = 'MID';
          }
        });
      });

      fixedWeekdayMAEmployees.forEach((emp) => {
        dateColumns.forEach(({ date, isoDate }) => {
          const dow = date.getDay();
          const isWeekend = dow === 0 || dow === 6;
          const isOff = isWeekend || ROMANIAN_HOLIDAYS.has(isoDate) || FIXED_STAFF_EXEMPT_OFF_DATES.has(isoDate);
          if (isOff) {
            candidateRows[emp][isoDate] = 'OFF';
          } else {
            const weekNum = getIsoWeekNumber(date);
            candidateRows[emp][isoDate] = (weekNum % 2 === 1) ? 'M' : 'A';
          }
        });
      });

      const tLen = config.template.length;
      rotatingEmployees.forEach((emp, eIdx) => {
        const slotOffset = config.slots[config.matching[eIdx]];
        dateColumns.forEach(({ index, isoDate }) => {
          candidateRows[emp][isoDate] = config.template[(slotOffset + index) % tLen];
        });
      });

      equalizeExactWorkingDays(
        candidateRows,
        dateColumns,
        targetWorkingDays,
        roster,
        targetYear,
        targetMonth,
        options.excludedEmployees ?? []
      );
      const candRelax: RelaxationRecord[] = [];
      const testReport = validateScheduleCompliance(
        candidateRows,
        dateColumns,
        roster,
        targetYear,
        targetMonth,
        targetWorkingDays,
        options.excludedEmployees ?? [],
        { tier: minTier, relaxations: candRelax }
      );

      const curHandoffSafe = isMonthEndHandoffSafe(candidateRows);
      const tier2Rep =
        minTier === 2
          ? testReport
          : validateScheduleCompliance(
              candidateRows,
              dateColumns,
              roster,
              targetYear,
              targetMonth,
              targetWorkingDays,
              options.excludedEmployees ?? [],
              { tier: 2 }
            );
      let hardCheckViol = 0;
      if (!tier2Rep.isFullyCompliant && minTier < 3) {
        const hardRep =
          minTier === 3
            ? testReport
            : validateScheduleCompliance(
                candidateRows,
                dateColumns,
                roster,
                targetYear,
                targetMonth,
                targetWorkingDays,
                options.excludedEmployees ?? [],
                { tier: 3 }
              );
        hardCheckViol = hardRep.totalViolations;
      }
      const curScore =
        hardCheckViol * 10000000 +
        (tier2Rep.isFullyCompliant ? 0 : 100000) +
        scoreReport(tier2Rep, curHandoffSafe) * 10 +
        scoreReport(testReport, curHandoffSafe);

      if (!bestReport || curScore < bestScoreValue) {
        bestResultRows = candidateRows;
        bestReport = testReport;
        bestRelaxations = candRelax;
        bestScoreValue = curScore;
        bestHandoffSafe = curHandoffSafe;
        finalSeed = config.attemptSeed;
        if (testReport.isFullyCompliant && curHandoffSafe) {
          break;
        }
        if (minTier < 2 && tier2Rep.isFullyCompliant && curHandoffSafe) {
          stage2FoundTier2Safe = true;
          break;
        }
      }
    }

    if ((bestReport?.isFullyCompliant && bestHandoffSafe) || stage2FoundTier2Safe) {
      break;
    }
  }

  if (bestReport?.isFullyCompliant && bestHandoffSafe) {
    return {
      month: targetMonth,
      year: targetYear,
      daysCount,
      targetWorkingDays,
      dateColumns,
      rows: bestResultRows,
      validation: bestReport,
      seed: finalSeed,
      excludedEmployees: options.excludedEmployees ?? [],
      tier: minTier,
      relaxations: bestRelaxations,
      needsManualReview: false,
      elapsedMs: Math.round(performance.now() - startTimeMs)
    };
  }

  const stage2BestRows = bestReport ? bestResultRows : null;
  const stage2BestSeed = finalSeed;

  const tryStage2AtTier = (t: ComplianceTier): GeneratedScheduleResult | null => {
    if (!stage2BestRows) return null;
    const s2Relax: RelaxationRecord[] = [];
    const s2Rep = validateScheduleCompliance(
      stage2BestRows,
      dateColumns,
      roster,
      targetYear,
      targetMonth,
      targetWorkingDays,
      options.excludedEmployees ?? [],
      { tier: t, relaxations: s2Relax }
    );
    if (s2Rep.isFullyCompliant) {
      return {
        month: targetMonth,
        year: targetYear,
        daysCount,
        targetWorkingDays,
        dateColumns,
        rows: stage2BestRows,
        validation: s2Rep,
        seed: stage2BestSeed,
        excludedEmployees: options.excludedEmployees ?? [],
        tier: t,
        relaxations: s2Relax,
        needsManualReview: false,
        elapsedMs: Math.round(performance.now() - startTimeMs)
      };
    }
    return null;
  };

  interface FallbackAttempt {
    rows: Record<string, Record<string, string>>;
    seed: number;
    handoffSafe: boolean;
  }
  const fallbackAttempts: FallbackAttempt[] = [];
  if (stage2BestRows) {
    fallbackAttempts.push({
      rows: stage2BestRows,
      seed: stage2BestSeed,
      handoffSafe: bestHandoffSafe
    });
  }

  // Stage 3: Exact Legal-Trajectory DP + Large Neighborhood Search across active tiers
  if (rotatingEmployees.length > 0 && targetWorkingDays > 0 && targetWorkingDays <= daysCount) {
    const ACT_SHIFT = ['A', 'M', 'N', 'OFF'] as const;
    const NEXT_SH = DAG_NEXT_SHIFT_ACT;
    const shToIdx = (s: string) => (s === 'A' ? 0 : s === 'M' ? 1 : s === 'N' ? 2 : 3);
    const numStatesPerDay = (targetWorkingDays + 1) * 4 * 13 * 2;
    const encState = encodeStage3DagState;

    interface DagEdge {
      act: number;
      nextNode: number;
      pen: number;
    }
    interface EmpDag {
      layers: DagEdge[][][]; // layers[d][nodeIdx] -> outgoing edges to layers[d+1]
      feasible: boolean;
      maxNodes: number;
      cAct: Int8Array[];
      cNext: Int32Array[];
    }

    const buildEmpDag = (
      history: string[],
      minNonNightRest: number,
      allowShortBlock2: boolean,
      adjLeaveBefore: boolean
    ): EmpDag => {
      const hasHistory = history.length > 0;
      const activeTailWork = getActiveTailBlock(history);
      const activeTailRest = activeTailWork.length === 0 ? getActiveTailRest(history) : [];
      let lastHistWork: string | null = null;
      for (let i = history.length - 1; i >= 0; i--) {
        if (history[i] === 'A' || history[i] === 'M' || history[i] === 'N') {
          lastHistWork = history[i];
          break;
        }
      }

      // Compute outgoing transitions from virtual start node (at d = 0 -> state at d = 1)
      const startEdges: Array<{ act: number; nextState: number; pen: number }> = [];
      if (activeTailWork.length > 0) {
        const curShift = activeTailWork[activeTailWork.length - 1];
        const curIdx = shToIdx(curShift);
        const totalWorkTail = activeTailWork.length;
        let sameLen = 0;
        for (let i = activeTailWork.length - 1; i >= 0 && activeTailWork[i] === curShift; i--) sameLen++;
        if (totalWorkTail < 5 && targetWorkingDays >= 1) {
          const nextMl = adjLeaveBefore ? 5 + totalWorkTail : totalWorkTail;
          startEdges.push({ act: curIdx, nextState: encState(1, curIdx, nextMl, 0), pen: 0 });
        }
        if (sameLen >= 3 || adjLeaveBefore) {
          startEdges.push({ act: 3, nextState: encState(0, curIdx, 10, 0), pen: 0 });
        }
      } else if (hasHistory) {
        const restLen = activeTailRest.length;
        const lastIdx = lastHistWork ? shToIdx(lastHistWork) : 3;
        if (restLen < 3) {
          startEdges.push({ act: 3, nextState: encState(0, lastIdx, 10 + restLen, 0), pen: 0 });
        }
        const reqRest = lastIdx === 2 ? 2 : (adjLeaveBefore ? 1 : minNonNightRest);
        if (targetWorkingDays >= 1 && (lastIdx === 3 || restLen >= reqRest)) {
          const pen = lastIdx !== 3 && lastIdx !== 2 && restLen === 1 && !adjLeaveBefore ? 15 : 0;
          const firstMl = adjLeaveBefore ? 5 : 0;
          if (lastIdx === 3) {
            for (const ns of [0, 1, 2]) {
              startEdges.push({ act: ns, nextState: encState(1, ns, firstMl, 0), pen: 0 });
            }
          } else {
            const ns = NEXT_SH[lastIdx];
            startEdges.push({ act: ns, nextState: encState(1, ns, firstMl, 0), pen });
            if (adjLeaveBefore) {
              startEdges.push({ act: lastIdx, nextState: encState(1, lastIdx, firstMl, 0), pen: 0 });
            }
          }
        }
      } else {
        startEdges.push({ act: 3, nextState: encState(0, 3, 10, 0), pen: 0 });
        if (targetWorkingDays >= 1) {
          for (const ns of [0, 1, 2]) {
            startEdges.push({ act: ns, nextState: encState(1, ns, 5, 0), pen: 0 });
          }
        }
      }

      if (startEdges.length === 0) {
        return { layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
      }

      // Forward reachability sets: activeStates[d] for d = 1..daysCount
      const activeStates: number[][] = Array.from({ length: daysCount + 1 }, () => []);
      const seen = new Uint8Array(numStatesPerDay);
      for (const e of startEdges) {
        if (!seen[e.nextState]) {
          seen[e.nextState] = 1;
          activeStates[1].push(e.nextState);
        }
      }

      const stepTransitions = (state: number): Array<{ act: number; nextState: number; pen: number }> => {
        const sb = state & 1;
        const t0 = state >> 1;
        const ml = t0 % 13;
        const tmp = (t0 - ml) / 13;
        const sh = tmp % 4;
        const w = (tmp - sh) / 4;
        const out: Array<{ act: number; nextState: number; pen: number }> = [];

        if (ml < 10) {
          const isUnb = ml >= 5;
          const len = isUnb ? ml - 4 : ml + 1;
          if (len < 5 && w + 1 <= targetWorkingDays) {
            out.push({ act: sh, nextState: encState(w + 1, sh, ml + 1, sb), pen: 0 });
          }
          if (len >= 3 || isUnb) {
            out.push({ act: 3, nextState: encState(w, sh, 10, sb), pen: 0 });
          } else if (allowShortBlock2 && len === 2 && sb === 0) {
            out.push({ act: 3, nextState: encState(w, sh, 10, 1), pen: 25 });
          }
        } else {
          const len = ml - 9;
          if (len < 3) {
            out.push({ act: 3, nextState: encState(w, sh, ml + 1, sb), pen: 0 });
          }
          if (w + 1 <= targetWorkingDays) {
            if (sh === 3) {
              for (const ns of [0, 1, 2]) {
                out.push({ act: ns, nextState: encState(w + 1, ns, 0, sb), pen: 0 });
              }
            } else if (len >= (sh === 2 ? 2 : minNonNightRest)) {
              const ns = NEXT_SH[sh];
              const pen = sh !== 2 && len === 1 ? 15 : 0;
              out.push({ act: ns, nextState: encState(w + 1, ns, 0, sb), pen });
            }
          }
        }
        return out;
      };

      for (let d = 1; d < daysCount; d++) {
        seen.fill(0);
        for (const st of activeStates[d]) {
          for (const tr of stepTransitions(st)) {
            if (!seen[tr.nextState]) {
              seen[tr.nextState] = 1;
              activeStates[d + 1].push(tr.nextState);
            }
          }
        }
      }

      // Backward reachability filter from terminal states (w === targetWorkingDays at d = daysCount)
      const reachNext = new Int32Array(numStatesPerDay).fill(-1);
      let termIdx = 0;
      for (const st of activeStates[daysCount]) {
        const sb = st & 1;
        const t0 = st >> 1;
        const ml = t0 % 13;
        const tmp = (t0 - ml) / 13;
        const sh = tmp % 4;
        const w = (tmp - sh) / 4;
        if (w === targetWorkingDays && (sb === 0 || allowShortBlock2)) {
          reachNext[st] = termIdx++;
        }
      }
      if (termIdx === 0) {
        return { layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
      }

      const layers: DagEdge[][][] = new Array(daysCount);
      const reachCur = new Int32Array(numStatesPerDay);
      for (let d = daysCount - 1; d >= 1; d--) {
        reachCur.fill(-1);
        const dayLayer: DagEdge[][] = [];
        for (const st of activeStates[d]) {
          const nodeEdges: DagEdge[] = [];
          for (const tr of stepTransitions(st)) {
            const nxtIdx = reachNext[tr.nextState];
            if (nxtIdx !== -1) {
              nodeEdges.push({ act: tr.act, nextNode: nxtIdx, pen: tr.pen });
            }
          }
          if (nodeEdges.length > 0) {
            reachCur[st] = dayLayer.length;
            dayLayer.push(nodeEdges);
          }
        }
        if (dayLayer.length === 0) {
          return { layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
        }
        layers[d] = dayLayer;
        reachNext.set(reachCur);
      }

      const rootEdges: DagEdge[] = [];
      for (const e of startEdges) {
        const nxtIdx = reachNext[e.nextState];
        if (nxtIdx !== -1) {
          rootEdges.push({ act: e.act, nextNode: nxtIdx, pen: e.pen });
        }
      }
      if (rootEdges.length === 0) {
        return { layers: [], feasible: false, maxNodes: 0, cAct: [], cNext: [] };
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
      return { layers, feasible: true, maxNodes, cAct, cNext };
    };

    const dayCostScratch = new Float64Array(daysCount * 4);
    let dpNextBuf = new Float64Array(512);
    let dpCurBuf = new Float64Array(512);

    const solveBestPath = (
      dag: EmpDag,
      otherM: Int32Array,
      otherA: Int32Array,
      otherN: Int32Array,
      dayWeights: Float64Array | null,
      noise: Float64Array | null,
      tier: ComplianceTier
    ): number[] => {
      if (dag.maxNodes > dpNextBuf.length) {
        dpNextBuf = new Float64Array(dag.maxNodes + 64);
        dpCurBuf = new Float64Array(dag.maxNodes + 64);
      }
      for (let d = 0; d < daysCount; d++) {
        const m0 = otherM[d];
        const a0 = otherA[d] + fixedAfternoonCoverage[d];
        const n0 = otherN[d];
        const dw = dayWeights ? dayWeights[d] : 1;
        const base = d * 4;
        dayCostScratch[base + 0] = tierCovPenalty(m0, a0 + 1, n0, dw, tier) + (noise ? noise[base + 0] : 0);
        dayCostScratch[base + 1] = tierCovPenalty(m0 + 1, a0, n0, dw, tier) + (noise ? noise[base + 1] : 0);
        dayCostScratch[base + 2] = tierCovPenalty(m0, a0, n0 + 1, dw, tier) + (noise ? noise[base + 2] : 0);
        dayCostScratch[base + 3] = tierCovPenalty(m0, a0, n0, dw, tier) + (noise ? noise[base + 3] : 0);
      }

      const { layers, cAct: dpChoicesAct, cNext: dpChoicesNext } = dag;

      for (let d = daysCount - 1; d >= 0; d--) {
        const layer = layers[d];
        const nNodes = layer.length;
        const cAct = dpChoicesAct[d];
        const cNext = dpChoicesNext[d];
        const base = d * 4;
        const isLast = d === daysCount - 1;
        const c0 = dayCostScratch[base + 0];
        const c1 = dayCostScratch[base + 1];
        const c2 = dayCostScratch[base + 2];
        const c3 = dayCostScratch[base + 3];

        for (let u = 0; u < nNodes; u++) {
          const edges = layer[u];
          const elen = edges.length;
          let bestVal = Infinity;
          let bestA = 3;
          let bestV = 0;
          for (let ei = 0; ei < elen; ei++) {
            const edge = edges[ei];
            const act = edge.act;
            const actCost = act === 0 ? c0 : act === 1 ? c1 : act === 2 ? c2 : c3;
            const vCost = isLast ? 0 : dpNextBuf[edge.nextNode];
            const val = actCost + edge.pen + vCost;
            if (val < bestVal) {
              bestVal = val;
              bestA = act;
              bestV = edge.nextNode;
            }
          }
          dpCurBuf[u] = bestVal;
          cAct[u] = bestA;
          cNext[u] = bestV;
        }
        const tmp = dpNextBuf;
        dpNextBuf = dpCurBuf;
        dpCurBuf = tmp;
      }

      const path = new Array<number>(daysCount);
      let curNode = 0;
      for (let d = 0; d < daysCount; d++) {
        path[d] = dpChoicesAct[d][curNode];
        curNode = dpChoicesNext[d][curNode];
      }
      return path;
    };

    const buildRowsFromPaths = (paths: number[][]): Record<string, Record<string, string>> => {
      const candidateRows: Record<string, Record<string, string>> = {};
      allEmployees.forEach((emp) => {
        candidateRows[emp] = {};
        if (excludedSet.has(emp)) {
          dateColumns.forEach(({ isoDate }) => {
            candidateRows[emp][isoDate] = 'OFF';
          });
        }
      });
      fixedMidEmployees.forEach((emp) => {
        dateColumns.forEach(({ date, isoDate }) => {
          candidateRows[emp][isoDate] = isFixedStaffOffDate(date, isoDate) ? 'OFF' : 'MID';
        });
      });
      fixedWeekdayMAEmployees.forEach((emp) => {
        dateColumns.forEach(({ date, isoDate }) => {
          candidateRows[emp][isoDate] = isFixedStaffOffDate(date, isoDate)
            ? 'OFF'
            : getIsoWeekNumber(date) % 2 === 1
              ? 'M'
              : 'A';
        });
      });
      rotatingEmployees.forEach((emp, eIdx) => {
        const path = paths[eIdx];
        dateColumns.forEach(({ index, isoDate }) => {
          candidateRows[emp][isoDate] = ACT_SHIFT[path[index]];
        });
      });
      return candidateRows;
    };

    let rngState = (seed ^ 0x9e3779b9) >>> 0;
    const nextRand = (): number => {
      rngState = (rngState + 0x6d2b79f5) >>> 0;
      let t = Math.imul(rngState ^ (rngState >>> 15), 1 | rngState);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
    const shuffleStage3 = (arr: number[]): void => {
      for (let i = arr.length - 1; i > 0; i--) {
        const j = Math.floor(nextRand() * (i + 1));
        const tmp = arr[i];
        arr[i] = arr[j];
        arr[j] = tmp;
      }
    };
    const isStage3PathValid = (dag: EmpDag, path: number[]): boolean =>
      isDagPathValid(dag, path, daysCount);
    const solveStage3ExtremeCount = (dag: EmpDag, targetAct: number, sign: 1 | -1): number =>
      solveDagExtremeCount(dag, daysCount, targetAct, sign);

    const evalCovWeighted = (
      mArr: Int32Array,
      aArr: Int32Array,
      nArr: Int32Array,
      dayWeights: Float64Array | null,
      tier: ComplianceTier
    ) => {
      let hard = 0;
      let total = 0;
      let weightedTotal = 0;
      for (let d = 0; d < daysCount; d++) {
        const m = mArr[d];
        const a = aArr[d] + fixedAfternoonCoverage[d];
        const n = nArr[d];
        if (!isCoverageFeasibleForTier(m, a, n, tier)) hard++;
        const baseP = tierCovPenalty(m, a, n, 1, tier);
        total += baseP;
        weightedTotal += dayWeights ? tierCovPenalty(m, a, n, dayWeights[d], tier) : baseP;
      }
      return { hard, total, weightedTotal };
    };

    const addPath = (path: number[], sign: number, mArr: Int32Array, aArr: Int32Array, nArr: Int32Array) => {
      for (let d = 0; d < daysCount; d++) {
        const act = path[d];
        if (act === 0) aArr[d] += sign;
        else if (act === 1) mArr[d] += sign;
        else if (act === 2) nArr[d] += sign;
      }
    };

    const buildDayActMasks = (dags: EmpDag[]): Uint8Array[] =>
      buildDagDayActMasks(dags, daysCount);

    const hasSufficientDagCapacity = (
      dags: EmpDag[],
      dayActMasks: Uint8Array[],
      tier: ComplianceTier
    ): boolean => {
      const availM = new Int32Array(daysCount);
      const availTotalA = new Int32Array(daysCount);
      const availN = new Int32Array(daysCount);
      for (let d = 0; d < daysCount; d++) {
        let mCnt = 0;
        let aCnt = fixedAfternoonCoverage[d];
        let nCnt = 0;
        for (let i = 0; i < dayActMasks.length; i++) {
          const mk = dayActMasks[i][d];
          if (mk & 1) aCnt++;
          if (mk & 2) mCnt++;
          if (mk & 4) nCnt++;
        }
        availM[d] = mCnt;
        availTotalA[d] = aCnt;
        availN[d] = nCnt;
      }
      let sumMinA = 0, sumMaxA = 0, sumMinM = 0, sumMaxM = 0, sumMinN = 0, sumMaxN = 0;
      for (let i = 0; i < dags.length; i++) {
        sumMinA += solveStage3ExtremeCount(dags[i], 0, 1);
        sumMaxA += solveStage3ExtremeCount(dags[i], 0, -1);
        sumMinM += solveStage3ExtremeCount(dags[i], 1, 1);
        sumMaxM += solveStage3ExtremeCount(dags[i], 1, -1);
        sumMinN += solveStage3ExtremeCount(dags[i], 2, 1);
        sumMaxN += solveStage3ExtremeCount(dags[i], 2, -1);
      }
      return canSatisfyDailyCapacityForTier({
        tier,
        daysCount,
        totalRotSlots: dags.length * targetWorkingDays,
        availM,
        availTotalA,
        availN,
        fixedMaxA: fixedAfternoonCoverage,
        sumMinRotA: sumMinA,
        sumMaxRotA: sumMaxA,
        sumMinM,
        sumMaxM,
        sumMinN,
        sumMaxN
      });
    };

    const selectStage3RuinGroup = (
      focusDay: number,
      k: number,
      curM: Int32Array,
      curA: Int32Array,
      curN: Int32Array,
      paths: number[][],
      dayActMasks: Uint8Array[],
      tier: ComplianceTier
    ): number[] => {
      shuffleStage3(empPool);
      if (nextRand() < 0.2) {
        return empPool.slice(0, k);
      }
      const cfg = getTierConfig(tier);
      const minReqM = cfg.allowCoverageGapMA ? 0 : 1;
      const minReqA = cfg.allowCoverageGapMA ? 1 : 2;
      const maxReqN = cfg.allowNightOver2 ? 3 : 2;
      const needAct =
        curN[focusDay] < 2
          ? 2
          : curM[focusDay] < minReqM
            ? 1
            : curA[focusDay] + fixedAfternoonCoverage[focusDay] < minReqA
              ? 0
              : curN[focusDay] > maxReqN
                ? -1
                : -2;
      const helpers: number[] = [];
      const capable: number[] = [];
      const others: number[] = [];
      for (let idx = 0; idx < empPool.length; idx++) {
        const e = empPool[idx];
        if (needAct >= 0) {
          if (canDoShift(dayActMasks, e, focusDay, needAct)) {
            if (paths[e][focusDay] !== needAct) helpers.push(e);
            else capable.push(e);
          } else {
            others.push(e);
          }
        } else if (needAct === -1 && paths[e][focusDay] === 2 && (dayActMasks[e][focusDay] & ~4) !== 0) {
          helpers.push(e);
        } else {
          others.push(e);
        }
      }
      if (helpers.length > 0) {
        const partners = [...helpers.slice(1), ...capable, ...others];
        shuffleStage3(partners);
        return [helpers[0], ...partners].slice(0, k);
      }
      if (capable.length > 0) {
        return [...capable, ...others].slice(0, k);
      }
      return empPool.slice(0, k);
    };

    const R = rotatingEmployees.length;
    const LEAVE_CODES = DAG_LEAVE_CODES;
    const stage3EmpData = rotatingEmployees.map((emp) => {
      const rawHist = rotatingRawHistories[emp] ?? [];
      const sanitizedHist = rawHist.map((s) =>
        s === 'HUP' || s === 'H8' || s === 'ABS' || s === 'AD' || s === 'CM' ? 'OFF' : s
      );
      const activeTailWork = getActiveTailBlock(sanitizedHist);
      const idxBefore = rawHist.length - 1 - activeTailWork.length;
      const adjLeaveBefore = idxBefore >= 0 && LEAVE_CODES.has(rawHist[idxBefore]);
      return { sanitizedHist, adjLeaveBefore };
    });

    const priorBestPaths = stage2BestRows
      ? rotatingEmployees.map((emp) =>
          dateColumns.map(({ isoDate }) => shToIdx(shiftKey(stage2BestRows[emp]?.[isoDate] ?? 'OFF')))
        )
      : [];
    const dayWeights = new Float64Array(daysCount);
    const noise = new Float64Array(daysCount * 4);
    const empPool = Array.from({ length: R }, (_, i) => i);

    let hasDagFallback = false;
    const addQuickDagFallback = (dags: EmpDag[], tier: ComplianceTier, fallbackSeed: number) => {
      if (hasDagFallback || dags.some((g) => !g.feasible)) return;
      const curM = new Int32Array(daysCount);
      const curA = new Int32Array(daysCount);
      const curN = new Int32Array(daysCount);
      const paths: number[][] = new Array(R);
      const order = Array.from({ length: R }, (_, i) => i);
      shuffleStage3(order);
      for (let oi = 0; oi < R; oi++) {
        const eIdx = order[oi];
        for (let k = 0; k < noise.length; k++) noise[k] = nextRand() * 6;
        const p = solveBestPath(dags[eIdx], curM, curA, curN, null, noise, tier);
        paths[eIdx] = p;
        addPath(p, 1, curM, curA, curN);
      }
      for (let sweep = 0; sweep < 2; sweep++) {
        let changed = false;
        for (let oi = 0; oi < R; oi++) {
          const eIdx = order[(oi + sweep) % R];
          const oldP = paths[eIdx];
          const before = evalCovWeighted(curM, curA, curN, null, tier).weightedTotal;
          addPath(oldP, -1, curM, curA, curN);
          const newP = solveBestPath(dags[eIdx], curM, curA, curN, null, null, tier);
          addPath(newP, 1, curM, curA, curN);
          const after = evalCovWeighted(curM, curA, curN, null, tier).weightedTotal;
          if (after < before) {
            paths[eIdx] = newP;
            changed = true;
          } else if (after > before) {
            addPath(newP, -1, curM, curA, curN);
            addPath(oldP, 1, curM, curA, curN);
          } else {
            paths[eIdx] = newP;
          }
        }
        if (!changed) break;
      }
      const candidateRows = buildRowsFromPaths(paths);
      fallbackAttempts.push({
        rows: candidateRows,
        seed: fallbackSeed,
        handoffSafe: isMonthEndHandoffSafe(candidateRows)
      });
      hasDagFallback = true;
    };

    for (let ti = 0; ti < activeTiers.length; ti++) {
      const tier = activeTiers[ti];
      const cfg = getTierConfig(tier);
      const dags = stage3EmpData.map(({ sanitizedHist, adjLeaveBefore }) =>
        buildEmpDag(sanitizedHist, cfg.minRest, cfg.allowShortBlock2, cfg.allowShortBlock2 && adjLeaveBefore)
      );
      if (dags.some((g) => !g.feasible)) {
        const s2Hit = tryStage2AtTier(tier);
        if (s2Hit) return s2Hit;
        continue;
      }

      const nowMs = performance.now();
      if (nowMs >= totalDeadlineMs) {
        addQuickDagFallback(dags, tier, seed);
        for (let tj = ti; tj < activeTiers.length; tj++) {
          const s2Hit = tryStage2AtTier(activeTiers[tj]);
          if (s2Hit) return s2Hit;
        }
        break;
      }

      const remBudgetMs = Math.max(0, totalDeadlineMs - nowMs);
      const remShares = activeTiers
        .slice(ti)
        .reduce((sum, t) => sum + getTierConfig(t).budgetShare, 0);
      const tierAllocMs = remShares > 0 ? (remBudgetMs * cfg.budgetShare) / remShares : remBudgetMs;
      const tierDeadlineMs = Math.min(totalDeadlineMs, nowMs + tierAllocMs);

      const dayActMasks = buildDayActMasks(dags);
      if (!hasSufficientDagCapacity(dags, dayActMasks, tier)) {
        if (ti === activeTiers.length - 1 || !hasDagFallback) {
          addQuickDagFallback(dags, tier, seed);
        }
        const s2Hit = tryStage2AtTier(tier);
        if (s2Hit) return s2Hit;
        continue;
      }

      const canSeedStage2 =
        priorBestPaths.length === R &&
        priorBestPaths.every((p, idx) => isStage3PathValid(dags[idx], p));
      const numRestarts = tier === 0 ? 6 : tier === 1 ? 3 : 10;
      let compliantRestarts = 0;
      let tierBestWinner: {
        rows: Record<string, Record<string, string>>;
        report: ScheduleValidationReport;
        relaxations: RelaxationRecord[];
        handoffSafe: boolean;
        seed: number;
        score: number;
      } | null = null;

      if (stage2BestRows) {
        const s2Relax: RelaxationRecord[] = [];
        const s2Rep = validateScheduleCompliance(
          stage2BestRows,
          dateColumns,
          roster,
          targetYear,
          targetMonth,
          targetWorkingDays,
          options.excludedEmployees ?? [],
          { tier, relaxations: s2Relax }
        );
        if (s2Rep.isFullyCompliant) {
          const s2Safe = isMonthEndHandoffSafe(stage2BestRows);
          tierBestWinner = {
            rows: stage2BestRows,
            report: s2Rep,
            relaxations: s2Relax,
            handoffSafe: s2Safe,
            seed: stage2BestSeed,
            score: (s2Safe ? 0 : 10000) + s2Relax.length * 10
          };
          if (s2Safe) {
            return {
              month: targetMonth,
              year: targetYear,
              daysCount,
              targetWorkingDays,
              dateColumns,
              rows: tierBestWinner.rows,
              validation: tierBestWinner.report,
              seed: tierBestWinner.seed,
              excludedEmployees: options.excludedEmployees ?? [],
              tier,
              relaxations: tierBestWinner.relaxations,
              needsManualReview: false,
              elapsedMs: Math.round(performance.now() - startTimeMs)
            };
          }
        }
      }

      for (let restart = 0; restart < numRestarts; restart++) {
        if (performance.now() >= tierDeadlineMs) break;
        if (tierBestWinner?.handoffSafe) break;
        if (tierBestWinner && (compliantRestarts >= 3 || restart >= 4)) break;

        dayWeights.fill(1);
        let paths: number[][] = new Array(R);
        const curM = new Int32Array(daysCount);
        const curA = new Int32Array(daysCount);
        const curN = new Int32Array(daysCount);
        const order = Array.from({ length: R }, (_, i) => i);
        shuffleStage3(order);

        if (restart === 0 && canSeedStage2) {
          for (let eIdx = 0; eIdx < R; eIdx++) {
            paths[eIdx] = priorBestPaths[eIdx];
            addPath(paths[eIdx], 1, curM, curA, curN);
          }
        } else {
          // Greedy randomized initialization
          for (let oi = 0; oi < R; oi++) {
            const eIdx = order[oi];
            if (restart > 0 && oi < 3) {
              for (let k = 0; k < noise.length; k++) noise[k] = nextRand() * 250;
            } else {
              for (let k = 0; k < noise.length; k++) noise[k] = nextRand() * (restart === 0 ? 10 : 8);
            }
            const p = solveBestPath(dags[eIdx], curM, curA, curN, null, noise, tier);
            paths[eIdx] = p;
            addPath(p, 1, curM, curA, curN);
          }
        }

        // Coordinate-descent best-response sweeps
        const runSweeps = (maxSweeps: number, useWeights: boolean) => {
          const wArr = useWeights ? dayWeights : null;
          for (let sweep = 0; sweep < maxSweeps; sweep++) {
            let changed = false;
            for (let oi = 0; oi < R; oi++) {
              const eIdx = order[(oi + sweep) % R];
              const oldP = paths[eIdx];
              const before = evalCovWeighted(curM, curA, curN, wArr, tier).weightedTotal;
              addPath(oldP, -1, curM, curA, curN);
              const newP = solveBestPath(dags[eIdx], curM, curA, curN, wArr, null, tier);
              addPath(newP, 1, curM, curA, curN);
              const after = evalCovWeighted(curM, curA, curN, wArr, tier).weightedTotal;
              if (after < before) {
                paths[eIdx] = newP;
                changed = true;
              } else if (after > before) {
                addPath(newP, -1, curM, curA, curN);
                addPath(oldP, 1, curM, curA, curN);
              } else {
                paths[eIdx] = newP;
              }
            }
            if (!changed) break;
          }
        };

        if (!(restart === 0 && canSeedStage2)) {
          runSweeps(3, false);
        }

        // Multi-employee Ruin-and-Recreate Simulated Annealing + Guided Local Search dayWeights
        let curEval = evalCovWeighted(curM, curA, curN, dayWeights, tier);
        let bestRestartHard = curEval.hard;
        let bestRestartTotal = curEval.total;
        let bestRestartPaths = paths.map((p) => p);

        let temp = 1600;
        const cool = 0.994;
        const saSteps = curEval.hard > 0 ? 1500 : 0;
        for (let step = 0; step < saSteps && bestRestartHard > 0; step++) {
          if ((step & 15) === 0 && performance.now() >= tierDeadlineMs) break;

          const badDays: number[] = [];
          for (let d = 0; d < daysCount; d++) {
            if (!isCoverageFeasibleForTier(curM[d], curA[d] + fixedAfternoonCoverage[d], curN[d], tier)) {
              badDays.push(d);
            }
          }
          if (badDays.length === 0) break;

          // Guided Local Search: boost weights on stubborn bad days every 35 steps
          if (step > 0 && step % 35 === 0) {
            for (const bd of badDays) {
              dayWeights[bd] = Math.min(8, dayWeights[bd] + 0.5);
            }
            if (step % 210 === 0) {
              for (let d = 0; d < daysCount; d++) dayWeights[d] = 1 + (dayWeights[d] - 1) * 0.5;
            }
            curEval = evalCovWeighted(curM, curA, curN, dayWeights, tier);
          }

          const focusDay = badDays[Math.floor(nextRand() * badDays.length)];
          const k = Math.min(R, step % 3 === 0 ? 3 : 2);
          const ruinGroup = selectStage3RuinGroup(focusDay, k, curM, curA, curN, paths, dayActMasks, tier);

          const savedSub: number[][] = ruinGroup.map((idx) => paths[idx]);

          for (let gi = 0; gi < ruinGroup.length; gi++) {
            addPath(paths[ruinGroup[gi]], -1, curM, curA, curN);
          }

          const noiseAmp = Math.min(350, temp * 0.32);
          for (let gi = 0; gi < ruinGroup.length; gi++) {
            const eIdx = ruinGroup[gi];
            if (gi === 0 && nextRand() < 0.5) {
              for (let nk = 0; nk < noise.length; nk++) noise[nk] = nextRand() * 6;
              for (let d = Math.max(0, focusDay - 5); d <= Math.min(daysCount - 1, focusDay + 5); d++) {
                for (let a = 0; a < 4; a++) noise[d * 4 + a] += (nextRand() - 0.5) * noiseAmp * 3;
              }
            } else {
              for (let nk = 0; nk < noise.length; nk++) noise[nk] = nextRand() * (4 + noiseAmp * 0.25);
            }
            const newP = solveBestPath(dags[eIdx], curM, curA, curN, dayWeights, noise, tier);
            paths[eIdx] = newP;
            addPath(newP, 1, curM, curA, curN);
          }

          let nextEval = evalCovWeighted(curM, curA, curN, dayWeights, tier);
          const improvedRecord = nextEval.hard < bestRestartHard;
          if (improvedRecord) {
            runSweeps(2, true);
            nextEval = evalCovWeighted(curM, curA, curN, dayWeights, tier);
          }

          const delta = nextEval.weightedTotal - curEval.weightedTotal;
          const accept =
            improvedRecord ||
            delta <= 0 ||
            nextRand() < Math.exp(-delta / Math.max(1, temp));

          if (accept) {
            curEval = nextEval;
            if (
              curEval.hard < bestRestartHard ||
              (curEval.hard === bestRestartHard && curEval.total < bestRestartTotal)
            ) {
              bestRestartHard = curEval.hard;
              bestRestartTotal = curEval.total;
              bestRestartPaths = paths.map((p) => p);
              if (bestRestartHard === 0) break;
            }
          } else {
            for (let gi = 0; gi < ruinGroup.length; gi++) {
              const eIdx = ruinGroup[gi];
              addPath(paths[eIdx], -1, curM, curA, curN);
              paths[eIdx] = savedSub[gi];
              addPath(savedSub[gi], 1, curM, curA, curN);
            }
          }

          temp *= cool;
        }

        paths = bestRestartPaths;
        curM.fill(0);
        curA.fill(0);
        curN.fill(0);
        for (let idx = 0; idx < R; idx++) addPath(paths[idx], 1, curM, curA, curN);
        runSweeps(2, false);

        const candidateRows = buildRowsFromPaths(paths);
        const candRelax: RelaxationRecord[] = [];
        const testReport = validateScheduleCompliance(
          candidateRows,
          dateColumns,
          roster,
          targetYear,
          targetMonth,
          targetWorkingDays,
          options.excludedEmployees ?? [],
          { tier, relaxations: candRelax }
        );
        const curHandoffSafe = isMonthEndHandoffSafe(candidateRows);
        fallbackAttempts.push({
          rows: candidateRows,
          seed: seed + restart,
          handoffSafe: curHandoffSafe
        });
        hasDagFallback = true;

        if (testReport.isFullyCompliant) {
          compliantRestarts++;
          const wScore = (curHandoffSafe ? 0 : 10000) + candRelax.length * 10;
          if (!tierBestWinner || wScore < tierBestWinner.score) {
            tierBestWinner = {
              rows: candidateRows,
              report: testReport,
              relaxations: candRelax,
              handoffSafe: curHandoffSafe,
              seed: seed + restart,
              score: wScore
            };
            if (curHandoffSafe) break;
          }
        }
      }

      if (!hasDagFallback) {
        addQuickDagFallback(dags, tier, seed);
      }

      if (tierBestWinner) {
        return {
          month: targetMonth,
          year: targetYear,
          daysCount,
          targetWorkingDays,
          dateColumns,
          rows: tierBestWinner.rows,
          validation: tierBestWinner.report,
          seed: tierBestWinner.seed,
          excludedEmployees: options.excludedEmployees ?? [],
          tier,
          relaxations: tierBestWinner.relaxations,
          needsManualReview: false,
          elapsedMs: Math.round(performance.now() - startTimeMs)
        };
      }

      const s2Hit = tryStage2AtTier(tier);
      if (s2Hit) return s2Hit;
    }
  }

  // No active tier achieved totalViolations === 0 (or Stage 3 skipped):
  // pick best attempt prioritizing 0 hard H1-H8 violations, then lowest totalViolations at maxTier.
  let chosenFallback: FallbackAttempt | null = null;
  let chosenScore = Infinity;
  for (const att of fallbackAttempts) {
    const hardRep = validateScheduleCompliance(
      att.rows,
      dateColumns,
      roster,
      targetYear,
      targetMonth,
      targetWorkingDays,
      options.excludedEmployees ?? [],
      { tier: 3 }
    );
    const maxRep = maxTier === 3
      ? hardRep
      : validateScheduleCompliance(
          att.rows,
          dateColumns,
          roster,
          targetYear,
          targetMonth,
          targetWorkingDays,
          options.excludedEmployees ?? [],
          { tier: maxTier }
        );
    const hardLegal =
      hardRep.streakViolationCount +
      hardRep.nightRestViolationCount +
      hardRep.excessiveRestViolationCount +
      hardRep.mixedBlockViolationCount +
      hardRep.workDaysMismatchCount +
      (hardRep.fixedStaffViolationCount ?? 0);
    const covTier3 = hardRep.understaffedDaysCount + hardRep.nightOverstaffedDaysCount;
    const sc =
      hardLegal * 1e9 +
      covTier3 * 1e5 +
      maxRep.totalViolations * 100 +
      (att.handoffSafe ? 0 : 1);
    if (!chosenFallback || sc < chosenScore) {
      chosenFallback = att;
      chosenScore = sc;
    }
  }

  const finalRows = chosenFallback?.rows ?? bestResultRows;
  const finalOutSeed = chosenFallback?.seed ?? finalSeed;
  const finalRelax: RelaxationRecord[] = [];
  const validation = validateScheduleCompliance(
    finalRows,
    dateColumns,
    roster,
    targetYear,
    targetMonth,
    targetWorkingDays,
    options.excludedEmployees ?? [],
    { tier: maxTier, relaxations: finalRelax }
  );

  return {
    month: targetMonth,
    year: targetYear,
    daysCount,
    targetWorkingDays,
    dateColumns,
    rows: finalRows,
    validation,
    seed: finalOutSeed,
    excludedEmployees: options.excludedEmployees ?? [],
    tier: maxTier,
    relaxations: finalRelax,
    needsManualReview: !validation.isFullyCompliant,
    elapsedMs: Math.round(performance.now() - startTimeMs)
  };
}
