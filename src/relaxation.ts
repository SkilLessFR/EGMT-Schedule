// src/relaxation.ts
import type { ComplianceTier, RelaxableRule, RelaxationRecord } from './scheduleTypes';

export type { ComplianceTier, RelaxableRule, RelaxationRecord };

export interface TierConfig {
  tier: ComplianceTier;
  minRest: 2 | 1;
  allowShortBlock2: boolean;
  allowNightOver2: boolean;
  allowCoverageGapMA: boolean;
  budgetShare: number;
}

export interface DailyCapacityTierInput {
  tier: ComplianceTier;
  daysCount: number;
  totalRotSlots: number;
  /** Per-day count of rotating employees capable of M (act 1). */
  availM: ArrayLike<number>;
  /** Per-day count of ALL employees (rotating + fixed) capable of A (act 0). */
  availTotalA: ArrayLike<number>;
  /** Per-day count of rotating employees capable of N (act 2). */
  availN: ArrayLike<number>;
  /** Per-day maximum fixed-staff Afternoon coverage. */
  fixedMaxA: ArrayLike<number>;
  sumMinRotA: number;
  sumMaxRotA: number;
  /** Optional monthly sum cap for fixed Afternoon shifts (defaults to sum of fixedMaxA). */
  sumMaxFixedA?: number;
  sumMinM: number;
  sumMaxM: number;
  sumMinN: number;
  sumMaxN: number;
}

const TIER_CONFIGS: Record<ComplianceTier, TierConfig> = {
  0: {
    tier: 0,
    minRest: 2,
    allowShortBlock2: false,
    allowNightOver2: false,
    allowCoverageGapMA: false,
    budgetShare: 0.50,
  },
  1: {
    tier: 1,
    minRest: 2,
    allowShortBlock2: false,
    allowNightOver2: true,
    allowCoverageGapMA: false,
    budgetShare: 0.25,
  },
  2: {
    tier: 2,
    minRest: 1,
    allowShortBlock2: true,
    allowNightOver2: true,
    allowCoverageGapMA: false,
    budgetShare: 0.15,
  },
  3: {
    tier: 3,
    minRest: 1,
    allowShortBlock2: true,
    allowNightOver2: true,
    allowCoverageGapMA: true,
    budgetShare: 0.10,
  },
};

/**
 * Returns the relaxation configuration and time-budget share for a given compliance tier.
 */
export function getTierConfig(tier: ComplianceTier): TierConfig {
  return { ...TIER_CONFIGS[tier] };
}

/**
 * Checks whether a day's (rotM, totalA, n) staffing satisfies the minimum/maximum coverage
 * rules permitted at the requested compliance tier.
 */
export function isCoverageFeasibleForTier(
  rotM: number,
  totalA: number,
  n: number,
  tier: ComplianceTier = 0
): boolean {
  const cfg = TIER_CONFIGS[tier];
  if (n < 2) return false;
  if (n > 3 || (!cfg.allowNightOver2 && n > 2)) return false;
  if (cfg.allowCoverageGapMA) {
    return rotM >= 0 && totalA >= 1;
  }
  return rotM >= 1 && totalA >= 2;
}

/**
 * Computes tier-aware daily coverage penalty for Stage 3 / Section 5 DP + SA solvers.
 * - Tier 0: enforces rotM >= 1, totalA >= 2, n === 2.
 * - Tier 1 & Tier 2: allows n === 3 with soft penalty; n < 2 or n > 3 receive heavy hard penalty.
 * - Tier 3: allows rotM === 0 or totalA === 1 with high penalty; never allows totalA === 0, n < 2, or n > 3.
 */
export function tierCovPenalty(
  rotM: number,
  totalA: number,
  n: number,
  dayWeight: number = 1,
  tier: ComplianceTier = 0
): number {
  const cfg = TIER_CONFIGS[tier];
  const overSoft =
    (rotM > 1 ? 10 * (rotM - 1) + (rotM > 2 ? 25 * (rotM - 2) : 0) : 0) +
    (totalA > 2 ? 10 * (totalA - 2) + (totalA > 3 ? 25 * (totalA - 3) : 0) : 0);

  if (tier === 0) {
    const hard =
      (rotM < 1 ? 1500 * dayWeight * (1 - rotM) : 0) +
      (totalA < 2 ? 1500 * dayWeight * (2 - totalA) : 0) +
      (n !== 2 ? 1800 * dayWeight * Math.abs(n - 2) : 0);
    return hard + overSoft;
  }

  const nightHard =
    (n < 2 ? 2200 * dayWeight * (2 - n) : 0) +
    (n > 3 ? 2200 * dayWeight * (n - 3) : 0);
  const nightSoft = n >= 3 ? 120 : 0;

  if (!cfg.allowCoverageGapMA) {
    const maHard =
      (rotM < 1 ? 1500 * dayWeight * (1 - rotM) : 0) +
      (totalA < 2 ? 1500 * dayWeight * (2 - totalA) : 0);
    return maHard + nightHard + nightSoft + overSoft;
  }

  const severeAHard = totalA < 1 ? 2200 * dayWeight * (1 - totalA) : 0;
  const negativeMHard = rotM < 0 ? 2200 * dayWeight * (-rotM) : 0;
  const gapPenalty =
    (rotM === 0 ? 450 * dayWeight : 0) +
    (totalA < 2 ? 450 * dayWeight * (2 - totalA) : 0);
  return severeAHard + negativeMHard + nightHard + gapPenalty + nightSoft + overSoft;
}

/**
 * Checks whether per-day shift availability and monthly shift min/max bounds across employee DAGs
 * can satisfy daily coverage requirements at the given compliance tier.
 */
export function canSatisfyDailyCapacityForTier(input: DailyCapacityTierInput): boolean {
  const {
    tier,
    daysCount,
    totalRotSlots,
    availM,
    availTotalA,
    availN,
    fixedMaxA,
    sumMinRotA,
    sumMaxRotA,
    sumMinM,
    sumMaxM,
    sumMinN,
    sumMaxN,
  } = input;

  const cfg = TIER_CONFIGS[tier];
  const minReqM = cfg.allowCoverageGapMA ? 0 : 1;
  const minReqTotalA = cfg.allowCoverageGapMA ? 1 : 2;
  const minReqN = 2;
  const maxReqN = cfg.allowNightOver2 ? 3 : 2;

  let sumFixedA = 0;
  let needRotA = 0;
  for (let d = 0; d < daysCount; d++) {
    if (availM[d] < minReqM) return false;
    if (availTotalA[d] < minReqTotalA) return false;
    if (availN[d] < minReqN) return false;
    const fA = fixedMaxA[d] ?? 0;
    sumFixedA += fA;
    if (fA < minReqTotalA) {
      needRotA += minReqTotalA - fA;
    }
  }

  const maxFixedASum = input.sumMaxFixedA ?? sumFixedA;
  const needM = daysCount * minReqM;
  const needTotalA = daysCount * minReqTotalA;
  const minNeedN = daysCount * minReqN;
  const maxNeedN = daysCount * maxReqN;

  if (sumMaxM < needM) return false;
  if (sumMaxRotA < needRotA || sumMaxRotA + maxFixedASum < needTotalA) return false;
  if (sumMaxN < minNeedN) return false;
  if (sumMinN > maxNeedN) return false;

  if (sumMinM + needRotA + minNeedN > totalRotSlots) return false;
  if (sumMinRotA + needM + minNeedN > totalRotSlots) return false;
  if (sumMinN + needM + needRotA > totalRotSlots) return false;

  return true;
}
