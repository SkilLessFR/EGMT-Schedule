export { minCostMatching } from './hungarian';

/**
 * Calculate cost for an employee starting at a given template offset
 */
export function evaluateOffsetCost(
  history: string[],
  template: string[],
  offset: number,
  employee: string,
  targetDaysCount: number,
  day1Anchor?: string | null,
  seed: number = 0,
  targetWorkingDays?: number
): number {
  let cost = 0;
  const tLen = template.length;
  const slotShifts: string[] = [];
  for (let d = 0; d < targetDaysCount; d++) {
    slotShifts.push(template[(offset + d) % tLen]);
  }

  const firstShift = slotShifts[0];
  const secondShift = slotShifts[1];

  // 1. Boundary streak check (combines tail from previous month + start of new month)
  let tailStreak = 0;
  for (let i = history.length - 1; i >= 0; i--) {
    if (history[i] !== 'OFF' && history[i] !== 'H8' && history[i] !== 'HUP') tailStreak++;
    else break;
  }

  let headStreak = 0;
  for (let i = 0; i < Math.min(5, targetDaysCount); i++) {
    if (slotShifts[i] !== 'OFF' && slotShifts[i] !== 'H8' && slotShifts[i] !== 'HUP') headStreak++;
    else break;
  }

  if (tailStreak >= 5) {
    // If employee reached 5 consecutive working days at end of previous month,
    // they MUST start the new month with at least 2 days of OFF!
    if (firstShift !== 'OFF' || (targetDaysCount > 1 && secondShift !== 'OFF')) {
      cost += 10000000;
    }
  } else if (tailStreak > 0) {
    const lastShift = history[history.length - 1];
    if (firstShift === 'OFF') {
      if (tailStreak < 3) {
        cost += 10000000; // work streak would be prematurely cut short
      } else {
        // Taking rest after >= 3 days of work is legal, but rest must be at least 2 days
        if (targetDaysCount > 1 && secondShift !== 'OFF') {
          cost += 10000000;
        }
      }
    } else {
      // Continuing the active work block into the new month: shift type must not change without rest
      if (firstShift !== lastShift) {
        cost += 10000000;
      } else {
        const combined = tailStreak + headStreak;
        if (combined < 3 || combined > 5) {
          cost += 10000000; // Block length must be 3..5 days
        }
        // Followed by at least 2 days of OFF
        if (headStreak < targetDaysCount && slotShifts[headStreak] !== 'OFF') cost += 5000000;
        if (headStreak + 1 < targetDaysCount && slotShifts[headStreak + 1] !== 'OFF') cost += 5000000;
      }
    }
  } else if (history.length > 0 && firstShift !== 'OFF') {
    // Previous month ended on rest: a new block starting Day 1 that ends in-month must be >= 3 days
    if (headStreak < 3 && headStreak < targetDaysCount) {
      cost += 10000000;
    }
  }

  // 2. Boundary rest streak check: rest periods must be at least 2 days, at most 3 days
  let tailOffStreak = 0;
  for (let i = history.length - 1; i >= 0; i--) {
    if (history[i] === 'OFF' || history[i] === 'H8' || history[i] === 'HUP') tailOffStreak++;
    else break;
  }
  let headOffStreak = 0;
  for (let i = 0; i < Math.min(5, targetDaysCount); i++) {
    if (slotShifts[i] === 'OFF' || slotShifts[i] === 'H8' || slotShifts[i] === 'HUP') headOffStreak++;
    else break;
  }

  // If employee was on day 1 of rest at end of previous month, Day 1 of new month MUST be OFF
  if (tailOffStreak === 1) {
    if (firstShift !== 'OFF') {
      cost += 10000000; // Minimum 2 days rest violation!
    }
  }

  // If employee already had >= 3 days of rest at end of previous month, adding OFF on Day 1 exceeds 3 days.
  // A 2-day tail rest + 1 OFF day on Day 1 yields a legal 3-day rest (soft preference against unless needed).
  if (tailOffStreak >= 3 && firstShift === 'OFF') {
    cost += 10000000;
  } else if (tailOffStreak === 2 && firstShift === 'OFF') {
    cost += 100;
  }

  if (tailOffStreak > 0 && headOffStreak > 0) {
    const boundaryOff = tailOffStreak + headOffStreak;
    if (boundaryOff > 3) {
      cost += 10000000 * (boundaryOff - 3); // Rest exceeded 3 days strictly forbidden
    }
  }

  // 3. Boundary night rest check:
  if (history.length > 0) {
    const lastShift = history[history.length - 1];
    if (lastShift === 'N' && (firstShift === 'M' || firstShift === 'A' || firstShift === 'MID')) {
      cost += 10000000; // STRICTLY FORBIDDEN!
    }
  }

  // 4. Anchor check if Day 01 was defined in previous month roster (never anchor to H8 or HUP)
  if (day1Anchor && day1Anchor !== 'OFF' && day1Anchor !== 'H8' && day1Anchor !== 'HUP') {
    const lastShift = history[history.length - 1];
    if (firstShift === day1Anchor && !(lastShift === 'N' && (firstShift === 'M' || firstShift === 'A'))) {
      cost -= 1500; // Reward matching anchor
    } else {
      cost += 800;
    }
  }

  // 5. Anti-repeat check & Shift Sequence continuity check:
  // The shift generation must strictly follow sequence A -> M -> N -> A
  const EXPECTED_NEXT_SHIFT: Record<string, string> = {
    A: 'M',
    M: 'N',
    N: 'A'
  };

  let lastWorkShift: string | null = null;
  for (let i = history.length - 1; i >= 0; i--) {
    const s = history[i];
    if (s && s !== 'OFF' && s !== 'H8' && s !== 'HUP' && s !== 'ABS' && s !== 'AD') {
      lastWorkShift = s;
      break;
    }
  }

  let firstWorkShift: string | null = null;
  for (let i = 0; i < slotShifts.length; i++) {
    const s = slotShifts[i];
    if (s !== 'OFF' && s !== 'H8' && s !== 'HUP') {
      firstWorkShift = s;
      break;
    }
  }

  if (lastWorkShift && firstWorkShift) {
    if (tailOffStreak > 0 || headOffStreak > 0 || slotShifts[0] === 'OFF') {
      // Starting a new block after rest: it MUST be the expected next shift in A -> M -> N -> A!
      const expectedNext = EXPECTED_NEXT_SHIFT[lastWorkShift];
      if (expectedNext) {
        if (firstWorkShift !== expectedNext) {
          cost += 10000000; // Strictly forbidden: breaks sequence A -> M -> N -> A
        }
      } else if (firstWorkShift === lastWorkShift) {
        cost += 5000000;
      }
    }
  }

  // 6. Target working days affinity: prefer offsets that naturally hit target working days
  if (targetWorkingDays !== undefined) {
    const rawWorkDays = slotShifts.filter((s) => s !== 'OFF' && s !== 'H8' && s !== 'HUP').length;
    if (rawWorkDays < targetWorkingDays) {
      cost += 500 * (targetWorkingDays - rawWorkDays);
    } else if (rawWorkDays > targetWorkingDays) {
      cost += 60 * (rawWorkDays - targetWorkingDays);
    }
  }

  // Seed variation factor
  const pseudo = ((employee.length * 37 + offset * 19 + seed * 97) % 100);
  cost += pseudo;

  return cost;
}
