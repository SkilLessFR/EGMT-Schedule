import { shiftKey, ROMANIAN_HOLIDAYS } from '../scheduleUtils';

export const FIXED_MID_EMPLOYEES = [
  'Palade Alexandru-Ionut'
];

export const FIXED_WEEKDAY_MA_EMPLOYEES = [
  'Tir George Cristian'
];

export const NON_ROTATING_EMPLOYEES = [
  ...FIXED_MID_EMPLOYEES,
  ...FIXED_WEEKDAY_MA_EMPLOYEES
];

export { ROMANIAN_HOLIDAYS, shiftKey };
export const ROMANIAN_HOLIDAYS_2026 = ROMANIAN_HOLIDAYS;

// Extra non-statutory dates on which fixed staff (Palade, Tir George) are scheduled OFF
export const FIXED_STAFF_EXEMPT_OFF_DATES = new Set<string>();

export const MONTHLY_WORKING_DAYS_2026: Record<number, number> = {
  0: 18, // Jan (22 weekdays minus Jan 1, 2, 6, 7 holidays)
  1: 20, // Feb
  2: 22, // Mar
  3: 20, // Apr
  4: 20, // May
  5: 21, // Jun
  6: 23, // Jul
  7: 21, // Aug
  8: 22, // Sep
  9: 22, // Oct
  10: 20, // Nov (20 in 11 Noiembrie.xlsx)
  11: 21  // Dec (21 statutory legal working days for 24/7 operations in 2026)
};

export const DAG_NEXT_SHIFT_ACT: readonly number[] = [1, 2, 0]; // A(0)->M(1), M(1)->N(2), N(2)->A(0)
export const DAG_LEAVE_CODES = new Set(['ABS', 'H8', 'CM']);
