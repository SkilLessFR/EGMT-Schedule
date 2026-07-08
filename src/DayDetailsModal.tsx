// DayDetailsModal.tsx
import type React from 'react';
import { memo, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { ChevronDown, ChevronLeft, ChevronRight } from 'lucide-react';
import type { ShiftEvent } from './types';
import {
  type AllShiftsGroup,
  type DailyRoster,
  buildAllShiftsGroups,
  groupForShift,
  initials,
  shiftHours,
  shiftLabel,
  solidColorFor,
  GLASS_SHEET,
} from './scheduleUtils';

// Tuning constants for the gesture system.
const AXIS_LOCK_THRESHOLD = 8;
const SWIPE_COMMIT_THRESHOLD = 70;
const DISMISS_COMMIT_THRESHOLD = 90;
const SWIPE_EXIT_DURATION = 200;
const WHEEL_DELTA_THRESHOLD = 24;

type DragState = { x: number; y: number; axis: 'x' | 'y' | null };

// ---------------------------------------------------------------------------
// Presentational subcomponents
// ---------------------------------------------------------------------------

const EmployeeList = memo(function EmployeeList({ employees }: { employees: { name: string; suffix?: string }[] }) {
  return <div className="space-y-1.5">
    {employees.map(({ name, suffix }) => <div key={`${name}-${suffix ?? ''}`} className="flex items-center gap-3 rounded-[18px] bg-zinc-100/70 p-2.5 dark:bg-zinc-800/60">
      <span className="flex size-8 shrink-0 items-center justify-center rounded-full bg-white text-[11px] font-bold text-zinc-600 shadow-sm dark:bg-zinc-700 dark:text-zinc-100">{initials(name)}</span>
      <span className="text-[15px] font-medium">{name}{suffix ? <span className="text-zinc-400 dark:text-zinc-500"> · {suffix}</span> : ''}</span>
    </div>)}
  </div>;
});

const WorkingSection = memo(function WorkingSection({ group }: { group: ReturnType<typeof groupForShift> }) {
  return <section className="rounded-[20px] bg-zinc-50/80 p-4 ring-1 ring-black/[0.03] dark:bg-white/[0.04] dark:ring-white/[0.05]">
    <h3 className="text-[11px] font-semibold uppercase tracking-wide text-zinc-400 dark:text-zinc-500">Employees on This Shift</h3>
    {group && group.employees.length > 0 ? <div className="mt-3">
      <EmployeeList employees={group.employees}/>
    </div> : <p className="mt-2 text-[15px] text-zinc-400 dark:text-zinc-500">No other employees are working this shift.</p>}
  </section>;
});

const AllShiftsSection = memo(function AllShiftsSection({ groups, expanded, onToggle }: { groups: AllShiftsGroup[]; expanded: boolean; onToggle: () => void }) {
  return <section className="rounded-[20px] bg-zinc-50/80 ring-1 ring-black/[0.03] dark:bg-white/[0.04] dark:ring-white/[0.05]">
    <button
      onClick={onToggle}
      className="flex min-h-11 w-full items-center justify-between px-4 py-3 text-left transition active:bg-zinc-950/[0.03] dark:active:bg-white/[0.05]"
      aria-expanded={expanded}
    >
      <span className="text-[15px] font-semibold">{expanded ? 'Hide all shifts' : 'Show all shifts'}</span>
      <ChevronDown className={`size-4 shrink-0 text-zinc-400 transition-transform duration-300 ${expanded ? 'rotate-180' : ''}`}/>
    </button>
    <div className={`grid transition-[grid-template-rows] duration-300 ease-out ${expanded ? 'grid-rows-[1fr]' : 'grid-rows-[0fr]'}`}>
      <div className="overflow-hidden">
        <div className="space-y-4 px-4 pb-4">
          {groups.length === 0 && <p className="text-[14px] text-zinc-400 dark:text-zinc-500">No shifts recorded for this day.</p>}
