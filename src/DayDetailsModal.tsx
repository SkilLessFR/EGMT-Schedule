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
          {groups.map((group) => <div key={group.code}>
            <div className="mb-1.5 flex items-center gap-1.5">
              <span className={`inline-block size-2.5 shrink-0 rounded-full ${solidColorFor(group.code)}`}/>
              <h4 className="text-[13px] font-semibold uppercase tracking-wide text-zinc-500 dark:text-zinc-400">{group.label}</h4>
            </div>
            <div className="space-y-1">
              {group.employees.map((name) => <div key={name} className="rounded-xl bg-white px-3 py-2 text-[14px] font-medium dark:bg-zinc-800/60">{name}</div>)}
            </div>
          </div>)}
        </div>
      </div>
    </div>
  </section>;
});

// ---------------------------------------------------------------------------
// Main modal
// ---------------------------------------------------------------------------

type DayDetailsModalProps = {
  event: ShiftEvent | null;
  dailyRoster: DailyRoster | undefined;
  selectedEmployee: string;
  canGoPrev: boolean;
  canGoNext: boolean;
  onPrev: () => void;
  onNext: () => void;
  onClose: () => void;
};

export default function DayDetailsModal({ event, dailyRoster, selectedEmployee, canGoPrev, canGoNext, onPrev, onNext, onClose }: DayDetailsModalProps) {
  const open = event != null;

  const workingGroup = useMemo(() => event ? groupForShift(event.shift, dailyRoster, selectedEmployee) : null, [event, dailyRoster, selectedEmployee]);
  const allShiftsGroups = useMemo(() => buildAllShiftsGroups(dailyRoster), [dailyRoster]);

  const dayLabel = useMemo(() => event ? new Intl.DateTimeFormat('en', { month: 'long', day: 'numeric' }).format(event.date) : '', [event]);
  const fullDateLabel = useMemo(
    () => event ? event.date.toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' }) : '',
    [event],
  );

  const [expanded, setExpanded] = useState(false);
  const wasOpen = useRef(false);
  useEffect(() => {
    if (open && !wasOpen.current) setExpanded(false);
    wasOpen.current = open;
  }, [open]);

  // ---- Gesture system: one pointer session, axis-locked to either the
  // vertical dismiss-drag or the horizontal day-swipe. Pointer capture pins
  // every event from this touch/mouse to the sheet element itself, so the
  // gesture can never be picked up by the calendar (or anything else)
  // rendered underneath. ----
  const [drag, setDrag] = useState<DragState>({ x: 0, y: 0, axis: null });
  const [isDragging, setIsDragging] = useState(false);
  const [noTransition, setNoTransition] = useState(false);
  const dragStart = useRef<{ x: number; y: number; scrollTop: number } | null>(null);
  const dragRef = useRef<DragState>({ x: 0, y: 0, axis: null });
  const animatingRef = useRef(false);
  const contentRef = useRef<HTMLDivElement>(null);
  const sheetRef = useRef<HTMLDivElement>(null);
  const capturedPointerId = useRef<number | null>(null);

  const setDragBoth = useCallback((next: DragState) => {
    dragRef.current = next;
    setDrag(next);
  }, []);

  const commitSwipe = useCallback((direction: 'next' | 'prev') => {
    if (animatingRef.current) return;
    animatingRef.current = true;
    setDragBoth({ x: direction === 'next' ? -window.innerWidth : window.innerWidth, y: 0, axis: 'x' });
    window.setTimeout(() => {
      if (direction === 'next') onNext(); else onPrev();
      setNoTransition(true);
      setDragBoth({ x: 0, y: 0, axis: null });
      requestAnimationFrame(() => requestAnimationFrame(() => setNoTransition(false)));
      animatingRef.current = false;
    }, SWIPE_EXIT_DURATION);
  }, [onNext, onPrev, setDragBoth]);

  const handlePointerDown = useCallback((e: React.PointerEvent) => {
    if (animatingRef.current) return;
    dragStart.current = { x: e.clientX, y: e.clientY, scrollTop: contentRef.current?.scrollTop ?? 0 };
    setIsDragging(true);
    e.currentTarget.setPointerCapture(e.pointerId);
    capturedPointerId.current = e.pointerId;
  }, []);

  useEffect(() => {
    if (!isDragging) return;

    const handleMove = (e: PointerEvent) => {
      const start = dragStart.current;
      if (!start) return;
      const dx = e.clientX - start.x;
      const dy = e.clientY - start.y;
      const current = dragRef.current;

      let axis = current.axis;
      if (!axis) {
        if (Math.abs(dx) < AXIS_LOCK_THRESHOLD && Math.abs(dy) < AXIS_LOCK_THRESHOLD) return;
        axis = Math.abs(dx) > Math.abs(dy) ? 'x' : 'y';
      }

      if (axis === 'x') {
        const blocked = (dx > 0 && !canGoPrev) || (dx < 0 && !canGoNext);
        e.preventDefault();
        setDragBoth({ x: blocked ? dx / 3 : dx, y: 0, axis });
        return;
      }

      const startedAtTop = start.scrollTop <= 0;
      if (dy > 0 && startedAtTop) {
        e.preventDefault();
        setDragBoth({ x: 0, y: dy, axis });
      } else if (!current.axis) {
        setDragBoth({ x: 0, y: 0, axis: null });
      }
    };

    const handleUp = () => {
      setIsDragging(false);
      if (capturedPointerId.current != null && sheetRef.current?.hasPointerCapture(capturedPointerId.current)) {
        sheetRef.current.releasePointerCapture(capturedPointerId.current);
      }
      capturedPointerId.current = null;
      const final = dragRef.current; // read once, decide once — never inside setState

      if (final.axis === 'x') {
        if (final.x <= -SWIPE_COMMIT_THRESHOLD && canGoNext) {
          commitSwipe('next');
        } else if (final.x >= SWIPE_COMMIT_THRESHOLD && canGoPrev) {
          commitSwipe('prev');
        } else {
          setDragBoth({ x: 0, y: 0, axis: null });
        }
      } else if (final.axis === 'y' && final.y > DISMISS_COMMIT_THRESHOLD) {
        setDragBoth({ x: 0, y: 0, axis: null });
        onClose();
      } else {
        setDragBoth({ x: 0, y: 0, axis: null });
      }

      dragStart.current = null;
    };

    window.addEventListener('pointermove', handleMove, { passive: false });
    window.addEventListener('pointerup', handleUp);
    window.addEventListener('pointercancel', handleUp);
    return () => {
      window.removeEventListener('pointermove', handleMove);
      window.removeEventListener('pointerup', handleUp);
      window.removeEventListener('pointercancel', handleUp);
    };
  }, [isDragging, canGoPrev, canGoNext, onClose, commitSwipe, setDragBoth]);

  useEffect(() => {
    if (!open) { setDragBoth({ x: 0, y: 0, axis: null }); setIsDragging(false); }
  }, [open, setDragBoth]);

  useEffect(() => {
    if (!open) return;
    const handleKey = (e: KeyboardEvent) => {
      if (e.key === 'ArrowLeft' && canGoPrev) onPrev();
      else if (e.key === 'ArrowRight' && canGoNext) onNext();
      else if (e.key === 'Escape') onClose();
    };
    window.addEventListener('keydown', handleKey);
    return () => window.removeEventListener('keydown', handleKey);
  }, [open, canGoPrev, canGoNext, onPrev, onNext, onClose]);

  const handleWheel = useCallback((e: React.WheelEvent) => {
    if (animatingRef.current) return;
    if (Math.abs(e.deltaX) < WHEEL_DELTA_THRESHOLD || Math.abs(e.deltaX) < Math.abs(e.deltaY)) return;
    if (e.deltaX > 0 && canGoNext) commitSwipe('next');
    else if (e.deltaX < 0 && canGoPrev) commitSwipe('prev');
  }, [canGoNext, canGoPrev, commitSwipe]);

  const toggleExpanded = useCallback(() => setExpanded((v) => !v), []);
  const noTransitionOrDragging = isDragging || noTransition;

  return <>
    <div
      onClick={onClose}
      className={`fixed inset-0 z-10 bg-black/30 backdrop-blur-sm transition-opacity duration-300 ${open ? 'opacity-100' : 'pointer-events-none opacity-0'}`}
    />

    <div
      ref={sheetRef}
      role="dialog"
      aria-modal="true"
      onPointerDown={handlePointerDown}
      style={{
        transform: open ? `translateY(${drag.axis === 'y' ? drag.y : 0}px)` : 'translateY(100%)',
        transition: noTransitionOrDragging ? 'none' : 'transform 300ms cubic-bezier(0.22,1,0.36,1)',
      }}
      className={`fixed inset-x-0 bottom-0 z-20 mx-auto flex max-h-[92vh] w-full max-w-[430px] flex-col overflow-hidden ${GLASS_SHEET} ${open ? '' : 'pointer-events-none'} ${isDragging ? 'select-none' : ''}`}
    >
      <div className="flex shrink-0 flex-col items-center pb-2 pt-2.5" style={{ touchAction: 'none' }}>
        <div className="h-1.5 w-10 rounded-full bg-zinc-300 dark:bg-zinc-600"/>
      </div>

      {event && <>
        <div className="flex shrink-0 items-center justify-between px-2 pb-1" style={{ touchAction: 'none' }}>
          <button
            onClick={onPrev}
            disabled={!canGoPrev}
            className="flex min-h-11 items-center gap-1 rounded-full px-3 text-[15px] font-medium text-blue-500 transition active:scale-[0.97] disabled:pointer-events-none disabled:opacity-30"
          >
            <ChevronLeft className="size-4"/>
            Previous
          </button>
          <p className="text-[13px] font-semibold text-zinc-400 dark:text-zinc-500">{dayLabel}</p>
          <button
            onClick={onNext}
            disabled={!canGoNext}
            className="flex min-h-11 items-center gap-1 rounded-full px-3 text-[15px] font-medium text-blue-500 transition active:scale-[0.97] disabled:pointer-events-none disabled:opacity-30"
          >
            Next
            <ChevronRight className="size-4"/>
          </button>
        </div>

        <div
          ref={contentRef}
          onWheel={handleWheel}
          style={{
            transform: `translateX(${drag.axis === 'x' ? drag.x : 0}px)`,
            transition: noTransitionOrDragging ? 'none' : 'transform 220ms cubic-bezier(0.22,1,0.36,1)',
            touchAction: drag.axis ? 'none' : 'pan-y',
            paddingBottom: 'max(1.5rem, env(safe-area-inset-bottom))',
          }}
          className="overflow-y-auto px-5 pt-1"
        >
          <p className="text-[13px] font-medium text-zinc-400 dark:text-zinc-500">{fullDateLabel}</p>
          <div className={`my-4 rounded-[28px] p-6 text-center shadow-[0_10px_26px_-12px_rgba(0,0,0,0.35)] ${solidColorFor(event.shift)}`}>
            <div className="text-[34px] font-bold leading-none tracking-tight">{shiftLabel(event.shift)}</div>
            <div className="mt-2 text-[15px] font-semibold opacity-90">{shiftHours(event.shift)}</div>
          </div>
          <div className="space-y-3">
            <WorkingSection group={workingGroup}/>
            <AllShiftsSection groups={allShiftsGroups} expanded={expanded} onToggle={toggleExpanded}/>
          </div>
          <button
            onClick={onClose}
            className="mt-5 w-full rounded-2xl bg-zinc-950/5 py-3.5 text-[17px] font-semibold text-zinc-950 transition active:scale-[0.97] dark:bg-white/10 dark:text-white"
          >
            Close
          </button>
        </div>
      </>}
    </div>
  </>;
}
