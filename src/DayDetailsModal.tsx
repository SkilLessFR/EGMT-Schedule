// DayDetailsModal.tsx
import type React from 'react';
import { memo, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { ChevronDown, ChevronLeft, ChevronRight, Activity, Users, ShieldAlert, Layers } from 'lucide-react';
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
} from './scheduleUtils';

const AXIS_LOCK_THRESHOLD = 8;
const SWIPE_COMMIT_THRESHOLD = 70;
const DISMISS_COMMIT_THRESHOLD = 90;
const SWIPE_EXIT_DURATION = 200;
const WHEEL_DELTA_THRESHOLD = 24;

type DragState = { x: number; y: number; axis: 'x' | 'y' | null };

// ---------------------------------------------------------------------------
// Advanced Liquid Plasma & Cyberglass CSS Engine
// ---------------------------------------------------------------------------
const CyberLiquidStyles = () => (
  <style dangerouslySetInnerHTML={{ __html: `
    @keyframes liquid-drift-1 {
      0% { transform: translate(0px, 0px) scale(1) rotate(0deg); }
      33% { transform: translate(40px, -60px) scale(1.4) rotate(120deg); }
      66% { transform: translate(-30px, 50px) scale(0.8) rotate(240deg); }
      100% { transform: translate(0px, 0px) scale(1) rotate(360deg); }
    }
    @keyframes liquid-drift-2 {
      0% { transform: translate(0px, 0px) scale(1.2) rotate(180deg); }
      50% { transform: translate(-50px, 40px) scale(0.7) rotate(0deg); }
      100% { transform: translate(0px, 0px) scale(1.2) rotate(180deg); }
    }
    .liquid-blob-1 {
      animation: liquid-drift-1 16s ease-in-out infinite;
      filter: blur(40px);
    }
    .liquid-blob-2 {
      animation: liquid-drift-2 12s ease-in-out infinite alternate;
      filter: blur(50px);
    }
    .cyber-panel-border {
      box-shadow: 
        inset 0 1px 1px rgba(255, 255, 255, 0.3),
        0 1px 2px rgba(0, 0, 0, 0.4),
        0 20px 40px -15px rgba(0, 0, 0, 0.7);
    }
    .text-glow {
      text-shadow: 0 0 12px rgba(255, 255, 255, 0.4);
    }
  `}} />
);

// ---------------------------------------------------------------------------
// Presentational subcomponents
// ---------------------------------------------------------------------------

const EmployeeList = memo(function EmployeeList({ employees }: { employees: { name: string; suffix?: string }[] }) {
  return (
    <div className="space-y-2">
      {employees.map(({ name, suffix }) => (
        <div 
          key={`${name}-${suffix ?? ''}`} 
          className="flex items-center justify-between rounded-xl border border-white/10 bg-zinc-950/40 p-3 backdrop-blur-md transition-all hover:bg-zinc-900/60"
        >
          <div className="flex items-center gap-3">
            <span className="flex size-7 shrink-0 items-center justify-center rounded bg-gradient-to-tr from-cyan-400 to-indigo-500 font-mono text-[10px] font-black text-black shadow-[0_0_10px_rgba(34,211,238,0.4)]">
              {initials(name)}
            </span>
            <span className="text-[14px] font-semibold tracking-wide text-zinc-100">{name}</span>
          </div>
          {suffix && (
            <span className="rounded bg-white/5 border border-white/5 px-2 py-0.5 font-mono text-[10px] uppercase text-cyan-400 tracking-wider">
              {suffix}
            </span>
          )}
        </div>
      ))}
    </div>
  );
});

const WorkingSection = memo(function WorkingSection({ group }: { group: ReturnType<typeof groupForShift> }) {
  return (
    <section className="cyber-panel-border rounded-2xl border border-white/10 bg-zinc-900/40 p-4 backdrop-blur-md">
      <div className="mb-3 flex items-center gap-2 border-b border-white/5 pb-2">
        <Users className="size-3.5 text-cyan-400" />
        <h3 className="font-mono text-[11px] font-black uppercase tracking-widest text-cyan-400">
          Assigned Personnel
        </h3>
      </div>
      {group && group.employees.length > 0 ? (
        <EmployeeList employees={group.employees}/>
      ) : (
        <p className="font-mono text-xs text-zinc-400 italic py-1">Standby status: No alternative assignments detected.</p>
      )}
    </section>
  );
});

const AllShiftsSection = memo(function AllShiftsSection({ groups, expanded, onToggle }: { groups: AllShiftsGroup[]; expanded: boolean; onToggle: () => void }) {
  return (
    <section className="cyber-panel-border rounded-2xl border border-white/10 bg-zinc-900/40 backdrop-blur-md">
      <button
        onClick={onToggle}
        className="flex min-h-12 w-full items-center justify-between px-4 py-3 text-left transition hover:bg-white/5"
        aria-expanded={expanded}
      >
        <div className="flex items-center gap-2">
          <Layers className="size-3.5 text-fuchsia-400" />
          <span className="font-mono text-[11px] font-black uppercase tracking-widest text-zinc-200">
            Timeline Overview
          </span>
        </div>
        <ChevronDown className={`size-4 shrink-0 text-fuchsia-400 transition-transform duration-500 cubic-bezier(0.16, 1, 0.3, 1) ${expanded ? 'rotate-180' : ''}`}/>
      </button>
      <div className={`grid transition-[grid-template-rows] duration-500 ease-[cubic-bezier(0.16,1,0.3,1)] ${expanded ? 'grid-rows-[1fr]' : 'grid-rows-[0fr]'}`}>
        <div className="overflow-hidden">
          <div className="space-y-4 px-4 pb-4 pt-2 border-t border-white/5">
            {groups.length === 0 && <p className="font-mono text-xs text-zinc-500">System idle.</p>}
            {groups.map((group) => (
              <div key={group.code} className="space-y-2">
                <div className="flex items-center gap-2">
                  <span className={`inline-block size-1.5 shrink-0 rounded-full shadow-[0_0_8px_currentColor] ${solidColorFor(group.code)}`}/>
                  <h4 className="font-mono text-[11px] font-bold uppercase tracking-wider text-zinc-400">{group.label}</h4>
                </div>
                <div className="grid grid-cols-1 gap-1">
                  {group.employees.map((name) => (
                    <div key={name} className="rounded-lg bg-zinc-950/50 border border-white/5 px-3 py-2 font-sans text-[13px] text-zinc-300">
                      {name}
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </section>
  );
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

  const [drag, setDrag] = useState<DragState>({ x: 0, y: 0, axis: null });
  const [isDragging, setIsDragging] = useState(false);
  const [noTransition, setNoTransition] = useState(false);
  const dragStart = useRef<{ x: number; y: number; scrollTop: number } | null>(null);
  const dragRef = useRef<DragState>({ x: 0, y: 0, axis: null });
  const animatingRef = useRef(false);
  const contentRef = useRef<HTMLDivElement>(null);
  const sheetRef = useRef<HTMLDivElement>(null);
  const dragPointerId = useRef<number | null>(null);
  const capturedPointerId = useRef<number | null>(null);
  const suppressNextClickRef = useRef(false);

  useEffect(() => {
    const swallowGhostClick = (e: MouseEvent) => {
      if (!suppressNextClickRef.current) return;
      suppressNextClickRef.current = false;
      e.preventDefault();
      e.stopPropagation();
    };
    window.addEventListener('click', swallowGhostClick, true);
    return () => window.removeEventListener('click', swallowGhostClick, true);
  }, []);

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
    dragPointerId.current = e.pointerId;
    setIsDragging(true);
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
        if (dragPointerId.current != null && sheetRef.current) {
          try {
            sheetRef.current.setPointerCapture(dragPointerId.current);
            capturedPointerId.current = dragPointerId.current;
          } catch {
            // Safe fallback
          }
        }
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
      dragPointerId.current = null;
      const final = dragRef.current;

      if (final.axis === 'x') {
        if (final.x <= -SWIPE_COMMIT_THRESHOLD && canGoNext) {
          commitSwipe('next');
        } else if (final.x >= SWIPE_COMMIT_THRESHOLD && canGoPrev) {
          commitSwipe('prev');
        } else {
          setDragBoth({ x: 0, y: 0, axis: null });
        }
      } else if (final.axis === 'y' && final.y > DISMISS_COMMIT_THRESHOLD) {
        suppressNextClickRef.current = true;
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

  return (
    <>
      <CyberLiquidStyles />
      
      {/* Absolute Dark Deep Backdrop Overlay */}
      <div
        onClick={onClose}
        className={`fixed inset-0 z-40 bg-zinc-950/70 backdrop-blur-md transition-opacity duration-500 ${
          open ? 'opacity-100' : 'pointer-events-none opacity-0'
        }`}
      />

      {/* Main Liquid Cyber Glass Sheet Structure */}
      <div
        ref={sheetRef}
        role="dialog"
        aria-modal="true"
        onPointerDown={handlePointerDown}
        style={{
          transform: open ? `translateY(${drag.axis === 'y' ? drag.y : 0}px)` : 'translateY(100%)',
          transition: noTransitionOrDragging ? 'none' : 'transform 400ms cubic-bezier(0.16, 1, 0.3, 1)',
        }}
        className={`fixed inset-x-0 bottom-0 z-50 mx-auto flex max-h-[92vh] w-full max-w-[430px] flex-col overflow-hidden rounded-t-[36px] border-t border-x border-white/20 bg-zinc-950/80 text-white shadow-[0_-20px_60px_rgba(0,0,0,0.8)] backdrop-blur-3xl ${
          open ? '' : 'pointer-events-none'
        } ${isDragging ? 'select-none' : ''}`}
      >
        
        {/* PHYSICAL LIQUID ENGINE: Moving organic elements trapped beneath the sheet */}
        <div className="absolute inset-0 -z-10 overflow-hidden opacity-50 mix-blend-screen pointer-events-none">
          <div className="liquid-blob-1 absolute -top-12 -left-12 size-52 rounded-full bg-cyan-500/30" />
          <div className="liquid-blob-2 absolute top-1/3 -right-20 size-64 rounded-full bg-fuchsia-500/25" />
          <div className="liquid-blob-1 absolute -bottom-16 left-1/4 size-48 rounded-full bg-indigo-500/30" />
        </div>

        {/* Top Control Notch Bar */}
        <div className="flex shrink-0 flex-col items-center pb-1 pt-3" style={{ touchAction: 'none' }}>
          <div className="h-1 w-14 rounded-full bg-white/20" />
        </div>

        {event && (
          <>
            {/* Header Telemetry Navigation Block */}
            <div className="flex shrink-0 items-center justify-between px-3 pb-2 border-b border-white/10" style={{ touchAction: 'none' }}>
              <button
                onClick={onPrev}
                disabled={!canGoPrev}
                className="flex h-9 items-center gap-1 rounded-xl px-2.5 font-mono text-[11px] font-black tracking-wider text-cyan-400 transition-all hover:bg-white/5 active:scale-95 disabled:pointer-events-none disabled:opacity-10"
              >
                <ChevronLeft className="size-4" />
                PREV
              </button>
              
              <div className="flex items-center gap-2 rounded-lg bg-zinc-900/60 border border-white/10 px-3 py-1">
                <Activity className="size-3.5 text-fuchsia-400 animate-pulse" />
                <p className="font-mono text-[11px] font-black tracking-widest text-zinc-100 uppercase">{dayLabel}</p>
              </div>

              <button
                onClick={onNext}
                disabled={!canGoNext}
                className="flex h-9 items-center gap-1 rounded-xl px-2.5 font-mono text-[11px] font-black tracking-wider text-cyan-400 transition-all hover:bg-white/5 active:scale-95 disabled:pointer-events-none disabled:opacity-10"
              >
                NEXT
                <ChevronRight className="size-4" />
              </button>
            </div>

            {/* Scrollable Container Body */}
            <div
              ref={contentRef}
              onWheel={handleWheel}
              style={{
                transform: `translateX(${drag.axis === 'x' ? drag.x : 0}px)`,
                transition: noTransitionOrDragging ? 'none' : 'transform 260ms cubic-bezier(0.16, 1, 0.3, 1)',
                touchAction: drag.axis ? 'none' : 'pan-y',
                paddingBottom: 'max(2rem, env(safe-area-inset-bottom))',
              }}
              className="overflow-y-auto px-5 pt-3"
            >
              <p className="font-mono text-[10px] uppercase tracking-widest text-zinc-400 text-center mb-3">
                Index // <span className="text-zinc-200">{fullDateLabel}</span>
              </p>

              {/* Liquid-Frosted Glass Hero Shift Badge */}
              <div className="cyber-panel-border relative my-4 overflow-hidden rounded-2xl bg-zinc-950/40 p-6 text-center border border-white/20 shadow-2xl">
                {/* Dynamic colored ambient backing syncs with actual code styles color */}
                <div className={`absolute -inset-10 -z-10 opacity-30 blur-2xl ${solidColorFor(event.shift)}`} />
                <div className="absolute inset-0 bg-gradient-to-b from-white/10 to-transparent mix-blend-overlay" />
                
                <div className="text-glow text-[38px] font-black tracking-tight uppercase leading-none text-white">
                  {shiftLabel(event.shift)}
                </div>
                <div className="mt-2 font-mono text-xs font-black tracking-widest text-zinc-300 uppercase">
                  {shiftHours(event.shift)}
                </div>
              </div>

              {/* Main Structural Information Groups */}
              <div className="space-y-4">
                <WorkingSection group={workingGroup}/>
                <AllShiftsSection groups={allShiftsGroups} expanded={expanded} onToggle={toggleExpanded}/>
              </div>

              {/* Modern Flush Close System Button */}
              <button
                onClick={onClose}
                className="mt-6 flex w-full items-center justify-center gap-2 rounded-xl border border-red-500/30 bg-red-950/20 py-3.5 font-mono text-xs font-black uppercase tracking-widest text-red-400 transition-all hover:bg-red-500/20 active:scale-98"
              >
                <ShieldAlert className="size-4" />
                Close Operational Log
              </button>
            </div>
          </>
        )}
      </div>
    </>
  );
}
