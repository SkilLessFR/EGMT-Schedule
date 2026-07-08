// DayDetailsModal.tsx
import type React from 'react';
import { memo, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { ChevronDown, ChevronLeft, ChevronRight, Activity, Users, ShieldAlert, Layers, User, ArrowLeft } from 'lucide-react';
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

const SWIPE_COMMIT_THRESHOLD = 60;
const DISMISS_COMMIT_THRESHOLD = 80;
const SWIPE_EXIT_DURATION = 220;

type DragState = { x: number; y: number; axis: 'x' | 'y' | null };

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
    .liquid-blob-1 { animation: liquid-drift-1 16s ease-in-out infinite; filter: blur(40px); }
    .liquid-blob-2 { animation: liquid-drift-2 12s ease-in-out infinite alternate; filter: blur(50px); }
    .cyber-panel-border { box-shadow: inset 0 1px 1px rgba(255, 255, 255, 0.15), 0 1px 2px rgba(0, 0, 0, 0.4), 0 20px 40px -15px rgba(0, 0, 0, 0.7); }
    .text-glow { text-shadow: 0 0 12px rgba(255, 255, 255, 0.4); }
  `}} />
);

type EmployeeListProps = {
  employees: { name: string; suffix?: string }[];
  onColleagueClick: (name: string) => void;
};

const EmployeeList = memo(function EmployeeList({ employees, onColleagueClick }: EmployeeListProps) {
  return (
    <div className="space-y-2">
      {employees.map(({ name, suffix }) => (
        <div 
          key={`${name}-${suffix ?? ''}`} 
          onClick={(e) => { e.stopPropagation(); onColleagueClick(name); }}
          className="flex items-center justify-between rounded-xl border border-white/10 bg-zinc-950/40 p-3 backdrop-blur-md transition-all hover:bg-zinc-900/60 cursor-pointer active:scale-[0.99]"
        >
          <div className="flex items-center gap-3">
            <span className="flex size-7 shrink-0 items-center justify-center rounded bg-gradient-to-tr from-cyan-400 to-indigo-500 font-mono text-[10px] font-black text-black">
              {initials(name)}
            </span>
            <span className="text-[14px] font-semibold tracking-wide text-zinc-100">{name}</span>
          </div>
          {suffix && (
            <span className="rounded bg-white/5 border border-white/5 px-2 py-0.5 font-mono text-[10px] uppercase text-cyan-400">
              {suffix}
            </span>
          )}
        </div>
      ))}
    </div>
  );
});

const WorkingSection = memo(function WorkingSection({ group, onColleagueClick }: { group: ReturnType<typeof groupForShift>; onColleagueClick: (name: string) => void }) {
  return (
    <section className="cyber-panel-border rounded-2xl border border-white/10 bg-zinc-900/40 p-4 backdrop-blur-md">
      <div className="mb-3 flex items-center gap-2 border-b border-white/5 pb-2">
        <Users className="size-3.5 text-cyan-400" />
        <h3 className="font-mono text-[11px] font-black uppercase tracking-widest text-cyan-400">
          Colleagues on shift
        </h3>
      </div>
      {group && group.employees.length > 0 ? (
        <EmployeeList employees={group.employees} onColleagueClick={onColleagueClick}/>
      ) : (
        <p className="font-mono text-xs text-zinc-400 italic py-1">Standby status: No assignments detected.</p>
      )}
    </section>
  );
});

const AllShiftsSection = memo(function AllShiftsSection({ groups, expanded, onToggle, onColleagueClick }: { groups: AllShiftsGroup[]; expanded: boolean; onToggle: () => void; onColleagueClick: (name: string) => void }) {
  return (
    <section className="cyber-panel-border rounded-2xl border border-white/10 bg-zinc-900/40 backdrop-blur-md">
      <button onClick={onToggle} className="flex min-h-12 w-full items-center justify-between px-4 py-3 text-left">
        <div className="flex items-center gap-2">
          <Layers className="size-3.5 text-fuchsia-400" />
          <span className="font-mono text-[11px] font-black uppercase tracking-widest text-zinc-200">
            Shift Overview
          </span>
        </div>
        <ChevronDown className={`size-4 shrink-0 text-fuchsia-400 transition-transform duration-500 ${expanded ? 'rotate-180' : ''}`}/>
      </button>
      <div className={`grid transition-[grid-template-rows] duration-500 ${expanded ? 'grid-rows-[1fr]' : 'grid-rows-[0fr]'}`}>
        <div className="overflow-hidden">
          <div className="space-y-4 px-4 pb-4 pt-2 border-t border-white/5 max-h-[220px] overflow-y-auto">
            {groups.length === 0 && <p className="font-mono text-xs text-zinc-500">System idle.</p>}
            {groups.map((group) => (
              <div key={group.code} className="space-y-2">
                <div className="flex items-center gap-2">
                  <span className={`inline-block size-1.5 rounded-full ${solidColorFor(group.code)}`}/>
                  <h4 className="font-mono text-[11px] font-bold uppercase tracking-wider text-zinc-400">{group.label}</h4>
                </div>
                <div className="grid grid-cols-1 gap-1">
                  {group.employees.map((name) => (
                    <div 
                      key={name} onClick={(e) => { e.stopPropagation(); onColleagueClick(name); }}
                      className="rounded-lg bg-zinc-950/50 border border-white/5 px-3 py-2 font-sans text-[13px] text-zinc-300 cursor-pointer hover:bg-zinc-900/40 transition-colors"
                    >
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

type DayDetailsModalProps = {
  event: ShiftEvent | null;
  dailyRoster: DailyRoster | undefined;
  selectedEmployee: string;
  canGoPrev: boolean;
  canGoNext: boolean;
  onPrev: () => void;
  onNext: () => void;
  onClose: () => void;
  onSwitchEmployee?: (name: string) => void;
};

export default function DayDetailsModal({ event, dailyRoster, selectedEmployee, canGoPrev, canGoNext, onPrev, onNext, onClose, onSwitchEmployee }: DayDetailsModalProps) {
  const open = event != null;

  const workingGroup = useMemo(() => event ? groupForShift(event.shift, dailyRoster, selectedEmployee) : null, [event, dailyRoster, selectedEmployee]);
  const allShiftsGroups = useMemo(() => buildAllShiftsGroups(dailyRoster), [dailyRoster]);
  const dayLabel = useMemo(() => event ? new Intl.DateTimeFormat('en', { month: 'long', day: 'numeric' }).format(event.date) : '', [event]);
  const fullDateLabel = useMemo(() => event ? event.date.toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric', year: 'numeric' }) : '', [event]);

  const displayHours = useMemo(() => {
    if (!event) return '';
    const label = shiftLabel(event.shift).toLowerCase();
    if (label.includes('mid') || event.shift === 'M' || event.shift === 'MID') {
      return '09:00 - 17:00';
    }
    return shiftHours(event.shift);
  }, [event]);

  const [expanded, setExpanded] = useState(false);
  const [activeColleagueMenu, setActiveColleagueMenu] = useState<string | null>(null);

  const wasOpen = useRef(false);
  useEffect(() => {
    if (open && !wasOpen.current) {
      setExpanded(false);
      setActiveColleagueMenu(null);
    }
    wasOpen.current = open;
  }, [open]);

  const [drag, setDrag] = useState<DragState>({ x: 0, y: 0, axis: null });
  const [isDragging, setIsDragging] = useState(false);
  const touchStart = useRef<{ x: number; y: number }>({ x: 0, y: 0 });
  const dragRef = useRef<DragState>({ x: 0, y: 0, axis: null });
  const animatingRef = useRef(false);
  const contentRef = useRef<HTMLDivElement>(null);

  const setDragBoth = useCallback((next: DragState) => {
    dragRef.current = next; setDrag(next);
  }, []);

  const commitSwipe = useCallback((direction: 'next' | 'prev') => {
    if (animatingRef.current) return; animatingRef.current = true;
    setDragBoth({ x: direction === 'next' ? -window.innerWidth : window.innerWidth, y: 0, axis: 'x' });
    window.setTimeout(() => {
      if (direction === 'next') onNext(); else onPrev();
      setDragBoth({ x: 0, y: 0, axis: null }); animatingRef.current = false;
    }, SWIPE_EXIT_DURATION);
  }, [onNext, onPrev, setDragBoth]);

  const handleTouchStart = useCallback((e: React.TouchEvent) => {
    if (animatingRef.current) return;
    const target = e.target as HTMLElement;
    if (target.closest('button') || target.closest('a') || target.closest('.cursor-pointer')) return;
    const touch = e.touches[0];
    touchStart.current = { x: touch.clientX, y: touch.clientY }; setIsDragging(true);
  }, []);

  const handleTouchMove = useCallback((e: React.TouchEvent) => {
    if (!isDragging) return; const touch = e.touches[0];
    const deltaX = touch.clientX - touchStart.current.x; const deltaY = touch.clientY - touchStart.current.y;
    const current = dragRef.current;
    let axis = current.axis;
    if (!axis) {
      if (Math.abs(deltaX) < 8 && Math.abs(deltaY) < 8) return;
      axis = Math.abs(deltaX) > Math.abs(deltaY) ? 'x' : 'y';
    }
    if (axis === 'x') {
      const blocked = (deltaX > 0 && !canGoPrev) || (deltaX < 0 && !canGoNext);
      setDragBoth({ x: blocked ? deltaX / 3 : deltaX, y: 0, axis });
    } else if (axis === 'y' && contentRef.current?.scrollTop === 0 && deltaY > 0) {
      if (e.cancelable) e.preventDefault();
      setDragBoth({ x: 0, y: Math.pow(deltaY, 0.85), axis });
    }
  }, [isDragging, canGoPrev, canGoNext, setDragBoth]);

  const handleTouchEnd = useCallback(() => {
    setIsDragging(false); const final = dragRef.current;
    if (final.axis === 'x') {
      if (final.x <= -SWIPE_COMMIT_THRESHOLD && canGoNext) commitSwipe('next');
      else if (final.x >= SWIPE_COMMIT_THRESHOLD && canGoPrev) commitSwipe('prev');
      else setDragBoth({ x: 0, y: 0, axis: null });
    } else if (final.axis === 'y' && final.y > DISMISS_COMMIT_THRESHOLD) {
      setDragBoth({ x: 0, y: 0, axis: null }); onClose();
    } else {
      setDragBoth({ x: 0, y: 0, axis: null });
    }
  }, [canGoNext, canGoPrev, commitSwipe, onClose, setDragBoth]);

  const handleColleagueClick = useCallback((name: string) => {
    setActiveColleagueMenu(name);
  }, []);

  const handleSelectColleagueShifts = useCallback(() => {
    if (activeColleagueMenu && onSwitchEmployee) {
      onSwitchEmployee(activeColleagueMenu);
      setActiveColleagueMenu(null);
      onClose();
    }
  }, [activeColleagueMenu, onSwitchEmployee, onClose]);

  return (
    <>
      <CyberLiquidStyles />
      <div onClick={onClose} className={`fixed inset-0 z-40 bg-zinc-950/70 backdrop-blur-md transition-opacity duration-500 ${open ? 'opacity-100' : 'pointer-events-none opacity-0'}`} />
      <div
        role="dialog" aria-modal="true" onTouchStart={handleTouchStart} onTouchMove={handleTouchMove} onTouchEnd={handleTouchEnd}
        style={{
          transform: open ? `translate3d(0, ${drag.axis === 'y' ? drag.y : 0}px, 0)` : 'translate3d(0, 100%, 0)',
          transition: isDragging ? 'none' : 'transform 420ms cubic-bezier(0.16, 1, 0.3, 1)',
        }}
        className={`fixed inset-x-0 bottom-0 z-50 mx-auto flex h-full max-h-[85vh] w-full max-w-[430px] flex-col overflow-hidden rounded-t-[36px] border-t border-white/20 bg-zinc-950/80 text-white backdrop-blur-3xl ${open ? '' : 'pointer-events-none'}`}
      >
        <div className="absolute inset-0 -z-10 overflow-hidden opacity-50 mix-blend-screen pointer-events-none">
          <div className="liquid-blob-1 absolute -top-12 -left-12 size-52 rounded-full bg-cyan-500/30" />
          <div className="liquid-blob-2 absolute top-1/3 -right-20 size-64 rounded-full bg-fuchsia-500/25" />
        </div>

        <div className="flex shrink-0 flex-col items-center pb-1 pt-3 touch-none"><div className="h-1 w-14 rounded-full bg-white/20" /></div>

        {event && (
          <div className="relative flex-1 flex flex-col overflow-hidden">
            <div className="flex shrink-0 items-center justify-between px-3 pb-2 border-b border-white/10 touch-none">
              <button onClick={() => canGoPrev && onPrev()} disabled={!canGoPrev} className="flex h-9 items-center gap-1 px-2.5 font-mono text-[11px] text-cyan-400 disabled:opacity-10"><ChevronLeft className="size-4" />PREV</button>
              <div className="flex items-center gap-2 rounded-lg bg-zinc-900/60 px-3 py-1"><Activity className="size-3.5 text-fuchsia-400 animate-pulse" /><p className="font-mono text-[11px] font-black tracking-widest uppercase text-zinc-100">{dayLabel}</p></div>
              <button onClick={() => canGoNext && onNext()} disabled={!canGoNext} className="flex h-9 items-center gap-1 px-2.5 font-mono text-[11px] text-cyan-400 disabled:opacity-10">NEXT<ChevronRight className="size-4" /></button>
            </div>

            <div ref={contentRef} style={{ transform: `translate3d(${drag.axis === 'x' ? drag.x : 0}px, 0, 0)`, transition: isDragging ? 'none' : 'transform 260ms' }} className="overflow-y-auto px-5 pt-3 space-y-4 flex-1 pb-10">
              <p className="font-mono text-[10px] uppercase text-zinc-400 text-center mb-3">Index // <span className="text-zinc-200">{fullDateLabel}</span></p>
              
              {/* FIXED: Rectangular blur glitch completely removed by using a centered radial circular mask core layout */}
              <div className="cyber-panel-border relative overflow-hidden rounded-2xl bg-zinc-950/40 p-6 text-center border border-white/15">
                <div className={`absolute left-1/2 top-1/2 -translate-x-1/2 -translate-y-1/2 size-36 rounded-full opacity-25 blur-3xl pointer-events-none ${solidColorFor(event.shift)}`} />
                <div className="text-glow text-[38px] font-black uppercase tracking-wide text-white">{shiftLabel(event.shift)}</div>
                <div className="mt-1.5 font-mono text-xs text-zinc-300 uppercase tracking-wider">{displayHours}</div>
              </div>

              <div className="space-y-4">
                <WorkingSection group={workingGroup} onColleagueClick={handleColleagueClick}/>
                <AllShiftsSection groups={allShiftsGroups} expanded={expanded} onToggle={() => setExpanded(!expanded)} onColleagueClick={handleColleagueClick}/>
              </div>
              <button onClick={onClose} className="mt-2 flex w-full items-center justify-center gap-2 rounded-xl border border-red-500/30 bg-red-950/20 py-3.5 font-mono text-xs text-red-400 uppercase tracking-widest font-black">Close shift info</button>
            </div>

            <div 
              onClick={() => setActiveColleagueMenu(null)}
              className={`absolute inset-0 z-50 flex flex-col justify-end bg-black/40 backdrop-blur-sm transition-opacity duration-300 ${activeColleagueMenu ? 'opacity-100 pointer-events-auto' : 'opacity-0 pointer-events-none'}`}
            >
              <div 
                onClick={(e) => e.stopPropagation()}
                className={`w-full bg-zinc-900 border-t border-white/10 rounded-t-[24px] p-4 pb-6 space-y-3 transition-transform duration-300 transform ${activeColleagueMenu ? 'translate-y-0' : 'translate-y-full'}`}
              >
                <div className="text-center px-4 py-2">
                  <p className="text-[15px] font-bold text-zinc-100 truncate">{activeColleagueMenu}</p>
                  <p className="text-[11px] text-zinc-400 font-mono uppercase tracking-wider mt-0.5">Roster Actions</p>
                </div>
                <div className="rounded-xl overflow-hidden bg-zinc-950/50 border border-white/5 divide-y divide-white/5">
                  <button onClick={handleSelectColleagueShifts} className="w-full py-3.5 text-center text-cyan-400 font-medium text-[14px] active:bg-white/5 transition-colors">
                    View {activeColleagueMenu}'s shifts
                  </button>
                </div>
                <button onClick={() => setActiveColleagueMenu(null)} className="w-full py-3.5 bg-zinc-800 text-zinc-300 font-medium text-[14px] rounded-xl active:bg-zinc-700 transition-colors">
                  Cancel
                </button>
              </div>
            </div>

          </div>
        )}
      </div>
    </>
  );
}
