// App.tsx
import type React from 'react';
import { useEffect, useMemo, useRef, useState, useCallback } from 'react';
import { BarChart3, Calendar as CalendarIcon, Check, ChevronLeft, ChevronRight, Moon, Search, Settings2, Sun, Users } from 'lucide-react';
import type { RosterData, ShiftEvent } from './types';
import { colorFor, shiftKey, shiftLabel, shiftHourValues, buildRosterIndex, eventForIso, GLASS_CARD, GLASS_NAV } from './scheduleUtils';
import DayDetailsModal from './DayDetailsModal';

const weekdays = ['M', 'T', 'W', 'T', 'F', 'S', 'S'];
const EMPLOYEE_STORAGE_KEY = 'work-schedule-employee';
const DARK_MODE_STORAGE_KEY = 'work-schedule-dark-mode';

const APP_DARK_BG = '#050505';
const APP_LIGHT_BG = '#f4f4f5';

const MONTH_AXIS_LOCK_THRESHOLD = 8;
const MONTH_SWIPE_COMMIT_THRESHOLD = 70;
const MONTH_SWIPE_DURATION = 260;
const NAV_INDICATOR_OFFSETS = ['translate-x-0', 'translate-x-[calc(100%+0.375rem)]', 'translate-x-[calc(200%+0.75rem)]'];

const IOS_SWITCH_ON = '#34c759';
const IOS_SWITCH_OFF = '#e5e5ea';

function monthDays(month: number, year: number) {
  const first = new Date(year, month, 1);
  const startOffset = (first.getDay() + 6) % 7;
  const start = new Date(year, month, 1 - startOffset);
  return Array.from({ length: 42 }, (_, index) => new Date(start.getFullYear(), start.getMonth(), start.getDate() + index));
}
function iso(date: Date) { return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`; }

function displayFileName(fileName: string | undefined | null) {
  if (!fileName || !fileName.trim()) return 'Untitled roster';
  const base = fileName.trim().split(/[\\/]/).pop() ?? fileName;
  return base.replace(/\.(xlsx|xls|xlsm|csv|json)$/i, '');
}

type ScheduleJson = {
  fileName: string;
  month: number;
  year: number;
  employees: string[];
  dateColumns: { isoDate: string }[];
  rows: Record<string, Record<string, string>>;
};

function hydrateRoster(json: ScheduleJson): RosterData {
  return {
    fileName: json.fileName,
    month: json.month,
    year: json.year,
    employees: json.employees,
    dateColumns: json.dateColumns.map(({ isoDate }, index) => ({ index, isoDate, date: new Date(isoDate) })),
    rows: json.rows,
  };
}

function eventsForEmployee(roster: RosterData, employee: string): ShiftEvent[] {
  return roster.dateColumns.reduce<ShiftEvent[]>((events, { isoDate }) => {
    const shift = roster.rows[employee]?.[isoDate] || 'OFF';
    events.push({ id: `${employee}-${isoDate}`, isoDate, shift, date: new Date(isoDate) });
    return events;
  }, []);
}

export default function App() {
  const [dark, setDark] = useState(() => {
    const saved = localStorage.getItem(DARK_MODE_STORAGE_KEY);
    return saved !== null ? saved === 'true' : true;
  });
  const [status, setStatus] = useState<'loading' | 'error' | 'loaded'>('loading');
  const [roster, setRoster] = useState<RosterData | null>(null);
  const [errorMessage, setErrorMessage] = useState('');
  const [selectedEmployee, setSelectedEmployee] = useState('');
  const [query, setQuery] = useState('');
  const [currentMonth, setCurrentMonth] = useState(new Date().getMonth());
  const [currentYear, setCurrentYear] = useState(new Date().getFullYear());
  const [selectedEvent, setSelectedEvent] = useState<ShiftEvent | null>(null);

  useEffect(() => {
    let mounted = true;
    fetch(`${import.meta.env.BASE_URL}schedule.json`)
      .then((res) => {
        if (!res.ok) throw new Error('Schedule file not found. Run npm run import.');
        return res.json() as Promise<ScheduleJson>;
      })
      .then((json) => {
        if (!mounted) return;
        const parsed = hydrateRoster(json);
        setRoster(parsed);
        setCurrentMonth(parsed.month);
        setCurrentYear(parsed.year);
        const savedEmployee = localStorage.getItem(EMPLOYEE_STORAGE_KEY);
        const initialEmployee = savedEmployee && parsed.employees.includes(savedEmployee) ? savedEmployee : (parsed.employees[0] ?? '');
        setSelectedEmployee(initialEmployee);
        setStatus('loaded');
      })
      .catch((err) => {
        if (!mounted) return;
        setErrorMessage(err instanceof Error ? err.message : 'Unable to load schedule.');
        setStatus('error');
      });
    return () => { mounted = false; };
  }, []);

  useEffect(() => {
    const html = document.documentElement;
    const body = document.body;
    const bg = dark ? APP_DARK_BG : APP_LIGHT_BG;
    
    const prev = {
      htmlBg: html.style.backgroundColor, bodyBg: body.style.backgroundColor,
      htmlOverscroll: html.style.overscrollBehaviorY, bodyOverscroll: body.style.overscrollBehaviorY,
      htmlHeight: html.style.height, bodyHeight: body.style.height,
    };
    
    html.style.backgroundColor = bg; body.style.backgroundColor = bg;
    html.style.overscrollBehaviorY = 'none'; body.style.overscrollBehaviorY = 'none';
    html.style.height = '100%'; body.style.height = '100%';
    
    return () => {
      html.style.backgroundColor = prev.htmlBg; body.style.backgroundColor = prev.bodyBg;
      html.style.overscrollBehaviorY = prev.htmlOverscroll; body.style.overscrollBehaviorY = prev.bodyOverscroll;
      html.style.height = prev.htmlHeight; body.style.height = prev.bodyHeight;
    };
  }, [dark]);

  const [activeTab, setActiveTab] = useState<'calendar' | 'reports' | 'settings'>('calendar');
  const handleTabChange = useCallback((tab: 'calendar' | 'reports' | 'settings') => {
    setActiveTab(tab);
    setSelectedEvent(null);
  }, []);

  const toggleAppearance = useCallback(() => {
    setDark((prev) => {
      const next = !prev;
      localStorage.setItem(DARK_MODE_STORAGE_KEY, String(next));
      return next;
    });
  }, []);

  const events = useMemo(() => (roster && selectedEmployee ? eventsForEmployee(roster, selectedEmployee) : []), [roster, selectedEmployee]);
  const eventMap = useMemo(() => Object.fromEntries(events.map((e) => [e.isoDate, e])), [events]);
  const rosterIndex = useMemo(() => buildRosterIndex(roster), [roster]);
  const rosterDateSet = useMemo(() => new Set(roster?.dateColumns.map((d) => d.isoDate) ?? []), [roster]);
  const filteredEmployees = useMemo(() => roster?.employees.filter((name) => name.toLowerCase().includes(query.toLowerCase())) ?? [], [query, roster]);
  const calendarDays = useMemo(() => monthDays(currentMonth, currentYear), [currentMonth, currentYear]);
  const title = useMemo(() => new Intl.DateTimeFormat('en', { month: 'long', year: 'numeric' }).format(new Date(currentYear, currentMonth)), [currentMonth, currentYear]);
  const todayIso = useMemo(() => iso(new Date()), []);

  const selectedDayIndex = useMemo(() => {
    if (!roster || !selectedEvent) return -1;
    return roster.dateColumns.findIndex((d) => d.isoDate === selectedEvent.isoDate);
  }, [roster, selectedEvent]);
  const canGoPrevDay = selectedDayIndex > 0;
  const canGoNextDay = roster ? selectedDayIndex !== -1 && selectedDayIndex < roster.dateColumns.length - 1 : false;

  const goToPrevDay = useCallback(() => {
    setSelectedEvent((current) => {
      if (!current || !roster) return current;
      const idx = roster.dateColumns.findIndex((d) => d.isoDate === current.isoDate);
      if (idx <= 0) return current;
      return eventForIso(roster, selectedEmployee, roster.dateColumns[idx - 1].isoDate);
    });
  }, [roster, selectedEmployee]);

  const goToNextDay = useCallback(() => {
    setSelectedEvent((current) => {
      if (!current || !roster) return current;
      const idx = roster.dateColumns.findIndex((d) => d.isoDate === current.isoDate);
      if (idx === -1 || idx >= roster.dateColumns.length - 1) return current;
      return eventForIso(roster, selectedEmployee, roster.dateColumns[idx + 1].isoDate);
    });
  }, [roster, selectedEmployee]);

  const currentMonthEvents = useMemo(
    () => events.filter((event) => event.date.getMonth() === currentMonth && event.date.getFullYear() === currentYear),
    [events, currentMonth, currentYear],
  );

  const monthlyStats = useMemo(() => {
    const counts: Record<string, number> = { M: 0, A: 0, N: 0, MID: 0, OFF: 0, H8: 0 };
    const monthDatesInRoster = roster ? roster.dateColumns.filter(
      (d) => d.date.getMonth() === currentMonth && d.date.getFullYear() === currentYear
    ) : [];

    monthDatesInRoster.forEach(({ isoDate }) => {
      const rawShift = roster?.rows[selectedEmployee]?.[isoDate];
      const code = shiftKey(rawShift ?? 'OFF');
      if (code in counts) {
        counts[code] += 1;
      } else {
        counts['OFF'] += 1;
      }
    });

    const workingShifts = counts.M + counts.A + counts.N + counts.MID;
    const totalHours = Object.entries(counts).reduce((sum, [code, count]) => sum + count * (shiftHourValues[code] ?? 0), 0);
    const daysOff = counts.OFF + counts.H8;
    return { counts, workingShifts, totalHours, daysOff };
  }, [roster, selectedEmployee, currentMonth, currentYear]);

  const mostSeenColleagues = useMemo(() => {
    if (!roster || !selectedEmployee) return { names: [], count: 0 };
    const matches: Record<string, number> = {};
    const targetMonthEvents = currentMonthEvents.filter(e => shiftKey(e.shift) !== 'OFF' && shiftKey(e.shift) !== 'H8');

    targetMonthEvents.forEach((userEvent) => {
      const activeUserShiftCode = shiftKey(userEvent.shift);
      roster.employees.forEach((colleague) => {
        if (colleague === selectedEmployee) return;
        const colleagueShiftRaw = roster.rows[colleague]?.[userEvent.isoDate];
        const colleagueShiftCode = shiftKey(colleagueShiftRaw ?? 'OFF');

        let isShared = false;
        if (activeUserShiftCode === colleagueShiftCode && colleagueShiftCode !== 'OFF') {
          isShared = true;
        } else if (activeUserShiftCode === 'MID' && (colleagueShiftCode === 'M' || colleagueShiftCode === 'A')) {
          isShared = true;
        } else if (colleagueShiftCode === 'MID' && (activeUserShiftCode === 'M' || activeUserShiftCode === 'A')) {
          isShared = true;
        }

        if (isShared) {
          matches[colleague] = (matches[colleague] || 0) + 1;
        }
      });
    });

    const scores = Object.values(matches);
    if (scores.length === 0) return { names: [], count: 0 };
    const maxScore = Math.max(...scores);
    const names = Object.keys(matches).filter(name => matches[name] === maxScore);
    return { names, count: maxScore };
  }, [roster, selectedEmployee, currentMonthEvents]);

  const goToPreviousMonth = useCallback(() => {
    setCurrentMonth((m) => { const d = new Date(currentYear, m - 1); setCurrentYear(d.getFullYear()); return d.getMonth(); });
  }, [currentYear]);

  const goToNextMonth = useCallback(() => {
    setCurrentMonth((m) => { const d = new Date(currentYear, m + 1); setCurrentYear(d.getFullYear()); return d.getMonth(); });
  }, [currentYear]);

  const monthGridRef = useRef<HTMLDivElement>(null);
  const monthDragStart = useRef<{ x: number; y: number } | null>(null);
  const monthAxisRef = useRef<'x' | 'y' | null>(null);
  const monthLastDx = useRef(0);
  const monthDragPointerId = useRef<number | null>(null);
  const monthCapturedPointerId = useRef<number | null>(null);
  const monthAnimatingRef = useRef(false);

  const setMonthTransform = useCallback((x: number, transition: boolean) => {
    const el = monthGridRef.current; if (!el) return;
    el.style.transition = transition ? `transform ${MONTH_SWIPE_DURATION}ms cubic-bezier(0.22,1,0.36,1)` : 'none';
    el.style.transform = `translateX(${x}px)`;
  }, []);

  const commitMonthSwipe = useCallback((dir: 'next' | 'prev') => {
    if (monthAnimatingRef.current) return; monthAnimatingRef.current = true;
    const dist = (monthGridRef.current?.offsetWidth ?? 300) + 24;
    setMonthTransform(dir === 'next' ? -dist : dist, true);
    window.setTimeout(() => {
      if (dir === 'next') goToNextMonth(); else goToPreviousMonth();
      setMonthTransform(0, false); monthAnimatingRef.current = false;
    }, MONTH_SWIPE_DURATION);
  }, [goToNextMonth, goToPreviousMonth, setMonthTransform]);

  const handleMonthPointerDown = useCallback((e: React.PointerEvent) => {
    if (monthAnimatingRef.current) return;
    monthDragStart.current = { x: e.clientX, y: e.clientY };
    monthDragPointerId.current = e.pointerId; monthAxisRef.current = null; monthLastDx.current = 0;
  }, []);

  useEffect(() => {
    const handleMove = (e: PointerEvent) => {
      const start = monthDragStart.current; const grid = monthGridRef.current;
      if (!start || !grid) return;
      const dx = e.clientX - start.x; const dy = e.clientY - start.y;
      let axis = monthAxisRef.current;
      if (!axis) {
        if (Math.abs(dx) < MONTH_AXIS_LOCK_THRESHOLD && Math.abs(dy) < MONTH_AXIS_LOCK_THRESHOLD) return;
        axis = Math.abs(dx) > Math.abs(dy) ? 'x' : 'y'; monthAxisRef.current = axis;
        if (axis === 'x') {
          grid.style.willChange = 'transform'; grid.classList.add('select-none');
          if (monthDragPointerId.current != null) {
            try { grid.setPointerCapture(monthDragPointerId.current); monthCapturedPointerId.current = monthDragPointerId.current; } catch {}
          }
        }
      }
      if (axis !== 'x') return;
      const applied = dx; monthLastDx.current = applied; e.preventDefault(); setMonthTransform(applied, false);
    };

    const handleUp = () => {
      const grid = monthGridRef.current; const axis = monthAxisRef.current;
      if (monthCapturedPointerId.current != null && grid?.hasPointerCapture(monthCapturedPointerId.current)) {
        grid.releasePointerCapture(monthCapturedPointerId.current);
      }
      monthCapturedPointerId.current = null; monthDragPointerId.current = null;
      if (axis === 'x' && grid) {
        grid.style.willChange = ''; grid.classList.remove('select-none');
        const dx = monthLastDx.current;
        if (dx <= -MONTH_SWIPE_COMMIT_THRESHOLD) commitMonthSwipe('next');
        else if (dx >= MONTH_SWIPE_COMMIT_THRESHOLD) commitMonthSwipe('prev');
        else setMonthTransform(0, true);
      }
      monthDragStart.current = null; monthAxisRef.current = null; monthLastDx.current = 0;
    };

    window.addEventListener('pointermove', handleMove, { passive: false });
    window.addEventListener('pointerup', handleUp); window.addEventListener('pointercancel', handleUp);
    return () => {
      window.removeEventListener('pointermove', handleMove);
      window.removeEventListener('pointerup', handleUp); window.removeEventListener('pointercancel', handleUp);
    };
  }, [commitMonthSwipe, setMonthTransform]);

  const handleSelectEmployee = useCallback((name: string) => {
    setSelectedEmployee(name);
    localStorage.setItem(EMPLOYEE_STORAGE_KEY, name);
  }, []);

  if (status === 'loading') {
    return <main className={dark ? 'dark' : ''}><div className="flex h-screen items-center justify-center overscroll-none bg-zinc-100 dark:bg-[#050505] text-zinc-400">Loading schedule…</div></main>;
  }

  if (status === 'error' || !roster) {
    return <main className={dark ? 'dark' : ''}><div className="flex h-screen items-center justify-center overscroll-none p-8 text-red-500 dark:bg-[#050505]">{errorMessage}</div></main>;
  }

  const sheetOpen = Boolean(selectedEvent);
  const tabs = [{ id: 'calendar' as const, label: 'Calendar', Icon: CalendarIcon }, { id: 'reports' as const, label: 'Reports', Icon: BarChart3 }, { id: 'settings' as const, label: 'Settings', Icon: Settings2 }];
  const activeTabIndex = tabs.findIndex((t) => t.id === activeTab);
  const statCards = [{ code: 'M', label: 'Morning' }, { code: 'A', label: 'Afternoon' }, { code: 'N', label: 'Night' }, { code: 'MID', label: 'Mid' }, { code: 'OFF', label: 'Off Days' }, { code: 'H8', label: 'Holiday' }];

  return (
    <main className={dark ? 'dark' : ''}>
      <div className="flex h-screen justify-center overscroll-none bg-zinc-200 dark:bg-[#050505]">
        <div className="relative flex h-screen w-full max-w-[430px] flex-col overflow-hidden bg-zinc-100 text-zinc-950 dark:bg-[#050505] dark:text-white">
          <div className={`flex min-h-0 flex-1 flex-col transition-[filter] duration-200 ${sheetOpen ? 'pointer-events-none select-none blur-[1px]' : ''}`} aria-hidden={sheetOpen}>
            
            {activeTab === 'calendar' && (
              <div className="flex min-h-0 flex-1 flex-col">
                <header className="shrink-0 px-5 pb-3 pt-6">
                  <div className="flex items-start justify-between">
                    <div className="min-w-0">
                      <h1 className="truncate text-[34px] font-bold leading-none tracking-tight">{title}</h1>
                      <p className="mt-1 truncate text-[15px] font-medium text-zinc-400 dark:text-zinc-500">{selectedEmployee}</p>
                    </div>
                  </div>
                  <div className="mt-3 flex items-center justify-center gap-1">
                    <button onClick={goToPreviousMonth} className="flex size-8 items-center justify-center rounded-full text-zinc-400"><ChevronLeft className="size-5"/></button>
                    <span className="text-[13px] font-semibold text-zinc-400">Swipe months</span>
                    <button onClick={goToNextMonth} className="flex size-8 items-center justify-center rounded-full text-zinc-400"><ChevronRight className="size-5"/></button>
                  </div>
                </header>

                <div className="flex min-h-0 flex-1 flex-col px-3 pb-24">
                  <div className="grid shrink-0 grid-cols-7 text-center text-[11px] font-semibold text-zinc-400">
                    {weekdays.map((day, i) => <div className="py-1.5" key={`${day}-${i}`}>{day}</div>)}
                  </div>
                  <div className={`relative min-h-0 flex-1 overflow-hidden ${GLASS_CARD}`}>
                    <div ref={monthGridRef} onPointerDown={handleMonthPointerDown} style={{ touchAction: 'pan-y' }} className="grid h-full grid-cols-7 grid-rows-6">
                      {calendarDays.map((day) => {
                        const dayIso = iso(day); const inRoster = rosterDateSet.has(dayIso);
                        const event = eventMap[dayIso] ?? (inRoster ? eventForIso(roster, selectedEmployee, dayIso) : undefined);
                        const colors = event ? colorFor(event.shift) : { bg: '', text: '' };
                        const isOff = event && shiftKey(event.shift) === 'OFF'; const isToday = dayIso === todayIso;
                        return <button
                          key={dayIso} onClick={() => inRoster && event && setSelectedEvent(event)} disabled={!inRoster}
                          className={`flex flex-col items-center justify-start gap-1 border-b border-r border-zinc-950/[0.04] pt-1.5 dark:border-white/[0.06] ${inRoster ? 'active:scale-[0.96]' : ''} ${day.getMonth() !== currentMonth ? 'opacity-30' : ''}`}
                        >
                          <span className={`flex size-8 items-center justify-center rounded-full text-[18px] ${isToday ? 'bg-red-500 text-white font-bold' : ''}`}>{day.getDate()}</span>
                          {event ? (isOff ? <span className="mt-0.5 size-1.5 rounded-full bg-zinc-300 dark:bg-zinc-600"/> : <span className={`mx-1 truncate rounded-full px-1.5 py-[1px] text-[9px] font-bold ${colors.bg} ${colors.text}`}>{event.shift}</span>) : null}
                        </button>;
                      })}
                    </div>
                  </div>
                </div>
              </div>
            )}

            {activeTab === 'reports' && (
              <div className="flex min-h-0 flex-1 flex-col overflow-y-auto px-5 pb-24 pt-6">
                <h1 className="text-[34px] font-bold tracking-tight">Reports</h1>
                <p className="mt-1 text-[15px] font-medium text-zinc-400 dark:text-zinc-500">{selectedEmployee} · {title}</p>

                <section className="mt-4">
                  <div className={`p-4 border border-blue-500/20 bg-blue-500/5 ${GLASS_CARD}`}>
                    <div className="flex items-center gap-2 text-blue-600 dark:text-blue-400 mb-1.5">
                      <Users className="size-4" />
                      <h3 className="text-[12px] font-bold uppercase tracking-wider">Most Seen Colleague This Month</h3>
                    </div>
                    {mostSeenColleagues.names.length > 0 ? (
                      <div>
                        <p className="text-[16px] font-bold tracking-wide">
                          {mostSeenColleagues.names.join(', ')}
                        </p>
                        <p className="text-[11px] font-mono text-zinc-400 mt-0.5">
                          Shared {mostSeenColleagues.count} operational shifts together
                        </p>
                      </div>
                    ) : (
                      <p className="text-xs font-medium text-zinc-400 italic">No shared shift duties recorded this month.</p>
                    )}
                  </div>
                </section>

                <div className="mt-4 grid grid-cols-2 gap-3">
                  {statCards.map(({ code, label }) => <div key={code} className={`p-4 ${GLASS_CARD}`}>
                    <p className="text-[13px] font-semibold text-zinc-400">{label}</p>
                    <p className="mt-1 text-[28px] font-bold tracking-tight">{monthlyStats.counts[code] ?? 0}</p>
                  </div>)}
                </div>

                <div className="mt-3 grid grid-cols-2 gap-3">
                  <div className="rounded-[28px] bg-blue-500/10 p-4"><p className="text-[13px] font-semibold text-blue-500">Working Shifts</p><p className="mt-1 text-[28px] font-bold text-blue-500">{monthlyStats.workingShifts}</p></div>
                  <div className="rounded-[28px] bg-green-500/10 p-4"><p className="text-[13px] font-semibold text-green-600 dark:text-green-400">Total Hours</p><p className="mt-1 text-[28px] font-bold text-green-600 dark:text-green-400">{monthlyStats.totalHours}h</p></div>
                </div>
              </div>
            )}

            {activeTab === 'settings' && (
              <div className="flex min-h-0 flex-1 flex-col overflow-y-auto px-5 pb-24 pt-6">
                <h1 className="text-[34px] font-bold tracking-tight">Settings</h1>
                
                {/* Restored Appearance Mode Toggler Section */}
                <section className="mt-6">
                  <h3 className="px-1 text-[13px] font-semibold uppercase tracking-wide text-zinc-400 dark:text-zinc-500">Appearance</h3>
                  <div className={`mt-2 flex items-center justify-between p-4 ${GLASS_CARD}`}>
                    <div className="flex items-center gap-3">
                      {dark ? <Moon className="size-5 text-zinc-400"/> : <Sun className="size-5 text-zinc-500"/>}
                      <div>
                        <p className="text-[15px] font-semibold">Dark Appearance</p>
                        <p className="text-[12px] text-zinc-400 dark:text-zinc-500">Optimizes display contrast</p>
                      </div>
                    </div>
                    <button
                      onClick={toggleAppearance}
                      style={{ backgroundColor: dark ? IOS_SWITCH_ON : IOS_SWITCH_OFF }}
                      className="relative h-[31px] w-[51px] shrink-0 rounded-full p-0.5 transition-colors duration-200"
                    >
                      <div className={`h-[27px] w-[27px] rounded-full bg-white shadow-sm ring-1 ring-black/[0.04] transition-transform duration-200 ${dark ? 'translate-x-[20px]' : 'translate-x-0'}`} />
                    </button>
                  </div>
                </section>

                {/* Restored Complete Employee Switcher List Section */}
                <section className="mt-6">
                  <div className="flex items-center justify-between px-1">
                    <h3 className="text-[13px] font-semibold uppercase tracking-wide text-zinc-400 dark:text-zinc-500">Active Roster</h3>
                    <span className="text-[11px] font-mono text-zinc-400 dark:text-zinc-500 truncate max-w-[180px]">{displayFileName(roster?.fileName)}</span>
                  </div>
                  <div className={`mt-2 ${GLASS_CARD}`}>
                    <div className="flex items-center gap-2 px-4 pt-3.5"><Search className="size-4 text-zinc-400"/><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Search employee" className="w-full bg-transparent py-2 text-[15px] outline-none"/></div>
                    <div className="max-h-64 divide-y divide-zinc-950/[0.06] overflow-y-auto dark:divide-white/[0.06]">{filteredEmployees.map((name) => <button key={name} onClick={() => handleSelectEmployee(name)} className="flex w-full items-center justify-between px-4 py-3 text-left"><span className="text-[15px]">{name}</span>{name === selectedEmployee && <Check className="size-4 text-blue-500"/>}</button>)}</div>
                  </div>
                </section>
              </div>
            )}
          </div>

          <DayDetailsModal
            event={selectedEvent}
            dailyRoster={selectedEvent ? rosterIndex[selectedEvent.isoDate] : undefined}
            selectedEmployee={selectedEmployee}
            canGoPrev={canGoPrevDay} canGoNext={canGoNextDay}
            onPrev={goToPrevDay} onNext={goToNextDay}
            onClose={() => setSelectedEvent(null)}
            onSwitchEmployee={handleSelectEmployee}
          />

          <nav aria-hidden={sheetOpen} className={`fixed inset-x-0 bottom-0 z-10 flex justify-center px-5 pb-3 transition-[filter] duration-200 ${sheetOpen ? 'blur-[1px] pointer-events-none' : ''}`}>
            <div className={`relative flex w-full max-w-[380px] p-1.5 ${GLASS_NAV}`}>
              <div style={{ width: 'calc((100% - 1.5rem) / 3)' }} className={`absolute inset-y-1.5 left-1.5 rounded-[22px] bg-white/85 shadow-md transition-transform duration-[220ms] dark:bg-white/[0.16] ${NAV_INDICATOR_OFFSETS[activeTabIndex] || NAV_INDICATOR_OFFSETS[0]}`} />
              {tabs.map(({ id, label, Icon }) => <button key={id} onClick={() => handleTabChange(id)} className={`relative z-10 flex flex-1 flex-col items-center gap-0.5 py-2 ${activeTab === id ? 'text-blue-500' : 'text-zinc-400'}`}><Icon className="size-5"/><span className="text-[10px] font-semibold">{label}</span></button>)}
            </div>
          </nav>
        </div>
      </div>
    </main>
  );
}
