// App.tsx
import { useEffect, useMemo, useState, useCallback } from 'react';
import { BarChart3, Calendar as CalendarIcon, Check, ChevronLeft, ChevronRight, Moon, Search, Settings2, Sun } from 'lucide-react';
import type { RosterData, ShiftEvent } from './types';
import { colorFor, shiftKey, shiftLabel, shiftHourValues, buildRosterIndex, eventForIso } from './scheduleUtils';
import DayDetailsModal from './DayDetailsModal';

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const weekdays = ['M', 'T', 'W', 'T', 'F', 'S', 'S'];
const EMPLOYEE_STORAGE_KEY = 'work-schedule-employee';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

type LoadStatus = 'loading' | 'error' | 'loaded';

/** Shape of the JSON file produced offline by `npm run import`. */
type ScheduleJson = {
  fileName: string;
  month: number;
  year: number;
  employees: string[];
  dateColumns: { isoDate: string }[];
  rows: Record<string, Record<string, string>>;
};

// ---------------------------------------------------------------------------
// Utility helpers
// ---------------------------------------------------------------------------

function monthDays(month: number, year: number) {
  const first = new Date(year, month, 1);
  const startOffset = (first.getDay() + 6) % 7;
  const start = new Date(year, month, 1 - startOffset);
  return Array.from({ length: 42 }, (_, index) => new Date(start.getFullYear(), start.getMonth(), start.getDate() + index));
}
function iso(date: Date) { return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`; }

/** Turns whatever `fileName` the import script wrote (full path, bare name, or missing) into a clean display string. */
function displayFileName(fileName: string | undefined | null) {
  if (!fileName || !fileName.trim()) return 'Untitled roster';
  const base = fileName.trim().split(/[\\/]/).pop() ?? fileName;
  return base.replace(/\.(xlsx|xls|xlsm|csv|json)$/i, '');
}

/** Converts the raw JSON payload into the RosterData shape the app works with. */
function hydrateRoster(json: ScheduleJson): RosterData {
  return {
    fileName: json.fileName,
    month: json.month,
    year: json.year,
    employees: json.employees,
    dateColumns: json.dateColumns.map(({ isoDate }, index) => ({
      index,
      isoDate,
      date: new Date(isoDate),
    })),
    rows: json.rows,
  };
}

/** Builds the per-employee shift events, reviving each isoDate string into a real Date object. */
function eventsForEmployee(roster: RosterData, employee: string): ShiftEvent[] {
  return roster.dateColumns.reduce<ShiftEvent[]>((events, { isoDate }) => {
    const shift = roster.rows[employee]?.[isoDate];
    if (shift) events.push({ id: `${employee}-${isoDate}`, isoDate, shift, date: new Date(isoDate) });
    return events;
  }, []);
}

// ---------------------------------------------------------------------------
// App
// ---------------------------------------------------------------------------

export default function App() {
  const [dark, setDark] = useState(true);
  const [status, setStatus] = useState<LoadStatus>('loading');
  const [roster, setRoster] = useState<RosterData | null>(null);
  const [errorMessage, setErrorMessage] = useState('');
  const [selectedEmployee, setSelectedEmployee] = useState('');
  const [query, setQuery] = useState('');
  const [currentMonth, setCurrentMonth] = useState(new Date().getMonth());
  const [currentYear, setCurrentYear] = useState(new Date().getFullYear());
  const [selectedEvent, setSelectedEvent] = useState<ShiftEvent | null>(null);

  // Load the schedule on mount. (React 18 StrictMode intentionally mounts
  // twice in dev — each mount gets its own `mounted` flag, so this is safe
  // and only ever commits state from the mount that's still active.)
  useEffect(() => {
    let mounted = true;

    fetch(`${import.meta.env.BASE_URL}schedule.json`)
      .then((response) => {
        if (!response.ok) throw new Error('Schedule file not found. Run npm run import -- path/to/roster.xlsx.');
        return response.json() as Promise<ScheduleJson>;
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

  // Force the real <html>/<body> to black and disable rubber-band scrolling so
  // iOS Safari never reveals a white flash behind the app during overscroll.
  useEffect(() => {
    const html = document.documentElement;
    const body = document.body;
    const prev = {
      htmlBg: html.style.backgroundColor,
      bodyBg: body.style.backgroundColor,
      htmlOverscroll: html.style.overscrollBehaviorY,
      bodyOverscroll: body.style.overscrollBehaviorY,
      htmlHeight: html.style.height,
      bodyHeight: body.style.height,
    };
    html.style.backgroundColor = '#000000';
    body.style.backgroundColor = '#000000';
    html.style.overscrollBehaviorY = 'none';
    body.style.overscrollBehaviorY = 'none';
    html.style.height = '100%';
    body.style.height = '100%';
    return () => {
      html.style.backgroundColor = prev.htmlBg;
      body.style.backgroundColor = prev.bodyBg;
      html.style.overscrollBehaviorY = prev.htmlOverscroll;
      body.style.overscrollBehaviorY = prev.bodyOverscroll;
      html.style.height = prev.htmlHeight;
      body.style.height = prev.bodyHeight;
    };
  }, []);

  // Three-tab navigation. Switching tabs never touches roster/employee/month
  // state, so returning to Calendar always shows what you left it showing.
  const [activeTab, setActiveTab] = useState<'calendar' | 'reports' | 'settings'>('calendar');
  const handleTabChange = useCallback((tab: 'calendar' | 'reports' | 'settings') => {
    setActiveTab(tab);
    setSelectedEvent(null);
  }, []);

  const events = useMemo(() => (roster && selectedEmployee ? eventsForEmployee(roster, selectedEmployee) : []), [roster, selectedEmployee]);
  const eventMap = useMemo(() => Object.fromEntries(events.map((event) => [event.isoDate, event])), [events]);
  const rosterIndex = useMemo(() => buildRosterIndex(roster), [roster]);
  const filteredEmployees = useMemo(() => roster?.employees.filter((name) => name.toLowerCase().includes(query.toLowerCase())) ?? [], [query, roster]);
  const calendarDays = useMemo(() => monthDays(currentMonth, currentYear), [currentMonth, currentYear]);
  const title = useMemo(() => new Intl.DateTimeFormat('en', { month: 'long', year: 'numeric' }).format(new Date(currentYear, currentMonth)), [currentMonth, currentYear]);
  const todayIso = useMemo(() => iso(new Date()), []);

  // Day-details navigation: bounds and handlers, based on the roster's own
  // date columns (not the padded calendar grid, which spills into other months).
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

  // Reports tab: everything derived from the roster already loaded above —
  // no hardcoded values, scoped to the selected employee + displayed month.
  const currentMonthEvents = useMemo(
    () => events.filter((event) => event.date.getMonth() === currentMonth && event.date.getFullYear() === currentYear),
    [events, currentMonth, currentYear],
  );
  const monthlyStats = useMemo(() => {
    const counts: Record<string, number> = { M: 0, A: 0, N: 0, MID: 0, OFF: 0, H8: 0 };
    currentMonthEvents.forEach((event) => {
      const code = shiftKey(event.shift);
      if (code in counts) counts[code] += 1;
    });
    const workingShifts = counts.M + counts.A + counts.N + counts.MID;
    const totalHours = Object.entries(counts).reduce((sum, [code, count]) => sum + count * (shiftHourValues[code] ?? 0), 0);
    const daysOff = counts.OFF + counts.H8;
    return { counts, workingShifts, totalHours, daysOff };
  }, [currentMonthEvents]);

  const goToPreviousMonth = useCallback(() => {
    setCurrentMonth((month) => {
      const d = new Date(currentYear, month - 1);
      setCurrentYear(d.getFullYear());
      return d.getMonth();
    });
  }, [currentYear]);

  const goToNextMonth = useCallback(() => {
    setCurrentMonth((month) => {
      const d = new Date(currentYear, month + 1);
      setCurrentYear(d.getFullYear());
      return d.getMonth();
    });
  }, [currentYear]);

  const handleSelectEmployee = useCallback((name: string) => {
    setSelectedEmployee(name);
    localStorage.setItem(EMPLOYEE_STORAGE_KEY, name);
  }, []);

  if (status === 'loading') {
    return <main className={dark ? 'dark' : ''}><div className="flex h-screen items-center justify-center overscroll-none bg-zinc-100 text-zinc-950 dark:bg-black dark:text-white">
      <p className="text-[15px] font-medium text-zinc-400 dark:text-zinc-500">Loading schedule…</p>
    </div></main>;
  }

  if (status === 'error' || !roster) {
    return <main className={dark ? 'dark' : ''}><div className="flex h-screen items-center justify-center overscroll-none bg-zinc-100 p-8 text-zinc-950 dark:bg-black dark:text-white">
      <div className="max-w-md rounded-3xl bg-red-500/10 p-6 text-center text-[15px] font-medium text-red-500">{errorMessage}</div>
    </div></main>;
  }

  const sheetOpen = Boolean(selectedEvent);
  const tabs = [
    { id: 'calendar' as const, label: 'Calendar', Icon: CalendarIcon },
    { id: 'reports' as const, label: 'Reports', Icon: BarChart3 },
    { id: 'settings' as const, label: 'Settings', Icon: Settings2 },
  ];
  const statCards = [
    { code: 'M', label: 'Morning' },
    { code: 'A', label: 'Afternoon' },
    { code: 'N', label: 'Night' },
    { code: 'MID', label: 'Mid' },
    { code: 'OFF', label: 'Off Days' },
    { code: 'H8', label: 'Holiday' },
  ];

  return <main className={dark ? 'dark' : ''}>
    <div className="flex h-screen justify-center overscroll-none bg-zinc-200 dark:bg-black">
      <div className="relative flex h-screen w-full max-w-[430px] flex-col overflow-hidden bg-zinc-100 text-zinc-950 dark:bg-black dark:text-white">

        <div
          className={`flex min-h-0 flex-1 flex-col transition-[filter] duration-200 ${sheetOpen ? 'pointer-events-none select-none blur-[1px]' : ''}`}
          aria-hidden={sheetOpen}
        >
          {activeTab === 'calendar' && <div className="flex min-h-0 flex-1 flex-col">
            <header className="shrink-0 px-5 pb-3 pt-6">
              <div className="flex items-start justify-between">
                <div className="min-w-0">
                  <h1 className="truncate text-[34px] font-bold leading-none tracking-tight">{title}</h1>
                  <p className="mt-1 truncate text-[15px] font-medium text-zinc-400 dark:text-zinc-500">{selectedEmployee || 'No employee selected'}</p>
                </div>
                <div className="flex shrink-0 items-center gap-1.5">
                  <button
                    onClick={() => setDark(!dark)}
                    className="flex size-9 items-center justify-center rounded-full bg-zinc-950/5 text-zinc-600 backdrop-blur transition active:scale-[0.97] dark:bg-white/10 dark:text-zinc-300"
                    aria-label="Toggle dark mode"
                  >
                    {dark ? <Sun className="size-4"/> : <Moon className="size-4"/>}
                  </button>
                </div>
              </div>

              <div className="mt-3 flex items-center justify-center gap-1">
                <button
                  onClick={goToPreviousMonth}
                  className="flex size-8 items-center justify-center rounded-full text-zinc-400 transition active:scale-[0.97] active:bg-zinc-950/5 dark:text-zinc-500 dark:active:bg-white/10"
                  aria-label="Previous month"
                >
                  <ChevronLeft className="size-5"/>
                </button>
                <span className="text-[13px] font-semibold text-zinc-400 dark:text-zinc-500">Swipe months</span>
                <button
                  onClick={goToNextMonth}
                  className="flex size-8 items-center justify-center rounded-full text-zinc-400 transition active:scale-[0.97] active:bg-zinc-950/5 dark:text-zinc-500 dark:active:bg-white/10"
                  aria-label="Next month"
                >
                  <ChevronRight className="size-5"/>
                </button>
              </div>
            </header>

            <div className="flex min-h-0 flex-1 flex-col px-3 pb-24">
              <div className="grid shrink-0 grid-cols-7 text-center text-[11px] font-semibold text-zinc-400 dark:text-zinc-500">
                {weekdays.map((day, i) => <div className="py-1.5" key={`${day}-${i}`}>{day}</div>)}
              </div>
              <div className="grid min-h-0 flex-1 grid-cols-7 grid-rows-6 overflow-hidden rounded-[28px] bg-white/70 shadow-[0_1px_2px_rgba(0,0,0,0.04)] backdrop-blur dark:bg-white/[0.04]">
                {calendarDays.map((day) => {
                  const dayIso = iso(day);
                  const event = eventMap[dayIso];
                  const colors = event ? colorFor(event.shift) : { bg: '', text: '' };
                  const isOff = event && shiftKey(event.shift) === 'OFF';
                  const isToday = dayIso === todayIso;
                  return <button
                    key={dayIso}
                    onClick={() => event && setSelectedEvent(event)}
                    className={`flex flex-col items-center justify-start gap-1 border-b border-r border-zinc-950/[0.04] pt-1.5 transition active:scale-[0.97] active:bg-zinc-950/[0.03] dark:border-white/[0.06] dark:active:bg-white/[0.05] ${day.getMonth() !== currentMonth ? 'opacity-30' : ''}`}
                  >
                    <span className={`flex size-8 shrink-0 items-center justify-center rounded-full text-[18px] font-medium transition sm:size-9 sm:text-[19px] ${isToday ? 'bg-red-500 font-semibold text-white' : ''}`}>{day.getDate()}</span>
                    {event ? (
                      isOff ? (
                        <span className="mt-0.5 flex size-1.5 items-center justify-center rounded-full bg-zinc-300 dark:bg-zinc-600"/>
                      ) : (
                        <span className={`mx-1 block max-w-[calc(100%-6px)] truncate rounded-full px-1.5 py-[1px] text-center text-[9px] font-bold leading-tight sm:text-[10px] ${colors.bg} ${colors.text}`}>
                          <span className="hidden sm:inline">{shiftLabel(event.shift)}</span>
                          <span className="sm:hidden">{event.shift}</span>
                        </span>
                      )
                    ) : null}
                  </button>;
                })}
              </div>
            </div>
          </div>}

          {activeTab === 'reports' && <div className="flex min-h-0 flex-1 flex-col overflow-y-auto px-5 pb-24 pt-6">
            <h1 className="text-[34px] font-bold leading-none tracking-tight">Reports</h1>
            <p className="mt-1 text-[15px] font-medium text-zinc-400 dark:text-zinc-500">{selectedEmployee || 'No employee selected'} · {title}</p>

            <div className="mt-5 grid grid-cols-2 gap-3">
              {statCards.map(({ code, label }) => <div key={code} className="rounded-3xl bg-white/70 p-4 backdrop-blur dark:bg-white/[0.04]">
                <p className="text-[13px] font-semibold text-zinc-400 dark:text-zinc-500">{label}</p>
                <p className="mt-1 text-[28px] font-bold tracking-tight">{monthlyStats.counts[code] ?? 0}</p>
              </div>)}
            </div>

            <div className="mt-3 grid grid-cols-2 gap-3">
              <div className="rounded-3xl bg-blue-500/10 p-4">
                <p className="text-[13px] font-semibold text-blue-500">Total Working Shifts</p>
                <p className="mt-1 text-[28px] font-bold tracking-tight text-blue-500">{monthlyStats.workingShifts}</p>
              </div>
              <div className="rounded-3xl bg-green-500/10 p-4">
                <p className="text-[13px] font-semibold text-green-600 dark:text-green-400">Total Hours</p>
                <p className="mt-1 text-[28px] font-bold tracking-tight text-green-600 dark:text-green-400">{monthlyStats.totalHours}h</p>
              </div>
            </div>

            <section className="mt-6">
              <h3 className="px-1 text-[13px] font-semibold uppercase tracking-wide text-zinc-400 dark:text-zinc-500">Monthly Summary</h3>
              <div className="mt-2 divide-y divide-zinc-950/[0.06] overflow-hidden rounded-3xl bg-white/70 backdrop-blur dark:divide-white/[0.06] dark:bg-white/[0.04]">
                <div className="flex items-center justify-between px-4 py-3.5">
                  <span className="text-[15px]">Working days</span>
                  <span className="text-[15px] font-semibold">{monthlyStats.workingShifts}</span>
                </div>
                <div className="flex items-center justify-between px-4 py-3.5">
                  <span className="text-[15px]">Days off</span>
                  <span className="text-[15px] font-semibold">{monthlyStats.daysOff}</span>
                </div>
                <div className="flex items-center justify-between px-4 py-3.5">
                  <span className="text-[15px]">Total hours</span>
                  <span className="text-[15px] font-semibold">{monthlyStats.totalHours}h</span>
                </div>
              </div>
            </section>
          </div>}

          {activeTab === 'settings' && <div className="flex min-h-0 flex-1 flex-col overflow-y-auto px-5 pb-24 pt-6">
            <h1 className="text-[34px] font-bold leading-none tracking-tight">Settings</h1>

            <section className="mt-6">
              <h3 className="px-1 text-[13px] font-semibold uppercase tracking-wide text-zinc-400 dark:text-zinc-500">Employee</h3>
              <div className="mt-2 rounded-3xl bg-white/70 backdrop-blur dark:bg-white/[0.04]">
                <div className="flex items-center gap-2 px-4 pt-3.5">
                  <Search className="size-4 shrink-0 text-zinc-400"/>
                  <input
                    value={query}
                    onChange={(e) => setQuery(e.target.value)}
                    placeholder="Tap to choose employee"
                    className="w-full min-w-0 bg-transparent py-2 text-[15px] outline-none placeholder:text-zinc-400"
                  />
                </div>
                <div className="max-h-64 divide-y divide-zinc-950/[0.06] overflow-y-auto dark:divide-white/[0.06]">
                  {filteredEmployees.map((name) => <button
                    key={name}
                    onClick={() => handleSelectEmployee(name)}
                    className="flex w-full items-center justify-between px-4 py-3 text-left transition active:bg-zinc-950/[0.03] dark:active:bg-white/[0.05]"
                  >
                    <span className="text-[15px] font-medium">{name}</span>
                    {name === selectedEmployee && <Check className="size-4 text-blue-500"/>}
                  </button>)}
                </div>
              </div>
            </section>

            <section className="mt-6">
              <h3 className="px-1 text-[13px] font-semibold uppercase tracking-wide text-zinc-400 dark:text-zinc-500">Appearance</h3>
              <div className="mt-2 rounded-3xl bg-white/70 backdrop-blur dark:bg-white/[0.04]">
                <div className="flex items-center justify-between px-4 py-3.5">
                  <span className="text-[15px]">Dark Mode</span>
                  <button
                    onClick={() => setDark(!dark)}
                    className={`relative h-[30px] w-[50px] rounded-full transition ${dark ? 'bg-green-500' : 'bg-zinc-300'}`}
                    aria-label="Toggle dark mode"
                    aria-pressed={dark}
                  >
                    <span className={`absolute top-0.5 size-[26px] rounded-full bg-white shadow transition-transform ${dark ? 'translate-x-[22px]' : 'translate-x-0.5'}`}/>
                  </button>
                </div>
              </div>
            </section>

            <section className="mt-6">
              <h3 className="px-1 text-[13px] font-semibold uppercase tracking-wide text-zinc-400 dark:text-zinc-500">About</h3>
              <div className="mt-2 divide-y divide-zinc-950/[0.06] overflow-hidden rounded-3xl bg-white/70 backdrop-blur dark:divide-white/[0.06] dark:bg-white/[0.04]">
                <div className="flex items-center justify-between px-4 py-3.5">
                  <span className="text-[15px]">Current month</span>
                  <span className="text-[15px] font-semibold text-zinc-400 dark:text-zinc-500">{title}</span>
                </div>
                <div className="flex items-center justify-between px-4 py-3.5">
                  <span className="text-[15px]">Roster filename</span>
                  <span className="truncate pl-4 text-[15px] font-semibold text-zinc-400 dark:text-zinc-500">{displayFileName(roster.fileName)}</span>
                </div>
                <div className="flex items-center justify-between px-4 py-3.5">
                  <span className="text-[15px]">Number of employees</span>
                  <span className="text-[15px] font-semibold text-zinc-400 dark:text-zinc-500">{roster.employees.length}</span>
                </div>
              </div>
            </section>
          </div>}
        </div>

        <DayDetailsModal
          event={selectedEvent}
          dailyRoster={selectedEvent ? rosterIndex[selectedEvent.isoDate] : undefined}
          selectedEmployee={selectedEmployee}
          canGoPrev={canGoPrevDay}
          canGoNext={canGoNextDay}
          onPrev={goToPrevDay}
          onNext={goToNextDay}
          onClose={() => setSelectedEvent(null)}
        />

        <nav
          aria-hidden={sheetOpen}
          className={`fixed inset-x-0 bottom-0 z-10 flex justify-center px-5 pb-3 transition-[filter] duration-200 ${sheetOpen ? 'pointer-events-none blur-[1px]' : ''}`}
          style={{ paddingBottom: 'max(0.75rem, env(safe-area-inset-bottom))' }}
        >
          <div className="mx-auto flex w-full max-w-[380px] items-stretch gap-1 rounded-[26px] bg-white/70 p-1.5 shadow-[0_8px_30px_rgba(0,0,0,0.15)] backdrop-blur-xl dark:bg-zinc-900/70">
            {tabs.map(({ id, label, Icon }) => <button
              key={id}
              onClick={() => handleTabChange(id)}
              className={`flex flex-1 flex-col items-center gap-0.5 rounded-[20px] py-2 transition active:scale-[0.97] ${activeTab === id ? 'text-blue-500' : 'text-zinc-400 dark:text-zinc-500'}`}
            >
              <Icon className="size-5" strokeWidth={activeTab === id ? 2.5 : 2}/>
              <span className="text-[10px] font-semibold">{label}</span>
            </button>)}
          </div>
        </nav>
      </div>
    </div>
  </main>;
}
