// App.tsx
import type React from 'react';
import { useEffect, useMemo, useRef, useState, useCallback } from 'react';
import { BarChart3, Calendar as CalendarIcon, Check, ChevronLeft, ChevronRight, Moon, Search, Settings2, Sun } from 'lucide-react';
import type { RosterData, ShiftEvent } from './types';
import { colorFor, shiftKey, shiftLabel, shiftHourValues, buildRosterIndex, eventForIso, GLASS_CARD, GLASS_NAV } from './scheduleUtils';
import DayDetailsModal from './DayDetailsModal';

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

const weekdays = ['M', 'T', 'W', 'T', 'F', 'S', 'S'];
const EMPLOYEE_STORAGE_KEY = 'work-schedule-employee';

// Near-black app background — warmer/lighter than pure black so shadows and
// the glass blur stay visible instead of crushing to a void.
const APP_DARK_BG = '#050505';

// Month-swipe gesture tuning — same axis-lock + slide-and-snap technique as
// DayDetailsModal's day swipe, applied to the calendar grid.
const MONTH_AXIS_LOCK_THRESHOLD = 8;
const MONTH_SWIPE_COMMIT_THRESHOLD = 70;
const MONTH_SWIPE_DURATION = 260;

// Sliding indicator offsets for the 3 bottom-nav tabs. These assume exactly 3
// equal-width tabs with p-1.5 container padding and gap-1.5 between them —
// adjust if the tab count or spacing ever changes.
const NAV_INDICATOR_OFFSETS = ['translate-x-0', 'translate-x-[calc(100%+0.375rem)]', 'translate-x-[calc(200%+0.75rem)]'];

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

type LoadStatus = 'loading' | 'error' | 'loaded';
type MonthDragState = { x: number; axis: 'x' | 'y' | null };

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

  // Force the real <html>/<body> to near-black and disable rubber-band
  // scrolling so iOS Safari never reveals a flash of the wrong color behind
  // the app during overscroll.
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
    html.style.backgroundColor = APP_DARK_BG;
    body.style.backgroundColor = APP_DARK_BG;
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
  // Which iso dates the loaded roster actually covers — used to decide which
  // calendar cells are clickable (adjacent-month padding days are excluded).
  const rosterDateSet = useMemo(() => new Set(roster?.dateColumns.map((d) => d.isoDate) ?? []), [roster]);
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
