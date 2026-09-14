// App.tsx
import type React from 'react';
import { useEffect, useMemo, useRef, useState, useCallback } from 'react';
import { BarChart3, Calendar as CalendarIcon, Check, ChevronLeft, ChevronRight, ChevronDown, Moon, Search, Settings2, Sun, Users, Wallet, CheckSquare, Plus, Trash2, Bell, AlertTriangle, Edit2, X, UploadCloud, Loader2, Cloud, ShieldCheck, LogOut, Lock, ArrowLeftRight, Save } from 'lucide-react';
import type { RosterData, ShiftEvent } from './types';
import { parseRoster, mergeRosters } from './parser';
import {
  colorFor,
  shiftKey,
  shiftHourValues,
  buildRosterIndex,
  eventForIso,
  GLASS_CARD,
  GLASS_NAV,
  findShiftTransformationPreview,
  shiftLabel,
  initials,
  isEmployeeOnShift,
  ONLY_ON_SHIFT_NOTIFS_KEY,
} from './scheduleUtils';
import DayDetailsModal from './DayDetailsModal';
import AuthPinModal from './AuthPinModal';
import { AUTH_STORAGE_KEY } from './authConfig';
import { NotificationSettingsCard, NotificationPromptBanner } from './NotificationManager';
import { registerServiceWorker, showAppNotification, NOTIFICATION_PREF_KEY, SHIFT_ALERTS_PREF_KEY } from './notificationService';
import { getIncomingSwapRequests, dismissIncomingSwapRequest, type SwapRequestItem } from './swapNotificationService';
import {
  type Task,
  type ScheduleType,
  type TargetShiftFilter,
  type TimeSelection,
  loadInitialTasks,
  fetchTasksFromCloud,
  saveTasksToCloud,
  TASKS_UPDATED_AT_KEY,
} from './taskService';

const weekdays = ['M', 'T', 'W', 'T', 'F', 'S', 'S'];
const WEEKDAYS_MAP = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
const WEEKDAY_LABELS = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun'];
const WEEKDAY_INDEXES = [1, 2, 3, 4, 5, 6, 0]; 

const EMPLOYEE_STORAGE_KEY = 'work-schedule-employee';
const DARK_MODE_STORAGE_KEY = 'work-schedule-dark-mode';
const SALARY_STORAGE_KEY = 'work-schedule-salaries';
const ADMIN_PIN_STORAGE_KEY = 'work-schedule-admin-pin';
const APP_BASE_URL = ((import.meta as ImportMeta & { env?: { BASE_URL?: string } }).env?.BASE_URL ?? '/');

const ROMANIAN_HOLIDAYS_2026 = new Set<string>([
  '2026-01-01', '2026-01-02', '2026-01-06', '2026-01-07', '2026-01-24',
  '2026-04-10', '2026-04-12', '2026-04-13', '2026-05-01', '2026-05-31',
  '2026-06-01', '2026-08-15', '2026-11-30', '2026-12-01', '2026-12-25', '2026-12-26'
]);

function isNationalHoliday(isoDate: string) {
  return ROMANIAN_HOLIDAYS_2026.has(isoDate);
}

function countWorkingDaysInMonth(month: number, year: number) {
  let count = 0;
  const daysInMonth = new Date(year, month + 1, 0).getDate();
  for (let day = 1; day <= daysInMonth; day++) {
    const date = new Date(year, month, day);
    const dayOfWeek = date.getDay();
    if (dayOfWeek === 0 || dayOfWeek === 6) continue;
    if (isNationalHoliday(iso(date))) continue;
    count += 1;
  }
  return count;
}

function normalizeName(name: string) {
  return name.toLowerCase().replace(/-/g, ' ').split(/\s+/).filter(Boolean).sort().join(' ');
}

const MOST_SEEN_EXCLUDED = [normalizeName('Palade Alexandru-Ionut'), normalizeName('Tir George Cristian')];

const APP_DARK_BG = '#050505';
const APP_LIGHT_BG = '#f4f4f5';

const MONTH_AXIS_LOCK_THRESHOLD = 8;
const MONTH_SWIPE_COMMIT_THRESHOLD = 70;
const MONTH_SWIPE_DURATION = 260;

const IOS_SWITCH_ON = '#34c759';
const IOS_SWITCH_OFF = '#e5e5ea';

function monthDays(month: number, year: number) {
  const first = new Date(year, month, 1);
  const startOffset = (first.getDay() + 6) % 7;
  const start = new Date(year, month, 1 - startOffset);
  return Array.from({ length: 42 }, (_, index) => new Date(start.getFullYear(), start.getMonth(), start.getDate() + index));
}
function iso(date: Date) { return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`; }

function formatSwapDateRange(dateText: string) {
  const dates = dateText.split(', ').filter(Boolean);
  if (dates.length <= 1) return dateText;
  const first = dates[0];
  const last = dates[dates.length - 1];
  const firstParts = first.split('-');
  const lastParts = last.split('-');
  if (firstParts.length === 3 && lastParts.length === 3 && firstParts[0] === lastParts[0] && firstParts[1] === lastParts[1]) {
    return `${first} → ${last}`;
  }
  return `${first} → ${last}`;
}

function formatRequiredSwap(swap: ReturnType<typeof findShiftTransformationPreview>['swaps'][number]) {
  const crossDateParts = swap.isoDate.split(' ↔ ');
  if (crossDateParts.length === 2) {
    const giveDate = new Date(`${crossDateParts[0]}T00:00:00`).toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
    const takeDate = new Date(`${crossDateParts[1]}T00:00:00`).toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
    return `Give your ${shiftLabel(swap.from)} on ${giveDate}; take ${swap.otherEmployee}'s ${shiftLabel(swap.to)} on ${takeDate}.`;
  }
  return `Take ${swap.otherEmployee}'s ${shiftLabel(swap.to)} on ${formatSwapDateRange(swap.isoDate)}.`;
}

function displayFileName(fileName: string | undefined | null) {
  if (!fileName || !fileName.trim()) return 'Untitled roster';
  const base = fileName.trim().split(/[\\/]/).pop() ?? fileName;
  return base.replace(/\.(xlsx|xls|xlsm|csv|json)$/i, '');
}

function formatRon(amount: number) {
  return `${amount.toLocaleString('ro-RO', { minimumFractionDigits: 2, maximumFractionDigits: 2 })} RON`;
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
    dateColumns: json.dateColumns.map(({ isoDate }, index) => ({ index, isoDate, date: new Date(`${isoDate}T00:00:00`) })),
    rows: json.rows,
  };
}

function eventsForEmployee(roster: RosterData, employee: string): ShiftEvent[] {
  return roster.dateColumns.reduce<ShiftEvent[]>((events, { isoDate }) => {
    const shift = shiftKey(roster.rows[employee]?.[isoDate] ?? 'OFF');
    events.push({ id: `${employee}-${isoDate}`, isoDate, shift, date: new Date(`${isoDate}T00:00:00`) });
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
  const [authenticatedEmployee, setAuthenticatedEmployee] = useState<string | null>(() => localStorage.getItem(AUTH_STORAGE_KEY));
  const [authModalTarget, setAuthModalTarget] = useState<string | null>(null);
  const [isEmployeeDialogOpen, setIsEmployeeDialogOpen] = useState(false);
  const [incomingSwapRequests, setIncomingSwapRequests] = useState<SwapRequestItem[]>([]);
  const seenSwapIdsRef = useRef<Set<string>>(new Set());
  const [query, setQuery] = useState('');
  const [currentMonth, setCurrentMonth] = useState(new Date().getMonth());
  const [currentYear, setCurrentYear] = useState(new Date().getFullYear());
  const [selectedEvent, setSelectedEvent] = useState<ShiftEvent | null>(null);
  const [salaries, setSalaries] = useState<Record<string, number>>(() => {
    try {
      const saved = localStorage.getItem(SALARY_STORAGE_KEY);
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });

  // Persistent Site Storage for Recurring Tasks (Cloudflare KV + Local Cache)
  const [tasks, setTasks] = useState<Task[]>(() => loadInitialTasks() || []);
  const [isSyncingTasks, setIsSyncingTasks] = useState(false);
  const [tasksLastUpdated, setTasksLastUpdated] = useState<string>(() => localStorage.getItem(TASKS_UPDATED_AT_KEY) || '');
  
  const [editingTaskId, setEditingTaskId] = useState<string | null>(null);
  const editingTaskIdRef = useRef<string | null>(null);
  editingTaskIdRef.current = editingTaskId;

  const [newTaskTitle, setNewTaskTitle] = useState('');
  const [scheduleType, setScheduleType] = useState<ScheduleType>('daily');
  const [targetShift, setTargetShift] = useState<TargetShiftFilter>('ALL_ACTIVE');
  const [selectedDays, setSelectedDays] = useState<number[]>([]);
  const [newTaskTimes, setNewTaskTimes] = useState<TimeSelection[]>([{ hour: '00', minute: '00' }]);
  
  const [activeAlarmTask, setActiveAlarmTask] = useState<{ id: string; taskTitle: string; time: string; type: ScheduleType } | null>(null);
  const audioRef = useRef<HTMLAudioElement | null>(null);
  const lastCheckedMinute = useRef<string>('');

  const hoursArray = useMemo(() => Array.from({ length: 24 }, (_, i) => String(i).padStart(2, '0')), []);
  const minutesArray = useMemo(() => Array.from({ length: 60 }, (_, i) => String(i).padStart(2, '0')), []);

  const handleSalaryChange = useCallback((employee: string, value: string) => {
    const parsed = value === '' ? 0 : Number(value);
    setSalaries((prev) => {
      const next = { ...prev, [employee]: Number.isFinite(parsed) ? parsed : 0 };
      localStorage.setItem(SALARY_STORAGE_KEY, JSON.stringify(next));
      return next;
    });
  }, []);

  // Hydrate and sync recurring tasks with Cloudflare KV site storage
  useEffect(() => {
    let isMounted = true;

    const syncTasks = async (isBackground = false) => {
      if (!isBackground) setIsSyncingTasks(true);
      try {
        const res = await fetchTasksFromCloud();
        if (!isMounted) return;
        if (res.tasks && res.tasks.length > 0) {
          // Do not overwrite form if user is actively in the middle of editing a task
          if (!editingTaskIdRef.current) {
            setTasks(res.tasks);
            if (res.updatedAt) setTasksLastUpdated(res.updatedAt);
          }
        }
      } finally {
        if (isMounted && !isBackground) setIsSyncingTasks(false);
      }
    };

    // Initial load from Cloudflare KV
    syncTasks(false);

    // Periodic real-time background sync every 15 seconds so changes on any device appear immediately
    const interval = setInterval(() => {
      syncTasks(true);
    }, 15000);

    // Sync when returning to tab/app (e.g. mobile unlock or tab switch)
    const handleVisibility = () => {
      if (document.visibilityState === 'visible') {
        syncTasks(true);
      }
    };
    window.addEventListener('visibilitychange', handleVisibility);
    window.addEventListener('focus', handleVisibility);

    return () => {
      isMounted = false;
      clearInterval(interval);
      window.removeEventListener('visibilitychange', handleVisibility);
      window.removeEventListener('focus', handleVisibility);
    };
  }, []);

  // Cloudflare KV Roster Upload & Sync State
  const [scheduleSource, setScheduleSource] = useState<'kv' | 'static'>('static');
  const [adminPin, setAdminPin] = useState(() => localStorage.getItem(ADMIN_PIN_STORAGE_KEY) ?? '');
  const [uploadFile, setUploadFile] = useState<File | null>(null);
  const [parsedPreview, setParsedPreview] = useState<RosterData | null>(null);
  const [mergeWithExisting, setMergeWithExisting] = useState(true);
  const [isUploading, setIsUploading] = useState(false);
  const [uploadStatus, setUploadStatus] = useState<{ type: 'success' | 'error'; message: string } | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileSelect = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setUploadFile(file);
    setUploadStatus(null);
    try {
      const parsed = await parseRoster(file);
      setParsedPreview(parsed);
    } catch (err) {
      setUploadStatus({
        type: 'error',
        message: err instanceof Error ? err.message : 'Failed to parse Excel schedule.',
      });
      setParsedPreview(null);
    }
  };

  const handlePublishRoster = async () => {
    if (!parsedPreview) return;
    setIsUploading(true);
    setUploadStatus(null);

    try {
      const payloadRoster = mergeWithExisting && roster
        ? mergeRosters(roster, parsedPreview)
        : parsedPreview;

      if (adminPin.trim()) {
        localStorage.setItem(ADMIN_PIN_STORAGE_KEY, adminPin.trim());
      } else {
        localStorage.removeItem(ADMIN_PIN_STORAGE_KEY);
      }

      const headers: Record<string, string> = {
        'Content-Type': 'application/json',
      };
      if (adminPin.trim()) {
        headers['X-Admin-Pin'] = adminPin.trim();
      }

      const res = await fetch('/api/schedule', {
        method: 'POST',
        headers,
        body: JSON.stringify(payloadRoster),
      });

      if (!res.ok) {
        const errorData = await res.json().catch(() => ({}));
        throw new Error(errorData.error || `Server responded with status ${res.status}`);
      }

      setRoster(payloadRoster);
      setScheduleSource('kv');
      setCurrentMonth(parsedPreview.month);
      setCurrentYear(parsedPreview.year);
      setUploadStatus({
        type: 'success',
        message: `Successfully published ${parsedPreview.fileName} to Cloudflare KV!`,
      });
      setUploadFile(null);
      setParsedPreview(null);
      if (fileInputRef.current) fileInputRef.current.value = '';
    } catch (err) {
      setUploadStatus({
        type: 'error',
        message: err instanceof Error ? err.message : 'Failed to publish schedule to Cloudflare KV.',
      });
    } finally {
      setIsUploading(false);
    }
  };

  useEffect(() => {
    registerServiceWorker().catch(() => {});
  }, []);

  // Poll Cloudflare KV for incoming shift swap requests directed to this employee
  useEffect(() => {
    if (!authenticatedEmployee) {
      setIncomingSwapRequests([]);
      return;
    }

    const checkIncomingSwaps = async () => {
      try {
        const requests = await getIncomingSwapRequests(authenticatedEmployee);
        setIncomingSwapRequests(requests);

        // Notify user about newly arrived pending swap requests
        for (const req of requests) {
          if (req.status === 'pending' && !seenSwapIdsRef.current.has(req.id)) {
            seenSwapIdsRef.current.add(req.id);
            showAppNotification(`🔄 Shift Swap Request from ${req.fromEmployee}`, {
              body: `${req.fromEmployee} requested to swap shifts with you for ${req.date}.`,
              tag: `swap-notification-${req.id}`,
            }).catch(() => {});
          }
        }
      } catch {
        // Non-blocking
      }
    };

    checkIncomingSwaps();
    const timer = setInterval(checkIncomingSwaps, 12000);
    return () => clearInterval(timer);
  }, [authenticatedEmployee]);

  useEffect(() => {
    audioRef.current = new Audio('/alarm.mp3');
    audioRef.current.loop = true;

    const timer = setInterval(() => {
      const now = new Date();
      const currentHHMM = `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
      const currentDayOfWeek = now.getDay();
      const currentIsoDate = iso(now);
      
      if (lastCheckedMinute.current === currentHHMM) return;
      lastCheckedMinute.current = currentHHMM;

      const activeEmployee = authenticatedEmployee || selectedEmployee;
      const taskAlertsEnabled = localStorage.getItem(NOTIFICATION_PREF_KEY) !== 'false';
      const onlyOnShiftEnabled = localStorage.getItem(ONLY_ON_SHIFT_NOTIFS_KEY) !== 'false';

      // Check recurring task alarms
      if (taskAlertsEnabled) {
        tasks.forEach((task) => {
          if (!task.times.includes(currentHHMM)) return;

          let shouldTrigger = false;

          if (task.scheduleType === 'daily') {
            shouldTrigger = true;
          } else if (task.scheduleType === 'once' && task.dateCreated === currentIsoDate) {
            shouldTrigger = true;
          } else if (task.scheduleType === 'weekly' && task.daysOfWeek?.includes(currentDayOfWeek)) {
            shouldTrigger = true;
          }

          if (!shouldTrigger) return;

          // Shift filtering: Only notify colleagues who are actively on shift
          if (onlyOnShiftEnabled && roster && activeEmployee) {
            const shiftStatus = isEmployeeOnShift(roster, activeEmployee, now);
            // If employee is scheduled OFF, H8, or outside shift hours, suppress alert
            if (!shiftStatus.onShift) {
              return;
            }
            // If task is targeted to a specific shift and colleague is working a different shift, suppress alert
            if (task.targetShift && task.targetShift !== 'ALL_ACTIVE' && shiftStatus.currentShift !== task.targetShift) {
              return;
            }
          }

          setActiveAlarmTask({ id: task.id, taskTitle: task.title, time: currentHHMM, type: task.scheduleType });
          if (audioRef.current) {
            audioRef.current.currentTime = 0;
            audioRef.current.play().catch(err => console.log("Audio deferred configuration:", err));
          }
          showAppNotification(`⏰ Alarm: ${task.title}`, {
            body: `Scheduled alarm triggered at ${currentHHMM}`,
            tag: `alarm-${task.id}`,
          }).catch(() => {});
        });
      }

      // Check upcoming shift alerts (1 hour before shift starts: M at 06:00, MID at 09:00, A at 14:00, N at 22:00)
      const shiftAlertsEnabled = localStorage.getItem(SHIFT_ALERTS_PREF_KEY) !== 'false';
      if (shiftAlertsEnabled && roster && activeEmployee) {
        const todayShift = shiftKey(roster.rows[activeEmployee]?.[currentIsoDate] ?? 'OFF');
        if (todayShift === 'M' && currentHHMM === '05:00') {
          showAppNotification('Shift Reminder: Morning (M)', {
            body: 'Your Morning shift starts in 1 hour at 06:00.',
            tag: `shift-${currentIsoDate}-M`,
          }).catch(() => {});
        } else if (todayShift === 'MID' && currentHHMM === '08:00') {
          showAppNotification('Shift Reminder: Mid (MID)', {
            body: 'Your Mid shift starts in 1 hour at 09:00.',
            tag: `shift-${currentIsoDate}-MID`,
          }).catch(() => {});
        } else if (todayShift === 'A' && currentHHMM === '13:00') {
          showAppNotification('Shift Reminder: Afternoon (A)', {
            body: 'Your Afternoon shift starts in 1 hour at 14:00.',
            tag: `shift-${currentIsoDate}-A`,
          }).catch(() => {});
        } else if (todayShift === 'N' && currentHHMM === '21:00') {
          showAppNotification('Shift Reminder: Night (N)', {
            body: 'Your Night shift starts in 1 hour at 22:00.',
            tag: `shift-${currentIsoDate}-N`,
          }).catch(() => {});
        }
      }
    }, 1000);

    return () => {
      clearInterval(timer);
      if (audioRef.current) {
        audioRef.current.pause();
        audioRef.current = null;
      }
    };
  }, [tasks, roster, selectedEmployee, authenticatedEmployee]);

  const dismissAlarm = useCallback(() => {
    if (audioRef.current) {
      audioRef.current.pause();
      audioRef.current.currentTime = 0;
    }
    
    if (activeAlarmTask && activeAlarmTask.type === 'once') {
      const nextStore = tasks.filter((t) => t.id !== activeAlarmTask.id);
      setTasks(nextStore);
      saveTasksToCloud(nextStore, authenticatedEmployee || selectedEmployee).catch(() => {});
    }
    
    setActiveAlarmTask(null);
  }, [activeAlarmTask, tasks, authenticatedEmployee, selectedEmployee]);

  const handleSaveTask = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newTaskTitle.trim()) return;
    if (scheduleType === 'weekly' && selectedDays.length === 0) return;

    const generatedTimes = newTaskTimes.map(t => `${t.hour}:${t.minute}`);
    const finalTimes = Array.from(new Set(generatedTimes)).sort();
    
    let nextStore: Task[] = [];

    if (editingTaskId) {
      nextStore = tasks.map((t) => 
        t.id === editingTaskId 
          ? { 
              ...t, 
              title: newTaskTitle.trim(), 
              times: finalTimes, 
              scheduleType, 
              daysOfWeek: scheduleType === 'weekly' ? [...selectedDays].sort() : undefined,
              targetShift,
            }
          : t
      );
    } else {
      const newTask: Task = {
        id: crypto.randomUUID(),
        title: newTaskTitle.trim(),
        times: finalTimes,
        scheduleType,
        daysOfWeek: scheduleType === 'weekly' ? [...selectedDays].sort() : undefined,
        dateCreated: iso(new Date()),
        targetShift,
      };
      nextStore = [...tasks, newTask];
    }

    setTasks(nextStore);
    resetForm();
    setIsSyncingTasks(true);
    try {
      const res = await saveTasksToCloud(nextStore, authenticatedEmployee || selectedEmployee);
      if (res.updatedAt) setTasksLastUpdated(res.updatedAt);
    } finally {
      setIsSyncingTasks(false);
    }
  };

  const handleRemoveTask = async (id: string) => {
    const nextStore = tasks.filter((t) => t.id !== id);
    setTasks(nextStore);
    if (editingTaskId === id) resetForm();
    setIsSyncingTasks(true);
    try {
      const res = await saveTasksToCloud(nextStore, authenticatedEmployee || selectedEmployee);
      if (res.updatedAt) setTasksLastUpdated(res.updatedAt);
    } finally {
      setIsSyncingTasks(false);
    }
  };

  const startEditingTask = (task: Task) => {
    setEditingTaskId(task.id);
    setNewTaskTitle(task.title);
    setScheduleType(task.scheduleType);
    setTargetShift(task.targetShift || 'ALL_ACTIVE');
    setSelectedDays(task.daysOfWeek || []);
    
    const splitTimes = task.times.map(t => {
      const parts = t.split(':');
      return { hour: parts[0] || '00', minute: parts[1] || '00' };
    });
    setNewTaskTimes(splitTimes);
  };

  const resetForm = () => {
    setEditingTaskId(null);
    setNewTaskTitle('');
    setScheduleType('daily');
    setTargetShift('ALL_ACTIVE');
    setSelectedDays([]);
    setNewTaskTimes([{ hour: '00', minute: '00' }]);
  };

  const toggleDaySelection = (dayIndex: number) => {
    setSelectedDays((prev) => 
      prev.includes(dayIndex) ? prev.filter(d => d !== dayIndex) : [...prev, dayIndex]
    );
  };

  const handleTimeDropdownChange = (index: number, field: 'hour' | 'minute', value: string) => {
    setNewTaskTimes((prev) => {
      const next = [...prev];
      next[index] = { ...next[index], [field]: value };
      return next;
    });
  };

  const addTimeInputField = () => {
    setNewTaskTimes((prev) => [...prev, { hour: '00', minute: '00' }]);
  };

  const removeTimeInputField = (index: number) => {
    if (newTaskTimes.length <= 1) return;
    setNewTaskTimes((prev) => prev.filter((_, i) => i !== index));
  };

  useEffect(() => {
    let mounted = true;

    const fetchSchedule = async () => {
      // 1. Try Cloudflare Pages Function connected to KV
      try {
        const kvRes = await fetch('/api/schedule');
        if (kvRes.ok) {
          const json = (await kvRes.json()) as ScheduleJson;
          if (!mounted) return;
          const parsed = hydrateRoster(json);
          setRoster(parsed);
          setScheduleSource('kv');
          const savedAuth = localStorage.getItem(AUTH_STORAGE_KEY);
          const savedEmployee = localStorage.getItem(EMPLOYEE_STORAGE_KEY);
          const initialEmployee = savedAuth && parsed.employees.includes(savedAuth)
            ? savedAuth
            : savedEmployee && parsed.employees.includes(savedEmployee)
              ? savedEmployee
              : (parsed.employees[0] ?? '');
          setSelectedEmployee(initialEmployee);
          if (savedAuth && parsed.employees.includes(savedAuth)) {
            setAuthenticatedEmployee(savedAuth);
          }
          setStatus('loaded');
          return;
        }
      } catch {
        // Fall back below
      }

      // 2. Fall back to bundled static JSON
      try {
        const res = await fetch(`${APP_BASE_URL}schedule.json`);
        if (!res.ok) throw new Error('Schedule structural asset missing from schedule.json.');
        const json = (await res.json()) as ScheduleJson;
        if (!mounted) return;
        const parsed = hydrateRoster(json);
        setRoster(parsed);
        setScheduleSource('static');
        const savedAuth = localStorage.getItem(AUTH_STORAGE_KEY);
        const savedEmployee = localStorage.getItem(EMPLOYEE_STORAGE_KEY);
        const initialEmployee = savedAuth && parsed.employees.includes(savedAuth)
          ? savedAuth
          : savedEmployee && parsed.employees.includes(savedEmployee)
            ? savedEmployee
            : (parsed.employees[0] ?? '');
        setSelectedEmployee(initialEmployee);
        if (savedAuth && parsed.employees.includes(savedAuth)) {
          setAuthenticatedEmployee(savedAuth);
        }
        setStatus('loaded');
      } catch (err) {
        if (!mounted) return;
        setErrorMessage(err instanceof Error ? err.message : 'Unable to query backend schema.');
        setStatus('error');
      }
    };

    fetchSchedule();
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

  const [activeTab, setActiveTab] = useState<'calendar' | 'reports' | 'tasks' | 'settings' | 'solver'>('calendar');
  const [solverTargetShift, setSolverTargetShift] = useState<string>('M');
  const [solverResult, setSolverResult] = useState<ReturnType<typeof findShiftTransformationPreview> | null>(null);
  const [solverSimulationVariant, setSolverSimulationVariant] = useState(0);
  const handleTabChange = useCallback((tab: 'calendar' | 'reports' | 'tasks' | 'settings' | 'solver') => {
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
  const isAlexandruStoian = useMemo(() => {
    const current = authenticatedEmployee || selectedEmployee;
    if (!current) return false;
    return (
      current === 'Stoian Alexandru-Gabriel' ||
      current === 'Alexandru Stoian' ||
      (/stoian/i.test(current) && /alexandru/i.test(current))
    );
  }, [authenticatedEmployee, selectedEmployee]);
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

    let weekendDays = 0;
    monthDatesInRoster.forEach(({ isoDate, date }) => {
      const rawShift = roster?.rows[selectedEmployee]?.[isoDate];
      const code = shiftKey(rawShift ?? 'OFF');
      if (code in counts) {
        counts[code] += 1;
      } else {
        counts['OFF'] += 1;
      }
      const isWeekendDate = date.getDay() === 0 || date.getDay() === 6;
      const isWorking = code !== 'OFF' && code !== 'H8';
      if (isWeekendDate && isWorking) weekendDays += 1;
    });

    const workingShifts = counts.M + counts.A + counts.N + counts.MID;
    const totalHours = Object.entries(counts).reduce((sum, [code, count]) => sum + count * (shiftHourValues[code] ?? 0), 0);
    const daysOff = counts.OFF + counts.H8;
    return { counts, workingShifts, totalHours, daysOff, weekendDays };
  }, [roster, selectedEmployee, currentMonth, currentYear]);

  const salaryBreakdown = useMemo(() => {
    const baseSalary = salaries[selectedEmployee] ?? 0;
    const workingDaysInMonth = countWorkingDaysInMonth(currentMonth, currentYear);
    const dailyWage = workingDaysInMonth > 0 ? baseSalary / workingDaysInMonth : 0;

    const monthDatesInRoster = roster ? roster.dateColumns.filter(
      (d) => d.date.getMonth() === currentMonth && d.date.getFullYear() === currentYear
    ) : [];

    let workedDays = 0;
    let nightCount = 0, weekendCount = 0, holidayCount = 0;
    let nightBonus = 0, weekendBonus = 0, holidayBonus = 0;
    let basePay = 0;

    monthDatesInRoster.forEach(({ isoDate, date }) => {
      const raw = roster?.rows[selectedEmployee]?.[isoDate];
      const code = shiftKey(raw ?? 'OFF');
      
      const isPaidDay = code !== 'OFF';
      if (!isPaidDay) return;

      workedDays += 1;
      basePay += dailyWage;

      if (code === 'H8') return;

      const isNight = code === 'N';
      const isWeekend = date.getDay() === 0 || date.getDay() === 6;
      const isHoliday = isNationalHoliday(isoDate);

      if (isNight) { nightCount += 1; nightBonus += dailyWage * 0.25; }
      if (isWeekend) { weekendCount += 1; weekendBonus += dailyWage * 0.10; }
      if (isHoliday) { holidayCount += 1; holidayBonus += dailyWage * 1.00; }
    });

    const total = basePay + nightBonus + weekendBonus + holidayBonus;
    return {
      baseSalary, workingDaysInMonth, dailyWage, workedDays, basePay,
      nightCount, weekendCount, holidayCount,
      nightBonus, weekendBonus, holidayBonus, total,
    };
  }, [roster, salaries, selectedEmployee, currentMonth, currentYear]);

  const mostSeenColleagues = useMemo(() => {
    if (!roster || !selectedEmployee) return { names: [], count: 0 };
    const matches: Record<string, number> = {};
    const targetMonthEvents = currentMonthEvents.filter(e => shiftKey(e.shift) !== 'OFF' && shiftKey(e.shift) !== 'H8');

    targetMonthEvents.forEach((userEvent) => {
      const activeUserShiftCode = shiftKey(userEvent.shift);
      roster.employees.forEach((colleague) => {
        if (colleague === selectedEmployee) return;
        if (MOST_SEEN_EXCLUDED.includes(normalizeName(colleague))) return;
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
    if (authenticatedEmployee === name) {
      setSelectedEmployee(name);
      localStorage.setItem(EMPLOYEE_STORAGE_KEY, name);
      return;
    }
    setAuthModalTarget(name);
  }, [authenticatedEmployee]);

  const handleAuthSuccess = useCallback((employee: string) => {
    setAuthenticatedEmployee(employee);
    setSelectedEmployee(employee);
    localStorage.setItem(AUTH_STORAGE_KEY, employee);
    localStorage.setItem(EMPLOYEE_STORAGE_KEY, employee);
  }, []);

  const handleLogOut = useCallback(() => {
    setAuthenticatedEmployee(null);
    localStorage.removeItem(AUTH_STORAGE_KEY);
  }, []);

  if (status === 'loading') {
    return <main className={dark ? 'dark' : ''}><div className="flex h-screen items-center justify-center overscroll-none bg-zinc-100 dark:bg-[#050505] text-zinc-400">Loading schedule…</div></main>;
  }

  if (status === 'error' || !roster) {
    return <main className={dark ? 'dark' : ''}><div className="flex h-screen items-center justify-center overscroll-none p-8 text-red-500 dark:bg-[#050505]">{errorMessage}</div></main>;
  }

  const sheetOpen = Boolean(selectedEvent) || Boolean(activeAlarmTask);
  const tabs = [
    { id: 'calendar' as const, label: 'Calendar', Icon: CalendarIcon }, 
    { id: 'reports' as const, label: 'Reports', Icon: BarChart3 }, 
    { id: 'solver' as const, label: 'Solver', Icon: Search },
    { id: 'tasks' as const, label: 'Tasks', Icon: CheckSquare },
    { id: 'settings' as const, label: 'Settings', Icon: Settings2 }
  ];
  const activeTabIndex = tabs.findIndex((t) => t.id === activeTab);
  
  const statCards = [
    { code: 'M', label: 'Morning' }, 
    { code: 'A', label: 'Afternoon' }, 
    { code: 'N', label: 'Night' }, 
    { code: 'MID', label: 'Mid' }, 
    { code: 'OFF', label: 'Off Days' }, 
    { code: 'H8', label: 'Holiday' }
  ];

  return (
    <main className={dark ? 'dark' : ''}>
      <div className="flex h-screen justify-center overscroll-none bg-zinc-200 dark:bg-[#050505]">
        <div className="relative flex h-screen w-full max-w-[430px] flex-col overflow-hidden bg-zinc-100 text-zinc-950 dark:bg-[#050505] dark:text-white lg:max-w-none lg:flex-row">

          {/* Desktop Nav Layout */}
          <nav className="hidden shrink-0 flex-col border-r border-zinc-950/[0.06] px-3 py-8 dark:border-white/[0.06] lg:flex lg:w-64">
            <p className="truncate px-3 text-[15px] font-bold tracking-tight">{title}</p>
            <p className="mt-0.5 truncate px-3 pb-6 text-[13px] font-medium text-zinc-400 dark:text-zinc-500">{selectedEmployee}</p>
            <div className="flex flex-col gap-1">
              {tabs.map(({ id, label, Icon }) => (
                <button
                  key={id}
                  onClick={() => handleTabChange(id)}
                  className={`flex items-center gap-3 rounded-2xl px-3 py-2.5 text-left text-[14px] font-semibold transition-colors ${activeTab === id ? 'bg-blue-500/10 text-blue-500' : 'text-zinc-500 hover:bg-zinc-950/[0.04] dark:text-zinc-400 dark:hover:bg-white/[0.06]'}`}
                >
                  <Icon className="size-5" />
                  {label}
                </button>
              ))}
            </div>
          </nav>

          <div className="flex min-h-0 flex-1 flex-col lg:overflow-y-auto">
          <div className={`flex min-h-0 flex-1 flex-col transition-[filter] duration-200 ${sheetOpen ? 'pointer-events-none select-none blur-[1px]' : ''}`} aria-hidden={sheetOpen}>
            <div className="mx-auto flex min-h-0 w-full flex-1 flex-col lg:max-w-4xl lg:px-10 lg:py-8">
            {incomingSwapRequests.length > 0 && (
              <div className="mx-4 lg:mx-0 mb-3 space-y-2">
                {incomingSwapRequests.map((req) => (
                  <div
                    key={req.id}
                    className="flex items-center justify-between gap-3 rounded-2xl border border-cyan-500/30 bg-cyan-500/10 p-3 text-cyan-600 dark:text-cyan-400 backdrop-blur-md animate-fadeIn"
                  >
                    <div className="flex items-center gap-2.5 min-w-0">
                      <div className="flex size-8 shrink-0 items-center justify-center rounded-xl bg-cyan-500/20">
                        <ArrowLeftRight className="size-4" />
                      </div>
                      <div className="min-w-0">
                        <p className="text-[13px] font-bold truncate">
                          Swap Request from {req.fromEmployee}
                        </p>
                        <p className="text-[11px] opacity-90 truncate">
                          Date: <strong>{req.date}</strong> ({shiftLabel(req.requesterShift)} ↔ {shiftLabel(req.candidateShift)})
                        </p>
                      </div>
                    </div>
                    <div className="flex items-center gap-1.5 shrink-0">
                      <button
                        type="button"
                        onClick={() => {
                          if (roster && selectedEmployee) {
                            const ev = eventForIso(roster, selectedEmployee, req.date);
                            setSelectedEvent(ev);
                          }
                        }}
                        className="rounded-xl bg-cyan-600 px-2.5 py-1.5 text-[11px] font-bold text-white shadow-sm hover:bg-cyan-500 transition-all active:scale-95 cursor-pointer"
                      >
                        View Day
                      </button>
                      <button
                        type="button"
                        onClick={async () => {
                          if (authenticatedEmployee) {
                            await dismissIncomingSwapRequest(authenticatedEmployee, req.id);
                            setIncomingSwapRequests((prev) => prev.filter((r) => r.id !== req.id));
                          }
                        }}
                        className="rounded-xl bg-zinc-950/5 p-1.5 text-zinc-400 hover:text-zinc-600 dark:bg-white/5 dark:hover:text-zinc-200 transition-all cursor-pointer"
                        title="Dismiss request"
                      >
                        <X className="size-4" />
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {activeTab === 'calendar' && (
              <div className="flex min-h-0 flex-1 flex-col">
                <header className="shrink-0 px-5 pb-3 pt-6 lg:hidden">
                  <div className="flex items-start justify-between">
                    <div className="min-w-0">
                      <h1 className="truncate text-[34px] font-bold leading-none tracking-tight">{title}</h1>
                      <p className="mt-1 truncate text-[15px] font-medium text-zinc-400 dark:text-zinc-500">{selectedEmployee}</p>
                    </div>
                    <button
                      onClick={() => handleTabChange('settings')}
                      className="flex size-10 items-center justify-center rounded-2xl bg-zinc-950/5 text-zinc-600 transition-colors hover:bg-zinc-950/10 dark:bg-white/5 dark:text-zinc-300 dark:hover:bg-white/10"
                      title="Settings & Notifications"
                    >
                      <Bell className="size-5" />
                    </button>
                  </div>
                  <div className="mt-3 flex items-center justify-center gap-1">
                    <button onClick={goToPreviousMonth} className="flex size-8 items-center justify-center rounded-full text-zinc-400"><ChevronLeft className="size-5"/></button>
                    <span className="text-[13px] font-semibold text-zinc-400">Swipe months</span>
                    <button onClick={goToNextMonth} className="flex size-8 items-center justify-center rounded-full text-zinc-400"><ChevronRight className="size-5"/></button>
                  </div>
                </header>

                <div className="hidden shrink-0 items-center justify-between pb-4 lg:flex">
                  <button onClick={goToPreviousMonth} className="flex size-9 items-center justify-center rounded-full text-zinc-400 hover:bg-zinc-950/[0.04] dark:hover:bg-white/[0.06]"><ChevronLeft className="size-5"/></button>
                  <h2 className="text-[20px] font-bold tracking-tight">{title}</h2>
                  <button onClick={goToNextMonth} className="flex size-9 items-center justify-center rounded-full text-zinc-400 hover:bg-zinc-950/[0.04] dark:hover:bg-white/[0.06]"><ChevronRight className="size-5"/></button>
                </div>

                <div className="flex min-h-0 flex-1 flex-col px-3 pb-24 lg:px-0 lg:pb-4">
                  <div className="grid shrink-0 grid-cols-7 text-center text-[11px] font-semibold text-zinc-400">
                    {weekdays.map((day, i) => <div className="py-1.5" key={`${day}-${i}`}>{day}</div>)}
                  </div>
                  <div className={`relative min-h-0 flex-1 overflow-hidden ${GLASS_CARD}`}>
                    <div ref={monthGridRef} onPointerDown={handleMonthPointerDown} style={{ touchAction: 'pan-y' }} className="grid h-full grid-cols-7 grid-rows-6">
                      {calendarDays.map((day) => {
                        const dayIso = iso(day); const inRoster = rosterDateSet.has(dayIso);
                        const event = eventMap[dayIso] ?? (inRoster && roster ? { id: `${selectedEmployee}-${dayIso}-synthetic`, isoDate: dayIso, date: day, shift: 'OFF' } : undefined);
                        const colors = event ? colorFor(event.shift) : { bg: '', text: '' };
                        const isOff = event && shiftKey(event.shift) === 'OFF'; const isToday = dayIso === todayIso;
                        const isHoliday = isNationalHoliday(dayIso);
                        return <button
                          key={dayIso} onClick={() => inRoster && event && setSelectedEvent(event)} disabled={!inRoster}
                          className={`flex flex-col items-center justify-start gap-1 border-b border-r border-zinc-950/[0.04] pt-1.5 dark:border-white/[0.06] ${inRoster ? 'active:scale-[0.96]' : ''} ${day.getMonth() !== currentMonth ? 'opacity-30' : ''} ${isHoliday ? 'bg-red-500/10' : ''}`}
                        >
                          <span className={`flex size-8 items-center justify-center rounded-full text-[18px] ${isToday ? 'bg-red-500 text-white font-bold' : isHoliday ? 'font-extrabold text-red-500 ring-2 ring-red-500' : ''}`}>{day.getDate()}</span>
                          {event ? (isOff ? <span className="mt-0.5 size-1.5 rounded-full bg-zinc-300 dark:bg-zinc-600"/> : <span className={`mx-1 truncate rounded-full px-1.5 py-[1px] text-[9px] font-bold ${colors.bg} ${colors.text}`}>{event.shift}</span>) : null}
                        </button>;
                      })}
                    </div>
                  </div>
                </div>
              </div>
            )}

            {activeTab === 'solver' && (
              <div className="flex min-h-0 w-full max-w-full flex-1 flex-col overflow-y-auto overflow-x-hidden px-4 pb-24 pt-6 sm:px-5 lg:px-0 lg:pb-8">
                <div className="mb-4 flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                  <div className="min-w-0">
                    <h1 className="text-[22px] font-bold tracking-tight sm:text-[30px]">Shift transformation</h1>
                    <p className="text-[12px] text-zinc-400 sm:text-[13px] truncate">Build a month-level swap preview for {selectedEmployee}</p>
                  </div>
                  <div className="flex items-center gap-2 shrink-0">
                    <button
                      onClick={() => {
                        if (!roster || !selectedEmployee) return;
                        setSolverSimulationVariant(0);
                        setSolverResult(findShiftTransformationPreview(roster, selectedEmployee, solverTargetShift, currentMonth, currentYear, 0));
                      }}
                      className="flex-1 sm:flex-none rounded-xl bg-blue-500 px-3 py-2 text-[12px] font-bold uppercase tracking-wider text-white text-center transition-all active:scale-95 cursor-pointer"
                    >
                      Run solver
                    </button>
                    <button
                      onClick={() => {
                        if (!roster || !selectedEmployee) return;
                        const nextVariant = solverSimulationVariant + 1;
                        setSolverSimulationVariant(nextVariant);
                        setSolverResult(findShiftTransformationPreview(roster, selectedEmployee, solverTargetShift, currentMonth, currentYear, nextVariant));
                      }}
                      disabled={!solverResult}
                      className="flex-1 sm:flex-none rounded-xl border border-white/15 px-3 py-2 text-[12px] font-bold uppercase tracking-wider text-zinc-200 disabled:cursor-not-allowed disabled:opacity-40 text-center transition-all active:scale-95 cursor-pointer"
                    >
                      Resimulate change
                    </button>
                  </div>
                </div>

                <div className="mb-4 grid grid-cols-2 gap-2.5 sm:gap-3 min-w-0">
                  <div className={`min-w-0 rounded-2xl border border-white/10 bg-zinc-900/40 p-3 ${GLASS_CARD}`}>
                    <label className="block text-[10px] font-bold uppercase tracking-wider text-zinc-400">Employee</label>
                    <div className="mt-2 text-[14px] sm:text-[16px] font-bold truncate" title={selectedEmployee}>{selectedEmployee}</div>
                  </div>
                  <div className={`min-w-0 rounded-2xl border border-white/10 bg-zinc-900/40 p-3 ${GLASS_CARD}`}>
                    <label className="block text-[10px] font-bold uppercase tracking-wider text-zinc-400">Target shift</label>
                    <select
                      value={solverTargetShift}
                      onChange={(e) => setSolverTargetShift(e.target.value)}
                      className="mt-2 w-full rounded-xl border border-zinc-700 bg-zinc-950 px-2 py-2 text-[13px] sm:text-[14px] text-zinc-100 outline-none [color-scheme:dark]"
                    >
                      {['M', 'A', 'N', 'OFF'].map((shift) => (
                        <option key={shift} value={shift} className="bg-zinc-900 text-zinc-100">
                          {shiftLabel(shift)}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                {solverResult && (
                  <>
                    <div className={`mb-4 rounded-2xl border ${solverResult.possible ? 'border-emerald-500/30 bg-emerald-500/5' : 'border-rose-500/30 bg-rose-500/5'} p-3 min-w-0`}>
                      <div className="font-mono text-[10px] font-bold uppercase tracking-wider text-zinc-300">Result</div>
                      <div className={`mt-1 text-[15px] font-bold ${solverResult.possible ? 'text-emerald-400' : 'text-rose-400'}`}>
                        {solverResult.possible ? 'Possible' : 'Impossible'}
                      </div>
                      <p className="mt-2 text-[12px] leading-relaxed text-zinc-200 break-words">{solverResult.summary}</p>
                      {solverResult.firstAttempt && (
                        <p className="mt-2 text-[11px] text-cyan-300 break-words">First attempted branch: {solverResult.firstAttempt}</p>
                      )}
                    </div>

                    <div className="mb-4 rounded-2xl border border-white/10 bg-zinc-900/40 p-3 min-w-0 overflow-hidden">
                      <div className="mb-2 font-mono text-[10px] font-bold uppercase tracking-wider text-zinc-400">Required swaps</div>
                      {solverResult.swaps.length > 0 ? (
                        <div className="space-y-2 min-w-0">
                          {solverResult.swaps.map((swap, index) => (
                            <div key={`${swap.isoDate}-${swap.otherEmployee}-${index}`} className="rounded-xl border border-white/5 bg-white/5 p-2.5 text-[12px] text-zinc-200 min-w-0 break-words">
                              <div className="font-semibold break-words">{formatRequiredSwap(swap)}</div>
                              <div className="mt-1 text-zinc-300 break-words">{swap.employee} ↔ {swap.otherEmployee}</div>
                              <div className="mt-1 flex flex-wrap items-center gap-2 font-mono text-[10px] uppercase tracking-wider">
                                <span className="text-cyan-300">{swap.from} → {swap.to}</span>
                                {swap.type && <span className={swap.status === 'good' ? 'text-emerald-400' : 'text-amber-300'}>{swap.type === 'WHOLE_BLOCK' ? 'Good' : swap.type === 'PARTIAL_BLOCK' ? 'Under review' : 'Eligible'}</span>}
                              </div>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <div className="text-[12px] text-zinc-300">No swaps were needed for this month preview.</div>
                      )}
                    </div>

                    <div className="rounded-2xl border border-white/10 bg-zinc-900/40 p-2.5 sm:p-3 min-w-0 overflow-hidden">
                      <div className="mb-2 font-mono text-[10px] font-bold uppercase tracking-wider text-zinc-400">Rendered preview</div>
                      <div className="grid grid-cols-7 gap-0.5 sm:gap-1 text-center text-[10px] text-zinc-400 min-w-0 w-full">
                        {['M', 'T', 'W', 'T', 'F', 'S', 'S'].map((day, i) => (
                          <div key={`${day}-${i}`} className="py-1 text-[10px] font-semibold">{day}</div>
                        ))}
                        {monthDays(currentMonth, currentYear).map((day) => {
                          const dayIso = iso(day);
                          const finalShift = shiftKey(solverResult.finalRoster[selectedEmployee]?.[dayIso] ?? 'OFF');
                          const sourceSwap = solverResult.swaps.find((swap) => {
                            const dateParts = swap.isoDate.split(' ↔ ');
                            const requesterDates = dateParts.length === 2 ? [dateParts[0]] : swap.isoDate.split(', ');
                            return requesterDates.includes(dayIso);
                          });
                          const previewTitle = sourceSwap ? `Taken from ${sourceSwap.otherEmployee}` : undefined;
                          return (
                            <div key={dayIso} title={previewTitle} aria-label={previewTitle} className={`min-w-0 rounded-lg border border-white/5 p-0.5 sm:p-1 ${day.getMonth() !== currentMonth ? 'opacity-30' : ''}`}>
                              <div className="text-[9px] sm:text-[10px] text-zinc-500">{day.getDate()}</div>
                              <div className={`mt-0.5 sm:mt-1 truncate rounded-full px-0.5 sm:px-1 py-0.5 text-[8px] sm:text-[9px] font-bold ${colorFor(finalShift).bg} ${colorFor(finalShift).text}`}>
                                {finalShift === 'OFF' ? 'OFF' : finalShift}
                              </div>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  </>
                )}
              </div>
            )}

            {activeTab === 'reports' && (
              <div className="flex min-h-0 flex-1 flex-col overflow-y-auto px-5 pb-24 pt-6 lg:px-0 lg:pb-8">
                <h1 className="text-[34px] font-bold tracking-tight lg:hidden">Reports</h1>
                <p className="mt-1 text-[15px] font-medium text-zinc-400 dark:text-zinc-500 lg:hidden">{selectedEmployee} · {title}</p>
                <h2 className="hidden text-[20px] font-bold tracking-tight lg:block">{title} overview</h2>

                <div className="lg:grid lg:grid-cols-3 lg:items-start lg:gap-4">
                <section className="mt-4 lg:col-span-1 lg:mt-4">
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
                      <p className="text-xs font-medium text-zinc-400 italic">No shared shifts recorded this month.</p>
                    )}
                  </div>
                </section>

                <div className="mt-4 grid grid-cols-2 gap-3 lg:col-span-2 lg:mt-4 lg:grid-cols-3">
                  {statCards.map(({ code, label }) => (
                    <div key={code} className={`p-4 ${GLASS_CARD}`}>
                      <p className="text-[13px] font-bold text-zinc-400">{label}</p>
                      <p className="mt-1 text-[28px] font-bold tracking-tight">{monthlyStats.counts[code] ?? 0}</p>
                    </div>
                  ))}
                  <div className={`p-4 ${GLASS_CARD}`}>
                    <p className="text-[13px] font-bold text-zinc-400">Weekend Shifts</p>
                    <p className="mt-1 text-[28px] font-bold tracking-tight">{monthlyStats.weekendDays}</p>
                  </div>
                </div>

                <div className="mt-3 grid grid-cols-2 gap-3 lg:col-span-3 lg:mt-1">
                  <div className="rounded-[28px] bg-blue-500/10 p-4"><p className="text-[13px] font-semibold text-blue-500">Working Shifts</p><p className="mt-1 text-[28px] font-bold text-blue-500">{monthlyStats.workingShifts}</p></div>
                  <div className="rounded-[28px] bg-green-500/10 p-4"><p className="text-[13px] font-semibold text-green-600 dark:text-green-400">Total Hours</p><p className="mt-1 text-[28px] font-bold text-green-600 dark:text-green-400">{monthlyStats.totalHours}h</p></div>
                </div>
                </div>

                <section className="mt-4">
                  <div className="flex items-center gap-2 px-1">
                    <Wallet className="size-4 text-emerald-500" />
                    <h3 className="text-[13px] font-semibold uppercase tracking-wide text-zinc-400 dark:text-zinc-500">Expected Salary — {title}</h3>
                  </div>
                  <div className={`mt-2 p-4 ${GLASS_CARD}`}>
                    <div className="flex items-center justify-between gap-3">
                      <label htmlFor="base-salary-input" className="shrink-0 text-[13px] font-semibold text-zinc-400">Base salary (RON)</label>
                      <input
                        id="base-salary-input"
                        type="number"
                        inputMode="decimal"
                        min={0}
                        value={salaries[selectedEmployee] ?? ''}
                        onChange={(e) => handleSalaryChange(selectedEmployee, e.target.value)}
                        placeholder="e.g. 4300"
                        className="w-32 rounded-xl border border-zinc-950/10 bg-transparent px-3 py-1.5 text-right text-[15px] font-semibold outline-none focus:border-blue-500 dark:border-white/10"
                      />
                    </div>

                    {salaryBreakdown.baseSalary > 0 ? (
                      <>
                        <div className="mt-4 flex items-center justify-between text-[12px] text-zinc-400">
                          <span>{salaryBreakdown.workingDaysInMonth} working days this month</span>
                          <span className="font-mono">{formatRon(salaryBreakdown.dailyWage)} / day</span>
                        </div>

                        <div className="mt-3 space-y-2 border-t border-zinc-950/10 pt-3 dark:border-white/10">
                          <div className="flex items-center justify-between text-[14px]">
                            <span>Base pay · {salaryBreakdown.workedDays} days worked</span>
                            <span className="font-mono font-semibold">{formatRon(salaryBreakdown.basePay)}</span>
                          </div>
                          {salaryBreakdown.nightCount > 0 && (
                            <div className="flex items-center justify-between text-[14px] text-indigo-400">
                              <span>{salaryBreakdown.nightCount} night shift{salaryBreakdown.nightCount === 1 ? '' : 's'} · +25%</span>
                              <span className="font-mono font-semibold">+{formatRon(salaryBreakdown.nightBonus)}</span>
                            </div>
                          )}
                          {salaryBreakdown.weekendCount > 0 && (
                            <div className="flex items-center justify-between text-[14px] text-cyan-400">
                              <span>{salaryBreakdown.weekendCount} weekend day{salaryBreakdown.weekendCount === 1 ? '' : 's'} · +10%</span>
                              <span className="font-mono font-semibold">+{formatRon(salaryBreakdown.weekendBonus)}</span>
                            </div>
                          )}
                          {salaryBreakdown.holidayCount > 0 && (
                            <div className="flex items-center justify-between text-[14px] text-red-400">
                              <span>{salaryBreakdown.holidayCount} national holiday{salaryBreakdown.holidayCount === 1 ? '' : 's'} · +100%</span>
                              <span className="font-mono font-semibold">+{formatRon(salaryBreakdown.holidayBonus)}</span>
                            </div>
                          )}
                        </div>

                        <div className="mt-3 flex items-center justify-between rounded-2xl bg-emerald-500/10 p-3.5">
                          <span className="text-[15px] font-bold text-emerald-500">Total Expected</span>
                          <span className="text-[22px] font-bold text-emerald-500">{formatRon(salaryBreakdown.total)}</span>
                        </div>
                      </>
                    ) : (
                      <p className="mt-3 text-xs italic text-zinc-400">Enter a base salary to see the earnings breakdown.</p>
                    )}
                  </div>
                </section>
              </div>
            )}

            {activeTab === 'tasks' && (
              <div className="flex min-h-0 flex-1 flex-col overflow-y-auto px-5 pb-24 pt-6 lg:px-0 lg:pb-8">
                <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-3">
                  <div>
                    <h1 className="text-[34px] font-bold tracking-tight">Recurring Tasks</h1>
                    <p className="mt-1 text-[15px] font-medium text-zinc-400 dark:text-zinc-500">
                      Persistent site storage · Synced in real-time across all devices
                    </p>
                  </div>
                  <div className="flex items-center gap-2 self-start sm:self-auto shrink-0">
                    {isSyncingTasks ? (
                      <span className="flex items-center gap-1.5 rounded-full bg-cyan-500/10 px-3 py-1.5 text-[11px] font-semibold text-cyan-500 border border-cyan-500/20">
                        <Loader2 className="size-3.5 animate-spin" />
                        <span>Syncing...</span>
                      </span>
                    ) : (
                      <span className="flex items-center gap-1.5 rounded-full bg-emerald-500/10 px-3 py-1.5 text-[11px] font-semibold text-emerald-500 border border-emerald-500/20">
                        <Cloud className="size-3.5" />
                        <span>Cloud Synced</span>
                      </span>
                    )}
                  </div>
                </div>
                
                <div className="mt-4">
                  <NotificationPromptBanner />
                </div>
                
                <section className="mt-6">
                  <form onSubmit={handleSaveTask} className={`p-4 ${GLASS_CARD} space-y-4`}>
                    <div className="flex items-center justify-between">
                      <h3 className="text-[14px] font-bold text-zinc-400 uppercase tracking-wide">
                        {editingTaskId ? 'Edit Alert Monitor' : 'Create Dynamic Alert'}
                      </h3>
                      {editingTaskId && (
                        <button 
                          type="button" 
                          onClick={resetForm}
                          className="flex items-center gap-1 text-xs text-zinc-400 hover:text-zinc-100"
                        >
                          <X className="size-3.5" /> Cancel Edit
                        </button>
                      )}
                    </div>
                    
                    <div>
                      <label className="block text-[12px] font-semibold text-zinc-400 mb-1">Task Title</label>
                      <input 
                        type="text"
                        value={newTaskTitle}
                        onChange={(e) => setNewTaskTitle(e.target.value)}
                        placeholder="e.g. Verify API endpoint updates"
                        className="w-full rounded-xl border border-zinc-950/10 bg-transparent px-3 py-2 text-[15px] outline-none focus:border-blue-500 dark:border-white/10"
                      />
                    </div>

                    <div>
                      <label className="block text-[12px] font-semibold text-zinc-400 mb-1.5">Schedule Configuration</label>
                      <div className="grid grid-cols-3 gap-1 rounded-xl bg-zinc-950/5 p-1 dark:bg-white/5">
                        {(['once', 'daily', 'weekly'] as ScheduleType[]).map((t) => (
                          <button
                            key={t}
                            type="button"
                            onClick={() => setScheduleType(t)}
                            className={`rounded-lg py-1.5 text-center text-[13px] font-semibold capitalize transition-all ${scheduleType === t ? 'bg-white shadow-sm dark:bg-zinc-800 text-blue-500' : 'text-zinc-400'}`}
                          >
                            {t === 'once' ? 'Run Once' : t}
                          </button>
                        ))}
                      </div>
                    </div>

                    <div>
                      <label className="block text-[12px] font-semibold text-zinc-400 mb-1.5">Target Shift</label>
                      <div className="grid grid-cols-5 gap-1 rounded-xl bg-zinc-950/5 p-1 dark:bg-white/5">
                        {[
                          { key: 'ALL_ACTIVE', label: 'All Shifts' },
                          { key: 'M', label: 'M (06-14)' },
                          { key: 'A', label: 'A (14-22)' },
                          { key: 'N', label: 'N (22-06)' },
                          { key: 'MID', label: 'MID (09-17)' },
                        ].map((item) => (
                          <button
                            key={item.key}
                            type="button"
                            onClick={() => setTargetShift(item.key as TargetShiftFilter)}
                            className={`rounded-lg py-1.5 text-center text-[11px] font-semibold transition-all ${
                              targetShift === item.key
                                ? 'bg-white shadow-sm dark:bg-zinc-800 text-blue-500'
                                : 'text-zinc-400'
                            }`}
                          >
                            {item.label}
                          </button>
                        ))}
                      </div>
                      <p className="mt-1 px-0.5 text-[11px] text-zinc-400">
                        {targetShift === 'ALL_ACTIVE'
                          ? 'Alerts whoever is actively on shift at the time (off-shift colleagues are silenced)'
                          : `Alerts only colleagues working Shift ${targetShift} (all others are silenced)`}
                      </p>
                    </div>

                    {scheduleType === 'weekly' && (
                      <div className="animate-fadeIn">
                        <label className="block text-[12px] font-semibold text-zinc-400 mb-1.5">Active Days</label>
                        <div className="flex items-center justify-between gap-1">
                          {WEEKDAY_LABELS.map((label, idx) => {
                            const systemDayVal = WEEKDAY_INDEXES[idx];
                            const isChosen = selectedDays.includes(systemDayVal);
                            return (
                              <button
                                key={label}
                                type="button"
                                onClick={() => toggleDaySelection(systemDayVal)}
                                className={`flex-1 rounded-xl py-2 text-center text-[12px] font-bold transition-all ${isChosen ? 'bg-blue-500 text-white shadow-md shadow-blue-500/10' : 'bg-zinc-950/5 text-zinc-400 dark:bg-white/5'}`}
                              >
                                {label}
                              </button>
                            );
                          })}
                        </div>
                      </div>
                    )}

                    <div className="space-y-2">
                      <label className="block text-[12px] font-semibold text-zinc-400">Alarm Triggers (24-Hour Format)</label>
                      {newTaskTimes.map((timeObj, index) => (
                        <div key={index} className="flex items-center gap-2">
                          <div className="flex flex-1 items-center gap-1 rounded-xl border border-zinc-950/10 bg-transparent px-2 py-1 dark:border-white/10">
                            <select 
                              value={timeObj.hour}
                              onChange={(e) => handleTimeDropdownChange(index, 'hour', e.target.value)}
                              className="w-full bg-transparent py-1 text-center text-[16px] font-semibold outline-none dark:text-white [color-scheme:dark]"
                            >
                              {hoursArray.map(h => <option key={h} value={h} className="bg-zinc-100 dark:bg-zinc-900">{h}</option>)}
                            </select>
                            <span className="text-zinc-400 font-bold">:</span>
                            <select 
                              value={timeObj.minute}
                              onChange={(e) => handleTimeDropdownChange(index, 'minute', e.target.value)}
                              className="w-full bg-transparent py-1 text-center text-[16px] font-semibold outline-none dark:text-white [color-scheme:dark]"
                            >
                              {minutesArray.map(m => <option key={m} value={m} className="bg-zinc-100 dark:bg-zinc-900">{m}</option>)}
                            </select>
                          </div>
                          {newTaskTimes.length > 1 && (
                            <button 
                              type="button" 
                              onClick={() => removeTimeInputField(index)}
                              className="p-2 text-red-500 hover:bg-red-500/10 rounded-xl shrink-0"
                            >
                              <Trash2 className="size-4" />
                            </button>
                          )}
                        </div>
                      ))}
                      <button 
                        type="button" 
                        onClick={addTimeInputField}
                        className="flex items-center gap-1.5 text-[13px] font-bold text-blue-500 mt-1 hover:underline"
                      >
                        <Plus className="size-4" /> Add another time
                      </button>
                    </div>

                    <button 
                      type="submit"
                      disabled={scheduleType === 'weekly' && selectedDays.length === 0}
                      className="w-full flex items-center justify-center gap-2 bg-blue-500 hover:bg-blue-600 disabled:opacity-40 disabled:hover:bg-blue-500 text-white font-semibold py-2.5 px-4 rounded-xl transition-colors cursor-pointer"
                    >
                      <Save className="size-4" /> {editingTaskId ? 'Update Task' : 'Save Task'}
                    </button>
                  </form>
                </section>

                <section className="mt-6">
                  <div className="flex items-center justify-between px-1 mb-2">
                    <h3 className="text-[13px] font-semibold uppercase tracking-wide text-zinc-400 dark:text-zinc-500">
                      Active Monitor Configuration ({tasks.length})
                    </h3>
                    {tasksLastUpdated && (
                      <span className="text-[10px] font-mono text-zinc-400 dark:text-zinc-500">
                        Synced {new Date(tasksLastUpdated).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}
                      </span>
                    )}
                  </div>
                  <div className="space-y-2">
                    {tasks.length === 0 ? (
                      <p className="text-xs italic text-zinc-400 px-1">No custom metrics monitored. Create a task above.</p>
                    ) : (
                      tasks.map((task) => (
                        <div key={task.id} className={`flex items-center justify-between p-4 ${GLASS_CARD} ${editingTaskId === task.id ? 'ring-2 ring-blue-500/50 bg-blue-500/[0.02]' : ''}`}>
                          <div className="min-w-0 flex-1 pr-4">
                            <div className="flex items-center gap-2 flex-wrap">
                              <p className="font-semibold text-[15px] truncate">{task.title}</p>
                              <span className={`rounded-full px-2 py-0.5 text-[9px] font-extrabold tracking-wide uppercase ${task.scheduleType === 'once' ? 'bg-amber-500/10 text-amber-500 border border-amber-500/20' : task.scheduleType === 'weekly' ? 'bg-indigo-500/10 text-indigo-400 border border-indigo-500/20' : 'bg-green-500/10 text-green-400 border border-green-500/20'}`}>
                                {task.scheduleType === 'once' ? 'One Time' : task.scheduleType}
                              </span>
                              {task.targetShift && task.targetShift !== 'ALL_ACTIVE' ? (
                                <span className="rounded-full px-2 py-0.5 text-[9px] font-extrabold tracking-wide uppercase bg-blue-500/10 text-blue-500 dark:text-blue-400 border border-blue-500/20">
                                  Shift {task.targetShift}
                                </span>
                              ) : (
                                <span className="rounded-full px-2 py-0.5 text-[9px] font-medium text-zinc-400 border border-zinc-500/20">
                                  On-Shift Only
                                </span>
                              )}
                            </div>
                            
                            {task.scheduleType === 'weekly' && task.daysOfWeek && (
                              <p className="text-[11px] text-zinc-400 dark:text-zinc-500 mt-1 font-medium">
                                Active on: {task.daysOfWeek.map(d => WEEKDAYS_MAP[d].substring(0,3)).join(', ')}
                              </p>
                            )}

                            <div className="flex flex-wrap gap-1 mt-1.5">
                              {task.times.map((t) => (
                                <span key={t} className="text-[10px] font-mono font-bold bg-zinc-950/5 dark:bg-white/10 text-zinc-500 dark:text-zinc-400 px-2 py-0.5 rounded-md">
                                  {t}
                                </span>
                              ))}
                            </div>

                            {/* Personal shift status badge for current user */}
                            {(() => {
                              const emp = authenticatedEmployee || selectedEmployee;
                              if (!roster || !emp) return null;
                              const status = isEmployeeOnShift(roster, emp, new Date());
                              if (!status.onShift) {
                                return (
                                  <p className="mt-1 text-[11px] text-zinc-400 dark:text-zinc-500 flex items-center gap-1">
                                    <span>💤</span> Silent for you (You are off shift)
                                  </p>
                                );
                              }
                              const matchesShift = !task.targetShift || task.targetShift === 'ALL_ACTIVE' || task.targetShift === status.currentShift;
                              if (matchesShift) {
                                return (
                                  <p className="mt-1 text-[11px] text-emerald-500 font-medium flex items-center gap-1.5">
                                    <span className="size-1.5 rounded-full bg-emerald-500 animate-pulse" />
                                    Active for you (Shift {status.currentShift})
                                  </p>
                                );
                              }
                              return (
                                <p className="mt-1 text-[11px] text-zinc-400 dark:text-zinc-500 flex items-center gap-1">
                                  <span>💤</span> Silent for you (You are on shift {status.currentShift})
                                </p>
                              );
                            })()}
                          </div>
                          
                          <div className="flex items-center gap-1 shrink-0">
                            <button 
                              onClick={() => startEditingTask(task)}
                              className="p-2.5 text-zinc-400 hover:text-blue-500 hover:bg-blue-500/10 rounded-xl transition-colors"
                              title="Edit Task parameters"
                            >
                              <Edit2 className="size-4" />
                            </button>
                            <button 
                              onClick={() => handleRemoveTask(task.id)}
                              className="p-2.5 text-zinc-400 hover:text-red-500 hover:bg-red-500/10 rounded-xl transition-colors"
                              title="Delete task configuration"
                            >
                              <Trash2 className="size-4" />
                            </button>
                          </div>
                        </div>
                      ))
                    )}
                  </div>
                </section>
              </div>
            )}

            {activeTab === 'settings' && (
              <div className="flex min-h-0 flex-1 flex-col overflow-y-auto px-5 pb-24 pt-6 lg:px-0 lg:pb-8">
                <h1 className="text-[34px] font-bold tracking-tight lg:hidden">Settings</h1>

                <div className="lg:grid lg:grid-cols-2 lg:items-start lg:gap-4">
                <section className="mt-6 lg:mt-0 space-y-4">
                  <div>
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
                  </div>

                  <div>
                    <h3 className="px-1 text-[13px] font-semibold uppercase tracking-wide text-zinc-400 dark:text-zinc-500">Notifications</h3>
                    <div className="mt-2">
                      <NotificationSettingsCard authenticatedEmployee={authenticatedEmployee} />
                    </div>
                  </div>
                </section>

                <section className="mt-6 lg:mt-0 space-y-4">
                  {isAlexandruStoian && (
                    <div className="hidden lg:block">
                      <div className="flex items-center justify-between px-1">
                        <h3 className="text-[13px] font-semibold uppercase tracking-wide text-zinc-400 dark:text-zinc-500">Cloudflare KV Schedule Sync</h3>
                        <span className={`inline-flex items-center gap-1 rounded-full px-2 py-0.5 font-mono text-[10px] font-semibold ${scheduleSource === 'kv' ? 'bg-lime-500/10 text-lime-500' : 'bg-zinc-500/10 text-zinc-400'}`}>
                          <Cloud className="size-3" />
                          {scheduleSource === 'kv' ? 'Live KV' : 'Local Bundle'}
                        </span>
                      </div>

                      <div className={`mt-2 p-4 space-y-3 ${GLASS_CARD}`}>
                        <input
                          ref={fileInputRef}
                          type="file"
                          accept=".xlsx,.xls"
                          onChange={handleFileSelect}
                          className="hidden"
                        />

                        <button
                          type="button"
                          onClick={() => fileInputRef.current?.click()}
                          className="flex w-full flex-col items-center justify-center gap-2 rounded-2xl border-2 border-dashed border-zinc-950/10 p-4 transition-colors hover:border-blue-500/40 hover:bg-blue-500/5 dark:border-white/10 dark:hover:border-blue-400/40 text-center cursor-pointer"
                        >
                          <UploadCloud className="size-6 text-blue-500" />
                          <div>
                            <p className="text-[14px] font-semibold">{uploadFile ? uploadFile.name : "Upload Boss's Excel (.xlsx)"}</p>
                            <p className="text-[11px] text-zinc-400 dark:text-zinc-500">Tap to select or drop updated schedule file</p>
                          </div>
                        </button>

                        {parsedPreview && (
                          <div className="rounded-xl bg-zinc-950/5 p-3.5 dark:bg-white/5 space-y-2 border border-zinc-950/5 dark:border-white/5 text-[13px]">
                            <div className="flex items-center justify-between">
                              <span className="text-zinc-400">Detected Period:</span>
                              <span className="font-semibold text-blue-500">
                                {new Intl.DateTimeFormat('en', { month: 'long', year: 'numeric' }).format(new Date(parsedPreview.year, parsedPreview.month))}
                              </span>
                            </div>
                            <div className="flex items-center justify-between">
                              <span className="text-zinc-400">Employees Found:</span>
                              <span className="font-mono font-semibold">{parsedPreview.employees.length}</span>
                            </div>
                            <div className="flex items-center justify-between">
                              <span className="text-zinc-400">Calendar Days:</span>
                              <span className="font-mono font-semibold">{parsedPreview.dateColumns.length} days</span>
                            </div>

                            <div className="pt-2 border-t border-zinc-950/5 dark:border-white/5 flex items-center justify-between">
                              <label className="text-[12px] flex items-center gap-2 cursor-pointer select-none">
                                <input
                                  type="checkbox"
                                  checked={mergeWithExisting}
                                  onChange={(e) => setMergeWithExisting(e.target.checked)}
                                  className="rounded text-blue-500"
                                />
                                Merge with existing months
                              </label>
                            </div>

                            <div className="pt-2 border-t border-zinc-950/5 dark:border-white/5">
                              <input
                                type="password"
                                value={adminPin}
                                onChange={(e) => setAdminPin(e.target.value)}
                                placeholder="Admin PIN (if configured)"
                                className="w-full rounded-lg bg-zinc-950/5 dark:bg-white/5 px-3 py-1.5 text-xs outline-none border border-zinc-950/10 dark:border-white/10"
                              />
                            </div>

                            <button
                              type="button"
                              onClick={handlePublishRoster}
                              disabled={isUploading}
                              className="w-full mt-2 flex items-center justify-center gap-2 rounded-xl bg-blue-500 hover:bg-blue-600 text-white font-bold py-2.5 px-4 text-[13px] shadow transition-all active:scale-98 disabled:opacity-50 cursor-pointer"
                            >
                              {isUploading ? <Loader2 className="size-4 animate-spin" /> : <Cloud className="size-4" />}
                              {isUploading ? 'Publishing to KV...' : 'Publish to Cloudflare KV'}
                            </button>
                          </div>
                        )}

                        {uploadStatus && (
                          <div className={`p-3 rounded-xl text-[13px] flex items-center gap-2 ${uploadStatus.type === 'success' ? 'bg-lime-500/10 text-lime-600 dark:text-lime-400 border border-lime-500/20' : 'bg-red-500/10 text-red-600 dark:text-red-400 border border-red-500/20'}`}>
                            {uploadStatus.type === 'success' ? <Check className="size-4 shrink-0" /> : <AlertTriangle className="size-4 shrink-0" />}
                            <p className="leading-snug">{uploadStatus.message}</p>
                          </div>
                        )}
                      </div>
                    </div>
                  )}

                  <div>
                    <div className="flex items-center justify-between px-1">
                      <h3 className="text-[13px] font-semibold uppercase tracking-wide text-zinc-400 dark:text-zinc-500">Employee Identity & Authentication</h3>
                      <span className="text-[11px] font-mono text-zinc-400 dark:text-zinc-500 truncate max-w-[180px]">{displayFileName(roster?.fileName)}</span>
                    </div>

                    {/* Authenticated Identity Status Banner */}
                    {authenticatedEmployee ? (
                      <div className="mt-2 mb-3 flex items-center justify-between rounded-2xl bg-emerald-500/10 p-3.5 border border-emerald-500/20">
                        <div className="flex items-center gap-2.5">
                          <div className="flex size-8 items-center justify-center rounded-xl bg-emerald-500/20 text-emerald-600 dark:text-emerald-400">
                            <ShieldCheck className="size-5" />
                          </div>
                          <div>
                            <p className="text-[14px] font-bold text-emerald-600 dark:text-emerald-400">
                              {authenticatedEmployee}
                            </p>
                            <p className="text-[11px] text-zinc-500 dark:text-zinc-400">
                              Authenticated Identity · Notifications Linked
                            </p>
                          </div>
                        </div>
                        <button
                          type="button"
                          onClick={handleLogOut}
                          className="flex items-center gap-1 rounded-xl bg-zinc-950/5 px-2.5 py-1.5 text-[12px] font-medium text-zinc-500 hover:bg-zinc-950/10 dark:bg-white/5 dark:text-zinc-400 dark:hover:bg-white/10 transition-colors"
                          title="Switch user or log out"
                        >
                          <LogOut className="size-3.5" />
                          <span>Switch</span>
                        </button>
                      </div>
                    ) : (
                      <div className="mt-2 mb-3 flex items-center gap-2.5 rounded-2xl bg-amber-500/10 p-3.5 border border-amber-500/20 text-amber-600 dark:text-amber-400">
                        <Lock className="size-5 shrink-0" />
                        <div>
                          <p className="text-[13px] font-semibold">Select your profile to authenticate</p>
                          <p className="text-[11px] opacity-85">Tap your name below and enter your secret code to log in.</p>
                        </div>
                      </div>
                    )}

                    <div className="mt-2">
                      <button
                        type="button"
                        onClick={() => {
                          setQuery('');
                          setIsEmployeeDialogOpen(true);
                        }}
                        className={`flex w-full items-center justify-between p-4 ${GLASS_CARD} text-left transition-all hover:bg-zinc-950/[0.03] dark:hover:bg-white/[0.03] active:scale-[0.99] cursor-pointer`}
                      >
                        <div className="flex items-center gap-3 min-w-0">
                          <div className="flex size-10 shrink-0 items-center justify-center rounded-2xl bg-blue-500/10 text-blue-500">
                            <Users className="size-5" />
                          </div>
                          <div className="min-w-0">
                            <p className="text-[11px] font-semibold uppercase tracking-wider text-zinc-400 dark:text-zinc-500">
                              Active Profile
                            </p>
                            <p className="text-[15px] font-bold truncate text-zinc-900 dark:text-zinc-100">
                              {selectedEmployee || 'Select Employee'}
                            </p>
                          </div>
                        </div>
                        <div className="flex items-center gap-2 shrink-0">
                          <span className="rounded-xl bg-blue-500 px-3 py-1.5 text-[12px] font-bold text-white shadow-sm hover:bg-blue-600 transition-colors">
                            Select Employee
                          </span>
                          <ChevronDown className="size-4 text-zinc-400" />
                        </div>
                      </button>
                    </div>
                  </div>
                </section>
                </div>
              </div>
            )}
            </div>
          </div>

          <DayDetailsModal
            event={selectedEvent}
            dailyRoster={selectedEvent ? rosterIndex[selectedEvent.isoDate] : undefined}
            roster={roster}
            selectedEmployee={selectedEmployee}
            canGoPrev={canGoPrevDay} canGoNext={canGoNextDay}
            onPrev={goToPrevDay} onNext={goToNextDay}
            onClose={() => setSelectedEvent(null)}
          />

          <AuthPinModal
            isOpen={Boolean(authModalTarget)}
            targetEmployee={authModalTarget}
            onClose={() => setAuthModalTarget(null)}
            onAuthenticated={handleAuthSuccess}
          />

          {/* Employee Selection Dialog Modal */}
          {isEmployeeDialogOpen && (
            <div
              className="fixed inset-0 z-50 flex items-end sm:items-center justify-center bg-black/60 p-0 sm:p-4 backdrop-blur-sm animate-fadeIn"
              onClick={() => setIsEmployeeDialogOpen(false)}
            >
              <div
                className="relative flex flex-col w-full max-w-lg rounded-t-3xl sm:rounded-3xl border border-zinc-950/10 bg-white/95 dark:border-white/10 dark:bg-zinc-900/95 shadow-2xl backdrop-blur-xl text-zinc-950 dark:text-white max-h-[85vh] sm:max-h-[75vh] overflow-hidden"
                onClick={(e) => e.stopPropagation()}
              >
                {/* Mobile drag handle */}
                <div className="flex sm:hidden justify-center pt-3 pb-1">
                  <div className="h-1 w-10 rounded-full bg-zinc-300 dark:bg-zinc-700" />
                </div>

                {/* Dialog Header */}
                <div className="flex items-center justify-between px-5 pt-3 pb-3 border-b border-zinc-950/5 dark:border-white/5">
                  <div>
                    <h3 className="text-[18px] font-bold tracking-tight">Select Employee</h3>
                    <p className="text-[12px] text-zinc-500 dark:text-zinc-400">
                      Choose an employee to view their schedule and shifts
                    </p>
                  </div>
                  <button
                    type="button"
                    onClick={() => setIsEmployeeDialogOpen(false)}
                    className="flex size-8 items-center justify-center rounded-full text-zinc-400 hover:bg-zinc-950/5 dark:hover:bg-white/10 transition-colors cursor-pointer"
                  >
                    <X className="size-4" />
                  </button>
                </div>

                {/* Search Input */}
                <div className="px-5 py-3 border-b border-zinc-950/5 dark:border-white/5">
                  <div className="flex items-center gap-2.5 rounded-2xl bg-zinc-950/5 dark:bg-white/5 px-3.5 py-2">
                    <Search className="size-4 shrink-0 text-zinc-400" />
                    <input
                      type="text"
                      value={query}
                      onChange={(e) => setQuery(e.target.value)}
                      placeholder="Search employee name..."
                      autoFocus
                      className="w-full bg-transparent text-[14px] outline-none placeholder:text-zinc-400"
                    />
                    {query && (
                      <button
                        type="button"
                        onClick={() => setQuery('')}
                        className="text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200 p-0.5 cursor-pointer"
                      >
                        <X className="size-3.5" />
                      </button>
                    )}
                  </div>
                </div>

                {/* Employee List with isolated, smooth scrolling */}
                <div className="flex-1 overflow-y-auto overscroll-contain divide-y divide-zinc-950/5 dark:divide-white/5 px-2 py-1">
                  {filteredEmployees.length === 0 ? (
                    <div className="py-12 text-center">
                      <p className="text-[14px] font-medium text-zinc-400">No employees found</p>
                      <p className="text-[12px] text-zinc-500 mt-1">Try a different search term</p>
                    </div>
                  ) : (
                    filteredEmployees.map((name) => {
                      const isAuth = name === authenticatedEmployee;
                      const isSelected = name === selectedEmployee;
                      return (
                        <button
                          key={name}
                          type="button"
                          onClick={() => {
                            handleSelectEmployee(name);
                            setIsEmployeeDialogOpen(false);
                          }}
                          className={`flex w-full items-center justify-between rounded-xl px-4 py-3 text-left transition-colors cursor-pointer ${
                            isSelected
                              ? 'bg-blue-500/10 text-blue-600 dark:text-blue-400 font-semibold'
                              : 'hover:bg-zinc-950/5 dark:hover:bg-white/5'
                          }`}
                        >
                          <div className="flex items-center gap-3 min-w-0">
                            <div className={`flex size-8 shrink-0 items-center justify-center rounded-xl font-mono text-[11px] font-bold ${
                              isSelected
                                ? 'bg-blue-500 text-white'
                                : isAuth
                                ? 'bg-emerald-500/20 text-emerald-500'
                                : 'bg-zinc-950/5 dark:bg-white/10 text-zinc-500 dark:text-zinc-400'
                            }`}>
                              {initials(name)}
                            </div>
                            <div className="min-w-0">
                              <span className={`block truncate text-[14px] ${isSelected || isAuth ? 'font-bold' : 'font-medium'}`}>
                                {name}
                              </span>
                              {isAuth && (
                                <span className="text-[11px] text-emerald-500 font-semibold">
                                  Authenticated Profile
                                </span>
                              )}
                            </div>
                          </div>
                          <div className="flex items-center gap-2 shrink-0">
                            {isAuth ? (
                              <span className="flex items-center gap-1 rounded-full bg-emerald-500/10 px-2 py-0.5 text-[10px] font-bold text-emerald-500 border border-emerald-500/20">
                                <ShieldCheck className="size-3" /> Verified
                              </span>
                            ) : (
                              <Lock className="size-3.5 text-zinc-400 opacity-50" />
                            )}
                            {isSelected && <Check className="size-4 text-blue-500" />}
                          </div>
                        </button>
                      );
                    })
                  )}
                </div>
              </div>
            </div>
          )}

          {/* Alarm Loop Intercept Modal */}
          {activeAlarmTask && (
            <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 p-6 backdrop-blur-md">
              <div className="w-full max-w-[360px] animate-bounce rounded-3xl bg-zinc-900 border border-red-500/30 p-6 text-center text-white shadow-2xl">
                <div className="mx-auto flex size-14 items-center justify-center rounded-full bg-red-500/20 text-red-500 mb-4 animate-pulse">
                  <AlertTriangle className="size-7" />
                </div>
                <h2 className="text-[20px] font-black tracking-tight">System Monitor Critical</h2>
                <p className="mt-1 font-mono text-zinc-400 text-xs">Triggered at {activeAlarmTask.time}</p>
                <div className="my-4 rounded-2xl bg-white/5 p-4 border border-white/5">
                  <p className="text-[15px] font-semibold break-words">{activeAlarmTask.taskTitle}</p>
                </div>
                <button 
                  onClick={dismissAlarm}
                  className="w-full bg-red-500 hover:bg-red-600 text-white font-bold py-3 px-4 rounded-xl shadow-lg shadow-red-500/20 active:scale-[0.98] transition-all text-[15px]"
                >
                  Dismiss Loop Alarm
                </button>
              </div>
            </div>
          )}

          <nav
            aria-hidden={sheetOpen}
            style={{ paddingBottom: 'max(0.75rem, env(safe-area-inset-bottom))' }}
            className={`fixed inset-x-0 bottom-0 z-10 flex justify-center px-3 sm:px-5 transition-[filter] duration-200 lg:hidden ${
              sheetOpen ? 'blur-[1px] pointer-events-none' : ''
            }`}
          >
            <div className={`relative flex w-full max-w-[400px] p-1.5 ${GLASS_NAV}`}>
              {/* Dynamic Liquid Glass Active Indicator */}
              <div
                style={{
                  width: `calc((100% - 0.75rem) / ${tabs.length})`,
                  transform: `translateX(${activeTabIndex * 100}%)`,
                }}
                className="absolute inset-y-1.5 left-1.5 rounded-[22px] bg-white/90 shadow-sm transition-transform duration-[220ms] ease-out dark:bg-white/[0.16] pointer-events-none"
              />
              {tabs.map(({ id, label, Icon }) => {
                const isActive = activeTab === id;
                return (
                  <button
                    key={id}
                    onClick={() => handleTabChange(id)}
                    className={`relative z-10 flex flex-1 flex-col items-center justify-center py-1.5 gap-0.5 transition-colors duration-150 select-none ${
                      isActive
                        ? 'text-blue-600 dark:text-blue-400 font-semibold'
                        : 'text-zinc-400 dark:text-zinc-500 hover:text-zinc-600 dark:hover:text-zinc-300'
                    }`}
                  >
                    <Icon className={`size-[18px] transition-transform duration-150 ${isActive ? 'scale-105' : ''}`} />
                    <span className="text-[10px] tracking-tight truncate max-w-full leading-tight">
                      {label}
                    </span>
                  </button>
                );
              })}
            </div>
          </nav>
          </div>
        </div>
      </div>
    </main>
  );
}