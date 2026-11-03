// src/ScheduleBuilderTab.tsx
import { useState, useMemo, useEffect } from 'react';
import {
  RotateCcw,
  Save,
  FileSpreadsheet,
  Download,
  ArrowLeftRight,
  ChevronLeft,
  ChevronRight,
  AlertTriangle,
  CheckCircle2,
  Shield,
  Lock,
  Clock,
  X
} from 'lucide-react';
import type { RosterData } from './types';
import {
  generateBalancedSchedule,
  validateScheduleCompliance,
  swapEmployeeShifts,
  exportRosterToExcelFile,
  countLegalWorkingDaysInMonth,
  isHistoricalMonthInRoster,
  loadHistoricalSchedule,
  FIXED_MID_EMPLOYEES,
  type GeneratedScheduleResult,
  type ScheduleValidationReport
} from './scheduleBuilder';
import { shiftKey, shiftColors, initials, GLASS_CARD } from './scheduleUtils';
import { mergeRosters } from './parser';

interface ScheduleBuilderTabProps {
  roster: RosterData | null;
  onApplySchedule: (updatedRoster: RosterData, targetMonth: number, targetYear: number) => void;
}

export default function ScheduleBuilderTab({ roster, onApplySchedule }: ScheduleBuilderTabProps) {
  // Target month & year state (defaults to current month or next month)
  const [targetYear, setTargetYear] = useState<number>(() => {
    if (roster && roster.dateColumns.length > 0) {
      const last = roster.dateColumns[roster.dateColumns.length - 1];
      const d = new Date(`${last.isoDate}T00:00:00`);
      return d.getFullYear();
    }
    return new Date().getFullYear();
  });

  const [targetMonth, setTargetMonth] = useState<number>(() => {
    if (roster && roster.dateColumns.length > 0) {
      const last = roster.dateColumns[roster.dateColumns.length - 1];
      const d = new Date(`${last.isoDate}T00:00:00`);
      return d.getMonth();
    }
    return new Date().getMonth();
  });

  // Target legal working days norm for the month (e.g. 21 in Dec 2026, 22 in Oct 2026) - set by law, not modifiable
  const targetWorkingDays = useMemo(
    () => countLegalWorkingDaysInMonth(targetMonth, targetYear),
    [targetMonth, targetYear]
  );

  // Whether the selected month is an archived/historical record in the roster (set in stone)
  const isHistorical = useMemo(() => {
    return isHistoricalMonthInRoster(roster, targetYear, targetMonth);
  }, [roster, targetYear, targetMonth]);

  // Current working generated schedule
  const [scheduleData, setScheduleData] = useState<GeneratedScheduleResult | null>(null);
  const [currentSeed, setCurrentSeed] = useState<number>(() => Math.floor(Math.random() * 999999));
  const [appliedNotification, setAppliedNotification] = useState<string | null>(null);

  // Multi-day selection state for swaps
  const [selectedDates, setSelectedDates] = useState<Set<string>>(new Set());
  const [swapEmployeeA, setSwapEmployeeA] = useState<string>('');
  const [swapEmployeeB, setSwapEmployeeB] = useState<string>('');
  const [swapFeedback, setSwapFeedback] = useState<{ type: 'success' | 'warning'; message: string } | null>(null);

  // Single cell editing popover
  const [editingCell, setEditingCell] = useState<{ employee: string; isoDate: string } | null>(null);
  const [hasUserNavigated, setHasUserNavigated] = useState(false);

  // Sync initial month and year once roster data loads if user hasn't explicitly navigated
  useEffect(() => {
    if (!hasUserNavigated && roster && roster.dateColumns.length > 0) {
      const last = roster.dateColumns[roster.dateColumns.length - 1];
      const d = new Date(`${last.isoDate}T00:00:00`);
      setTargetYear(d.getFullYear());
      setTargetMonth(d.getMonth());
    }
  }, [roster, hasUserNavigated]);

  // Generate initial schedule or load historical schedule if already set in stone
  const runGeneration = (seedValue: number = currentSeed) => {
    if (isHistoricalMonthInRoster(roster, targetYear, targetMonth)) {
      setScheduleData(loadHistoricalSchedule(roster!, targetYear, targetMonth));
    } else {
      const result = generateBalancedSchedule(roster, targetYear, targetMonth, {
        seed: seedValue,
        targetWorkingDays
      });
      setScheduleData(result);
    }
    setSelectedDates(new Set());
    setSwapFeedback(null);
  };

  useEffect(() => {
    runGeneration(currentSeed);
  }, [targetYear, targetMonth, targetWorkingDays, roster]);

  // Live validation whenever schedule rows change
  const validation: ScheduleValidationReport = useMemo(() => {
    if (!scheduleData) {
      return {
        isFullyCompliant: true,
        totalViolations: 0,
        understaffedDaysCount: 0,
        streakViolationCount: 0,
        nightRestViolationCount: 0,
        excessiveRestViolationCount: 0,
        mixedBlockViolationCount: 0,
        workDaysMismatchCount: 0,
        dailyCoverage: {},
        employeeStats: {},
        warnings: []
      };
    }
    return validateScheduleCompliance(
      scheduleData.rows,
      scheduleData.dateColumns,
      roster,
      targetYear,
      targetMonth,
      targetWorkingDays
    );
  }, [scheduleData, roster, targetYear, targetMonth, targetWorkingDays]);

  // Employee list
  const employees = useMemo(() => {
    if (scheduleData) return Object.keys(scheduleData.rows);
    return roster?.employees ?? [];
  }, [scheduleData, roster]);

  const rotatingEmployees = useMemo(() => {
    return employees.filter((emp) => !FIXED_MID_EMPLOYEES.includes(emp));
  }, [employees]);

  // Handlers for month navigation
  const handlePrevMonth = () => {
    setHasUserNavigated(true);
    setTargetMonth((prev) => {
      if (prev === 0) {
        setTargetYear((y) => y - 1);
        return 11;
      }
      return prev - 1;
    });
  };

  const handleNextMonth = () => {
    setHasUserNavigated(true);
    setTargetMonth((prev) => {
      if (prev === 11) {
        setTargetYear((y) => y + 1);
        return 0;
      }
      return prev + 1;
    });
  };

  const handleRegenerateVariant = () => {
    if (isHistorical) return;
    const nextSeed = Math.floor(Math.random() * 999999);
    setCurrentSeed(nextSeed);
    runGeneration(nextSeed);
  };

  // Date multi-selection toggle
  const toggleDateSelection = (isoDate: string) => {
    if (isHistorical) return;
    setSelectedDates((prev) => {
      const next = new Set(prev);
      if (next.has(isoDate)) next.delete(isoDate);
      else next.add(isoDate);
      return next;
    });
  };

  const selectAllWeekdays = () => {
    if (!scheduleData) return;
    const weekdays = scheduleData.dateColumns
      .filter((d) => d.date.getDay() !== 0 && d.date.getDay() !== 6)
      .map((d) => d.isoDate);
    setSelectedDates(new Set(weekdays));
  };

  const selectAllWeekends = () => {
    if (!scheduleData) return;
    const weekends = scheduleData.dateColumns
      .filter((d) => d.date.getDay() === 0 || d.date.getDay() === 6)
      .map((d) => d.isoDate);
    setSelectedDates(new Set(weekends));
  };

  const clearSelection = () => {
    setSelectedDates(new Set());
    setSwapFeedback(null);
  };

  // Apply swap between two employees on selected dates
  const handleExecuteSwap = () => {
    if (isHistorical || !scheduleData || !swapEmployeeA || !swapEmployeeB || selectedDates.size === 0) return;
    if (swapEmployeeA === swapEmployeeB) {
      setSwapFeedback({ type: 'warning', message: 'Select two different employees to swap shifts.' });
      return;
    }

    const datesArray = Array.from(selectedDates);
    const updatedRows = swapEmployeeShifts(
      scheduleData.rows,
      swapEmployeeA,
      swapEmployeeB,
      datesArray
    );

    // Validate updated
    const testReport = validateScheduleCompliance(
      updatedRows,
      scheduleData.dateColumns,
      roster,
      targetYear,
      targetMonth
    );

    setScheduleData({
      ...scheduleData,
      rows: updatedRows,
      validation: testReport
    });

    if (testReport.isFullyCompliant) {
      setSwapFeedback({
        type: 'success',
        message: `Swapped ${datesArray.length} day(s) between ${swapEmployeeA.split(' ')[0]} and ${swapEmployeeB.split(' ')[0]}. Schedule remains 100% compliant!`
      });
    } else {
      setSwapFeedback({
        type: 'warning',
        message: `Swapped shifts with ${testReport.totalViolations} legal/coverage alert(s). Review summary below.`
      });
    }
  };

  // Manual shift override
  const handleSetCellShift = (shiftCode: string) => {
    if (isHistorical || !editingCell || !scheduleData) return;
    const { employee, isoDate } = editingCell;

    const updatedRows = {
      ...scheduleData.rows,
      [employee]: {
        ...scheduleData.rows[employee],
        [isoDate]: shiftCode
      }
    };

    const testReport = validateScheduleCompliance(
      updatedRows,
      scheduleData.dateColumns,
      roster,
      targetYear,
      targetMonth
    );

    setScheduleData({
      ...scheduleData,
      rows: updatedRows,
      validation: testReport
    });

    setEditingCell(null);
  };

  // Apply schedule to app active roster
  const handleApplyToCalendar = () => {
    if (!scheduleData) return;

    const newRosterPiece: RosterData = {
      fileName: `Generated_${targetYear}_${String(targetMonth + 1).padStart(2, '0')}.xlsx`,
      month: targetMonth,
      year: targetYear,
      employees: Object.keys(scheduleData.rows),
      dateColumns: scheduleData.dateColumns,
      rows: scheduleData.rows
    };

    const merged = roster ? mergeRosters(roster, newRosterPiece) : newRosterPiece;
    onApplySchedule(merged, targetMonth, targetYear);
    setAppliedNotification(`Schedule for ${monthDisplay} applied to Calendar & Reports!`);
    setTimeout(() => setAppliedNotification(null), 4000);
  };

  // Export to Excel
  const handleExportExcel = async () => {
    if (!scheduleData) return;
    await exportRosterToExcelFile({
      fileName: `Schedule_${targetYear}_${String(targetMonth + 1).padStart(2, '0')}.xlsx`,
      month: targetMonth,
      year: targetYear,
      employees: Object.keys(scheduleData.rows),
      dateColumns: scheduleData.dateColumns,
      rows: scheduleData.rows
    });
  };

  // Export to JSON
  const handleExportJson = () => {
    if (!scheduleData) return;
    const payload = {
      fileName: `Schedule_${targetYear}_${String(targetMonth + 1).padStart(2, '0')}.json`,
      month: targetMonth,
      year: targetYear,
      employees: Object.keys(scheduleData.rows),
      dateColumns: scheduleData.dateColumns,
      rows: scheduleData.rows
    };
    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `Schedule_${targetYear}_${String(targetMonth + 1).padStart(2, '0')}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const monthDisplay = new Intl.DateTimeFormat('en', { month: 'long', year: 'numeric' }).format(
    new Date(targetYear, targetMonth, 1)
  );

  return (
    <div className="w-full space-y-6 pb-20">
      {/* Top Header & Context */}
      <div className="flex flex-col gap-4 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <div className="flex items-center gap-3">
            <h1 className="text-[28px] font-bold tracking-tight">Schedule Builder</h1>
            <span className="rounded-full bg-blue-500/10 px-2.5 py-0.5 text-[11px] font-semibold text-blue-500 border border-blue-500/20">
              Desktop Pro
            </span>
            {isHistorical && (
              <span className="rounded-full bg-amber-500/10 px-2.5 py-0.5 text-[11px] font-semibold text-amber-600 dark:text-amber-400 border border-amber-500/20 flex items-center gap-1">
                <Lock className="size-3" /> Historical Record (Set in Stone)
              </span>
            )}
          </div>
          <p className="mt-1 text-[13px] text-zinc-500 dark:text-zinc-400">
            {isHistorical
              ? "Archived monthly schedule record loaded directly from roster data (locked & set in stone)"
              : "Automated, balanced schedule generation based on previous months' roster rules"}
          </p>
        </div>

        {/* Month Selector Pill */}
        <div className="flex items-center gap-2">
          <div className="flex items-center gap-1 rounded-2xl border border-zinc-950/10 dark:border-white/10 bg-zinc-950/5 dark:bg-white/5 p-1">
            <button
              type="button"
              onClick={handlePrevMonth}
              className="flex size-8 items-center justify-center rounded-xl text-zinc-600 dark:text-zinc-300 hover:bg-zinc-950/5 dark:hover:bg-white/10 active:scale-95 transition-all cursor-pointer"
              title="Previous month"
            >
              <ChevronLeft className="size-4" />
            </button>
            <span className="px-3 text-[13px] font-bold text-zinc-800 dark:text-zinc-200 whitespace-nowrap min-w-[140px] text-center">
              {monthDisplay}
            </span>
            <button
              type="button"
              onClick={handleNextMonth}
              className="flex size-8 items-center justify-center rounded-xl text-zinc-600 dark:text-zinc-300 hover:bg-zinc-950/5 dark:hover:bg-white/10 active:scale-95 transition-all cursor-pointer"
              title="Next month"
            >
              <ChevronRight className="size-4" />
            </button>
          </div>
        </div>
      </div>

      {/* Applied Banner */}
      {appliedNotification && (
        <div className="rounded-2xl border border-emerald-500/30 bg-emerald-500/10 p-3.5 text-[13px] text-emerald-600 dark:text-emerald-400 flex items-center justify-between shadow-sm animate-in fade-in duration-200">
          <div className="flex items-center gap-2.5">
            <CheckCircle2 className="size-5 shrink-0" />
            <span className="font-semibold">{appliedNotification}</span>
          </div>
          <button
            onClick={() => setAppliedNotification(null)}
            className="text-emerald-500 hover:text-emerald-700 cursor-pointer"
          >
            <X className="size-4" />
          </button>
        </div>
      )}

      {/* Constraints & Overview Banner */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
        {/* Card 1: Fixed Management Rule */}
        <div className={`p-4 ${GLASS_CARD} flex items-start gap-3`}>
          <div className="rounded-xl bg-blue-500/10 p-2 text-blue-500 shrink-0">
            <Shield className="size-5" />
          </div>
          <div className="text-[12px]">
            <p className="font-semibold text-zinc-800 dark:text-zinc-200 text-[13px]">Fixed Management</p>
            <p className="text-zinc-500 dark:text-zinc-400 mt-0.5 leading-relaxed">
              <span className="font-semibold text-blue-500">Alexandru Ionut Palade</span> & <span className="font-semibold text-blue-500">George Tir</span> are locked to <span className="font-semibold">MID</span> Mon–Fri, <span className="font-semibold">OFF</span> Sat–Sun.
            </p>
          </div>
        </div>

        {/* Card 2: 24/7 Coverage Targets */}
        <div className={`p-4 ${GLASS_CARD} flex items-start gap-3`}>
          <div className="rounded-xl bg-emerald-500/10 p-2 text-emerald-500 shrink-0">
            <Clock className="size-5" />
          </div>
          <div className="text-[12px]">
            <p className="font-semibold text-zinc-800 dark:text-zinc-200 text-[13px]">24/7 Shift Coverage</p>
            <div className="mt-1 flex flex-wrap gap-2 text-[11px] font-medium">
              <span className="rounded bg-green-500/15 text-green-700 dark:text-green-300 px-1.5 py-0.5">Morning: 2–3</span>
              <span className="rounded bg-orange-500/15 text-orange-700 dark:text-orange-300 px-1.5 py-0.5">Afternoon: 2–3</span>
              <span className="rounded bg-purple-500/15 text-purple-700 dark:text-purple-300 px-1.5 py-0.5">Night: 2</span>
            </div>
          </div>
        </div>

        {/* Card 3: Legal Compliance Status or Historical Record */}
        {isHistorical ? (
          <div className={`p-4 ${GLASS_CARD} flex items-start gap-3`}>
            <div className="rounded-xl p-2 shrink-0 bg-amber-500/10 text-amber-500">
              <Lock className="size-5" />
            </div>
            <div className="text-[12px]">
              <div className="flex items-center justify-between">
                <p className="font-semibold text-zinc-800 dark:text-zinc-200 text-[13px]">
                  Archived Historical Record
                </p>
                <span className="font-semibold text-[10px] text-amber-600 dark:text-amber-400 bg-amber-500/10 px-2 py-0.5 rounded-full border border-amber-500/20">
                  Set in Stone
                </span>
              </div>
              <p className="text-zinc-500 dark:text-zinc-400 mt-0.5 leading-relaxed">
                Official historical schedule loaded directly from roster data. Locked and non-regenerated.
              </p>
            </div>
          </div>
        ) : (
          <div className={`p-4 ${GLASS_CARD} flex items-start gap-3`}>
            <div className={`rounded-xl p-2 shrink-0 ${validation.isFullyCompliant ? 'bg-emerald-500/10 text-emerald-500' : 'bg-amber-500/10 text-amber-500'}`}>
              {validation.isFullyCompliant ? <CheckCircle2 className="size-5" /> : <AlertTriangle className="size-5" />}
            </div>
            <div className="text-[12px]">
              <div className="flex items-center justify-between">
                <p className="font-semibold text-zinc-800 dark:text-zinc-200 text-[13px]">
                  {validation.isFullyCompliant ? '100% Legal & Balanced' : 'Action Needed'}
                </p>
                <span className="font-mono text-[10px] text-zinc-400">Seed #{scheduleData?.seed}</span>
              </div>
              <p className="text-zinc-500 dark:text-zinc-400 mt-0.5 leading-relaxed">
                {validation.isFullyCompliant
                  ? `All employees meet the monthly norm of ${targetWorkingDays} working days with compliant rests and coverage.`
                  : `${validation.totalViolations} warning(s): review days with alerts.`}
              </p>
            </div>
          </div>
        )}
      </div>

      {/* Main Action Bar */}
      <div className={`p-3.5 ${GLASS_CARD} flex flex-wrap items-center justify-between gap-3`}>
        <div className="flex flex-wrap items-center gap-2">
          {/* Regenerate Variant Button */}
          <button
            type="button"
            disabled={isHistorical}
            onClick={handleRegenerateVariant}
            className={`flex items-center gap-2 rounded-xl font-semibold px-4 py-2 text-[13px] shadow transition-all ${
              isHistorical
                ? 'bg-zinc-200 dark:bg-zinc-800 text-zinc-400 dark:text-zinc-500 cursor-not-allowed opacity-60'
                : 'bg-blue-500 hover:bg-blue-600 text-white active:scale-98 cursor-pointer'
            }`}
            title={isHistorical ? 'Historical records are set in stone and cannot be regenerated' : 'Generate new balanced variant'}
          >
            {isHistorical ? <Lock className="size-4" /> : <RotateCcw className="size-4" />}
            <span>{isHistorical ? 'Locked: Set in Stone' : 'Regenerate Variant'}</span>
          </button>

          {/* Legal Monthly Working Days Norm (Fixed, non-modifiable variable) */}
          <div className="flex items-center gap-1.5 rounded-xl border border-zinc-200/60 dark:border-zinc-800/60 bg-white/50 dark:bg-zinc-800/50 px-3 py-1.5 text-[12px]">
            <span className="text-zinc-500 dark:text-zinc-400 font-medium">Legal Monthly Norm:</span>
            <span className="font-semibold text-blue-600 dark:text-blue-400 font-mono text-[13px] px-1">
              {targetWorkingDays} Days
            </span>
          </div>

          {/* Quick Info text on block variability */}
          <span className="hidden xl:inline text-[12px] text-zinc-400 dark:text-zinc-500 pl-2">
            {isHistorical
              ? 'Archived Monthly Roster Record'
              : 'Blocks of 2, 3, 4, 5 days • Max 2 days rest (rarely 3)'}
          </span>
        </div>

        <div className="flex items-center gap-2">
          {/* Apply to Calendar */}
          {isHistorical ? (
            <button
              type="button"
              disabled
              className="flex items-center gap-2 rounded-xl bg-zinc-200 dark:bg-zinc-800 text-zinc-400 dark:text-zinc-500 font-semibold px-4 py-2 text-[13px] shadow transition-all cursor-not-allowed opacity-60"
              title="Historical shifts are already saved in the active roster"
            >
              <Save className="size-4" />
              <span>Already in Calendar</span>
            </button>
          ) : (
            <button
              type="button"
              onClick={handleApplyToCalendar}
              className="flex items-center gap-2 rounded-xl bg-emerald-600 hover:bg-emerald-700 text-white font-semibold px-4 py-2 text-[13px] shadow transition-all active:scale-98 cursor-pointer"
              title="Integrate this generated month into live calendar"
            >
              <Save className="size-4" />
              <span>Apply to Calendar</span>
            </button>
          )}

          {/* Export Excel */}
          <button
            type="button"
            onClick={handleExportExcel}
            className="flex items-center gap-2 rounded-xl border border-zinc-950/10 dark:border-white/10 hover:bg-zinc-950/5 dark:hover:bg-white/5 font-semibold px-3 py-2 text-[13px] transition-colors cursor-pointer"
            title="Download formatted Excel workbook"
          >
            <FileSpreadsheet className="size-4 text-emerald-500" />
            <span className="hidden sm:inline">Excel (.xlsx)</span>
          </button>

          {/* Export JSON */}
          <button
            type="button"
            onClick={handleExportJson}
            className="flex items-center gap-2 rounded-xl border border-zinc-950/10 dark:border-white/10 hover:bg-zinc-950/5 dark:hover:bg-white/5 font-semibold px-3 py-2 text-[13px] transition-colors cursor-pointer"
            title="Download JSON schedule data"
          >
            <Download className="size-4 text-zinc-400" />
            <span className="hidden sm:inline">JSON</span>
          </button>
        </div>
      </div>

      {/* Multi-Select Date & Replace / Swap Employee Toolbar */}
      {isHistorical ? (
        <div className={`p-4 ${GLASS_CARD} flex items-center justify-between`}>
          <div className="flex items-center gap-2.5 text-zinc-500 dark:text-zinc-400 text-[13px]">
            <Lock className="size-4 text-amber-500" />
            <span>This is an official historical roster record (set in stone). Shift modifications and swaps are disabled.</span>
          </div>
          <span className="rounded-full bg-zinc-500/10 px-2.5 py-0.5 text-[11px] font-semibold text-zinc-500 border border-zinc-500/20">
            Read Only
          </span>
        </div>
      ) : (
        <div className={`p-4 ${GLASS_CARD} space-y-3 border-2 ${selectedDates.size > 0 ? 'border-blue-500/30 bg-blue-500/[0.02]' : 'border-transparent'}`}>
          <div className="flex flex-wrap items-center justify-between gap-3">
            <div className="flex items-center gap-2">
              <ArrowLeftRight className="size-5 text-blue-500" />
              <h2 className="text-[14px] font-bold">Swap & Replace Shifts</h2>
              <span className="text-[12px] text-zinc-400">
                {selectedDates.size > 0
                  ? `(${selectedDates.size} date${selectedDates.size === 1 ? '' : 's'} selected)`
                  : '— Click column dates or use quick range'}
              </span>
            </div>

            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={selectAllWeekdays}
                className="rounded-lg bg-zinc-950/5 dark:bg-white/5 hover:bg-zinc-950/10 dark:hover:bg-white/10 px-2 py-1 text-[11px] font-semibold text-zinc-600 dark:text-zinc-400 cursor-pointer"
              >
                + Weekdays
              </button>
              <button
                type="button"
                onClick={selectAllWeekends}
                className="rounded-lg bg-zinc-950/5 dark:bg-white/5 hover:bg-zinc-950/10 dark:hover:bg-white/10 px-2 py-1 text-[11px] font-semibold text-zinc-600 dark:text-zinc-400 cursor-pointer"
              >
                + Weekends
              </button>
              {selectedDates.size > 0 && (
                <button
                  type="button"
                  onClick={clearSelection}
                  className="text-[11px] font-semibold text-rose-500 hover:text-rose-600 cursor-pointer ml-1"
                >
                  Clear
                </button>
              )}
            </div>
          </div>

          {/* Swap Controls Row */}
          <div className="flex flex-wrap items-center gap-3">
            <div className="flex items-center gap-2">
              <span className="text-[12px] text-zinc-500 font-medium">Employee A:</span>
              <select
                value={swapEmployeeA}
                onChange={(e) => setSwapEmployeeA(e.target.value)}
                className="rounded-xl border border-zinc-950/10 dark:border-white/10 bg-zinc-950/5 dark:bg-white/5 px-3 py-1.5 text-[12px] font-semibold outline-none focus:ring-2 focus:ring-blue-500/30"
              >
                <option value="">Select Employee...</option>
                {rotatingEmployees.map((emp) => (
                  <option key={emp} value={emp}>
                    {emp}
                  </option>
                ))}
              </select>
            </div>

            <span className="text-zinc-400">⇄</span>

            <div className="flex items-center gap-2">
              <span className="text-[12px] text-zinc-500 font-medium">Employee B:</span>
              <select
                value={swapEmployeeB}
                onChange={(e) => setSwapEmployeeB(e.target.value)}
                className="rounded-xl border border-zinc-950/10 dark:border-white/10 bg-zinc-950/5 dark:bg-white/5 px-3 py-1.5 text-[12px] font-semibold outline-none focus:ring-2 focus:ring-blue-500/30"
              >
                <option value="">Select Employee...</option>
                {rotatingEmployees.map((emp) => (
                  <option key={emp} value={emp}>
                    {emp}
                  </option>
                ))}
              </select>
            </div>

            <button
              type="button"
              disabled={!swapEmployeeA || !swapEmployeeB || selectedDates.size === 0}
              onClick={handleExecuteSwap}
              className="flex items-center gap-2 rounded-xl bg-blue-500 hover:bg-blue-600 disabled:opacity-40 text-white font-semibold px-4 py-1.5 text-[12px] shadow transition-all active:scale-98 cursor-pointer"
            >
              <ArrowLeftRight className="size-3.5" />
              <span>Swap on {selectedDates.size} Day(s)</span>
            </button>
          </div>

          {/* Swap Feedback Message */}
          {swapFeedback && (
            <div className={`rounded-xl p-2.5 text-[12px] flex items-center gap-2 ${
              swapFeedback.type === 'success'
                ? 'bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border border-emerald-500/20'
                : 'bg-amber-500/10 text-amber-600 dark:text-amber-400 border border-amber-500/20'
            }`}>
              {swapFeedback.type === 'success' ? <CheckCircle2 className="size-4 shrink-0" /> : <AlertTriangle className="size-4 shrink-0" />}
              <span>{swapFeedback.message}</span>
            </div>
          )}
        </div>
      )}

      {/* Interactive Schedule Matrix Grid */}
      <div className={`overflow-hidden ${GLASS_CARD}`}>
        <div className="overflow-x-auto select-none">
          <table className="w-full text-left border-collapse text-[11px]">
            {/* Table Header: Days of the Month */}
            <thead>
              <tr className="border-b border-zinc-950/10 dark:border-white/10 bg-zinc-950/[0.03] dark:bg-white/[0.03]">
                {/* Employee Column Header */}
                <th className="sticky left-0 z-20 min-w-[190px] p-2.5 font-bold text-zinc-600 dark:text-zinc-400 bg-zinc-100 dark:bg-[#121212] border-r border-zinc-950/10 dark:border-white/10">
                  Employee ({employees.length})
                </th>

                {/* Day Columns */}
                {scheduleData?.dateColumns.map(({ isoDate, date }) => {
                  const dayNum = date.getDate();
                  const dayName = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][date.getDay()];
                  const isWeekend = date.getDay() === 0 || date.getDay() === 6;
                  const isSelected = selectedDates.has(isoDate);

                  return (
                    <th
                      key={isoDate}
                      onClick={() => !isHistorical && toggleDateSelection(isoDate)}
                      className={`min-w-[42px] max-w-[42px] p-1 text-center font-semibold transition-colors border-r border-zinc-950/5 dark:border-white/5 ${
                        isHistorical
                          ? isWeekend
                            ? 'bg-zinc-950/[0.05] dark:bg-white/[0.05] text-zinc-500 cursor-default'
                            : 'text-zinc-700 dark:text-zinc-300 cursor-default'
                          : isSelected
                            ? 'bg-blue-500/20 text-blue-600 dark:text-blue-400 font-bold cursor-pointer'
                            : isWeekend
                              ? 'bg-zinc-950/[0.05] dark:bg-white/[0.05] text-zinc-500 cursor-pointer'
                              : 'text-zinc-700 dark:text-zinc-300 hover:bg-zinc-950/[0.03] dark:hover:bg-white/[0.03] cursor-pointer'
                      }`}
                      title={isHistorical ? `${isoDate} (Historical record)` : `Click to select/deselect ${isoDate}`}
                    >
                      <div className="text-[9px] uppercase tracking-tighter opacity-70">{dayName}</div>
                      <div className="text-[12px] font-bold leading-tight">{dayNum}</div>
                    </th>
                  );
                })}

                {/* Summary Columns Header */}
                <th className="sticky right-0 z-20 min-w-[140px] p-2 font-bold text-zinc-600 dark:text-zinc-400 bg-zinc-100 dark:bg-[#121212] border-l border-zinc-950/10 dark:border-white/10 text-center">
                  Monthly Stats
                </th>
              </tr>
            </thead>

            {/* Table Body: Employee Rows */}
            <tbody className="divide-y divide-zinc-950/5 dark:divide-white/5">
              {employees.map((emp) => {
                const isFixedMid = FIXED_MID_EMPLOYEES.includes(emp);
                const stats = validation.employeeStats[emp];

                return (
                  <tr
                    key={emp}
                    className="hover:bg-zinc-950/[0.02] dark:hover:bg-white/[0.02] transition-colors"
                  >
                    {/* Sticky Employee Name Cell */}
                    <td className="sticky left-0 z-10 p-2 font-medium bg-zinc-50/90 dark:bg-[#101010]/95 backdrop-blur border-r border-zinc-950/10 dark:border-white/10">
                      <div className="flex items-center gap-2">
                        <div className={`flex size-6 shrink-0 items-center justify-center rounded-lg text-[10px] font-bold ${
                          isFixedMid ? 'bg-blue-500/15 text-blue-500' : 'bg-zinc-500/15 text-zinc-600 dark:text-zinc-400'
                        }`}>
                          {initials(emp)}
                        </div>
                        <div className="min-w-0">
                          <p className="truncate font-semibold text-[12px] text-zinc-900 dark:text-zinc-100">
                            {emp}
                          </p>
                          {isFixedMid && (
                            <span className="text-[9px] font-bold text-blue-500 uppercase">Fixed Mid</span>
                          )}
                        </div>
                      </div>
                    </td>

                    {/* Shift Cells */}
                    {scheduleData?.dateColumns.map(({ isoDate, date }) => {
                      const rawShift = scheduleData.rows[emp]?.[isoDate] ?? 'OFF';
                      const sKey = shiftKey(rawShift);
                      const isWeekend = date.getDay() === 0 || date.getDay() === 6;
                      const isDateSelected = selectedDates.has(isoDate);
                      const isEditing = editingCell?.employee === emp && editingCell?.isoDate === isoDate;

                      // Badge styles
                      const colors = shiftColors[sKey] ?? {
                        bg: 'bg-zinc-500/10',
                        text: 'text-zinc-500'
                      };

                      return (
                        <td
                          key={isoDate}
                          onClick={() => {
                            if (!isHistorical && !isFixedMid) {
                              setEditingCell({ employee: emp, isoDate });
                            }
                          }}
                          className={`p-1 text-center border-r border-zinc-950/5 dark:border-white/5 relative ${
                            isHistorical || isFixedMid ? 'cursor-default' : 'cursor-pointer'
                          } ${
                            isDateSelected ? 'bg-blue-500/[0.06]' : isWeekend ? 'bg-zinc-950/[0.015] dark:bg-white/[0.015]' : ''
                          }`}
                          title={`${emp} on ${isoDate}: ${rawShift}${isHistorical ? ' (Historical Record - Set in Stone)' : isFixedMid ? ' (Fixed MID)' : ' (Click to edit)'}`}
                        >
                          <div
                            className={`inline-flex items-center justify-center size-6 rounded-lg font-bold text-[10px] transition-transform ${
                              isHistorical || isFixedMid ? '' : 'hover:scale-110 active:scale-95'
                            } ${
                              sKey === 'OFF'
                                ? 'text-zinc-300 dark:text-zinc-600 hover:bg-zinc-950/5 dark:hover:bg-white/5'
                                : `${colors.bg} ${colors.text} shadow-xs`
                            }`}
                          >
                            {sKey === 'OFF' ? '·' : sKey}
                          </div>

                          {/* Quick Shift Edit Popover */}
                          {isEditing && (
                            <div
                              onClick={(e) => e.stopPropagation()}
                              className="absolute z-50 top-full left-1/2 -translate-x-1/2 mt-1 rounded-xl bg-white dark:bg-[#1a1a1a] shadow-xl border border-zinc-950/10 dark:border-white/10 p-2 flex gap-1 animate-in fade-in zoom-in-95"
                            >
                              {['M', 'A', 'N', 'MID', 'OFF', 'H8'].map((code) => (
                                <button
                                  key={code}
                                  onClick={() => handleSetCellShift(code)}
                                  className={`px-2 py-1 rounded-lg text-[10px] font-bold cursor-pointer transition-colors ${
                                    sKey === code ? 'bg-blue-500 text-white' : 'hover:bg-zinc-100 dark:hover:bg-white/10'
                                  }`}
                                >
                                  {code}
                                </button>
                              ))}
                              <button
                                onClick={() => setEditingCell(null)}
                                className="px-1.5 py-1 text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200 cursor-pointer"
                              >
                                <X className="size-3" />
                              </button>
                            </div>
                          )}
                        </td>
                      );
                    })}

                    {/* Summary Stats Cell */}
                    <td className="sticky right-0 z-10 p-2 text-center bg-zinc-50/90 dark:bg-[#101010]/95 backdrop-blur border-l border-zinc-950/10 dark:border-white/10 font-mono text-[10px]">
                      {stats && (
                        <div className="flex items-center justify-center gap-1.5">
                          <span className="text-green-600 dark:text-green-400 font-bold" title="Morning shifts">{stats.M}M</span>
                          <span className="text-orange-600 dark:text-orange-400 font-bold" title="Afternoon shifts">{stats.A}A</span>
                          <span className="text-purple-600 dark:text-purple-400 font-bold" title="Night shifts">{stats.N}N</span>
                          {stats.MID > 0 && <span className="text-blue-500 font-bold" title="Mid shifts">{stats.MID}MID</span>}
                          <span className="text-zinc-400" title="Off days">{stats.OFF}O</span>
                          <span
                            className={`font-bold px-1.5 py-0.5 rounded text-[10px] ${
                              stats.totalWorkingDays === stats.targetWorkingDays
                                ? 'bg-emerald-500/15 text-emerald-600 dark:text-emerald-400'
                                : 'bg-amber-500/15 text-amber-600 dark:text-amber-400'
                            }`}
                            title={`Total working days: ${stats.totalWorkingDays}/${stats.targetWorkingDays} required`}
                          >
                            {stats.totalWorkingDays}/{stats.targetWorkingDays}d
                          </span>
                        </div>
                      )}
                    </td>
                  </tr>
                );
              })}
            </tbody>

            {/* Table Footer: Daily Coverage Rows */}
            <tfoot className="border-t-2 border-zinc-950/10 dark:border-white/10 bg-zinc-950/[0.04] dark:bg-white/[0.04] font-semibold">
              {/* Morning Coverage Row */}
              <tr>
                <td className="sticky left-0 z-10 p-2 font-bold text-green-700 dark:text-green-400 bg-zinc-100 dark:bg-[#141414] border-r border-zinc-950/10 dark:border-white/10">
                  Morning (M)
                </td>
                {scheduleData?.dateColumns.map(({ isoDate }) => {
                  const cov = validation.dailyCoverage[isoDate];
                  const under = cov ? cov.M < 2 : false;
                  return (
                    <td
                      key={isoDate}
                      className={`p-1 text-center font-mono text-[11px] border-r border-zinc-950/5 dark:border-white/5 ${
                        under ? 'bg-rose-500/20 text-rose-600 font-bold' : 'text-green-700 dark:text-green-400'
                      }`}
                    >
                      {cov?.M ?? 0}
                    </td>
                  );
                })}
                <td className="sticky right-0 z-10 p-1 text-center bg-zinc-100 dark:bg-[#141414] border-l border-zinc-950/10 dark:border-white/10 font-bold text-green-700 dark:text-green-400">
                  M ≥ 2
                </td>
              </tr>

              {/* Afternoon Coverage Row */}
              <tr>
                <td className="sticky left-0 z-10 p-2 font-bold text-orange-700 dark:text-orange-400 bg-zinc-100 dark:bg-[#141414] border-r border-zinc-950/10 dark:border-white/10">
                  Afternoon (A)
                </td>
                {scheduleData?.dateColumns.map(({ isoDate }) => {
                  const cov = validation.dailyCoverage[isoDate];
                  const under = cov ? cov.A < 2 : false;
                  return (
                    <td
                      key={isoDate}
                      className={`p-1 text-center font-mono text-[11px] border-r border-zinc-950/5 dark:border-white/5 ${
                        under ? 'bg-rose-500/20 text-rose-600 font-bold' : 'text-orange-700 dark:text-orange-400'
                      }`}
                    >
                      {cov?.A ?? 0}
                    </td>
                  );
                })}
                <td className="sticky right-0 z-10 p-1 text-center bg-zinc-100 dark:bg-[#141414] border-l border-zinc-950/10 dark:border-white/10 font-bold text-orange-700 dark:text-orange-400">
                  A ≥ 2
                </td>
              </tr>

              {/* Night Coverage Row */}
              <tr>
                <td className="sticky left-0 z-10 p-2 font-bold text-purple-700 dark:text-purple-400 bg-zinc-100 dark:bg-[#141414] border-r border-zinc-950/10 dark:border-white/10">
                  Night (N)
                </td>
                {scheduleData?.dateColumns.map(({ isoDate }) => {
                  const cov = validation.dailyCoverage[isoDate];
                  const under = cov ? cov.N < 2 : false;
                  return (
                    <td
                      key={isoDate}
                      className={`p-1 text-center font-mono text-[11px] border-r border-zinc-950/5 dark:border-white/5 ${
                        under ? 'bg-rose-500/20 text-rose-600 font-bold' : 'text-purple-700 dark:text-purple-400'
                      }`}
                    >
                      {cov?.N ?? 0}
                    </td>
                  );
                })}
                <td className="sticky right-0 z-10 p-1 text-center bg-zinc-100 dark:bg-[#141414] border-l border-zinc-950/10 dark:border-white/10 font-bold text-purple-700 dark:text-purple-400">
                  N ≥ 2
                </td>
              </tr>

              {/* Mid Coverage Row */}
              <tr>
                <td className="sticky left-0 z-10 p-2 font-bold text-blue-600 dark:text-blue-400 bg-zinc-100 dark:bg-[#141414] border-r border-zinc-950/10 dark:border-white/10">
                  Mid (MID)
                </td>
                {scheduleData?.dateColumns.map(({ isoDate }) => {
                  const cov = validation.dailyCoverage[isoDate];
                  return (
                    <td
                      key={isoDate}
                      className="p-1 text-center font-mono text-[11px] border-r border-zinc-950/5 dark:border-white/5 text-blue-600 dark:text-blue-400"
                    >
                      {cov?.MID ?? 0}
                    </td>
                  );
                })}
                <td className="sticky right-0 z-10 p-1 text-center bg-zinc-100 dark:bg-[#141414] border-l border-zinc-950/10 dark:border-white/10 font-bold text-blue-600 dark:text-blue-400">
                  2 (Mon–Fri)
                </td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

      {/* Warnings & Diagnostics Panel (if any) */}
      {validation.warnings.length > 0 && (
        <div className={`p-4 ${GLASS_CARD} border-amber-500/30 bg-amber-500/5 space-y-2`}>
          <div className="flex items-center gap-2 text-amber-600 dark:text-amber-400 font-bold text-[13px]">
            <AlertTriangle className="size-4.5" />
            <span>Validation Warnings ({validation.warnings.length})</span>
          </div>
          <ul className="list-disc list-inside text-[12px] text-zinc-600 dark:text-zinc-400 space-y-1">
            {validation.warnings.slice(0, 8).map((w, i) => (
              <li key={i}>{w}</li>
            ))}
            {validation.warnings.length > 8 && (
              <li className="font-semibold text-zinc-500">
                ...and {validation.warnings.length - 8} more warnings
              </li>
            )}
          </ul>
        </div>
      )}
    </div>
  );
}
