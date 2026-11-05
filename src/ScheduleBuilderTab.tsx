import { useState, useMemo, useEffect, useCallback, useRef } from 'react';
import {
  RotateCcw,
  Save,
  FileSpreadsheet,
  Download,
  ChevronLeft,
  ChevronRight,
  AlertTriangle,
  CheckCircle2,
  Shield,
  Lock,
  Clock,
  X,
  Loader2,
  MousePointerClick
} from 'lucide-react';
import type { RosterData } from './types';
import {
  generateBalancedSchedule,
  validateScheduleCompliance,
  exportRosterToExcelFile,
  countLegalWorkingDaysInMonth,
  isHistoricalMonthInRoster,
  loadHistoricalSchedule,
  FIXED_MID_EMPLOYEES,
  FIXED_WEEKDAY_MA_EMPLOYEES,
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

  // Multi-cell selection state for block changes
  const [selectedCells, setSelectedCells] = useState<Set<string>>(new Set());
  const [lastClickedCell, setLastClickedCell] = useState<{ employee: string; isoDate: string } | null>(null);
  const [bulkFeedback, setBulkFeedback] = useState<string | null>(null);
  const isMouseDownRef = useRef(false);

  // Clear dragging on mouseup
  useEffect(() => {
    const handleMouseUp = () => {
      isMouseDownRef.current = false;
    };
    window.addEventListener('mouseup', handleMouseUp);
    return () => window.removeEventListener('mouseup', handleMouseUp);
  }, []);

  // Single cell editing popover
  const [editingCell, setEditingCell] = useState<{ employee: string; isoDate: string } | null>(null);
  const popoverRef = useRef<HTMLDivElement>(null);
  const [hasUserNavigated, setHasUserNavigated] = useState(false);

  // Close editing popover when clicking outside
  useEffect(() => {
    if (!editingCell) return;
    const handleClickOutside = (e: MouseEvent | PointerEvent) => {
      if (popoverRef.current && !popoverRef.current.contains(e.target as Node)) {
        setEditingCell(null);
      }
    };
    const timer = setTimeout(() => {
      document.addEventListener('pointerdown', handleClickOutside);
    }, 0);
    return () => {
      clearTimeout(timer);
      document.removeEventListener('pointerdown', handleClickOutside);
    };
  }, [editingCell]);

  // Sync initial month and year once roster data loads if user hasn't explicitly navigated
  useEffect(() => {
    if (!hasUserNavigated && roster && roster.dateColumns.length > 0) {
      const last = roster.dateColumns[roster.dateColumns.length - 1];
      const d = new Date(`${last.isoDate}T00:00:00`);
      setTargetYear(d.getFullYear());
      setTargetMonth(d.getMonth());
    }
  }, [roster, hasUserNavigated]);

  const [isGenerating, setIsGenerating] = useState(false);

  // Generate initial schedule or load historical schedule if already set in stone
  const runGeneration = useCallback(
    (seedValue: number = currentSeed, y: number = targetYear, m: number = targetMonth) => {
      setIsGenerating(true);
      setTimeout(() => {
        if (isHistoricalMonthInRoster(roster, y, m)) {
          setScheduleData(loadHistoricalSchedule(roster!, y, m));
        } else {
          const result = generateBalancedSchedule(roster, y, m, {
            seed: seedValue,
            targetWorkingDays: countLegalWorkingDaysInMonth(m, y)
          });
          setScheduleData(result);
        }
        setSelectedCells(new Set());
        setBulkFeedback(null);
        setIsGenerating(false);
      }, 20);
    },
    [roster, currentSeed, targetYear, targetMonth]
  );

  useEffect(() => {
    runGeneration(currentSeed, targetYear, targetMonth);
  }, [targetYear, targetMonth, targetWorkingDays, roster]);

  // Live validation whenever schedule rows change
  const validation: ScheduleValidationReport = useMemo(() => {
    if (!scheduleData) {
      return {
        isFullyCompliant: true,
        totalViolations: 0,
        understaffedDaysCount: 0,
        nightOverstaffedDaysCount: 0,
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
    if (scheduleData.validation) {
      return scheduleData.validation;
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

  // Handlers for month navigation
  const handlePrevMonth = () => {
    if (isGenerating) return;
    setHasUserNavigated(true);
    const prev = new Date(targetYear, targetMonth - 1, 1);
    setTargetMonth(prev.getMonth());
    setTargetYear(prev.getFullYear());
  };

  const handleNextMonth = () => {
    if (isGenerating) return;
    setHasUserNavigated(true);
    const next = new Date(targetYear, targetMonth + 1, 1);
    setTargetMonth(next.getMonth());
    setTargetYear(next.getFullYear());
  };

  const handleRegenerateVariant = () => {
    if (isHistorical || isGenerating) return;
    const nextSeed = Math.floor(Math.random() * 999999);
    setCurrentSeed(nextSeed);
    runGeneration(nextSeed, targetYear, targetMonth);
  };

  // Multi-cell selection & block change handlers
  const handleCellMouseDown = (e: React.MouseEvent, emp: string, isoDate: string) => {
    if ((e.target as HTMLElement).closest('.shift-popover')) return;

    const cellKey = `${emp}::${isoDate}`;
    isMouseDownRef.current = true;

    if (e.ctrlKey || e.metaKey) {
      // Holding Ctrl / Cmd: Toggle cell into/out of multi-selection
      setEditingCell(null);
      setSelectedCells((prev) => {
        const next = new Set(prev);
        if (next.has(cellKey)) {
          next.delete(cellKey);
        } else {
          next.add(cellKey);
        }
        return next;
      });
      setLastClickedCell({ employee: emp, isoDate });
    } else if (e.shiftKey && lastClickedCell && scheduleData) {
      // Holding Shift: Select rectangular block range between last clicked cell and current
      setEditingCell(null);
      const allDates = scheduleData.dateColumns.map((c) => c.isoDate);
      const allEmps = Object.keys(scheduleData.rows);

      const startEmpIdx = allEmps.indexOf(lastClickedCell.employee);
      const endEmpIdx = allEmps.indexOf(emp);
      const startDateIdx = allDates.indexOf(lastClickedCell.isoDate);
      const endDateIdx = allDates.indexOf(isoDate);

      if (startEmpIdx !== -1 && endEmpIdx !== -1 && startDateIdx !== -1 && endDateIdx !== -1) {
        const [lowEmp, highEmp] = startEmpIdx <= endEmpIdx ? [startEmpIdx, endEmpIdx] : [endEmpIdx, startEmpIdx];
        const [lowDate, highDate] = startDateIdx <= endDateIdx ? [startDateIdx, endDateIdx] : [endDateIdx, startDateIdx];

        setSelectedCells((prev) => {
          const next = new Set(prev);
          for (let eIdx = lowEmp; eIdx <= highEmp; eIdx++) {
            for (let dIdx = lowDate; dIdx <= highDate; dIdx++) {
              next.add(`${allEmps[eIdx]}::${allDates[dIdx]}`);
            }
          }
          return next;
        });
      }
    } else {
      // Normal click without Ctrl/Shift:
      if (selectedCells.size > 0 && selectedCells.has(cellKey)) {
        return;
      }
      setSelectedCells(new Set());
      setEditingCell({ employee: emp, isoDate });
      setLastClickedCell({ employee: emp, isoDate });
    }
  };

  const handleCellMouseEnter = (e: React.MouseEvent, emp: string, isoDate: string) => {
    if (isMouseDownRef.current && (e.ctrlKey || e.metaKey)) {
      const cellKey = `${emp}::${isoDate}`;
      setSelectedCells((prev) => {
        const next = new Set(prev);
        next.add(cellKey);
        return next;
      });
    }
  };

  // Toggle selection for an entire day column
  const handleToggleColumnSelection = (isoDate: string, isCtrl: boolean) => {
    if (!scheduleData) return;
    const emps = Object.keys(scheduleData.rows);
    const colKeys = emps.map((emp) => `${emp}::${isoDate}`);
    setSelectedCells((prev) => {
      const next = isCtrl ? new Set(prev) : new Set<string>();
      const allSelected = colKeys.every((k) => prev.has(k));
      if (allSelected && isCtrl) {
        colKeys.forEach((k) => next.delete(k));
      } else {
        colKeys.forEach((k) => next.add(k));
      }
      return next;
    });
    setEditingCell(null);
  };

  // Toggle selection for an entire employee row
  const handleToggleRowSelection = (emp: string, isCtrl: boolean) => {
    if (!scheduleData) return;
    const rowKeys = scheduleData.dateColumns.map((c) => `${emp}::${c.isoDate}`);
    setSelectedCells((prev) => {
      const next = isCtrl ? new Set(prev) : new Set<string>();
      const allSelected = rowKeys.every((k) => prev.has(k));
      if (allSelected && isCtrl) {
        rowKeys.forEach((k) => next.delete(k));
      } else {
        rowKeys.forEach((k) => next.add(k));
      }
      return next;
    });
    setEditingCell(null);
  };

  // Apply shift code to all currently selected cells (block change)
  const handleApplyShiftToSelected = (shiftCode: string) => {
    if (!scheduleData || selectedCells.size === 0) return;

    const updatedRows = { ...scheduleData.rows };
    selectedCells.forEach((cellKey) => {
      const [emp, isoDate] = cellKey.split('::');
      if (updatedRows[emp]) {
        updatedRows[emp] = {
          ...updatedRows[emp],
          [isoDate]: shiftCode
        };
      }
    });

    const testReport = validateScheduleCompliance(
      updatedRows,
      scheduleData.dateColumns,
      roster,
      targetYear,
      targetMonth,
      targetWorkingDays
    );

    setScheduleData({
      ...scheduleData,
      rows: updatedRows,
      validation: testReport
    });

    setBulkFeedback(`Changed ${selectedCells.size} shift(s) to ${shiftCode}`);
    setTimeout(() => setBulkFeedback(null), 3000);
  };

  const handleClearSelectedCells = () => {
    setSelectedCells(new Set());
    setEditingCell(null);
  };

  // Keyboard shortcuts for block shifts and deselecting
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement) return;

      if (e.key === 'Escape') {
        setSelectedCells(new Set());
        setEditingCell(null);
        return;
      }

      if (selectedCells.size === 0) return;

      const key = e.key.toUpperCase();
      if (key === 'M') {
        e.preventDefault();
        handleApplyShiftToSelected('M');
      } else if (key === 'A') {
        e.preventDefault();
        handleApplyShiftToSelected('A');
      } else if (key === 'N') {
        e.preventDefault();
        handleApplyShiftToSelected('N');
      } else if (key === 'O' || e.key === 'Delete' || e.key === 'Backspace') {
        e.preventDefault();
        handleApplyShiftToSelected('OFF');
      } else if (key === 'H') {
        e.preventDefault();
        handleApplyShiftToSelected('H8');
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [selectedCells, scheduleData, roster, targetYear, targetMonth, targetWorkingDays]);

  // Manual shift override
  const handleSetCellShift = (shiftCode: string) => {
    if (!editingCell || !scheduleData) return;
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
      targetMonth,
      targetWorkingDays
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
            disabled={isHistorical || isGenerating}
            onClick={handleRegenerateVariant}
            className={`flex items-center gap-2 rounded-xl font-semibold px-4 py-2 text-[13px] shadow transition-all ${
              isHistorical || isGenerating
                ? 'bg-zinc-200 dark:bg-zinc-800 text-zinc-400 dark:text-zinc-500 cursor-not-allowed opacity-60'
                : 'bg-blue-500 hover:bg-blue-600 text-white active:scale-98 cursor-pointer'
            }`}
            title={isHistorical ? 'Historical records are set in stone and cannot be regenerated' : 'Generate new balanced variant'}
          >
            {isHistorical ? (
              <Lock className="size-4" />
            ) : isGenerating ? (
              <Loader2 className="size-4 animate-spin" />
            ) : (
              <RotateCcw className="size-4" />
            )}
            <span>{isHistorical ? 'Locked: Set in Stone' : isGenerating ? 'Generating...' : 'Regenerate Variant'}</span>
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

      {/* Quick Help & Status Bar */}
      <div className={`p-3 px-4 ${GLASS_CARD} flex flex-wrap items-center justify-between gap-2.5 text-[12px]`}>
        <div className="flex items-center gap-2 text-zinc-600 dark:text-zinc-400">
          <MousePointerClick className="size-4 text-blue-500 shrink-0" />
          <span>
            <strong className="text-zinc-800 dark:text-zinc-200">Block Selection:</strong> Hold <kbd className="px-1.5 py-0.5 rounded bg-zinc-950/10 dark:bg-white/10 font-mono text-[11px] text-blue-500 font-bold">Ctrl</kbd> (or <kbd className="px-1.5 py-0.5 rounded bg-zinc-950/10 dark:bg-white/10 font-mono text-[11px] text-blue-500 font-bold">⌘</kbd>) and click/drag to select multiple shifts, or use <kbd className="px-1.5 py-0.5 rounded bg-zinc-950/10 dark:bg-white/10 font-mono text-[11px] font-bold">Shift+Click</kbd> for a range.
          </span>
        </div>
        {bulkFeedback ? (
          <span className="font-semibold text-emerald-600 dark:text-emerald-400 animate-in fade-in flex items-center gap-1.5">
            <CheckCircle2 className="size-4 shrink-0" />
            {bulkFeedback}
          </span>
        ) : (
          <span className="text-zinc-400 text-[11px]">
            Hotkeys: <span className="font-mono font-semibold">M</span>, <span className="font-mono font-semibold">A</span>, <span className="font-mono font-semibold">N</span>, <span className="font-mono font-semibold">O</span> (off), <span className="font-mono font-semibold">H</span> (h8)
          </span>
        )}
      </div>


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
                {/* Day Columns */}
                {scheduleData?.dateColumns.map(({ isoDate, date }) => {
                  const dayNum = date.getDate();
                  const dayName = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][date.getDay()];
                  const isWeekend = date.getDay() === 0 || date.getDay() === 6;

                  return (
                    <th
                      key={isoDate}
                      onClick={(e) => handleToggleColumnSelection(isoDate, e.ctrlKey || e.metaKey)}
                      className={`min-w-[42px] max-w-[42px] p-1 text-center font-semibold transition-colors border-r border-zinc-950/5 dark:border-white/5 cursor-pointer ${
                        isWeekend
                          ? 'bg-zinc-950/[0.05] dark:bg-white/[0.05] text-zinc-500 hover:bg-zinc-950/10 dark:hover:bg-white/10'
                          : 'text-zinc-700 dark:text-zinc-300 hover:bg-zinc-950/[0.03] dark:hover:bg-white/[0.03]'
                      }`}
                      title={`Click to select/deselect all shifts on ${isoDate}`}
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
                const isWeekdayMA = FIXED_WEEKDAY_MA_EMPLOYEES.includes(emp);
                const stats = validation.employeeStats[emp];

                return (
                  <tr
                    key={emp}
                    className="hover:bg-zinc-950/[0.02] dark:hover:bg-white/[0.02] transition-colors"
                  >
                    {/* Sticky Employee Name Cell */}
                    <td
                      onClick={(e) => handleToggleRowSelection(emp, e.ctrlKey || e.metaKey)}
                      className="sticky left-0 z-10 p-2 font-medium bg-zinc-50/90 dark:bg-[#101010]/95 backdrop-blur border-r border-zinc-950/10 dark:border-white/10 cursor-pointer hover:bg-zinc-100 dark:hover:bg-[#181818] transition-colors"
                      title={`Click to select/deselect all shifts for ${emp}`}
                    >
                      <div className="flex items-center gap-2">
                        <div className={`flex size-6 shrink-0 items-center justify-center rounded-lg text-[10px] font-bold ${
                          isFixedMid
                            ? 'bg-blue-500/15 text-blue-500'
                            : isWeekdayMA
                            ? 'bg-emerald-500/15 text-emerald-500'
                            : 'bg-zinc-500/15 text-zinc-600 dark:text-zinc-400'
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
                          {isWeekdayMA && (
                            <span className="text-[9px] font-bold text-emerald-500 uppercase">Weekday M/A</span>
                          )}
                        </div>
                      </div>
                    </td>

                    {/* Shift Cells */}
                    {scheduleData?.dateColumns.map(({ isoDate, date }) => {
                      const rawShift = scheduleData.rows[emp]?.[isoDate] ?? 'OFF';
                      const sKey = shiftKey(rawShift);
                      const isWeekend = date.getDay() === 0 || date.getDay() === 6;
                      const cellKey = `${emp}::${isoDate}`;
                      const isCellSelected = selectedCells.has(cellKey);
                      const isEditing = editingCell?.employee === emp && editingCell?.isoDate === isoDate;

                      // Badge styles
                      const colors = shiftColors[sKey] ?? {
                        bg: 'bg-zinc-500/10',
                        text: 'text-zinc-500'
                      };

                      return (
                        <td
                          key={isoDate}
                          onMouseDown={(e) => handleCellMouseDown(e, emp, isoDate)}
                          onMouseEnter={(e) => handleCellMouseEnter(e, emp, isoDate)}
                          className={`p-1 text-center border-r border-zinc-950/5 dark:border-white/5 relative cursor-pointer select-none transition-all ${
                            isCellSelected
                              ? 'bg-blue-500/25 dark:bg-blue-500/35 ring-2 ring-blue-500 ring-inset z-20'
                              : isEditing
                                ? 'z-30 ring-2 ring-blue-500/50 rounded-lg'
                                : isWeekend
                                  ? 'bg-zinc-950/[0.015] dark:bg-white/[0.015]'
                                  : ''
                          }`}
                          title={`${emp} on ${isoDate}: ${rawShift} (Click to edit, Ctrl+click to multi-select)`}
                        >
                          <div
                            className={`inline-flex items-center justify-center size-6 rounded-lg font-bold text-[10px] transition-transform ${
                              isCellSelected
                                ? 'scale-110 shadow-md ring-1 ring-blue-400 bg-blue-500 text-white'
                                : 'hover:scale-110 active:scale-95'
                            } ${
                              isCellSelected
                                ? ''
                                : sKey === 'OFF'
                                  ? 'text-zinc-300 dark:text-zinc-600 hover:bg-zinc-950/5 dark:hover:bg-white/5'
                                  : `${colors.bg} ${colors.text} shadow-xs`
                            }`}
                          >
                            {sKey === 'OFF' && !isCellSelected ? '·' : sKey}
                          </div>

                          {/* Quick Shift Edit Popover (only when not multi-selecting) */}
                          {isEditing && selectedCells.size === 0 && (
                            <div
                              ref={popoverRef}
                              onClick={(e) => e.stopPropagation()}
                              className={`shift-popover absolute z-50 top-full ${
                                date.getDate() <= 2 ? 'left-0' : date.getDate() >= 29 ? 'right-0' : 'left-1/2 -translate-x-1/2'
                              } mt-1 rounded-xl bg-white dark:bg-[#1c1c1e] shadow-2xl border border-zinc-950/15 dark:border-white/15 p-1.5 flex items-center gap-1 animate-in fade-in zoom-in-95`}
                            >
                              {(['M', 'A', 'N', 'MID', 'OFF', 'H8'] as const).map((code) => (
                                <button
                                  key={code}
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleSetCellShift(code);
                                  }}
                                  className={`px-2 py-1 rounded-lg text-[10px] font-bold cursor-pointer transition-colors ${
                                    sKey === code ? 'bg-blue-500 text-white shadow-xs' : 'hover:bg-zinc-100 dark:hover:bg-white/10 text-zinc-700 dark:text-zinc-300'
                                  }`}
                                >
                                  {code}
                                </button>
                              ))}
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setEditingCell(null);
                                }}
                                className="p-1 text-zinc-400 hover:text-zinc-600 dark:hover:text-zinc-200 cursor-pointer rounded-lg hover:bg-zinc-100 dark:hover:bg-white/10"
                                title="Close"
                              >
                                <X className="size-3.5" />
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
                          {stats.H8 > 0 && <span className="text-pink-500 font-bold" title="H8 Holiday shifts">{stats.H8}H8</span>}
                          <span className="text-zinc-400" title="Off days">{stats.OFF}O</span>
                          <span
                            className={`font-bold px-1.5 py-0.5 rounded text-[10px] ${
                              stats.totalWorkingDays === stats.targetWorkingDays
                                ? 'bg-emerald-500/15 text-emerald-600 dark:text-emerald-400'
                                : 'bg-rose-500/15 text-rose-600 dark:text-rose-400'
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
                  const nightIssue = cov ? cov.N !== 2 : false;
                  return (
                    <td
                      key={isoDate}
                      className={`p-1 text-center font-mono text-[11px] border-r border-zinc-950/5 dark:border-white/5 ${
                        nightIssue ? 'bg-rose-500/20 text-rose-600 font-bold' : 'text-purple-700 dark:text-purple-400'
                      }`}
                    >
                      {cov?.N ?? 0}
                    </td>
                  );
                })}
                <td className="sticky right-0 z-10 p-1 text-center bg-zinc-100 dark:bg-[#141414] border-l border-zinc-950/10 dark:border-white/10 font-bold text-purple-700 dark:text-purple-400">
                  N = 2
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

      {/* Floating Block Shift Edit Toolbar */}
      {selectedCells.size > 0 && (
        <div className="fixed bottom-6 left-1/2 -translate-x-1/2 z-50 animate-in fade-in slide-in-from-bottom-4 duration-200">
          <div className={`p-2 px-4 ${GLASS_CARD} flex flex-wrap items-center gap-3 shadow-2xl border border-blue-500/40 ring-2 ring-blue-500/20 backdrop-blur-2xl bg-white/95 dark:bg-[#181818]/95`}>
            <div className="flex items-center gap-2">
              <span className="flex size-6 items-center justify-center rounded-full bg-blue-500 text-white text-[11px] font-bold shadow-xs">
                {selectedCells.size}
              </span>
              <span className="text-[12px] font-bold text-zinc-900 dark:text-zinc-100 whitespace-nowrap">
                Shifts Selected
              </span>
            </div>

            <div className="h-5 w-px bg-zinc-950/10 dark:bg-white/10" />

            <span className="text-[11px] font-semibold text-zinc-500 dark:text-zinc-400 whitespace-nowrap">
              Set Block:
            </span>

            <div className="flex items-center gap-1.5">
              {(['M', 'A', 'N', 'MID', 'OFF', 'H8'] as const).map((code) => {
                const colors = shiftColors[code] ?? { bg: 'bg-zinc-500/10', text: 'text-zinc-500' };
                return (
                  <button
                    key={code}
                    type="button"
                    onClick={() => handleApplyShiftToSelected(code)}
                    className={`px-3 py-1.5 rounded-xl text-[11px] font-bold cursor-pointer transition-all active:scale-95 shadow-xs border border-zinc-950/5 dark:border-white/5 hover:scale-105 ${colors.bg} ${colors.text} hover:ring-2 hover:ring-blue-500/50`}
                    title={`Change all ${selectedCells.size} selected shifts to ${code} (or press ${code === 'OFF' ? 'O' : code === 'H8' ? 'H' : code})`}
                  >
                    {code}
                  </button>
                );
              })}
            </div>

            <div className="h-5 w-px bg-zinc-950/10 dark:bg-white/10" />

            <button
              type="button"
              onClick={handleClearSelectedCells}
              className="flex items-center gap-1 text-[11px] font-semibold text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200 px-2 py-1 rounded-lg hover:bg-zinc-950/5 dark:hover:bg-white/5 cursor-pointer whitespace-nowrap"
              title="Deselect all (Esc)"
            >
              <X className="size-3.5" />
              <span>Deselect (Esc)</span>
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
