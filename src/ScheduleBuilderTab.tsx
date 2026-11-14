import { useState, useMemo, useEffect, useCallback, useRef } from 'react';
import {
  RotateCcw,
  RefreshCw,
  Save,
  FileSpreadsheet,
  Download,
  ChevronLeft,
  ChevronRight,
  ChevronDown,
  ChevronUp,
  AlertTriangle,
  CheckCircle2,
  Shield,
  ShieldAlert,
  Sparkles,
  Info,
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
  calculateAbsReadjustments,
  exportRosterToExcelFile,
  countLegalWorkingDaysInMonth,
  isHistoricalMonthInRoster,
  loadHistoricalSchedule,
  ROMANIAN_HOLIDAYS,
  FIXED_MID_EMPLOYEES,
  FIXED_WEEKDAY_MA_EMPLOYEES,
  type ComplianceTier,
  type RelaxationRecord,
  type ScheduleBuilderOptions,
  type GeneratedScheduleResult,
  type ScheduleValidationReport,
  type ReadjustmentProposal
} from './scheduleBuilder';
import { shiftKey, shiftColors, shiftSolid, initials, GLASS_CARD } from './scheduleUtils';
import { mergeRosters } from './parser';

export const RELAXED_CELL_RING_CLASS =
  'border-2 border-dashed border-amber-500 dark:border-amber-400 ring-2 ring-amber-400/40';

export type ScheduleComplianceStatusKind = 'compliant' | 'relaxed' | 'review';

export interface ResolvedScheduleComplianceState {
  tier: ComplianceTier;
  relaxations: RelaxationRecord[];
  needsManualReview: boolean;
  requiresReview: boolean;
  warnings: string[];
  status: ScheduleComplianceStatusKind;
  statusTitle: string;
  hasNightOver2Relaxation: boolean;
  isNightTargetRelaxed: boolean;
  relaxedNightTargetLabel: string;
}

function getMaxTierFromRelaxations(relaxations: RelaxationRecord[]): ComplianceTier {
  return relaxations.reduce<ComplianceTier>((maxT, r) => (r.tier > maxT ? r.tier : maxT), 0);
}

/**
 * Resolves unified compliance state (Tier 0..3, relaxations, and manual-review flag),
 * staying backward-compatible when tier/relaxations/needsManualReview are undefined.
 */
export function resolveScheduleComplianceState(
  scheduleData: GeneratedScheduleResult | null,
  validation: ScheduleValidationReport
): ResolvedScheduleComplianceState {
  const warnings = validation.warnings ?? [];
  const relaxations: RelaxationRecord[] = scheduleData?.relaxations ?? [];
  const inferredTier = getMaxTierFromRelaxations(relaxations);
  const tier: ComplianceTier = scheduleData?.tier ?? inferredTier;
  const needsManualReview =
    scheduleData?.needsManualReview !== undefined
      ? scheduleData.needsManualReview
      : !validation.isFullyCompliant || warnings.length > 0;
  const requiresReview = needsManualReview || warnings.length > 0;
  const status: ScheduleComplianceStatusKind = requiresReview
    ? 'review'
    : tier > 0 || relaxations.length > 0
      ? 'relaxed'
      : 'compliant';

  const statusTitle =
    status === 'compliant'
      ? 'Fully compliant'
      : status === 'relaxed'
        ? `Compliant with ${relaxations.length} approved relaxation${relaxations.length === 1 ? '' : 's'} (Tier ${tier})`
        : 'Needs manager review';

  const hasNightOver2Relaxation = relaxations.some((r) => r.rule === 'NIGHT_OVER_2');
  const isNightTargetRelaxed = tier >= 1 || hasNightOver2Relaxation;
  const relaxedNightTargetLabel = isNightTargetRelaxed ? 'N = 2 (3 relaxed)' : 'N = 2';

  return {
    tier,
    relaxations,
    needsManualReview,
    requiresReview,
    warnings,
    status,
    statusTitle,
    hasNightOver2Relaxation,
    isNightTargetRelaxed,
    relaxedNightTargetLabel
  };
}

function noteMentionsMorning(note: string): boolean {
  return /\bmorning\b|\brotm\b/i.test(note) || /\bM\b(?=\s*[=:<(/0-9]|$)/.test(note);
}

function noteMentionsAfternoon(note: string): boolean {
  return /\bafternoon\b|\btotala\b/i.test(note) || /\bA\b(?=\s*[=:<(/0-9]|$)/.test(note);
}

export interface RelaxationLookup {
  getCellRelaxations: (emp: string, isoDate: string, shiftCode: string) => RelaxationRecord[];
  getDayHeaderRelaxations: (isoDate: string) => RelaxationRecord[];
  getMorningFooterRelaxations: (isoDate: string) => RelaxationRecord[];
  getAfternoonFooterRelaxations: (isoDate: string) => RelaxationRecord[];
  getNightFooterRelaxations: (isoDate: string) => RelaxationRecord[];
}

/**
 * Indexes RelaxationRecord entries so every cell, day header, and footer coverage cell
 * can trace back to exact or day-wide relaxations.
 */
export function buildRelaxationLookup(
  relaxations: RelaxationRecord[],
  dailyCoverage: Record<string, ScheduleValidationReport['dailyCoverage'][string]>
): RelaxationLookup {
  const byCell = new Map<string, RelaxationRecord[]>();
  const byDateAll = new Map<string, RelaxationRecord[]>();
  const dayWideByDate = new Map<string, RelaxationRecord[]>();

  for (const r of relaxations) {
    const dateList = byDateAll.get(r.isoDate);
    if (dateList) dateList.push(r);
    else byDateAll.set(r.isoDate, [r]);

    if (r.employee) {
      const key = `${r.employee}::${r.isoDate}`;
      const cellList = byCell.get(key);
      if (cellList) cellList.push(r);
      else byCell.set(key, [r]);
    } else {
      const dayWideList = dayWideByDate.get(r.isoDate);
      if (dayWideList) dayWideList.push(r);
      else dayWideByDate.set(r.isoDate, [r]);
    }
  }

  return {
    getCellRelaxations(emp: string, isoDate: string, shiftCode: string): RelaxationRecord[] {
      const exact = byCell.get(`${emp}::${isoDate}`) ?? [];
      const dayWide = dayWideByDate.get(isoDate) ?? [];
      if (shiftCode !== 'N') return exact;
      const nightWide = dayWide.filter((r) => r.rule === 'NIGHT_OVER_2');
      return nightWide.length > 0 ? [...exact, ...nightWide] : exact;
    },

    getDayHeaderRelaxations(isoDate: string): RelaxationRecord[] {
      return dayWideByDate.get(isoDate) ?? [];
    },

    getMorningFooterRelaxations(isoDate: string): RelaxationRecord[] {
      const records = (byDateAll.get(isoDate) ?? []).filter((r) => r.rule === 'COVERAGE_GAP');
      if (records.length === 0) return [];
      const cov = dailyCoverage[isoDate];
      const isMorningUnder = cov ? (cov.rotM !== undefined ? cov.rotM < 1 : cov.M < 1) : true;
      if (!isMorningUnder) return [];
      return records.filter((r) => noteMentionsMorning(r.note) || !noteMentionsAfternoon(r.note));
    },

    getAfternoonFooterRelaxations(isoDate: string): RelaxationRecord[] {
      const records = (byDateAll.get(isoDate) ?? []).filter((r) => r.rule === 'COVERAGE_GAP');
      if (records.length === 0) return [];
      const cov = dailyCoverage[isoDate];
      if (cov && cov.A === 0) return [];
      const isAfternoonUnder = cov ? cov.A < 2 : true;
      if (!isAfternoonUnder) return [];
      return records.filter((r) => noteMentionsAfternoon(r.note) || !noteMentionsMorning(r.note));
    },

    getNightFooterRelaxations(isoDate: string): RelaxationRecord[] {
      return (byDateAll.get(isoDate) ?? []).filter((r) => r.rule === 'NIGHT_OVER_2');
    }
  };
}

/**
 * Determines whether a stricter-tier retry candidate should replace the active schedule.
 */
export function shouldAcceptStricterVariant(
  current: GeneratedScheduleResult,
  candidate: GeneratedScheduleResult
): boolean {
  const currentTier: ComplianceTier =
    current.tier ?? getMaxTierFromRelaxations(current.relaxations ?? []);
  const candidateTier: ComplianceTier =
    candidate.tier ?? getMaxTierFromRelaxations(candidate.relaxations ?? []);
  const candidateNeedsReview =
    candidate.needsManualReview ??
    (!candidate.validation.isFullyCompliant || candidate.validation.warnings.length > 0);
  return (
    candidateTier < currentTier &&
    !candidateNeedsReview &&
    candidate.validation.warnings.length === 0
  );
}

function getWarningActionNote(warning: string): string | null {
  const normMatch = warning.match(/^(.+?):\s*worked\s+(\d+)\s+days\s*\(target is\s*(\d+)\s*days/i);
  if (!normMatch) return null;
  const empName = normMatch[1].trim();
  const surname = empName.split(/\s+/)[0];
  const worked = parseInt(normMatch[2], 10);
  const target = parseInt(normMatch[3], 10);
  const diff = worked - target;
  if (diff > 0) {
    return `We need to give to ${surname} ${diff} day${diff > 1 ? 's' : ''} off after this change`;
  } else if (diff < 0) {
    return `We need to give to ${surname} ${Math.abs(diff)} more work day${Math.abs(diff) > 1 ? 's' : ''} to reach the monthly norm`;
  }
  return null;
}

function CoverageCountBadge({ count, isRelaxed }: { count: number; isRelaxed: boolean }) {
  if (!isRelaxed) return <>{count}</>;
  return (
    <span
      className={`inline-flex items-center justify-center min-w-[24px] h-6 px-1 rounded-lg bg-amber-500/15 text-amber-700 dark:text-amber-300 font-bold ${RELAXED_CELL_RING_CLASS}`}
    >
      {count}
    </span>
  );
}

interface ScheduleBuilderTabProps {
  roster: RosterData | null;
  onApplySchedule: (updatedRoster: RosterData, targetMonth: number, targetYear: number) => void;
  initialSchedule?: GeneratedScheduleResult | null;
}

export default function ScheduleBuilderTab({
  roster,
  onApplySchedule,
  initialSchedule = null
}: ScheduleBuilderTabProps) {
  // Target month & year state (defaults to initialSchedule, current roster tail, or current month)
  const [targetYear, setTargetYear] = useState<number>(() => {
    if (initialSchedule) return initialSchedule.year;
    if (roster && roster.dateColumns.length > 0) {
      const last = roster.dateColumns[roster.dateColumns.length - 1];
      const d = new Date(`${last.isoDate}T00:00:00`);
      return d.getFullYear();
    }
    return new Date().getFullYear();
  });

  const [targetMonth, setTargetMonth] = useState<number>(() => {
    if (initialSchedule) return initialSchedule.month;
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

  // Current working generated schedule
  const [scheduleData, setScheduleData] = useState<GeneratedScheduleResult | null>(initialSchedule);

  // Whether the selected month is an archived/historical record in the roster (set in stone)
  const isHistorical = useMemo(() => {
    return Boolean(scheduleData?.isHistorical) || isHistoricalMonthInRoster(roster, targetYear, targetMonth);
  }, [scheduleData?.isHistorical, roster, targetYear, targetMonth]);

  const [currentSeed, setCurrentSeed] = useState<number>(() =>
    initialSchedule?.seed ?? Math.floor(Math.random() * 999999)
  );
  const [appliedNotification, setAppliedNotification] = useState<string | null>(null);
  const [stricterStatusNote, setStricterStatusNote] = useState<{
    kind: 'success' | 'info';
    message: string;
  } | null>(null);
  const [isRelaxationsOpen, setIsRelaxationsOpen] = useState<boolean>(true);
  const [showAllWarnings, setShowAllWarnings] = useState<boolean>(false);
  const [isTryingStricter, setIsTryingStricter] = useState<boolean>(false);
  const [excludedEmployees, setExcludedEmployees] = useState<Set<string>>(
    () => new Set(initialSchedule?.excludedEmployees ?? [])
  );

  // Whether any employee in the schedule has ABS assigned
  const hasAbsShifts = useMemo(() => {
    if (!scheduleData?.rows) return false;
    return Object.values(scheduleData.rows).some((row) =>
      Object.values(row).some((s) => s === 'ABS' || shiftKey(s) === 'ABS')
    );
  }, [scheduleData]);

  // Readjustment modal & proposal state
  const [readjustProposal, setReadjustProposal] = useState<ReadjustmentProposal | null>(null);
  const [isReadjustModalOpen, setIsReadjustModalOpen] = useState(false);
  const [readjustSeed, setReadjustSeed] = useState(0);

  const handleOpenReadjustModal = (seed = 0) => {
    if (!scheduleData) return;
    const proposal = calculateAbsReadjustments(
      scheduleData.rows,
      scheduleData.dateColumns,
      roster,
      targetYear,
      targetMonth,
      targetWorkingDays,
      seed
    );
    setReadjustProposal(proposal);
    setReadjustSeed(seed);
    setIsReadjustModalOpen(true);
  };

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
      targetWorkingDays,
      Array.from(excludedEmployees)
    );
  }, [scheduleData, roster, targetYear, targetMonth, targetWorkingDays, excludedEmployees]);

  const complianceState = useMemo(
    () => resolveScheduleComplianceState(scheduleData, validation),
    [scheduleData, validation]
  );

  const relaxationLookup = useMemo(
    () => buildRelaxationLookup(complianceState.relaxations, validation.dailyCoverage),
    [complianceState.relaxations, validation.dailyCoverage]
  );

  const revalidateEditedSchedule = useCallback(
    (
      baseSchedule: GeneratedScheduleResult,
      updatedRows: Record<string, Record<string, string>>,
      precomputedValidation?: ScheduleValidationReport
    ): GeneratedScheduleResult => {
      const currentTier: ComplianceTier =
        baseSchedule.tier ?? getMaxTierFromRelaxations(baseSchedule.relaxations ?? []);
      if (precomputedValidation && currentTier === 0) {
        return {
          ...baseSchedule,
          rows: updatedRows,
          validation: precomputedValidation,
          tier: 0,
          relaxations: [],
          needsManualReview:
            !precomputedValidation.isFullyCompliant || precomputedValidation.warnings.length > 0
        };
      }
      const nextRelaxations: RelaxationRecord[] = [];
      const testReport = validateScheduleCompliance(
        updatedRows,
        baseSchedule.dateColumns,
        roster,
        targetYear,
        targetMonth,
        targetWorkingDays,
        Array.from(excludedEmployees),
        { tier: currentTier, relaxations: nextRelaxations }
      );
      const nextTier = getMaxTierFromRelaxations(nextRelaxations);
      const nextNeedsReview = !testReport.isFullyCompliant || testReport.warnings.length > 0;
      return {
        ...baseSchedule,
        rows: updatedRows,
        validation: testReport,
        tier: nextNeedsReview ? currentTier : nextTier,
        relaxations: nextRelaxations,
        needsManualReview: nextNeedsReview
      };
    },
    [roster, targetYear, targetMonth, targetWorkingDays, excludedEmployees]
  );

  const handleApplyReadjustments = () => {
    if (!readjustProposal || !scheduleData) return;
    setScheduleData(
      revalidateEditedSchedule(scheduleData, readjustProposal.updatedRows, readjustProposal.validation)
    );
    setStricterStatusNote(null);
    setIsReadjustModalOpen(false);
    setAppliedNotification(`Applied ${readjustProposal.adjustments.length} shift readjustments to cover ABS!`);
    setTimeout(() => setAppliedNotification(null), 4000);
  };

  // Multi-cell selection state for block changes
  const [selectedCells, setSelectedCells] = useState<Set<string>>(new Set());
  const [bulkFeedback, setBulkFeedback] = useState<string | null>(null);
  const isMouseDownRef = useRef(false);
  const dragStartCellRef = useRef<{ employee: string; isoDate: string } | null>(null);
  const hasDraggedRef = useRef(false);

  // Clear dragging on mouseup
  useEffect(() => {
    const handleMouseUp = () => {
      isMouseDownRef.current = false;
      dragStartCellRef.current = null;
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
    if (!initialSchedule && !hasUserNavigated && roster && roster.dateColumns.length > 0) {
      const last = roster.dateColumns[roster.dateColumns.length - 1];
      const d = new Date(`${last.isoDate}T00:00:00`);
      setTargetYear(d.getFullYear());
      setTargetMonth(d.getMonth());
    }
  }, [initialSchedule, roster, hasUserNavigated]);

  const [isGenerating, setIsGenerating] = useState(false);
  const skipInitialGenerationRef = useRef(Boolean(initialSchedule));

  // Generate initial schedule or load historical schedule if already set in stone
  const runGeneration = useCallback(
    (
      seedValue: number = currentSeed,
      y: number = targetYear,
      m: number = targetMonth,
      exclusions: Set<string> = excludedEmployees
    ) => {
      setIsGenerating(true);
      setStricterStatusNote(null);
      setShowAllWarnings(false);
      setTimeout(() => {
        if (isHistoricalMonthInRoster(roster, y, m)) {
          setScheduleData(loadHistoricalSchedule(roster!, y, m));
        } else {
          const result = generateBalancedSchedule(roster, y, m, {
            seed: seedValue,
            targetWorkingDays: countLegalWorkingDaysInMonth(m, y),
            excludedEmployees: Array.from(exclusions)
          });
          setScheduleData(result);
        }
        setSelectedCells(new Set());
        setBulkFeedback(null);
        setIsGenerating(false);
      }, 20);
    },
    [roster, currentSeed, targetYear, targetMonth, excludedEmployees]
  );

  useEffect(() => {
    if (skipInitialGenerationRef.current) {
      skipInitialGenerationRef.current = false;
      return;
    }
    runGeneration(currentSeed, targetYear, targetMonth, excludedEmployees);
  }, [targetYear, targetMonth, targetWorkingDays, roster]);

  // Employee list
  const employees = useMemo(() => {
    if (scheduleData) return Object.keys(scheduleData.rows);
    return roster?.employees ?? [];
  }, [scheduleData, roster]);

  // Handlers for month navigation
  const handlePrevMonth = () => {
    if (isGenerating) return;
    setHasUserNavigated(true);
    setExcludedEmployees(new Set());
    setStricterStatusNote(null);
    const prev = new Date(targetYear, targetMonth - 1, 1);
    setTargetMonth(prev.getMonth());
    setTargetYear(prev.getFullYear());
  };

  const handleNextMonth = () => {
    if (isGenerating) return;
    setHasUserNavigated(true);
    setExcludedEmployees(new Set());
    setStricterStatusNote(null);
    const next = new Date(targetYear, targetMonth + 1, 1);
    setTargetMonth(next.getMonth());
    setTargetYear(next.getFullYear());
  };

  const handleToggleEmployeeActive = (emp: string) => {
    if (isHistorical || isGenerating) return;
    const next = new Set(excludedEmployees);
    if (next.has(emp)) {
      next.delete(emp);
    } else {
      next.add(emp);
    }
    setExcludedEmployees(next);
    runGeneration(currentSeed, targetYear, targetMonth, next);
  };

  const handleToggleAllEmployees = () => {
    if (isHistorical || isGenerating) return;
    if (excludedEmployees.size > 0) {
      const next = new Set<string>();
      setExcludedEmployees(next);
      runGeneration(currentSeed, targetYear, targetMonth, next);
    }
  };

  const handleRegenerateVariant = () => {
    if (isHistorical || isGenerating) return;
    const nextSeed = Math.floor(Math.random() * 999999);
    setCurrentSeed(nextSeed);
    runGeneration(nextSeed, targetYear, targetMonth, excludedEmployees);
  };

  const handleTryStricterVariant = () => {
    if (isHistorical || isGenerating || !scheduleData) return;
    const activeTier = complianceState.tier;
    if (activeTier <= 0) return;

    const maxAllowedTier = (activeTier - 1) as ComplianceTier;
    const nextSeed = Math.floor(Math.random() * 999999);
    setIsGenerating(true);
    setIsTryingStricter(true);
    setStricterStatusNote(null);

    setTimeout(() => {
      const options: ScheduleBuilderOptions = {
        seed: nextSeed,
        targetWorkingDays: countLegalWorkingDaysInMonth(targetMonth, targetYear),
        excludedEmployees: Array.from(excludedEmployees),
        minTier: 0,
        maxTier: maxAllowedTier
      };
      const res = generateBalancedSchedule(roster, targetYear, targetMonth, options);

      if (shouldAcceptStricterVariant(scheduleData, res)) {
        const appliedTier =
          res.tier ?? getMaxTierFromRelaxations(res.relaxations ?? []);
        setCurrentSeed(nextSeed);
        setScheduleData(res);
        setSelectedCells(new Set());
        setBulkFeedback(null);
        setShowAllWarnings(false);
        setStricterStatusNote({
          kind: 'success',
          message: `Applied stricter Tier ${appliedTier} schedule (seed #${nextSeed}).`
        });
      } else {
        const tierLabel = maxAllowedTier === 0 ? 'Tier 0' : `Tier 0–${maxAllowedTier}`;
        setStricterStatusNote({
          kind: 'info',
          message: `No stricter compliant schedule (${tierLabel}) found within budget (seed #${nextSeed}). Keeping current Tier ${activeTier} schedule.`
        });
      }

      setIsTryingStricter(false);
      setIsGenerating(false);
    }, 20);
  };

  // Finds the full contiguous shift block containing targetIsoDate for emp.
  // If emp is on an extended run of OFF (> 4 days, e.g. deselected for whole month), returns a standard 4-day block.
  const getBlockCellsForEmployee = (
    data: GeneratedScheduleResult,
    emp: string,
    targetIsoDate: string
  ): string[] => {
    const dateCols = data.dateColumns;
    const targetIdx = dateCols.findIndex((c) => c.isoDate === targetIsoDate);
    if (targetIdx === -1) return [`${emp}::${targetIsoDate}`];

    const row = data.rows[emp] || {};
    const currentShift = shiftKey(row[targetIsoDate] ?? 'OFF');

    // Find contiguous range of same shift for this employee
    let startIdx = targetIdx;
    while (startIdx > 0 && shiftKey(row[dateCols[startIdx - 1].isoDate] ?? 'OFF') === currentShift) {
      startIdx--;
    }

    let endIdx = targetIdx;
    while (endIdx < dateCols.length - 1 && shiftKey(row[dateCols[endIdx + 1].isoDate] ?? 'OFF') === currentShift) {
      endIdx++;
    }

    const contiguousCount = endIdx - startIdx + 1;

    // If extended OFF run (> 4 days, e.g. deselected employee off for month),
    // select a standard 4-day block starting from targetIdx (or clamped to end of month/run)
    if (currentShift === 'OFF' && contiguousCount > 4) {
      const blockStart = targetIdx;
      const blockEnd = Math.min(dateCols.length - 1, targetIdx + 3);
      const result: string[] = [];
      for (let i = blockStart; i <= blockEnd; i++) {
        result.push(`${emp}::${dateCols[i].isoDate}`);
      }
      return result;
    }

    // Standard block: all contiguous days of this shift
    const result: string[] = [];
    for (let i = startIdx; i <= endIdx; i++) {
      result.push(`${emp}::${dateCols[i].isoDate}`);
    }
    return result;
  };

  // Multi-cell selection & block change handlers
  const handleCellMouseDown = (e: React.MouseEvent, emp: string, isoDate: string) => {
    if ((e.target as HTMLElement).closest('.shift-popover')) return;

    const cellKey = `${emp}::${isoDate}`;
    hasDraggedRef.current = false;

    // 1. Shift-click: Select entire shift block!
    if (e.shiftKey && scheduleData) {
      e.preventDefault();
      setEditingCell(null);
      const blockKeys = getBlockCellsForEmployee(scheduleData, emp, isoDate);

      setSelectedCells((prev) => {
        // If holding Ctrl/Cmd as well, add to existing selection; otherwise replace
        if (e.ctrlKey || e.metaKey) {
          const next = new Set(prev);
          blockKeys.forEach((k) => next.add(k));
          return next;
        }
        return new Set(blockKeys);
      });
      isMouseDownRef.current = false;
      dragStartCellRef.current = null;
      return;
    }

    // 2. Ctrl/Cmd click: Toggle single cell into/out of multi-selection
    if (e.ctrlKey || e.metaKey) {
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
      isMouseDownRef.current = true;
      dragStartCellRef.current = { employee: emp, isoDate };
      return;
    }

    // 3. Normal mouse down: prepare for click or drag
    isMouseDownRef.current = true;
    dragStartCellRef.current = { employee: emp, isoDate };
  };

  const handleCellMouseEnter = (e: React.MouseEvent, emp: string, isoDate: string) => {
    if (!isMouseDownRef.current || !dragStartCellRef.current || !scheduleData) return;

    // Holding Ctrl/Cmd while dragging: toggle/add hovered cells
    if (e.ctrlKey || e.metaKey) {
      hasDraggedRef.current = true;
      const cellKey = `${emp}::${isoDate}`;
      setSelectedCells((prev) => {
        if (prev.has(cellKey)) return prev;
        const next = new Set(prev);
        next.add(cellKey);
        return next;
      });
      return;
    }

    // Normal drag across cells: select range on the current employee or rectangular area
    hasDraggedRef.current = true;
    setEditingCell(null);

    const allDates = scheduleData.dateColumns.map((c) => c.isoDate);
    const allEmps = Object.keys(scheduleData.rows);

    const startEmpIdx = allEmps.indexOf(dragStartCellRef.current.employee);
    const endEmpIdx = allEmps.indexOf(emp);
    const startDateIdx = allDates.indexOf(dragStartCellRef.current.isoDate);
    const endDateIdx = allDates.indexOf(isoDate);

    if (startEmpIdx !== -1 && endEmpIdx !== -1 && startDateIdx !== -1 && endDateIdx !== -1) {
      const [lowEmp, highEmp] = startEmpIdx <= endEmpIdx ? [startEmpIdx, endEmpIdx] : [endEmpIdx, startEmpIdx];
      const [lowDate, highDate] = startDateIdx <= endDateIdx ? [startDateIdx, endDateIdx] : [endDateIdx, startDateIdx];

      const newSet = new Set<string>();
      for (let eIdx = lowEmp; eIdx <= highEmp; eIdx++) {
        for (let dIdx = lowDate; dIdx <= highDate; dIdx++) {
          newSet.add(`${allEmps[eIdx]}::${allDates[dIdx]}`);
        }
      }
      setSelectedCells(newSet);
    }
  };

  const handleCellClick = (e: React.MouseEvent, emp: string, isoDate: string) => {
    if ((e.target as HTMLElement).closest('.shift-popover')) return;
    if (e.shiftKey || e.ctrlKey || e.metaKey) return;
    if (hasDraggedRef.current) return;

    const cellKey = `${emp}::${isoDate}`;
    if (selectedCells.size > 0 && selectedCells.has(cellKey)) {
      return;
    }

    setSelectedCells(new Set());
    setEditingCell({ employee: emp, isoDate });
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

    setScheduleData(revalidateEditedSchedule(scheduleData, updatedRows));
    setStricterStatusNote(null);
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
      } else if (key === 'B') {
        e.preventDefault();
        handleApplyShiftToSelected('ABS');
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [selectedCells, scheduleData, roster, targetYear, targetMonth, targetWorkingDays, excludedEmployees]);

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

    setScheduleData(revalidateEditedSchedule(scheduleData, updatedRows));
    setStricterStatusNote(null);
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
              <span className="font-semibold text-blue-500">Palade Alexandru-Ionut</span> (<span className="font-semibold">MID</span>) &amp; <span className="font-semibold text-emerald-500">Tir George Cristian</span> (<span className="font-semibold">M/A</span> by ISO week) work Mon–Fri, <span className="font-semibold">OFF</span> Sat–Sun &amp; legal holidays.
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
              <span className="rounded bg-green-500/15 text-green-700 dark:text-green-300 px-1.5 py-0.5">Morning: ≥ 1 (rot)</span>
              <span className="rounded bg-orange-500/15 text-orange-700 dark:text-orange-300 px-1.5 py-0.5">Afternoon: ≥ 2</span>
              <span className="rounded bg-purple-500/15 text-purple-700 dark:text-purple-300 px-1.5 py-0.5">Night: 2</span>
            </div>
          </div>
        </div>

        {/* Card 3: Legal Compliance Status (3-state) or Historical Record */}
        {isHistorical ? (
          <div className={`p-4 ${GLASS_CARD} flex items-start gap-3`}>
            <div className="rounded-xl p-2 shrink-0 bg-amber-500/10 text-amber-500">
              <Lock className="size-5" />
            </div>
            <div className="text-[12px] flex-1 min-w-0">
              <div className="flex items-center justify-between gap-2">
                <p className="font-semibold text-zinc-800 dark:text-zinc-200 text-[13px]">
                  Archived Historical Record
                </p>
                <span className="font-semibold text-[10px] text-amber-600 dark:text-amber-400 bg-amber-500/10 px-2 py-0.5 rounded-full border border-amber-500/20 shrink-0">
                  Set in Stone
                </span>
              </div>
              <p className="text-zinc-500 dark:text-zinc-400 mt-0.5 leading-relaxed">
                Official historical schedule loaded directly from roster data. Locked and non-regenerated.
              </p>
            </div>
          </div>
        ) : complianceState.status === 'compliant' ? (
          <div className={`p-4 ${GLASS_CARD} border border-emerald-500/30 bg-emerald-500/[0.06] dark:bg-emerald-500/[0.08] flex items-start gap-3`}>
            <div className="rounded-xl p-2 shrink-0 bg-emerald-500/15 text-emerald-600 dark:text-emerald-400">
              <CheckCircle2 className="size-5" />
            </div>
            <div className="text-[12px] flex-1 min-w-0">
              <div className="flex items-center justify-between gap-2">
                <p className="font-semibold text-emerald-800 dark:text-emerald-200 text-[13px]">
                  {complianceState.statusTitle}
                </p>
                <span className="font-mono text-[10px] text-zinc-500 dark:text-zinc-400 shrink-0">
                  Tier 0 · Seed #{scheduleData?.seed ?? '—'}
                </span>
              </div>
              <p className="text-zinc-600 dark:text-zinc-300 mt-0.5 leading-relaxed">
                All employees meet the monthly norm of {targetWorkingDays} working days with strict rests and 24/7 coverage.
              </p>
            </div>
          </div>
        ) : complianceState.status === 'relaxed' ? (
          <div className={`p-4 ${GLASS_CARD} border border-amber-500/35 bg-amber-500/[0.06] dark:bg-amber-500/[0.08] flex items-start gap-3`}>
            <div className="rounded-xl p-2 shrink-0 bg-amber-500/15 text-amber-600 dark:text-amber-400">
              <AlertTriangle className="size-5" />
            </div>
            <div className="text-[12px] flex-1 min-w-0">
              <div className="flex items-center justify-between gap-2">
                <p className="font-semibold text-amber-800 dark:text-amber-200 text-[13px]">
                  {complianceState.statusTitle}
                </p>
                <span className="font-mono text-[10px] text-zinc-500 dark:text-zinc-400 shrink-0">
                  Seed #{scheduleData?.seed ?? '—'}
                </span>
              </div>
              <p className="text-zinc-600 dark:text-zinc-300 mt-0.5 leading-relaxed">
                Hard rest/sequence rules remain satisfied; dashed amber cells mark approved Tier {complianceState.tier} relaxations.
              </p>
            </div>
          </div>
        ) : (
          <div className={`p-4 ${GLASS_CARD} border border-rose-500/40 bg-rose-500/[0.07] dark:bg-rose-500/[0.10] flex items-start gap-3`}>
            <div className="rounded-xl p-2 shrink-0 bg-rose-500/15 text-rose-600 dark:text-rose-400">
              <ShieldAlert className="size-5" />
            </div>
            <div className="text-[12px] flex-1 min-w-0">
              <div className="flex items-center justify-between gap-2">
                <p className="font-semibold text-rose-800 dark:text-rose-200 text-[13px]">
                  {complianceState.statusTitle}
                </p>
                <span className="font-mono text-[10px] text-zinc-500 dark:text-zinc-400 shrink-0">
                  {complianceState.tier > 0 ? `Tier ${complianceState.tier} · ` : ''}Seed #{scheduleData?.seed ?? '—'}
                </span>
              </div>
              <p className="text-zinc-600 dark:text-zinc-300 mt-0.5 leading-relaxed">
                {complianceState.warnings.length > 0
                  ? `${complianceState.warnings.length} hard or coverage rule violation${complianceState.warnings.length === 1 ? '' : 's'} detected — inspect warnings below and adjust manually.`
                  : 'Hard or coverage rules could not be fully satisfied within budget — inspect cells and adjust manually.'}
              </p>
            </div>
          </div>
        )}
      </div>

      {/* Main Action Bar */}
      <div className={`p-3.5 ${GLASS_CARD} flex flex-wrap items-center justify-between gap-3`}>
        <div className="flex flex-wrap items-center gap-2">
          {/* Regenerate Variant Button Container */}
          <div className="flex flex-col gap-1.5">
            <div className="flex flex-wrap items-center gap-2">
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
                ) : isGenerating && !isTryingStricter ? (
                  <Loader2 className="size-4 animate-spin" />
                ) : (
                  <RotateCcw className="size-4" />
                )}
                <span>
                  {isHistorical
                    ? 'Locked: Set in Stone'
                    : isGenerating && !isTryingStricter
                      ? 'Generating...'
                      : 'Regenerate Variant'}
                </span>
              </button>

              {!isHistorical && scheduleData && complianceState.tier > 0 && (
                <button
                  type="button"
                  disabled={isGenerating}
                  onClick={handleTryStricterVariant}
                  className={`flex items-center gap-1.5 rounded-xl font-semibold px-3.5 py-2 text-[13px] shadow transition-all ${
                    isGenerating
                      ? 'bg-zinc-200 dark:bg-zinc-800 text-zinc-400 dark:text-zinc-500 cursor-not-allowed opacity-60'
                      : 'bg-amber-500 hover:bg-amber-600 text-white active:scale-98 cursor-pointer'
                  }`}
                  title={`Attempt to generate a stricter schedule (Tier 0${complianceState.tier - 1 > 0 ? `–${complianceState.tier - 1}` : ''})`}
                >
                  {isTryingStricter ? (
                    <Loader2 className="size-4 animate-spin" />
                  ) : (
                    <Sparkles className="size-4" />
                  )}
                  <span>{isTryingStricter ? 'Trying Stricter...' : 'Try stricter variant'}</span>
                </button>
              )}
            </div>

            {/* Readjust Shifts Button (pops up beneath regenerate button when colleagues have ABS) */}
            {hasAbsShifts && !isHistorical && (
              <button
                type="button"
                onClick={() => handleOpenReadjustModal(0)}
                className="flex items-center justify-center gap-2 rounded-xl bg-gradient-to-r from-amber-500 to-orange-500 hover:from-amber-600 hover:to-orange-600 text-white font-semibold px-3 py-1.5 text-[12px] shadow-sm transition-all cursor-pointer active:scale-98 animate-in fade-in slide-in-from-top-1"
                title="Calculate minimal shift readjustments to makeup for colleagues on ABS"
              >
                <RefreshCw className="size-3.5" />
                <span>Readjust Shifts</span>
                <span className="bg-black/30 text-white text-[10px] px-1.5 py-0.2 rounded font-mono font-bold">
                  ABS
                </span>
              </button>
            )}
          </div>

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
              : 'Blocks of 3–5 days (2 in Tier 2) • Rest 2–3 days'}
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

      {/* Stricter Variant Inline Status Note */}
      {stricterStatusNote && (
        <div
          className={`rounded-2xl border px-3.5 py-2.5 text-[12px] flex items-center justify-between gap-2 ${
            stricterStatusNote.kind === 'success'
              ? 'border-emerald-500/30 bg-emerald-500/10 text-emerald-700 dark:text-emerald-300'
              : 'border-amber-500/30 bg-amber-500/10 text-amber-800 dark:text-amber-200'
          }`}
        >
          <div className="flex items-center gap-2">
            <Info className="size-4 shrink-0" />
            <span className="font-medium">{stricterStatusNote.message}</span>
          </div>
          <button
            type="button"
            onClick={() => setStricterStatusNote(null)}
            className="opacity-70 hover:opacity-100 cursor-pointer"
            title="Dismiss"
          >
            <X className="size-3.5" />
          </button>
        </div>
      )}

      {/* Relaxations Used Panel (never hides any relaxation) */}
      {complianceState.relaxations.length > 0 && (
        <div className={`p-4 ${GLASS_CARD} border border-amber-500/35 bg-amber-500/[0.06] dark:bg-amber-500/[0.08] space-y-3`}>
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2 text-amber-700 dark:text-amber-300 font-bold text-[13px]">
              <AlertTriangle className="size-4 shrink-0 text-amber-500" />
              <span>Relaxations used ({complianceState.relaxations.length})</span>
              <span className="rounded-full bg-amber-500/15 border border-amber-500/30 px-2 py-0.5 text-[10px] font-bold text-amber-700 dark:text-amber-300">
                Tier {complianceState.tier}
              </span>
            </div>
            <button
              type="button"
              onClick={() => setIsRelaxationsOpen((prev) => !prev)}
              className="flex items-center gap-1 rounded-lg px-2.5 py-1 text-[11px] font-semibold text-amber-700 dark:text-amber-300 hover:bg-amber-500/15 transition-colors cursor-pointer"
              aria-expanded={isRelaxationsOpen}
            >
              <span>
                {isRelaxationsOpen
                  ? 'Hide details'
                  : `Show all ${complianceState.relaxations.length}`}
              </span>
              {isRelaxationsOpen ? (
                <ChevronUp className="size-3.5" />
              ) : (
                <ChevronDown className="size-3.5" />
              )}
            </button>
          </div>

          {isRelaxationsOpen && (
            <ul className="space-y-1.5 max-h-72 overflow-y-auto pr-1 text-[12px]">
              {complianceState.relaxations.map((r, idx) => (
                <li
                  key={`${r.isoDate}-${r.rule}-${r.employee ?? 'coverage'}-${idx}`}
                  className="flex flex-wrap items-center justify-between gap-2 rounded-xl border border-amber-500/20 bg-white/75 dark:bg-zinc-900/70 px-3 py-2"
                >
                  <div className="flex flex-wrap items-center gap-2 min-w-0">
                    <span className="shrink-0 rounded-full bg-amber-500/20 border border-amber-500/35 px-2 py-0.5 text-[10px] font-bold text-amber-800 dark:text-amber-200">
                      Tier {r.tier}
                    </span>
                    <span className="font-mono font-bold text-[11px] text-zinc-900 dark:text-zinc-100 shrink-0">
                      {r.isoDate}
                    </span>
                    <span className="text-zinc-400">•</span>
                    <span className="font-semibold text-zinc-800 dark:text-zinc-200 shrink-0">
                      {r.employee ?? 'Coverage'}
                    </span>
                    <span className="rounded bg-amber-500/15 text-amber-800 dark:text-amber-300 px-1.5 py-0.5 font-mono text-[10px] font-bold shrink-0">
                      {r.rule}
                    </span>
                    <span className="text-zinc-600 dark:text-zinc-300 text-[11px] break-words">
                      {r.note}
                    </span>
                  </div>
                </li>
              ))}
            </ul>
          )}
        </div>
      )}

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
            Hotkeys: <span className="font-mono font-semibold">M</span>, <span className="font-mono font-semibold">A</span>, <span className="font-mono font-semibold">N</span>, <span className="font-mono font-semibold">O</span> (off), <span className="font-mono font-semibold">H</span> (h8), <span className="font-mono font-semibold">B</span> (abs)
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
                <th className="sticky left-0 z-20 min-w-[210px] p-2.5 font-bold text-zinc-600 dark:text-zinc-400 bg-zinc-100 dark:bg-[#121212] border-r border-zinc-950/10 dark:border-white/10">
                  <div className="flex items-center justify-between gap-2">
                    <div className="flex items-center gap-2">
                      <input
                        type="checkbox"
                        checked={excludedEmployees.size === 0}
                        onChange={handleToggleAllEmployees}
                        disabled={isHistorical || isGenerating}
                        className="rounded border-zinc-300 dark:border-zinc-700 text-blue-600 focus:ring-blue-500 cursor-pointer size-3.5"
                        title={excludedEmployees.size > 0 ? "Click to reactivate all employees" : "All employees active"}
                      />
                      <span>Employee</span>
                    </div>
                    <span className="text-[10px] font-medium px-1.5 py-0.5 rounded-full bg-zinc-200 dark:bg-zinc-800 text-zinc-600 dark:text-zinc-400">
                      {employees.length - excludedEmployees.size}/{employees.length} Active
                    </span>
                  </div>
                </th>

                {/* Day Columns */}
                {scheduleData?.dateColumns.map(({ isoDate, date }) => {
                  const dayNum = date.getDate();
                  const dayName = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'][date.getDay()];
                  const isWeekend = date.getDay() === 0 || date.getDay() === 6;
                  const isHoliday = ROMANIAN_HOLIDAYS.has(isoDate);
                  const dayHeaderRelaxations = relaxationLookup.getDayHeaderRelaxations(isoDate);
                  const hasDayRelaxation = dayHeaderRelaxations.length > 0;
                  const holidayNote = isHoliday ? ' (National Holiday)' : '';
                  const headerTooltip = hasDayRelaxation
                    ? `${isoDate}${holidayNote} — Relaxations: ${dayHeaderRelaxations.map((r) => `${r.rule} (Tier ${r.tier}): ${r.note}`).join('; ')}`
                    : `Click to select/deselect all shifts on ${isoDate}${holidayNote}`;

                  return (
                    <th
                      key={isoDate}
                      onClick={(e) => handleToggleColumnSelection(isoDate, e.ctrlKey || e.metaKey)}
                      className={`min-w-[42px] max-w-[42px] p-1 text-center font-semibold transition-colors border-r border-zinc-950/5 dark:border-white/5 cursor-pointer ${
                        hasDayRelaxation
                          ? 'bg-amber-500/10 text-amber-800 dark:text-amber-200 hover:bg-amber-500/15'
                          : isHoliday
                            ? 'bg-rose-500/[0.08] dark:bg-rose-500/[0.12] text-rose-600 dark:text-rose-400 hover:bg-rose-500/15'
                            : isWeekend
                              ? 'bg-zinc-950/[0.05] dark:bg-white/[0.05] text-zinc-500 hover:bg-zinc-950/10 dark:hover:bg-white/10'
                              : 'text-zinc-700 dark:text-zinc-300 hover:bg-zinc-950/[0.03] dark:hover:bg-white/[0.03]'
                      }`}
                      title={headerTooltip}
                    >
                      <div
                        className={`rounded-md py-0.5 ${
                          hasDayRelaxation ? RELAXED_CELL_RING_CLASS : ''
                        }`}
                      >
                        <div className="text-[9px] uppercase tracking-tighter opacity-70">{dayName}</div>
                        <div className="text-[12px] font-bold leading-tight">{dayNum}</div>
                      </div>
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
                const isExcluded = excludedEmployees.has(emp);
                const isFixedMid = FIXED_MID_EMPLOYEES.includes(emp);
                const isWeekdayMA = FIXED_WEEKDAY_MA_EMPLOYEES.includes(emp);
                const stats = validation.employeeStats[emp];
                const hasCustomShifts = isExcluded && (stats?.totalWorkingDays ?? 0) > 0;

                return (
                  <tr
                    key={emp}
                    className={`transition-colors ${
                      isExcluded && !hasCustomShifts
                        ? 'bg-zinc-950/[0.03] dark:bg-white/[0.015]'
                        : 'hover:bg-zinc-950/[0.02] dark:hover:bg-white/[0.02]'
                    }`}
                  >
                    {/* Sticky Employee Name Cell */}
                    <td
                      onClick={(e) => handleToggleRowSelection(emp, e.ctrlKey || e.metaKey)}
                      className="sticky left-0 z-10 p-2 font-medium bg-zinc-100 dark:bg-[#121212] border-r border-zinc-950/10 dark:border-white/10 cursor-pointer hover:bg-zinc-200 dark:hover:bg-[#1a1a1a] transition-colors"
                      title={`Click to select/deselect all shifts for ${emp}`}
                    >
                      <div className="flex items-center gap-2">
                        <input
                          type="checkbox"
                          checked={!isExcluded}
                          onChange={(e) => {
                            e.stopPropagation();
                            handleToggleEmployeeActive(emp);
                          }}
                          disabled={isHistorical || isGenerating}
                          className="rounded border-zinc-300 dark:border-zinc-700 text-blue-600 focus:ring-blue-500 cursor-pointer size-3.5 shrink-0"
                          title={
                            isExcluded
                              ? hasCustomShifts
                                ? `${emp} has custom shifts assigned. Check to restore full auto-rotation`
                                : `Include ${emp} in monthly roster generation`
                              : `Exclude ${emp} from monthly roster generation (Off whole month)`
                          }
                        />
                        <div className={`flex size-6 shrink-0 items-center justify-center rounded-lg text-[10px] font-bold ${
                          isExcluded && !hasCustomShifts
                            ? 'bg-zinc-500/10 text-zinc-400'
                            : isExcluded && hasCustomShifts
                            ? 'bg-amber-500/15 text-amber-600 dark:text-amber-400'
                            : isFixedMid
                            ? 'bg-blue-500/15 text-blue-500'
                            : isWeekdayMA
                            ? 'bg-emerald-500/15 text-emerald-500'
                            : 'bg-zinc-500/15 text-zinc-600 dark:text-zinc-400'
                        }`}>
                          {initials(emp)}
                        </div>
                        <div className="min-w-0">
                          <p className={`truncate font-semibold text-[12px] ${
                            isExcluded && !hasCustomShifts
                              ? 'line-through text-zinc-400 dark:text-zinc-500'
                              : 'text-zinc-900 dark:text-zinc-100'
                          }`}>
                            {emp}
                          </p>
                          {isExcluded ? (
                            hasCustomShifts ? (
                              <span className="text-[9px] font-bold text-amber-500 dark:text-amber-400 uppercase tracking-tight">
                                Off / Custom ({stats?.totalWorkingDays}d)
                              </span>
                            ) : (
                              <span className="text-[9px] font-bold text-amber-500 dark:text-amber-400 uppercase tracking-tight">Month Off</span>
                            )
                          ) : isFixedMid ? (
                            <span className="text-[9px] font-bold text-blue-500 uppercase">Fixed Mid</span>
                          ) : isWeekdayMA ? (
                            <span className="text-[9px] font-bold text-emerald-500 uppercase">Weekday M/A</span>
                          ) : null}
                        </div>
                      </div>
                    </td>

                    {/* Shift Cells */}
                    {scheduleData?.dateColumns.map(({ isoDate, date }) => {
                      const rawShift = scheduleData.rows[emp]?.[isoDate] ?? 'OFF';
                      const sKey = shiftKey(rawShift);
                      const isWeekend = date.getDay() === 0 || date.getDay() === 6;
                      const isHoliday = ROMANIAN_HOLIDAYS.has(isoDate);
                      const cellKey = `${emp}::${isoDate}`;
                      const isCellSelected = selectedCells.has(cellKey);
                      const isEditing = editingCell?.employee === emp && editingCell?.isoDate === isoDate;
                      const cellRelaxations = relaxationLookup.getCellRelaxations(emp, isoDate, sKey);
                      const isCellRelaxed = cellRelaxations.length > 0;
                      const relaxationTitleSuffix = isCellRelaxed
                        ? ` | Relaxations: ${cellRelaxations.map((r) => `Tier ${r.tier} ${r.rule}: ${r.note}`).join('; ')}`
                        : '';

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
                          onClick={(e) => handleCellClick(e, emp, isoDate)}
                          className={`p-1 text-center border-r border-zinc-950/5 dark:border-white/5 relative cursor-pointer select-none transition-all ${
                            isCellSelected
                              ? 'bg-blue-500/25 dark:bg-blue-500/35 ring-2 ring-blue-500 ring-inset z-20'
                              : isEditing
                                ? 'z-30 ring-2 ring-blue-500/50 rounded-lg'
                                : isCellRelaxed
                                  ? 'bg-amber-500/[0.06] dark:bg-amber-400/[0.08]'
                                  : isHoliday
                                    ? 'bg-rose-500/[0.03] dark:bg-rose-500/[0.04]'
                                    : isWeekend
                                      ? 'bg-zinc-950/[0.015] dark:bg-white/[0.015]'
                                      : ''
                          }`}
                          title={`${emp} on ${isoDate}: ${rawShift}${relaxationTitleSuffix} (Click to edit, Shift+click to select block, drag to select range)`}
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
                                  ? isCellRelaxed
                                    ? 'bg-amber-500/15 text-amber-700 dark:text-amber-300'
                                    : 'text-zinc-300 dark:text-zinc-600 hover:bg-zinc-950/5 dark:hover:bg-white/5'
                                  : `${colors.bg} ${colors.text} shadow-xs`
                            } ${!isCellSelected && isCellRelaxed ? RELAXED_CELL_RING_CLASS : ''}`}
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
                              {(['M', 'A', 'N', 'MID', 'OFF', 'H8', 'ABS'] as const).map((code) => (
                                <button
                                  key={code}
                                  type="button"
                                  onClick={(e) => {
                                    e.stopPropagation();
                                    handleSetCellShift(code);
                                  }}
                                  className={`px-2 py-1 rounded-lg text-[10px] font-bold cursor-pointer transition-colors ${
                                    sKey === code
                                      ? 'bg-blue-500 text-white shadow-xs'
                                      : code === 'ABS'
                                        ? 'bg-black text-white hover:bg-zinc-800'
                                        : 'hover:bg-zinc-100 dark:hover:bg-white/10 text-zinc-700 dark:text-zinc-300'
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
                    <td className="sticky right-0 z-10 p-2 text-center bg-zinc-100 dark:bg-[#121212] border-l border-zinc-950/10 dark:border-white/10 font-mono text-[10px]">
                      {isExcluded && !hasCustomShifts ? (
                        <span className="inline-flex items-center px-2 py-0.5 rounded font-sans text-[10px] font-bold bg-zinc-500/10 text-zinc-500 dark:text-zinc-400">
                          Off for month
                        </span>
                      ) : stats && (
                        <div className="flex items-center justify-center gap-1.5">
                          <span className="text-green-600 dark:text-green-400 font-bold" title="Morning shifts">{stats.M}M</span>
                          <span className="text-orange-600 dark:text-orange-400 font-bold" title="Afternoon shifts">{stats.A}A</span>
                          <span className="text-purple-600 dark:text-purple-400 font-bold" title="Night shifts">{stats.N}N</span>
                          {stats.MID > 0 && <span className="text-blue-500 font-bold" title="Mid shifts">{stats.MID}MID</span>}
                          {stats.H8 > 0 && <span className="text-pink-500 font-bold" title="H8 Holiday shifts">{stats.H8}H8</span>}
                          {(stats.ABS ?? 0) > 0 && (
                            <span className="bg-black text-white px-1 py-0.5 rounded text-[9px] font-bold" title="Absent / Medical leave">
                              {stats.ABS}ABS
                            </span>
                          )}
                          <span className="text-zinc-400" title="Off days">{stats.OFF}O</span>
                          <span
                            className={`font-bold px-1.5 py-0.5 rounded text-[10px] ${
                              isExcluded
                                ? 'bg-amber-500/15 text-amber-600 dark:text-amber-400'
                                : stats.totalWorkingDays === stats.targetWorkingDays
                                  ? 'bg-emerald-500/15 text-emerald-600 dark:text-emerald-400'
                                  : 'bg-rose-500/15 text-rose-600 dark:text-rose-400'
                            }`}
                            title={isExcluded ? `Custom shifts on deselected employee: ${stats.totalWorkingDays} working days` : `Total working days: ${stats.totalWorkingDays}/${stats.targetWorkingDays} required`}
                          >
                            {isExcluded ? `${stats.totalWorkingDays}d` : `${stats.totalWorkingDays}/${stats.targetWorkingDays}d`}
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
                  const under = cov ? (cov.rotM !== undefined ? cov.rotM < 1 : cov.M < 1) : false;
                  const morningRelaxations = relaxationLookup.getMorningFooterRelaxations(isoDate);
                  const isRelaxed = morningRelaxations.length > 0;
                  const isError = under && !isRelaxed;
                  const baseTitle = cov ? `Total Morning: ${cov.M} (${cov.rotM ?? cov.M} rotating)` : 'Morning coverage';
                  const title = isRelaxed
                    ? `${baseTitle} — Relaxed: ${morningRelaxations.map((r) => r.note).join('; ')}`
                    : baseTitle;

                  return (
                    <td
                      key={isoDate}
                      className={`p-1 text-center font-mono text-[11px] border-r border-zinc-950/5 dark:border-white/5 ${
                        isError
                          ? 'bg-rose-500/20 text-rose-600 font-bold'
                          : isRelaxed
                            ? 'bg-amber-500/5'
                            : 'text-green-700 dark:text-green-400'
                      }`}
                      title={title}
                    >
                      <CoverageCountBadge count={cov?.M ?? 0} isRelaxed={isRelaxed} />
                    </td>
                  );
                })}
                <td className="sticky right-0 z-10 p-1 text-center bg-zinc-100 dark:bg-[#141414] border-l border-zinc-950/10 dark:border-white/10 font-bold text-green-700 dark:text-green-400">
                  M ≥ 1 (rot)
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
                  const afternoonRelaxations = relaxationLookup.getAfternoonFooterRelaxations(isoDate);
                  const isRelaxed = afternoonRelaxations.length > 0;
                  const isError = under && !isRelaxed;
                  const title = isRelaxed
                    ? `Afternoon: ${cov?.A ?? 0} — Relaxed: ${afternoonRelaxations.map((r) => r.note).join('; ')}`
                    : undefined;

                  return (
                    <td
                      key={isoDate}
                      className={`p-1 text-center font-mono text-[11px] border-r border-zinc-950/5 dark:border-white/5 ${
                        isError
                          ? 'bg-rose-500/20 text-rose-600 font-bold'
                          : isRelaxed
                            ? 'bg-amber-500/5'
                            : 'text-orange-700 dark:text-orange-400'
                      }`}
                      title={title}
                    >
                      <CoverageCountBadge count={cov?.A ?? 0} isRelaxed={isRelaxed} />
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
                  const nightRelaxations = relaxationLookup.getNightFooterRelaxations(isoDate);
                  const isRelaxed =
                    nightRelaxations.length > 0 ||
                    (cov?.N === 3 &&
                      complianceState.isNightTargetRelaxed &&
                      !complianceState.requiresReview);
                  const isError = cov
                    ? cov.N < 2 || cov.N > 3 || (cov.N === 3 && !isRelaxed)
                    : false;
                  const title = isRelaxed
                    ? `Night: ${cov?.N ?? 3} (relaxed from N = 2)${nightRelaxations.length > 0 ? ` — ${nightRelaxations.map((r) => r.note).join('; ')}` : ''}`
                    : undefined;

                  return (
                    <td
                      key={isoDate}
                      className={`p-1 text-center font-mono text-[11px] border-r border-zinc-950/5 dark:border-white/5 ${
                        isError
                          ? 'bg-rose-500/20 text-rose-600 font-bold'
                          : isRelaxed
                            ? 'bg-amber-500/5'
                            : 'text-purple-700 dark:text-purple-400'
                      }`}
                      title={title}
                    >
                      <CoverageCountBadge count={cov?.N ?? 0} isRelaxed={isRelaxed} />
                    </td>
                  );
                })}
                <td
                  className={`sticky right-0 z-10 p-1 text-center bg-zinc-100 dark:bg-[#141414] border-l border-zinc-950/10 dark:border-white/10 font-bold ${
                    complianceState.isNightTargetRelaxed
                      ? 'text-amber-700 dark:text-amber-300'
                      : 'text-purple-700 dark:text-purple-400'
                  }`}
                >
                  {complianceState.relaxedNightTargetLabel}
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
                  1 (Mon–Fri, excl. holidays)
                </td>
              </tr>
            </tfoot>
          </table>
        </div>
      </div>

      {/* Warnings & Diagnostics Panel (true hard/coverage violations) */}
      {validation.warnings.length > 0 && (
        <div className={`p-4 ${GLASS_CARD} border border-rose-500/35 bg-rose-500/[0.06] dark:bg-rose-500/[0.08] space-y-2`}>
          <div className="flex flex-wrap items-center justify-between gap-2">
            <div className="flex items-center gap-2 text-rose-700 dark:text-rose-300 font-bold text-[13px]">
              <AlertTriangle className="size-4 shrink-0 text-rose-500" />
              <span>Validation Warnings ({validation.warnings.length})</span>
            </div>
            {validation.warnings.length > 8 && (
              <button
                type="button"
                onClick={() => setShowAllWarnings((prev) => !prev)}
                className="text-[11px] font-semibold text-rose-700 dark:text-rose-300 hover:underline cursor-pointer"
              >
                {showAllWarnings
                  ? 'Show top 8'
                  : `Show all ${validation.warnings.length}`}
              </button>
            )}
          </div>
          <ul className="space-y-1.5 text-[12px] text-zinc-700 dark:text-zinc-300">
            {(showAllWarnings ? validation.warnings : validation.warnings.slice(0, 8)).map((w, i) => {
              const actionNote = getWarningActionNote(w);
              return (
                <li key={i} className="space-y-0.5">
                  <div className="flex items-start gap-1.5 font-mono text-[11px]">
                    <span className="text-rose-500 font-bold">•</span>
                    <span>{w}</span>
                  </div>
                  {actionNote && (
                    <div className="pl-3.5 text-[11px] font-sans font-semibold text-amber-700 dark:text-amber-300 flex items-center gap-1.5">
                      <span className="text-[12px]">👉</span>
                      <span>{actionNote}</span>
                    </div>
                  )}
                </li>
              );
            })}
            {!showAllWarnings && validation.warnings.length > 8 && (
              <li className="font-semibold text-zinc-500 pl-3">
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
              {(['M', 'A', 'N', 'MID', 'OFF', 'H8', 'ABS'] as const).map((code) => {
                const colors = shiftColors[code] ?? { bg: 'bg-zinc-500/10', text: 'text-zinc-500' };
                return (
                  <button
                    key={code}
                    type="button"
                    onClick={() => handleApplyShiftToSelected(code)}
                    className={`px-3 py-1.5 rounded-xl text-[11px] font-bold cursor-pointer transition-all active:scale-95 shadow-xs border border-zinc-950/5 dark:border-white/5 hover:scale-105 ${colors.bg} ${colors.text} hover:ring-2 hover:ring-blue-500/50`}
                    title={`Change all ${selectedCells.size} selected shifts to ${code} (or press ${code === 'OFF' ? 'O' : code === 'H8' ? 'H' : code === 'ABS' ? 'B' : code})`}
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

      {/* Readjustment Suggestion Modal */}
      {isReadjustModalOpen && readjustProposal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-sm animate-in fade-in duration-200">
          <div className="relative w-full max-w-2xl max-h-[90vh] flex flex-col rounded-2xl bg-white dark:bg-[#18181b] border border-zinc-200 dark:border-zinc-800 shadow-2xl overflow-hidden">
            {/* Header */}
            <div className="p-5 border-b border-zinc-100 dark:border-zinc-800 flex items-center justify-between bg-zinc-50/50 dark:bg-zinc-900/50">
              <div className="flex items-center gap-3">
                <div className="size-10 rounded-xl bg-amber-500/15 text-amber-600 dark:text-amber-400 flex items-center justify-center">
                  <RefreshCw className="size-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-zinc-900 dark:text-white flex items-center gap-2">
                    <span>Shift Readjustment Suggestions</span>
                    <span className="text-[11px] font-mono px-2 py-0.5 rounded-full bg-amber-500/15 text-amber-600 dark:text-amber-400 font-bold">
                      Variant #{readjustSeed + 1}
                    </span>
                  </h3>
                  <p className="text-xs text-zinc-500 dark:text-zinc-400 mt-0.5">
                    Minimally invasive coverage for colleagues on medical leave (ABS)
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsReadjustModalOpen(false)}
                className="p-2 rounded-xl text-zinc-400 hover:text-zinc-700 dark:hover:text-zinc-200 hover:bg-zinc-100 dark:hover:bg-zinc-800 transition-colors cursor-pointer"
                title="Close"
              >
                <X className="size-5" />
              </button>
            </div>

            {/* Content Area */}
            <div className="p-5 overflow-y-auto space-y-4 text-xs">
              {/* If no adjustments needed */}
              {readjustProposal.adjustments.length === 0 && readjustProposal.unresolvedDeficits.length === 0 ? (
                <div className="p-4 rounded-xl bg-emerald-500/10 border border-emerald-500/20 text-emerald-800 dark:text-emerald-300 flex items-start gap-3">
                  <CheckCircle2 className="size-5 text-emerald-500 shrink-0 mt-0.5" />
                  <div>
                    <h4 className="font-bold text-sm">No Shifts Need Readjustment</h4>
                    <p className="mt-1 leading-relaxed text-zinc-600 dark:text-zinc-300">
                      Even with colleagues on ABS, every shift still meets the minimum staffing requirement (M ≥ 1, A ≥ 2, N = 2).
                      No colleagues need their shifts changed (0 invasiveness).
                    </p>
                  </div>
                </div>
              ) : (
                <>
                  {/* Status Banner */}
                  <div className="p-3.5 rounded-xl bg-blue-500/10 border border-blue-500/20 text-blue-900 dark:text-blue-200 flex items-center justify-between">
                    <div className="flex items-center gap-2.5">
                      <Clock className="size-4 text-blue-500 shrink-0" />
                      <span>
                        Found <strong className="font-bold">{readjustProposal.adjustments.length}</strong> minimal shift readjustment(s) to restore required staffing.
                      </span>
                    </div>
                  </div>

                  {/* List of Proposed Adjustments */}
                  <div className="space-y-2.5">
                    <h4 className="font-semibold text-zinc-700 dark:text-zinc-300 uppercase tracking-wider text-[11px]">
                      Recommended Shift Changes:
                    </h4>
                    <div className="divide-y divide-zinc-100 dark:divide-zinc-800 border border-zinc-200 dark:border-zinc-800 rounded-xl overflow-hidden">
                      {readjustProposal.adjustments.map((adj, idx) => (
                        <div key={idx} className="p-3.5 bg-white dark:bg-zinc-900 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
                          <div className="space-y-1">
                            <div className="flex items-center gap-2">
                              <span className="font-bold text-zinc-900 dark:text-zinc-100 font-mono text-[12px]">
                                {adj.isoDate}
                              </span>
                              <span className="text-zinc-400">•</span>
                              <span className="font-semibold text-zinc-800 dark:text-zinc-200">
                                {adj.employee}
                              </span>
                            </div>
                            <p className="text-zinc-500 dark:text-zinc-400 text-[11px] leading-relaxed">
                              {adj.reason}
                            </p>
                            <div className="text-[10px] text-zinc-400 flex items-center gap-1">
                              <span>Absent colleague:</span>
                              <span className="bg-black text-white px-1.5 py-0.2 rounded font-mono font-bold">
                                {adj.absColleague}
                              </span>
                            </div>
                          </div>

                          <div className="flex items-center gap-2 shrink-0 self-start sm:self-center">
                            <span className={`px-2 py-1 rounded font-bold text-[11px] ${shiftColors[adj.fromShift]?.bg ?? 'bg-zinc-100'} ${shiftColors[adj.fromShift]?.text ?? 'text-zinc-700'}`}>
                              {adj.fromShift}
                            </span>
                            <span className="text-zinc-400 font-bold">➔</span>
                            <span className={`px-2.5 py-1 rounded font-bold text-[12px] shadow-xs ${shiftSolid[adj.toShift] ?? 'bg-blue-500 text-white'}`}>
                              {adj.toShift}
                            </span>
                          </div>
                        </div>
                      ))}
                    </div>
                  </div>

                  {/* Unresolved Deficits (if any) */}
                  {readjustProposal.unresolvedDeficits.length > 0 && (
                    <div className="p-3.5 rounded-xl bg-amber-500/10 border border-amber-500/20 text-amber-900 dark:text-amber-200 space-y-1">
                      <div className="flex items-center gap-2 font-bold text-amber-800 dark:text-amber-300">
                        <AlertTriangle className="size-4 shrink-0" />
                        <span>Unresolved Deficit Notice</span>
                      </div>
                      <p className="text-[11px] leading-relaxed">
                        {readjustProposal.unresolvedDeficits.map((u, i) => (
                          <span key={i} className="block">
                            • {u.isoDate}: Need {u.needed} on shift {u.shift}. {u.reason}
                          </span>
                        ))}
                      </p>
                    </div>
                  )}

                  {/* Compliance Preview & Warning Details */}
                  <div className="space-y-2">
                    <div className="p-3 rounded-xl bg-zinc-50 dark:bg-zinc-900/60 border border-zinc-200 dark:border-zinc-800 flex items-center justify-between text-[11px]">
                      <span className="text-zinc-500 dark:text-zinc-400">Post-adjustment Compliance:</span>
                      <span className={`font-semibold flex items-center gap-1 ${readjustProposal.validation.isFullyCompliant ? 'text-emerald-600 dark:text-emerald-400' : 'text-amber-600 dark:text-amber-400'}`}>
                        {readjustProposal.validation.isFullyCompliant ? (
                          <>
                            <CheckCircle2 className="size-3.5 text-emerald-500" />
                            <span>Fully Compliant (All Target Norms &amp; Rest Rules Met)</span>
                          </>
                        ) : (
                          <>
                            <AlertTriangle className="size-3.5 text-amber-500" />
                            <span>{readjustProposal.validation.totalViolations} Warning(s) Remaining</span>
                          </>
                        )}
                      </span>
                    </div>

                    {/* Detailed Warnings Breakdown */}
                    {readjustProposal.validation.warnings.length > 0 ? (
                      <div className="rounded-xl border border-amber-500/20 bg-amber-500/5 p-3.5 space-y-2">
                        <div className="flex flex-wrap items-center justify-between gap-2">
                          <div className="flex items-center gap-2 font-bold text-amber-800 dark:text-amber-300 text-xs">
                            <AlertTriangle className="size-4 shrink-0 text-amber-500" />
                            <span>Compliance Warning Details ({readjustProposal.validation.warnings.length})</span>
                          </div>
                          <div className="flex items-center gap-1.5 flex-wrap text-[10px]">
                            {readjustProposal.validation.workDaysMismatchCount > 0 && (
                              <span className="bg-amber-500/15 text-amber-700 dark:text-amber-300 px-2 py-0.5 rounded-full font-mono font-medium">
                                {readjustProposal.validation.workDaysMismatchCount} norm mismatch
                              </span>
                            )}
                            {readjustProposal.validation.understaffedDaysCount > 0 && (
                              <span className="bg-rose-500/15 text-rose-700 dark:text-rose-300 px-2 py-0.5 rounded-full font-mono font-medium">
                                {readjustProposal.validation.understaffedDaysCount} understaffed
                              </span>
                            )}
                            {readjustProposal.validation.nightOverstaffedDaysCount > 0 && (
                              <span className="bg-purple-500/15 text-purple-700 dark:text-purple-300 px-2 py-0.5 rounded-full font-mono font-medium">
                                {readjustProposal.validation.nightOverstaffedDaysCount} night &gt; 2
                              </span>
                            )}
                            {readjustProposal.validation.nightRestViolationCount > 0 && (
                              <span className="bg-purple-500/15 text-purple-700 dark:text-purple-300 px-2 py-0.5 rounded-full font-mono font-medium">
                                {readjustProposal.validation.nightRestViolationCount} rest issue
                              </span>
                            )}
                            {readjustProposal.validation.excessiveRestViolationCount > 0 && (
                              <span className="bg-amber-500/15 text-amber-700 dark:text-amber-300 px-2 py-0.5 rounded-full font-mono font-medium">
                                {readjustProposal.validation.excessiveRestViolationCount} rest &gt; 3d
                              </span>
                            )}
                            {readjustProposal.validation.streakViolationCount > 0 && (
                              <span className="bg-orange-500/15 text-orange-700 dark:text-orange-300 px-2 py-0.5 rounded-full font-mono font-medium">
                                {readjustProposal.validation.streakViolationCount} block/streak issue
                              </span>
                            )}
                            {readjustProposal.validation.mixedBlockViolationCount > 0 && (
                              <span className="bg-rose-500/15 text-rose-700 dark:text-rose-300 px-2 py-0.5 rounded-full font-mono font-medium">
                                {readjustProposal.validation.mixedBlockViolationCount} mixed block
                              </span>
                            )}
                            {(readjustProposal.validation.fixedStaffViolationCount ?? 0) > 0 && (
                              <span className="bg-blue-500/15 text-blue-700 dark:text-blue-300 px-2 py-0.5 rounded-full font-mono font-medium">
                                {readjustProposal.validation.fixedStaffViolationCount} fixed staff
                              </span>
                            )}
                          </div>
                        </div>

                        <div className="max-h-52 overflow-y-auto space-y-1.5 pr-1 font-mono text-[11px]">
                          {readjustProposal.validation.warnings.map((w, idx) => {
                            const actionNote = getWarningActionNote(w);
                            return (
                              <div key={idx} className="p-2 rounded-lg bg-white/80 dark:bg-zinc-900/80 border border-amber-500/20 text-amber-900 dark:text-amber-200 space-y-1">
                                <div className="flex items-start gap-2">
                                  <span className="text-amber-500 shrink-0 font-bold">•</span>
                                  <span className="leading-snug break-words font-medium">{w}</span>
                                </div>
                                {actionNote && (
                                  <div className="pl-4 text-[11px] font-sans font-semibold text-amber-700 dark:text-amber-300 flex items-center gap-1.5 pt-1 border-t border-amber-500/10">
                                    <span className="text-[12px]">👉</span>
                                    <span>{actionNote}</span>
                                  </div>
                                )}
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    ) : (
                      <div className="p-3 rounded-xl bg-emerald-500/10 border border-emerald-500/20 text-emerald-800 dark:text-emerald-300 flex items-center justify-between text-[11px]">
                        <div className="flex items-center gap-2">
                          <CheckCircle2 className="size-4 text-emerald-500 shrink-0" />
                          <span className="font-semibold">Zero Compliance Warnings</span>
                        </div>
                        <span className="text-emerald-700 dark:text-emerald-300">
                          All legal monthly working norms and mandatory rest rules are fully preserved.
                        </span>
                      </div>
                    )}
                  </div>
                </>
              )}
            </div>

            {/* Footer Buttons */}
            <div className="p-4 border-t border-zinc-100 dark:border-zinc-800 bg-zinc-50/50 dark:bg-zinc-900/50 flex flex-wrap items-center justify-between gap-3">
              <button
                type="button"
                onClick={() => setIsReadjustModalOpen(false)}
                className="px-4 py-2 rounded-xl text-xs font-semibold text-zinc-600 dark:text-zinc-300 hover:bg-zinc-200/60 dark:hover:bg-zinc-800 transition-colors cursor-pointer"
              >
                Cancel
              </button>

              <div className="flex items-center gap-2">
                {readjustProposal.adjustments.length > 0 && (
                  <button
                    type="button"
                    onClick={() => handleOpenReadjustModal(readjustSeed + 1)}
                    className="flex items-center gap-1.5 px-3.5 py-2 rounded-xl text-xs font-semibold text-zinc-700 dark:text-zinc-200 bg-zinc-200/70 hover:bg-zinc-300 dark:bg-zinc-800 dark:hover:bg-zinc-700 transition-all cursor-pointer active:scale-98"
                    title="Generate alternative candidate assignment"
                  >
                    <RefreshCw className="size-3.5" />
                    <span>Try Another Variant</span>
                  </button>
                )}

                {readjustProposal.adjustments.length > 0 && (
                  <button
                    type="button"
                    onClick={handleApplyReadjustments}
                    className="flex items-center gap-1.5 px-5 py-2 rounded-xl text-xs font-semibold text-white bg-emerald-600 hover:bg-emerald-700 active:scale-98 transition-all shadow cursor-pointer"
                  >
                    <CheckCircle2 className="size-4" />
                    <span>Accept & Apply Changes</span>
                  </button>
                )}
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
