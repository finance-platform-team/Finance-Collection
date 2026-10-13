import Icon from '../components/Icon';
import { Fragment, useEffect, useMemo, useState } from 'react';
import { Cfm_aragingsService } from '../generated/services/Cfm_aragingsService';
import { Strategy_kpisesService } from '../generated/services/Strategy_kpisesService';
import { Pm_kpiachievmentsService } from '../generated/services/Pm_kpiachievmentsService';
import { Cfm_collectionplanweeksService } from '../generated/services/Cfm_collectionplanweeksService';
import type { Cfm_aragings } from '../generated/models/Cfm_aragingsModel';
import type { Pm_kpiachievments, Pm_kpiachievmentsBase, Pm_kpiachievmentspm_month } from '../generated/models/Pm_kpiachievmentsModel';
import type { Cfm_collectionplanweeks, Cfm_collectionplanweeksBase } from '../generated/models/Cfm_collectionplanweeksModel';
import { operationErrorMessage } from '../lib/errors';
import { KSA_BU_CODES, type Region } from '../lib/useRegion';
import { escapeForCsv, downloadCsv } from '../lib/csvExport';
import './WeeklyTargets.css';

/**
 * Entity chain (confirmed against the live generated schema / Ref.html):
 *   strategy_kpises ("Collection Achievement") -> pm_kpiachievments (one row per
 *   BU+period) -> cfm_collectionplanweeks (one row per week, linked via the
 *   cfm_KPIAchievement / _cfm_kpiachievement_value lookup).
 * The BU lookup on pm_kpiachievments points at the standard Dataverse
 * `businessunits` table (confirmed earlier, matches Ref.html's own bind).
 */
const KPI_NAME = 'Collection Achievement';
const BU_ENTITY_SET = 'businessunits';

const ARAGING_SELECT = ['cfm_aragingid', '_cfm_type_value', 'cfm_targetplan'];
const PARENT_SELECT = ['pm_kpiachievmentid', 'pm_target', 'pm_actual', '_pm_businessunit_value', 'createdon'];
const WEEK_SELECT = [
  'cfm_collectionplanweekid',
  'cfm_weeknumber',
  'cfm_weekfrom',
  'cfm_weekto',
  'cfm_targetpercentage',
  'cfm_targetplan',
  'cfm_weeklytarget',
  'cfm_actualcollected',
  'cfm_achievement',
  '_cfm_kpiachievement_value',
];

function formattedValue(record: unknown, attribute: string): string {
  const raw = (record as Record<string, unknown>)[`${attribute}@OData.Community.Display.V1.FormattedValue`];
  return typeof raw === 'string' ? raw : '';
}

const N = (v: number | null | undefined): number => Number(v || 0);
const fmtSar = (n: number): string => Math.round(n || 0).toLocaleString('en-US');
const fmtDate = (iso: string): string => (iso ? iso.slice(0, 10) : '—');
const fmtPct = (n: number): string => `${n.toFixed(1)}%`;
const norm = (s: string | null | undefined): string => (s == null ? '' : String(s)).trim().toUpperCase();

const MONTHS = ['', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
/** "YYYY-MM" key from a plan's start date — the same key used for both the
 * month filter's options and its default (the current month). */
function monthKeyOf(iso: string): string {
  return (iso || '').slice(0, 7);
}
function monthLabelOf(key: string): string {
  const [y, m] = key.split('-');
  return `${MONTHS[parseInt(m, 10)] || m} ${y}`;
}
function currentMonthKey(): string {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}
const KSA_BU_SET = new Set(KSA_BU_CODES);

type Tone = 'ok' | 'warn' | 'bad';
function achievementTone(pct: number): Tone {
  if (pct < 90) return 'bad';
  if (pct < 105) return 'warn';
  return 'ok';
}
/** kpi-card/kpi-icon variants are named kpi-danger/kpi-icon-danger for the bad
 * tone (shared.css), while the acc-visit-status pill family uses "-bad" — same
 * three tones, different suffixes per shared component. */
const KPI_TONE_SUFFIX: Record<Tone, string> = { ok: 'ok', warn: 'warn', bad: 'danger' };
function achievementPill(pct: number) {
  return <span className={`acc-visit-status acc-visit-status-${achievementTone(pct)}`}>{fmtPct(pct)}</span>;
}

function toISODate(d: Date): string {
  return d.toISOString().slice(0, 10);
}
function addDays(d: Date, n: number): Date {
  const x = new Date(d);
  x.setUTCDate(x.getUTCDate() + n);
  return x;
}
/** Fill the whole From–To span with N contiguous, (near-)equal-length weeks —
 * remainder folded into the last week so the dates always land exactly on `to`. */
function buildWeekRanges(fromStr: string, toStr: string, n: number): { from: string; to: string }[] {
  const from = new Date(`${fromStr}T00:00:00Z`);
  const to = new Date(`${toStr}T00:00:00Z`);
  const totalDays = Math.round((to.getTime() - from.getTime()) / 86400000) + 1;
  const per = Math.floor(totalDays / n);
  const ranges: { from: string; to: string }[] = [];
  let cursor = new Date(from);
  for (let i = 0; i < n; i++) {
    const wFrom = new Date(cursor);
    const wTo = i === n - 1 ? new Date(to) : addDays(wFrom, per - 1);
    ranges.push({ from: toISODate(wFrom), to: toISODate(wTo) });
    cursor = addDays(wTo, 1);
  }
  return ranges;
}

/** From/To auto-fill for the create drawer: starts at the 1st of the current
 * month, spans exactly weekCount×7 days so buildWeekRanges lands on clean
 * ~weekly chunks. */
function computeAutoRange(n: number): { from: string; to: string } {
  const now = new Date();
  const start = new Date(Date.UTC(now.getFullYear(), now.getMonth(), 1));
  const end = addDays(start, n * 7 - 1);
  return { from: toISODate(start), to: toISODate(end) };
}

async function fetchAll<T>(
  run: (skipToken?: string) => Promise<{ success: boolean; data?: T[] | null; error?: unknown; skipToken?: string }>,
  fallback: string
): Promise<T[]> {
  let out: T[] = [];
  let token: string | undefined;
  let guard = 0;
  do {
    const res = await run(token);
    if (!res.success) throw new Error(operationErrorMessage(res.error, fallback));
    out = out.concat(res.data || []);
    token = res.skipToken;
    guard++;
  } while (token && guard < 50);
  return out;
}

interface BuAggregate {
  id: string;
  name: string;
  targetPlanAuto: number;
}

interface WeekRow {
  id: string;
  weekNumber: number;
  from: string;
  to: string;
  targetPlanAuto: number;
  percent: number;
  weeklyTarget: number;
  actualCollected: number;
  achievement: number;
}

interface PlanGroup {
  id: string;
  buId: string;
  buName: string;
  from: string;
  to: string;
  targetPlanTotal: number;
  createdOn: string;
  weeks: WeekRow[];
}

interface GroupTotals {
  targetPlan: number;
  manualPct: number;
  weeklyTarget: number;
  actualCollected: number;
  achievement: number;
}

function groupTotals(g: PlanGroup): GroupTotals {
  const targetPlan = g.targetPlanTotal;
  const weeklyTarget = g.weeks.reduce((s, w) => s + w.weeklyTarget, 0);
  const actualCollected = g.weeks.reduce((s, w) => s + w.actualCollected, 0);
  const manualPct = g.weeks.reduce((s, w) => s + w.percent, 0);
  const achievement = weeklyTarget > 0 ? (actualCollected / weeklyTarget) * 100 : 0;
  return { targetPlan, manualPct, weeklyTarget, actualCollected, achievement };
}

/** Target Plan (Auto) for every week is the plan's Target Plan Total, unsplit —
 * the same full number on every week row, never fractured across weeks.
 * Weekly Target is that Total scaled by the week's own % Manual (each week
 * independently, not a share that has to sum to 100% across the plan). */
function recalcWeek(w: WeekRow, targetPlanTotal: number): WeekRow {
  const targetPlanAuto = Math.round(targetPlanTotal);
  const weeklyTarget = Math.round((targetPlanTotal * w.percent) / 100);
  const achievement = weeklyTarget > 0 ? (w.actualCollected / weeklyTarget) * 100 : 0;
  return { ...w, targetPlanAuto, weeklyTarget, achievement };
}

interface WeeklyTargetsPageProps {
  region: Region | null;
}

export default function WeeklyTargetsPage({ region }: WeeklyTargetsPageProps) {
  const [kpiId, setKpiId] = useState<string | null>(null);
  const [kpiError, setKpiError] = useState<string | null>(null);

  const [allBuAggregates, setAllBuAggregates] = useState<BuAggregate[]>([]);
  const [allGroups, setAllGroups] = useState<PlanGroup[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [expanded, setExpanded] = useState<Set<string>>(new Set());
  const [dirtyWeekIds, setDirtyWeekIds] = useState<Set<string>>(new Set());
  const [saving, setSaving] = useState(false);
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(15);
  // Defaults to the current month; '' means "All months".
  const [monthFilter, setMonthFilter] = useState(currentMonthKey);

  const [createOpen, setCreateOpen] = useState(false);
  const [buSearch, setBuSearch] = useState('');
  const [buResultsOpen, setBuResultsOpen] = useState(false);
  const [selectedBu, setSelectedBu] = useState<BuAggregate | null>(null);
  const [from, setFrom] = useState('');
  const [to, setTo] = useState('');
  const [weekCount, setWeekCount] = useState('4');
  const [creating, setCreating] = useState(false);
  const [createError, setCreateError] = useState('');

  const [toastMsg, setToastMsg] = useState('');
  const [toastOn, setToastOn] = useState(false);

  function showToast(msg: string, durationMs = 3000) {
    setToastMsg(msg);
    setToastOn(true);
    setTimeout(() => setToastOn(false), durationMs);
  }

  async function load() {
    setLoading(true);
    setLoadError(null);
    let resolvedKpiId = kpiId;
    if (!resolvedKpiId) {
      try {
        const kpiRes = await Strategy_kpisesService.getAll({
          select: ['strategy_kpisid'],
          filter: `strategy_newcolumn eq '${KPI_NAME.replace(/'/g, "''")}' and statecode eq 0`,
          top: 1,
        });
        if (!kpiRes.success) throw new Error(operationErrorMessage(kpiRes.error, 'Could not resolve KPI.'));
        resolvedKpiId = kpiRes.data?.[0]?.strategy_kpisid || null;
        if (!resolvedKpiId) throw new Error(`KPI "${KPI_NAME}" not found.`);
        setKpiId(resolvedKpiId);
      } catch (e) {
        setKpiError(e instanceof Error ? e.message : String(e));
        setLoading(false);
        return;
      }
    }

    try {
      const [aragingRows, parentRows, weekRows] = await Promise.all([
        fetchAll<Cfm_aragings>(
          (skipToken) => Cfm_aragingsService.getAll({ select: ARAGING_SELECT, filter: 'statecode eq 0', maxPageSize: 5000, skipToken }),
          'Could not load AR aging.'
        ),
        fetchAll<Pm_kpiachievments>(
          (skipToken) =>
            Pm_kpiachievmentsService.getAll({
              select: PARENT_SELECT,
              filter: `_pm_kpi_value eq ${resolvedKpiId} and statecode eq 0`,
              orderBy: ['createdon desc'],
              maxPageSize: 5000,
              skipToken,
            }),
          'Could not load Weekly Target plans.'
        ),
        fetchAll<Cfm_collectionplanweeks>(
          (skipToken) => Cfm_collectionplanweeksService.getAll({ select: WEEK_SELECT, filter: 'statecode eq 0', maxPageSize: 5000, skipToken }),
          'Could not load Collection Plan Weeks.'
        ),
      ]);

      // Requirement: group cfm_aragings by _cfm_type_value (the BU lookup) and sum
      // cfm_targetplan per type — this sum is the "Target Plan (Auto)" for that BU.
      const buMap = new Map<string, BuAggregate>();
      aragingRows.forEach((r) => {
        const id = r._cfm_type_value;
        if (!id) return;
        const name = formattedValue(r, '_cfm_type_value') || '—';
        const existing = buMap.get(id);
        if (existing) existing.targetPlanAuto += N(r.cfm_targetplan);
        else buMap.set(id, { id, name, targetPlanAuto: N(r.cfm_targetplan) });
      });
      const aggregates = Array.from(buMap.values()).sort((a, b) => a.name.localeCompare(b.name));

      const weeksByParent = new Map<string, WeekRow[]>();
      weekRows.forEach((w) => {
        const parentId = w._cfm_kpiachievement_value;
        if (!parentId) return;
        const row: WeekRow = {
          id: w.cfm_collectionplanweekid,
          weekNumber: N(w.cfm_weeknumber),
          from: w.cfm_weekfrom || '',
          to: w.cfm_weekto || '',
          targetPlanAuto: N(w.cfm_targetplan),
          percent: N(w.cfm_targetpercentage),
          weeklyTarget: N(w.cfm_weeklytarget),
          actualCollected: N(w.cfm_actualcollected),
          achievement: N(w.cfm_achievement),
        };
        const list = weeksByParent.get(parentId);
        if (list) list.push(row);
        else weeksByParent.set(parentId, [row]);
      });
      weeksByParent.forEach((list) => list.sort((a, b) => a.weekNumber - b.weekNumber));

      const planGroups: PlanGroup[] = parentRows.map((p) => {
        const targetPlanTotal = N(p.pm_target);
        // Recompute every week's Target Plan (Auto) / Weekly Target from the
        // parent's live total rather than trusting whatever was last persisted
        // on cfm_targetplan/cfm_weeklytarget — this keeps plans created before
        // the Total-per-week model (or edited outside this page) from showing
        // stale, per-week-fractured amounts. Corrects itself in the DB the
        // next time any week on the plan is saved.
        const weeks = (weeksByParent.get(p.pm_kpiachievmentid) || []).map((w) => recalcWeek(w, targetPlanTotal));
        return {
          id: p.pm_kpiachievmentid,
          buId: p._pm_businessunit_value || '',
          buName: formattedValue(p, '_pm_businessunit_value') || '—',
          from: weeks[0]?.from || '',
          to: weeks[weeks.length - 1]?.to || '',
          targetPlanTotal,
          createdOn: p.createdon || '',
          weeks,
        };
      });
      planGroups.sort((a, b) => b.createdOn.localeCompare(a.createdOn));

      setAllBuAggregates(aggregates);
      setAllGroups(planGroups);
      setDirtyWeekIds(new Set());
    } catch (e) {
      setLoadError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    load();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /** The BU column partitions rows by region (KSA_BU_CODES vs everything else),
   * same hard partition used on Collection Plan — applied before anything else
   * on the page (cards, table, and the BU picker inside the create drawer). */
  const buAggregates = useMemo(() => {
    if (!region) return allBuAggregates;
    return allBuAggregates.filter((b) => (region === 'KSA' ? KSA_BU_SET.has(norm(b.name)) : !KSA_BU_SET.has(norm(b.name))));
  }, [allBuAggregates, region]);

  const regionGroups = useMemo(() => {
    if (!region) return allGroups;
    return allGroups.filter((g) => (region === 'KSA' ? KSA_BU_SET.has(norm(g.buName)) : !KSA_BU_SET.has(norm(g.buName))));
  }, [allGroups, region]);

  /** Options are built from the region-scoped set (before the month filter
   * itself is applied), so switching regions never leaves a stale month
   * selected that has no data in the new region. */
  const monthOptions = useMemo(() => {
    const map = new Map<string, string>();
    regionGroups.forEach((g) => {
      const key = monthKeyOf(g.from);
      if (key && !map.has(key)) map.set(key, monthLabelOf(key));
    });
    return Array.from(map.entries()).sort((a, b) => b[0].localeCompare(a[0]));
  }, [regionGroups]);

  const groups = useMemo(() => {
    if (!monthFilter) return regionGroups;
    return regionGroups.filter((g) => monthKeyOf(g.from) === monthFilter);
  }, [regionGroups, monthFilter]);

  /** BU picker source for the create drawer: union of BUs with AR Aging data
   * (which carry a real Target Plan Auto) and BUs that already have a plan but
   * no AR Aging yet (Target Plan Auto shows as 0 for those). Deliberately sourced
   * from the unfiltered allBuAggregates/allGroups (not the region-scoped
   * buAggregates/groups above) — the region split only governs what's shown on
   * the dashboard, not what business units can be searched when creating a plan. */
  const buOptions = useMemo(() => {
    const map = new Map(allBuAggregates.map((b) => [b.id, b]));
    allGroups.forEach((g) => {
      if (g.buId && !map.has(g.buId)) map.set(g.buId, { id: g.buId, name: g.buName, targetPlanAuto: 0 });
    });
    return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name));
  }, [allBuAggregates, allGroups]);

  const allExpanded = groups.length > 0 && groups.every((g) => expanded.has(g.id));

  function toggleExpand(groupId: string) {
    setExpanded((prev) => {
      const next = new Set(prev);
      if (next.has(groupId)) next.delete(groupId);
      else next.add(groupId);
      return next;
    });
  }

  function toggleExtendAll() {
    setExpanded(allExpanded ? new Set() : new Set(groups.map((g) => g.id)));
  }

  /** A BU can have more than one plan (e.g. re-created for a new period) — the
   * summary card rolls up every plan for that BU, and clicking it opens/closes
   * all of that BU's rows in the table together, rather than reflecting (or
   * toggling) only whichever plan happened to be found first. */
  function buGroupsFor(buId: string): PlanGroup[] {
    return groups.filter((g) => g.buId === buId);
  }

  function buTotals(buId: string): GroupTotals {
    const list = buGroupsFor(buId);
    let targetPlan = 0;
    let weeklyTarget = 0;
    let actualCollected = 0;
    list.forEach((g) => {
      const t = groupTotals(g);
      targetPlan += t.targetPlan;
      weeklyTarget += t.weeklyTarget;
      actualCollected += t.actualCollected;
    });
    const achievement = weeklyTarget > 0 ? (actualCollected / weeklyTarget) * 100 : 0;
    const manualPct = list.length ? list.reduce((s, g) => s + groupTotals(g).manualPct, 0) / list.length : 0;
    return { targetPlan, manualPct, weeklyTarget, actualCollected, achievement };
  }

  function toggleExpandBu(buId: string) {
    const ids = buGroupsFor(buId).map((g) => g.id);
    const allOpen = ids.length > 0 && ids.every((id) => expanded.has(id));
    setExpanded((prev) => {
      const next = new Set(prev);
      ids.forEach((id) => (allOpen ? next.delete(id) : next.add(id)));
      return next;
    });
  }

  function updateWeekField(groupId: string, weekId: string, field: 'percent' | 'actualCollected', value: number): void;
  function updateWeekField(groupId: string, weekId: string, field: 'from' | 'to', value: string): void;
  function updateWeekField(groupId: string, weekId: string, field: 'percent' | 'actualCollected' | 'from' | 'to', value: number | string) {
    setAllGroups((prev) =>
      prev.map((g) => {
        if (g.id !== groupId) return g;
        return {
          ...g,
          weeks: g.weeks.map((w) => {
            if (w.id !== weekId) return w;
            if (field === 'from' || field === 'to') return { ...w, [field]: value as string };
            return recalcWeek({ ...w, [field]: value as number }, g.targetPlanTotal);
          }),
        };
      })
    );
    setDirtyWeekIds((prev) => new Set(prev).add(weekId));
  }

  const dirtyCount = dirtyWeekIds.size;

  async function saveUpdates() {
    if (dirtyCount === 0) return;
    setSaving(true);
    try {
      const dirtyWeeks = groups.flatMap((g) => g.weeks.filter((w) => dirtyWeekIds.has(w.id)));
      const weekResults = await Promise.all(
        dirtyWeeks.map((w) =>
          Cfm_collectionplanweeksService.update(w.id, {
            cfm_targetpercentage: Math.round(w.percent),
            cfm_targetplan: w.targetPlanAuto,
            cfm_weeklytarget: w.weeklyTarget,
            cfm_actualcollected: Math.round(w.actualCollected),
            cfm_achievement: w.achievement,
            cfm_weekfrom: w.from,
            cfm_weekto: w.to,
          })
        )
      );
      const weekFailed = weekResults.filter((r) => !r.success).length;

      const affectedGroups = groups.filter((g) => g.weeks.some((w) => dirtyWeekIds.has(w.id)));
      const headerResults = await Promise.all(
        affectedGroups.map((g) => {
          const totals = groupTotals(g);
          return Pm_kpiachievmentsService.update(g.id, { pm_actual: Math.round(totals.actualCollected) });
        })
      );
      const headerFailed = headerResults.filter((r) => !r.success).length;

      setDirtyWeekIds(new Set());
      const totalFailed = weekFailed + headerFailed;
      showToast(totalFailed ? `Saved with ${totalFailed} error(s) — some updates may not have persisted.` : 'Updates saved.', totalFailed ? 5000 : 2600);
      await load();
    } catch (e) {
      showToast(e instanceof Error ? e.message : 'Save failed.', 5000);
    } finally {
      setSaving(false);
    }
  }

  function openCreateDrawer(presetBu?: BuAggregate) {
    setSelectedBu(presetBu || null);
    setBuSearch('');
    setBuResultsOpen(false);
    const auto = computeAutoRange(4);
    setFrom(auto.from);
    setTo(auto.to);
    setWeekCount('4');
    setCreateError('');
    setCreateOpen(true);
  }
  function closeCreateDrawer() {
    setCreateOpen(false);
  }

  /** Behaves like a dropdown, not a search-first box: with nothing typed it
   * lists every available BU; typing narrows it. Opens on focus either way. */
  const buMatches = useMemo(() => {
    const q = buSearch.trim().toLowerCase();
    return q ? buOptions.filter((b) => b.name.toLowerCase().includes(q)) : buOptions;
  }, [buSearch, buOptions]);

  function selectBu(bu: BuAggregate) {
    setSelectedBu(bu);
    setBuSearch('');
    setBuResultsOpen(false);
  }
  function clearBu() {
    setSelectedBu(null);
  }

  /** Every week gets the full, un-split Target Plan Total (read from the BU's
   * AR Aging sum) and starts at 100% Manual (so Weekly Target = Total at first
   * — adjust % per week afterward), fills the picked From–To range with
   * contiguous week dates, then creates one cfm_collectionplanweeks child per
   * week bound to the new pm_kpiachievments parent via
   * cfm_KPIAchievement@odata.bind. */
  async function handleCreate() {
    if (!selectedBu) {
      setCreateError('Select a Business Unit.');
      return;
    }
    const total = Math.round(selectedBu.targetPlanAuto);
    if (!from || !to) {
      setCreateError('Enter From and To.');
      return;
    }
    if (new Date(to) < new Date(from)) {
      setCreateError('To must be after From.');
      return;
    }
    const n = parseInt(weekCount, 10);
    if (!n || n < 1 || n > 52) {
      setCreateError('Enter a valid week count (1-52).');
      return;
    }

    setCreating(true);
    setCreateError('');
    try {
      const ranges = buildWeekRanges(from, to, n);
      const fromDate = new Date(`${from}T00:00:00Z`);

      const parentBody = {
        'pm_businessunit@odata.bind': `/${BU_ENTITY_SET}(${selectedBu.id})`,
        'pm_kpi@odata.bind': `/strategy_kpises(${kpiId})`,
        pm_month: (fromDate.getUTCMonth() + 1) as Pm_kpiachievmentspm_month,
        pm_year: fromDate.getUTCFullYear(),
        pm_target: total,
        pm_actual: 0,
      };
      const parentResult = await Pm_kpiachievmentsService.create(parentBody as unknown as Omit<Pm_kpiachievmentsBase, 'pm_kpiachievmentid'>);
      if (!parentResult.success) throw new Error(operationErrorMessage(parentResult.error, 'Could not create the plan.'));
      const parentId = (parentResult.data as unknown as { pm_kpiachievmentid?: string })?.pm_kpiachievmentid;
      if (!parentId) throw new Error('Plan was created but no id was returned.');

      const weekBodies = ranges.map((r, i) => ({
        cfm_weeknumber: i + 1,
        cfm_weekfrom: `${r.from}T00:00:00Z`,
        cfm_weekto: `${r.to}T00:00:00Z`,
        cfm_targetpercentage: 100,
        cfm_targetplan: total,
        cfm_weeklytarget: total,
        cfm_actualcollected: 0,
        cfm_achievement: 0,
        'cfm_KPIAchievement@odata.bind': `/pm_kpiachievments(${parentId})`,
      }));
      const weekResults = await Promise.all(
        weekBodies.map((b) => Cfm_collectionplanweeksService.create(b as unknown as Omit<Cfm_collectionplanweeksBase, 'cfm_collectionplanweekid'>))
      );
      const failedWeeks = weekResults.filter((r) => !r.success).length;

      setCreateOpen(false);
      showToast(failedWeeks ? `Plan created, but ${failedWeeks} of ${n} week(s) failed.` : `Weekly Target created for ${selectedBu.name}.`, failedWeeks ? 5000 : 3000);
      setExpanded((prev) => new Set(prev).add(parentId));
      await load();
    } catch (e) {
      setCreateError(e instanceof Error ? e.message : 'Create failed.');
    } finally {
      setCreating(false);
    }
  }

  const grandTotals = useMemo(() => {
    let targetPlan = 0;
    let weeklyTarget = 0;
    let actualCollected = 0;
    groups.forEach((g) => {
      const t = groupTotals(g);
      targetPlan += t.targetPlan;
      weeklyTarget += t.weeklyTarget;
      actualCollected += t.actualCollected;
    });
    const achievement = weeklyTarget > 0 ? (actualCollected / weeklyTarget) * 100 : 0;
    return { targetPlan, weeklyTarget, actualCollected, achievement };
  }, [groups]);

  // Grand total (footer) still sums every group, not just the current page.
  const totalPages = Math.max(1, Math.ceil(groups.length / pageSize));
  const clampedPage = Math.min(page, totalPages);
  const pageGroups = groups.slice((clampedPage - 1) * pageSize, (clampedPage - 1) * pageSize + pageSize);

  useEffect(() => {
    setPage(1);
  }, [groups.length]);

  function pagerButtons(): (number | 'ellipsis')[] {
    const pages: (number | 'ellipsis')[] = [];
    for (let p = 1; p <= totalPages; p++) {
      if (p === 1 || p === totalPages || Math.abs(p - clampedPage) <= 1) pages.push(p);
      else if (Math.abs(p - clampedPage) === 2) pages.push('ellipsis');
    }
    return pages;
  }

  function exportCsv() {
    if (!groups.length) {
      showToast('Nothing to export.');
      return;
    }
    const header = ['Business Unit', 'Week', 'From', 'To', 'Target Plan (Auto)', '% Manual', 'Weekly Target', 'Actual Collected', 'Achievement %'];
    const lines = [header.join(',')];
    groups.forEach((g) => {
      const t = groupTotals(g);
      lines.push(
        [g.buName, `${g.weeks.length} Weeks (Total)`, fmtDate(g.from), fmtDate(g.to), Math.round(t.targetPlan), t.manualPct.toFixed(0), Math.round(t.weeklyTarget), Math.round(t.actualCollected), t.achievement.toFixed(1)]
          .map(escapeForCsv)
          .join(',')
      );
      g.weeks.forEach((w) => {
        lines.push(
          [g.buName, `Week ${w.weekNumber}`, fmtDate(w.from), fmtDate(w.to), Math.round(w.targetPlanAuto), w.percent, Math.round(w.weeklyTarget), Math.round(w.actualCollected), w.achievement.toFixed(1)]
            .map(escapeForCsv)
            .join(',')
        );
      });
    });
    downloadCsv(`weekly_targets_${new Date().toISOString().slice(0, 10)}.csv`, lines);
  }

  return (
    <>
      <div className="acc-wrap wt-page">
        <div className="acc-head">
          <div>
            <h1 className="acc-title">Weekly Targets &amp; Achievement</h1>
            <div className="acc-sub">Collection plan execution, Daily Collection entries, AR decision feedback, follow-up and dispute operations.</div>
          </div>
          <div className="acc-actions">
            <button className="acc-btn acc-btn-gold" onClick={() => openCreateDrawer()}>
              <Icon className="fa-solid fa-plus" /> New
            </button>
            <button className="acc-btn" onClick={toggleExtendAll}>
              <Icon className={`fa-solid ${allExpanded ? 'fa-angles-up' : 'fa-angles-down'}`} /> {allExpanded ? 'Collapse all' : 'Extend all'}
            </button>
            <button className="acc-btn" onClick={exportCsv}>
              <Icon className="fa-solid fa-file-export" /> Export
            </button>
            <button className={`acc-btn${dirtyCount > 0 ? ' acc-btn-primary' : ''}`} disabled={saving || dirtyCount === 0} onClick={saveUpdates}>
              <Icon className="fa-solid fa-floppy-disk" /> {saving ? 'Saving…' : 'Save updates'}
              {dirtyCount > 0 && !saving && <span className="wt-count-badge">{dirtyCount}</span>}
            </button>
          </div>
        </div>

        {kpiError ? (
          <div className="acc-state">
            <Icon className="fa-solid fa-triangle-exclamation" />
            {kpiError}
          </div>
        ) : !loading && buAggregates.length === 0 && groups.length === 0 ? (
          <div className="wt-empty-hero">
            <div className="wt-empty-icon">
              <Icon className="fa-solid fa-bullseye" />
            </div>
            <h3>Set your first weekly target</h3>
            <p>Track collection targets against actual achievement for this region, week by week.</p>
            <div className="wt-empty-note">
              <Icon className="fa-solid fa-circle-info" />
              {region
                ? 'No AR Aging or collection plans found for this region yet.'
                : "No AR Aging records found, so there's no Target Plan (Auto) to build a plan from yet."}
            </div>
            <button className="acc-btn acc-btn-gold" onClick={() => openCreateDrawer()}>
              <Icon className="fa-solid fa-plus" /> New weekly target
            </button>
          </div>
        ) : (
          <>
            {(loading || buAggregates.length > 0) && (
              <div className="kpi-grid">
                {loading
                  ? Array.from({ length: 3 }).map((_, i) => (
                      <div className="kpi-card" key={i}>
                        <div className="acc-skel" style={{ width: '50%', height: 14, marginBottom: 14 }}></div>
                        <div className="acc-skel" style={{ width: '35%', height: 22, marginBottom: 10 }}></div>
                        <div className="acc-skel" style={{ width: '75%', height: 12 }}></div>
                      </div>
                    ))
                  : buAggregates.map((bu) => {
                      const buGroups = buGroupsFor(bu.id);
                      if (buGroups.length === 0) {
                        return (
                          <div className="kpi-card" key={bu.id}>
                            <div className="kpi-top">
                              <span className="kpi-icon kpi-icon-gold">
                                <Icon className="fa-solid fa-building-shield" />
                              </span>
                              <span className="kpi-label">{bu.name}</span>
                            </div>
                            <div className="kpi-value">—</div>
                            <div className="kpi-context">Target Plan (Auto) {fmtSar(bu.targetPlanAuto)} · no plan yet</div>
                          </div>
                        );
                      }
                      const t = buTotals(bu.id);
                      const isOpen = buGroups.every((g) => expanded.has(g.id));
                      const toneSuffix = KPI_TONE_SUFFIX[achievementTone(t.achievement)];
                      return (
                        <div
                          className={`kpi-card kpi-clickable kpi-${toneSuffix}`}
                          key={bu.id}
                          onClick={() => toggleExpandBu(bu.id)}
                          title={isOpen ? 'Click to collapse weekly details' : 'Click to open weekly details'}
                        >
                          <div className="kpi-top">
                            <span className={`kpi-icon kpi-icon-${toneSuffix}`}>
                              <Icon className={`fa-solid ${isOpen ? 'fa-chevron-up' : 'fa-chevron-down'}`} />
                            </span>
                            <span className="kpi-label">
                              {bu.name}
                              {buGroups.length > 1 ? ` · ${buGroups.length} plans` : ''}
                            </span>
                          </div>
                          <div className="kpi-value">{fmtPct(t.achievement)}</div>
                          <div className="kpi-context">
                            Weekly Target {fmtSar(t.weeklyTarget)} · Collected {fmtSar(t.actualCollected)}
                          </div>
                        </div>
                      );
                    })}
              </div>
            )}

            <div className="tbl-shell">
              <div className="tbl-hdr">
                <select className="acc-f" value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)}>
                  <option value="">All months</option>
                  {monthOptions.map(([key, label]) => (
                    <option key={key} value={key}>
                      {label}
                    </option>
                  ))}
                </select>
              </div>
              <div className="wt-tbl-scroll">
                <table className="wt-tbl">
                  <thead>
                    <tr>
                      <th>Business Unit</th>
                      <th>Weeks</th>
                      <th>From</th>
                      <th>To</th>
                      <th className="wt-numeric">Target Plan (Auto)</th>
                      <th className="wt-numeric">% (Manual)</th>
                      <th className="wt-numeric">Weekly Target</th>
                      <th className="wt-numeric">Actual Collected</th>
                      <th className="wt-numeric">Achievement %</th>
                    </tr>
                  </thead>
                  <tbody>
                    {loading ? (
                      Array.from({ length: 4 }).map((_, i) => (
                        <tr key={i}>
                          <td colSpan={9}>
                            <div className="acc-skel" style={{ width: `${55 + ((i * 13) % 35)}%` }}></div>
                          </td>
                        </tr>
                      ))
                    ) : loadError ? (
                      <tr>
                        <td colSpan={9}>
                          <div className="acc-state">
                            <Icon className="fa-solid fa-triangle-exclamation" />
                            Couldn't load Weekly Targets.
                            <br />
                            <small>{loadError}</small>
                          </div>
                        </td>
                      </tr>
                    ) : groups.length === 0 ? (
                      <tr>
                        <td colSpan={9}>
                          <div className="acc-state">
                            <Icon className="fa-solid fa-inbox" />
                            {monthFilter
                              ? `No weekly targets for ${monthLabelOf(monthFilter)} — try "All months" or use "New" to create one.`
                              : 'No weekly targets yet — use "New" to create one.'}
                          </div>
                        </td>
                      </tr>
                    ) : (
                      pageGroups.map((g) => {
                        const t = groupTotals(g);
                        const isOpen = expanded.has(g.id);
                        return (
                          <Fragment key={g.id}>
                            <tr className="wt-group-row" onClick={() => toggleExpand(g.id)}>
                              <td className="wt-col-bold">{g.buName}</td>
                              <td>
                                <span className="wt-weeks-select">
                                  {g.weeks.length} week{g.weeks.length === 1 ? '' : 's'}
                                  <Icon className={`fa-solid ${isOpen ? 'fa-chevron-up' : 'fa-chevron-down'}`} />
                                </span>
                              </td>
                              <td>{fmtDate(g.from)}</td>
                              <td>{fmtDate(g.to)}</td>
                              <td className="wt-numeric wt-col-bold">{fmtSar(t.targetPlan)}</td>
                              <td className="wt-numeric">{t.manualPct.toFixed(0)}%</td>
                              <td className="wt-numeric wt-col-bold">{fmtSar(t.weeklyTarget)}</td>
                              <td className="wt-numeric wt-col-bold">{fmtSar(t.actualCollected)}</td>
                              <td className="wt-numeric">{achievementPill(t.achievement)}</td>
                            </tr>
                            {isOpen &&
                              (g.weeks.length === 0 ? (
                                <tr>
                                  <td colSpan={9}>
                                    <div className="acc-state" style={{ padding: '20px' }}>
                                      No weeks configured for this plan.
                                    </div>
                                  </td>
                                </tr>
                              ) : (
                                g.weeks.map((w, wi) => (
                                  <tr className="wt-week-row" key={w.id}>
                                    <td>{g.buName}</td>
                                    <td>Week {w.weekNumber}</td>
                                    <td>
                                      {wi === 0 ? (
                                        fmtDate(w.from)
                                      ) : (
                                        <input
                                          type="date"
                                          className="wt-date-input"
                                          value={fmtDate(w.from)}
                                          onClick={(e) => e.stopPropagation()}
                                          onChange={(e) => updateWeekField(g.id, w.id, 'from', `${e.target.value}T00:00:00Z`)}
                                        />
                                      )}
                                    </td>
                                    <td>
                                      {wi === g.weeks.length - 1 ? (
                                        fmtDate(w.to)
                                      ) : (
                                        <input
                                          type="date"
                                          className="wt-date-input"
                                          value={fmtDate(w.to)}
                                          onClick={(e) => e.stopPropagation()}
                                          onChange={(e) => updateWeekField(g.id, w.id, 'to', `${e.target.value}T00:00:00Z`)}
                                        />
                                      )}
                                    </td>
                                    <td className="wt-numeric">{fmtSar(w.targetPlanAuto)}</td>
                                    <td className="wt-numeric">
                                      <span className="wt-pct-input">
                                        <input
                                          type="number"
                                          min={0}
                                          max={100}
                                          step={1}
                                          value={w.percent}
                                          onClick={(e) => e.stopPropagation()}
                                          onChange={(e) => {
                                            const v = Math.max(0, Math.min(100, Number(e.target.value) || 0));
                                            updateWeekField(g.id, w.id, 'percent', v);
                                          }}
                                        />
                                        <span>%</span>
                                      </span>
                                    </td>
                                    <td className="wt-numeric">{fmtSar(w.weeklyTarget)}</td>
                                    <td className="wt-numeric">
                                      <input
                                        type="number"
                                        min={0}
                                        step={1}
                                        className="wt-num-input"
                                        value={w.actualCollected}
                                        onClick={(e) => e.stopPropagation()}
                                        onChange={(e) => {
                                          const v = Math.max(0, Number(e.target.value) || 0);
                                          updateWeekField(g.id, w.id, 'actualCollected', v);
                                        }}
                                      />
                                    </td>
                                    <td className="wt-numeric">{achievementPill(w.achievement)}</td>
                                  </tr>
                                ))
                              ))}
                          </Fragment>
                        );
                      })
                    )}
                  </tbody>
                  {!loading && !loadError && groups.length > 0 && (
                    <tfoot>
                      <tr>
                        <td colSpan={4} className="wt-col-bold">
                          <span style={{ color: 'var(--gold)', marginRight: 6, fontWeight: 700 }}>Σ</span>
                          Grand total
                        </td>
                        <td className="wt-numeric wt-col-bold">{fmtSar(grandTotals.targetPlan)}</td>
                        <td></td>
                        <td className="wt-numeric wt-col-bold">{fmtSar(grandTotals.weeklyTarget)}</td>
                        <td className="wt-numeric wt-col-bold">{fmtSar(grandTotals.actualCollected)}</td>
                        <td className="wt-numeric">{achievementPill(grandTotals.achievement)}</td>
                      </tr>
                    </tfoot>
                  )}
                </table>
              </div>
              {!loading && !loadError && groups.length > 0 && (
                <div className="acc-foot">
                  <span>
                    {groups.length} business unit{groups.length !== 1 ? 's' : ''}
                  </span>
                  <div className="acc-pg">
                    <button disabled={clampedPage <= 1} onClick={() => setPage(clampedPage - 1)}>
                      ‹
                    </button>
                    {pagerButtons().map((p, i) =>
                      p === 'ellipsis' ? (
                        <span key={`e${i}`} style={{ padding: '0 4px' }}>
                          …
                        </span>
                      ) : (
                        <button key={p} className={p === clampedPage ? 'active' : ''} onClick={() => setPage(p)}>
                          {p}
                        </button>
                      )
                    )}
                    <button disabled={clampedPage >= totalPages} onClick={() => setPage(clampedPage + 1)}>
                      ›
                    </button>
                  </div>
                  <select
                    className="acc-f"
                    style={{ minWidth: 90 }}
                    value={pageSize}
                    onChange={(e) => {
                      setPageSize(parseInt(e.target.value, 10));
                      setPage(1);
                    }}
                  >
                    <option value={10}>10</option>
                    <option value={15}>15</option>
                    <option value={20}>20</option>
                  </select>
                </div>
              )}
            </div>
          </>
        )}
      </div>

      <div className={`acc-scrim${createOpen ? ' on' : ''}`} onClick={closeCreateDrawer}></div>
      <aside className={`acc-drawer${createOpen ? ' on' : ''}`}>
        <div className="acc-dhead">
          <h2>New Weekly Target</h2>
          <button className="acc-dclose" onClick={closeCreateDrawer}>
            <Icon className="fa-solid fa-xmark" />
          </button>
        </div>
        <div className="acc-dbody">
          <div className={`acc-err${createError ? ' on' : ''}`}>{createError}</div>
          <div className="acc-fld">
            <label>
              Business Unit <span className="dm-req">*</span>
            </label>
            <div className="wt-lookup">
              <input
                placeholder="Type to search, or click to pick a business unit…"
                autoComplete="off"
                value={buSearch}
                onChange={(e) => {
                  setBuSearch(e.target.value);
                  setBuResultsOpen(true);
                }}
                onFocus={() => setBuResultsOpen(true)}
                onBlur={() => setTimeout(() => setBuResultsOpen(false), 150)}
              />
              {buResultsOpen && (
                <div className="wt-lookup-results open">
                  {buMatches.length === 0 ? (
                    <div className="wt-lookup-empty">No matches</div>
                  ) : (
                    buMatches.map((b) => (
                      <div key={b.id} className="wt-lookup-item" onMouseDown={() => selectBu(b)}>
                        {b.name}
                      </div>
                    ))
                  )}
                </div>
              )}
            </div>
            {selectedBu && (
              <span className="wt-selected-bu">
                {selectedBu.name}{' '}
                <button type="button" title="Clear" onClick={clearBu}>
                  ×
                </button>
              </span>
            )}
          </div>
          <div className="acc-fld">
            <label>Target Plan Total</label>
            <div className="acc-ro-wrap">
              <div className="acc-ro-label">Sum of cfm_targetplan on AR Aging for this BU</div>
              <div className="acc-ro-value">{selectedBu ? fmtSar(selectedBu.targetPlanAuto) : '—'}</div>
            </div>
          </div>
          <div className="acc-fld">
            <label>
              From <span className="dm-req">*</span>
            </label>
            <input type="date" value={from} onChange={(e) => setFrom(e.target.value)} />
          </div>
          <div className="acc-fld">
            <label>
              To <span className="dm-req">*</span>
            </label>
            <input type="date" value={to} onChange={(e) => setTo(e.target.value)} />
          </div>
          <div className="acc-fld">
            <label>
              # Weeks <span className="dm-req">*</span>
            </label>
            <input type="number" min={1} max={52} value={weekCount} onChange={(e) => setWeekCount(e.target.value)} />
          </div>
        </div>
        <div className="acc-dfoot">
          <button className="acc-btn acc-btn-ghost" onClick={closeCreateDrawer}>
            Cancel
          </button>
          <button className="acc-btn acc-btn-primary" disabled={creating} onClick={handleCreate}>
            {creating ? (
              <>
                <Icon className="fa-solid fa-spinner fa-spin" /> Creating…
              </>
            ) : (
              'Create Plan'
            )}
          </button>
        </div>
      </aside>

      <div className={`acc-toast${toastOn ? ' on' : ''}`}>{toastMsg}</div>
    </>
  );
}
