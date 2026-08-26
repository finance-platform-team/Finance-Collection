import { Fragment, useEffect, useMemo, useState } from 'react';
import { Pm_kpiachievmentsService } from '../generated/services/Pm_kpiachievmentsService';
import { Stf_kpiachievmentbreakdownsService } from '../generated/services/Stf_kpiachievmentbreakdownsService';
import { Cfm_collectionplanweeksService } from '../generated/services/Cfm_collectionplanweeksService';
import { Strategy_kpisesService } from '../generated/services/Strategy_kpisesService';
import { Hr_employeesService } from '../generated/services/Hr_employeesService';
import type { Pm_kpiachievments, Pm_kpiachievmentsBase } from '../generated/models/Pm_kpiachievmentsModel';
import type { Stf_kpiachievmentbreakdowns, Stf_kpiachievmentbreakdownsBase } from '../generated/models/Stf_kpiachievmentbreakdownsModel';
import type { Cfm_collectionplanweeks, Cfm_collectionplanweeksBase } from '../generated/models/Cfm_collectionplanweeksModel';
import type { Hr_employees } from '../generated/models/Hr_employeesModel';
import type { Strategy_kpises } from '../generated/models/Strategy_kpisesModel';
import { operationErrorMessage } from '../lib/errors';
import { fetchAll } from '../lib/dataverse';
import './VisitPlan.css';

/**
 * "Visits — Planned & Actual" — mirrors Ref.html's standalone wizard/dashboard,
 * ported onto this app's generated services + acc-* shared UI. The stf_Employee
 * lookup on the breakdown is bound to hr_employees (Hr_employeesModel/Service are
 * hand-authored — see their module comments — since pac code add-data-source
 * couldn't resolve this table's metadata in this environment).
 */

const BU_ID = '0e650dd0-5d4f-f011-877a-6045bd9979f4'; // AHQ
const DEFAULT_KPI_NAME = 'No. of Visits';
const BREAKDOWN_TYPE = 5;

const MONTHS = ['', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

const ACHIEVEMENT_SELECT = ['pm_kpiachievmentid', 'pm_month', 'pm_year', '_pm_kpi_value'];
const BREAKDOWN_SELECT = ['stf_kpiachievmentbreakdownid', 'stf_value', 'comp_breakdowntarget', '_stf_total_value', '_stf_employee_value'];
const WEEK_SELECT = ['cfm_collectionplanweekid', 'cfm_weeknumber', 'cfm_plan', 'cfm_actual', '_cfm_kpiachievementbreakdown_value'];

function formattedValue(record: unknown, attribute: string): string {
  const raw = (record as Record<string, unknown>)[`${attribute}@OData.Community.Display.V1.FormattedValue`];
  return typeof raw === 'string' ? raw : '';
}
function N(v: number | null | undefined): number {
  return Number(v || 0);
}
function fmtNum(n: number): string {
  return Math.round(n || 0).toLocaleString('en-US');
}
function initials(name: string): string {
  return name
    .split(/\s+/)
    .map((p) => p[0] || '')
    .join('')
    .slice(0, 2)
    .toUpperCase();
}
function pctTone(p: number): 'bad' | 'warn' | 'ok' {
  if (p < 30) return 'bad';
  if (p < 35) return 'warn';
  return 'ok';
}
function achievementPill(p: number) {
  return <span className={`acc-visit-status acc-visit-status-${pctTone(p)}`}>{p.toFixed(1)}%</span>;
}
// kpi-card/kpi-icon variants are named kpi-danger/kpi-icon-danger for the bad tone
// (shared.css), while the acc-visit-status pill family uses "-bad" — same three
// tones, different class-name suffix per shared component.
const KPI_TONE_SUFFIX: Record<'bad' | 'warn' | 'ok', string> = { bad: 'danger', warn: 'warn', ok: 'ok' };
function resolveYear(m: number): number {
  const now = new Date();
  const cm = now.getMonth() + 1;
  let y = now.getFullYear();
  if (m < cm - 6) y += 1;
  else if (m > cm + 6) y -= 1;
  return y;
}
function isoWeekRange(weekIndex: number): { from: string; to: string } {
  const base = new Date();
  const from = new Date(base);
  from.setDate(base.getDate() + (weekIndex - 1) * 7);
  const to = new Date(from);
  to.setDate(from.getDate() + 6);
  return { from: from.toISOString(), to: to.toISOString() };
}
function escapeForCsv(value: string | number): string {
  return `"${String(value).replace(/"/g, '""')}"`;
}
function monthKeyOf(year: number, month: number): string {
  return `${year}-${String(month).padStart(2, '0')}`;
}

const AVATAR_TONES = ['gold', 'ok', 'warn', 'info', 'danger'] as const;

interface WeekRow {
  id: string;
  number: number;
  plan: number;
  actual: number;
}

interface OwnerRow {
  breakdownId: string;
  achievementId: string;
  month: number;
  year: number;
  employeeId: string;
  name: string;
  target: number;
  weeks: WeekRow[];
}

interface AchievementOption {
  id: string;
  kpiName: string;
  month: number;
  year: number;
}

interface EmployeeOption {
  id: string;
  name: string;
}

interface KpiOption {
  id: string;
  name: string;
  code: string;
}

function ownerTotals(o: OwnerRow): { planned: number; achieved: number; rate: number } {
  const planned = o.weeks.reduce((s, w) => s + w.plan, 0);
  const achieved = o.weeks.reduce((s, w) => s + w.actual, 0);
  const rate = planned > 0 ? (achieved / planned) * 100 : 0;
  return { planned, achieved, rate };
}

interface OwnerCard {
  key: string;
  employeeId: string;
  query: string;
  plans: number[];
}

type CellState = 'saving' | 'saved' | 'error';

export default function VisitPlanPage() {
  const [allOwners, setAllOwners] = useState<OwnerRow[]>([]);
  const [achievements, setAchievements] = useState<AchievementOption[]>([]);
  const [employees, setEmployees] = useState<EmployeeOption[]>([]);
  const [kpis, setKpis] = useState<KpiOption[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  // Empty set = "All months". monthKey is `${year}-${month}` so same-named months
  // in different years (e.g. two "August"s a year apart) never collide.
  const [selectedMonths, setSelectedMonths] = useState<Set<string>>(new Set());
  const [monthMenuOpen, setMonthMenuOpen] = useState(false);
  // Empty set = "All owners".
  const [selectedOwners, setSelectedOwners] = useState<Set<string>>(new Set());
  const [ownerMenuOpen, setOwnerMenuOpen] = useState(false);
  const [ownerSearch, setOwnerSearch] = useState('');
  const [expanded, setExpanded] = useState<Set<string>>(new Set());

  const [cellState, setCellState] = useState<Record<string, CellState>>({});
  const [planDrafts, setPlanDrafts] = useState<Record<string, string>>({});

  const [toastMsg, setToastMsg] = useState('');
  const [toastOn, setToastOn] = useState(false);

  // ---- wizard state ----
  const [wizardOpen, setWizardOpen] = useState(false);
  const [wizStep, setWizStep] = useState<1 | 2>(1);
  const [mode, setMode] = useState<'existing' | 'new'>('existing');
  const [selectedAchievementId, setSelectedAchievementId] = useState('');
  const [newMonth, setNewMonth] = useState(String(new Date().getMonth() + 1));
  const [selectedKpiId, setSelectedKpiId] = useState('');
  const [weekCount, setWeekCount] = useState(4);
  const [ownerCards, setOwnerCards] = useState<OwnerCard[]>([]);
  const [employeePickerOpen, setEmployeePickerOpen] = useState<string | null>(null);
  const [wizError, setWizError] = useState('');
  const [wizBusy, setWizBusy] = useState(false);
  const [ownerSeq, setOwnerSeq] = useState(0);

  function showToast(msg: string, durationMs = 3000) {
    setToastMsg(msg);
    setToastOn(true);
    setTimeout(() => setToastOn(false), durationMs);
  }

  async function load(): Promise<OwnerRow[]> {
    setLoading(true);
    setLoadError(null);
    try {
      // pm_kpiachievments/stf_kpiachievmentbreakdowns are shared across the whole
      // KPI/strategy module (breakdown type "Employee" alone is used well beyond
      // this page — e.g. an org-wide headcount snapshot showed up here before this
      // fix). Resolve "No. of Visits"'s KPI id first and scope the achievement
      // query to it, so unrelated KPIs' achievements/breakdowns never load.
      const [employeeRows, kpiRows] = await Promise.all([
        fetchAll<Hr_employees>(
          (skipToken) =>
            Hr_employeesService.getAll({
              select: ['hr_employeeid', 'hr_fullname'],
              filter: 'statecode eq 0',
              orderBy: ['hr_fullname'],
              maxPageSize: 5000,
              skipToken,
            }),
          'Could not load task owners.'
        ),
        fetchAll<Strategy_kpises>(
          (skipToken) =>
            Strategy_kpisesService.getAll({
              select: ['strategy_kpisid', 'strategy_newcolumn', 'cr18c_kpicode'],
              filter: 'statecode eq 0',
              orderBy: ['createdon desc'],
              maxPageSize: 5000,
              skipToken,
            }),
          'Could not load KPIs.'
        ),
      ]);

      const defaultKpi = kpiRows.find((k) => (k.strategy_newcolumn || '').trim().toLowerCase() === DEFAULT_KPI_NAME.toLowerCase());
      if (!defaultKpi) throw new Error(`KPI "${DEFAULT_KPI_NAME}" not found.`);
      const kpiId = defaultKpi.strategy_kpisid;

      const achRows = await fetchAll<Pm_kpiachievments>(
        (skipToken) =>
          Pm_kpiachievmentsService.getAll({
            select: ACHIEVEMENT_SELECT,
            filter: `_pm_kpi_value eq ${kpiId} and statecode eq 0`,
            orderBy: ['createdon desc'],
            maxPageSize: 5000,
            skipToken,
          }),
        'Could not load KPI achievements.'
      );

      // Scope breakdowns to these specific achievements too — stf_kpiachievmentbreakdowns
      // is shared org-wide (this is exactly the table that surfaced the unrelated
      // 851-row headcount data), so fetching it unfiltered doesn't scale and risks
      // this class of bug again even with the achievement-id join guard below.
      const achIdFilter = achRows.map((a) => `_stf_total_value eq ${a.pm_kpiachievmentid}`).join(' or ');
      const [brkRows, weekRows] = await Promise.all([
        achRows.length === 0
          ? Promise.resolve([])
          : fetchAll<Stf_kpiachievmentbreakdowns>(
              (skipToken) =>
                Stf_kpiachievmentbreakdownsService.getAll({
                  select: BREAKDOWN_SELECT,
                  filter: `(${achIdFilter}) and statecode eq 0`,
                  orderBy: ['createdon asc'],
                  maxPageSize: 5000,
                  skipToken,
                }),
              'Could not load KPI achievement breakdowns.'
            ),
        fetchAll<Cfm_collectionplanweeks>(
          (skipToken) =>
            Cfm_collectionplanweeksService.getAll({ select: WEEK_SELECT, filter: 'statecode eq 0', orderBy: ['cfm_weeknumber asc'], maxPageSize: 5000, skipToken }),
          'Could not load collection plan weeks.'
        ),
      ]);

      const employeeOptions: EmployeeOption[] = employeeRows.map((e) => ({ id: e.hr_employeeid, name: e.hr_fullname || '(employee)' }));
      const employeeMap = new Map(employeeOptions.map((e) => [e.id, e.name]));

      const weeksByBrk = new Map<string, WeekRow[]>();
      weekRows.forEach((w) => {
        const brkId = w._cfm_kpiachievementbreakdown_value;
        if (!brkId) return;
        const row: WeekRow = { id: w.cfm_collectionplanweekid, number: N(w.cfm_weeknumber), plan: N(w.cfm_plan), actual: N(w.cfm_actual) };
        const list = weeksByBrk.get(brkId);
        if (list) list.push(row);
        else weeksByBrk.set(brkId, [row]);
      });
      weeksByBrk.forEach((list) => list.sort((a, b) => a.number - b.number));

      const achById = new Map(achRows.map((a) => [a.pm_kpiachievmentid, a]));
      const owners: OwnerRow[] = [];
      brkRows.forEach((b) => {
        const achId = b._stf_total_value;
        const ach = achId ? achById.get(achId) : undefined;
        if (!ach) return;
        const employeeId = b._stf_employee_value || '';
        owners.push({
          breakdownId: b.stf_kpiachievmentbreakdownid,
          achievementId: ach.pm_kpiachievmentid,
          month: N(ach.pm_month),
          year: N(ach.pm_year),
          employeeId,
          name: employeeMap.get(employeeId) || '(owner)',
          target: N(b.comp_breakdowntarget),
          weeks: weeksByBrk.get(b.stf_kpiachievmentbreakdownid) || [],
        });
      });

      setAchievements(
        achRows.map((a) => ({ id: a.pm_kpiachievmentid, kpiName: formattedValue(a, '_pm_kpi_value') || '(KPI)', month: N(a.pm_month), year: N(a.pm_year) }))
      );
      setEmployees(employeeOptions);
      setKpis(kpiRows.map((k) => ({ id: k.strategy_kpisid, name: k.strategy_newcolumn || '(unnamed KPI)', code: k.cr18c_kpicode || '' })));
      setAllOwners(owners);

      // Drop any selected month that no longer exists in the fresh data (e.g. its
      // last task owner was removed) — defaults to "All months" (empty set) otherwise.
      const keys = new Set(owners.map((o) => monthKeyOf(o.year, o.month)));
      setSelectedMonths((prev) => {
        const next = new Set(Array.from(prev).filter((k) => keys.has(k)));
        return next.size === prev.size ? prev : next;
      });
      return owners;
    } catch (e) {
      setLoadError(e instanceof Error ? e.message : String(e));
      return [];
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    load();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    if (!monthMenuOpen) return;
    function onDocMouseDown() {
      setMonthMenuOpen(false);
    }
    document.addEventListener('mousedown', onDocMouseDown);
    return () => document.removeEventListener('mousedown', onDocMouseDown);
  }, [monthMenuOpen]);

  useEffect(() => {
    if (!ownerMenuOpen) return;
    function onDocMouseDown() {
      setOwnerMenuOpen(false);
    }
    document.addEventListener('mousedown', onDocMouseDown);
    return () => document.removeEventListener('mousedown', onDocMouseDown);
  }, [ownerMenuOpen]);

  const monthOptions = useMemo(() => {
    const map = new Map<string, string>();
    allOwners.forEach((o) => {
      const key = monthKeyOf(o.year, o.month);
      if (!map.has(key)) map.set(key, `${MONTHS[o.month] || o.month} ${o.year}`);
    });
    return Array.from(map.entries()).sort((a, b) => b[0].localeCompare(a[0]));
  }, [allOwners]);

  const allMonthsSelected = selectedMonths.size === 0;
  const monthOwners = useMemo(
    () => (allMonthsSelected ? allOwners : allOwners.filter((o) => selectedMonths.has(monthKeyOf(o.year, o.month)))),
    [allOwners, selectedMonths, allMonthsSelected]
  );

  const ownerOptions = useMemo(() => Array.from(new Set(monthOwners.map((o) => o.name))).sort(), [monthOwners]);

  useEffect(() => {
    setSelectedOwners((prev) => {
      const next = new Set(Array.from(prev).filter((n) => ownerOptions.includes(n)));
      return next.size === prev.size ? prev : next;
    });
  }, [ownerOptions]);

  const allOwnersSelected = selectedOwners.size === 0;
  const view = useMemo(
    () => (allOwnersSelected ? monthOwners : monthOwners.filter((o) => selectedOwners.has(o.name))),
    [monthOwners, selectedOwners, allOwnersSelected]
  );

  function selectAllMonths() {
    setSelectedMonths(new Set());
    setSelectedOwners(new Set());
  }
  function toggleMonth(key: string) {
    setSelectedMonths((prev) => {
      const next = new Set(prev);
      if (next.has(key)) next.delete(key);
      else next.add(key);
      return next;
    });
    setSelectedOwners(new Set());
  }
  function selectAllOwners() {
    setSelectedOwners(new Set());
  }
  function toggleOwner(name: string) {
    setSelectedOwners((prev) => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name);
      else next.add(name);
      return next;
    });
  }
  const ownerButtonLabel = allOwnersSelected
    ? 'All owners'
    : selectedOwners.size === 1
      ? Array.from(selectedOwners)[0]
      : `${selectedOwners.size} owners selected`;
  const ownerMenuMatches = useMemo(() => {
    const q = ownerSearch.trim().toLowerCase();
    return q ? ownerOptions.filter((n) => n.toLowerCase().includes(q)) : ownerOptions;
  }, [ownerOptions, ownerSearch]);
  const monthButtonLabel = allMonthsSelected
    ? 'All months'
    : selectedMonths.size === 1
      ? monthOptions.find(([k]) => selectedMonths.has(k))?.[1] || '1 month'
      : `${selectedMonths.size} months selected`;

  const summary = useMemo(() => {
    const planned = view.reduce((s, o) => s + ownerTotals(o).planned, 0);
    const achieved = view.reduce((s, o) => s + ownerTotals(o).achieved, 0);
    const rate = planned > 0 ? (achieved / planned) * 100 : 0;
    return { planned, achieved, rate };
  }, [view]);

  function toggleExpand(id: string) {
    setExpanded((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function planValue(w: WeekRow): string {
    return w.id in planDrafts ? planDrafts[w.id] : String(w.plan);
  }

  async function commitPlan(owner: OwnerRow, week: WeekRow) {
    const draft = planDrafts[week.id];
    if (draft === undefined) return;
    const value = Math.max(0, Number(draft) || 0);
    if (value === week.plan) {
      setPlanDrafts((prev) => {
        const next = { ...prev };
        delete next[week.id];
        return next;
      });
      return;
    }
    const key = week.id;
    setCellState((prev) => ({ ...prev, [key]: 'saving' }));
    try {
      const weekResult = await Cfm_collectionplanweeksService.update(week.id, { cfm_plan: value });
      if (!weekResult.success) throw new Error(operationErrorMessage(weekResult.error, 'Could not save plan.'));

      // Sum from a fresh server read of this breakdown's weeks rather than the
      // client-side owner.weeks snapshot: saving several weeks back-to-back (faster
      // than each PATCH round-trips) means each save's local snapshot doesn't yet
      // know about the others still in flight, so a client-computed total can
      // undercount. A fresh read always reflects every commit that has actually landed.
      const freshWeeks = await fetchAll<Cfm_collectionplanweeks>(
        (skipToken) =>
          Cfm_collectionplanweeksService.getAll({
            select: ['cfm_plan'],
            filter: `_cfm_kpiachievementbreakdown_value eq ${owner.breakdownId} and statecode eq 0`,
            maxPageSize: 5000,
            skipToken,
          }),
        'Could not verify week totals.'
      );
      const totalPlan = freshWeeks.reduce((s, w) => s + N(w.cfm_plan), 0);
      const brkResult = await Stf_kpiachievmentbreakdownsService.update(owner.breakdownId, { comp_breakdowntarget: totalPlan });
      if (!brkResult.success) throw new Error(operationErrorMessage(brkResult.error, 'Could not update breakdown target.'));

      setAllOwners((prev) =>
        prev.map((o) => (o.breakdownId !== owner.breakdownId ? o : { ...o, target: totalPlan, weeks: o.weeks.map((w) => (w.id === week.id ? { ...w, plan: value } : w)) }))
      );
      setPlanDrafts((prev) => {
        const next = { ...prev };
        delete next[week.id];
        return next;
      });
      setCellState((prev) => ({ ...prev, [key]: 'saved' }));
      setTimeout(() => setCellState((prev) => { const n = { ...prev }; delete n[key]; return n; }), 900);
    } catch (e) {
      setCellState((prev) => ({ ...prev, [key]: 'error' }));
      setTimeout(() => setCellState((prev) => { const n = { ...prev }; delete n[key]; return n; }), 1400);
      showToast(e instanceof Error ? e.message : 'Save failed.', 4500);
    }
  }

  async function deleteOwner(owner: OwnerRow) {
    if (!window.confirm(`Remove "${owner.name}"? This deletes their plan weeks and cannot be undone.`)) return;
    try {
      await Promise.all(owner.weeks.map((w) => Cfm_collectionplanweeksService.delete(w.id)));
      await Stf_kpiachievmentbreakdownsService.delete(owner.breakdownId);
      setAllOwners((prev) => prev.filter((o) => o.breakdownId !== owner.breakdownId));
      showToast(`Removed ${owner.name}.`);
    } catch (e) {
      showToast(e instanceof Error ? e.message : 'Delete failed.', 4500);
    }
  }

  function exportCsv() {
    if (!view.length) {
      showToast('Nothing to export.');
      return;
    }
    const header = ['Task owner', 'Monthly planned', 'Monthly achieved', 'Achievement %', 'Month'];
    const lines = [header.join(',')];
    view.forEach((o) => {
      const t = ownerTotals(o);
      lines.push([o.name, Math.round(t.planned), Math.round(t.achieved), t.rate.toFixed(1) + '%', `${MONTHS[o.month] || o.month} ${o.year}`].map(escapeForCsv).join(','));
    });
    const blob = new Blob([lines.join('\n')], { type: 'text/csv;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'visits_planned_actual.csv';
    link.click();
    URL.revokeObjectURL(url);
  }

  // ================= wizard =================
  function openWizard() {
    setWizError('');
    setWizStep(1);
    setMode('existing');
    setNewMonth(String(new Date().getMonth() + 1));
    const defaultKpi = kpis.find((k) => k.name.trim().toLowerCase() === DEFAULT_KPI_NAME.toLowerCase());
    setSelectedKpiId(defaultKpi?.id || kpis[0]?.id || '');
    const defaultAch =
      achievements.find((a) => a.kpiName.trim().toLowerCase() === DEFAULT_KPI_NAME.toLowerCase()) || achievements[0];
    setSelectedAchievementId(defaultAch?.id || '');
    setWeekCount(4);
    setOwnerCards([]);
    setOwnerSeq(0);
    setWizardOpen(true);
  }
  function closeWizard() {
    setWizardOpen(false);
  }

  function addOwnerCard() {
    const key = `owner-${ownerSeq}`;
    setOwnerSeq((n) => n + 1);
    setOwnerCards((prev) => [...prev, { key, employeeId: '', query: '', plans: Array.from({ length: weekCount }, () => 1) }]);
  }
  function removeOwnerCard(key: string) {
    setOwnerCards((prev) => prev.filter((c) => c.key !== key));
    setEmployeePickerOpen((open) => (open === key ? null : open));
  }
  function setCardQuery(key: string, query: string) {
    setOwnerCards((prev) => prev.map((c) => (c.key === key ? { ...c, query, employeeId: '' } : c)));
  }
  function pickCardEmployee(key: string, employee: EmployeeOption) {
    setOwnerCards((prev) => prev.map((c) => (c.key === key ? { ...c, employeeId: employee.id, query: employee.name } : c)));
    setEmployeePickerOpen(null);
  }
  function employeeMatches(query: string): EmployeeOption[] {
    const q = query.trim().toLowerCase();
    const list = q ? employees.filter((e) => e.name.toLowerCase().includes(q)) : employees;
    return list.slice(0, 30);
  }
  function setCardPlan(key: string, weekIdx: number, value: number) {
    setOwnerCards((prev) => prev.map((c) => (c.key === key ? { ...c, plans: c.plans.map((p, i) => (i === weekIdx ? value : p)) } : c)));
  }
  function changeWeeks(delta: number) {
    setWeekCount((prev) => {
      const next = Math.max(1, Math.min(52, prev + delta));
      setOwnerCards((cards) =>
        cards.map((c) => {
          const plans = c.plans.slice(0, next);
          while (plans.length < next) plans.push(1);
          return { ...c, plans };
        })
      );
      return next;
    });
  }
  const totalPlanned = useMemo(() => ownerCards.reduce((s, c) => s + c.plans.reduce((a, b) => a + (Number(b) || 0), 0), 0), [ownerCards]);

  async function onContinue() {
    setWizError('');
    setWizBusy(true);
    try {
      let achievementId: string;
      if (mode === 'existing') {
        if (!selectedAchievementId) throw new Error('Select a KPI achievement first.');
        achievementId = selectedAchievementId;
      } else {
        if (!selectedKpiId) throw new Error('No KPI selected.');
        const month = parseInt(newMonth, 10);
        const year = resolveYear(month);
        const body: Record<string, unknown> = {
          pm_month: month,
          pm_year: year,
          'pm_kpi@odata.bind': `/strategy_kpises(${selectedKpiId})`,
          'pm_businessunit@odata.bind': `/businessunits(${BU_ID})`,
        };
        const result = await Pm_kpiachievmentsService.create(body as unknown as Omit<Pm_kpiachievmentsBase, 'pm_kpiachievmentid'>);
        if (!result.success) throw new Error(operationErrorMessage(result.error, 'Could not create the KPI achievement.'));
        const id = (result.data as unknown as { pm_kpiachievmentid?: string })?.pm_kpiachievmentid;
        if (!id) throw new Error('KPI achievement was created but no id was returned.');
        achievementId = id;
      }
      setSelectedAchievementId(achievementId);
      setWizStep(2);
      if (ownerCards.length === 0) addOwnerCard();
    } catch (e) {
      setWizError(e instanceof Error ? e.message : 'Could not continue.');
    } finally {
      setWizBusy(false);
    }
  }
  function backToStep1() {
    setWizStep(1);
    setWizError('');
  }

  async function onCreateAll() {
    setWizError('');
    if (!selectedAchievementId) {
      setWizError('No KPI achievement selected.');
      return;
    }
    if (!ownerCards.length) {
      setWizError('Add at least one task owner.');
      return;
    }
    for (const card of ownerCards) {
      if (!card.employeeId) {
        setWizError('Select a task owner in every row.');
        return;
      }
    }

    setWizBusy(true);
    try {
      for (const card of ownerCards) {
        const target = card.plans.reduce((a, b) => a + (Number(b) || 0), 0);
        const brkBody: Record<string, unknown> = {
          stf_breakdowntype: BREAKDOWN_TYPE,
          comp_breakdowntarget: target,
          stf_value: 0,
          'stf_Total@odata.bind': `/pm_kpiachievments(${selectedAchievementId})`,
          'stf_Employee@odata.bind': `/hr_employees(${card.employeeId})`,
        };
        const brkResult = await Stf_kpiachievmentbreakdownsService.create(brkBody as unknown as Omit<Stf_kpiachievmentbreakdownsBase, 'stf_kpiachievmentbreakdownid'>);
        if (!brkResult.success) throw new Error(operationErrorMessage(brkResult.error, 'Could not create the task owner breakdown.'));
        const breakdownId = (brkResult.data as unknown as { stf_kpiachievmentbreakdownid?: string })?.stf_kpiachievmentbreakdownid;
        if (!breakdownId) throw new Error('Breakdown was created but no id was returned.');

        const weekResults = await Promise.all(
          card.plans.map((plan, i) => {
            const w = i + 1;
            const range = isoWeekRange(w);
            const weekBody: Record<string, unknown> = {
              cfm_weeknumber: w,
              cfm_plan: Number(plan) || 0,
              cfm_actual: null,
              cfm_weekfrom: range.from,
              cfm_weekto: range.to,
              'cfm_KPIAchievementBreakdown@odata.bind': `/stf_kpiachievmentbreakdowns(${breakdownId})`,
            };
            return Cfm_collectionplanweeksService.create(weekBody as unknown as Omit<Cfm_collectionplanweeksBase, 'cfm_collectionplanweekid'>);
          })
        );
        const failed = weekResults.filter((r) => !r.success).length;
        if (failed) throw new Error(`${failed} of ${card.plans.length} week(s) failed to save.`);
      }

      const createdAchievementId = selectedAchievementId;
      setWizardOpen(false);
      showToast('KPI achievement plan created.');
      const freshOwners = await load();
      const created = freshOwners.find((o) => o.achievementId === createdAchievementId);
      if (created) setSelectedMonths(new Set([monthKeyOf(created.year, created.month)]));
    } catch (e) {
      setWizError(e instanceof Error ? e.message : 'Create failed.');
    } finally {
      setWizBusy(false);
    }
  }

  const wizKpiChip = mode === 'existing' ? achievements.find((a) => a.id === selectedAchievementId)?.kpiName || 'KPI' : kpis.find((k) => k.id === selectedKpiId)?.name || 'KPI';
  const wizMonthChip = mode === 'existing' ? (() => { const a = achievements.find((x) => x.id === selectedAchievementId); return a ? `${MONTHS[a.month]} ${a.year}` : ''; })() : MONTHS[parseInt(newMonth, 10)];

  return (
    <>
      <div className="acc-wrap">
        <div className="acc-head">
          <div>
            <div className="vp-eyebrow">
              <span className="vp-dot"></span> KPI ACHIEVEMENT{!allMonthsSelected ? ` · ${monthButtonLabel.toUpperCase()}` : ''}
            </div>
            <h1 className="acc-title">Visits — Planned &amp; Actual</h1>
            <div className="acc-sub">Task owner · Monthly planned / achieved · Weekly planned / achieved · Achievement %</div>
          </div>
          <div className="acc-actions">
            <button className="acc-btn acc-btn-primary" onClick={openWizard}>
              <i className="fa-solid fa-plus"></i> Create KPI achievement
            </button>
            <button className="acc-btn" onClick={exportCsv}>
              <i className="fa-solid fa-download"></i> Export
            </button>
          </div>
        </div>

        <div className="kpi-grid">
          <div className="kpi-card">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-gold">
                <i className="fa-solid fa-bullseye"></i>
              </span>
              <span className="kpi-label">Monthly planned</span>
            </div>
            <div className="kpi-value">{loading ? '—' : fmtNum(summary.planned)}</div>
            <div className="kpi-context">across {view.length} task owner{view.length === 1 ? '' : 's'}</div>
          </div>
          <div className="kpi-card kpi-info">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-info">
                <i className="fa-solid fa-clipboard-check"></i>
              </span>
              <span className="kpi-label">Monthly achieved</span>
            </div>
            <div className="kpi-value">{loading ? '—' : fmtNum(summary.achieved)}</div>
            <div className="kpi-context">recorded so far</div>
          </div>
          <div className={`kpi-card kpi-${KPI_TONE_SUFFIX[pctTone(summary.rate)]}`}>
            <div className="kpi-top">
              <span className={`kpi-icon kpi-icon-${KPI_TONE_SUFFIX[pctTone(summary.rate)]}`}>
                <i className="fa-solid fa-gauge-high"></i>
              </span>
              <span className="kpi-label">Achievement rate</span>
            </div>
            <div className="kpi-value">{loading ? '—' : `${summary.rate.toFixed(1)}%`}</div>
            <div className="vp-bar">
              <span style={{ width: `${Math.min(100, summary.rate)}%` }}></span>
            </div>
          </div>
        </div>

        <div className="tbl-shell">
          <div className="tbl-hdr">
            <div className="dm-filter-bar">
              <div className="vp-multiselect" onMouseDown={(e) => e.stopPropagation()}>
                <button type="button" className="acc-f vp-multiselect-btn" onClick={() => setMonthMenuOpen((v) => !v)}>
                  <i className="fa-solid fa-calendar-days vp-multiselect-ico"></i>
                  <span>{monthButtonLabel}</span>
                  <i className={`fa-solid fa-chevron-down vp-multiselect-caret${monthMenuOpen ? ' open' : ''}`}></i>
                </button>
                {monthMenuOpen && (
                  <div className="vp-multiselect-menu">
                    <label className={`vp-multiselect-item vp-multiselect-all${allMonthsSelected ? ' checked' : ''}`}>
                      <input type="checkbox" checked={allMonthsSelected} onChange={selectAllMonths} />
                      All months
                    </label>
                    <div className="vp-multiselect-sep"></div>
                    {monthOptions.length === 0 ? (
                      <div className="vp-multiselect-empty">No data</div>
                    ) : (
                      <div className="vp-multiselect-scroll">
                        {monthOptions.map(([key, label]) => (
                          <label className="vp-multiselect-item" key={key}>
                            <input type="checkbox" checked={selectedMonths.has(key)} onChange={() => toggleMonth(key)} />
                            {label}
                          </label>
                        ))}
                      </div>
                    )}
                  </div>
                )}
              </div>

              <div className="vp-multiselect" onMouseDown={(e) => e.stopPropagation()}>
                <button
                  type="button"
                  className="acc-f vp-multiselect-btn"
                  onClick={() => {
                    setOwnerMenuOpen((v) => !v);
                    setOwnerSearch('');
                  }}
                >
                  <i className="fa-solid fa-user-tag vp-multiselect-ico"></i>
                  <span>{ownerButtonLabel}</span>
                  <i className={`fa-solid fa-chevron-down vp-multiselect-caret${ownerMenuOpen ? ' open' : ''}`}></i>
                </button>
                {ownerMenuOpen && (
                  <div className="vp-multiselect-menu vp-multiselect-menu-wide">
                    <div className="vp-multiselect-search">
                      <i className="fa-solid fa-magnifying-glass"></i>
                      <input
                        type="text"
                        placeholder="Search task owners…"
                        autoFocus
                        value={ownerSearch}
                        onChange={(e) => setOwnerSearch(e.target.value)}
                      />
                    </div>
                    <label className={`vp-multiselect-item vp-multiselect-all${allOwnersSelected ? ' checked' : ''}`}>
                      <input type="checkbox" checked={allOwnersSelected} onChange={selectAllOwners} />
                      All owners
                    </label>
                    <div className="vp-multiselect-sep"></div>
                    <div className="vp-multiselect-scroll">
                      {ownerMenuMatches.length === 0 ? (
                        <div className="vp-multiselect-empty">No matches.</div>
                      ) : (
                        ownerMenuMatches.map((n) => (
                          <label className="vp-multiselect-item" key={n}>
                            <input type="checkbox" checked={selectedOwners.has(n)} onChange={() => toggleOwner(n)} />
                            {n}
                          </label>
                        ))
                      )}
                    </div>
                  </div>
                )}
              </div>

              {(!allMonthsSelected || !allOwnersSelected) && (
                <button
                  type="button"
                  className="vp-clear-filters"
                  onClick={() => {
                    setSelectedMonths(new Set());
                    setSelectedOwners(new Set());
                  }}
                >
                  <i className="fa-solid fa-xmark"></i> Clear filters
                </button>
              )}
            </div>
            <div className="dm-count">
              <i className="fa-solid fa-users"></i> Task owners <strong>{view.length}</strong>
            </div>
          </div>
          <div className="acc-tablewrap">
            <table>
              <thead>
                <tr>
                  <th>Task owner</th>
                  <th className="acc-num">Monthly planned</th>
                  <th className="acc-num">Monthly achieved</th>
                  <th className="acc-num">Achievement %</th>
                  <th className="acc-num">Month</th>
                  <th></th>
                </tr>
              </thead>
              <tbody>
                {loading ? (
                  Array.from({ length: 4 }).map((_, i) => (
                    <tr key={i}>
                      {Array.from({ length: 6 }).map((__, j) => (
                        <td key={j}>
                          <div className="acc-skel"></div>
                        </td>
                      ))}
                    </tr>
                  ))
                ) : loadError ? (
                  <tr>
                    <td colSpan={6}>
                      <div className="acc-state">
                        <i className="fa-solid fa-triangle-exclamation"></i>
                        Couldn't load visit plans.
                        <br />
                        <small>{loadError}</small>
                      </div>
                    </td>
                  </tr>
                ) : view.length === 0 ? (
                  <tr>
                    <td colSpan={6}>
                      <div className="acc-state">
                        <i className="fa-solid fa-inbox"></i>
                        No task owners for this selection.
                      </div>
                    </td>
                  </tr>
                ) : (
                  view.map((o, idx) => {
                    const t = ownerTotals(o);
                    const isOpen = expanded.has(o.breakdownId);
                    const tone = AVATAR_TONES[idx % AVATAR_TONES.length];
                    return (
                      <Fragment key={o.breakdownId}>
                        <tr className="acc-row" onClick={() => toggleExpand(o.breakdownId)}>
                          <td>
                            <div className="vp-who">
                              <i className={`fa-solid ${isOpen ? 'fa-chevron-down' : 'fa-chevron-right'} vp-caret`}></i>
                              <div className={`vp-ava vp-ava-${tone}`}>{initials(o.name)}</div>
                              <div>
                                <div className="acc-name">{o.name}</div>
                                <div className="vp-tags">
                                  <span className="vp-chip">
                                    {o.weeks.length} week{o.weeks.length === 1 ? '' : 's'}
                                  </span>
                                </div>
                              </div>
                            </div>
                          </td>
                          <td className="acc-num vp-planned">{fmtNum(t.planned)}</td>
                          <td className="acc-num vp-achieved">{fmtNum(t.achieved)}</td>
                          <td className="acc-num">{achievementPill(t.rate)}</td>
                          <td className="acc-num">
                            <span className="vp-chip">{MONTHS[o.month]} {o.year}</span>
                          </td>
                          <td className="acc-num" onClick={(e) => e.stopPropagation()}>
                            <button className="vp-del" title="Remove" onClick={() => deleteOwner(o)}>
                              <i className="fa-solid fa-trash"></i>
                            </button>
                          </td>
                        </tr>
                        {isOpen && (
                          <tr className="vp-weeks-row">
                            <td colSpan={6}>
                              <div className="vp-weeks-grid">
                                {o.weeks.length === 0 ? (
                                  <div className="vp-no-weeks">No weeks configured.</div>
                                ) : (
                                  o.weeks.map((w) => {
                                    const key = w.id;
                                    const draftPlan = Number(planValue(w)) || 0;
                                    const prog = draftPlan > 0 ? Math.min(100, (w.actual / draftPlan) * 100) : 0;
                                    return (
                                      <div className="vp-wk" key={w.id}>
                                        <div className="vp-wtitle">Week {w.number}</div>
                                        <div className="vp-pair">
                                          <div>
                                            <small>Plan</small>
                                            <input
                                              type="number"
                                              min={0}
                                              value={planValue(w)}
                                              onChange={(e) => setPlanDrafts((prev) => ({ ...prev, [w.id]: e.target.value }))}
                                              onBlur={() => commitPlan(o, w)}
                                            />
                                          </div>
                                          <div>
                                            <small>Actual</small>
                                            <input type="text" value={w.actual} readOnly />
                                          </div>
                                        </div>
                                        <div className="vp-wbar">
                                          <span style={{ width: `${prog}%` }}></span>
                                        </div>
                                        {cellState[key] === 'saving' && <i className="fa-solid fa-spinner fa-spin vp-cell-ico"></i>}
                                        {cellState[key] === 'saved' && <i className="fa-solid fa-check vp-cell-ico vp-ok"></i>}
                                        {cellState[key] === 'error' && <i className="fa-solid fa-triangle-exclamation vp-cell-ico vp-bad"></i>}
                                      </div>
                                    );
                                  })
                                )}
                              </div>
                            </td>
                          </tr>
                        )}
                      </Fragment>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>
        </div>

        <div className="tbl-shell vp-chart-panel">
          <h2 className="vp-chh">
            <i className="fa-solid fa-chart-simple"></i> Actual vs target by task owner
          </h2>
          <div className="vp-chart-body">
            {view.length === 0 ? (
              <div className="acc-state">
                <i className="fa-solid fa-chart-bar"></i>
                Nothing to chart yet.
              </div>
            ) : (
              view.map((o) => {
                const t = ownerTotals(o);
                const tone = pctTone(t.rate);
                return (
                  <div className="vp-crow" key={o.breakdownId}>
                    <div className="vp-cname">{o.name}</div>
                    <div className="vp-ctrack">
                      <span className={`vp-ctrack-fill vp-ctrack-${tone}`} style={{ width: `${Math.min(100, t.rate)}%` }}></span>
                    </div>
                    <div className="vp-cval">
                      {fmtNum(t.achieved)} / {fmtNum(t.planned)}
                    </div>
                    <div className={`vp-cpct vp-cpct-${tone}`}>{t.rate.toFixed(1)}%</div>
                  </div>
                );
              })
            )}
          </div>
        </div>
      </div>

      {/* ===== Create wizard modal ===== */}
      <div className={`vp-modal-overlay${wizardOpen ? ' on' : ''}`} onClick={closeWizard}>
        <div className="vp-modal" onClick={(e) => e.stopPropagation()}>
          <button className="vp-modal-close" onClick={closeWizard}>
            <i className="fa-solid fa-xmark"></i>
          </button>
          <div className="vp-weyebrow">New objective</div>
          <h2 className="vp-wtitle">Create KPI achievement</h2>

          <div className="vp-steps">
            <div className={`vp-step${wizStep === 1 ? ' active' : ' done'}`}>
              <div className="vp-scircle">{wizStep === 1 ? '1' : <i className="fa-solid fa-check"></i>}</div>
              <div className="vp-stext">
                <b>Define KPI</b>
                <span>Name &amp; month</span>
              </div>
            </div>
            <div className="vp-sline"></div>
            <div className={`vp-step${wizStep === 2 ? ' active' : ''}`}>
              <div className="vp-scircle">2</div>
              <div className="vp-stext">
                <b>Add visit plan</b>
                <span>Owners &amp; weeks</span>
              </div>
            </div>
          </div>
          <div className="vp-wdivider"></div>

          <div className="vp-wbody">
            <div className={`acc-err${wizError ? ' on' : ''}`}>{wizError}</div>

            {wizStep === 1 ? (
              <>
                <div className="vp-toggle">
                  <button className={mode === 'existing' ? 'on' : ''} onClick={() => setMode('existing')}>
                    Use existing
                  </button>
                  <button className={mode === 'new' ? 'on' : ''} onClick={() => setMode('new')}>
                    Create new
                  </button>
                </div>

                {mode === 'existing' ? (
                  <div className="vp-field">
                    <label>KPI achievement</label>
                    <select disabled value={selectedAchievementId} onChange={() => {}}>
                      {achievements.length === 0 ? (
                        <option value="">No achievements yet</option>
                      ) : (
                        achievements.map((a) => (
                          <option key={a.id} value={a.id}>
                            {a.kpiName} — {MONTHS[a.month]} {a.year}
                          </option>
                        ))
                      )}
                    </select>
                    <p className="vp-hint">Auto-selected: latest "{DEFAULT_KPI_NAME}" achievement. Weeks and breakdowns link to this one.</p>
                  </div>
                ) : (
                  <>
                    <div className="vp-field">
                      <label>
                        KPI <span className="vp-badge-ro">read only</span>
                      </label>
                      <select disabled value={selectedKpiId} onChange={() => {}}>
                        {kpis.length === 0 ? (
                          <option value="">Loading KPIs…</option>
                        ) : (
                          kpis.map((k) => (
                            <option key={k.id} value={k.id}>
                              {k.name}
                              {k.code ? ` — ${k.code}` : ''}
                            </option>
                          ))
                        )}
                      </select>
                      <p className="vp-hint">Defaults to "{DEFAULT_KPI_NAME}". Loaded live from the KPI table.</p>
                    </div>
                    <div className="vp-row2">
                      <div>
                        <label>Month</label>
                        <select value={newMonth} onChange={(e) => setNewMonth(e.target.value)}>
                          {MONTHS.slice(1).map((m, i) => (
                            <option key={i + 1} value={i + 1}>
                              {m}
                            </option>
                          ))}
                        </select>
                        <p className="vp-hint">
                          Year derived: <b>{resolveYear(parseInt(newMonth, 10))}</b>
                        </p>
                      </div>
                      <div>
                        <label>
                          Business unit <span className="vp-badge-ro">read only</span>
                        </label>
                        <select disabled value={BU_ID} onChange={() => {}}>
                          <option value={BU_ID}>AHQ</option>
                        </select>
                      </div>
                    </div>
                  </>
                )}
              </>
            ) : (
              <>
                <div className="vp-wk-head">
                  <div>
                    <label>Number of weeks</label>
                    <div className="vp-weeks-ctl">
                      <button onClick={() => changeWeeks(-1)}>−</button>
                      <div className="vp-weeks-val">{weekCount}</div>
                      <button onClick={() => changeWeeks(1)}>+</button>
                    </div>
                  </div>
                  <div className="vp-total-planned">
                    <div className="vp-lbl">Total planned visits</div>
                    <div className="vp-num">{fmtNum(totalPlanned)}</div>
                  </div>
                </div>

                {ownerCards.map((card) => (
                  <div className="vp-owner-card" key={card.key}>
                    <button className="vp-del" title="Remove" onClick={() => removeOwnerCard(card.key)}>
                      <i className="fa-solid fa-trash"></i>
                    </button>
                    <label>Task owner</label>
                    <div className="acc-combo vp-owner-combo">
                      <input
                        type="text"
                        placeholder="Search task owner…"
                        autoComplete="off"
                        value={card.query}
                        onChange={(e) => setCardQuery(card.key, e.target.value)}
                        onFocus={() => setEmployeePickerOpen(card.key)}
                        onBlur={() => setTimeout(() => setEmployeePickerOpen((open) => (open === card.key ? null : open)), 150)}
                      />
                      {employeePickerOpen === card.key && (
                        <div className="acc-combo-list">
                          {employeeMatches(card.query).length === 0 ? (
                            <div className="acc-combo-empty">No matches.</div>
                          ) : (
                            employeeMatches(card.query).map((e) => (
                              <div key={e.id} className="acc-combo-item" onMouseDown={(ev) => ev.preventDefault()} onClick={() => pickCardEmployee(card.key, e)}>
                                {e.name}
                              </div>
                            ))
                          )}
                        </div>
                      )}
                    </div>
                    <div className="vp-wk-inputs">
                      {card.plans.map((p, wi) => (
                        <div className="vp-wk2" key={wi}>
                          <small>Week {wi + 1}</small>
                          <input type="number" min={0} value={p} onChange={(e) => setCardPlan(card.key, wi, Number(e.target.value) || 0)} />
                        </div>
                      ))}
                      <div className="vp-rowtotal">
                        <div className="vp-lbl">Row total</div>
                        <div className="vp-num">{fmtNum(card.plans.reduce((a, b) => a + (Number(b) || 0), 0))}</div>
                      </div>
                    </div>
                  </div>
                ))}
                <button className="vp-add-owner-btn" onClick={addOwnerCard}>
                  ＋ Add task owner
                </button>
              </>
            )}
          </div>

          <div className="vp-wfooter">
            <div className="vp-wfoot-chips">
              <span className="vp-wchip">{wizKpiChip}</span>
              <span className="vp-wchip">{wizMonthChip}</span>
            </div>
            <div className="vp-wfoot-actions">
              <button className="vp-wbtn-link" onClick={closeWizard}>
                Cancel
              </button>
              {wizStep === 2 && (
                <button className="vp-wbtn-link" onClick={backToStep1}>
                  Back
                </button>
              )}
              {wizStep === 1 ? (
                <button className="vp-wbtn-primary" disabled={wizBusy} onClick={onContinue}>
                  {wizBusy ? 'Saving…' : <>Continue ›</>}
                </button>
              ) : (
                <button className="vp-wbtn-primary" disabled={wizBusy} onClick={onCreateAll}>
                  {wizBusy ? 'Creating…' : '✓ Create KPI'}
                </button>
              )}
            </div>
          </div>
        </div>
      </div>

      <div className={`acc-toast${toastOn ? ' on' : ''}`}>{toastMsg}</div>
    </>
  );
}
