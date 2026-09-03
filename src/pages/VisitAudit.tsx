import { useEffect, useMemo, useState } from 'react';
import { Cr301_visitsService } from '../generated/services/Cr301_visitsService';
import { Cr301_agentsService } from '../generated/services/Cr301_agentsService';
import { Hr_employeesService } from '../generated/services/Hr_employeesService';
import { Strategy_kpisesService } from '../generated/services/Strategy_kpisesService';
import { Pm_kpiachievmentsService } from '../generated/services/Pm_kpiachievmentsService';
import { Stf_kpiachievmentbreakdownsService } from '../generated/services/Stf_kpiachievmentbreakdownsService';
import { Cfm_collectionplanweeksService } from '../generated/services/Cfm_collectionplanweeksService';
import type { Cr301_visits } from '../generated/models/Cr301_visitsModel';
import type { Hr_employees } from '../generated/models/Hr_employeesModel';
import type { Pm_kpiachievments } from '../generated/models/Pm_kpiachievmentsModel';
import type { Stf_kpiachievmentbreakdowns } from '../generated/models/Stf_kpiachievmentbreakdownsModel';
import type { Cfm_collectionplanweeks } from '../generated/models/Cfm_collectionplanweeksModel';
import { fetchAll } from '../lib/dataverse';
import { escapeForCsv, downloadCsv } from '../lib/csvExport';
import type { Region } from '../lib/useRegion';
import Icon from '../components/Icon';
import './VisitAudit.css';

/**
 * Cross-checks every recorded visit (cr301_visits) against its Visit Plan week
 * (cfm_collectionplanweeks, via stf_kpiachievmentbreakdowns / pm_kpiachievments
 * for the "No. of Visits" KPI — same chain VisitPlan.tsx and Visits.tsx's
 * adjustVisitPlanActual use) to flag visits that were logged late or landed in
 * the wrong week bucket, and rolls that up per task owner.
 *
 * cfm_agent2 (Task Owner 2, double visits) isn't in the generated Cr301_visits
 * model — see the same gap documented in Visits.tsx. Selected here as a raw
 * string key for the same reason: it's a real, confirmed attribute name.
 */
const VISIT_KPI_NAME = 'No. of Visits';
const TAG_COLLECTION = 766340000; // cfm_tagename — "Collection"
// Same cr18c_area choice on cr301_agents / same Egypt-vs-KSA mapping as
// Agents.tsx, VisitPlan.tsx and Visits.tsx (KSA = the real KSA choice value,
// Egypt = everything else, including unset).
const AREA_KSA = 989230004;

const MONTHS = ['', 'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];

// Same fixed calendar-day buckets as VisitPlan.tsx's MONTH_WEEK_BOUNDS and
// Visits.tsx's monthDayBucket — must stay in sync with both.
function monthDayBucket(day: number): number {
  if (day <= 10) return 1;
  if (day <= 18) return 2;
  if (day <= 25) return 3;
  return 4;
}

// Same tolerant name match Agents.tsx / Visits.tsx use for Agent <-> Employee.
function firstWords(name: string, n: number): string {
  return name.trim().toLowerCase().split(/\s+/).slice(0, n).join(' ');
}

function N(v: number | string | null | undefined): number {
  return Number(v || 0);
}

function daysBetween(a: Date, b: Date): number {
  const ua = Date.UTC(a.getFullYear(), a.getMonth(), a.getDate());
  const ub = Date.UTC(b.getFullYear(), b.getMonth(), b.getDate());
  return Math.round((ub - ua) / 86400000);
}

function monthKeyOf(year: number, month: number): string {
  return `${year}-${String(month).padStart(2, '0')}`;
}

interface UnmappedVisitFields {
  _cfm_agent2_value?: string;
}

type AuditStatus = 'on-time' | 'late-same-week' | 'wrong-week' | 'wrong-month';

const STATUS_LABEL: Record<AuditStatus, string> = {
  'on-time': 'On-Time',
  'late-same-week': 'Late – Same Week',
  'wrong-week': 'Wrong Week',
  'wrong-month': 'Wrong Month',
};
const STATUS_TONE: Record<AuditStatus, 'ok' | 'warn' | 'bad'> = {
  'on-time': 'ok',
  'late-same-week': 'warn',
  'wrong-week': 'bad',
  'wrong-month': 'bad',
};

interface AuditRow {
  key: string;
  visitId: string;
  ownerName: string;
  employeeId: string;
  visitDate: string;
  createdOn: string;
  expectedWeek: number;
  expectedMonth: number;
  expectedYear: number;
  createdWeek: number;
  createdMonth: number;
  createdYear: number;
  daysDiff: number;
  plan: number | null;
  actual: number | null;
  variance: number | null;
  status: AuditStatus;
  region: 'EG' | 'KSA' | '';
}

interface OwnerSummary {
  name: string;
  total: number;
  correct: number;
  late: number;
  wrongWeek: number;
  planned: number;
  actual: number;
  variance: number;
  timingCompliance: number;
  achievementPct: number;
}

function fmtDate(iso: string): string {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '—';
  return d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
}
function fmtDateTime(iso: string): string {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '—';
  return d.toLocaleString(undefined, { year: 'numeric', month: 'short', day: 'numeric', hour: '2-digit', minute: '2-digit' });
}

interface VisitAuditPageProps {
  region: Region | null;
}

export default function VisitAuditPage({ region }: VisitAuditPageProps) {
  const [rows, setRows] = useState<AuditRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [query, setQuery] = useState('');
  const [monthFilter, setMonthFilter] = useState('');
  const [ownerFilter, setOwnerFilter] = useState('');
  const [statusFilter, setStatusFilter] = useState<'' | AuditStatus>('');
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(15);

  async function load() {
    setLoading(true);
    setLoadError(null);
    try {
      const [employeeRows, agentRows, visitRows, kpiRes] = await Promise.all([
        fetchAll<Hr_employees>(
          (skipToken) => Hr_employeesService.getAll({ select: ['hr_employeeid', 'hr_fullname'], filter: 'statecode eq 0', maxPageSize: 5000, skipToken }),
          'Could not load employees.'
        ),
        fetchAll<{ cr301_agentid: string; cr301_newcolumn?: string; cr18c_area?: number }>(
          (skipToken) => Cr301_agentsService.getAll({ select: ['cr301_agentid', 'cr301_newcolumn', 'cr18c_area'], filter: 'cfm_iscollection eq true', maxPageSize: 5000, skipToken }),
          'Could not load task owners.'
        ),
        fetchAll<Cr301_visits>(
          (skipToken) =>
            Cr301_visitsService.getAll({
              select: ['cr301_visitid', 'cr301_visitdate', 'createdon', '_cr301_agent_value', '_cfm_agent2_value'],
              filter: `cfm_tagename eq ${TAG_COLLECTION}`,
              maxPageSize: 5000,
              skipToken,
            }),
          'Could not load visits.'
        ),
        Strategy_kpisesService.getAll({ select: ['strategy_kpisid'], filter: `strategy_newcolumn eq '${VISIT_KPI_NAME}' and statecode eq 0`, top: 1 }),
      ]);

      const kpiId = kpiRes.success ? kpiRes.data?.[0]?.strategy_kpisid : undefined;
      if (!kpiId) throw new Error(`KPI "${VISIT_KPI_NAME}" not found.`);

      const achRows = await fetchAll<Pm_kpiachievments>(
        (skipToken) =>
          Pm_kpiachievmentsService.getAll({
            select: ['pm_kpiachievmentid', 'pm_month', 'pm_year'],
            filter: `_pm_kpi_value eq ${kpiId} and statecode eq 0`,
            maxPageSize: 5000,
            skipToken,
          }),
        'Could not load KPI achievements.'
      );

      const achIdFilter = achRows.map((a) => `_stf_total_value eq ${a.pm_kpiachievmentid}`).join(' or ');
      const brkRows = achRows.length
        ? await fetchAll<Stf_kpiachievmentbreakdowns>(
            (skipToken) =>
              Stf_kpiachievmentbreakdownsService.getAll({
                select: ['stf_kpiachievmentbreakdownid', '_stf_total_value', '_stf_employee_value'],
                filter: `(${achIdFilter}) and statecode eq 0`,
                maxPageSize: 5000,
                skipToken,
              }),
            'Could not load KPI achievement breakdowns.'
          )
        : [];

      // Not filtered by a breakdown-id OR-chain like achRows/brkRows are — with
      // one breakdown per employee per month, that list gets long enough to
      // risk the query silently failing/truncating (this is why Plan/Actual
      // showed as "—" for every row). VisitPlan.tsx avoids the same trap by
      // fetching all active weeks unscoped and joining client-side — mirrored
      // here for the same reason.
      const weekRows = await fetchAll<Cfm_collectionplanweeks>(
        (skipToken) =>
          Cfm_collectionplanweeksService.getAll({
            select: ['cfm_weeknumber', 'cfm_plan', 'cfm_actual', '_cfm_kpiachievementbreakdown_value'],
            filter: 'statecode eq 0',
            maxPageSize: 5000,
            skipToken,
          }),
        'Could not load collection plan weeks.'
      );

      const achById = new Map(achRows.map((a) => [a.pm_kpiachievmentid, a]));
      const brkById = new Map(brkRows.map((b) => [b.stf_kpiachievmentbreakdownid, b]));
      const employeeByKey = new Map(employeeRows.map((e) => [firstWords(e.hr_fullname || '', 3), e]));
      const agentById = new Map(agentRows.map((a) => [a.cr301_agentid, a]));

      // employeeId|year|month|week -> { plan, actual }
      const planMap = new Map<string, { plan: number; actual: number }>();
      weekRows.forEach((w) => {
        const brkId = w._cfm_kpiachievementbreakdown_value;
        const brk = brkId ? brkById.get(brkId) : undefined;
        const achId = brk?._stf_total_value;
        const ach = achId ? achById.get(achId) : undefined;
        const empId = brk?._stf_employee_value;
        if (!ach || !empId) return;
        const key = `${empId}|${N(ach.pm_year)}|${N(ach.pm_month)}|${N(w.cfm_weeknumber)}`;
        planMap.set(key, { plan: N(w.cfm_plan), actual: N(w.cfm_actual) });
      });

      const auditRows: AuditRow[] = [];
      visitRows.forEach((v) => {
        const visitDateIso = v.cr301_visitdate;
        const createdOnIso = v.createdon;
        if (!visitDateIso || !createdOnIso) return;
        const visitDate = new Date(visitDateIso);
        const createdOn = new Date(createdOnIso);
        if (Number.isNaN(visitDate.getTime()) || Number.isNaN(createdOn.getTime())) return;

        const extra = v as unknown as UnmappedVisitFields;
        const agentIds = Array.from(new Set([v._cr301_agent_value, extra._cfm_agent2_value].filter((id): id is string => !!id)));
        if (!agentIds.length) return;

        // Local calendar date on both sides — pm_month/pm_year and the day
        // buckets are both defined in local wall-clock terms (see the
        // Visits.tsx adjustVisitPlanActual comment on the same subtlety).
        const expectedYear = visitDate.getFullYear();
        const expectedMonth = visitDate.getMonth() + 1;
        const expectedWeek = monthDayBucket(visitDate.getDate());
        const createdYear = createdOn.getFullYear();
        const createdMonth = createdOn.getMonth() + 1;
        const createdWeek = monthDayBucket(createdOn.getDate());
        const daysDiff = daysBetween(visitDate, createdOn);

        let status: AuditStatus;
        if (expectedYear !== createdYear || expectedMonth !== createdMonth) status = 'wrong-month';
        else if (expectedWeek !== createdWeek) status = 'wrong-week';
        else if (daysDiff > 0) status = 'late-same-week';
        else status = 'on-time';

        agentIds.forEach((agentId) => {
          const agent = agentById.get(agentId);
          const ownerName = agent?.cr301_newcolumn || '(unknown owner)';
          const employee = employeeByKey.get(firstWords(ownerName, 3));
          const planEntry = employee ? planMap.get(`${employee.hr_employeeid}|${expectedYear}|${expectedMonth}|${expectedWeek}`) : undefined;
          const areaValue = agent?.cr18c_area == null ? null : Number(agent.cr18c_area);
          const rowRegion: 'EG' | 'KSA' | '' = areaValue == null ? '' : areaValue === AREA_KSA ? 'KSA' : 'EG';

          auditRows.push({
            key: `${v.cr301_visitid}-${agentId}`,
            visitId: v.cr301_visitid,
            ownerName,
            employeeId: employee?.hr_employeeid || '',
            visitDate: visitDateIso,
            createdOn: createdOnIso,
            expectedWeek,
            expectedMonth,
            expectedYear,
            createdWeek,
            createdMonth,
            createdYear,
            daysDiff,
            plan: planEntry ? planEntry.plan : null,
            actual: planEntry ? planEntry.actual : null,
            variance: planEntry ? planEntry.actual - planEntry.plan : null,
            status,
            region: rowRegion,
          });
        });
      });

      auditRows.sort((a, b) => b.createdOn.localeCompare(a.createdOn));
      setRows(auditRows);
    } catch (e) {
      setLoadError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    load();
  }, []);

  const monthOptions = useMemo(() => {
    const map = new Map<string, string>();
    rows.forEach((r) => {
      const key = monthKeyOf(r.expectedYear, r.expectedMonth);
      if (!map.has(key)) map.set(key, `${MONTHS[r.expectedMonth] || r.expectedMonth} ${r.expectedYear}`);
    });
    return Array.from(map.entries()).sort((a, b) => b[0].localeCompare(a[0]));
  }, [rows]);

  const ownerOptions = useMemo(() => Array.from(new Set(rows.map((r) => r.ownerName))).sort(), [rows]);

  const view = useMemo(() => {
    const q = query.trim().toLowerCase();
    return rows.filter((r) => {
      if (q && !r.ownerName.toLowerCase().includes(q)) return false;
      if (monthFilter && monthKeyOf(r.expectedYear, r.expectedMonth) !== monthFilter) return false;
      if (ownerFilter && r.ownerName !== ownerFilter) return false;
      if (statusFilter && r.status !== statusFilter) return false;
      // App-wide Region toggle (Layout header) — unset region on a visit's
      // owner counts as Egypt, same "everything else is Egypt" convention
      // useRegion.ts uses for BUs.
      if (region && (region === 'KSA' ? r.region !== 'KSA' : r.region === 'KSA')) return false;
      return true;
    });
  }, [rows, query, monthFilter, ownerFilter, statusFilter, region]);

  useEffect(() => {
    setPage(1);
  }, [query, monthFilter, ownerFilter, statusFilter, region]);

  const totalPages = Math.max(1, Math.ceil(view.length / pageSize));
  const clampedPage = Math.min(page, totalPages);
  const pageRows = view.slice((clampedPage - 1) * pageSize, (clampedPage - 1) * pageSize + pageSize);

  const summary = useMemo(() => {
    let onTime = 0;
    let late = 0;
    let wrong = 0;
    view.forEach((r) => {
      if (r.status === 'on-time') onTime++;
      else if (r.status === 'late-same-week') late++;
      else wrong++;
    });
    const compliance = view.length ? (onTime / view.length) * 100 : 0;
    return { total: view.length, onTime, late, wrong, compliance };
  }, [view]);

  const ownerSummaries = useMemo(() => {
    const groups = new Map<string, { name: string; total: number; correct: number; late: number; wrongWeek: number; weekKeys: Set<string> }>();
    const planLookup = new Map<string, { plan: number; actual: number }>();
    view.forEach((r) => {
      const groupKey = r.employeeId || r.ownerName;
      const g = groups.get(groupKey) || { name: r.ownerName, total: 0, correct: 0, late: 0, wrongWeek: 0, weekKeys: new Set<string>() };
      g.total++;
      if (r.status === 'on-time') g.correct++;
      else if (r.status === 'late-same-week') g.late++;
      else g.wrongWeek++;
      if (r.employeeId && r.plan != null) {
        const wk = `${r.employeeId}|${r.expectedYear}|${r.expectedMonth}|${r.expectedWeek}`;
        g.weekKeys.add(wk);
        planLookup.set(wk, { plan: r.plan, actual: r.actual || 0 });
      }
      groups.set(groupKey, g);
    });
    const list: OwnerSummary[] = Array.from(groups.values()).map((g) => {
      let planned = 0;
      let actual = 0;
      g.weekKeys.forEach((wk) => {
        const entry = planLookup.get(wk);
        if (entry) {
          planned += entry.plan;
          actual += entry.actual;
        }
      });
      return {
        name: g.name,
        total: g.total,
        correct: g.correct,
        late: g.late,
        wrongWeek: g.wrongWeek,
        planned,
        actual,
        variance: actual - planned,
        timingCompliance: g.total ? (g.correct / g.total) * 100 : 0,
        achievementPct: planned > 0 ? (actual / planned) * 100 : 0,
      };
    });
    return list.sort((a, b) => b.total - a.total);
  }, [view]);

  const weekSummaries = useMemo(() => {
    const groups = new Map<
      string,
      { name: string; year: number; month: number; week: number; total: number; correct: number; late: number; wrongWeek: number; plan: number | null; actual: number | null; variance: number | null }
    >();
    view.forEach((r) => {
      const groupKey = `${r.employeeId || r.ownerName}|${r.expectedYear}|${r.expectedMonth}|${r.expectedWeek}`;
      const g = groups.get(groupKey) || {
        name: r.ownerName,
        year: r.expectedYear,
        month: r.expectedMonth,
        week: r.expectedWeek,
        total: 0,
        correct: 0,
        late: 0,
        wrongWeek: 0,
        plan: r.plan,
        actual: r.actual,
        variance: r.variance,
      };
      g.total++;
      if (r.status === 'on-time') g.correct++;
      else if (r.status === 'late-same-week') g.late++;
      else g.wrongWeek++;
      groups.set(groupKey, g);
    });
    return Array.from(groups.values()).sort((a, b) => b.year - a.year || b.month - a.month || a.week - b.week || a.name.localeCompare(b.name));
  }, [view]);

  const monthSummaries = useMemo(() => {
    const groups = new Map<
      string,
      { name: string; year: number; month: number; total: number; correct: number; late: number; wrongWeek: number; weekKeys: Set<string> }
    >();
    const planLookup = new Map<string, { plan: number; actual: number }>();
    view.forEach((r) => {
      const groupKey = `${r.employeeId || r.ownerName}|${r.expectedYear}|${r.expectedMonth}`;
      const g = groups.get(groupKey) || { name: r.ownerName, year: r.expectedYear, month: r.expectedMonth, total: 0, correct: 0, late: 0, wrongWeek: 0, weekKeys: new Set<string>() };
      g.total++;
      if (r.status === 'on-time') g.correct++;
      else if (r.status === 'late-same-week') g.late++;
      else g.wrongWeek++;
      if (r.employeeId && r.plan != null) {
        const wk = `${r.employeeId}|${r.expectedYear}|${r.expectedMonth}|${r.expectedWeek}`;
        g.weekKeys.add(wk);
        planLookup.set(wk, { plan: r.plan, actual: r.actual || 0 });
      }
      groups.set(groupKey, g);
    });
    return Array.from(groups.values())
      .map((g) => {
        let planned = 0;
        let actual = 0;
        g.weekKeys.forEach((wk) => {
          const entry = planLookup.get(wk);
          if (entry) {
            planned += entry.plan;
            actual += entry.actual;
          }
        });
        return {
          name: g.name,
          year: g.year,
          month: g.month,
          total: g.total,
          correct: g.correct,
          late: g.late,
          wrongWeek: g.wrongWeek,
          planned,
          actual,
          variance: actual - planned,
          timingCompliance: g.total ? (g.correct / g.total) * 100 : 0,
          achievementPct: planned > 0 ? (actual / planned) * 100 : 0,
        };
      })
      .sort((a, b) => b.year - a.year || b.month - a.month || a.name.localeCompare(b.name));
  }, [view]);

  function exportDetail() {
    if (!view.length) return;
    const header = ['Task Owner', 'Region', 'Visit Date', 'Created On', 'Expected Week', 'Created Week', 'Days Difference', 'Plan', 'Actual', 'Variance', 'Audit Status'];
    const lines = [header.join(',')];
    view.forEach((r) => {
      lines.push(
        [
          r.ownerName,
          r.region === 'KSA' ? 'KSA' : r.region === 'EG' ? 'Egypt' : '—',
          fmtDate(r.visitDate),
          fmtDateTime(r.createdOn),
          `Week ${r.expectedWeek} (${MONTHS[r.expectedMonth]} ${r.expectedYear})`,
          `Week ${r.createdWeek} (${MONTHS[r.createdMonth]} ${r.createdYear})`,
          r.daysDiff,
          r.plan ?? '—',
          r.actual ?? '—',
          r.variance ?? '—',
          STATUS_LABEL[r.status],
        ]
          .map(escapeForCsv)
          .join(',')
      );
    });
    downloadCsv('visit_audit_detail.csv', lines);
  }

  function exportSummary() {
    if (!ownerSummaries.length) return;
    const header = ['Task Owner', 'Total Visits', 'Correct', 'Late', 'Wrong Week', 'Planned', 'Actual', 'Variance', 'Timing Compliance %', 'Achievement %'];
    const lines = [header.join(',')];
    ownerSummaries.forEach((s) => {
      lines.push(
        [s.name, s.total, s.correct, s.late, s.wrongWeek, s.planned, s.actual, s.variance, s.timingCompliance.toFixed(1) + '%', s.achievementPct.toFixed(1) + '%']
          .map(escapeForCsv)
          .join(',')
      );
    });
    downloadCsv('visit_audit_summary.csv', lines);
  }

  function exportByWeek() {
    if (!weekSummaries.length) return;
    const header = ['Task Owner', 'Month', 'Week', 'Total Visits', 'Correct', 'Late', 'Wrong Week', 'Plan', 'Actual', 'Variance'];
    const lines = [header.join(',')];
    weekSummaries.forEach((s) => {
      lines.push(
        [s.name, `${MONTHS[s.month]} ${s.year}`, `Week ${s.week}`, s.total, s.correct, s.late, s.wrongWeek, s.plan ?? '—', s.actual ?? '—', s.variance ?? '—']
          .map(escapeForCsv)
          .join(',')
      );
    });
    downloadCsv('visit_audit_by_week.csv', lines);
  }

  function exportByMonth() {
    if (!monthSummaries.length) return;
    const header = ['Task Owner', 'Month', 'Total Visits', 'Correct', 'Late', 'Wrong Week', 'Planned', 'Actual', 'Variance', 'Timing Compliance %', 'Achievement %'];
    const lines = [header.join(',')];
    monthSummaries.forEach((s) => {
      lines.push(
        [
          s.name,
          `${MONTHS[s.month]} ${s.year}`,
          s.total,
          s.correct,
          s.late,
          s.wrongWeek,
          s.planned,
          s.actual,
          s.variance,
          s.timingCompliance.toFixed(1) + '%',
          s.achievementPct.toFixed(1) + '%',
        ]
          .map(escapeForCsv)
          .join(',')
      );
    });
    downloadCsv('visit_audit_by_month.csv', lines);
  }

  function statusPill(status: AuditStatus) {
    return <span className={`acc-visit-status acc-visit-status-${STATUS_TONE[status]}`}>{STATUS_LABEL[status]}</span>;
  }

  function pagerButtons(): (number | 'ellipsis')[] {
    const pages: (number | 'ellipsis')[] = [];
    for (let p = 1; p <= totalPages; p++) {
      if (p === 1 || p === totalPages || Math.abs(p - clampedPage) <= 1) pages.push(p);
      else if (Math.abs(p - clampedPage) === 2) pages.push('ellipsis');
    }
    return pages;
  }

  return (
    <div className="acc-wrap">
      <div className="acc-head">
        <div>
          <h1 className="acc-title">Visit Audit</h1>
          <div className="acc-sub">Visit Date vs. Created On · Expected Week vs. Created Week · Plan vs. Actual · per task owner</div>
        </div>
        <div className="acc-actions">
          <button className="acc-btn" onClick={exportDetail}>
            <Icon className="fa-solid fa-download" /> Export Detail
          </button>
          <button className="acc-btn" onClick={exportSummary}>
            <Icon className="fa-solid fa-download" /> Export Summary
          </button>
          <button className="acc-btn" onClick={exportByWeek}>
            <Icon className="fa-solid fa-download" /> Export by Week
          </button>
          <button className="acc-btn" onClick={exportByMonth}>
            <Icon className="fa-solid fa-download" /> Export by Month
          </button>
        </div>
      </div>

      <div className="kpi-grid">
        <div className="kpi-card">
          <div className="kpi-top">
            <span className="kpi-icon kpi-icon-gold">
              <Icon className="fa-solid fa-list-check" />
            </span>
            <span className="kpi-label">Total visits</span>
          </div>
          <div className="kpi-value">{loading ? '—' : summary.total}</div>
        </div>
        <div className="kpi-card kpi-ok">
          <div className="kpi-top">
            <span className="kpi-icon kpi-icon-ok">
              <Icon className="fa-solid fa-circle-check" />
            </span>
            <span className="kpi-label">On-time</span>
          </div>
          <div className="kpi-value">{loading ? '—' : summary.onTime}</div>
        </div>
        <div className="kpi-card kpi-warn">
          <div className="kpi-top">
            <span className="kpi-icon kpi-icon-warn">
              <Icon className="fa-solid fa-clock" />
            </span>
            <span className="kpi-label">Late (same week)</span>
          </div>
          <div className="kpi-value">{loading ? '—' : summary.late}</div>
        </div>
        <div className="kpi-card kpi-danger">
          <div className="kpi-top">
            <span className="kpi-icon kpi-icon-danger">
              <Icon className="fa-solid fa-triangle-exclamation" />
            </span>
            <span className="kpi-label">Wrong week / month</span>
          </div>
          <div className="kpi-value">{loading ? '—' : summary.wrong}</div>
        </div>
      </div>

      {loadError && (
        <div className="acc-state acc-state-error">
          <Icon className="fa-solid fa-triangle-exclamation" />
          {loadError}
        </div>
      )}

      <div className="tbl-shell">
        <div className="tbl-hdr">
          <h2 className="va-sec-title">Per task owner</h2>
        </div>
        <div className="acc-tablewrap">
          <table>
            <thead>
              <tr>
                <th>Task Owner</th>
                <th className="acc-num">Total</th>
                <th className="acc-num">Correct</th>
                <th className="acc-num">Late</th>
                <th className="acc-num">Wrong Week</th>
                <th className="acc-num">Planned</th>
                <th className="acc-num">Actual</th>
                <th className="acc-num">Variance</th>
                <th className="acc-num">Timing Compliance</th>
                <th className="acc-num">Achievement %</th>
              </tr>
            </thead>
            <tbody>
              {loading ? (
                Array.from({ length: 4 }).map((_, i) => (
                  <tr key={i}>
                    {Array.from({ length: 10 }).map((__, j) => (
                      <td key={j}>
                        <div className="acc-skel"></div>
                      </td>
                    ))}
                  </tr>
                ))
              ) : ownerSummaries.length === 0 ? (
                <tr>
                  <td colSpan={10}>
                    <div className="acc-state">
                      <Icon className="fa-solid fa-inbox" />
                      No visits match your filters.
                    </div>
                  </td>
                </tr>
              ) : (
                ownerSummaries.map((s) => (
                  <tr key={s.name}>
                    <td className="acc-name">{s.name}</td>
                    <td className="acc-num">{s.total}</td>
                    <td className="acc-num">{s.correct}</td>
                    <td className="acc-num">{s.late}</td>
                    <td className="acc-num">{s.wrongWeek}</td>
                    <td className="acc-num">{s.planned}</td>
                    <td className="acc-num">{s.actual}</td>
                    <td className="acc-num">{s.variance}</td>
                    <td className="acc-num">{s.timingCompliance.toFixed(1)}%</td>
                    <td className="acc-num">{s.achievementPct.toFixed(1)}%</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>

      <div className="tbl-shell">
        <div className="tbl-hdr">
          <div className="dm-filter-bar">
            <div className="acc-search">
              <Icon className="fa-solid fa-magnifying-glass" />
              <input placeholder="Search task owner…" value={query} onChange={(e) => setQuery(e.target.value)} />
            </div>
            <select className="acc-f" value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)}>
              <option value="">All months</option>
              {monthOptions.map(([key, label]) => (
                <option key={key} value={key}>
                  {label}
                </option>
              ))}
            </select>
            <select className="acc-f" value={ownerFilter} onChange={(e) => setOwnerFilter(e.target.value)}>
              <option value="">All owners</option>
              {ownerOptions.map((n) => (
                <option key={n} value={n}>
                  {n}
                </option>
              ))}
            </select>
            <select className="acc-f" value={statusFilter} onChange={(e) => setStatusFilter(e.target.value as '' | AuditStatus)}>
              <option value="">All statuses</option>
              {(Object.keys(STATUS_LABEL) as AuditStatus[]).map((s) => (
                <option key={s} value={s}>
                  {STATUS_LABEL[s]}
                </option>
              ))}
            </select>
          </div>
          <div className="dm-count">
            <Icon className="fa-solid fa-list-check" /> {view.length} visit{view.length === 1 ? '' : 's'}
          </div>
        </div>
        <div className="acc-tablewrap">
          <table>
            <thead>
              <tr>
                <th>Task Owner</th>
                <th>Region</th>
                <th>Visit Date</th>
                <th>Created On</th>
                <th>Expected Week</th>
                <th>Created Week</th>
                <th className="acc-num">Days Diff</th>
                <th className="acc-num">Plan</th>
                <th className="acc-num">Actual</th>
                <th className="acc-num">Variance</th>
                <th>Audit Status</th>
              </tr>
            </thead>
            <tbody>
              {loading ? (
                Array.from({ length: 6 }).map((_, i) => (
                  <tr key={i}>
                    {Array.from({ length: 11 }).map((__, j) => (
                      <td key={j}>
                        <div className="acc-skel"></div>
                      </td>
                    ))}
                  </tr>
                ))
              ) : pageRows.length === 0 ? (
                <tr>
                  <td colSpan={11}>
                    <div className="acc-state">
                      <Icon className="fa-solid fa-inbox" />
                      No visits match your filters.
                    </div>
                  </td>
                </tr>
              ) : (
                pageRows.map((r) => (
                  <tr key={r.key}>
                    <td className="acc-name">{r.ownerName}</td>
                    <td>{r.region ? (r.region === 'KSA' ? 'KSA' : 'Egypt') : <span className="vis-muted">—</span>}</td>
                    <td>{fmtDate(r.visitDate)}</td>
                    <td>{fmtDateTime(r.createdOn)}</td>
                    <td>
                      Week {r.expectedWeek} <span className="vis-muted">({MONTHS[r.expectedMonth]} {r.expectedYear})</span>
                    </td>
                    <td>
                      Week {r.createdWeek} <span className="vis-muted">({MONTHS[r.createdMonth]} {r.createdYear})</span>
                    </td>
                    <td className="acc-num">{r.daysDiff}</td>
                    <td className="acc-num">{r.plan ?? <span className="vis-muted">—</span>}</td>
                    <td className="acc-num">{r.actual ?? <span className="vis-muted">—</span>}</td>
                    <td className="acc-num">{r.variance ?? <span className="vis-muted">—</span>}</td>
                    <td>{statusPill(r.status)}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>

        {!loading && view.length > 0 && (
          <div className="acc-foot">
            <div className="dm-page-size">
              <span>Rows per page:</span>
              <select
                className="acc-f"
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
          </div>
        )}
      </div>
    </div>
  );
}
