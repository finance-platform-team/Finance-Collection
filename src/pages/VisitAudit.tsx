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
// strategy_kpises can hold more than one KPI named "No. of Visits" across
// different Function/Department combinations — pin down the one this page
// means instead of taking whichever same-named record the server returns first.
// Function/Department are lookups (strategy_function / strategy_department),
// not text columns, so there's nothing to filter server-side by name — the
// display text only exists as the @OData FormattedValue annotation on the
// selected _value field, read client-side via formattedValue() below.
const VISIT_KPI_FUNCTION = 'Collection/Finance';
const VISIT_KPI_DEPARTMENT = 'Finance';

function formattedValue(record: unknown, attribute: string): string {
  const raw = (record as Record<string, unknown>)[`${attribute}@OData.Community.Display.V1.FormattedValue`];
  return typeof raw === 'string' ? raw : '';
}
const TAG_COLLECTION = 766340000; // cfm_tagename — "Collection"
// cr301_visitstatus "Cancelled" — same value as VisitPlan.tsx / Visits.tsx.
// A cancelled visit never happened, so it's left out of the audit entirely.
const VISIT_STATUS_CANCELLED = 983080002;
// cr301_visittype "Double Visit" — same value as Visits.tsx.
const VISIT_TYPE_DOUBLE = 983080001;
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
  agentId: string;
  ownerName: string;
  accountId: string;
  company: string;
  isDouble: boolean;
  doubleWith: string;
  /** Visits by this owner to this company in the same expected week (incl. this one). */
  repeatCount: number;
  modifiedOn: string;
  createdBy: string;
  modifiedBy: string;
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
  // Defaults to the current month; '' means "All months".
  const [monthFilter, setMonthFilter] = useState(() => monthKeyOf(new Date().getFullYear(), new Date().getMonth() + 1));
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
              select: [
                'cr301_visitid',
                'cr301_visitdate',
                'createdon',
                'modifiedon',
                'cr301_visitstatus',
                'cr301_visittype',
                '_cr301_account_value',
                '_cr301_agent_value',
                '_cfm_agent2_value',
                '_createdby_value',
                '_modifiedby_value',
              ],
              filter: `cfm_tagename eq ${TAG_COLLECTION}`,
              maxPageSize: 5000,
              skipToken,
            }),
          'Could not load visits.'
        ),
        Strategy_kpisesService.getAll({
          select: ['strategy_kpisid', '_strategy_function_value', '_strategy_department_value'],
          filter: `strategy_newcolumn eq '${VISIT_KPI_NAME}' and statecode eq 0`,
        }),
      ]);

      const matchingKpi = kpiRes.success
        ? (kpiRes.data || []).find(
            (k) =>
              formattedValue(k, '_strategy_function_value').trim().toLowerCase() === VISIT_KPI_FUNCTION.toLowerCase() &&
              formattedValue(k, '_strategy_department_value').trim().toLowerCase() === VISIT_KPI_DEPARTMENT.toLowerCase()
          )
        : undefined;
      const kpiId = matchingKpi?.strategy_kpisid;
      if (!kpiId) throw new Error(`KPI "${VISIT_KPI_NAME}" (${VISIT_KPI_FUNCTION} / ${VISIT_KPI_DEPARTMENT}) not found.`);

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
        if (v.cr301_visitstatus === VISIT_STATUS_CANCELLED) return;
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

        const agentName = (id: string) =>
          agentById.get(id)?.cr301_newcolumn ||
          formattedValue(v, id === v._cr301_agent_value ? '_cr301_agent_value' : '_cfm_agent2_value') ||
          '(unknown owner)';
        const isDouble = agentIds.length > 1 || v.cr301_visittype === VISIT_TYPE_DOUBLE;

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
            agentId,
            ownerName,
            accountId: v._cr301_account_value || '',
            company: formattedValue(v, '_cr301_account_value'),
            isDouble,
            doubleWith: agentIds.filter((id) => id !== agentId).map(agentName).join(', '),
            repeatCount: 1,
            modifiedOn: v.modifiedon || '',
            createdBy: formattedValue(v, '_createdby_value'),
            modifiedBy: formattedValue(v, '_modifiedby_value'),
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

      // Repeated visits: the same owner logging the same company more than
      // once inside one week bucket (by visit date) — this also catches the
      // same visit entered twice at the same time.
      const repeatKey = (r: AuditRow) => `${r.agentId}|${r.accountId}|${r.expectedYear}|${r.expectedMonth}|${r.expectedWeek}`;
      const repeatCounts = new Map<string, number>();
      auditRows.forEach((r) => {
        if (r.accountId) repeatCounts.set(repeatKey(r), (repeatCounts.get(repeatKey(r)) || 0) + 1);
      });
      auditRows.forEach((r) => {
        if (r.accountId) r.repeatCount = repeatCounts.get(repeatKey(r)) || 1;
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
    // Always offer the current month so the default selection shows in the dropdown.
    const now = new Date();
    const curKey = monthKeyOf(now.getFullYear(), now.getMonth() + 1);
    if (!map.has(curKey)) map.set(curKey, `${MONTHS[now.getMonth() + 1]} ${now.getFullYear()}`);
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

  const [exporting, setExporting] = useState(false);

  /**
   * One audit workbook — Summary / Double Visits / Repeated Visits / Details
   * Data — built from the currently filtered rows. Replaces the old separate
   * Detail / Summary / By Week / By Month CSV exports.
   */
  async function exportAudit() {
    if (!view.length) return;
    setExporting(true);
    try {
      const ExcelJS = await import('exceljs');
      const wb = new ExcelJS.Workbook();

      const monthLabel = (r: AuditRow) => `${MONTHS[r.expectedMonth]} ${r.expectedYear}`;
      const weekLabel = (r: AuditRow) => `Week ${r.expectedWeek}`;
      const target = (r: AuditRow) => r.plan ?? '';
      const asDate = (iso: string) => (iso ? new Date(iso) : '');
      const ownerWeekKey = (r: AuditRow) => `${r.agentId}|${r.expectedYear}|${r.expectedMonth}|${r.expectedWeek}`;
      const sorted = view
        .slice()
        .sort(
          (a, b) =>
            a.ownerName.localeCompare(b.ownerName) ||
            a.expectedYear - b.expectedYear ||
            a.expectedMonth - b.expectedMonth ||
            a.expectedWeek - b.expectedWeek ||
            a.visitDate.localeCompare(b.visitDate)
        );

      type Col = { header: string; width: number; date?: 'date' | 'datetime' };
      type Cell = string | number | Date;
      const border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } } as const;

      // Same layout as the audit template: table starts at B2, column A is a margin.
      const addSheet = (name: string, columns: Col[], data: Cell[][]) => {
        const ws = wb.addWorksheet(name, { views: [{ state: 'frozen', ySplit: 2 }] });
        ws.getColumn(1).width = 3;
        columns.forEach((c, i) => {
          const col = ws.getColumn(i + 2);
          col.width = c.width;
          if (c.date) col.numFmt = c.date === 'date' ? 'dd-mmm-yyyy' : 'dd-mmm-yyyy hh:mm';
        });
        const header = ws.getRow(2);
        header.height = 32;
        columns.forEach((c, i) => {
          const cell = header.getCell(i + 2);
          cell.value = c.header;
          cell.font = { bold: true, color: { argb: 'FFFFFFFF' } };
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };
          cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          cell.border = border;
        });
        data.forEach((values, ri) => {
          const row = ws.getRow(ri + 3);
          values.forEach((val, ci) => {
            const cell = row.getCell(ci + 2);
            cell.value = val === '' ? null : val;
            cell.border = border;
            cell.alignment = { vertical: 'middle', horizontal: typeof val === 'number' || val instanceof Date ? 'center' : 'left' };
          });
        });
        ws.autoFilter = { from: { row: 2, column: 2 }, to: { row: 2, column: columns.length + 1 } };
      };

      // ---- Summary: one row per owner per week ----
      const weekGroups = new Map<string, { r: AuditRow; same: number; wrong: number; total: number; double: number; repeated: number }>();
      sorted.forEach((r) => {
        const k = ownerWeekKey(r);
        const g = weekGroups.get(k) || { r, same: 0, wrong: 0, total: 0, double: 0, repeated: 0 };
        g.total++;
        if (r.status === 'on-time' || r.status === 'late-same-week') g.same++;
        else g.wrong++;
        if (r.isDouble) g.double++;
        if (r.repeatCount > 1) g.repeated++;
        weekGroups.set(k, g);
      });
      addSheet(
        'Summary',
        [
          { header: 'Agent', width: 28 },
          { header: 'Month', width: 16 },
          { header: 'Week', width: 10 },
          { header: 'Weekly Target', width: 12 },
          { header: 'Visits Done in Same Week', width: 16 },
          { header: 'Visits Added in Wrong Week', width: 16 },
          { header: 'Total Visits', width: 12 },
          { header: 'Double Visits', width: 12 },
          { header: 'Repeated Visits', width: 12 },
        ],
        Array.from(weekGroups.values()).map((g) => [g.r.ownerName, monthLabel(g.r), weekLabel(g.r), target(g.r), g.same, g.wrong, g.total, g.double, g.repeated])
      );

      // ---- Double Visits: every double visit + the owner's double count that week ----
      const doubleRows = sorted.filter((r) => r.isDouble);
      const doubleCount = new Map<string, number>();
      doubleRows.forEach((r) => doubleCount.set(ownerWeekKey(r), (doubleCount.get(ownerWeekKey(r)) || 0) + 1));
      addSheet(
        'Double Visits',
        [
          { header: 'Agent Name', width: 28 },
          { header: 'Month', width: 16 },
          { header: 'Week', width: 10 },
          { header: 'Weekly Target', width: 12 },
          { header: 'Num. Double Visits', width: 12 },
          { header: 'Visit Date', width: 14, date: 'date' },
          { header: 'Double Visit With', width: 28 },
          { header: 'Company Name', width: 34 },
        ],
        doubleRows.map((r) => [r.ownerName, monthLabel(r), weekLabel(r), target(r), doubleCount.get(ownerWeekKey(r)) || 0, asDate(r.visitDate), r.doubleWith, r.company])
      );

      // ---- Repeated Visits: same owner + same company more than once in a week ----
      addSheet(
        'Repeated Visits',
        [
          { header: 'Agent Name', width: 28 },
          { header: 'Month', width: 16 },
          { header: 'Week', width: 10 },
          { header: 'Weekly Target', width: 12 },
          { header: 'Company Name', width: 34 },
          { header: 'No. of Visits in Same Week', width: 14 },
          { header: 'Visit Date', width: 14, date: 'date' },
          { header: 'Created On', width: 18, date: 'datetime' },
          { header: 'Double With', width: 28 },
        ],
        sorted
          .filter((r) => r.repeatCount > 1)
          .sort((a, b) => a.ownerName.localeCompare(b.ownerName) || a.company.localeCompare(b.company) || a.visitDate.localeCompare(b.visitDate))
          .map((r) => [r.ownerName, monthLabel(r), weekLabel(r), target(r), r.company, r.repeatCount, asDate(r.visitDate), asDate(r.createdOn), r.doubleWith])
      );

      // ---- Details Data: every visit row ----
      addSheet(
        'Details Data',
        [
          { header: 'Task Owner', width: 28 },
          { header: 'Region', width: 10 },
          { header: 'Month', width: 16 },
          { header: 'Week', width: 10 },
          { header: 'Weekly Target', width: 12 },
          { header: 'Visit Date', width: 14, date: 'date' },
          { header: 'Company', width: 34 },
          { header: 'Created Week', width: 22 },
          { header: 'Days Difference', width: 11 },
          { header: 'Audit Status', width: 18 },
          { header: 'Wrong Week', width: 10 },
          { header: 'Double Visit?', width: 10 },
          { header: 'Double With', width: 28 },
          { header: 'Repeated Visits', width: 11 },
          { header: 'Created On', width: 18, date: 'datetime' },
          { header: 'Modified On', width: 18, date: 'datetime' },
          { header: 'Created By', width: 24 },
          { header: 'Modified By', width: 24 },
        ],
        sorted.map((r) => [
          r.ownerName,
          r.region === 'KSA' ? 'KSA' : r.region === 'EG' ? 'Egypt' : '',
          monthLabel(r),
          weekLabel(r),
          target(r),
          asDate(r.visitDate),
          r.company,
          `Week ${r.createdWeek} (${MONTHS[r.createdMonth]} ${r.createdYear})`,
          r.daysDiff,
          STATUS_LABEL[r.status],
          r.status === 'wrong-week' || r.status === 'wrong-month' ? 'Yes' : 'No',
          r.isDouble ? 'Yes' : 'No',
          r.doubleWith,
          r.repeatCount > 1 ? r.repeatCount : '',
          asDate(r.createdOn),
          asDate(r.modifiedOn),
          r.createdBy,
          r.modifiedBy,
        ])
      );

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'visit_audit.xlsx';
      link.click();
      URL.revokeObjectURL(url);
    } finally {
      setExporting(false);
    }
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
          <button className="acc-btn" onClick={exportAudit} disabled={exporting || loading || !view.length}>
            <Icon className="fa-solid fa-download" /> {exporting ? 'Exporting…' : 'Export Audit'}
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
