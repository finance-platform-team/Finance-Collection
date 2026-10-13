import Icon from '../components/Icon';
import { useEffect, useMemo, useRef, useState } from 'react';
import { getClient } from '@microsoft/power-apps/data';
import { dataSourcesInfo } from '../../.power/schemas/appschemas/dataSourcesInfo';
import { Cr301_visitsService } from '../generated/services/Cr301_visitsService';
import { Cr301_accountsService } from '../generated/services/Cr301_accountsService';
import { Cr301_agentsService } from '../generated/services/Cr301_agentsService';
import { Stf_kpiachievmentbreakdownsService } from '../generated/services/Stf_kpiachievmentbreakdownsService';
import { Cfm_collectionplanweeksService } from '../generated/services/Cfm_collectionplanweeksService';
import { Pm_kpiachievmentsService } from '../generated/services/Pm_kpiachievmentsService';
import { Strategy_kpisesService } from '../generated/services/Strategy_kpisesService';
import { Cr301_visitscr301_visitstatus } from '../generated/models/Cr301_visitsModel';
import type { Cr301_visitsBase, Cr301_visitscr301_visittype } from '../generated/models/Cr301_visitsModel';
import type { Cfm_collectionplanweeks } from '../generated/models/Cfm_collectionplanweeksModel';
import type { Strategy_kpises } from '../generated/models/Strategy_kpisesModel';
import { operationErrorMessage } from '../lib/errors';
import { fetchAll } from '../lib/dataverse';
import { visitStatusClass } from '../lib/visitStatus';
import { escapeForCsv, downloadCsv } from '../lib/csvExport';
import type { Region } from '../lib/useRegion';
import './Visits.css';

const VISIT_KPI_NAME = 'No. of Visits';
// Same cr18c_area choice on cr301_agents / same Egypt-vs-KSA mapping as
// Agents.tsx and VisitPlan.tsx (KSA = the real KSA choice value, Egypt =
// everything else, including unset).
const AREA_KSA = 989230004;
// Same cr301_country KSA value used on Accounts.tsx.
const ACCOUNT_REGION_KSA = 983080001;

// Same Function/Department pinning VisitPlan.tsx uses to disambiguate which
// "No. of Visits" strategy_kpises row is the real one (more than one can
// exist across different Function/Department combos).
const DEFAULT_KPI_FUNCTION = 'Collection/Finance';
const DEFAULT_KPI_DEPARTMENT = 'Finance';
// Function/Department are lookups with no separate name column to filter on
// server-side — the display text comes from the @OData FormattedValue
// annotation, same as VisitPlan.tsx's formattedValue().
function fmtVal(record: unknown, attribute: string): string {
  const raw = (record as Record<string, unknown>)[`${attribute}@OData.Community.Display.V1.FormattedValue`];
  return typeof raw === 'string' ? raw : '';
}
function isVisitsKpi(k: { strategy_newcolumn?: string }): boolean {
  if ((k.strategy_newcolumn || '').trim().toLowerCase() !== VISIT_KPI_NAME.toLowerCase()) return false;
  const fn = fmtVal(k, '_strategy_function_value').trim().toLowerCase();
  const dept = fmtVal(k, '_strategy_department_value').trim().toLowerCase();
  return fn === DEFAULT_KPI_FUNCTION.toLowerCase() && dept === DEFAULT_KPI_DEPARTMENT.toLowerCase();
}

type AdjustOutcome =
  | { status: 'ok'; weekNumber?: number }
  | { status: 'no-employee' | 'bad-date' | 'no-week'; agentName: string }
  | { status: 'no-kpi'; agentName: string }
  | { status: 'no-achievement'; agentName: string; month: number; year: number }
  | { status: 'no-breakdown'; agentName: string; month: number; year: number }
  | { status: 'error'; agentName: string; message: string };

// Same fixed calendar-day buckets VisitPlan.tsx uses when creating weeks
// (1-10, 11-18, 19-25, 26-end-of-month) — must stay in sync with
// MONTH_WEEK_BOUNDS there, since this is how a visit's date resolves to a
// cfm_weeknumber to adjust.
function monthDayBucket(day: number): number {
  if (day <= 10) return 1;
  if (day <= 18) return 2;
  if (day <= 25) return 3;
  return 4;
}

/**
 * For an agent's linked employee (cr301_agents.cfm_employee, a real Dataverse
 * lookup — no more matching hr_employees by name) + a reference date (the
 * Visit's createdon system timestamp — NOT the user-entered Visit Date), find
 * that employee's active "No. of Visits" KPI achievement breakdown for that
 * DATE'S OWN MONTH+YEAR (not just year — a September visit must land on the
 * September achievement, never fall back to August's), and the Collection
 * Plan Week under that breakdown whose cfm_weeknumber matches the reference
 * date's day-of-month bucket (day 1-10 -> week 1, 11-18 -> week 2, 19-25 ->
 * week 3, 26-end-of-month -> week 4) — then adjust that week's cfm_actual by
 * `delta` (+1 for a created/moved-in visit, -1 for a removed/moved-out one,
 * clamped at 0).
 *
 * Resolves the breakdown the same way VisitPlan.tsx does: find the "No. of
 * Visits" strategy_kpises row (pinned by Function/Department, since more than
 * one KPI can share that name), find its achievement(s) for this exact
 * pm_month/pm_year, then find this employee's breakdown whose
 * _stf_total_value points at one of those achievements. Deliberately does NOT
 * filter breakdowns on pms_kpinamefx — that's a formula/rollup column that
 * stays null on a freshly created breakdown until Dataverse's async
 * calculation job runs, so it made brand-new plans invisible to this lookup.
 *
 * Deliberately does NOT compare against cfm_weekfrom/cfm_weekto: matching by
 * cfm_weeknumber is the one plan-week-count-agnostic anchor both this code
 * and the plan's own week cards agree on, and doesn't require the two pages
 * to agree on exact date-range formatting.
 *
 * Mirrors what a Power Automate flow on cr301_visits would do, run inline
 * instead. Returns which step it stopped at rather than throwing, so the
 * caller can tell the user exactly why nothing was adjusted instead of
 * failing silently.
 */
async function adjustVisitPlanActual(agentName: string, employeeId: string, referenceDateIso: string, delta: 1 | -1): Promise<AdjustOutcome> {
  try {
    // No cfm_employee link on this agent (never saved with one, or saved
    // before Agents.tsx started binding it) — nothing to join the KPI side
    // to, so this has to stop here rather than guess by name.
    if (!employeeId) return { status: 'no-employee', agentName };

    const referenceDate = new Date(referenceDateIso);
    if (Number.isNaN(referenceDate.getTime())) return { status: 'bad-date', agentName };
    // Local calendar date, not UTC: pm_month/pm_year on the achievement (set from
    // `new Date().getMonth()` in VisitPlan.tsx) and the 1-10/11-18/19-25/26-31
    // day buckets are both defined in local wall-clock terms. Reading UTC here
    // instead disagreed near midnight in UTC+2/UTC+3 — a visit at 00:30 local on
    // the 1st is still the last day of the prior month in UTC, so it missed the
    // new month's achievement entirely ("no active KPI plan found").
    const year = referenceDate.getFullYear();
    const month = referenceDate.getMonth() + 1;
    const weekNumber = monthDayBucket(referenceDate.getDate());

    // pms_kpinamefx is a formula/rollup column that stays null on a freshly
    // created breakdown until Dataverse's async calculation job runs, so
    // filtering on it misses brand-new plans entirely. Resolve the breakdown
    // the same way VisitPlan.tsx does instead: KPI -> its month/year
    // achievements -> this employee's breakdown under one of them.
    const kpiRows = await fetchAll<Strategy_kpises>(
      (skipToken) =>
        Strategy_kpisesService.getAll({
          select: ['strategy_kpisid', 'strategy_newcolumn', '_strategy_function_value', '_strategy_department_value'],
          filter: 'statecode eq 0',
          maxPageSize: 5000,
          skipToken,
        }),
      'Could not load KPIs.'
    );
    const visitsKpi = kpiRows.find(isVisitsKpi);
    if (!visitsKpi) return { status: 'no-kpi', agentName };

    const achievements = await Pm_kpiachievmentsService.getAll({
      select: ['pm_kpiachievmentid'],
      filter: `_pm_kpi_value eq ${visitsKpi.strategy_kpisid} and pm_month eq ${month} and pm_year eq ${year} and statecode eq 0`,
      maxPageSize: 5000,
    });
    const achIds = new Set((achievements.success ? achievements.data : undefined)?.map((a) => a.pm_kpiachievmentid) || []);
    if (!achIds.size) return { status: 'no-achievement', agentName, month, year };

    const employeeBreakdowns = await Stf_kpiachievmentbreakdownsService.getAll({
      select: ['stf_kpiachievmentbreakdownid', '_stf_total_value'],
      filter: `_stf_employee_value eq ${employeeId} and statecode eq 0`,
      maxPageSize: 5000,
    });
    const candidates = (employeeBreakdowns.success ? employeeBreakdowns.data : undefined) || [];
    const breakdownId = candidates.find((b) => b._stf_total_value && achIds.has(b._stf_total_value))?.stf_kpiachievmentbreakdownid;
    if (!breakdownId) return { status: 'no-breakdown', agentName, month, year };

    const weeks = await fetchAll<Cfm_collectionplanweeks>(
      (skipToken) =>
        Cfm_collectionplanweeksService.getAll({
          select: ['cfm_collectionplanweekid', 'cfm_weeknumber', 'cfm_actual'],
          filter: `_cfm_kpiachievementbreakdown_value eq ${breakdownId} and statecode eq 0`,
          maxPageSize: 5000,
          skipToken,
        }),
      'Could not load collection plan weeks.'
    );
    const target = weeks.find((w) => Number(w.cfm_weeknumber) === weekNumber);
    if (!target) return { status: 'no-week', agentName };

    const nextActual = Math.max(0, (Number(target.cfm_actual) || 0) + delta);
    const weekResult = await Cfm_collectionplanweeksService.update(target.cfm_collectionplanweekid, { cfm_actual: nextActual });
    if (!weekResult.success) return { status: 'error', agentName, message: operationErrorMessage(weekResult.error, 'Could not update the week.') };

    // Recompute the breakdown's total from a fresh read of every week — see the
    // same reasoning in VisitPlan.tsx's commitActual: a client-computed running
    // total can undercount if edits race, a fresh read never does.
    const freshWeeks = await fetchAll<Cfm_collectionplanweeks>(
      (skipToken) =>
        Cfm_collectionplanweeksService.getAll({
          select: ['cfm_actual'],
          filter: `_cfm_kpiachievementbreakdown_value eq ${breakdownId} and statecode eq 0`,
          maxPageSize: 5000,
          skipToken,
        }),
      'Could not verify week totals.'
    );
    const totalActual = freshWeeks.reduce((s, w) => s + (Number(w.cfm_actual) || 0), 0);
    await Stf_kpiachievmentbreakdownsService.update(breakdownId, { stf_value: totalActual });
    return { status: 'ok', weekNumber: target.cfm_weeknumber };
  } catch (e) {
    return { status: 'error', agentName, message: e instanceof Error ? e.message : String(e) };
  }
}

function describeAdjustOutcome(o: AdjustOutcome): string | null {
  switch (o.status) {
    case 'ok':
      return null;
    case 'no-employee':
      return `"${o.agentName}" has no linked Employee — open it in Agents and re-select the Employee`;
    case 'no-kpi':
      return `the "No. of Visits" KPI (Collection/Finance) itself wasn't found in Strategy KPIs`;
    case 'no-achievement':
      return `no "No. of Visits" achievement exists for ${o.month}/${o.year} — the KPI plan hasn't been created for that month yet`;
    case 'no-breakdown':
      return `"${o.agentName}" has a "No. of Visits" plan for ${o.month}/${o.year}, but no breakdown row under it linked to this exact employee — check which Employee the plan's owner card was created against`;
    case 'no-week':
      return `no Collection Plan Week found for "${o.agentName}"`;
    case 'bad-date':
      return `invalid visit date`;
    case 'error':
      return `${o.agentName}: ${o.message}`;
  }
}

const TYPE_INSURANCE = 983080002; // cr301_accounttype — Insurance (same value used on the Accounts page)
const TYPE_DIRECT = 766340001; // cr301_accounttype — Direct (same value used on the Accounts page)

const VISIT_TYPE_SINGLE: Cr301_visitscr301_visittype = 983080000; // 'Single Visit'
const VISIT_TYPE_DOUBLE: Cr301_visitscr301_visittype = 983080001; // 'Double Visit'

const VISIT_STATUS_OPEN = 983080000;
const VISIT_STATUS_CANCELLED = 983080002;
const VISIT_STATUS_OPTIONS = [
  { value: VISIT_STATUS_OPEN, label: 'Open' },
  { value: 983080001, label: 'Completed' },
  { value: VISIT_STATUS_CANCELLED, label: 'Cancelled' },
  { value: 766340001, label: 'Partially Closed' },
];
/** Every status except Cancelled counts toward the KPI Actual. */
function countsTowardActual(status: number): boolean {
  return status !== VISIT_STATUS_CANCELLED;
}

/**
 * cfm_tagename, cfm_revisitdate and cfm_agent2 (Task Owner 2) do not exist
 * anywhere in the generated Cr301_visitsModel — this table's Code App data source was set
 * up without these columns. There is no generic-client fallback for a missing COLUMN the
 * way there was for a whole missing TABLE (Accounts' classification lookup): cr301_visits
 * IS a connected data source, so the SDK's retrieveMultipleRecordsAsync/createRecordAsync
 * will actually forward whatever raw string keys are put in `select`/the write body
 * straight to Dataverse — TypeScript just has no typed knowledge of them. These names are
 * carried over from Ref.html / user confirmation; if any is wrong the request fails fast
 * with a clear 400 rather than silently corrupting data.
 *
 * cfm_agent2's logical (attribute) name is confirmed as "cfm_agent2" — safe to select as
 * `_cfm_agent2_value`, since Dataverse attribute/logical names are always lowercase. The
 * `@odata.bind` nav-property casing is NOT confirmed, though: Dataverse preserves whatever
 * casing the maker originally typed for a lookup's schema name (e.g. cr301_salesteam binds
 * via "cr301_SalesTeam", not "cr301_salesteam"), and that can't be derived mechanically.
 * "cfm_Agent2" is used below as the best-justified guess (single word + suffix, same shape
 * as cr301_agent → cr301_Agent) — if writes to it 400, that casing is the first thing to
 * check.
 */
const TAG_COLLECTION = 766340000; // cfm_tagename — "Collection"
const AGENT2_BIND_NAV = 'cfm_Agent2'; // confirmed by the regenerated model ("cfm_Agent2@odata.bind")
const VISITS_TABLE = 'cr301_visits';

/**
 * Bind Task Owner 2 on its own PATCH through the raw client, then re-read
 * _cfm_agent2_value to confirm it stuck. Mixing an extra @odata.bind key into
 * the typed service's create/update body can silently drop it (same issue
 * VisitPlan.tsx's bindLookups works around) — that's how Double visits were
 * ending up saved with no second owner. A 200/204 isn't proof; only the re-read is.
 */
async function bindAgent2(visitId: string, agent2Id: string): Promise<boolean> {
  try {
    const client = getClient(dataSourcesInfo);
    await client.updateRecordAsync(VISITS_TABLE, visitId, { [`${AGENT2_BIND_NAV}@odata.bind`]: `/cr301_agents(${agent2Id})` });
    const check = await Cr301_visitsService.get(visitId, { select: ['_cfm_agent2_value'] });
    return !!check.success && (check.data?._cfm_agent2_value || '').toLowerCase() === agent2Id.toLowerCase();
  } catch {
    return false;
  }
}

interface UnmappedVisitFields {
  cfm_tagename?: number;
  cfm_revisitdate?: string;
  _cfm_agent2_value?: string;
}

interface AccountOption {
  id: string;
  name: string;
  code: string;
  region: 'EG' | 'KSA' | '';
}

interface AgentOption {
  id: string;
  name: string;
  region: 'EG' | 'KSA' | '';
  /** cr301_agents.cfm_employee (_cfm_employee_value) — the real hr_employees
   *  link Agents.tsx now binds on save. Empty for an agent that hasn't been
   *  re-saved since (or was saved before that binding existed), in which case
   *  the KPI-actual adjustment below can't be made — see adjustVisitPlanActual. */
  employeeId: string;
}

interface VisitRow {
  id: string;
  accountId: string;
  company: string;
  code: string;
  agentId: string;
  owner: string;
  agent2Id: string;
  owner2: string;
  date: string;
  createdon: string;
  isDouble: boolean;
  status: number;
  statusLabel: string;
  revisit: string;
  notes: string;
  region: 'EG' | 'KSA' | '';
}

type TypeFilter = '' | 'single' | 'double';

function fmtDate(iso: string): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
}

// True once local calendar "today" has moved past the local calendar day the
// visit was created on — a visit is only editable-by-date on its own creation
// day, then Visit Date locks for good. Local day, not UTC, for the same
// midnight-boundary reason as adjustVisitPlanActual's month/year read.
function isPastCreationDay(createdonIso: string): boolean {
  const created = new Date(createdonIso);
  if (Number.isNaN(created.getTime())) return false;
  const now = new Date();
  return (
    created.getFullYear() !== now.getFullYear() ||
    created.getMonth() !== now.getMonth() ||
    created.getDate() !== now.getDate()
  );
}

interface VisitsPageProps {
  region: Region | null;
}

export default function VisitsPage({ region }: VisitsPageProps) {
  const [visits, setVisits] = useState<VisitRow[]>([]);
  const [accounts, setAccounts] = useState<AccountOption[]>([]);
  const [agents, setAgents] = useState<AgentOption[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [query, setQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState('');
  const [typeFilter, setTypeFilter] = useState<TypeFilter>('');
  // Defaults to the current month; '' means "All Months".
  const [monthFilter, setMonthFilter] = useState(() => {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  });
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(15);

  const [drawerOpen, setDrawerOpen] = useState(false);
  const [editId, setEditId] = useState<string | null>(null);
  // Snapshot of the row being edited, so a saved date/agent change can decrement
  // the KPI week it's moving out of, not just increment the one it's moving into.
  const [editOriginal, setEditOriginal] = useState<VisitRow | null>(null);
  const [formAccountId, setFormAccountId] = useState('');
  const [companyQuery, setCompanyQuery] = useState('');
  const [companyPickerOpen, setCompanyPickerOpen] = useState(false);
  const [formAgentId, setFormAgentId] = useState('');
  const [agentQuery, setAgentQuery] = useState('');
  const [agentPickerOpen, setAgentPickerOpen] = useState(false);
  const [formDate, setFormDate] = useState('');
  const [formStatus, setFormStatus] = useState(String(VISIT_STATUS_OPEN));
  const [formVisitType, setFormVisitType] = useState<'single' | 'double'>('single');
  const [formAgent2Id, setFormAgent2Id] = useState('');
  const [agent2Query, setAgent2Query] = useState('');
  const [agent2PickerOpen, setAgent2PickerOpen] = useState(false);
  const [formRevisit, setFormRevisit] = useState('');
  const [formNotes, setFormNotes] = useState('');
  const [formError, setFormError] = useState('');
  const [saving, setSaving] = useState(false);

  const [toastMsg, setToastMsg] = useState('');
  const [toastOn, setToastOn] = useState(false);
  const toastTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  function showToast(msg: string, durationMs = 2200) {
    setToastMsg(msg);
    setToastOn(true);
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToastOn(false), durationMs);
  }

  async function load() {
    setLoading(true);
    setLoadError(null);
    try {
      const [accResult, agentResult, visitResult] = await Promise.all([
        Cr301_accountsService.getAll({
          select: ['cr301_accountid', 'cr301_acccountname', 'cr301_accountcode', 'cr301_country'],
          filter: `cr301_accounttype eq ${TYPE_INSURANCE} or cr301_accounttype eq ${TYPE_DIRECT}`,
          orderBy: ['cr301_acccountname'],
        }),
        Cr301_agentsService.getAll({
          select: ['cr301_agentid', 'cr301_newcolumn', 'cr18c_area', '_cfm_employee_value'],
          filter: 'cfm_iscollection eq true',
          orderBy: ['cr301_newcolumn'],
        }),
        Cr301_visitsService.getAll({
          select: [
            'cr301_visitid',
            'cr301_visitdate',
            'createdon',
            'cr301_visitstatus',
            'cr301_visittype',
            'cr301_notes',
            '_cr301_account_value',
            '_cr301_agent_value',
            // Not in the generated model — see the module comment above.
            'cfm_tagename',
            'cfm_revisitdate',
            '_cfm_agent2_value',
          ],
          filter: `cfm_tagename eq ${TAG_COLLECTION}`,
          orderBy: ['cr301_visitdate desc'],
        }),
      ]);

      if (!accResult.success) {
        throw new Error(operationErrorMessage(accResult.error, 'Could not load companies.'));
      }
      if (!agentResult.success) {
        throw new Error(operationErrorMessage(agentResult.error, 'Could not load task owners.'));
      }
      if (!visitResult.success) {
        throw new Error(operationErrorMessage(visitResult.error, 'Could not load visits.'));
      }

      const accountOptions: AccountOption[] = (accResult.data || []).map((a) => ({
        id: a.cr301_accountid,
        name: a.cr301_acccountname || '(no name)',
        code: a.cr301_accountcode || '',
        region: a.cr301_country == null ? '' : Number(a.cr301_country) === ACCOUNT_REGION_KSA ? 'KSA' : 'EG',
      }));
      const agentOptions: AgentOption[] = (agentResult.data || []).map((a) => {
        const areaValue = a.cr18c_area == null ? null : Number(a.cr18c_area);
        return {
          id: a.cr301_agentid,
          name: a.cr301_newcolumn || '(agent)',
          region: areaValue == null ? '' : areaValue === AREA_KSA ? 'KSA' : 'EG',
          employeeId: a._cfm_employee_value || '',
        };
      });
      const accMap = new Map(accountOptions.map((a) => [a.id, a]));
      const agentMap = new Map(agentOptions.map((a) => [a.id, a]));

      const rows: VisitRow[] = (visitResult.data || []).map((r) => {
        const extra = r as unknown as UnmappedVisitFields;
        const accountId = r._cr301_account_value || '';
        const agentId = r._cr301_agent_value || '';
        const agent2Id = extra._cfm_agent2_value || '';
        const acc = accMap.get(accountId);
        const agent = agentMap.get(agentId);
        const agent2 = agent2Id ? agentMap.get(agent2Id) : undefined;
        return {
          id: r.cr301_visitid,
          accountId,
          company: acc?.name || '—',
          code: acc?.code || '',
          agentId,
          owner: agent?.name || '—',
          agent2Id,
          // Formatted-value fallback: Task Owner 2 may be an agent outside the
          // collection roster loaded above, which used to show as "—" even though it's saved.
          owner2: agent2?.name || (agent2Id ? fmtVal(r, '_cfm_agent2_value') || '(agent)' : ''),
          date: r.cr301_visitdate || '',
          createdon: r.createdon || '',
          isDouble: r.cr301_visittype === VISIT_TYPE_DOUBLE,
          status: r.cr301_visitstatus ?? VISIT_STATUS_OPEN,
          statusLabel: Cr301_visitscr301_visitstatus[r.cr301_visitstatus ?? VISIT_STATUS_OPEN] || '—',
          revisit: extra.cfm_revisitdate || '',
          notes: r.cr301_notes || '',
          region: agent?.region || '',
        };
      });

      setAccounts(accountOptions);
      setAgents(agentOptions);
      setVisits(rows);
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
    const seen = new Map<string, string>();
    visits.forEach((v) => {
      if (!v.date) return;
      const d = new Date(v.date);
      if (Number.isNaN(d.getTime())) return;
      const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
      if (!seen.has(key)) {
        seen.set(key, d.toLocaleDateString(undefined, { year: 'numeric', month: 'long' }));
      }
    });
    // Always offer the current month so the default selection shows in the dropdown.
    const now = new Date();
    const curKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    if (!seen.has(curKey)) seen.set(curKey, now.toLocaleDateString(undefined, { year: 'numeric', month: 'long' }));
    return Array.from(seen.entries())
      .sort((a, b) => b[0].localeCompare(a[0]))
      .map(([value, label]) => ({ value, label }));
  }, [visits]);

  const view = useMemo(() => {
    const q = query.trim().toLowerCase();
    const matches = visits.filter((v) => {
      if (q && !`${v.owner} ${v.owner2} ${v.company} ${v.code} ${v.notes}`.toLowerCase().includes(q)) return false;
      if (statusFilter && String(v.status) !== statusFilter) return false;
      if (typeFilter === 'single' && v.isDouble) return false;
      if (typeFilter === 'double' && !v.isDouble) return false;
      if (monthFilter && v.date.slice(0, 7) !== monthFilter) return false;
      if (dateFrom && (!v.date || v.date.slice(0, 10) < dateFrom)) return false;
      if (dateTo && (!v.date || v.date.slice(0, 10) > dateTo)) return false;
      // App-wide Region toggle (Layout header) — unset region on a visit's
      // owner counts as Egypt, same "everything else is Egypt" convention
      // useRegion.ts uses for BUs.
      if (region && (region === 'KSA' ? v.region !== 'KSA' : v.region === 'KSA')) return false;
      return true;
    });
    return [...matches].sort((a, b) => b.createdon.localeCompare(a.createdon));
  }, [visits, query, statusFilter, typeFilter, monthFilter, dateFrom, dateTo, region]);

  // Add/Edit Visit drawer's Company Name / Task Owner pickers — same App-wide
  // Region toggle convention as the table above (unset region counts as Egypt).
  const regionAccountOptions = useMemo(
    () => (!region ? accounts : accounts.filter((a) => (region === 'KSA' ? a.region === 'KSA' : a.region !== 'KSA'))),
    [accounts, region]
  );
  const regionAgentOptions = useMemo(
    () => (!region ? agents : agents.filter((a) => (region === 'KSA' ? a.region === 'KSA' : a.region !== 'KSA'))),
    [agents, region]
  );

  /** Searchable Company Name picker — same style/behavior as Task Owner below. */
  function companyMatches(q: string): AccountOption[] {
    const query = q.trim().toLowerCase();
    if (!query) return regionAccountOptions;
    return regionAccountOptions.filter((a) => `${a.name} ${a.code}`.toLowerCase().includes(query));
  }
  function pickCompany(a: AccountOption) {
    setFormAccountId(a.id);
    setCompanyQuery(a.name);
    setCompanyPickerOpen(false);
  }

  /** Searchable Task Owner / Task Owner 2 pickers — shows the full (region-
   * scoped) roster right away, filtering as the user types. */
  function agentMatches(q: string): AgentOption[] {
    const query = q.trim().toLowerCase();
    if (!query) return regionAgentOptions;
    return regionAgentOptions.filter((a) => a.name.toLowerCase().includes(query));
  }
  function pickAgent(a: AgentOption) {
    setFormAgentId(a.id);
    setAgentQuery(a.name);
    setAgentPickerOpen(false);
  }
  function pickAgent2(a: AgentOption) {
    setFormAgent2Id(a.id);
    setAgent2Query(a.name);
    setAgent2PickerOpen(false);
  }

  useEffect(() => {
    setPage(1);
  }, [query, statusFilter, typeFilter, monthFilter, dateFrom, dateTo, region]);

  const totalPages = Math.max(1, Math.ceil(view.length / pageSize));
  const clampedPage = Math.min(page, totalPages);
  const pageRows = view.slice((clampedPage - 1) * pageSize, (clampedPage - 1) * pageSize + pageSize);

  const kpis = useMemo(() => {
    const uniqueCompanies = new Set(view.filter((v) => v.accountId).map((v) => v.accountId)).size;
    const open = view.filter((v) => v.status === VISIT_STATUS_OPEN).length;
    const completed = view.filter((v) => v.status === 983080001).length;
    return { total: view.length, uniqueCompanies, open, completed };
  }, [view]);

  const selectedAccount = useMemo(() => accounts.find((a) => a.id === formAccountId) || null, [accounts, formAccountId]);
  // Visit Date stays editable on the day the visit was created, then locks —
  // never locked for a brand-new (not-yet-saved) visit.
  const visitDateLocked = !!editId && !!editOriginal?.createdon && isPastCreationDay(editOriginal.createdon);
  // Task Owner 2 locks once saved, but stays pickable on a Double visit that
  // was saved without one, so those records can be repaired.
  const agent2Locked = !!editId && !!editOriginal?.agent2Id;

  function openDrawer(row?: VisitRow) {
    setFormError('');
    setEditOriginal(row || null);
    if (row) {
      setEditId(row.id);
      setFormAccountId(row.accountId);
      setCompanyQuery(row.company);
      setFormAgentId(row.agentId);
      setAgentQuery(row.owner);
      setFormDate(row.date ? row.date.substring(0, 10) : '');
      setFormStatus(String(row.status));
      setFormVisitType(row.isDouble ? 'double' : 'single');
      setFormAgent2Id(row.agent2Id);
      setAgent2Query(row.owner2);
      setFormRevisit(row.revisit ? row.revisit.substring(0, 10) : '');
      setFormNotes(row.notes);
    } else {
      setEditId(null);
      setFormAccountId('');
      setCompanyQuery('');
      setFormAgentId('');
      setAgentQuery('');
      setFormDate(new Date().toISOString().substring(0, 10));
      setFormStatus(String(VISIT_STATUS_OPEN));
      setFormVisitType('single');
      setFormAgent2Id('');
      setAgent2Query('');
      setFormRevisit('');
      setFormNotes('');
    }
    setCompanyPickerOpen(false);
    setAgentPickerOpen(false);
    setAgent2PickerOpen(false);
    setDrawerOpen(true);
  }

  function closeDrawer() {
    setDrawerOpen(false);
  }

  async function handleSave() {
    if (!formAccountId) {
      setFormError('Company is required.');
      return;
    }
    if (!formAgentId) {
      setFormError('Task Owner is required.');
      return;
    }
    const isDouble = formVisitType === 'double';
    if (isDouble && !formAgent2Id) {
      setFormError('Task Owner 2 is required for a double visit.');
      return;
    }
    if (isDouble && formAgent2Id && formAgentId === formAgent2Id) {
      setFormError('Task Owner 2 must be different from Task Owner.');
      return;
    }
    if (parseInt(formStatus, 10) !== VISIT_STATUS_OPEN && !formDate) {
      setFormError('Visit Date is required once the status is no longer Open.');
      return;
    }

    setSaving(true);
    setFormError('');
    try {
      const body: Record<string, unknown> = {
        'cr301_Account@odata.bind': `/cr301_accounts(${formAccountId})`,
        'cr301_Agent@odata.bind': `/cr301_agents(${formAgentId})`,
        cr301_visittype: isDouble ? VISIT_TYPE_DOUBLE : VISIT_TYPE_SINGLE,
        cr301_visitdate: formDate ? new Date(formDate).toISOString() : null,
        // Planned Visit Date has no field of its own in this form — it just mirrors
        // Visit Date on every save, per user confirmation.
        cr301_plannedvisitdate: formDate ? new Date(formDate).toISOString() : null,
        cr301_visitstatus: parseInt(formStatus, 10),
        cr301_notes: formNotes.trim() || null,
        // Not in the generated model — see the module comment above.
        cfm_tagename: TAG_COLLECTION,
        cfm_revisitdate: formRevisit ? new Date(formRevisit).toISOString() : null,
      };
      // Task Owner 2 is NOT put in this body — it's bound separately below
      // (see bindAgent2). Visit Type can't change on edit, so there's no unlink case.

      let result;
      if (editId) {
        result = await Cr301_visitsService.update(editId, body as Partial<Omit<Cr301_visitsBase, 'cr301_visitid'>>);
      } else {
        result = await Cr301_visitsService.create(body as unknown as Omit<Cr301_visitsBase, 'cr301_visitid'>);
      }
      if (!result.success) {
        throw new Error(operationErrorMessage(result.error, 'Save failed.'));
      }

      const visitId = editId || result.data?.cr301_visitid || '';
      let agent2Saved = true;
      if (isDouble && formAgent2Id && formAgent2Id !== editOriginal?.agent2Id) {
        agent2Saved = !!visitId && (await bindAgent2(visitId, formAgent2Id));
      }

      // Four directions now, not two: creating an (active) visit increments its
      // week's Actual; editing its task owner(s) while it stays active moves
      // the count across owners; and — the piece that was missing — flipping
      // Visit Status to/from Cancelled gives the count back / re-adds it,
      // since a cancelled visit never really happened. The reference date is
      // the visit's own createdon (system timestamp), not the user-entered
      // Visit Date — see adjustVisitPlanActual's doc comment for why, and it's
      // immutable, so editing an existing visit can only ever change WHICH
      // owner(s)/whether it counts, never WHICH week. A failure/no-match here
      // never undoes the fact that the visit itself already saved fine — the
      // outcome is surfaced in the toast, not swallowed.
      //
      // Adjustments run ONE AT A TIME (not Promise.all) — a Double Visit's two
      // owners can end up resolving to the same employee (e.g. two agents
      // mistakenly linked to the same hr_employees record), and running their
      // adjustVisitPlanActual calls concurrently is a lost-update race: both
      // read the week's current cfm_actual before either writes back, so two
      // +1's land as only +1 overall. Sequential awaits close that window —
      // each call's read only ever sees the previous call's write.
      const adjustNotes: string[] = [];
      const referenceDateIso = editId ? editOriginal?.createdon || '' : result.data?.createdon || '';
      if (referenceDateIso) {
        const toAdjustTargets = (ids: string[]) =>
          ids
            .filter(Boolean)
            .map((id) => agents.find((a) => a.id === id))
            .filter((a): a is AgentOption => !!a);
        const applyAdjustments = async (targets: AgentOption[], delta: 1 | -1): Promise<AdjustOutcome[]> => {
          const results: AdjustOutcome[] = [];
          for (const a of targets) {
            results.push(await adjustVisitPlanActual(a.name, a.employeeId, referenceDateIso, delta));
          }
          return results;
        };
        // Only count Task Owner 2 if the link actually landed — otherwise the
        // KPI would say they did a visit the record doesn't show.
        const newAgents = toAdjustTargets([formAgentId, isDouble && agent2Saved ? formAgent2Id : '']);
        const newCounts = countsTowardActual(parseInt(formStatus, 10));
        const outcomes: AdjustOutcome[] = [];
        if (!editId) {
          // A visit created directly as Cancelled never counted in the first place.
          if (newCounts) {
            outcomes.push(...(await applyAdjustments(newAgents, 1)));
          }
        } else if (editOriginal) {
          const oldAgents = toAdjustTargets([editOriginal.agentId, editOriginal.isDouble ? editOriginal.agent2Id : '']);
          const oldCounts = countsTowardActual(editOriginal.status);
          const sameAgents =
            JSON.stringify([...oldAgents.map((a) => a.id)].sort()) === JSON.stringify([...newAgents.map((a) => a.id)].sort());
          if (oldCounts && !newCounts) {
            // Just cancelled — give back whatever it was counted toward.
            outcomes.push(...(await applyAdjustments(oldAgents, -1)));
          } else if (!oldCounts && newCounts) {
            // Un-cancelled — count it now, toward its current owner(s).
            outcomes.push(...(await applyAdjustments(newAgents, 1)));
          } else if (oldCounts && newCounts && !sameAgents) {
            // Stayed active/counted, owners changed — move the count across owners.
            outcomes.push(...(await applyAdjustments(oldAgents, -1)));
            outcomes.push(...(await applyAdjustments(newAgents, 1)));
          }
          // else: stayed Cancelled, or stayed active/counted with the same owners — no change.
        }
        adjustNotes.push(...outcomes.map(describeAdjustOutcome).filter((n): n is string => n !== null));
      }

      if (!agent2Saved) {
        adjustNotes.unshift('Task Owner 2 could not be linked — open the visit and pick Task Owner 2 again');
      }

      setDrawerOpen(false);
      if (adjustNotes.length) {
        showToast(`${editId ? 'Visit updated' : 'Visit recorded'} — KPI actual not updated (${adjustNotes.join('; ')}).`, 6000);
      } else {
        showToast(editId ? 'Visit updated' : 'Visit recorded');
      }
      await load();
    } catch (e) {
      setFormError(e instanceof Error ? e.message : String(e));
    } finally {
      setSaving(false);
    }
  }

  function exportCsv() {
    const header = ['Task Owner', 'Task Owner 2', 'Company', 'Company Code', 'Visit Date', 'Visit Status', 'Visit Type', 'Revisit Date', 'Comments', 'Created On', 'Region'];
    const rows = view.map((v) => [
      v.owner,
      v.owner2,
      v.company,
      v.code,
      fmtDate(v.date),
      v.statusLabel,
      v.isDouble ? 'Double' : 'Single',
      fmtDate(v.revisit),
      v.notes,
      fmtDate(v.createdon),
      v.region === 'KSA' ? 'KSA' : v.region === 'EG' ? 'Egypt' : '—',
    ]);
    downloadCsv('visits_log.csv', [header, ...rows].map((r) => r.map(escapeForCsv).join(',')));
  }

  function pagerButtons() {
    const pages: (number | 'ellipsis')[] = [];
    for (let p = 1; p <= totalPages; p++) {
      if (p === 1 || p === totalPages || Math.abs(p - clampedPage) <= 1) {
        pages.push(p);
      } else if (Math.abs(p - clampedPage) === 2) {
        pages.push('ellipsis');
      }
    }
    return pages;
  }

  return (
    <>
      <div className="acc-wrap">
        <div className="acc-head">
          <div>
            <h1 className="acc-title">Visits Log</h1>
            <div className="acc-sub">Recorded visits — Task Owner · Company · Visit Date · Visit Status · Visit Type · Revisit · Comments</div>
          </div>
          <div className="acc-actions">
            <button className="acc-btn acc-btn-primary" onClick={() => openDrawer()}>
              <Icon className="fa-solid fa-plus" /> Add Visit
            </button>
            <button className="acc-btn" onClick={exportCsv}>
              <Icon className="fa-solid fa-download" /> Export
            </button>
          </div>
        </div>

        <div className="kpi-grid">
          <div className="kpi-card">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-gold">
                <Icon className="fa-solid fa-clipboard-check" />
              </span>
              <span className="kpi-label">Total visits</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.total}</div>
            <div className="kpi-context">&nbsp;</div>
          </div>
          <div className="kpi-card">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-gold">
                <Icon className="fa-solid fa-building" />
              </span>
              <span className="kpi-label">Unique companies</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.uniqueCompanies}</div>
            <div className="kpi-context">&nbsp;</div>
          </div>
          <div className="kpi-card kpi-warn">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-warn">
                <Icon className="fa-solid fa-door-open" />
              </span>
              <span className="kpi-label">Open</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.open}</div>
            <div className="kpi-context">
              {loading ? '' : kpis.total > 0 ? `${Math.round((kpis.open / kpis.total) * 100)}% of visits` : '—'}
            </div>
          </div>
          <div className="kpi-card kpi-ok">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-ok">
                <Icon className="fa-solid fa-circle-check" />
              </span>
              <span className="kpi-label">Completed</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.completed}</div>
            <div className="kpi-context">
              {loading ? '' : kpis.total > 0 ? `${Math.round((kpis.completed / kpis.total) * 100)}% of visits` : '—'}
            </div>
          </div>
        </div>

        <div className="tbl-shell">
          <div className="tbl-hdr">
            <div className="acc-search">
              <Icon className="fa-solid fa-magnifying-glass" />
              <input placeholder="Search owner, company or notes…" value={query} onChange={(e) => setQuery(e.target.value)} />
            </div>
            <select className="acc-f" value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
              <option value="">All Statuses</option>
              {VISIT_STATUS_OPTIONS.map((s) => (
                <option key={s.value} value={s.value}>
                  {s.label}
                </option>
              ))}
            </select>
            <select className="acc-f" value={typeFilter} onChange={(e) => setTypeFilter(e.target.value as TypeFilter)}>
              <option value="">All Types</option>
              <option value="single">Single visit</option>
              <option value="double">Double visit</option>
            </select>
            <select className="acc-f" value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)}>
              <option value="">All Months</option>
              {monthOptions.map((m) => (
                <option key={m.value} value={m.value}>
                  {m.label}
                </option>
              ))}
            </select>
            <div className="acc-date-range">
              <input type="date" className="acc-f" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)} title="Visit date from" />
              <span>–</span>
              <input type="date" className="acc-f" value={dateTo} onChange={(e) => setDateTo(e.target.value)} title="Visit date to" />
            </div>
          </div>

          <div className="acc-tablewrap">
          <table>
            <thead>
              <tr>
                <th>Task Owner</th>
                <th>Task Owner 2</th>
                <th>Company</th>
                <th>Visit Date</th>
                <th>Visit Status</th>
                <th>Type</th>
                <th>Revisit</th>
                <th>Comments</th>
                <th>Created On</th>
                <th>Region</th>
              </tr>
            </thead>
            <tbody>
              {loading ? (
                Array.from({ length: 6 }).map((_, i) => (
                  <tr key={i}>
                    {Array.from({ length: 10 }).map((__, j) => (
                      <td key={j}>
                        <div className="acc-skel"></div>
                      </td>
                    ))}
                  </tr>
                ))
              ) : loadError ? (
                <tr>
                  <td colSpan={10}>
                    <div className="acc-state">
                      <Icon className="fa-solid fa-triangle-exclamation" />
                      Could not load visits.
                      <br />
                      <small>{loadError}</small>
                    </div>
                  </td>
                </tr>
              ) : pageRows.length === 0 ? (
                <tr>
                  <td colSpan={10}>
                    <div className="acc-state">
                      <Icon className="fa-solid fa-inbox" />
                      No visits match your filters.
                    </div>
                  </td>
                </tr>
              ) : (
                pageRows.map((v) => {
                  return (
                    <tr key={v.id} className="acc-row" onClick={() => openDrawer(v)}>
                      <td className="acc-name">{v.owner}</td>
                      <td>{v.owner2 ? <span className="acc-name">{v.owner2}</span> : <span className="vis-muted">—</span>}</td>
                      <td>
                        <div className="acc-name">{v.company}</div>
                        {v.code && <div className="acc-code">{v.code}</div>}
                      </td>
                      <td>{v.date ? fmtDate(v.date) : <span className="vis-muted">—</span>}</td>
                      <td>
                        <span className={`acc-visit-status ${visitStatusClass(v.statusLabel)}`}>{v.statusLabel}</span>
                      </td>
                      <td>
                        {v.isDouble ? (
                          <span className="vis-vt vis-vt-double">
                            <Icon className="fa-solid fa-users" /> Double
                          </span>
                        ) : (
                          <span className="vis-vt">
                            <Icon className="fa-solid fa-user" /> Single
                          </span>
                        )}
                      </td>
                      <td>{v.revisit ? fmtDate(v.revisit) : <span className="vis-muted">—</span>}</td>
                      <td>
                        <div className="vis-notes">{v.notes || <span className="vis-muted">—</span>}</div>
                      </td>
                      <td>{v.createdon ? fmtDate(v.createdon) : <span className="vis-muted">—</span>}</td>
                      <td>{v.region ? (v.region === 'KSA' ? 'KSA' : 'Egypt') : <span className="vis-muted">—</span>}</td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
          <div className="acc-foot">
            <span>
              {view.length} visit{view.length !== 1 ? 's' : ''}
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
        </div>
        </div>
      </div>

      <div className={`acc-scrim${drawerOpen ? ' on' : ''}`} onClick={closeDrawer}></div>
      <aside className={`acc-drawer${drawerOpen ? ' on' : ''}`}>
        <div className="acc-dhead">
          <h2>
            <Icon className={`fa-solid ${editId ? 'fa-pen' : 'fa-plus'}`} style={{ color: 'var(--gold)', fontSize: editId ? 15 : 16 }} />
            {editId ? 'Edit Visit' : 'Add Visit'}
          </h2>
          <button className="acc-dclose" onClick={closeDrawer}>
            <Icon className="fa-solid fa-xmark" />
          </button>
        </div>
        <div className="acc-dbody">
          <div className={`acc-err${formError ? ' on' : ''}`}>{formError}</div>

          <div className="acc-fld">
            <label>Company Name</label>
            <div className="acc-combo">
              <div className="acc-search">
                <Icon className="fa-solid fa-magnifying-glass" />
                <input
                  placeholder="Search company or code…"
                  autoComplete="off"
                  disabled={!!editId}
                  value={companyQuery}
                  onChange={(e) => {
                    setCompanyQuery(e.target.value);
                    setFormAccountId('');
                    setCompanyPickerOpen(true);
                  }}
                  onFocus={() => setCompanyPickerOpen(true)}
                  onBlur={() => setTimeout(() => setCompanyPickerOpen(false), 150)}
                />
              </div>
              {companyPickerOpen && !editId && (
                <div className="acc-combo-list">
                  {companyMatches(companyQuery).length === 0 ? (
                    <div className="acc-combo-empty">No matches.</div>
                  ) : (
                    companyMatches(companyQuery).map((a) => (
                      <div key={a.id} className="acc-combo-item" onMouseDown={(ev) => ev.preventDefault()} onClick={() => pickCompany(a)}>
                        {a.name}
                      </div>
                    ))
                  )}
                </div>
              )}
            </div>
          </div>

          <div className="acc-fld">
            <div className="acc-ro-wrap">
              <div className="acc-ro-label">Company Code</div>
              <div className="acc-ro-value">{selectedAccount?.code || '—'}</div>
            </div>
          </div>

          <div className="acc-fld">
            <label>Task Owner</label>
            <div className="acc-combo">
              <div className="acc-search">
                <Icon className="fa-solid fa-magnifying-glass" />
                <input
                  placeholder="Search users…"
                  autoComplete="off"
                  disabled={!!editId}
                  value={agentQuery}
                  onChange={(e) => {
                    setAgentQuery(e.target.value);
                    setFormAgentId('');
                    setAgentPickerOpen(true);
                  }}
                  onFocus={() => setAgentPickerOpen(true)}
                  onBlur={() => setTimeout(() => setAgentPickerOpen(false), 150)}
                />
              </div>
              {agentPickerOpen && !editId && (
                <div className="acc-combo-list">
                  {agentMatches(agentQuery).length === 0 ? (
                    <div className="acc-combo-empty">No matches.</div>
                  ) : (
                    agentMatches(agentQuery).map((a) => (
                      <div key={a.id} className="acc-combo-item" onMouseDown={(ev) => ev.preventDefault()} onClick={() => pickAgent(a)}>
                        {a.name}
                      </div>
                    ))
                  )}
                </div>
              )}
            </div>
          </div>

          <div className="acc-fld">
            <label>Visit Date</label>
            <input type="date" value={formDate} disabled={visitDateLocked} onChange={(e) => setFormDate(e.target.value)} />
          </div>

          <div className="acc-fld">
            <label>Visit Status</label>
            <select value={formStatus} onChange={(e) => setFormStatus(e.target.value)}>
              {VISIT_STATUS_OPTIONS.map((s) => (
                <option key={s.value} value={s.value}>
                  {s.label}
                </option>
              ))}
            </select>
          </div>

          <div className="acc-fld">
            <label>Visit Type</label>
            <select value={formVisitType} disabled={!!editId} onChange={(e) => setFormVisitType(e.target.value as 'single' | 'double')}>
              <option value="single">Single visit</option>
              <option value="double">Double visit</option>
            </select>
          </div>

          {formVisitType === 'double' && (
            <div className="acc-fld">
              <label>Task Owner 2</label>
              <div className="acc-combo">
                <div className="acc-search">
                  <Icon className="fa-solid fa-magnifying-glass" />
                  <input
                    placeholder="Search users…"
                    autoComplete="off"
                    disabled={agent2Locked}
                    value={agent2Query}
                    onChange={(e) => {
                      setAgent2Query(e.target.value);
                      setFormAgent2Id('');
                      setAgent2PickerOpen(true);
                    }}
                    onFocus={() => setAgent2PickerOpen(true)}
                    onBlur={() => setTimeout(() => setAgent2PickerOpen(false), 150)}
                  />
                </div>
                {agent2PickerOpen && !agent2Locked && (
                  <div className="acc-combo-list">
                    {agentMatches(agent2Query).length === 0 ? (
                      <div className="acc-combo-empty">No matches.</div>
                    ) : (
                      agentMatches(agent2Query).map((a) => (
                        <div key={a.id} className="acc-combo-item" onMouseDown={(ev) => ev.preventDefault()} onClick={() => pickAgent2(a)}>
                          {a.name}
                        </div>
                      ))
                    )}
                  </div>
                )}
              </div>
              <div className="acc-hint">
                The visit is counted as achieved for <b>both</b> task owners.
              </div>
            </div>
          )}

          <div className="acc-fld">
            <label>
              Revisit Date <span className="vis-muted">(optional)</span>
            </label>
            <input type="date" value={formRevisit} onChange={(e) => setFormRevisit(e.target.value)} />
          </div>

          <div className="acc-fld">
            <label>Comments</label>
            <textarea
              placeholder="What was discussed / agreed / next steps…"
              value={formNotes}
              onChange={(e) => setFormNotes(e.target.value)}
            />
          </div>
        </div>
        <div className="acc-dfoot">
          <button className="acc-btn acc-btn-ghost" onClick={closeDrawer}>
            Cancel
          </button>
          <button className="acc-btn acc-btn-primary" disabled={saving} onClick={handleSave}>
            {saving ? (
              <>
                <Icon className="fa-solid fa-spinner fa-spin" /> Saving…
              </>
            ) : (
              'Save Visit'
            )}
          </button>
        </div>
      </aside>

      <div className={`acc-toast${toastOn ? ' on' : ''}`}>{toastMsg}</div>
    </>
  );
}
