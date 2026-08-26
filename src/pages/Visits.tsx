import { useEffect, useMemo, useRef, useState } from 'react';
import { Cr301_visitsService } from '../generated/services/Cr301_visitsService';
import { Cr301_accountsService } from '../generated/services/Cr301_accountsService';
import { Cr301_agentsService } from '../generated/services/Cr301_agentsService';
import { Hr_employeesService } from '../generated/services/Hr_employeesService';
import { Stf_kpiachievmentbreakdownsService } from '../generated/services/Stf_kpiachievmentbreakdownsService';
import { Cfm_collectionplanweeksService } from '../generated/services/Cfm_collectionplanweeksService';
import { Cr301_visitscr301_visitstatus } from '../generated/models/Cr301_visitsModel';
import type { Cr301_visitsBase, Cr301_visitscr301_visittype } from '../generated/models/Cr301_visitsModel';
import type { Cfm_collectionplanweeks } from '../generated/models/Cfm_collectionplanweeksModel';
import type { Hr_employees } from '../generated/models/Hr_employeesModel';
import { operationErrorMessage } from '../lib/errors';
import { fetchAll } from '../lib/dataverse';
import { visitStatusClass } from '../lib/visitStatus';
import './Visits.css';

const VISIT_KPI_NAME = 'No. of Visits';

// First 3 words of a name, normalized — same tolerant match used on the Agents
// page for Employee <-> System User, reused here for Agent <-> Employee since
// Arabic names commonly carry 3-4+ parts and HR vs. agent records don't always
// agree on how many are kept.
function firstWords(name: string, n: number): string {
  return name.trim().toLowerCase().split(/\s+/).slice(0, n).join(' ');
}

type AdjustOutcome =
  | { status: 'ok'; weekNumber?: number }
  | { status: 'no-employee' | 'no-breakdown' | 'no-week' | 'bad-date'; agentName: string }
  | { status: 'error'; agentName: string; message: string };

/**
 * For an agent + a reference date (the Visit's createdon), find that agent's
 * matching hr_employees record (by name), their active "No. of Visits" KPI
 * achievement breakdown for that date's year, and the Collection Plan Week
 * under that breakdown whose cfm_weeknumber matches the reference date's
 * day-of-month bucket (day 1-7 -> week 1, 8-14 -> week 2, 15-21 -> week 3,
 * ...) — then adjust that week's cfm_actual by `delta` (+1 for a
 * created/moved-in visit, -1 for a removed/moved-out one, clamped at 0).
 *
 * Deliberately does NOT compare against cfm_weekfrom/cfm_weekto: those aren't
 * reliable real calendar boundaries (the create wizard in VisitPlan.tsx just
 * stamps them as sequential 7-day ranges from whenever the plan happened to be
 * created, not from a fixed month start), so date-range matching against them
 * can miss every week. cfm_weeknumber is the one plan-week-count-agnostic
 * anchor both this code and the plan's own week cards agree on.
 *
 * Mirrors what a Power Automate flow on cr301_visits would do, run inline
 * instead. Returns which step it stopped at rather than throwing, so the
 * caller can tell the user exactly why nothing was adjusted instead of
 * failing silently.
 */
async function adjustVisitPlanActual(agentName: string, referenceDateIso: string, delta: 1 | -1): Promise<AdjustOutcome> {
  try {
    const employeeRows = await fetchAll<Hr_employees>(
      (skipToken) => Hr_employeesService.getAll({ select: ['hr_employeeid', 'hr_fullname'], filter: 'statecode eq 0', maxPageSize: 5000, skipToken }),
      'Could not load employees.'
    );
    const agentKey = firstWords(agentName, 3);
    const employee = agentKey ? employeeRows.find((e) => firstWords(e.hr_fullname || '', 3) === agentKey) : undefined;
    if (!employee) return { status: 'no-employee', agentName };

    const referenceDate = new Date(referenceDateIso);
    if (Number.isNaN(referenceDate.getTime())) return { status: 'bad-date', agentName };
    const year = referenceDate.getUTCFullYear();
    const weekNumber = Math.ceil(referenceDate.getUTCDate() / 7);

    const breakdowns = await Stf_kpiachievmentbreakdownsService.getAll({
      select: ['stf_kpiachievmentbreakdownid'],
      filter: `_stf_employee_value eq ${employee.hr_employeeid} and pms_year eq ${year} and pms_kpinamefx eq '${VISIT_KPI_NAME}' and statecode eq 0`,
      orderBy: ['createdon desc'],
      top: 1,
    });
    const breakdownId = breakdowns.success ? breakdowns.data?.[0]?.stf_kpiachievmentbreakdownid : undefined;
    if (!breakdownId) return { status: 'no-breakdown', agentName };

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
      return `no employee match for "${o.agentName}"`;
    case 'no-breakdown':
      return `no active "No. of Visits" KPI plan found for "${o.agentName}" this year`;
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
const VISIT_STATUS_OPTIONS = [
  { value: VISIT_STATUS_OPEN, label: 'Open' },
  { value: 983080001, label: 'Completed' },
  { value: 983080002, label: 'Cancelled' },
  { value: 766340001, label: 'Partially Closed' },
];

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
const AGENT2_BIND_NAV = 'cfm_Agent2'; // unconfirmed casing — see comment above

interface UnmappedVisitFields {
  cfm_tagename?: number;
  cfm_revisitdate?: string;
  _cfm_agent2_value?: string;
}

interface AccountOption {
  id: string;
  name: string;
  code: string;
}

interface AgentOption {
  id: string;
  name: string;
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
}

type TypeFilter = '' | 'single' | 'double';

function fmtDate(iso: string): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return iso;
  return d.toLocaleDateString(undefined, { year: 'numeric', month: 'short', day: 'numeric' });
}

export default function VisitsPage() {
  const [visits, setVisits] = useState<VisitRow[]>([]);
  const [accounts, setAccounts] = useState<AccountOption[]>([]);
  const [agents, setAgents] = useState<AgentOption[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [query, setQuery] = useState('');
  const [statusFilter, setStatusFilter] = useState('');
  const [typeFilter, setTypeFilter] = useState<TypeFilter>('');
  const [monthFilter, setMonthFilter] = useState('');
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
  const [formAgentId, setFormAgentId] = useState('');
  const [formDate, setFormDate] = useState('');
  const [formStatus, setFormStatus] = useState(String(VISIT_STATUS_OPEN));
  const [formVisitType, setFormVisitType] = useState<'single' | 'double'>('single');
  const [formAgent2Id, setFormAgent2Id] = useState('');
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
          select: ['cr301_accountid', 'cr301_acccountname', 'cr301_accountcode'],
          filter: `cr301_accounttype eq ${TYPE_INSURANCE} or cr301_accounttype eq ${TYPE_DIRECT}`,
          orderBy: ['cr301_acccountname'],
        }),
        Cr301_agentsService.getAll({
          select: ['cr301_agentid', 'cr301_newcolumn'],
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
      }));
      const agentOptions: AgentOption[] = (agentResult.data || []).map((a) => ({
        id: a.cr301_agentid,
        name: a.cr301_newcolumn || '(agent)',
      }));
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
          owner2: agent2?.name || '',
          date: r.cr301_visitdate || '',
          createdon: r.createdon || '',
          isDouble: r.cr301_visittype === VISIT_TYPE_DOUBLE,
          status: r.cr301_visitstatus ?? VISIT_STATUS_OPEN,
          statusLabel: Cr301_visitscr301_visitstatus[r.cr301_visitstatus ?? VISIT_STATUS_OPEN] || '—',
          revisit: extra.cfm_revisitdate || '',
          notes: r.cr301_notes || '',
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
    return Array.from(seen.entries())
      .sort((a, b) => b[0].localeCompare(a[0]))
      .map(([value, label]) => ({ value, label }));
  }, [visits]);

  const view = useMemo(() => {
    const q = query.trim().toLowerCase();
    return visits.filter((v) => {
      if (q && !`${v.owner} ${v.owner2} ${v.company} ${v.code} ${v.notes}`.toLowerCase().includes(q)) return false;
      if (statusFilter && String(v.status) !== statusFilter) return false;
      if (typeFilter === 'single' && v.isDouble) return false;
      if (typeFilter === 'double' && !v.isDouble) return false;
      if (monthFilter && v.date.slice(0, 7) !== monthFilter) return false;
      if (dateFrom && (!v.date || v.date.slice(0, 10) < dateFrom)) return false;
      if (dateTo && (!v.date || v.date.slice(0, 10) > dateTo)) return false;
      return true;
    });
  }, [visits, query, statusFilter, typeFilter, monthFilter, dateFrom, dateTo]);

  useEffect(() => {
    setPage(1);
  }, [query, statusFilter, typeFilter, monthFilter, dateFrom, dateTo]);

  const totalPages = Math.max(1, Math.ceil(view.length / pageSize));
  const clampedPage = Math.min(page, totalPages);
  const pageRows = view.slice((clampedPage - 1) * pageSize, (clampedPage - 1) * pageSize + pageSize);

  const kpis = useMemo(() => {
    const uniqueCompanies = new Set(visits.filter((v) => v.accountId).map((v) => v.accountId)).size;
    const open = visits.filter((v) => v.status === VISIT_STATUS_OPEN).length;
    const completed = visits.filter((v) => v.status === 983080001).length;
    return { total: visits.length, uniqueCompanies, open, completed };
  }, [visits]);

  const selectedAccount = useMemo(() => accounts.find((a) => a.id === formAccountId) || null, [accounts, formAccountId]);

  function openDrawer(row?: VisitRow) {
    setFormError('');
    setEditOriginal(row || null);
    if (row) {
      setEditId(row.id);
      setFormAccountId(row.accountId);
      setFormAgentId(row.agentId);
      setFormDate(row.date ? row.date.substring(0, 10) : '');
      setFormStatus(String(row.status));
      setFormVisitType(row.isDouble ? 'double' : 'single');
      setFormAgent2Id(row.agent2Id);
      setFormRevisit(row.revisit ? row.revisit.substring(0, 10) : '');
      setFormNotes(row.notes);
    } else {
      setEditId(null);
      setFormAccountId('');
      setFormAgentId('');
      setFormDate(new Date().toISOString().substring(0, 10));
      setFormStatus(String(VISIT_STATUS_OPEN));
      setFormVisitType('single');
      setFormAgent2Id('');
      setFormRevisit('');
      setFormNotes('');
    }
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
      if (isDouble && formAgent2Id) {
        body[`${AGENT2_BIND_NAV}@odata.bind`] = `/cr301_agents(${formAgent2Id})`;
      } else if (editId) {
        // Switched back to Single (or cleared) on an existing visit — unlink the second owner.
        body[`${AGENT2_BIND_NAV}@odata.bind`] = null;
      }

      let result;
      if (editId) {
        result = await Cr301_visitsService.update(editId, body as Partial<Omit<Cr301_visitsBase, 'cr301_visitid'>>);
      } else {
        result = await Cr301_visitsService.create(body as unknown as Omit<Cr301_visitsBase, 'cr301_visitid'>);
      }
      if (!result.success) {
        throw new Error(operationErrorMessage(result.error, 'Save failed.'));
      }

      // Both directions: creating a visit increments its week's Actual; editing
      // its task owner(s) decrements the old owner's assignment and increments
      // the new one's. The reference date is the visit's own createdon (system
      // timestamp), not the user-entered Visit Date — see adjustVisitPlanActual's
      // doc comment for why, and it's immutable, so editing an existing visit
      // can only ever change WHICH owner(s) it counts toward, never WHICH week.
      // A failure/no-match here never undoes the fact that the visit itself
      // already saved fine — the outcome is surfaced in the toast, not swallowed.
      const adjustNotes: string[] = [];
      const referenceDateIso = editId ? editOriginal?.createdon || '' : result.data?.createdon || '';
      if (referenceDateIso) {
        const newAgentNames = [formAgentId, isDouble ? formAgent2Id : '']
          .filter(Boolean)
          .map((id) => agents.find((a) => a.id === id)?.name || '')
          .filter(Boolean);
        const outcomes: AdjustOutcome[] = [];
        if (!editId) {
          outcomes.push(...(await Promise.all(newAgentNames.map((name) => adjustVisitPlanActual(name, referenceDateIso, 1)))));
        } else if (editOriginal) {
          const oldAgentNames = [editOriginal.agentId, editOriginal.isDouble ? editOriginal.agent2Id : '']
            .filter(Boolean)
            .map((id) => agents.find((a) => a.id === id)?.name || '')
            .filter(Boolean);
          const sameAgents = JSON.stringify([...oldAgentNames].sort()) === JSON.stringify([...newAgentNames].sort());
          if (!sameAgents) {
            outcomes.push(...(await Promise.all(oldAgentNames.map((name) => adjustVisitPlanActual(name, referenceDateIso, -1)))));
            outcomes.push(...(await Promise.all(newAgentNames.map((name) => adjustVisitPlanActual(name, referenceDateIso, 1)))));
          }
        }
        adjustNotes.push(...outcomes.map(describeAdjustOutcome).filter((n): n is string => n !== null));
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
    const escapeForCsv = (value: string) => `"${value.replace(/"/g, '""')}"`;
    const header = ['Task Owner', 'Task Owner 2', 'Company', 'Company Code', 'Visit Date', 'Visit Status', 'Visit Type', 'Revisit Date', 'Comments'];
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
    ]);
    const csv = [header, ...rows].map((r) => r.map(escapeForCsv).join(',')).join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'visits_log.csv';
    link.click();
    URL.revokeObjectURL(url);
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
              <i className="fa-solid fa-plus"></i> Add Visit
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
                <i className="fa-solid fa-clipboard-check"></i>
              </span>
              <span className="kpi-label">Total visits</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.total}</div>
            <div className="kpi-context">&nbsp;</div>
          </div>
          <div className="kpi-card">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-gold">
                <i className="fa-solid fa-building"></i>
              </span>
              <span className="kpi-label">Unique companies</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.uniqueCompanies}</div>
            <div className="kpi-context">&nbsp;</div>
          </div>
          <div className="kpi-card kpi-warn">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-warn">
                <i className="fa-solid fa-door-open"></i>
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
                <i className="fa-solid fa-circle-check"></i>
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
              <i className="fa-solid fa-magnifying-glass"></i>
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
              </tr>
            </thead>
            <tbody>
              {loading ? (
                Array.from({ length: 6 }).map((_, i) => (
                  <tr key={i}>
                    {Array.from({ length: 8 }).map((__, j) => (
                      <td key={j}>
                        <div className="acc-skel"></div>
                      </td>
                    ))}
                  </tr>
                ))
              ) : loadError ? (
                <tr>
                  <td colSpan={8}>
                    <div className="acc-state">
                      <i className="fa-solid fa-triangle-exclamation"></i>
                      Could not load visits.
                      <br />
                      <small>{loadError}</small>
                    </div>
                  </td>
                </tr>
              ) : pageRows.length === 0 ? (
                <tr>
                  <td colSpan={8}>
                    <div className="acc-state">
                      <i className="fa-solid fa-inbox"></i>
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
                            <i className="fa-solid fa-users"></i> Double
                          </span>
                        ) : (
                          <span className="vis-vt">
                            <i className="fa-solid fa-user"></i> Single
                          </span>
                        )}
                      </td>
                      <td>{v.revisit ? fmtDate(v.revisit) : <span className="vis-muted">—</span>}</td>
                      <td>
                        <div className="vis-notes">{v.notes || <span className="vis-muted">—</span>}</div>
                      </td>
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
            <i className={`fa-solid ${editId ? 'fa-pen' : 'fa-plus'}`} style={{ color: 'var(--gold)', fontSize: editId ? 15 : 16 }}></i>
            {editId ? 'Edit Visit' : 'Add Visit'}
          </h2>
          <button className="acc-dclose" onClick={closeDrawer}>
            <i className="fa-solid fa-xmark"></i>
          </button>
        </div>
        <div className="acc-dbody">
          <div className={`acc-err${formError ? ' on' : ''}`}>{formError}</div>

          <div className="acc-fld">
            <label>Company Name</label>
            <select value={formAccountId} onChange={(e) => setFormAccountId(e.target.value)}>
              <option value="">— Select company —</option>
              {accounts.map((a) => (
                <option key={a.id} value={a.id}>
                  {a.name}
                </option>
              ))}
            </select>
          </div>

          <div className="acc-fld">
            <div className="acc-ro-wrap">
              <div className="acc-ro-label">Company Code</div>
              <div className="acc-ro-value">{selectedAccount?.code || '—'}</div>
            </div>
          </div>

          <div className="acc-fld">
            <label>Task Owner</label>
            <select value={formAgentId} onChange={(e) => setFormAgentId(e.target.value)}>
              <option value="">— Select agent —</option>
              {agents.map((a) => (
                <option key={a.id} value={a.id}>
                  {a.name}
                </option>
              ))}
            </select>
          </div>

          <div className="acc-fld">
            <label>Visit Date</label>
            <input type="date" value={formDate} onChange={(e) => setFormDate(e.target.value)} />
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
            <select value={formVisitType} onChange={(e) => setFormVisitType(e.target.value as 'single' | 'double')}>
              <option value="single">Single visit</option>
              <option value="double">Double visit</option>
            </select>
          </div>

          {formVisitType === 'double' && (
            <div className="acc-fld">
              <label>Task Owner 2</label>
              <select value={formAgent2Id} onChange={(e) => setFormAgent2Id(e.target.value)}>
                <option value="">— Second agent —</option>
                {agents.map((a) => (
                  <option key={a.id} value={a.id}>
                    {a.name}
                  </option>
                ))}
              </select>
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
            {saving ? 'Saving…' : 'Save Visit'}
          </button>
        </div>
      </aside>

      <div className={`acc-toast${toastOn ? ' on' : ''}`}>{toastMsg}</div>
    </>
  );
}
