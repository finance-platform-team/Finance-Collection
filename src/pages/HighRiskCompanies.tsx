import { useEffect, useMemo, useState } from 'react';
import { Cfm_highriskcompaniesesService } from '../generated/services/Cfm_highriskcompaniesesService';
import { Cfm_aragingsService } from '../generated/services/Cfm_aragingsService';
import type { Cfm_highriskcompanieses, Cfm_highriskcompaniesesBase } from '../generated/models/Cfm_highriskcompaniesesModel';
import type { Cfm_aragings } from '../generated/models/Cfm_aragingsModel';
import { operationErrorMessage } from '../lib/errors';
import { fetchAll } from '../lib/dataverse';
import { KSA_BU_CODES, type Region } from '../lib/useRegion';
import { escapeForCsv, downloadCsv } from '../lib/csvExport';
import './HighRiskCompanies.css';

/**
 * High Risk Companies — flagged accounts under active watch (cfm_highriskcompanieses),
 * linked back to their source AR aging row (cfm_aragings) for B.U / code / financials /
 * task owner. Mirrors the standalone Dataverse web resource (Ref.html) this tab replaces:
 * one high-risk record per matching aging row, financials + task owner copied in at
 * create time, status editable inline afterwards.
 */

const HIGH_RISK_SELECT = [
  'cfm_highriskcompaniesid',
  'cfm_month',
  'cfm_duescp',
  'cfm_totalcollectioncp',
  'cfm_comment',
  'cfm_status',
  'statecode',
  'statuscode',
  '_cfm_companyname_value',
  '_cfm_taskownercp_value',
  'createdon',
];

const COMPANY_SELECT = ['cfm_aragingid', 'cfm_companyname', 'cfm_companycode', 'cfm_totaldues', 'cfm_collectedamount', '_cfm_taskowner_value', '_cfm_type_value'];

const STATUS = {
  LEGAL: 766340000,
  ESCALATED: 766340001,
  UNDER_REVIEW: 766340002,
  RESOLVED: 766340003,
  ACTIVE: 766340004,
  INACTIVE: 766340005,
} as const;

const STATUS_LABEL: Record<number, string> = {
  [STATUS.LEGAL]: 'Legal',
  [STATUS.ESCALATED]: 'Escalated',
  [STATUS.UNDER_REVIEW]: 'Under Review',
  [STATUS.RESOLVED]: 'Resolved',
  [STATUS.ACTIVE]: 'Active',
  [STATUS.INACTIVE]: 'InActive',
};

const STATUS_OPTIONS = [STATUS.ACTIVE, STATUS.UNDER_REVIEW, STATUS.ESCALATED, STATUS.LEGAL, STATUS.RESOLVED, STATUS.INACTIVE];

function statusPillClass(v: number | null): string {
  if (v === STATUS.LEGAL) return 'hr-status-legal';
  if (v === STATUS.ESCALATED) return 'hr-status-escalated';
  if (v === STATUS.UNDER_REVIEW) return 'hr-status-underreview';
  if (v === STATUS.RESOLVED) return 'hr-status-resolved';
  if (v === STATUS.INACTIVE) return 'hr-status-inactive';
  return 'hr-status-active';
}

function formattedValue(record: unknown, attribute: string): string {
  const raw = (record as Record<string, unknown>)[`${attribute}@OData.Community.Display.V1.FormattedValue`];
  return typeof raw === 'string' ? raw : '';
}

const N = (v: number | string | null | undefined): number => Number(v || 0);
const fmt = (v: number | null): string => (v == null || v === 0 ? '—' : v.toLocaleString('en-US', { maximumFractionDigits: 0 }));

function fmtMonth(iso: string | null | undefined): string {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '—';
  return d.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
}

function monthKeyOf(iso: string | null | undefined): string {
  if (!iso) return '';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '';
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

interface AragingLite {
  code: string;
  buText: string;
}

interface CompanyOption {
  name: string;
  code: string;
  buText: string;
  ids: string[];
}

interface RiskRow {
  id: string;
  company: string;
  code: string;
  buText: string;
  month: string;
  monthKey: string;
  dues: number | null;
  collection: number | null;
  ownerName: string;
  status: number | null;
  active: boolean;
  comment: string;
}

function normaliseRow(row: Cfm_highriskcompanieses, aragingById: Map<string, AragingLite>): RiskRow {
  const aid = row._cfm_companyname_value || '';
  const a = aragingById.get(aid);
  const status = row.cfm_status == null ? null : Number(row.cfm_status);
  return {
    id: row.cfm_highriskcompaniesid,
    company: formattedValue(row, '_cfm_companyname_value') || '—',
    code: a?.code || '—',
    buText: a?.buText || '—',
    month: fmtMonth(row.cfm_month),
    monthKey: monthKeyOf(row.cfm_month),
    dues: row.cfm_duescp != null ? N(row.cfm_duescp) : null,
    collection: row.cfm_totalcollectioncp != null ? N(row.cfm_totalcollectioncp) : null,
    ownerName: formattedValue(row, '_cfm_taskownercp_value') || '—',
    status,
    active: row.statecode === 0,
    comment: row.cfm_comment || '',
  };
}

type CellState = 'saving' | 'saved' | 'error';

interface HighRiskCompaniesPageProps {
  region: Region | null;
}

const KSA_BU_SET = new Set(KSA_BU_CODES);
function norm(s: string): string {
  return (s || '').trim().toUpperCase();
}

export default function HighRiskCompaniesPage({ region }: HighRiskCompaniesPageProps) {
  const [rows, setRows] = useState<RiskRow[]>([]);
  const [companies, setCompanies] = useState<CompanyOption[]>([]);
  const [aragingById, setAragingById] = useState<Map<string, Cfm_aragings>>(new Map());
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [query, setQuery] = useState('');
  const [monthFilter, setMonthFilter] = useState('all');
  const [statusFilter, setStatusFilter] = useState('all');
  const [page, setPage] = useState(1);
  const pageSize = 25;

  const [cellState, setCellState] = useState<Record<string, CellState>>({});
  const [commentDrafts, setCommentDrafts] = useState<Record<string, string>>({});
  const [statusOpenId, setStatusOpenId] = useState<string | null>(null);
  const [statusMenuUp, setStatusMenuUp] = useState(false);

  // Decide up/down from the button's actual position in the viewport, not row
  // index — a fixed "last 2 rows open upward" guess overshoots past the table's
  // own header on a short table where there isn't really room above either.
  const STATUS_MENU_HEIGHT_ESTIMATE = STATUS_OPTIONS.length * 37 + 12;
  function toggleStatusMenu(rowId: string, button: HTMLButtonElement | null) {
    if (statusOpenId === rowId) {
      setStatusOpenId(null);
      return;
    }
    if (button) {
      const rect = button.getBoundingClientRect();
      const spaceBelow = window.innerHeight - rect.bottom;
      const spaceAbove = rect.top;
      setStatusMenuUp(spaceBelow < STATUS_MENU_HEIGHT_ESTIMATE && spaceAbove > spaceBelow);
    }
    setStatusOpenId(rowId);
  }

  const [toastMsg, setToastMsg] = useState('');
  const [toastOn, setToastOn] = useState(false);

  // Add drawer
  const [drawerOpen, setDrawerOpen] = useState(false);
  const [coQuery, setCoQuery] = useState('');
  const [coListOpen, setCoListOpen] = useState(false);
  const [selectedCo, setSelectedCo] = useState<CompanyOption | null>(null);
  const [formStatus, setFormStatus] = useState(String(STATUS.ACTIVE));
  const [formDate, setFormDate] = useState('');
  const [saving, setSaving] = useState(false);
  const [formError, setFormError] = useState('');

  function showToast(msg: string, durationMs = 3200) {
    setToastMsg(msg);
    setToastOn(true);
    setTimeout(() => setToastOn(false), durationMs);
  }

  function flashCell(key: string, state: CellState, clearAfter?: number) {
    setCellState((prev) => ({ ...prev, [key]: state }));
    if (clearAfter) {
      setTimeout(() => {
        setCellState((prev) => {
          const next = { ...prev };
          delete next[key];
          return next;
        });
      }, clearAfter);
    }
  }

  async function loadAragings(): Promise<{ options: CompanyOption[]; byId: Map<string, AragingLite>; fullById: Map<string, Cfm_aragings> }> {
    const list = await fetchAll<Cfm_aragings>(
      (skipToken) => Cfm_aragingsService.getAll({ select: COMPANY_SELECT, filter: 'cfm_companyname ne null', maxPageSize: 5000, skipToken }),
      'Could not load companies.'
    );
    const byId = new Map<string, AragingLite>();
    const fullById = new Map<string, Cfm_aragings>();
    const byName = new Map<string, CompanyOption>();
    list.forEach((r) => {
      const buText = formattedValue(r, '_cfm_type_value');
      byId.set(r.cfm_aragingid, { code: r.cfm_companycode || '—', buText: buText || '—' });
      fullById.set(r.cfm_aragingid, r);
      const name = (r.cfm_companyname || '').trim();
      if (!name) return;
      const existing = byName.get(name);
      if (existing) existing.ids.push(r.cfm_aragingid);
      else byName.set(name, { name, code: r.cfm_companycode || '', buText: buText || '—', ids: [r.cfm_aragingid] });
    });
    const options = Array.from(byName.values()).sort((a, b) => a.name.localeCompare(b.name));
    return { options, byId, fullById };
  }

  async function load() {
    setLoading(true);
    setLoadError(null);
    try {
      const { options, byId, fullById } = await loadAragings();
      setCompanies(options);
      setAragingById(fullById);
      const riskRows = await fetchAll<Cfm_highriskcompanieses>(
        (skipToken) => Cfm_highriskcompaniesesService.getAll({ select: HIGH_RISK_SELECT, orderBy: ['createdon desc'], maxPageSize: 5000, skipToken }),
        'Could not load high risk companies.'
      );
      setRows(riskRows.map((r) => normaliseRow(r, byId)));
    } catch (e) {
      setLoadError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    load();
  }, []);

  useEffect(() => {
    if (!statusOpenId) return;
    function onDocMouseDown() {
      setStatusOpenId(null);
    }
    document.addEventListener('mousedown', onDocMouseDown);
    return () => document.removeEventListener('mousedown', onDocMouseDown);
  }, [statusOpenId]);

  const regionRows = useMemo(() => {
    if (!region) return rows;
    return rows.filter((r) => {
      const inKsa = KSA_BU_SET.has(norm(r.buText));
      return region === 'KSA' ? inKsa : !inKsa;
    });
  }, [rows, region]);

  const monthOptions = useMemo(() => {
    const seen = new Map<string, string>();
    regionRows.forEach((r) => {
      if (r.monthKey && !seen.has(r.monthKey)) seen.set(r.monthKey, r.month);
    });
    return Array.from(seen.entries()).sort((a, b) => b[0].localeCompare(a[0]));
  }, [regionRows]);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    const matches = regionRows.filter((r) => {
      if (monthFilter !== 'all' && r.monthKey !== monthFilter) return false;
      if (statusFilter !== 'all' && String(r.status) !== statusFilter) return false;
      if (q && !`${r.company} ${r.code}`.toLowerCase().includes(q)) return false;
      return true;
    });
    // Resolved rows sink to the end of the list — everything else keeps its
    // existing relative order (Array.prototype.sort is a stable sort).
    return [...matches].sort((a, b) => (a.status === STATUS.RESOLVED ? 1 : 0) - (b.status === STATUS.RESOLVED ? 1 : 0));
  }, [regionRows, query, monthFilter, statusFilter]);

  useEffect(() => {
    setPage(1);
  }, [query, monthFilter, statusFilter, region]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const clampedPage = Math.min(page, totalPages);
  const pageRows = filtered.slice((clampedPage - 1) * pageSize, (clampedPage - 1) * pageSize + pageSize);

  function pagerButtons(): (number | 'ellipsis')[] {
    const pages: (number | 'ellipsis')[] = [];
    for (let p = 1; p <= totalPages; p++) {
      if (p === 1 || p === totalPages || Math.abs(p - clampedPage) <= 1) pages.push(p);
      else if (Math.abs(p - clampedPage) === 2) pages.push('ellipsis');
    }
    return pages;
  }

  const kpis = useMemo(() => {
    const totalDues = regionRows.reduce((t, r) => t + (r.dues || 0), 0);
    const totalCollected = regionRows.reduce((t, r) => t + (r.collection || 0), 0);
    const activeCount = regionRows.filter((r) => r.status === STATUS.ACTIVE || (r.status == null && r.active)).length;
    return { total: regionRows.length, totalDues, totalCollected, activeCount };
  }, [regionRows]);

  const coMatches = useMemo(() => {
    const q = coQuery.trim().toLowerCase();
    if (!q) return companies.slice(0, 30);
    return companies.filter((c) => `${c.code} ${c.name}`.toLowerCase().includes(q)).slice(0, 30);
  }, [companies, coQuery]);

  function pickCompany(c: CompanyOption) {
    setSelectedCo(c);
    setCoQuery(`${c.code} — ${c.name}`);
    setCoListOpen(false);
  }

  function openDrawer() {
    setSelectedCo(null);
    setCoQuery('');
    setCoListOpen(false);
    setFormStatus(String(STATUS.ACTIVE));
    setFormDate(new Date().toISOString().slice(0, 10));
    setFormError('');
    setDrawerOpen(true);
  }

  function closeDrawer() {
    setDrawerOpen(false);
  }

  const canSave = !saving && !!selectedCo;

  async function addHighRisk() {
    if (saving || !selectedCo) return;
    setSaving(true);
    setFormError('');
    try {
      const jobs = selectedCo.ids.map(async (aid) => {
        const a = aragingById.get(aid);
        const body: Record<string, unknown> = {
          'cfm_CompanyName@odata.bind': `/cfm_aragings(${aid})`,
          cfm_status: parseInt(formStatus, 10),
          statecode: 0,
          statuscode: 1,
        };
        if (a?.cfm_totaldues != null) body.cfm_duescp = a.cfm_totaldues;
        if (a?.cfm_collectedamount != null) body.cfm_totalcollectioncp = a.cfm_collectedamount;
        if (a?._cfm_taskowner_value) body['cfm_TaskOwnerCP@odata.bind'] = `/systemusers(${a._cfm_taskowner_value})`;
        if (formDate) body.cfm_month = `${formDate}T00:00:00Z`;
        const result = await Cfm_highriskcompaniesesService.create(body as unknown as Omit<Cfm_highriskcompaniesesBase, 'cfm_highriskcompaniesid'>);
        if (!result.success) throw new Error(operationErrorMessage(result.error, 'Save failed.'));
      });
      await Promise.all(jobs);
      showToast(`${selectedCo.ids.length} record${selectedCo.ids.length === 1 ? '' : 's'} added for ${selectedCo.name}.`);
      closeDrawer();
      await load();
    } catch (e) {
      setFormError(e instanceof Error ? e.message : String(e));
    } finally {
      setSaving(false);
    }
  }

  async function updateStatus(row: RiskRow, value: number) {
    if (value === row.status) return;
    const prev = row.status;
    const key = `${row.id}:status`;
    setRows((list) => list.map((r) => (r.id === row.id ? { ...r, status: value } : r)));
    flashCell(key, 'saving');
    try {
      const result = await Cfm_highriskcompaniesesService.update(row.id, { cfm_status: value } as Partial<Omit<Cfm_highriskcompaniesesBase, 'cfm_highriskcompaniesid'>>);
      if (!result.success) throw new Error(operationErrorMessage(result.error, 'Save failed.'));
      flashCell(key, 'saved', 900);
    } catch (e) {
      setRows((list) => list.map((r) => (r.id === row.id ? { ...r, status: prev } : r)));
      flashCell(key, 'error', 1400);
      showToast(e instanceof Error ? e.message : 'Save failed.', 4500);
    }
  }

  function commentValue(row: RiskRow): string {
    return row.id in commentDrafts ? commentDrafts[row.id] : row.comment;
  }

  async function commitComment(row: RiskRow) {
    const draft = commentDrafts[row.id];
    if (draft === undefined || draft === row.comment) return;
    const key = `${row.id}:comment`;
    flashCell(key, 'saving');
    try {
      const result = await Cfm_highriskcompaniesesService.update(row.id, { cfm_comment: draft.trim() || null } as Partial<Omit<Cfm_highriskcompaniesesBase, 'cfm_highriskcompaniesid'>>);
      if (!result.success) throw new Error(operationErrorMessage(result.error, 'Save failed.'));
      setRows((list) => list.map((r) => (r.id === row.id ? { ...r, comment: draft.trim() } : r)));
      setCommentDrafts((prev) => {
        const next = { ...prev };
        delete next[row.id];
        return next;
      });
      flashCell(key, 'saved', 900);
    } catch (e) {
      flashCell(key, 'error', 1400);
      showToast(e instanceof Error ? e.message : 'Save failed.', 4500);
    }
  }

  function exportCsv() {
    const header = ['B.U', 'Company', 'Code', 'Month', 'Dues (CP)', 'Collection (CP)', 'Task owner', 'Status', 'Comment'];
    const body = filtered.map((r) => [
      r.buText,
      r.company,
      r.code,
      r.month,
      r.dues ?? '',
      r.collection ?? '',
      r.ownerName,
      r.status != null ? STATUS_LABEL[r.status] || '' : '',
      r.comment,
    ]);
    downloadCsv('high_risk_companies.csv', [header, ...body].map((r) => r.map(escapeForCsv).join(',')));
  }

  return (
    <>
      <div className="acc-wrap hr-page">
        <div className="acc-head">
          <div>
            <h1 className="acc-title">High Risk Companies</h1>
            <div className="acc-sub">Flagged accounts under active watch, with collection dues, task owners, and remediation plans.</div>
          </div>
          <div className="acc-actions">
            <button className="acc-btn acc-btn-primary" onClick={openDrawer}>
              <i className="fa-solid fa-plus"></i> Add high risk
            </button>
            <button className="acc-btn acc-btn-ghost" onClick={exportCsv}>
              <i className="fa-solid fa-file-export"></i> Export
            </button>
          </div>
        </div>

        <div className="kpi-grid">
          <div className="kpi-card">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-gold">
                <i className="fa-solid fa-triangle-exclamation"></i>
              </span>
              <span className="kpi-label">Flagged companies</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.total}</div>
          </div>
          <div className="kpi-card kpi-danger">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-danger">
                <i className="fa-solid fa-sack-dollar"></i>
              </span>
              <span className="kpi-label">Total dues (CP)</span>
            </div>
            <div className="kpi-value">{loading ? '—' : fmt(kpis.totalDues)}</div>
          </div>
          <div className="kpi-card kpi-ok">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-ok">
                <i className="fa-solid fa-hand-holding-dollar"></i>
              </span>
              <span className="kpi-label">Total collection (CP)</span>
            </div>
            <div className="kpi-value">{loading ? '—' : fmt(kpis.totalCollected)}</div>
          </div>
          <div className="kpi-card kpi-info">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-info">
                <i className="fa-solid fa-circle-check"></i>
              </span>
              <span className="kpi-label">Active</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.activeCount}</div>
          </div>
        </div>

        <div className="tbl-shell">
          <div className="tbl-hdr">
            <div className="dm-filter-bar">
              <div className="acc-search dm-search">
                <i className="fa-solid fa-magnifying-glass"></i>
                <input placeholder="Search company, code…" value={query} onChange={(e) => setQuery(e.target.value)} />
              </div>
              <select className="acc-f" value={monthFilter} onChange={(e) => setMonthFilter(e.target.value)}>
                <option value="all">Month: All</option>
                {monthOptions.map(([key, label]) => (
                  <option key={key} value={key}>
                    Month: {label}
                  </option>
                ))}
              </select>
              <select className="acc-f" value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
                <option value="all">Status: All</option>
                {STATUS_OPTIONS.map((v) => (
                  <option key={v} value={v}>
                    Status: {STATUS_LABEL[v]}
                  </option>
                ))}
              </select>
            </div>
            <div className="dm-count">
              <i className="fa-solid fa-table"></i> Showing <strong>{filtered.length}</strong> of <strong>{regionRows.length}</strong>
            </div>
          </div>

          <div className="dm-tablewrap">
            <table>
              <thead>
                <tr>
                  <th>B.U</th>
                  <th>Company</th>
                  <th>Code</th>
                  <th>Month</th>
                  <th className="dm-numeric">Dues (CP)</th>
                  <th className="dm-numeric">Collection (CP)</th>
                  <th>Task owner</th>
                  <th>Status</th>
                  <th>Comment</th>
                </tr>
              </thead>
              <tbody>
                {loading ? (
                  Array.from({ length: 6 }).map((_, i) => (
                    <tr key={i}>
                      {Array.from({ length: 9 }).map((__, j) => (
                        <td key={j}>
                          <div className="acc-skel"></div>
                        </td>
                      ))}
                    </tr>
                  ))
                ) : loadError ? (
                  <tr>
                    <td colSpan={9}>
                      <div className="acc-state">
                        <i className="fa-solid fa-triangle-exclamation"></i>
                        Could not load high risk companies.
                        <br />
                        <small>{loadError}</small>
                      </div>
                    </td>
                  </tr>
                ) : filtered.length === 0 ? (
                  <tr>
                    <td colSpan={9}>
                      <div className="acc-state">
                        <i className="fa-solid fa-inbox"></i>
                        No high risk companies match your filters.
                      </div>
                    </td>
                  </tr>
                ) : (
                  pageRows.map((r) => {
                    const statusKey = `${r.id}:status`;
                    const commentKey = `${r.id}:comment`;
                    return (
                      <tr key={r.id}>
                        <td>
                          <span className="hr-bu">{r.buText}</span>
                        </td>
                        <td className="dm-company">{r.company}</td>
                        <td className="hr-code">{r.code}</td>
                        <td>{r.month}</td>
                        <td className="dm-numeric hr-dues">{fmt(r.dues)}</td>
                        <td className="dm-numeric">{fmt(r.collection)}</td>
                        <td className="hr-owner">{r.ownerName}</td>
                        <td onClick={(e) => e.stopPropagation()}>
                          <div className="hr-status-wrap">
                            <button
                              type="button"
                              className={`hr-status-select ${statusPillClass(r.status)}`}
                              onClick={(e) => toggleStatusMenu(r.id, e.currentTarget)}
                            >
                              {r.status != null ? STATUS_LABEL[r.status] : '—'}
                              <i className="fa-solid fa-chevron-down hr-status-caret"></i>
                            </button>
                            {statusOpenId === r.id && (
                              <div className={`hr-status-menu${statusMenuUp ? ' hr-status-menu-up' : ''}`} onMouseDown={(e) => e.stopPropagation()}>
                                {STATUS_OPTIONS.map((v) => (
                                  <div
                                    key={v}
                                    className={`hr-status-item ${statusPillClass(v)}${v === r.status ? ' active' : ''}`}
                                    onClick={() => {
                                      updateStatus(r, v);
                                      setStatusOpenId(null);
                                    }}
                                  >
                                    {STATUS_LABEL[v]}
                                    {v === r.status && <i className="fa-solid fa-check hr-status-item-check"></i>}
                                  </div>
                                ))}
                              </div>
                            )}
                          </div>
                          {cellState[statusKey] === 'saving' && <i className="fa-solid fa-spinner fa-spin hr-cell-ico"></i>}
                        </td>
                        <td onClick={(e) => e.stopPropagation()}>
                          <input
                            className={`hr-comment-inp${cellState[commentKey] === 'error' ? ' err' : ''}`}
                            placeholder="Add a comment…"
                            value={commentValue(r)}
                            onChange={(e) => setCommentDrafts((prev) => ({ ...prev, [r.id]: e.target.value }))}
                            onBlur={() => commitComment(r)}
                          />
                          {cellState[commentKey] === 'saving' && <i className="fa-solid fa-spinner fa-spin hr-cell-ico"></i>}
                          {cellState[commentKey] === 'saved' && <i className="fa-solid fa-check hr-cell-ico ok"></i>}
                        </td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {!loading && !loadError && filtered.length > 0 && (
            <div className="dm-pagination">
              <div className="dm-page-info">
                Showing <strong>{(clampedPage - 1) * pageSize + 1}</strong>–<strong>{Math.min(clampedPage * pageSize, filtered.length)}</strong> of{' '}
                <strong>{filtered.length}</strong>
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

      {/* Add high risk company drawer */}
      <div className={`acc-scrim${drawerOpen ? ' on' : ''}`} onClick={closeDrawer}></div>
      <aside className={`acc-drawer${drawerOpen ? ' on' : ''}`}>
        <div className="acc-dhead">
          <h2>New · High Risk Company</h2>
          <button className="acc-dclose" onClick={closeDrawer}>
            <i className="fa-solid fa-xmark"></i>
          </button>
        </div>
        <div className="acc-dbody">
          <div className={`acc-err${formError ? ' on' : ''}`}>{formError}</div>

          <div className="acc-combo hr-co-wrap">
            <div className="acc-fld" style={{ marginBottom: 0 }}>
              <label>
                Company <span className="dm-req">*</span>
              </label>
              <input
                placeholder="Search by name or code…"
                value={coQuery}
                onChange={(e) => {
                  setCoQuery(e.target.value);
                  setCoListOpen(true);
                  setSelectedCo(null);
                }}
                onFocus={() => setCoListOpen(true)}
              />
            </div>
            {coListOpen && (
              <div className="acc-combo-list">
                {coMatches.length === 0 ? (
                  <div className="acc-combo-empty">No matches.</div>
                ) : (
                  coMatches.map((c) => (
                    <div key={c.name} className="acc-combo-item" onMouseDown={(e) => e.preventDefault()} onClick={() => pickCompany(c)}>
                      <div>
                        {c.code} — {c.name}
                      </div>
                      <small>
                        {c.buText}
                        {c.ids.length > 1 ? ` · ${c.ids.length} matching records` : ''}
                      </small>
                    </div>
                  ))
                )}
              </div>
            )}
            {selectedCo && (
              <div className="acc-hint">
                <b>{selectedCo.ids.length}</b> matching aging record{selectedCo.ids.length === 1 ? '' : 's'} will be added.
              </div>
            )}
          </div>

          <div className="acc-fld">
            <label>Status</label>
            <select value={formStatus} onChange={(e) => setFormStatus(e.target.value)}>
              {STATUS_OPTIONS.map((v) => (
                <option key={v} value={v}>
                  {STATUS_LABEL[v]}
                </option>
              ))}
            </select>
          </div>

          <div className="acc-fld">
            <label>Date</label>
            <input type="date" value={formDate} onChange={(e) => setFormDate(e.target.value)} />
          </div>

          <div className="acc-hint">Dues, collection and task owner are pulled automatically from the Collection Plan.</div>
        </div>
        <div className="acc-dfoot">
          <div style={{ display: 'flex', gap: 8, marginLeft: 'auto' }}>
            <button className="acc-btn acc-btn-ghost" onClick={closeDrawer}>
              Cancel
            </button>
            <button className="acc-btn acc-btn-primary" disabled={!canSave} onClick={addHighRisk}>
              {saving ? 'Adding…' : '+ Add high risk'}
            </button>
          </div>
        </div>
      </aside>

      <div className={`acc-toast${toastOn ? ' on' : ''}`}>{toastMsg}</div>
    </>
  );
}
