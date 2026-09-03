import Icon from '../components/Icon';
import { useEffect, useMemo, useRef, useState } from 'react';
import { Ar_disputemanagementsService } from '../generated/services/Ar_disputemanagementsService';
import { Cfm_disputemanagementupdatesesService } from '../generated/services/Cfm_disputemanagementupdatesesService';
import { Cfm_insurancecompaniesService } from '../generated/services/Cfm_insurancecompaniesService';
import type { Ar_disputemanagements, Ar_disputemanagementsBase } from '../generated/models/Ar_disputemanagementsModel';
import type { Cfm_disputemanagementupdatesesBase } from '../generated/models/Cfm_disputemanagementupdatesesModel';
import type { Cfm_insurancecompanies } from '../generated/models/Cfm_insurancecompaniesModel';
import { operationErrorMessage } from '../lib/errors';
import { fetchAll } from '../lib/dataverse';
import type { DocFile } from '../lib/docs';
import { docExt, docIconClass, docIcon, fmtBytes, fileToBase64, docsFlow, listDocs, uploadDoc } from '../lib/docs';
import { useConfirm } from '../lib/useConfirm';
import { KSA_BU_CODES, type Region } from '../lib/useRegion';
import './DisputeManagement.css';

/**
 * Dispute Management — AR decision/follow-up tracking for rejected or disputed
 * collections. Parent: ar_disputemanagements. Company lookup is cfm_aragings, same
 * table/picker as Daily Collection. A lightweight child log (cfm_disputemanagementupdates)
 * records updates against a dispute, added from that dispute's detail drawer;
 * it does not itself change the dispute's AR Review status (that stays a direct,
 * explicit inline edit on the row).
 */

const DISPUTE_SELECT = [
  'ar_disputemanagementid',
  'ar_code',
  'ar_amount',
  'ar_category',
  'ar_due',
  'ar_rejection',
  'ar_totalclaim',
  'ar_agreement',
  'cfm_arreview',
  'cfm_companyname',
  'cfm_proof',
  '_cfm_companycode_value',
  '_createdby_value',
  'createdon',
  'statecode',
];

// ar_disputemanagements.cfm_CompanyCode targets cfm_insurancecompanies, NOT
// cfm_aragings (a different table with unrelated ids, despite the similar
// "company code" naming) — binding an cfm_aragings id here fails with
// "Entity 'cfm_InsuranceCompany' ... Does Not Exist".
const COMPANY_SELECT = ['cfm_insurancecompanyid', 'cfm_code', 'cfm_companyname', 'cfm_name', 'cfm_companytype', 'cfm_butext'];

const KSA_BU_SET = new Set(KSA_BU_CODES);
function normBu(s: string): string {
  return (s || '').trim().toUpperCase();
}

const CATEGORY = {
  REJECTED: 0,
  LEGAL_ISSUES: 1,
  CLAIMS_ISSUE: 2,
  BANKRUPTCY: 3,
  STOPPED: 4,
} as const;

const CATEGORY_LABEL: Record<number, string> = {
  0: 'Rejected',
  1: 'Legal Issues',
  2: 'Claims Issue',
  3: 'Bankruptcy',
  4: 'Stopped',
};

const CATEGORY_PILLS = [
  { value: null as number | null, label: 'All' },
  { value: CATEGORY.REJECTED, label: 'Rejected' },
  { value: CATEGORY.LEGAL_ISSUES, label: 'Legal Issues' },
  { value: CATEGORY.BANKRUPTCY, label: 'Bankruptcy' },
  { value: CATEGORY.CLAIMS_ISSUE, label: 'Claims Issue' },
  { value: CATEGORY.STOPPED, label: 'Stopped' },
];

const AR_REVIEW = {
  REJECTED: 0,
  PENDING: 1,
  APPROVED: 2,
  RECONCILED: 3,
} as const;

const AR_REVIEW_LABEL: Record<number, string> = {
  0: 'Rejected',
  1: 'Pending',
  2: 'Approved',
  3: 'Reconciled',
};

const AR_REVIEW_OPTIONS = [
  { value: AR_REVIEW.PENDING, label: 'Pending' },
  { value: AR_REVIEW.APPROVED, label: 'Approved' },
  { value: AR_REVIEW.REJECTED, label: 'Rejected' },
  { value: AR_REVIEW.RECONCILED, label: 'Reconciled' },
];

const UPDATE_TYPE_OPTIONS = [
  { value: 766340000, label: 'Manual Update' },
  { value: 766340001, label: 'Customer Call' },
  { value: 766340002, label: 'Payment Update' },
  { value: 766340003, label: 'Legal Action' },
  { value: 766340004, label: 'Bankruptcy' },
  { value: 766340005, label: 'Dispute Resolution' },
  { value: 766340006, label: 'Escalation' },
  { value: 766340007, label: 'Follow-up' },
  { value: 766340008, label: 'Other' },
];

function formattedValue(record: unknown, attribute: string): string {
  const raw = (record as Record<string, unknown>)[`${attribute}@OData.Community.Display.V1.FormattedValue`];
  return typeof raw === 'string' ? raw : '';
}

const N = (v: number | string | null | undefined): number => Number(v || 0);
const fmt = (n: number): string => (n == null || n === 0 ? '—' : n.toLocaleString('en-US', { maximumFractionDigits: 0 }));

function fmtDue(iso: string | null | undefined): string {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '—';
  return d.toLocaleDateString('en-US', { month: 'short', year: 'numeric' });
}

interface CompanyOption {
  id: string;
  code: string;
  name: string;
  entityName: string;
  buText: string;
}

interface HistoryEntry {
  id: string;
  updateType: number | null;
  updateTypeLabel: string;
  date: string;
  description: string;
}

function updateTypeLabelFor(value: number | null): string {
  if (value == null) return '—';
  return UPDATE_TYPE_OPTIONS.find((o) => o.value === value)?.label || '—';
}

interface DisputeRow {
  id: string;
  code: string;
  companyName: string;
  companyId: string;
  amount: number;
  rejectionPct: number;
  totalClaim: number;
  due: string;
  category: number | null;
  categoryLabel: string;
  ownerId: string;
  ownerName: string;
  proof: string;
  agreement: string;
  arReview: number;
  arReviewLabel: string;
  createdOn: string;
  active: boolean;
  region: 'EG' | 'KSA' | '';
}

function normaliseDispute(row: Ar_disputemanagements, companies: CompanyOption[]): DisputeRow {
  const coId = row._cfm_companycode_value || '';
  const co = coId ? companies.find((c) => c.id === coId) : undefined;
  const category = row.ar_category == null ? null : Number(row.ar_category);
  const arReview = row.cfm_arreview == null ? AR_REVIEW.PENDING : Number(row.cfm_arreview);
  return {
    id: row.ar_disputemanagementid,
    code: row.ar_code || '—',
    companyName: row.cfm_companyname || (co ? co.name : '') || '—',
    companyId: coId,
    amount: N(row.ar_amount),
    rejectionPct: N(row.ar_rejection),
    totalClaim: N(row.ar_totalclaim),
    due: row.ar_due || '',
    category,
    categoryLabel: formattedValue(row, 'ar_category') || (category != null ? CATEGORY_LABEL[category] : '—'),
    ownerId: row._createdby_value || '',
    ownerName: formattedValue(row, '_createdby_value') || '—',
    proof: row.cfm_proof || '',
    agreement: row.ar_agreement || '',
    arReview,
    arReviewLabel: formattedValue(row, 'cfm_arreview') || AR_REVIEW_LABEL[arReview],
    createdOn: row.createdon || '',
    active: Number(row.statecode) === 0,
    region: co ? (KSA_BU_SET.has(normBu(co.buText)) ? 'KSA' : 'EG') : '',
  };
}

function arReviewPillClass(v: number): string {
  if (v === AR_REVIEW.APPROVED) return 'dm-review-approved';
  if (v === AR_REVIEW.REJECTED) return 'dm-review-rejected';
  if (v === AR_REVIEW.RECONCILED) return 'dm-review-reconciled';
  return 'dm-review-pending';
}

// SharePoint folder for a dispute case's documents — same helper flow as Daily
// Collection (see ../lib/docs), keyed by company + code instead of company + code + date.
function docsFolderName(row: DisputeRow): string {
  const strip = (s: string | null | undefined) => (s == null ? '' : String(s)).replace(/["*:<>?/\\|]/g, '').replace(/\s+/g, ' ').trim();
  const name = strip(row.companyName) || 'Company';
  const code = strip(row.code) || 'NA';
  return `${name} - ${code} - (Dispute)`;
}

type CellState = 'saving' | 'saved' | 'error';

interface DisputeManagementPageProps {
  region: Region | null;
}

export default function DisputeManagementPage({ region }: DisputeManagementPageProps) {
  const [rows, setRows] = useState<DisputeRow[]>([]);
  const [companies, setCompanies] = useState<CompanyOption[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [query, setQuery] = useState('');
  const [reviewFilter, setReviewFilter] = useState<string>('');
  const [categoryFilter, setCategoryFilter] = useState<number | null>(null);
  // Defaults to only showing active disputes, same as before inactive ones were
  // even loaded — pick "Inactive" or "All" to see deactivated ones too.
  const [statusFilter, setStatusFilter] = useState<'active' | 'inactive' | 'all'>('active');
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(15);

  const { confirm, dialog } = useConfirm();

  const [cellState, setCellState] = useState<Record<string, CellState>>({});
  const cellTimers = useRef<Record<string, ReturnType<typeof setTimeout>>>({});
  const [agreementDrafts, setAgreementDrafts] = useState<Record<string, string>>({});

  const [toastMsg, setToastMsg] = useState('');
  const [toastOn, setToastOn] = useState(false);
  const toastTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  // Create Dispute Case drawer — Code doubles as the company-code search box;
  // Company is a read-only field auto-filled from whichever match gets picked.
  const [createOpen, setCreateOpen] = useState(false);
  const [coListOpen, setCoListOpen] = useState(false);
  const [selectedCo, setSelectedCo] = useState<CompanyOption | null>(null);
  const [formCode, setFormCode] = useState('');
  const [formAmount, setFormAmount] = useState('');
  const [formCategory, setFormCategory] = useState(String(CATEGORY.REJECTED));
  const [formDue, setFormDue] = useState('');
  const [saving, setSaving] = useState(false);
  const [formError, setFormError] = useState('');

  // AR Review pill dropdown (single row open at a time)
  const [reviewOpenId, setReviewOpenId] = useState<string | null>(null);
  const reviewMenuRef = useRef<HTMLDivElement | null>(null);

  // Dispute detail drawer (opened by clicking a row) — stat summary + update
  // history + an inline "Add Update" form scoped to that one dispute.
  const [detailOpen, setDetailOpen] = useState(false);
  const [detailDispute, setDetailDispute] = useState<DisputeRow | null>(null);
  const [history, setHistory] = useState<HistoryEntry[]>([]);
  const [historyLoading, setHistoryLoading] = useState(false);
  const [historyError, setHistoryError] = useState('');
  const [updateDate, setUpdateDate] = useState('');
  const [updateType, setUpdateType] = useState(String(UPDATE_TYPE_OPTIONS[0].value));
  const [updateDesc, setUpdateDesc] = useState('');
  const [updateSaving, setUpdateSaving] = useState(false);
  const [updateError, setUpdateError] = useState('');

  // Files drawer (opened by "View files") — lists/uploads/deletes SharePoint
  // documents for one dispute case, via the shared docs flow.
  const [docsOpen, setDocsOpen] = useState(false);
  const [docsRow, setDocsRow] = useState<DisputeRow | null>(null);
  const [docs, setDocs] = useState<DocFile[]>([]);
  const [docsLoading, setDocsLoading] = useState(false);
  const [docsError, setDocsError] = useState('');
  const [docsUploadingNames, setDocsUploadingNames] = useState<string[]>([]);

  function showToast(msg: string, durationMs = 3200) {
    setToastMsg(msg);
    setToastOn(true);
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToastOn(false), durationMs);
  }

  function flashCell(key: string, state: CellState, clearAfter?: number) {
    clearTimeout(cellTimers.current[key]);
    setCellState((prev) => ({ ...prev, [key]: state }));
    if (clearAfter) {
      cellTimers.current[key] = setTimeout(() => {
        setCellState((prev) => {
          const next = { ...prev };
          delete next[key];
          return next;
        });
      }, clearAfter);
    }
  }

  async function loadCompanies(): Promise<CompanyOption[]> {
    const list = await fetchAll<Cfm_insurancecompanies>(
      (skipToken) => Cfm_insurancecompaniesService.getAll({ select: COMPANY_SELECT, filter: 'statecode eq 0', maxPageSize: 5000, skipToken }),
      'Could not load companies.'
    );
    return list.map((r) => ({
      id: r.cfm_insurancecompanyid,
      code: r.cfm_code || '',
      name: r.cfm_companyname || r.cfm_name || '',
      entityName: formattedValue(r, 'cfm_companytype') || '—',
      buText: r.cfm_butext || '',
    }));
  }

  async function load() {
    setLoading(true);
    setLoadError(null);
    try {
      const [companyOptions, disputeRows] = await Promise.all([
        loadCompanies(),
        fetchAll<Ar_disputemanagements>(
          // Not scoped to statecode eq 0 — inactive disputes are loaded too so the
          // status filter can show/toggle them instead of them just disappearing.
          (skipToken) => Ar_disputemanagementsService.getAll({ select: DISPUTE_SELECT, orderBy: ['createdon desc'], maxPageSize: 5000, skipToken }),
          'Could not load dispute cases.'
        ),
      ]);
      setCompanies(companyOptions);
      setRows(disputeRows.map((r) => normaliseDispute(r, companyOptions)));
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
    if (!reviewOpenId) return;
    function onDocMouseDown(e: MouseEvent) {
      if (reviewMenuRef.current && !reviewMenuRef.current.contains(e.target as Node)) {
        setReviewOpenId(null);
      }
    }
    document.addEventListener('mousedown', onDocMouseDown);
    return () => document.removeEventListener('mousedown', onDocMouseDown);
  }, [reviewOpenId]);

  // App-wide Region toggle (Layout header) — unset region on a dispute counts
  // as Egypt, same convention used everywhere else this filter appears.
  const regionRows = useMemo(
    () => (!region ? rows : rows.filter((r) => (region === 'KSA' ? r.region === 'KSA' : r.region !== 'KSA'))),
    [rows, region]
  );

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    const matches = regionRows.filter((r) => {
      if (statusFilter !== 'all' && r.active !== (statusFilter === 'active')) return false;
      if (categoryFilter != null && r.category !== categoryFilter) return false;
      if (reviewFilter !== '' && r.arReview !== Number(reviewFilter)) return false;
      if (q && !`${r.code} ${r.companyName} ${r.agreement}`.toLowerCase().includes(q)) return false;
      return true;
    });
    // Reconciled disputes are settled — keep them out of the way at the bottom of
    // the list instead of mixed in by date. Everything else keeps its existing order.
    return [...matches].sort((a, b) => {
      const aLast = a.arReview === AR_REVIEW.RECONCILED ? 1 : 0;
      const bLast = b.arReview === AR_REVIEW.RECONCILED ? 1 : 0;
      return aLast - bLast;
    });
  }, [regionRows, query, reviewFilter, categoryFilter, statusFilter]);

  useEffect(() => {
    setPage(1);
  }, [query, reviewFilter, categoryFilter, statusFilter, region]);

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
    // Same rows the KPIs always reflected — inactive disputes are now loaded
    // (for the status filter/toggle) but stay out of these totals, same as
    // before they were loaded at all.
    const activeRows = regionRows.filter((r) => r.active);
    const totalDisputes = activeRows.length;
    const disputedAmount = activeRows.reduce((t, r) => t + r.amount, 0);
    const pendingAr = activeRows.filter((r) => r.arReview === AR_REVIEW.PENDING);
    const approved = activeRows.filter((r) => r.arReview === AR_REVIEW.APPROVED).length;
    return { totalDisputes, disputedAmount, pendingAr, approved };
  }, [regionRows]);

  const coMatches = useMemo(() => {
    const q = formCode.trim().toLowerCase();
    if (!q) return [];
    return companies.filter((c) => c.code.toLowerCase().includes(q)).slice(0, 20);
  }, [companies, formCode]);

  function pickCompany(c: CompanyOption) {
    setSelectedCo(c);
    setFormCode(c.code);
    setCoListOpen(false);
  }

  function openCreateDrawer() {
    setSelectedCo(null);
    setCoListOpen(false);
    setFormCode('');
    setFormAmount('');
    setFormCategory(String(CATEGORY.REJECTED));
    setFormDue('');
    setFormError('');
    setCreateOpen(true);
  }

  function closeCreateDrawer() {
    setCreateOpen(false);
  }

  async function saveCreate() {
    if (saving) return;
    if (!formCode.trim()) {
      setFormError('Enter a company code.');
      return;
    }
    if (!selectedCo) {
      setFormError('Pick a company from the matches under Code — typing a code alone is not enough.');
      return;
    }
    if (!formAmount || N(formAmount) <= 0) {
      setFormError('Enter a valid Amount.');
      return;
    }
    if (!formDue) {
      setFormError('Enter a Due date.');
      return;
    }
    setSaving(true);
    setFormError('');
    try {
      const body: Record<string, unknown> = {
        'cfm_CompanyCode@odata.bind': `/cfm_insurancecompanies(${selectedCo.id})`,
        ar_code: formCode.trim(),
        cfm_companyname: selectedCo.name,
        ar_amount: N(formAmount),
        ar_category: parseInt(formCategory, 10),
        ar_due: formDue ? `${formDue}T00:00:00Z` : null,
        cfm_arreview: AR_REVIEW.PENDING,
      };
      const result = await Ar_disputemanagementsService.create(body as unknown as Omit<Ar_disputemanagementsBase, 'ar_disputemanagementid'>);
      if (!result.success) throw new Error(operationErrorMessage(result.error, 'Save failed.'));
      showToast('Dispute case created.');
      closeCreateDrawer();
      await load();
    } catch (e) {
      setFormError(e instanceof Error ? e.message : String(e));
    } finally {
      setSaving(false);
    }
  }

  async function loadHistory(disputeId: string) {
    setHistoryLoading(true);
    setHistoryError('');
    try {
      const res = await Cfm_disputemanagementupdatesesService.getAll({
        select: ['cfm_disputemanagementupdatesid', 'cfm_updatetype', 'cfm_updatedate', 'cfm_description'],
        filter: `statecode eq 0 and _cfm_disputemanagement_value eq ${disputeId}`,
        orderBy: ['cfm_updatedate desc'],
      });
      if (!res.success) throw new Error(operationErrorMessage(res.error, 'Could not load history.'));
      const entries: HistoryEntry[] = (res.data || []).map((h) => {
        const t = h.cfm_updatetype == null ? null : Number(h.cfm_updatetype);
        return {
          id: h.cfm_disputemanagementupdatesid,
          updateType: t,
          updateTypeLabel: updateTypeLabelFor(t),
          date: h.cfm_updatedate || '',
          description: h.cfm_description || '',
        };
      });
      setHistory(entries);
    } catch (e) {
      setHistoryError(e instanceof Error ? e.message : String(e));
    } finally {
      setHistoryLoading(false);
    }
  }

  function openDetail(row: DisputeRow) {
    setDetailDispute(row);
    setUpdateDate(new Date().toISOString().slice(0, 10));
    setUpdateType(String(UPDATE_TYPE_OPTIONS[0].value));
    setUpdateDesc('');
    setUpdateError('');
    setHistory([]);
    setDetailOpen(true);
    loadHistory(row.id);
  }

  function closeDetail() {
    setDetailOpen(false);
  }

  const canAddUpdate = !updateSaving && !!detailDispute && updateDesc.trim() !== '';

  async function addHistoryEntry() {
    if (updateSaving || !detailDispute) return;
    setUpdateSaving(true);
    setUpdateError('');
    try {
      const body: Record<string, unknown> = {
        'cfm_DisputeManagement@odata.bind': `/ar_disputemanagements(${detailDispute.id})`,
        cfm_updatetype: parseInt(updateType, 10),
        cfm_description: updateDesc.trim(),
        cfm_updatedate: updateDate ? `${updateDate}T00:00:00Z` : new Date().toISOString(),
      };
      const result = await Cfm_disputemanagementupdatesesService.create(body as unknown as Omit<Cfm_disputemanagementupdatesesBase, 'cfm_disputemanagementupdatesid'>);
      if (!result.success) throw new Error(operationErrorMessage(result.error, 'Save failed.'));
      setUpdateDesc('');
      showToast('Saved — the history entry was saved.');
      await loadHistory(detailDispute.id);
    } catch (e) {
      setUpdateError(e instanceof Error ? e.message : String(e));
    } finally {
      setUpdateSaving(false);
    }
  }

  async function updateArReview(row: DisputeRow, value: number) {
    if (value === row.arReview) return;
    const prev = row.arReview;
    const key = `${row.id}:review`;
    setRows((list) => list.map((r) => (r.id === row.id ? { ...r, arReview: value, arReviewLabel: AR_REVIEW_LABEL[value] } : r)));
    flashCell(key, 'saving');
    try {
      const result = await Ar_disputemanagementsService.update(row.id, { cfm_arreview: value } as Partial<Omit<Ar_disputemanagementsBase, 'ar_disputemanagementid'>>);
      if (!result.success) throw new Error(operationErrorMessage(result.error, 'Save failed.'));
      flashCell(key, 'saved', 900);
    } catch (e) {
      setRows((list) => list.map((r) => (r.id === row.id ? { ...r, arReview: prev, arReviewLabel: AR_REVIEW_LABEL[prev] } : r)));
      flashCell(key, 'error', 1400);
      showToast(e instanceof Error ? e.message : 'Save failed.', 4500);
    }
  }

  async function toggleDisputeStatus(row: DisputeRow) {
    const goingActive = !row.active;
    const ok = await confirm({
      title: goingActive ? `Activate ${row.code}?` : `Deactivate ${row.code}?`,
      message: goingActive
        ? 'This dispute case shows up in the list and totals again.'
        : 'This dispute case is kept (nothing is deleted), but drops out of the default list and totals until reactivated.',
      danger: !goingActive,
      confirmLabel: goingActive ? 'Activate' : 'Deactivate',
    });
    if (!ok) return;
    try {
      const result = await Ar_disputemanagementsService.update(row.id, {
        statecode: goingActive ? 0 : 1,
        statuscode: goingActive ? 1 : 2,
      } as Partial<Omit<Ar_disputemanagementsBase, 'ar_disputemanagementid'>>);
      if (!result.success) throw new Error(operationErrorMessage(result.error, 'Could not update status.'));
      setRows((list) => list.map((r) => (r.id === row.id ? { ...r, active: goingActive } : r)));
      showToast(goingActive ? `${row.code} activated.` : `${row.code} deactivated.`);
    } catch (e) {
      showToast(e instanceof Error ? e.message : 'Could not update status.', 4500);
    }
  }

  function agreementValue(row: DisputeRow): string {
    return row.id in agreementDrafts ? agreementDrafts[row.id] : row.agreement;
  }

  async function commitAgreement(row: DisputeRow) {
    const draft = agreementDrafts[row.id];
    if (draft === undefined || draft === row.agreement) return;
    const key = `${row.id}:agreement`;
    flashCell(key, 'saving');
    try {
      const result = await Ar_disputemanagementsService.update(row.id, { ar_agreement: draft.trim() || null } as Partial<Omit<Ar_disputemanagementsBase, 'ar_disputemanagementid'>>);
      if (!result.success) throw new Error(operationErrorMessage(result.error, 'Save failed.'));
      setRows((list) => list.map((r) => (r.id === row.id ? { ...r, agreement: draft.trim() } : r)));
      setAgreementDrafts((prev) => {
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

  async function loadDocs(forRow?: DisputeRow | null) {
    const row = forRow ?? docsRow;
    if (!row) return;
    const folder = docsFolderName(row);
    setDocsLoading(true);
    setDocsError('');
    try {
      setDocs(await listDocs(folder));
    } catch (err) {
      const msg = err instanceof Error ? err.message : String(err);
      // A "folder/list not found" result just means no uploads yet — treat as empty, quietly.
      if (!/not\s*found|404|does not exist|itemnotfound|list not found|no\s*such/i.test(msg)) {
        setDocsError(msg);
      }
      setDocs([]);
    } finally {
      setDocsLoading(false);
    }
  }

  function openDocsDrawer(row: DisputeRow) {
    setDocsRow(row);
    setDocs([]);
    setDocsError('');
    setDocsOpen(true);
    loadDocs(row);
  }

  function closeDocsDrawer() {
    setDocsOpen(false);
  }

  async function onDocsSelected(files: File[]) {
    if (!files.length || !docsRow) return;
    const folder = docsFolderName(docsRow);
    let ok = 0;
    let fail = 0;
    let firstErr = '';
    for (const file of files) {
      setDocsUploadingNames((prev) => [...prev, file.name]);
      try {
        const b64 = await fileToBase64(file);
        await uploadDoc({
          folderName: folder,
          collectionId: docsRow.id.replace(/[{}]/g, ''),
          collectionRef: docsRow.code,
          fileName: file.name,
          fileContentBase64: b64,
        });
        ok++;
      } catch (err) {
        fail++;
        if (!firstErr) firstErr = err instanceof Error ? err.message : String(err);
      } finally {
        setDocsUploadingNames((prev) => prev.filter((n) => n !== file.name));
      }
    }
    if (ok) showToast(`${ok} file${ok === 1 ? '' : 's'} attached${fail ? `, ${fail} failed` : ''}.`);
    else if (fail) showToast(`Upload failed: ${firstErr}`);
    await loadDocs();
  }

  async function deleteDocFile(path: string, name: string) {
    if (!docsRow) return;
    if (!window.confirm(`Delete "${name}"? This cannot be undone.`)) return;
    const folder = docsFolderName(docsRow);
    try {
      await docsFlow({ action: 'delete', folderName: folder, filePath: path });
      showToast(`Deleted ${name}.`);
      await loadDocs();
    } catch (err) {
      showToast(`Delete failed: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  return (
    <>
      <div className="acc-wrap">
        <div className="acc-head">
          <div>
            <h1 className="acc-title">Dispute Management</h1>
            <div className="acc-sub">Collection plan execution, Daily Collection entries, AR decision feedback, follow-up and dispute operations.</div>
          </div>
          <div className="acc-actions">
            <button className="acc-btn acc-btn-primary" onClick={openCreateDrawer}>
              <Icon className="fa-solid fa-plus" /> Create Dispute Case
            </button>
          </div>
        </div>

        <div className="kpi-grid">
          <div className="kpi-card">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-gold">
                <Icon className="fa-solid fa-folder-open" />
              </span>
              <span className="kpi-label">Total disputes</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.totalDisputes}</div>
          </div>
          <div className="kpi-card kpi-gold">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-gold">
                <Icon className="fa-solid fa-sack-dollar" />
              </span>
              <span className="kpi-label">Disputed amount</span>
            </div>
            <div className="kpi-value">{loading ? '—' : fmt(kpis.disputedAmount)}</div>
          </div>
          <div className="kpi-card kpi-warn">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-warn">
                <Icon className="fa-solid fa-clock" />
              </span>
              <span className="kpi-label">Pending AR</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.pendingAr.length}</div>
          </div>
          <div className="kpi-card kpi-ok">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-ok">
                <Icon className="fa-solid fa-circle-check" />
              </span>
              <span className="kpi-label">Approved</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.approved}</div>
          </div>
        </div>

        <div className="tbl-shell">
          <div className="tbl-hdr">
            <div className="dm-filter-bar">
              <div className="acc-search dm-search">
                <Icon className="fa-solid fa-magnifying-glass" />
                <input placeholder="Search code, company, agreement…" value={query} onChange={(e) => setQuery(e.target.value)} />
              </div>
              <select className="acc-f" value={reviewFilter} onChange={(e) => setReviewFilter(e.target.value)}>
                <option value="">AR Review: All</option>
                {AR_REVIEW_OPTIONS.map((o) => (
                  <option key={o.value} value={o.value}>
                    AR Review: {o.label}
                  </option>
                ))}
              </select>
              <select className="acc-f" value={statusFilter} onChange={(e) => setStatusFilter(e.target.value as 'active' | 'inactive' | 'all')}>
                <option value="active">Active only</option>
                <option value="inactive">Inactive only</option>
                <option value="all">All statuses</option>
              </select>
            </div>
            <div className="dm-count">
              <Icon className="fa-solid fa-table" /> Showing <strong>{filtered.length}</strong> of <strong>{regionRows.length}</strong>
            </div>
          </div>

          <div className="dm-category-bar">
            <span className="dm-category-label">Category</span>
            {CATEGORY_PILLS.map((p) => (
              <button
                key={String(p.value)}
                className={`dm-cat-pill${categoryFilter === p.value ? ' active' : ''}`}
                onClick={() => setCategoryFilter(p.value)}
              >
                {p.label}
              </button>
            ))}
          </div>

          <div className="dm-tablewrap">
            <table>
              <thead>
                <tr>
                  <th>Code</th>
                  <th>Company</th>
                  <th className="dm-numeric">Amount</th>
                  <th className="dm-numeric">Rej %</th>
                  <th>Due</th>
                  <th>Category</th>
                  <th>Owner</th>
                  <th>Proof</th>
                  <th>Agreement</th>
                  <th>AR Review</th>
                  <th></th>
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
                ) : loadError ? (
                  <tr>
                    <td colSpan={11}>
                      <div className="acc-state">
                        <Icon className="fa-solid fa-triangle-exclamation" />
                        Could not load dispute cases.
                        <br />
                        <small>{loadError}</small>
                      </div>
                    </td>
                  </tr>
                ) : filtered.length === 0 ? (
                  <tr>
                    <td colSpan={11}>
                      <div className="acc-state">
                        <Icon className="fa-solid fa-inbox" />
                        No dispute cases match your filters.
                      </div>
                    </td>
                  </tr>
                ) : (
                  pageRows.map((r) => {
                    const reviewKey = `${r.id}:review`;
                    const agreementKey = `${r.id}:agreement`;
                    return (
                      <tr key={r.id} className={`dm-row${r.active ? '' : ' dm-row-inactive'}`} onClick={() => openDetail(r)}>
                        <td className="dm-code">
                          {r.code}
                          {!r.active && <span className="dm-pill dm-pill-inactive">Inactive</span>}
                        </td>
                        <td className="dm-company">{r.companyName}</td>
                        <td className="dm-numeric dm-amt">{fmt(r.amount)}</td>
                        <td className="dm-numeric">{r.rejectionPct ? `${fmt(r.rejectionPct)}%` : '—'}</td>
                        <td>{fmtDue(r.due)}</td>
                        <td>
                          <span className="dm-pill dm-pill-category">{r.categoryLabel}</span>
                        </td>
                        <td className="dm-owner">{r.ownerName}</td>
                        <td onClick={(e) => e.stopPropagation()}>
                          <button className="dm-viewfiles" onClick={() => openDocsDrawer(r)}>
                            <Icon className="fa-solid fa-folder-open" /> View files
                          </button>
                        </td>
                        <td onClick={(e) => e.stopPropagation()}>
                          <input
                            className={`dm-agreement-inp${cellState[agreementKey] === 'error' ? ' err' : ''}`}
                            placeholder="Agreement…"
                            value={agreementValue(r)}
                            onChange={(e) => setAgreementDrafts((prev) => ({ ...prev, [r.id]: e.target.value }))}
                            onBlur={() => commitAgreement(r)}
                          />
                          {cellState[agreementKey] === 'saving' && <Icon className="fa-solid fa-spinner fa-spin dm-cell-ico" />}
                          {cellState[agreementKey] === 'saved' && <Icon className="fa-solid fa-check dm-cell-ico ok" />}
                        </td>
                        <td onClick={(e) => e.stopPropagation()}>
                          <div className="dm-review-wrap" ref={reviewOpenId === r.id ? reviewMenuRef : undefined}>
                            <button
                              type="button"
                              className={`dm-review-select ${arReviewPillClass(r.arReview)}`}
                              onClick={() => setReviewOpenId(reviewOpenId === r.id ? null : r.id)}
                            >
                              {AR_REVIEW_LABEL[r.arReview]}
                              <Icon className="fa-solid fa-chevron-down dm-review-caret" />
                            </button>
                            {reviewOpenId === r.id && (
                              <div className="dm-review-menu">
                                {AR_REVIEW_OPTIONS.map((o) => (
                                  <div
                                    key={o.value}
                                    className={`dm-review-item ${arReviewPillClass(o.value)}${o.value === r.arReview ? ' active' : ''}`}
                                    onClick={() => {
                                      updateArReview(r, o.value);
                                      setReviewOpenId(null);
                                    }}
                                  >
                                    {o.label}
                                    {o.value === r.arReview && <Icon className="fa-solid fa-check dm-review-item-check" />}
                                  </div>
                                ))}
                              </div>
                            )}
                          </div>
                          {cellState[reviewKey] === 'saving' && <Icon className="fa-solid fa-spinner fa-spin dm-cell-ico" />}
                        </td>
                        <td onClick={(e) => e.stopPropagation()}>
                          <button
                            className={`dm-status-toggle${r.active ? ' dm-status-toggle-active' : ''}`}
                            title={r.active ? 'Deactivate' : 'Activate'}
                            onClick={() => toggleDisputeStatus(r)}
                          >
                            <Icon className={`fa-solid ${r.active ? 'fa-toggle-on' : 'fa-toggle-off'}`} />
                          </button>
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
              <div className="dm-page-info">
                Showing <strong>{(clampedPage - 1) * pageSize + 1}</strong>–
                <strong>{Math.min(clampedPage * pageSize, filtered.length)}</strong> of <strong>{filtered.length}</strong>
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

      {/* Create Dispute Case drawer */}
      <div className={`acc-scrim${createOpen ? ' on' : ''}`} onClick={closeCreateDrawer}></div>
      <aside className={`acc-drawer${createOpen ? ' on' : ''}`}>
        <div className="acc-dhead">
          <h2>New · Dispute Case</h2>
          <button className="acc-dclose" onClick={closeCreateDrawer}>
            <Icon className="fa-solid fa-xmark" />
          </button>
        </div>
        <div className="acc-dbody">
          <div className={`acc-err${formError ? ' on' : ''}`}>{formError}</div>

          <div className="acc-combo dm-co-wrap">
            <div className="acc-fld" style={{ marginBottom: 0 }}>
              <label>
                Code <span className="dm-req">*</span>
              </label>
              <input
                placeholder="Type a company code…"
                value={formCode}
                onChange={(e) => {
                  setFormCode(e.target.value);
                  setCoListOpen(true);
                  setSelectedCo(null);
                }}
                onFocus={() => setCoListOpen(true)}
              />
            </div>
            {coListOpen && formCode.trim() && (
              <div className="acc-combo-list">
                {coMatches.length === 0 ? (
                  <div className="acc-combo-empty">No matches.</div>
                ) : (
                  coMatches.map((c) => (
                    <div key={c.id} className="acc-combo-item" onClick={() => pickCompany(c)}>
                      <div>
                        {c.code} — {c.name}
                      </div>
                      <small>{c.entityName}</small>
                    </div>
                  ))
                )}
              </div>
            )}
          </div>

          <div className="acc-fld">
            <label>Company</label>
            <div className="acc-ro-wrap">
              <div className="acc-ro-label">Auto-filled from selection</div>
              <div className="acc-ro-value">{selectedCo ? selectedCo.name : '—'}</div>
            </div>
          </div>

          <div className="acc-fld">
            <label>Category</label>
            <select value={formCategory} onChange={(e) => setFormCategory(e.target.value)}>
              {Object.entries(CATEGORY_LABEL).map(([v, label]) => (
                <option key={v} value={v}>
                  {label}
                </option>
              ))}
            </select>
          </div>

          <div className="acc-fld">
            <label>
              Amount <span className="dm-req">*</span>
            </label>
            <input type="number" step="any" min={0} placeholder="0.00" value={formAmount} onChange={(e) => setFormAmount(e.target.value)} />
          </div>

          <div className="acc-fld">
            <label>
              Due (Month / Year) <span className="dm-req">*</span>
            </label>
            <input type="date" value={formDue} onChange={(e) => setFormDue(e.target.value)} />
          </div>
        </div>
        <div className="acc-dfoot">
          <div style={{ display: 'flex', gap: 8, marginLeft: 'auto' }}>
            <button className="acc-btn acc-btn-ghost" onClick={closeCreateDrawer}>
              Cancel
            </button>
            <button className="acc-btn acc-btn-primary" disabled={saving} onClick={saveCreate}>
              {saving ? (
                <>
                  <Icon className="fa-solid fa-spinner fa-spin" /> Saving…
                </>
              ) : (
                'Create dispute case'
              )}
            </button>
          </div>
        </div>
      </aside>

      {/* Dispute detail drawer — opened by clicking a row */}
      <div className={`acc-scrim${detailOpen ? ' on' : ''}`} onClick={closeDetail}></div>
      <aside className={`acc-drawer${detailOpen ? ' on' : ''}`}>
        <div className="acc-dhead">
          <h2 className="dm-detail-head">
            <Icon className="fa-solid fa-scale-balanced" />
            <span>
              {detailDispute?.arReviewLabel} · {detailDispute?.code} · {detailDispute?.companyName}
            </span>
          </h2>
          <button className="acc-dclose" onClick={closeDetail}>
            <Icon className="fa-solid fa-xmark" />
          </button>
        </div>
        <div className="acc-dbody">
          {detailDispute && (
            <>
              <div className="dm-detail-grid">
                <div className="dm-stat">
                  <div className="dm-stat-label">Amount</div>
                  <div className="dm-stat-value">{fmt(detailDispute.amount)}</div>
                </div>
                <div className="dm-stat">
                  <div className="dm-stat-label">Due (Y/M)</div>
                  <div className="dm-stat-value">{fmtDue(detailDispute.due)}</div>
                </div>
                <div className="dm-stat">
                  <div className="dm-stat-label">AR review</div>
                  <div className="dm-stat-value">{detailDispute.arReviewLabel}</div>
                </div>
              </div>

              <div className="dm-section-title">History</div>
              {historyLoading ? (
                <div className="dm-history-empty">
                  <Icon className="fa-solid fa-spinner fa-spin" /> Loading…
                </div>
              ) : historyError ? (
                <div className="dm-history-empty" style={{ color: 'var(--bad)' }}>
                  {historyError}
                </div>
              ) : history.length === 0 ? (
                <div className="dm-history-empty">No updates logged yet.</div>
              ) : (
                <div className="dm-history">
                  {history.map((h) => (
                    <div key={h.id} className="dm-history-item">
                      <span className="dm-history-dot"></span>
                      <div className="dm-history-head">
                        <span className="dm-history-type">{h.updateTypeLabel}</span>
                        <span className="dm-history-date">{h.date ? h.date.slice(0, 10) : ''}</span>
                      </div>
                      {h.description && <div className="dm-history-desc">{h.description}</div>}
                    </div>
                  ))}
                </div>
              )}

              <div className="dm-section-title">Add update</div>
              <div className={`acc-err${updateError ? ' on' : ''}`}>{updateError}</div>
              <div className="acc-fld">
                <label>Date</label>
                <input type="date" value={updateDate} onChange={(e) => setUpdateDate(e.target.value)} />
              </div>
              <div className="acc-fld">
                <label>Update type</label>
                <select value={updateType} onChange={(e) => setUpdateType(e.target.value)}>
                  {UPDATE_TYPE_OPTIONS.map((o) => (
                    <option key={o.value} value={o.value}>
                      {o.label}
                    </option>
                  ))}
                </select>
              </div>
              <div className="acc-fld">
                <label>What happened?</label>
                <textarea rows={4} placeholder="Free text…" value={updateDesc} onChange={(e) => setUpdateDesc(e.target.value)} />
              </div>
            </>
          )}
        </div>
        <div className="acc-dfoot">
          <div style={{ display: 'flex', gap: 8, marginLeft: 'auto' }}>
            <button className="acc-btn acc-btn-ghost" onClick={closeDetail}>
              Cancel
            </button>
            <button className="acc-btn acc-btn-primary" disabled={!canAddUpdate} onClick={addHistoryEntry}>
              <Icon className="fa-solid fa-paper-plane" /> {updateSaving ? 'Saving…' : 'Add Update'}
            </button>
          </div>
        </div>
      </aside>

      {/* Files drawer — opened by "View files"; lists/uploads/deletes SharePoint documents */}
      <div className={`acc-scrim${docsOpen ? ' on' : ''}`} onClick={closeDocsDrawer}></div>
      <aside className={`acc-drawer${docsOpen ? ' on' : ''}`}>
        <div className="acc-dhead">
          <h2>
            <Icon className="fa-solid fa-folder-open" /> {docsRow?.code}
          </h2>
          <button className="acc-dclose" onClick={closeDocsDrawer}>
            <Icon className="fa-solid fa-xmark" />
          </button>
        </div>
        <div className="acc-dbody">
          <div className="dc-docs-sec" style={{ marginTop: 0 }}>
            <div className="dc-docs-hdr">
              <div className="dc-docs-title">
                Documents
                {docs.length > 0 && <span className="dc-docs-count">({docs.length})</span>}
              </div>
              <label className="dc-docs-upload-btn">
                <Icon className="fa-solid fa-upload" /> Upload
                <input
                  type="file"
                  multiple
                  onChange={(e) => {
                    const files = Array.from(e.target.files || []);
                    e.target.value = '';
                    onDocsSelected(files);
                  }}
                />
              </label>
            </div>
            <div className="dc-docs-list">
              {docsLoading ? (
                <div className="dc-doc-uploading">
                  <Icon className="fa-solid fa-spinner fa-spin" /> Loading documents…
                </div>
              ) : docsError ? (
                <div className="dc-docs-empty">
                  <Icon className="fa-solid fa-circle-info" /> Couldn&apos;t load documents right now. You can still upload — click Upload.
                </div>
              ) : docsUploadingNames.length > 0 || docs.length > 0 ? (
                <>
                  {docsUploadingNames.map((n, i) => (
                    <div key={`up-${i}-${n}`} className="dc-doc-uploading">
                      <Icon className="fa-solid fa-spinner fa-spin" /> Uploading {n}…
                    </div>
                  ))}
                  {docs.map((f) => {
                    const ext = docExt(f.name);
                    return (
                      <div key={f.path || f.name} className="dc-doc-row">
                        <div className={`dc-doc-ico dc-doc-ico-${docIconClass(ext)}`}>
                          <Icon className={`fa-solid ${docIcon(ext)}`} />
                        </div>
                        <div className="dc-doc-meta">
                          <div className="dc-doc-name" title={f.name}>
                            {f.name}
                          </div>
                          <div className="dc-doc-sub">
                            {ext.toUpperCase()}
                            {f.size ? ` · ${fmtBytes(f.size)}` : ''}
                          </div>
                        </div>
                        <div className="dc-doc-act">
                          <a href={f.url} target="_blank" rel="noopener noreferrer" title="Open">
                            <Icon className="fa-solid fa-arrow-up-right-from-square" />
                          </a>
                          <button title="Delete" onClick={() => deleteDocFile(f.path, f.name)}>
                            <Icon className="fa-solid fa-trash" />
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </>
              ) : (
                <div className="dc-docs-empty">
                  <Icon className="fa-solid fa-folder-open" /> No documents yet — click Upload to attach files.
                </div>
              )}
            </div>
            <div className="dc-docs-note">Any file type. Stored on SharePoint, not in Dataverse.</div>
          </div>
        </div>
      </aside>

      <div className={`acc-toast${toastOn ? ' on' : ''}`}>{toastMsg}</div>
      {dialog}
    </>
  );
}
