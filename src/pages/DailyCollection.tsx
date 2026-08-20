import { useEffect, useMemo, useRef, useState } from 'react';
import { getContext } from '@microsoft/power-apps/app';
import { Cfm_dailycollectionsService } from '../generated/services/Cfm_dailycollectionsService';
import { Cfm_collectionmonthallocationsService } from '../generated/services/Cfm_collectionmonthallocationsService';
import { Cfm_aragingsService } from '../generated/services/Cfm_aragingsService';
import { Cr301_accountsService } from '../generated/services/Cr301_accountsService';
import { SystemusersService } from '../generated/services/SystemusersService';
import { Cfm_dailycollectionscfm_paymentmethod } from '../generated/models/Cfm_dailycollectionsModel';
import type { Cfm_dailycollections, Cfm_dailycollectionsBase } from '../generated/models/Cfm_dailycollectionsModel';
import type { Cfm_collectionmonthallocations, Cfm_collectionmonthallocationsBase } from '../generated/models/Cfm_collectionmonthallocationsModel';
import type { Cfm_aragings } from '../generated/models/Cfm_aragingsModel';
import type { Cr301_accounts } from '../generated/models/Cr301_accountsModel';
import { operationErrorMessage } from '../lib/errors';
import { fetchAll } from '../lib/dataverse';
import type { DocFile } from '../lib/docs';
import { docExt, docIconClass, docIcon, fmtBytes, fileToBase64, docsFlow, listDocs, uploadDoc } from '../lib/docs';
import './DailyCollection.css';

/**
 * Ported from Ref.html (the "CFM · Daily Collection" Dataverse HTML web resource).
 * Parent: cfm_dailycollections. Child: cfm_collectionmonthallocations, linked via the
 * confirmed nav property cfm_DailyCollection (odata.bind). Company lookup is
 * cfm_aragings — only companies with an AR Aging row ("a Plan") can be collected
 * against, so the picker's source list IS that table, not a separate flag.
 *
 * New entries are created as "Awaiting documents" (766340005) and move to "Pending
 * AR Verification" automatically once the first document is uploaded — see
 * promoteAfterUpload(). Documents are stored on SharePoint (not Dataverse) via the
 * CFM_CollectionDocs Power Automate flow, called directly from the browser the same
 * way Ref.html does it.
 */

const ENTRY_SELECT = [
  'cfm_dailycollectionid',
  'cfm_name',
  'cfm_companynamefx',
  'cfm_entityfx',
  'cfm_collectiondate',
  'cfm_paymentmethod',
  'cfm_collectedamountnet',
  'cfm_proof',
  'cfm_note',
  'cfm_paymentreference',
  'cfm_statusaction',
  '_cfm_companycode_value',
  '_cfm_bankaccount_value',
  'createdon',
  '_createdby_value',
];

const COMPANY_SELECT = ['cfm_aragingid', 'cfm_companycode', 'cfm_companyname', '_cfm_type_value'];

const BANK_ACCOUNT_SELECT = ['cr301_accountid', 'cr301_acccountname', 'cr301_accountcode'];

const ALLOC_COUNT_SELECT = ['cfm_collectionmonthallocationid', '_cfm_dailycollection_value'];

const ALLOC_DETAIL_SELECT = [
  'cfm_collectionmonthallocationid',
  'cfm_allocatetomonth',
  'cfm_allocationamount',
  'cfm_tax',
  'cfm_volumediscount',
  'cfm_earlypaymentdisc',
  'cfm_adminfees',
  'cfm_rejection',
];

const AR_STATUS = {
  PENDING_VERIFY: 766340000,
  REJECTED: 766340001,
  APPROVED: 766340002,
  ALLOCATED: 766340003,
  PENDING_BANK: 766340004,
  AWAITING_DOCS: 766340005,
} as const;

const AR_STATUS_LABEL: Record<number, string> = {
  766340000: 'Pending AR Verification',
  766340001: 'Rejected from AR',
  766340002: 'Approved',
  766340003: 'Allocated',
  766340004: 'Pending Bank Statement',
  766340005: 'Awaiting documents',
};

const LOCK_REASON: Record<number, string> = {
  [AR_STATUS.APPROVED]: 'This entry has been approved by AR. Editing is disabled.',
  [AR_STATUS.ALLOCATED]: 'This entry has been allocated. Editing is disabled.',
  [AR_STATUS.PENDING_BANK]: 'This entry is awaiting the bank statement from Treasury. Editing is disabled until Treasury acts.',
};

function isLocked(status: number | null): boolean {
  if (status == null) return false;
  if (status === AR_STATUS.AWAITING_DOCS) return false;
  if (status === AR_STATUS.PENDING_VERIFY) return false;
  if (status === AR_STATUS.REJECTED) return false;
  return true;
}

const PAYMENT_METHOD_OPTIONS = Object.entries(Cfm_dailycollectionscfm_paymentmethod).map(([value, label]) => ({
  value: Number(value),
  label: label.trim(),
}));

interface MonthOption {
  value: string;
  label: string;
}

/** 24 months back through 35 months forward, as Date-Only "YYYY-MM-01" values. */
function generateMonthOptions(): MonthOption[] {
  const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const opts: MonthOption[] = [];
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth() - 24, 1);
  for (let i = 0; i < 60; i++) {
    const d = new Date(start.getFullYear(), start.getMonth() + i, 1);
    opts.push({ value: `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-01`, label: `${MONTHS[d.getMonth()]} ${d.getFullYear()}` });
  }
  return opts;
}
const MONTH_OPTIONS = generateMonthOptions();
function defaultMonthValue(): string {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-01`;
}

function formattedValue(record: unknown, attribute: string): string {
  const raw = (record as Record<string, unknown>)[`${attribute}@OData.Community.Display.V1.FormattedValue`];
  return typeof raw === 'string' ? raw : '';
}

const N = (v: number | string | null | undefined): number => Number(v || 0);
const fmt = (n: number): string => (n == null || n === 0 ? '—' : n.toLocaleString('en-US', { maximumFractionDigits: 0 }));

function fmtDate(iso: string): string {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '—';
  return d.toISOString().slice(0, 10);
}
function fmtDateTime(iso: string): string {
  if (!iso) return '—';
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return '—';
  const dd = String(d.getDate()).padStart(2, '0');
  const mm = String(d.getMonth() + 1).padStart(2, '0');
  const ap = d.getHours() >= 12 ? 'PM' : 'AM';
  const hh = d.getHours() % 12 || 12;
  const min = String(d.getMinutes()).padStart(2, '0');
  return `${dd}-${mm}-${d.getFullYear()} ${hh}:${min} ${ap}`;
}

interface CompanyOption {
  id: string;
  code: string;
  name: string;
  entityId: string;
  entityName: string;
}

interface BankAccountOption {
  id: string;
  name: string;
  code: string;
}

function bankAccountLabel(b: BankAccountOption): string {
  return b.code ? `${b.name} — ${b.code}` : b.name;
}

interface EntryRow {
  id: string;
  ref: string;
  entity: string;
  companyCode: string;
  companyName: string;
  companyId: string;
  bankAccountId: string;
  date: string;
  method: number | null;
  methodLabel: string;
  net: number;
  proof: string;
  note: string;
  paymentReference: string;
  status: number | null;
  statusLabel: string;
  createdBy: string;
  createdById: string;
  createdOn: string;
  allocCount: number;
}

interface AllocRow {
  childId: string | null;
  month: string;
  amount: number;
  tax: number;
  vol: number;
  early: number;
  admin: number;
  reject: number;
}

type Mode = 'create' | 'edit' | 'reedit';

function normaliseEntry(row: Cfm_dailycollections, companies: CompanyOption[]): EntryRow {
  const coId = row._cfm_companycode_value || '';
  const co = coId ? companies.find((c) => c.id === coId) : undefined;
  const status = row.cfm_statusaction == null ? null : Number(row.cfm_statusaction);
  const method = row.cfm_paymentmethod == null ? null : Number(row.cfm_paymentmethod);
  return {
    id: row.cfm_dailycollectionid,
    ref: row.cfm_name || '—',
    entity: co ? co.entityName || '—' : row.cfm_entityfx || '—',
    companyCode: co ? co.code || '—' : '—',
    companyName: co ? co.name || '—' : row.cfm_companynamefx || '—',
    companyId: coId,
    bankAccountId: row._cfm_bankaccount_value || '',
    date: row.cfm_collectiondate || '',
    method,
    methodLabel: formattedValue(row, 'cfm_paymentmethod') || (method == null ? '—' : PAYMENT_METHOD_OPTIONS.find((o) => o.value === method)?.label || String(method)),
    net: N(row.cfm_collectedamountnet),
    proof: row.cfm_proof || '',
    note: row.cfm_note || '',
    paymentReference: row.cfm_paymentreference || '',
    status,
    statusLabel: (formattedValue(row, 'cfm_statusaction') || (status != null ? AR_STATUS_LABEL[status] : 'Awaiting documents')).trim(),
    createdBy: formattedValue(row, '_createdby_value') || '—',
    createdById: (row._createdby_value || '').replace(/[{}]/g, '').toLowerCase(),
    createdOn: row.createdon || '',
    allocCount: 0,
  };
}

// Status is a real Dataverse value now (including "Awaiting documents"), so the
// display label is a direct lookup — no more deriving it from the proof link.
function displayStatusLabel(r: EntryRow): string {
  if (r.status === AR_STATUS.AWAITING_DOCS || r.status == null) return 'Awaiting documents';
  return r.statusLabel;
}

function statusPillClass(status: number | null): string {
  switch (status) {
    case AR_STATUS.APPROVED:
      return 'dc-pill-approved';
    case AR_STATUS.REJECTED:
      return 'dc-pill-rejected';
    case AR_STATUS.ALLOCATED:
      return 'dc-pill-allocated';
    case AR_STATUS.PENDING_BANK:
      return 'dc-pill-bank';
    case AR_STATUS.AWAITING_DOCS:
      return 'dc-pill-awaiting';
    default:
      return 'dc-pill-pending';
  }
}
function statusIcon(status: number | null): string {
  switch (status) {
    case AR_STATUS.APPROVED:
      return 'fa-circle-check';
    case AR_STATUS.REJECTED:
      return 'fa-circle-xmark';
    case AR_STATUS.ALLOCATED:
      return 'fa-layer-group';
    case AR_STATUS.PENDING_BANK:
      return 'fa-building-columns';
    case AR_STATUS.AWAITING_DOCS:
      return 'fa-file-circle-plus';
    default:
      return 'fa-clock';
  }
}

/* ========================================================================
 * Documents (SharePoint via the CFM_CollectionDocs Power Automate flow) — see
 * ../lib/docs. folderName = "{CompanyName} - {Code} - {dd-mm-yyyy} - ({Ref})".
 * ======================================================================== */
function docsFolderName(entry: EntryRow): string {
  const strip = (s: string | null | undefined) =>
    (s == null ? '' : String(s)).replace(/["*:<>?/\\|]/g, '').replace(/\s+/g, ' ').trim();
  const name = strip(entry.companyName) || 'Company';
  const code = strip(entry.companyCode) || 'NA';
  let dateStr = 'NODATE';
  if (entry.date) {
    const d = new Date(entry.date);
    if (!Number.isNaN(d.getTime())) {
      const dd = String(d.getDate()).padStart(2, '0');
      const mm = String(d.getMonth() + 1).padStart(2, '0');
      dateStr = `${dd}-${mm}-${d.getFullYear()}`;
    }
  }
  const refInParen = strip(entry.ref) || 'REF';
  return `${name} - ${code} - ${dateStr} - (${refInParen})`;
}

/**
 * Runs once, right after a Daily Collection record is successfully created (never on
 * edit/re-submit, never on load/refresh): reads the entry's net collected amount and its
 * AR Aging lookup, then ADDS the amount onto that AR Aging row's Pending AR Verification
 * field (cfm_pendingconfirmation, treating a null existing value as 0). Only that field is
 * touched — no other AR Aging column is read or written here, and no AR Aging row is ever
 * created. Returns an error message on failure so the caller can surface it without
 * rolling back the already-created Daily Collection entry.
 */
async function addCollectedAmountToArAging(dailyCollectionId: string): Promise<string | null> {
  const dcRes = await Cfm_dailycollectionsService.get(dailyCollectionId, {
    select: ['cfm_collectedamountnet', '_cfm_companycode_value'],
  });
  if (!dcRes.success || !dcRes.data) {
    return operationErrorMessage(dcRes.error, 'Could not re-read the saved entry to update AR Aging.');
  }
  const aragingId = dcRes.data._cfm_companycode_value;
  if (!aragingId) return null; // no AR Aging lookup on the entry — nothing to update

  const arRes = await Cfm_aragingsService.get(aragingId, { select: ['cfm_pendingconfirmation'] });
  if (!arRes.success || !arRes.data) {
    return operationErrorMessage(arRes.error, 'Could not read the matching AR Aging record.');
  }

  const newPending = N(arRes.data.cfm_pendingconfirmation) + N(dcRes.data.cfm_collectedamountnet);
  const updateRes = await Cfm_aragingsService.update(aragingId, { cfm_pendingconfirmation: newPending });
  if (!updateRes.success) {
    return operationErrorMessage(updateRes.error, 'Could not update AR Aging Pending AR Verification.');
  }
  return null;
}

export default function DailyCollectionPage() {
  const [entries, setEntries] = useState<EntryRow[]>([]);
  const [companies, setCompanies] = useState<CompanyOption[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [currentUser, setCurrentUser] = useState<{ id: string; name: string }>({ id: '', name: '' });

  const [query, setQuery] = useState('');
  const [entityFilter, setEntityFilter] = useState('');
  const [methodFilter, setMethodFilter] = useState('');
  const [statusFilter, setStatusFilter] = useState('');
  const [byMe, setByMe] = useState(false);
  const [rejectedOnly, setRejectedOnly] = useState(false);
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState<number | 'all'>(25);

  const [drawerOpen, setDrawerOpen] = useState(false);
  const [mode, setMode] = useState<Mode>('create');
  const [editId, setEditId] = useState<string | null>(null);
  const [editEntry, setEditEntry] = useState<EntryRow | null>(null);
  const [locked, setLocked] = useState(false);

  const [coQuery, setCoQuery] = useState('');
  const [coListOpen, setCoListOpen] = useState(false);
  const [selectedCo, setSelectedCo] = useState<CompanyOption | null>(null);

  const [bankAccounts, setBankAccounts] = useState<BankAccountOption[]>([]);
  const [bankQuery, setBankQuery] = useState('');
  const [bankListOpen, setBankListOpen] = useState(false);

  const [formDate, setFormDate] = useState('');
  const [formMethod, setFormMethod] = useState('');
  const [formNet, setFormNet] = useState('');
  const [formPaymentRef, setFormPaymentRef] = useState('');
  const [formBankAccountId, setFormBankAccountId] = useState('');
  const [formNote, setFormNote] = useState('');

  const [allocs, setAllocs] = useState<AllocRow[]>([]);
  const [originalChildIds, setOriginalChildIds] = useState<string[]>([]);
  const [allocsLoading, setAllocsLoading] = useState(false);
  const [allocsError, setAllocsError] = useState('');

  const [docs, setDocs] = useState<DocFile[]>([]);
  const [docsLoading, setDocsLoading] = useState(false);
  const [docsError, setDocsError] = useState('');
  const [docsUploadingNames, setDocsUploadingNames] = useState<string[]>([]);

  const [saving, setSaving] = useState(false);
  const [formError, setFormError] = useState('');

  const [toastMsg, setToastMsg] = useState('');
  const [toastOn, setToastOn] = useState(false);
  const toastTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  function showToast(msg: string, durationMs = 3200) {
    setToastMsg(msg);
    setToastOn(true);
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToastOn(false), durationMs);
  }

  async function loadCompanies(): Promise<CompanyOption[]> {
    const rows = await fetchAll<Cfm_aragings>(
      (skipToken) => Cfm_aragingsService.getAll({ select: COMPANY_SELECT, filter: 'statecode eq 0', maxPageSize: 5000, skipToken }),
      'Could not load companies.'
    );
    return rows.map((r) => ({
      id: r.cfm_aragingid,
      code: r.cfm_companycode || '',
      name: r.cfm_companyname || '',
      entityId: r._cfm_type_value || '',
      entityName: formattedValue(r, '_cfm_type_value') || '—',
    }));
  }

  async function loadBankAccounts(): Promise<BankAccountOption[]> {
    // cfm_BankAccount is a lookup to cr301_accounts (the Accounts table), confirmed
    // against the live Dataverse relationship metadata — not a separate "bank
    // accounts" table, despite the field's display name.
    const rows = await fetchAll<Cr301_accounts>(
      (skipToken) => Cr301_accountsService.getAll({ select: BANK_ACCOUNT_SELECT, filter: 'statecode eq 0', maxPageSize: 5000, skipToken }),
      'Could not load bank accounts.'
    );
    return rows.map((r) => ({
      id: r.cr301_accountid,
      name: r.cr301_acccountname || '—',
      code: r.cr301_accountcode || '',
    }));
  }

  async function loadAllocCounts(ids: string[]): Promise<Record<string, number>> {
    if (!ids.length) return {};
    const rows = await fetchAll<Cfm_collectionmonthallocations>(
      (skipToken) => Cfm_collectionmonthallocationsService.getAll({ select: ALLOC_COUNT_SELECT, filter: 'statecode eq 0', maxPageSize: 5000, skipToken }),
      'Could not load allocation counts.'
    );
    const map: Record<string, number> = {};
    rows.forEach((r) => {
      const p = r._cfm_dailycollection_value;
      if (p) map[p] = (map[p] || 0) + 1;
    });
    return map;
  }

  async function load() {
    setLoading(true);
    setLoadError(null);
    try {
      const [companyOptions, bankAccountOptions, entryRows] = await Promise.all([
        loadCompanies(),
        loadBankAccounts(),
        fetchAll<Cfm_dailycollections>(
          (skipToken) => Cfm_dailycollectionsService.getAll({ select: ENTRY_SELECT, filter: 'statecode eq 0', orderBy: ['createdon desc'], maxPageSize: 5000, skipToken }),
          'Could not load daily collection entries.'
        ),
      ]);
      setCompanies(companyOptions);
      setBankAccounts(bankAccountOptions);
      const rows = entryRows.map((r) => normaliseEntry(r, companyOptions));
      const counts = await loadAllocCounts(rows.map((r) => r.id));
      rows.forEach((r) => {
        r.allocCount = counts[r.id] || 0;
      });
      setEntries(rows);
      setPage(1);
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
    getContext()
      .then(async (ctx) => {
        const objectId = ctx.user?.objectId;
        if (!objectId) {
          setCurrentUser({ id: '', name: ctx.user?.fullName || '' });
          return;
        }
        const res = await SystemusersService.getAll({ select: ['systemuserid'], filter: `azureactivedirectoryobjectid eq ${objectId}`, top: 1 });
        const id = res.success && res.data && res.data[0] ? res.data[0].systemuserid : '';
        setCurrentUser({ id, name: ctx.user?.fullName || '' });
      })
      .catch(() => {});
  }, []);

  function uniqueSorted(list: EntryRow[], pick: (r: EntryRow) => string): string[] {
    const set = new Set<string>();
    list.forEach((r) => {
      const v = pick(r);
      if (v && v !== '—') set.add(v);
    });
    return Array.from(set).sort();
  }

  const entityOptions = useMemo(() => uniqueSorted(entries, (r) => r.entity), [entries]);
  const methodOptions = useMemo(() => uniqueSorted(entries, (r) => r.methodLabel), [entries]);
  const statusOptions = useMemo(() => {
    const set = new Set<string>();
    entries.forEach((r) => {
      const l = displayStatusLabel(r);
      if (l && l !== '—') set.add(l);
    });
    return Array.from(set).sort();
  }, [entries]);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return entries.filter((r) => {
      if (rejectedOnly && r.status !== AR_STATUS.REJECTED) return false;
      if (byMe) {
        if (currentUser.id) {
          if (r.createdById !== currentUser.id) return false;
        } else if (currentUser.name) {
          if (r.createdBy.trim().toLowerCase() !== currentUser.name.trim().toLowerCase()) return false;
        }
      }
      if (entityFilter && r.entity !== entityFilter) return false;
      if (methodFilter && r.methodLabel !== methodFilter) return false;
      if (statusFilter && displayStatusLabel(r) !== statusFilter) return false;
      if (q && !`${r.ref} ${r.companyCode} ${r.companyName}`.toLowerCase().includes(q)) return false;
      return true;
    });
  }, [entries, query, entityFilter, methodFilter, statusFilter, byMe, rejectedOnly, currentUser]);

  useEffect(() => {
    setPage(1);
  }, [query, entityFilter, methodFilter, statusFilter, byMe, rejectedOnly]);

  const totalPages = pageSize === 'all' ? 1 : Math.max(1, Math.ceil(filtered.length / pageSize));
  const clampedPage = Math.min(page, totalPages);
  const pageRows = pageSize === 'all' ? filtered : filtered.slice((clampedPage - 1) * pageSize, (clampedPage - 1) * pageSize + pageSize);

  const kpis = useMemo(() => {
    const today = new Date().toISOString().slice(0, 10);
    const monthStart = today.slice(0, 7);
    let entriesToday = 0;
    let collectedToday = 0;
    let mtd = 0;
    let rejected = 0;
    entries.forEach((r) => {
      const d = (r.date || '').slice(0, 10);
      if (d === today) {
        entriesToday++;
        collectedToday += r.net;
      }
      if (d.startsWith(monthStart)) mtd += r.net;
      if (r.status === AR_STATUS.REJECTED) rejected++;
    });
    return { entriesToday, collectedToday, mtd, rejected };
  }, [entries]);

  function pagerButtons(): (number | 'ellipsis')[] {
    const pages: (number | 'ellipsis')[] = [];
    for (let p = 1; p <= totalPages; p++) {
      if (p === 1 || p === totalPages || Math.abs(p - clampedPage) <= 1) pages.push(p);
      else if (Math.abs(p - clampedPage) === 2) pages.push('ellipsis');
    }
    return pages;
  }

  const coMatches = useMemo(() => {
    const q = coQuery.trim().toLowerCase();
    if (!q) return [];
    return companies.filter((c) => `${c.code} ${c.name}`.toLowerCase().includes(q)).slice(0, 20);
  }, [companies, coQuery]);

  function pickCompany(c: CompanyOption) {
    setSelectedCo(c);
    setCoQuery(`${c.code} — ${c.name}`);
    setCoListOpen(false);
  }

  const bankMatches = useMemo(() => {
    const q = bankQuery.trim().toLowerCase();
    if (!q) return bankAccounts.slice(0, 20);
    return bankAccounts.filter((b) => bankAccountLabel(b).toLowerCase().includes(q)).slice(0, 20);
  }, [bankAccounts, bankQuery]);

  function pickBankAccount(b: BankAccountOption | null) {
    setFormBankAccountId(b ? b.id : '');
    setBankQuery(b ? bankAccountLabel(b) : '');
    setBankListOpen(false);
  }

  function grossOf(a: AllocRow): number {
    return N(a.amount) + N(a.tax) + N(a.vol) + N(a.early) + N(a.admin) + N(a.reject);
  }

  function addAllocRow(prefill?: Partial<AllocRow>) {
    setAllocs((prev) => [
      ...prev,
      {
        childId: prefill?.childId ?? null,
        month: prefill?.month ?? defaultMonthValue(),
        amount: prefill?.amount ?? 0,
        tax: prefill?.tax ?? 0,
        vol: prefill?.vol ?? 0,
        early: prefill?.early ?? 0,
        admin: prefill?.admin ?? 0,
        reject: prefill?.reject ?? 0,
      },
    ]);
  }

  function updateAllocRow(idx: number, field: keyof Omit<AllocRow, 'childId'>, value: string) {
    setAllocs((prev) => prev.map((a, i) => (i === idx ? { ...a, [field]: field === 'month' ? value : Number(value) || 0 } : a)));
  }

  function removeAllocRow(idx: number) {
    if (allocs.length === 1) {
      showToast('At least one allocation is required.');
      return;
    }
    setAllocs((prev) => prev.filter((_, i) => i !== idx));
  }

  const netVal = N(formNet);
  const sumAllocs = useMemo(() => allocs.reduce((t, a) => t + N(a.amount), 0), [allocs]);
  const sumOk = netVal > 0 && Math.abs(sumAllocs - netVal) < 0.01;

  // Allocations are normally read-only once an entry leaves 'create'/'reedit' — but
  // Rejected (reedit) and Pending AR Verification are still open for correction before
  // AR has acted on them, so 'edit' mode also unlocks allocation editing for those.
  const allocsEditable = mode !== 'edit' || editEntry?.status === AR_STATUS.PENDING_VERIFY;

  const canSave = useMemo(() => {
    if (saving || locked) return false;
    if (!formDate || formMethod === '' || netVal <= 0) return false;
    if (!allocsEditable) return true;
    return !!selectedCo && sumOk;
  }, [saving, locked, formDate, formMethod, netVal, allocsEditable, selectedCo, sumOk]);

  function openDrawerCreate() {
    setMode('create');
    setEditId(null);
    setEditEntry(null);
    setLocked(false);
    setSelectedCo(null);
    setCoQuery('');
    setCoListOpen(false);
    setFormDate(new Date().toISOString().slice(0, 10));
    setFormMethod('');
    setFormNet('');
    setFormPaymentRef('');
    setFormBankAccountId('');
    setBankQuery('');
    setBankListOpen(false);
    setFormNote('');
    setAllocs([{ childId: null, month: defaultMonthValue(), amount: 0, tax: 0, vol: 0, early: 0, admin: 0, reject: 0 }]);
    setOriginalChildIds([]);
    setAllocsError('');
    setFormError('');
    setDocs([]);
    setDocsError('');
    setDocsUploadingNames([]);
    setDrawerOpen(true);
  }

  async function openDrawerEdit(entry: EntryRow) {
    const entryLocked = isLocked(entry.status);
    const reedit = entry.status === AR_STATUS.REJECTED;
    setMode(reedit ? 'reedit' : 'edit');
    setEditId(entry.id);
    setEditEntry(entry);
    setLocked(entryLocked);
    setSelectedCo(entry.companyId ? { id: entry.companyId, code: entry.companyCode, name: entry.companyName, entityId: '', entityName: entry.entity } : null);
    setCoQuery(entryLocked ? '' : `${entry.companyCode} — ${entry.companyName}`);
    setCoListOpen(false);
    setFormDate(entry.date ? entry.date.slice(0, 10) : '');
    setFormMethod(entry.method == null ? '' : String(entry.method));
    setFormNet(entry.net ? String(entry.net) : '');
    setFormPaymentRef(entry.paymentReference || '');
    setFormBankAccountId(entry.bankAccountId || '');
    const currentBank = entry.bankAccountId ? bankAccounts.find((b) => b.id === entry.bankAccountId) : undefined;
    setBankQuery(currentBank ? bankAccountLabel(currentBank) : '');
    setBankListOpen(false);
    setFormNote(entry.note || '');
    setFormError('');
    setAllocs([]);
    setOriginalChildIds([]);
    setAllocsError('');
    setAllocsLoading(true);
    setDocs([]);
    setDocsError('');
    setDocsUploadingNames([]);
    setDrawerOpen(true);
    loadDocs(entry);

    try {
      const res = await Cfm_collectionmonthallocationsService.getAll({
        select: ALLOC_DETAIL_SELECT,
        filter: `statecode eq 0 and _cfm_dailycollection_value eq ${entry.id}`,
        orderBy: ['cfm_allocatetomonth'],
      });
      if (!res.success) throw new Error(operationErrorMessage(res.error, 'Could not load allocations.'));
      const rows = res.data || [];
      setOriginalChildIds(rows.map((a) => a.cfm_collectionmonthallocationid).filter(Boolean));
      const asRows = (list: Cfm_collectionmonthallocations[]): AllocRow[] =>
        list.map((a) => ({
          childId: a.cfm_collectionmonthallocationid,
          month: (a.cfm_allocatetomonth || '').slice(0, 10),
          amount: N(a.cfm_allocationamount),
          tax: N(a.cfm_tax),
          vol: N(a.cfm_volumediscount),
          early: N(a.cfm_earlypaymentdisc),
          admin: N(a.cfm_adminfees),
          reject: N(a.cfm_rejection),
        }));
      if (reedit) {
        setAllocs(rows.length ? asRows(rows) : [{ childId: null, month: defaultMonthValue(), amount: 0, tax: 0, vol: 0, early: 0, admin: 0, reject: 0 }]);
      } else {
        setAllocs(asRows(rows));
      }
    } catch (e) {
      setAllocsError(e instanceof Error ? e.message : String(e));
    } finally {
      setAllocsLoading(false);
    }
  }

  function closeDrawer() {
    setDrawerOpen(false);
  }

  async function loadDocs(forEntry?: EntryRow | null) {
    const entry = forEntry ?? editEntry;
    if (!entry) return;
    const folder = docsFolderName(entry);
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

  // After the first successful upload, promote Awaiting documents (or legacy null status)
  // to Pending AR Verification — both in Dataverse and in the local table/drawer state.
  async function promoteAfterUpload() {
    if (!editId || !editEntry) return;
    if (editEntry.status !== AR_STATUS.AWAITING_DOCS && editEntry.status != null) return;
    try {
      const result = await Cfm_dailycollectionsService.update(editId, { cfm_statusaction: AR_STATUS.PENDING_VERIFY });
      if (!result.success) throw new Error(operationErrorMessage(result.error, 'Could not update status.'));
      const patch = { status: AR_STATUS.PENDING_VERIFY, statusLabel: AR_STATUS_LABEL[AR_STATUS.PENDING_VERIFY] };
      setEditEntry((prev) => (prev ? { ...prev, ...patch } : prev));
      setEntries((prev) => prev.map((r) => (r.id === editId ? { ...r, ...patch } : r)));
      showToast('Status updated: moved to Pending AR Verification.');
    } catch (err) {
      showToast(`Documents uploaded, but status could not change: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  async function onDocsSelected(files: File[]) {
    if (locked) {
      showToast('This entry is locked — cannot upload.');
      return;
    }
    if (!files.length || !editEntry) return;
    const folder = docsFolderName(editEntry);
    let ok = 0;
    let fail = 0;
    let firstErr = '';
    for (const file of files) {
      setDocsUploadingNames((prev) => [...prev, file.name]);
      try {
        const b64 = await fileToBase64(file);
        await uploadDoc({
          folderName: folder,
          collectionId: editEntry.id.replace(/[{}]/g, ''),
          collectionRef: editEntry.ref,
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
    if (ok) await promoteAfterUpload();
    await loadDocs();
  }

  async function deleteDoc(path: string, name: string) {
    if (locked || !editEntry) return;
    if (!window.confirm(`Delete "${name}"? This cannot be undone.`)) return;
    const folder = docsFolderName(editEntry);
    try {
      await docsFlow({ action: 'delete', folderName: folder, filePath: path });
      showToast(`Deleted ${name}.`);
      await loadDocs();
    } catch (err) {
      showToast(`Delete failed: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  async function saveEntry() {
    if (saving) return;
    setSaving(true);
    setFormError('');
    try {
      const parentBody: Record<string, unknown> = {};
      if (selectedCo) parentBody['cfm_CompanyCode@odata.bind'] = `/cfm_aragings(${selectedCo.id})`;
      if (formBankAccountId) parentBody['cfm_BankAccount@odata.bind'] = `/cr301_accounts(${formBankAccountId})`;
      parentBody.cfm_collectiondate = formDate ? `${formDate}T00:00:00Z` : null;
      parentBody.cfm_paymentmethod = formMethod === '' ? null : parseInt(formMethod, 10);
      parentBody.cfm_collectedamountnet = netVal;
      parentBody.cfm_paymentreference = formPaymentRef.trim() || null;
      parentBody.cfm_note = formNote.trim() || null;

      if (mode === 'edit' || mode === 'reedit') {
        if (!editId) throw new Error('Missing entry id.');
        if (mode === 'reedit') parentBody.cfm_statusaction = AR_STATUS.PENDING_VERIFY;
        const result = await Cfm_dailycollectionsService.update(editId, parentBody as Partial<Omit<Cfm_dailycollectionsBase, 'cfm_dailycollectionid'>>);
        if (!result.success) throw new Error(operationErrorMessage(result.error, 'Update failed.'));

        if (!allocsEditable) {
          showToast('Entry updated.');
          closeDrawer();
          await load();
          return;
        }

        const keptIds = new Set(allocs.map((a) => a.childId).filter((v): v is string => !!v));
        let childFail = 0;
        let firstErr = '';
        for (const a of allocs) {
          const amt = N(a.amount);
          const body: Record<string, unknown> = {
            cfm_allocatetomonth: a.month || null,
            cfm_allocatetoclaim: amt,
            cfm_allocationamount: amt,
            cfm_tax: N(a.tax),
            cfm_volumediscount: N(a.vol),
            cfm_earlypaymentdisc: N(a.early),
            cfm_adminfees: N(a.admin),
            cfm_rejection: N(a.reject),
          };
          try {
            if (a.childId) {
              const r = await Cfm_collectionmonthallocationsService.update(a.childId, body as Partial<Omit<Cfm_collectionmonthallocationsBase, 'cfm_collectionmonthallocationid'>>);
              if (!r.success) throw new Error(operationErrorMessage(r.error, 'Allocation update failed.'));
            } else {
              body['cfm_DailyCollection@odata.bind'] = `/cfm_dailycollections(${editId})`;
              const r = await Cfm_collectionmonthallocationsService.create(body as unknown as Omit<Cfm_collectionmonthallocationsBase, 'cfm_collectionmonthallocationid'>);
              if (!r.success) throw new Error(operationErrorMessage(r.error, 'Allocation create failed.'));
            }
          } catch (err) {
            childFail++;
            if (!firstErr) firstErr = err instanceof Error ? err.message : String(err);
          }
        }
        for (const oldId of originalChildIds) {
          if (!keptIds.has(oldId)) {
            try {
              await Cfm_collectionmonthallocationsService.delete(oldId);
            } catch (err) {
              childFail++;
              if (!firstErr) firstErr = err instanceof Error ? err.message : String(err);
            }
          }
        }
        showToast(
          mode === 'reedit'
            ? childFail > 0
              ? `Sent back to AR, but ${childFail} allocation change(s) failed: ${firstErr}`
              : 'Sent back to AR as Pending AR Verification.'
            : childFail > 0
              ? `Entry updated, but ${childFail} allocation change(s) failed: ${firstErr}`
              : 'Entry and allocations updated.'
        );
        closeDrawer();
        await load();
        return;
      }

      // create — new entries start as Awaiting documents until the first upload
      parentBody.cfm_statusaction = AR_STATUS.AWAITING_DOCS;
      const parentResult = await Cfm_dailycollectionsService.create(parentBody as unknown as Omit<Cfm_dailycollectionsBase, 'cfm_dailycollectionid'>);
      if (!parentResult.success) throw new Error(operationErrorMessage(parentResult.error, 'Save failed.'));
      const parentId = parentResult.data?.cfm_dailycollectionid;
      if (!parentId) throw new Error('Entry saved but no id was returned.');

      let childOk = 0;
      let childFail = 0;
      let firstErr = '';
      for (const a of allocs) {
        const amt = N(a.amount);
        const body: Record<string, unknown> = {
          'cfm_DailyCollection@odata.bind': `/cfm_dailycollections(${parentId})`,
          cfm_allocatetomonth: a.month || null,
          cfm_allocatetoclaim: amt,
          cfm_allocationamount: amt,
          cfm_tax: N(a.tax),
          cfm_volumediscount: N(a.vol),
          cfm_earlypaymentdisc: N(a.early),
          cfm_adminfees: N(a.admin),
          cfm_rejection: N(a.reject),
        };
        try {
          const r = await Cfm_collectionmonthallocationsService.create(body as unknown as Omit<Cfm_collectionmonthallocationsBase, 'cfm_collectionmonthallocationid'>);
          if (!r.success) throw new Error(operationErrorMessage(r.error, 'Allocation save failed.'));
          childOk++;
        } catch (err) {
          childFail++;
          if (!firstErr) firstErr = err instanceof Error ? err.message : String(err);
        }
      }
      if (childFail > 0 && childOk === 0) {
        try {
          await Cfm_dailycollectionsService.delete(parentId);
        } catch {
          /* best-effort rollback */
        }
        throw new Error(`Allocations failed to save: ${firstErr}`);
      }

      const arAgingError = await addCollectedAmountToArAging(parentId);

      showToast(
        `Saved as Awaiting documents. ${childOk} allocation${childOk === 1 ? '' : 's'} recorded${childFail ? `, ${childFail} failed` : ''}. Upload a document to move it to Pending AR Verification.` +
          (arAgingError ? ` AR Aging update failed: ${arAgingError}` : '')
      );
      closeDrawer();
      await load();
    } catch (e) {
      setFormError(e instanceof Error ? e.message : String(e));
    } finally {
      setSaving(false);
    }
  }

  return (
    <>
      <div className="acc-wrap">
        <div className="acc-head">
          <div>
            <h1 className="acc-title">Daily Collection</h1>
            <div className="acc-sub">Every collection recorded by the team, and one place to enter a new one.</div>
          </div>
          <div className="acc-actions">
            <button className="acc-btn" onClick={load}>
              <i className="fa-solid fa-rotate"></i> Reload
            </button>
            <button className="acc-btn acc-btn-primary" onClick={openDrawerCreate}>
              <i className="fa-solid fa-plus"></i> New entry
            </button>
          </div>
        </div>

        <div className="kpi-grid">
          <div className="kpi-card">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-gold">
                <i className="fa-solid fa-list-check"></i>
              </span>
              <span className="kpi-label">Entries today</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.entriesToday}</div>
            <div className="kpi-context">&nbsp;</div>
          </div>
          <div className="kpi-card kpi-ok">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-ok">
                <i className="fa-solid fa-hand-holding-dollar"></i>
              </span>
              <span className="kpi-label">Collected today</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.collectedToday.toLocaleString('en-US')}</div>
            <div className="kpi-context">&nbsp;</div>
          </div>
          <div className="kpi-card kpi-info">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-info">
                <i className="fa-solid fa-calendar-days"></i>
              </span>
              <span className="kpi-label">Month to date</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.mtd.toLocaleString('en-US')}</div>
            <div className="kpi-context">&nbsp;</div>
          </div>
          <div
            className="kpi-card kpi-danger kpi-clickable"
            onClick={() => setRejectedOnly((v) => !v)}
            title="Click to show only rejected entries"
          >
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-danger">
                <i className="fa-solid fa-circle-xmark"></i>
              </span>
              <span className="kpi-label">Rejected · action needed</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.rejected}</div>
            <div className="kpi-context">{rejectedOnly ? 'Filtering — click to clear' : ' '}</div>
          </div>
        </div>

        <div className="tbl-shell">
          <div className="tbl-hdr">
            <div className="dc-filter-bar">
              <button className={`dc-toggle${byMe ? ' active' : ''}`} onClick={() => setByMe((v) => !v)} title="Show only entries I created">
                <i className="fa-solid fa-user"></i> By me
              </button>
              <div className="acc-search dc-search">
                <i className="fa-solid fa-magnifying-glass"></i>
                <input placeholder="Search company, code, ref…" value={query} onChange={(e) => setQuery(e.target.value)} />
              </div>
              <select className="acc-f" value={entityFilter} onChange={(e) => setEntityFilter(e.target.value)}>
                <option value="">All entities</option>
                {entityOptions.map((v) => (
                  <option key={v} value={v}>
                    {v}
                  </option>
                ))}
              </select>
              <select className="acc-f" value={methodFilter} onChange={(e) => setMethodFilter(e.target.value)}>
                <option value="">All methods</option>
                {methodOptions.map((v) => (
                  <option key={v} value={v}>
                    {v}
                  </option>
                ))}
              </select>
              <select className="acc-f" value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)}>
                <option value="">All AR status</option>
                {statusOptions.map((v) => (
                  <option key={v} value={v}>
                    {v}
                  </option>
                ))}
              </select>
            </div>
            <div className="dc-count">
              <i className="fa-solid fa-table"></i> Showing <strong>{filtered.length}</strong> of <strong>{entries.length}</strong>
            </div>
          </div>

          <div className="dc-tablewrap">
            <table>
              <thead>
                <tr>
                  <th>Ref</th>
                  <th>Entity</th>
                  <th>Company code</th>
                  <th>Company name</th>
                  <th>Collection date</th>
                  <th>Payment method</th>
                  <th className="dc-numeric">Collected (net)</th>
                  <th>Allocations</th>
                  <th>Status &amp; proof</th>
                  <th>Created by</th>
                  <th>Created on</th>
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
                        <i className="fa-solid fa-triangle-exclamation"></i>
                        Could not load entries.
                        <br />
                        <small>{loadError}</small>
                      </div>
                    </td>
                  </tr>
                ) : pageRows.length === 0 ? (
                  <tr>
                    <td colSpan={11}>
                      <div className="acc-state">
                        <i className="fa-solid fa-inbox"></i>
                        No entries match your filters.
                      </div>
                    </td>
                  </tr>
                ) : (
                  pageRows.map((r) => {
                    const rowLocked = isLocked(r.status);
                    const rejected = r.status === AR_STATUS.REJECTED;
                    const dispLabel = displayStatusLabel(r);
                    const isAwaiting = r.status === AR_STATUS.AWAITING_DOCS || r.status == null;
                    const rowCls = rejected ? ' dc-row-rejected' : rowLocked ? ' dc-row-locked' : '';
                    return (
                      <tr key={r.id} className={`dc-row${rowCls}`} onClick={() => openDrawerEdit(r)}>
                        <td className="dc-ref">
                          {rejected ? (
                            <i className="fa-solid fa-triangle-exclamation dc-reject-ico" title="Rejected by AR"></i>
                          ) : rowLocked ? (
                            <i className="fa-solid fa-lock dc-lock-ico" title="Locked by AR"></i>
                          ) : null}
                          {r.ref}
                        </td>
                        <td>
                          <span className="dc-pill dc-pill-entity">{r.entity}</span>
                        </td>
                        <td>
                          <span className="dc-code">{r.companyCode}</span>
                        </td>
                        <td className="dc-company">{r.companyName}</td>
                        <td>{fmtDate(r.date)}</td>
                        <td>
                          <span className="dc-pill dc-pill-method">{r.methodLabel}</span>
                        </td>
                        <td className="dc-numeric dc-amt">{fmt(r.net)}</td>
                        <td>{r.allocCount ? `${r.allocCount} month${r.allocCount === 1 ? '' : 's'}` : <span className="dc-muted">—</span>}</td>
                        <td>
                          {rejected ? (
                            <span className="dc-pill dc-pill-rejected-strong">
                              <i className="fa-solid fa-rotate-left"></i> {r.statusLabel} — re-work
                            </span>
                          ) : isAwaiting ? (
                            <span className="dc-pill dc-pill-awaiting">
                              <i className="fa-solid fa-file-circle-plus"></i> Awaiting documents
                            </span>
                          ) : (
                            <span className={`dc-pill ${statusPillClass(r.status)}`}>
                              <i className={`fa-solid ${statusIcon(r.status)}`}></i> {dispLabel}
                            </span>
                          )}
                        </td>
                        <td>
                          <span className="dc-createdby">
                            <i className="fa-solid fa-user"></i>
                            {r.createdBy}
                          </span>
                        </td>
                        <td className="dc-createdon">{fmtDateTime(r.createdOn)}</td>
                      </tr>
                    );
                  })
                )}
              </tbody>
            </table>
          </div>

          {!loading && !loadError && filtered.length > 0 && (
            <div className="dc-pagination">
              <div className="dc-page-size">
                <span>Rows per page:</span>
                <select
                  className="acc-f"
                  value={pageSize}
                  onChange={(e) => {
                    setPageSize(e.target.value === 'all' ? 'all' : parseInt(e.target.value, 10));
                    setPage(1);
                  }}
                >
                  <option value={25}>25</option>
                  <option value={50}>50</option>
                  <option value={100}>100</option>
                  <option value="all">All</option>
                </select>
              </div>
              <div className="dc-page-info">
                Showing <strong>{pageSize === 'all' ? 1 : (clampedPage - 1) * pageSize + 1}</strong>–
                <strong>{pageSize === 'all' ? filtered.length : Math.min(clampedPage * pageSize, filtered.length)}</strong> of <strong>{filtered.length}</strong>
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

      <div className={`acc-scrim${drawerOpen ? ' on' : ''}`} onClick={closeDrawer}></div>
      <aside className={`acc-drawer dc-drawer${drawerOpen ? ' on' : ''}`}>
        <div className="acc-dhead">
          <h2>
            {locked && <i className="fa-solid fa-lock" style={{ color: 'var(--gold)', fontSize: 15, marginRight: 6 }}></i>}
            {mode === 'create' ? 'New · Daily Collection entry' : locked ? `View · ${editEntry?.ref}` : mode === 'reedit' ? `Re-submit · ${editEntry?.ref}` : `Edit · ${editEntry?.ref}`}
          </h2>
          <button className="acc-dclose" onClick={closeDrawer}>
            <i className="fa-solid fa-xmark"></i>
          </button>
        </div>
        <div className="acc-dbody">
          <div className={`acc-err${formError ? ' on' : ''}`}>{formError}</div>

          {!locked && (
            <div className="acc-combo dc-co-wrap">
              <div className="acc-fld" style={{ marginBottom: 0 }}>
                <label>
                  Company <span className="dc-req">*</span>
                  <span className="dc-co-hint">— from AR Aging (Plan required)</span>
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
              {coListOpen && coQuery.trim() && (
                <div className="acc-combo-list">
                  {coMatches.length === 0 ? (
                    <div className="acc-combo-empty">No matches — company must exist in AR Aging.</div>
                  ) : (
                    coMatches.map((c) => (
                      <div key={c.id} className="acc-combo-item" onMouseDown={(e) => e.preventDefault()} onClick={() => pickCompany(c)}>
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
          )}

          <div className={`dc-co-panel${selectedCo ? '' : ' empty'}`}>
            {selectedCo ? (
              <>
                <div>
                  <div className="k">Company name</div>
                  <div className="v">{selectedCo.name || '—'}</div>
                </div>
                <div>
                  <div className="k">Entity</div>
                  <div className="v">{selectedCo.entityName || '—'}</div>
                </div>
              </>
            ) : (
              <div>
                <i className="fa-solid fa-circle-info" style={{ color: 'var(--gold)', marginRight: 6 }}></i>
                Pick a company — name &amp; entity fill in automatically.
              </div>
            )}
          </div>

          <div className="dc-fld-row">
            <div className="acc-fld">
              <label>
                Collection date <span className="dc-req">*</span>
              </label>
              <input type="date" value={formDate} disabled={locked} onChange={(e) => setFormDate(e.target.value)} />
            </div>
            <div className="acc-fld">
              <label>
                Payment method <span className="dc-req">*</span>
              </label>
              <select value={formMethod} disabled={locked} onChange={(e) => setFormMethod(e.target.value)}>
                <option value="">Select…</option>
                {PAYMENT_METHOD_OPTIONS.map((o) => (
                  <option key={o.value} value={o.value}>
                    {o.label}
                  </option>
                ))}
              </select>
            </div>
          </div>

          <div className="dc-fld-row">
            <div className="acc-fld">
              <label>
                Collected amount (net) <span className="dc-req">*</span>
              </label>
              <input type="number" step="any" min={0} placeholder="0.00" value={formNet} disabled={locked} onChange={(e) => setFormNet(e.target.value)} />
            </div>
            <div className="acc-fld">
              <label>Payment reference #</label>
              <input placeholder="e.g. TT-887421" value={formPaymentRef} disabled={locked} onChange={(e) => setFormPaymentRef(e.target.value)} />
            </div>
          </div>

          <div className="acc-combo dc-co-wrap">
            <div className="acc-fld" style={{ marginBottom: 0 }}>
              <label>Bank account</label>
              <input
                placeholder="Search bank name or account number…"
                value={bankQuery}
                disabled={locked}
                onChange={(e) => {
                  setBankQuery(e.target.value);
                  setBankListOpen(true);
                  setFormBankAccountId('');
                }}
                onFocus={() => setBankListOpen(true)}
              />
            </div>
            {bankListOpen && (
              <div className="acc-combo-list">
                <div className="acc-combo-item" onMouseDown={(e) => e.preventDefault()} onClick={() => pickBankAccount(null)}>
                  <div>None</div>
                </div>
                {bankMatches.length === 0 ? (
                  <div className="acc-combo-empty">No matches.</div>
                ) : (
                  bankMatches.map((b) => (
                    <div key={b.id} className="acc-combo-item" onMouseDown={(e) => e.preventDefault()} onClick={() => pickBankAccount(b)}>
                      <div>{bankAccountLabel(b)}</div>
                    </div>
                  ))
                )}
              </div>
            )}
          </div>

          <div className="acc-fld">
            <label>Notes</label>
            <textarea rows={2} placeholder="Any note about this entry…" value={formNote} disabled={locked} onChange={(e) => setFormNote(e.target.value)} />
          </div>

          <div className="dc-alloc-hdr">
            <div className="dc-alloc-title">
              <i className="fa-solid fa-layer-group"></i> Month allocations
            </div>
            {allocsEditable && (
              <button className="dc-alloc-add" disabled={locked} onClick={() => addAllocRow()}>
                <i className="fa-solid fa-plus"></i> Add month
              </button>
            )}
          </div>

          {allocsLoading ? (
            <div className="dc-alloc-readonly-note">
              <i className="fa-solid fa-spinner fa-spin"></i> Loading allocations…
            </div>
          ) : allocsError ? (
            <div className="dc-alloc-readonly-note" style={{ color: 'var(--bad)' }}>
              {allocsError}
            </div>
          ) : allocs.length === 0 ? (
            <div className="dc-alloc-readonly-note">No allocations found.</div>
          ) : (
            <div className="dc-alloc-list">
              {allocs.map((a, idx) => {
                const gross = grossOf(a);
                const readOnlyCard = !allocsEditable;
                return (
                  <div key={idx} className="dc-alloc-card">
                    <div className="dc-alloc-card-hdr">
                      <select value={a.month} disabled={readOnlyCard || locked} onChange={(e) => updateAllocRow(idx, 'month', e.target.value)}>
                        {MONTH_OPTIONS.map((o) => (
                          <option key={o.value} value={o.value}>
                            {o.label}
                          </option>
                        ))}
                      </select>
                      {!readOnlyCard && !locked && (
                        <button className="dc-alloc-del" onClick={() => removeAllocRow(idx)} aria-label="Remove">
                          <i className="fa-solid fa-trash"></i>
                        </button>
                      )}
                    </div>
                    <div className="dc-alloc-card-body">
                      <div className="dc-alloc-fg">
                        <div className="dc-alloc-fg-lbl">Amount</div>
                        <input className="dc-alloc-inp" type="number" step="any" value={a.amount || ''} disabled={readOnlyCard || locked} onChange={(e) => updateAllocRow(idx, 'amount', e.target.value)} />
                      </div>
                      <div className="dc-alloc-fg">
                        <div className="dc-alloc-fg-lbl">Tax</div>
                        <input className="dc-alloc-inp" type="number" step="any" value={a.tax || ''} disabled={readOnlyCard || locked} onChange={(e) => updateAllocRow(idx, 'tax', e.target.value)} />
                      </div>
                      <div className="dc-alloc-fg">
                        <div className="dc-alloc-fg-lbl">Volume discount</div>
                        <input className="dc-alloc-inp" type="number" step="any" value={a.vol || ''} disabled={readOnlyCard || locked} onChange={(e) => updateAllocRow(idx, 'vol', e.target.value)} />
                      </div>
                      <div className="dc-alloc-fg">
                        <div className="dc-alloc-fg-lbl">Early payment disc</div>
                        <input className="dc-alloc-inp" type="number" step="any" value={a.early || ''} disabled={readOnlyCard || locked} onChange={(e) => updateAllocRow(idx, 'early', e.target.value)} />
                      </div>
                      <div className="dc-alloc-fg">
                        <div className="dc-alloc-fg-lbl">Admin fees</div>
                        <input className="dc-alloc-inp" type="number" step="any" value={a.admin || ''} disabled={readOnlyCard || locked} onChange={(e) => updateAllocRow(idx, 'admin', e.target.value)} />
                      </div>
                      <div className="dc-alloc-fg">
                        <div className="dc-alloc-fg-lbl">Rejection</div>
                        <input className="dc-alloc-inp" type="number" step="any" value={a.reject || ''} disabled={readOnlyCard || locked} onChange={(e) => updateAllocRow(idx, 'reject', e.target.value)} />
                      </div>
                      <div className="dc-alloc-fg wide">
                        <div className="dc-alloc-fg-lbl">
                          Gross amount <span className="auto">(auto: amount + tax + vol + early + admin + rejection)</span>
                        </div>
                        <input className="dc-alloc-inp readonly" readOnly value={fmt(gross)} />
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {!locked && allocsEditable && (
            <div className={`dc-sum-pill ${sumOk ? 'ok' : 'bad'}`}>
              <span>
                <i className={`fa-solid ${sumOk ? 'fa-circle-check' : 'fa-circle-exclamation'}`}></i>{' '}
                {sumOk ? 'Σ allocations match collected net' : 'Σ allocations must equal collected net'}
              </span>
              <span className="amounts">
                {fmt(sumAllocs)} / {fmt(netVal)}
              </span>
            </div>
          )}
          {!allocsEditable && (
            <div className="dc-sum-pill ok">
              <span>
                <i className="fa-solid fa-circle-info"></i> Allocations are read-only in this status
              </span>
              <span className="amounts">
                {fmt(sumAllocs)} / {fmt(netVal)}
              </span>
            </div>
          )}

          {mode === 'create' ? (
            <div className="dc-note warn">
              <i className="fa-solid fa-file-circle-plus"></i>
              <span>
                This entry will be saved as <strong>Awaiting documents</strong>. Upload the first document to move it to <strong>Pending AR Verification</strong>.
              </span>
            </div>
          ) : locked ? (
            <div className="dc-note bad">
              <i className="fa-solid fa-lock"></i>
              <span>{(editEntry?.status != null && LOCK_REASON[editEntry.status]) || 'This entry is locked.'}</span>
            </div>
          ) : mode === 'reedit' ? (
            <div className="dc-note bad">
              <i className="fa-solid fa-circle-xmark"></i>
              <span>
                <strong>Rejected by AR.</strong> Fix the details and re-submit — the entry will go back to AR as <strong>Pending AR Verification</strong>.
              </span>
            </div>
          ) : editEntry?.status === AR_STATUS.AWAITING_DOCS || editEntry?.status == null ? (
            <div className="dc-note warn">
              <i className="fa-solid fa-file-circle-plus"></i>
              <span>
                <strong>Awaiting documents.</strong> Upload the first document below to move this entry to <strong>Pending AR Verification</strong>.
              </span>
            </div>
          ) : editEntry?.proof ? (
            <div className="dc-note">
              <i className="fa-solid fa-link" style={{ color: 'var(--ok)' }}></i>
              <span>
                Proof:{' '}
                <a href={editEntry.proof} target="_blank" rel="noopener noreferrer">
                  {editEntry.proof}
                </a>
              </span>
            </div>
          ) : (
            <div className="dc-note">
              <i className="fa-solid fa-file-signature"></i>
              <span>The proof link is created on OneDrive automatically after the entry is saved.</span>
            </div>
          )}

          {mode !== 'create' && editEntry && (
            <div className="dc-docs-sec">
              <div className="dc-docs-hdr">
                <div className="dc-docs-title">
                  <i className="fa-solid fa-paperclip"></i> Documents
                  {docs.length > 0 && <span className="dc-docs-count">({docs.length})</span>}
                </div>
                {!locked && (
                  <label className="dc-docs-upload-btn">
                    <i className="fa-solid fa-upload"></i> Upload
                    <input
                      type="file"
                      multiple
                      onChange={(e) => {
                        // Snapshot into a real array before clearing the input — some
                        // browsers reuse/clear the live FileList object on reset, which
                        // would silently empty this same reference if read afterwards.
                        const files = Array.from(e.target.files || []);
                        e.target.value = '';
                        onDocsSelected(files);
                      }}
                    />
                  </label>
                )}
              </div>
              <div className="dc-docs-list">
                {docsLoading ? (
                  <div className="dc-doc-uploading">
                    <i className="fa-solid fa-spinner fa-spin"></i> Loading documents…
                  </div>
                ) : docsError ? (
                  <div className="dc-docs-empty">
                    <i className="fa-solid fa-circle-info"></i> Couldn&apos;t load documents right now.{!locked && ' You can still upload — click Upload.'}
                  </div>
                ) : docsUploadingNames.length > 0 || docs.length > 0 ? (
                  <>
                    {docsUploadingNames.map((n, i) => (
                      <div key={`up-${i}-${n}`} className="dc-doc-uploading">
                        <i className="fa-solid fa-spinner fa-spin"></i> Uploading {n}…
                      </div>
                    ))}
                    {docs.map((f) => {
                      const ext = docExt(f.name);
                      return (
                        <div key={f.path || f.name} className="dc-doc-row">
                          <div className={`dc-doc-ico dc-doc-ico-${docIconClass(ext)}`}>
                            <i className={`fa-solid ${docIcon(ext)}`}></i>
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
                              <i className="fa-solid fa-arrow-up-right-from-square"></i>
                            </a>
                            {!locked && (
                              <button title="Delete" onClick={() => deleteDoc(f.path, f.name)}>
                                <i className="fa-solid fa-trash"></i>
                              </button>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </>
                ) : (
                  <div className="dc-docs-empty">
                    <i className="fa-solid fa-folder-open"></i> {locked ? 'No documents attached.' : 'No documents yet — click Upload to attach files.'}
                  </div>
                )}
              </div>
              <div className="dc-docs-note">{locked ? 'This entry is locked — documents are view-only.' : 'Any file type. Stored on SharePoint, not in Dataverse.'}</div>
            </div>
          )}
        </div>
        <div className="acc-dfoot">
          <div className="acc-hint">
            <i className="fa-solid fa-triangle-exclamation" style={{ color: 'var(--warn)', marginRight: 4 }}></i>
            Company must exist in AR Aging
          </div>
          <div style={{ display: 'flex', gap: 8 }}>
            <button className="acc-btn acc-btn-ghost" onClick={closeDrawer}>
              Cancel
            </button>
            {!locked && (
              <button className="acc-btn acc-btn-primary" disabled={!canSave} onClick={saveEntry}>
                {saving ? 'Saving…' : mode === 'create' ? 'Save entry' : mode === 'reedit' ? 'Re-submit entry' : 'Update entry'}
              </button>
            )}
          </div>
        </div>
      </aside>

      <div className={`acc-toast${toastOn ? ' on' : ''}`}>{toastMsg}</div>
    </>
  );
}
