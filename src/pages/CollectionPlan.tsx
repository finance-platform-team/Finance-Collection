import Icon from '../components/Icon';
﻿import { useEffect, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import { Cfm_aragingsService } from '../generated/services/Cfm_aragingsService';
import { Cfm_insurancecompaniesService } from '../generated/services/Cfm_insurancecompaniesService';
import { Cfm_companymonthlybalancesesService } from '../generated/services/Cfm_companymonthlybalancesesService';
import { SystemusersService } from '../generated/services/SystemusersService';
import type { Cfm_aragings, Cfm_aragingsBase } from '../generated/models/Cfm_aragingsModel';
import type { Cfm_insurancecompanies } from '../generated/models/Cfm_insurancecompaniesModel';
import type { Cfm_companymonthlybalanceses } from '../generated/models/Cfm_companymonthlybalancesesModel';
import { operationErrorMessage } from '../lib/errors';
import { KSA_BU_CODES, type Region } from '../lib/useRegion';
import { escapeForCsv, downloadCsv } from '../lib/csvExport';
import './CollectionPlan.css';

/**
 * Payment Term, Customer Class and Company Type are NOT read from cfm_aragings even
 * though that table carries its own copies of those lookups — Ref.html's join rule (the
 * insurance company record is the source of truth) is preserved here: values are resolved
 * from cfm_insurancecompanies by matching cfm_unique (code+BU), falling back to code+BU
 * text and then to code alone.
 *
 * The generated models advertise "name" shadow properties for lookups/choices (e.g.
 * cfm_typename, cfm_taskownername, cfm_paymenttermname, cfm_customerclassname) as if they
 * were selectable columns, but Dataverse rejects them in $select with a 400. The display
 * text lives on the real attribute's own `@OData.Community.Display.V1.FormattedValue`
 * annotation instead — select the real attribute and read that annotation directly.
 */
const ARAGING_SELECT = [
  'cfm_aragingid',
  'cfm_companycode',
  'cfm_companyname',
  'cfm_unique',
  '_cfm_type_value',
  'cfm_earlypayment',
  'cfm_31to60',
  'cfm_61to90',
  'cfm_91to120',
  'cfm_150days',
  'cfm_agreedrecon',
  'cfm_totaldues',
  'cfm_targetplan',
  'cfm_collectedamount',
  'cfm_collectedtaxrej',
  'cfm_tax',
  'cfm_rejections',
  'cfm_legalissues',
  'cfm_bankruptcy',
  'cfm_claimsissue',
  'cfm_stopped',
  'cfm_notdue',
  'cfm_pendingconfirmation',
  '_cfm_taskowner_value',
  '_cfm_supervisor_value',
];

const INSURANCE_SELECT = ['cfm_insurancecompanyid', 'cfm_code', 'cfm_butext', 'cfm_unique', '_cfm_paymentterm_value', 'cfm_customerclass', 'cfm_companytype'];

/**
 * cfm_businessunit is a lookup (same pattern as cfm_aragings' cfm_type — see the note at the
 * top of this file): the "cfm_businessunitname" shadow property the generated model advertises
 * is rejected by Dataverse in $select. Select the real lookup attribute and read the BU text off
 * its @OData.Community.Display.V1.FormattedValue annotation instead, same as buText below.
 */
const MONTHLY_BALANCE_SELECT = ['cfm_companymonthlybalancesid', 'cfm_companycode', '_cfm_businessunit_value', 'cfm_month', 'cfm_balance', 'cfm_totalclaims'];

function formattedValue(record: unknown, attribute: string): string {
  const raw = (record as Record<string, unknown>)[`${attribute}@OData.Community.Display.V1.FormattedValue`];
  return typeof raw === 'string' ? raw : '';
}

interface InsuranceRecord {
  paymentTerm: string;
  customerClass: string;
  companyType: string;
}

interface MonthlyBalanceRecord {
  monthLabel: string;
  monthShort: string;
  sortKey: number;
  balance: number;
  totalClaims: number;
}

interface PlanRow {
  id: string;
  code: string;
  name: string;
  unique: string;
  buText: string;
  taskOwnerId: string;
  taskOwner: string;
  supervisorId: string;
  supervisor: string;
  early: number;
  d60: number;
  d90: number;
  d120: number;
  over120: number;
  recon: number;
  totalDues: number;
  target: number;
  collected: number;
  collectedPlus: number;
  pending: number;
  remDues: number;
  remTarget: number;
  legal: number;
  bankruptcy: number;
  claims: number;
  stopped: number;
  notDue: number;
  paymentTerm: string;
  customerClass: string;
  companyType: string;
  insVia: 'unique' | 'code+bu' | 'code-only' | 'none';
  balanceRecords: MonthlyBalanceRecord[];
}

interface UserOption {
  id: string;
  name: string;
}

type Which = 'taskOwner' | 'supervisor';
type CellState = 'saving' | 'saved' | 'error';

const N = (v: number | null | undefined): number => Number(v || 0);
const norm = (s: string | null | undefined): string => (s == null ? '' : String(s)).trim().toUpperCase();
const fmt = (n: number): string => (n == null || n === 0 ? '—' : n.toLocaleString('en-US', { maximumFractionDigits: 0 }));

/** e.g. 442006808 -> "442.0M". Used for the large KPI headline figures. */
function abbreviate(n: number): string {
  const abs = Math.abs(n);
  const sign = n < 0 ? '-' : '';
  if (abs >= 1_000_000_000) return `${sign}${(abs / 1_000_000_000).toFixed(1)}B`;
  if (abs >= 1_000_000) return `${sign}${(abs / 1_000_000).toFixed(1)}M`;
  if (abs >= 1_000) return `${sign}${(abs / 1_000).toFixed(1)}K`;
  return n.toLocaleString('en-US', { maximumFractionDigits: 0 });
}

/** Government/Petroleum read as public-sector (green), Banks as financial (gold), everything else neutral. */
function typeBadgeClass(type: string): string {
  const t = type.trim().toLowerCase();
  if (t.startsWith('government') || t.startsWith('petroleum')) return 'cp-type-green';
  if (t.startsWith('bank')) return 'cp-type-gold';
  return 'cp-type-gray';
}

function hasData(rec: InsuranceRecord): boolean {
  return !!(rec.paymentTerm || rec.customerClass || rec.companyType);
}

function buildInsuranceMaps(insRows: Cfm_insurancecompanies[]): { byUnique: Map<string, InsuranceRecord>; byCode: Map<string, InsuranceRecord> } {
  const byUnique = new Map<string, InsuranceRecord>();
  const byCode = new Map<string, InsuranceRecord>();
  insRows.forEach((ir) => {
    const rec: InsuranceRecord = {
      paymentTerm: formattedValue(ir, '_cfm_paymentterm_value'),
      customerClass: formattedValue(ir, 'cfm_customerclass'),
      companyType: formattedValue(ir, 'cfm_companytype'),
    };
    const bu = norm(ir.cfm_butext);
    const uKey = norm(ir.cfm_unique);
    if (uKey && (!byUnique.has(uKey) || hasData(rec))) byUnique.set(uKey, rec);
    const cKey = norm(ir.cfm_code) + bu;
    if (cKey && (!byUnique.has(cKey) || hasData(rec))) byUnique.set(cKey, rec);
    const codeOnly = norm(ir.cfm_code);
    const existing = byCode.get(codeOnly);
    if (codeOnly && (!existing || (hasData(rec) && !hasData(existing)))) byCode.set(codeOnly, rec);
  });
  return { byUnique, byCode };
}

const MONTH_SHORT = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** Keyed by company code + BU text (the Ref.html join rule for this table: cfm_businessunit's
 * display text + cfm_companycode), one entry per cfm_companymonthlybalanceses row so a code
 * with several monthly records keeps all of them. */
function buildMonthlyBalanceMap(balRows: Cfm_companymonthlybalanceses[]): Map<string, MonthlyBalanceRecord[]> {
  const map = new Map<string, MonthlyBalanceRecord[]>();
  balRows.forEach((b) => {
    const code = norm(b.cfm_companycode);
    const bu = norm(formattedValue(b, '_cfm_businessunit_value'));
    if (!code || !b.cfm_month) return;
    const d = new Date(b.cfm_month);
    if (isNaN(d.getTime())) return;
    const rec: MonthlyBalanceRecord = {
      monthLabel: `${MONTH_SHORT[d.getUTCMonth()]} ${d.getUTCFullYear()}`,
      monthShort: MONTH_SHORT[d.getUTCMonth()],
      sortKey: d.getUTCFullYear() * 12 + d.getUTCMonth(),
      balance: N(b.cfm_balance),
      totalClaims: N(b.cfm_totalclaims),
    };
    const key = `${code}|${bu}`;
    const list = map.get(key);
    if (list) list.push(rec);
    else map.set(key, [rec]);
  });
  return map;
}

function resolveInsurance(
  code: string,
  bu: string,
  unique: string,
  maps: { byUnique: Map<string, InsuranceRecord>; byCode: Map<string, InsuranceRecord> }
): { rec: InsuranceRecord | null; via: PlanRow['insVia'] } {
  const nCode = norm(code);
  const nBu = norm(bu);
  const nUnique = norm(unique);
  if (nUnique && maps.byUnique.has(nUnique)) return { rec: maps.byUnique.get(nUnique)!, via: 'unique' };
  if (nCode && nBu && maps.byUnique.has(nCode + nBu)) return { rec: maps.byUnique.get(nCode + nBu)!, via: 'code+bu' };
  if (nCode && maps.byCode.has(nCode)) return { rec: maps.byCode.get(nCode)!, via: 'code-only' };
  return { rec: null, via: 'none' };
}

function normaliseRow(
  row: Cfm_aragings,
  maps: { byUnique: Map<string, InsuranceRecord>; byCode: Map<string, InsuranceRecord> },
  balanceMap: Map<string, MonthlyBalanceRecord[]>,
  currentYear: number
): PlanRow {
  const collected = N(row.cfm_collectedamount);
  const tax = N(row.cfm_tax);
  const rejections = N(row.cfm_rejections);
  const collectedPlus = row.cfm_collectedtaxrej != null ? N(row.cfm_collectedtaxrej) : collected + tax + rejections;
  const totalDues = N(row.cfm_totaldues);
  const target = N(row.cfm_targetplan);
  const code = row.cfm_companycode || '';
  const buText = formattedValue(row, '_cfm_type_value');
  const unique = row.cfm_unique || '';
  const { rec, via } = resolveInsurance(code, buText, unique, maps);
  const balanceRecords = (balanceMap.get(`${norm(code)}|${norm(buText)}`) || [])
    .filter((b) => Math.floor(b.sortKey / 12) === currentYear)
    .sort((a, b) => a.sortKey - b.sortKey);
  return {
    id: row.cfm_aragingid,
    code: code || '—',
    name: row.cfm_companyname || '—',
    unique,
    buText,
    taskOwnerId: row._cfm_taskowner_value || '',
    taskOwner: formattedValue(row, '_cfm_taskowner_value'),
    supervisorId: row._cfm_supervisor_value || '',
    supervisor: formattedValue(row, '_cfm_supervisor_value'),
    early: N(row.cfm_earlypayment),
    d60: N(row.cfm_31to60),
    d90: N(row.cfm_61to90),
    d120: N(row.cfm_91to120),
    over120: N(row.cfm_150days),
    recon: N(row.cfm_agreedrecon),
    totalDues,
    target,
    collected,
    collectedPlus,
    pending: N(row.cfm_pendingconfirmation),
    remDues: totalDues - collected,
    remTarget: target - collected,
    legal: N(row.cfm_legalissues),
    bankruptcy: N(row.cfm_bankruptcy),
    claims: N(row.cfm_claimsissue),
    stopped: N(row.cfm_stopped),
    notDue: N(row.cfm_notdue),
    paymentTerm: rec ? rec.paymentTerm || '—' : '—',
    customerClass: rec ? rec.customerClass || '—' : '—',
    companyType: rec ? rec.companyType || '—' : '—',
    insVia: via,
    balanceRecords,
  };
}

async function fetchAll<T>(run: (skipToken?: string) => Promise<{ success: boolean; data?: T[] | null; error?: unknown; skipToken?: string }>, fallback: string): Promise<T[]> {
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

const INS_VIA_LABEL: Record<PlanRow['insVia'], string> = {
  unique: 'Matched (code + BU key)',
  'code+bu': 'Matched (code + BU text)',
  'code-only': 'Matched (code only, fallback)',
  none: 'Not matched',
};

interface CollectionPlanPageProps {
  region: Region | null;
}

const KSA_BU_SET = new Set(KSA_BU_CODES);

export default function CollectionPlanPage({ region }: CollectionPlanPageProps) {
  const [rows, setRows] = useState<PlanRow[]>([]);
  const [users, setUsers] = useState<UserOption[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [query, setQuery] = useState('');
  const [buFilter, setBuFilter] = useState('');
  const [paymentFilter, setPaymentFilter] = useState('');
  const [classFilter, setClassFilter] = useState('');
  const [unassignedOnly, setUnassignedOnly] = useState(false);
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(15);

  const [detailId, setDetailId] = useState<string | null>(null);

  const [cellState, setCellState] = useState<Record<string, CellState>>({});
  const cellTimers = useRef<Record<string, ReturnType<typeof setTimeout>>>({});

  const [picker, setPicker] = useState<{ rowId: string; which: Which; rect: DOMRect } | null>(null);
  const [pickerQuery, setPickerQuery] = useState('');
  const pickerPopRef = useRef<HTMLDivElement | null>(null);

  const [toastMsg, setToastMsg] = useState('');
  const [toastOn, setToastOn] = useState(false);
  const toastTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  function showToast(msg: string, durationMs = 3200) {
    setToastMsg(msg);
    setToastOn(true);
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToastOn(false), durationMs);
  }

  async function load() {
    setLoading(true);
    setLoadError(null);
    try {
      const [aragingRows, insuranceRows, userRows, balanceRows] = await Promise.all([
        fetchAll<Cfm_aragings>(
          (skipToken) => Cfm_aragingsService.getAll({ select: ARAGING_SELECT, filter: 'statecode eq 0', maxPageSize: 5000, skipToken }),
          'Could not load AR aging.'
        ),
        fetchAll<Cfm_insurancecompanies>(
          (skipToken) => Cfm_insurancecompaniesService.getAll({ select: INSURANCE_SELECT, filter: 'statecode eq 0', maxPageSize: 5000, skipToken }),
          'Could not load insurance companies.'
        ),
        fetchAll<{ systemuserid: string; fullname?: string }>(
          (skipToken) =>
            SystemusersService.getAll({
              select: ['systemuserid', 'fullname'],
              filter: 'isdisabled eq false and applicationid eq null',
              orderBy: ['fullname'],
              maxPageSize: 5000,
              skipToken,
            }),
          'Could not load users.'
        ),
        fetchAll<Cfm_companymonthlybalanceses>(
          (skipToken) =>
            Cfm_companymonthlybalancesesService.getAll({ select: MONTHLY_BALANCE_SELECT, filter: 'statecode eq 0', maxPageSize: 5000, skipToken }),
          'Could not load company monthly balances.'
        ),
      ]);

      const maps = buildInsuranceMaps(insuranceRows);
      const balanceMap = buildMonthlyBalanceMap(balanceRows);
      const currentYear = new Date().getFullYear();
      const planRows = aragingRows.map((r) => normaliseRow(r, maps, balanceMap, currentYear));
      const userOptions = userRows.filter((u) => u.fullname).map((u) => ({ id: u.systemuserid, name: u.fullname! }));

      setRows(planRows);
      setUsers(userOptions);
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
    if (!picker) return;
    function onDocDown(e: MouseEvent) {
      const target = e.target as Node;
      if (pickerPopRef.current && !pickerPopRef.current.contains(target)) {
        setPicker(null);
      }
    }
    function onScrollOrResize(e: Event) {
      const target = e.target as Node;
      if (pickerPopRef.current && pickerPopRef.current.contains(target)) return;
      setPicker(null);
    }
    document.addEventListener('mousedown', onDocDown, true);
    window.addEventListener('scroll', onScrollOrResize, true);
    window.addEventListener('resize', onScrollOrResize);
    return () => {
      document.removeEventListener('mousedown', onDocDown, true);
      window.removeEventListener('scroll', onScrollOrResize, true);
      window.removeEventListener('resize', onScrollOrResize);
    };
  }, [picker]);

  /** The BU column (cfm_type) partitions rows by region: Saudi Arabia is the fixed
   * KSA_BU_CODES list, Egypt is everything else. This is a hard partition applied
   * before any of the page's own filters, not just another dropdown option. */
  const regionRows = useMemo(() => {
    if (!region) return rows;
    return rows.filter((r) => {
      const inKsa = KSA_BU_SET.has(norm(r.buText));
      return region === 'KSA' ? inKsa : !inKsa;
    });
  }, [rows, region]);

  const matchStats = useMemo(() => {
    const matched = regionRows.filter((r) => r.insVia !== 'none').length;
    return { matched, unmatched: regionRows.length - matched };
  }, [regionRows]);

  /** One trailing table column per distinct month that actually has a matched Company
   * Monthly Balances record this year — not a fixed Jan–Dec grid, so the table only grows
   * as many columns as there is real data for. */
  const monthColumns = useMemo(() => {
    const byKey = new Map<number, { sortKey: number; monthShort: string; monthLabel: string }>();
    regionRows.forEach((r) => r.balanceRecords.forEach((b) => {
      if (!byKey.has(b.sortKey)) byKey.set(b.sortKey, b);
    }));
    return Array.from(byKey.values()).sort((a, b) => a.sortKey - b.sortKey);
  }, [regionRows]);
  const totalCols = 26 + monthColumns.length;

  const buOptions = useMemo(() => uniqueSorted(regionRows, (r) => r.buText), [regionRows]);
  const paymentOptions = useMemo(() => uniqueSorted(regionRows, (r) => r.paymentTerm), [regionRows]);
  const classOptions = useMemo(() => uniqueSorted(regionRows, (r) => r.customerClass), [regionRows]);

  function uniqueSorted(list: PlanRow[], pick: (r: PlanRow) => string): string[] {
    const set = new Set<string>();
    list.forEach((r) => {
      const v = pick(r);
      if (v && v !== '—') set.add(v);
    });
    return Array.from(set).sort();
  }

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return regionRows.filter((r) => {
      if (buFilter && r.buText !== buFilter) return false;
      if (classFilter && r.customerClass !== classFilter) return false;
      if (paymentFilter && r.paymentTerm !== paymentFilter) return false;
      if (unassignedOnly && r.taskOwner && r.supervisor) return false;
      if (q && !`${r.code} ${r.name} ${r.taskOwner} ${r.supervisor}`.toLowerCase().includes(q)) return false;
      return true;
    });
  }, [regionRows, query, buFilter, paymentFilter, classFilter, unassignedOnly]);

  useEffect(() => {
    setPage(1);
  }, [query, buFilter, paymentFilter, classFilter, unassignedOnly]);

  const totals = useMemo(() => {
    const t = {
      early: 0,
      d60: 0,
      d90: 0,
      d120: 0,
      over120: 0,
      recon: 0,
      totalDues: 0,
      target: 0,
      collected: 0,
      collectedPlus: 0,
      pending: 0,
      remDues: 0,
      remTarget: 0,
      notDue: 0,
      monthTotals: {} as Record<number, number>,
    };
    filtered.forEach((r) => {
      t.early += r.early;
      t.d60 += r.d60;
      t.d90 += r.d90;
      t.d120 += r.d120;
      t.over120 += r.over120;
      t.recon += r.recon;
      t.totalDues += r.totalDues;
      t.target += r.target;
      t.collected += r.collected;
      t.collectedPlus += r.collectedPlus;
      t.pending += r.pending;
      t.remDues += r.remDues;
      t.remTarget += r.remTarget;
      t.notDue += r.notDue;
      r.balanceRecords.forEach((b) => {
        t.monthTotals[b.sortKey] = (t.monthTotals[b.sortKey] || 0) + b.balance;
      });
    });
    return t;
  }, [filtered]);

  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
  const clampedPage = Math.min(page, totalPages);
  const pageRows = filtered.slice((clampedPage - 1) * pageSize, (clampedPage - 1) * pageSize + pageSize);

  const detailRow = useMemo(() => rows.find((r) => r.id === detailId) || null, [rows, detailId]);

  function openDetail(id: string) {
    setDetailId(id);
  }

  function closeDetail() {
    setDetailId(null);
  }

  function clearFilters() {
    setQuery('');
    setBuFilter('');
    setPaymentFilter('');
    setClassFilter('');
    setUnassignedOnly(false);
  }

  function pagerButtons(): (number | 'ellipsis')[] {
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

  async function handleAssign(row: PlanRow, which: Which, userId: string) {
    const prevId = which === 'taskOwner' ? row.taskOwnerId : row.supervisorId;
    if ((userId || '') === (prevId || '')) return;
    const prevName = which === 'taskOwner' ? row.taskOwner : row.supervisor;
    const user = userId ? users.find((u) => u.id === userId) : undefined;
    const key = `${row.id}:${which}`;

    setRows((prev) =>
      prev.map((r) =>
        r.id === row.id
          ? which === 'taskOwner'
            ? { ...r, taskOwnerId: userId, taskOwner: user ? user.name : '' }
            : { ...r, supervisorId: userId, supervisor: user ? user.name : '' }
          : r
      )
    );
    flashCell(key, 'saving');
    try {
      const bindKey = which === 'taskOwner' ? 'cfm_TaskOwner@odata.bind' : 'cfm_Supervisor@odata.bind';
      const body: Record<string, unknown> = { [bindKey]: userId ? `/systemusers(${userId})` : null };
      const result = await Cfm_aragingsService.update(row.id, body as Partial<Omit<Cfm_aragingsBase, 'cfm_aragingid'>>);
      if (!result.success) {
        throw new Error(operationErrorMessage(result.error, 'Save failed.'));
      }
      flashCell(key, 'saved', 900);
      showToast(`${which === 'taskOwner' ? 'Task owner' : 'Supervisor'} updated${user ? ` → ${user.name}` : ' (cleared)'}.`);
    } catch (e) {
      setRows((prev) =>
        prev.map((r) =>
          r.id === row.id
            ? which === 'taskOwner'
              ? { ...r, taskOwnerId: prevId, taskOwner: prevName }
              : { ...r, supervisorId: prevId, supervisor: prevName }
            : r
        )
      );
      flashCell(key, 'error', 1400);
      showToast(e instanceof Error ? e.message : 'Save failed.', 4500);
    }
  }

  function exportCsv() {
    if (!filtered.length) {
      showToast('Nothing to export.');
      return;
    }
    const header = [
      'BU',
      'Supervisor',
      'Task Owner',
      'Payment Term',
      'Customer Class',
      'Company Type',
      'Code',
      'Company Name',
      'Early Payment',
      '31-60',
      '60-90',
      '90-120',
      'Over 120',
      'Agreed Reconciliation',
      'Total Dues',
      'Target',
      'Total-collected',
      'Total collected+tax+rej',
      'Pending Confirmation',
      'Remaining from Dues',
      'Remaining from Target',
      'Legal Issues',
      'Bankruptcy',
      'Claims Issue',
      'Stopped',
      'Not Due',
      ...monthColumns.map((c) => `Balance ${c.monthLabel}`),
    ];
    const lines = [header.join(',')].concat(
      filtered.map((r) =>
        [
          r.buText,
          r.supervisor,
          r.taskOwner,
          r.paymentTerm,
          r.customerClass,
          r.companyType,
          r.code,
          r.name,
          r.early,
          r.d60,
          r.d90,
          r.d120,
          r.over120,
          r.recon,
          r.totalDues,
          r.target,
          r.collected,
          r.collectedPlus,
          r.pending,
          r.remDues,
          r.remTarget,
          r.legal,
          r.bankruptcy,
          r.claims,
          r.stopped,
          r.notDue,
          ...monthColumns.map((c) => r.balanceRecords.find((b) => b.sortKey === c.sortKey)?.balance ?? ''),
        ]
          .map(escapeForCsv)
          .join(',')
      )
    );
    downloadCsv(`collection_plan_${new Date().toISOString().slice(0, 10)}.csv`, lines);
    showToast(`Exported ${filtered.length} rows to CSV.`);
  }

  function closePicker() {
    setPicker(null);
  }

  function pickUser(row: PlanRow, which: Which, userId: string) {
    closePicker();
    handleAssign(row, which, userId);
  }

  /** Editable Task Owner / Supervisor trigger — opens the searchable picker popover (renderPickerPopover). */
  function assignButton(row: PlanRow, which: Which) {
    const currentName = which === 'taskOwner' ? row.taskOwner : row.supervisor;
    const key = `${row.id}:${which}`;
    const state = cellState[key];
    const isOpen = picker?.rowId === row.id && picker?.which === which;
    return (
      <div className={`cp-edit-wrap${state ? ` cp-cell-${state}` : ''}`}>
        <button
          type="button"
          className={`cp-user-pick${isOpen ? ' open' : ''}`}
          disabled={state === 'saving'}
          onClick={(e) => {
            e.stopPropagation();
            if (isOpen) {
              closePicker();
              return;
            }
            setPickerQuery('');
            setPicker({ rowId: row.id, which, rect: e.currentTarget.getBoundingClientRect() });
          }}
        >
          <span className="cp-user-pick-name">{currentName || '— Assign —'}</span>
          <Icon className="fa-solid fa-chevron-down cp-user-pick-caret" />
        </button>
      </div>
    );
  }

  function renderPickerPopover() {
    if (!picker) return null;
    const row = rows.find((r) => r.id === picker.rowId);
    if (!row) return null;
    const currentId = picker.which === 'taskOwner' ? row.taskOwnerId : row.supervisorId;
    const q = pickerQuery.trim().toLowerCase();
    const items = q ? users.filter((u) => u.name.toLowerCase().includes(q)) : users;
    const width = Math.max(picker.rect.width, 240);
    const left = Math.min(picker.rect.left, Math.max(8, window.innerWidth - width - 8));
    const estHeight = 320;
    let top = picker.rect.bottom + 4;
    if (top + estHeight > window.innerHeight) {
      top = Math.max(8, picker.rect.top - estHeight - 4);
    }
    return createPortal(
      <div className="cp-picker" style={{ top, left, width }} ref={pickerPopRef}>
        <div className="cp-picker-search">
          <Icon className="fa-solid fa-magnifying-glass" />
          <input autoFocus placeholder="Search users…" value={pickerQuery} onChange={(e) => setPickerQuery(e.target.value)} />
        </div>
        <div className="cp-picker-list">
          <button
            type="button"
            className={`cp-picker-opt cp-picker-clear${!currentId ? ' sel' : ''}`}
            onClick={() => pickUser(row, picker.which, '')}
          >
            <Icon className="fa-solid fa-user-slash" /> Unassigned
          </button>
          {items.map((u) => (
            <button
              key={u.id}
              type="button"
              className={`cp-picker-opt${u.id === currentId ? ' sel' : ''}`}
              onClick={() => pickUser(row, picker.which, u.id)}
            >
              <span className="cp-picker-opt-name">{u.name}</span>
              {u.id === currentId && <Icon className="fa-solid fa-check" />}
            </button>
          ))}
          {!items.length && <div className="cp-picker-empty">No users match.</div>}
        </div>
      </div>,
      document.body
    );
  }

  function stat(label: string, value: string, tone?: 'pos' | 'neg') {
    return (
      <div className="cp-stat">
        <div className="cp-stat-label">{label}</div>
        <div className={`cp-stat-value${tone ? ` cp-stat-${tone}` : ''}`}>{value}</div>
      </div>
    );
  }

  function flagStat(label: string, v: number) {
    return (
      <div className="cp-stat">
        <div className="cp-stat-label">{label}</div>
        <div className="cp-stat-value">{v > 0 ? <span className="cp-flag-badge">{fmt(v)}</span> : <span className="cp-dim">None</span>}</div>
      </div>
    );
  }

  function classBadgeClass(cls: string): string {
    const c = cls.trim().toUpperCase();
    if (c.startsWith('A')) return 'acc-cls-A';
    if (c.startsWith('B')) return 'acc-cls-B';
    if (c.startsWith('C')) return 'acc-cls-C';
    return 'acc-cls-none';
  }

  function classCell(cls: string) {
    return cls && cls !== '—' ? (
      <td>
        <span className={`acc-cls cp-cls-badge ${classBadgeClass(cls)}`}>{cls}</span>
      </td>
    ) : (
      <td>
        <span className="cp-chip-miss">— not matched —</span>
      </td>
    );
  }

  function monthCell(r: PlanRow, sortKey: number) {
    const rec = r.balanceRecords.find((b) => b.sortKey === sortKey);
    return (
      <td key={sortKey} className="cp-col-month cp-numeric">
        {rec ? fmt(rec.balance) : <span className="cp-dim">—</span>}
      </td>
    );
  }

  function textCell(value: string) {
    return value && value !== '—' ? (
      <td>{value}</td>
    ) : (
      <td>
        <span className="cp-chip-miss">— not matched —</span>
      </td>
    );
  }

  function flagCell(v: number) {
    return v > 0 ? (
      <td>
        <span className="acc-cover">
          <span className="acc-dot acc-bad"></span>
          {fmt(v)}
        </span>
      </td>
    ) : (
      <td className="cp-flag-none">—</td>
    );
  }

  return (
    <>
      <div className="acc-wrap cp-page">
        <div className="acc-head">
          <div>
            <h1 className="acc-title">Collection Plan</h1>
            <div className="acc-sub">
              {loading
                ? 'Loading…'
                : `${regionRows.length} compan${regionRows.length === 1 ? 'y' : 'ies'}${
                    matchStats.unmatched > 0 ? ` · ${matchStats.unmatched} unmatched to insurance` : ''
                  }`}
            </div>
          </div>
          <div className="acc-actions">
            <button className="acc-btn" onClick={exportCsv}>
              <Icon className="fa-solid fa-file-arrow-down" /> <span className="cp-btn-label">Export</span>
            </button>
            <button className="acc-btn" onClick={load}>
              <Icon className="fa-solid fa-rotate" /> <span className="cp-btn-label">Reload</span>
            </button>
          </div>
        </div>

        <div className="kpi-grid">
          <div className="kpi-card">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-gold">
                <Icon className="fa-solid fa-coins" />
              </span>
              <span className="kpi-label">Total dues</span>
            </div>
            <div className="kpi-value">{loading ? '—' : abbreviate(totals.totalDues)}</div>
            <div className="kpi-context">{loading ? '' : totals.totalDues.toLocaleString('en-US')}</div>
          </div>

          <div className="kpi-card">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-gold">
                <Icon className="fa-solid fa-bullseye" />
              </span>
              <span className="kpi-label">Target plan</span>
            </div>
            <div className="kpi-value">{loading ? '—' : totals.target > 0 ? totals.target.toLocaleString('en-US') : '—'}</div>
            <div className="kpi-context">{loading ? '' : totals.target > 0 ? ' ' : 'Not set'}</div>
          </div>

          <div className="kpi-card kpi-ok">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-ok">
                <Icon className="fa-solid fa-sack-dollar" />
              </span>
              <span className="kpi-label">Total collected</span>
            </div>
            <div className="kpi-value">{loading ? '—' : abbreviate(totals.collected)}</div>
            <div className="kpi-context">
              {loading ? '' : totals.totalDues > 0 ? `${Math.round((totals.collected / totals.totalDues) * 100)}% of dues` : '—'}
            </div>
          </div>

          <div className="kpi-card kpi-gold">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-gold">
                <Icon className="fa-solid fa-hourglass-half" />
              </span>
              <span className="kpi-label">Remaining dues</span>
            </div>
            <div className="kpi-value">{loading ? '—' : abbreviate(totals.remDues)}</div>
            <div className="kpi-context">{loading ? '' : totals.remDues.toLocaleString('en-US')}</div>
          </div>

          <div className="kpi-card kpi-warn">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-warn">
                <Icon className="fa-solid fa-clock" />
              </span>
              <span className="kpi-label">Pending confirmation</span>
            </div>
            <div className="kpi-value">{loading ? '—' : abbreviate(totals.pending)}</div>
            <div className="kpi-context">{loading ? '' : totals.pending.toLocaleString('en-US')}</div>
          </div>

          <div className="kpi-card kpi-info">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-info">
                <Icon className="fa-solid fa-money-bill-trend-up" />
              </span>
              <span className="kpi-label">Pending confirmation + collected</span>
            </div>
            <div className="kpi-value">{loading ? '—' : abbreviate(totals.pending + totals.collected)}</div>
            <div className="kpi-context">{loading ? '' : (totals.pending + totals.collected).toLocaleString('en-US')}</div>
          </div>
        </div>

        <div className="tbl-shell">
          <div className="tbl-hdr">
            <div className="cp-filter-bar">
              <div className="acc-search cp-search">
                <Icon className="fa-solid fa-magnifying-glass" />
                <input placeholder="Search company, code, owner…" value={query} onChange={(e) => setQuery(e.target.value)} />
              </div>
              <select className="acc-f" value={buFilter} onChange={(e) => setBuFilter(e.target.value)}>
                <option value="">All BU</option>
                {buOptions.map((v) => (
                  <option key={v} value={v}>
                    {v}
                  </option>
                ))}
              </select>
              <select className="acc-f" value={paymentFilter} onChange={(e) => setPaymentFilter(e.target.value)}>
                <option value="">All Payments</option>
                {paymentOptions.map((v) => (
                  <option key={v} value={v}>
                    {v}
                  </option>
                ))}
              </select>
              <select className="acc-f" value={classFilter} onChange={(e) => setClassFilter(e.target.value)}>
                <option value="">All Classes</option>
                {classOptions.map((v) => (
                  <option key={v} value={v}>
                    {v}
                  </option>
                ))}
              </select>
              <button className={`cp-toggle${unassignedOnly ? ' active' : ''}`} onClick={() => setUnassignedOnly((v) => !v)}>
                <Icon className="fa-solid fa-user-slash" /> Unassigned only
              </button>
              <button className="cp-clear" onClick={clearFilters}>
                <Icon className="fa-solid fa-xmark" /> Clear
              </button>
            </div>
            <div className="cp-count">
              <Icon className="fa-solid fa-table" /> Showing <strong>{filtered.length}</strong> of <strong>{regionRows.length}</strong>
            </div>
          </div>

          <div className="cp-tbl-scroll">
            <table className="cp-tbl">
              <thead>
                <tr className="cp-head-row">
                  <th>BU</th>
                  <th>Supervisor</th>
                  <th className="cp-group-end">Task owner</th>
                  <th>Payment term</th>
                  <th>Customer class</th>
                  <th className="cp-group-end">Company type</th>
                  <th className="cp-numeric cp-group-end">Code</th>
                  <th className="cp-stick-col cp-group-end">Company name</th>
                  <th className="cp-numeric">Early payment</th>
                  <th className="cp-numeric">31–60</th>
                  <th className="cp-numeric">60–90</th>
                  <th className="cp-numeric">90–120</th>
                  <th className="cp-numeric cp-group-end">Over 120</th>
                  <th className="cp-numeric">Agreed reconciliation</th>
                  <th className="cp-numeric cp-group-end">Total dues</th>
                  <th className="cp-numeric">Target</th>
                  <th className="cp-numeric">Total collected</th>
                  <th className="cp-numeric">Collected + tax + rej.</th>
                  <th className="cp-numeric">Pending confirmation</th>
                  <th className="cp-numeric">Remaining from dues</th>
                  <th className="cp-numeric cp-group-end">Remaining from target</th>
                  <th>Legal issues</th>
                  <th>Bankruptcy</th>
                  <th>Claims issue</th>
                  <th>Stopped</th>
                  <th className={`cp-numeric${monthColumns.length ? ' cp-group-end' : ''}`}>Not due</th>
                  {monthColumns.map((c) => (
                    <th key={c.sortKey} className="cp-numeric cp-col-month" title={c.monthLabel}>
                      {c.monthShort}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {loading ? (
                  Array.from({ length: 5 }).map((_, i) => (
                    <tr key={i} className="cp-skel-row">
                      <td colSpan={totalCols}>
                        <div className="acc-skel" style={{ width: `${55 + ((i * 13) % 35)}%` }}></div>
                      </td>
                    </tr>
                  ))
                ) : loadError ? (
                  <tr>
                    <td colSpan={totalCols}>
                      <div className="acc-state">
                        <Icon className="fa-solid fa-triangle-exclamation" />
                        Couldn't load Collection Plan.
                        <br />
                        <small>{loadError}</small>
                      </div>
                    </td>
                  </tr>
                ) : pageRows.length === 0 ? (
                  <tr>
                    <td colSpan={totalCols}>
                      <div className="acc-state">
                        <Icon className="fa-solid fa-filter-circle-xmark" />
                        No records match the current filters.
                      </div>
                    </td>
                  </tr>
                ) : (
                  pageRows.map((r) => (
                    <tr key={r.id} className="cp-row" onClick={() => openDetail(r.id)}>
                      <td>{r.buText || <span className="cp-dim">—</span>}</td>
                      <td onClick={(e) => e.stopPropagation()}>{assignButton(r, 'supervisor')}</td>
                      <td className="cp-group-end" onClick={(e) => e.stopPropagation()}>
                        {assignButton(r, 'taskOwner')}
                      </td>
                      {textCell(r.paymentTerm)}
                      {classCell(r.customerClass)}
                      <td className="cp-group-end">
                        {r.companyType && r.companyType !== '—' ? (
                          <span className={`cp-type-badge ${typeBadgeClass(r.companyType)}`}>{r.companyType.trim()}</span>
                        ) : (
                          <span className="cp-chip-miss">— not matched —</span>
                        )}
                      </td>
                      <td className="cp-col-id cp-numeric cp-group-end">{r.code}</td>
                      <td className="cp-col-bold cp-stick-col cp-group-end" dir="auto">
                        {r.name}
                      </td>
                      <td className="cp-col-amt cp-numeric">{fmt(r.early)}</td>
                      <td className="cp-col-amt cp-numeric">{fmt(r.d60)}</td>
                      <td className="cp-col-amt cp-numeric">{fmt(r.d90)}</td>
                      <td className="cp-col-amt cp-numeric">{fmt(r.d120)}</td>
                      <td className="cp-col-amt cp-numeric cp-group-end">{fmt(r.over120)}</td>
                      <td className="cp-col-amt cp-numeric">{fmt(r.recon)}</td>
                      <td className="cp-col-amt cp-numeric cp-group-end">{fmt(r.totalDues)}</td>
                      <td className="cp-col-pos cp-numeric">{fmt(r.target)}</td>
                      <td className="cp-col-amt cp-numeric">{fmt(r.collected)}</td>
                      <td className="cp-col-amt cp-numeric">{fmt(r.collectedPlus)}</td>
                      <td className="cp-col-amt cp-numeric">{fmt(r.pending)}</td>
                      <td className={`cp-numeric ${r.remDues < 0 ? 'cp-col-pos' : 'cp-col-neg'}`}>{fmt(r.remDues)}</td>
                      <td className={`cp-numeric cp-group-end ${r.remTarget < 0 ? 'cp-col-pos' : 'cp-col-neg'}`}>{fmt(r.remTarget)}</td>
                      {flagCell(r.legal)}
                      {flagCell(r.bankruptcy)}
                      {flagCell(r.claims)}
                      {flagCell(r.stopped)}
                      <td className={`cp-col-amt cp-numeric${monthColumns.length ? ' cp-group-end' : ''}`}>{fmt(r.notDue)}</td>
                      {monthColumns.map((c) => monthCell(r, c.sortKey))}
                    </tr>
                  ))
                )}
              </tbody>
              {!loading && !loadError && filtered.length > 0 && (
                <tfoot>
                  <tr>
                    <td colSpan={3} className="cp-col-bold cp-group-end">
                      <Icon className="fa-solid fa-square-poll-vertical" style={{ color: 'var(--gold)', marginRight: 6 }} />
                      Total ({filtered.length})
                    </td>
                    <td colSpan={3} className="cp-group-end"></td>
                    <td className="cp-numeric cp-group-end"></td>
                    <td className="cp-stick-col cp-group-end"></td>
                    <td className="cp-col-amt cp-numeric">{totals.early.toLocaleString('en-US')}</td>
                    <td className="cp-col-amt cp-numeric">{totals.d60.toLocaleString('en-US')}</td>
                    <td className="cp-col-amt cp-numeric">{totals.d90.toLocaleString('en-US')}</td>
                    <td className="cp-col-amt cp-numeric">{totals.d120.toLocaleString('en-US')}</td>
                    <td className="cp-col-amt cp-numeric cp-group-end">{totals.over120.toLocaleString('en-US')}</td>
                    <td className="cp-col-amt cp-numeric">{totals.recon.toLocaleString('en-US')}</td>
                    <td className="cp-col-amt cp-numeric cp-group-end">{totals.totalDues.toLocaleString('en-US')}</td>
                    <td className="cp-col-pos cp-numeric">{totals.target.toLocaleString('en-US')}</td>
                    <td className="cp-col-amt cp-numeric">{totals.collected.toLocaleString('en-US')}</td>
                    <td className="cp-col-amt cp-numeric">{totals.collectedPlus.toLocaleString('en-US')}</td>
                    <td className="cp-col-amt cp-numeric">{totals.pending.toLocaleString('en-US')}</td>
                    <td className="cp-col-neg cp-numeric">{totals.remDues.toLocaleString('en-US')}</td>
                    <td className="cp-col-neg cp-numeric cp-group-end">{totals.remTarget.toLocaleString('en-US')}</td>
                    <td colSpan={4}></td>
                    <td className={`cp-col-amt cp-numeric${monthColumns.length ? ' cp-group-end' : ''}`}>
                      {totals.notDue.toLocaleString('en-US')}
                    </td>
                    {monthColumns.map((c) => (
                      <td key={c.sortKey} className="cp-col-month cp-numeric">
                        {(totals.monthTotals[c.sortKey] || 0).toLocaleString('en-US')}
                      </td>
                    ))}
                  </tr>
                </tfoot>
              )}
            </table>
          </div>

          {/* Mobile-only card list (CSS-toggled with .cp-tbl-scroll — see CollectionPlan.css).
              The table is ~26 columns; rather than cram every field into a phone card, each
              card shows the identity/priority fields and tapping it opens the same detail
              drawer the desktop table row already opens (openDetail), which already lists
              every field. */}
          <div className="cp-cards">
            {loading ? (
              Array.from({ length: 5 }).map((_, i) => (
                <div className="cp-card cp-card-skel" key={i}>
                  <div className="acc-skel" style={{ width: `${55 + ((i * 13) % 35)}%` }}></div>
                </div>
              ))
            ) : loadError ? (
              <div className="acc-state">
                <Icon className="fa-solid fa-triangle-exclamation" />
                Couldn't load Collection Plan.
                <br />
                <small>{loadError}</small>
              </div>
            ) : pageRows.length === 0 ? (
              <div className="acc-state">
                <Icon className="fa-solid fa-filter-circle-xmark" />
                No records match the current filters.
              </div>
            ) : (
              pageRows.map((r) => (
                <div key={r.id} className="cp-card" onClick={() => openDetail(r.id)}>
                  <div className="cp-card-hdr">
                    <div className="cp-card-title" dir="auto">
                      {r.name}
                    </div>
                    {r.customerClass && r.customerClass !== '—' ? (
                      <span className={`acc-cls cp-cls-badge ${classBadgeClass(r.customerClass)}`}>{r.customerClass}</span>
                    ) : (
                      <span className="cp-chip-miss">—</span>
                    )}
                  </div>
                  <div className="cp-card-code">{r.code}</div>
                  <div className="cp-card-grid">
                    <div className="cp-card-field">
                      <span className="cp-card-lbl">BU</span>
                      <span className="cp-card-val">{r.buText || '—'}</span>
                    </div>
                    <div className="cp-card-field">
                      <span className="cp-card-lbl">Supervisor</span>
                      <span className="cp-card-val">{r.supervisor || '—'}</span>
                    </div>
                    <div className="cp-card-field">
                      <span className="cp-card-lbl">Task owner</span>
                      <span className="cp-card-val">{r.taskOwner || '—'}</span>
                    </div>
                    <div className="cp-card-field">
                      <span className="cp-card-lbl">Payment term</span>
                      <span className="cp-card-val">{r.paymentTerm || '—'}</span>
                    </div>
                    <div className="cp-card-field">
                      <span className="cp-card-lbl">Company type</span>
                      <span className="cp-card-val">{r.companyType || '—'}</span>
                    </div>
                    <div className="cp-card-field">
                      <span className="cp-card-lbl">Monthly balance</span>
                      <span className="cp-card-val">
                        {r.balanceRecords.length ? r.balanceRecords.map((b) => `${b.monthShort} ${fmt(b.balance)}`).join(' · ') : '—'}
                      </span>
                    </div>
                    <div className="cp-card-field">
                      <span className="cp-card-lbl">Early payment</span>
                      <span className="cp-card-val">{fmt(r.early)}</span>
                    </div>
                    <div className="cp-card-field">
                      <span className="cp-card-lbl">Total dues</span>
                      <span className="cp-card-val">{fmt(r.totalDues)}</span>
                    </div>
                    <div className="cp-card-field">
                      <span className="cp-card-lbl">Target</span>
                      <span className="cp-card-val">{fmt(r.target)}</span>
                    </div>
                    <div className="cp-card-field">
                      <span className="cp-card-lbl">Total collected</span>
                      <span className="cp-card-val">{fmt(r.collected)}</span>
                    </div>
                  </div>
                </div>
              ))
            )}
          </div>

          {!loading && !loadError && filtered.length > 0 && (
            <div className="cp-pagination">
              <div className="cp-page-size">
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
              <div className="cp-page-info">
                Showing <strong>{(clampedPage - 1) * pageSize + 1}</strong>–
                <strong>{Math.min(clampedPage * pageSize, filtered.length)}</strong> of{' '}
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

      <div className={`acc-scrim${detailRow ? ' on' : ''}`} onClick={closeDetail}></div>
      <aside className={`acc-drawer cp-detail${detailRow ? ' on' : ''}`}>
        {detailRow && (
          <>
            <div className="acc-dhead">
              <h2>
                <span className="cp-detail-name" dir="auto">
                  {detailRow.name}
                </span>
              </h2>
              <button className="acc-dclose" onClick={closeDetail}>
                <Icon className="fa-solid fa-xmark" />
              </button>
            </div>
            <div className="acc-dbody">
              <div className="cp-detail-sub">
                <span className="cp-col-id">{detailRow.code}</span>
                {detailRow.buText && <span> · {detailRow.buText}</span>}
              </div>

              <div className="cp-detail-section">
                <div className="cp-detail-section-title">Assignment</div>
                <div className="cp-detail-grid cp-detail-grid-2">
                  <div className="cp-field">
                    <label>Supervisor</label>
                    {assignButton(detailRow, 'supervisor')}
                  </div>
                  <div className="cp-field">
                    <label>Task owner</label>
                    {assignButton(detailRow, 'taskOwner')}
                  </div>
                </div>
              </div>

              <div className="cp-detail-section">
                <div className="cp-detail-section-title">Insurance lookup</div>
                <div className="cp-detail-grid cp-detail-grid-2">
                  {stat('Payment term', detailRow.paymentTerm)}
                  {stat('Customer class', detailRow.customerClass)}
                  {stat('Company type', detailRow.companyType)}
                  {stat('Match status', INS_VIA_LABEL[detailRow.insVia])}
                </div>
              </div>

              <div className="cp-detail-section">
                <div className="cp-detail-section-title">Company monthly balance ({new Date().getFullYear()})</div>
                {detailRow.balanceRecords.length ? (
                  <div className="cp-detail-grid cp-detail-grid-3">
                    {detailRow.balanceRecords.map((b) => (
                      <div className="cp-stat" key={b.monthLabel}>
                        <div className="cp-stat-label">{b.monthLabel}</div>
                        <div className="cp-stat-value">{fmt(b.balance)}</div>
                        <div className="cp-detail-sub">Total claims: {fmt(b.totalClaims)}</div>
                      </div>
                    ))}
                  </div>
                ) : (
                  <div className="cp-detail-sub">No matching Company Monthly Balances record (BU + Code) for this year.</div>
                )}
              </div>

              <div className="cp-detail-section">
                <div className="cp-detail-section-title">Aging breakdown</div>
                <div className="cp-detail-grid cp-detail-grid-3">
                  {stat('Early payment', fmt(detailRow.early))}
                  {stat('31–60', fmt(detailRow.d60))}
                  {stat('60–90', fmt(detailRow.d90))}
                  {stat('90–120', fmt(detailRow.d120))}
                  {stat('Over 120', fmt(detailRow.over120))}
                  {stat('Agreed reconciliation', fmt(detailRow.recon))}
                </div>
              </div>

              <div className="cp-detail-section">
                <div className="cp-detail-section-title">Financials &amp; collection</div>
                <div className="cp-detail-grid cp-detail-grid-2">
                  {stat('Total dues', fmt(detailRow.totalDues))}
                  {stat('Target', fmt(detailRow.target))}
                  {stat('Total collected', fmt(detailRow.collected))}
                  {stat('Total collected + tax + rej.', fmt(detailRow.collectedPlus))}
                  {stat('Pending confirmation', fmt(detailRow.pending))}
                  {stat('Remaining from dues', fmt(detailRow.remDues), detailRow.remDues < 0 ? 'pos' : 'neg')}
                  {stat('Remaining from target', fmt(detailRow.remTarget), detailRow.remTarget < 0 ? 'pos' : 'neg')}
                </div>
              </div>

              <div className="cp-detail-section">
                <div className="cp-detail-section-title">Risk flags</div>
                <div className="cp-detail-grid cp-detail-grid-3">
                  {flagStat('Legal issues', detailRow.legal)}
                  {flagStat('Bankruptcy', detailRow.bankruptcy)}
                  {flagStat('Claims issue', detailRow.claims)}
                  {flagStat('Stopped', detailRow.stopped)}
                  {stat('Not due', fmt(detailRow.notDue))}
                </div>
              </div>
            </div>
          </>
        )}
      </aside>

      <div className={`acc-toast${toastOn ? ' on' : ''}`}>{toastMsg}</div>
      {renderPickerPopover()}
    </>
  );
}
