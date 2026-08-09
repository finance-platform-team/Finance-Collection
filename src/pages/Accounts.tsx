import { useEffect, useMemo, useRef, useState } from 'react';
import { getClient } from '@microsoft/power-apps/data';
import { dataSourcesInfo } from '../../.power/schemas/appschemas/dataSourcesInfo';
import { Cr301_accountsService } from '../generated/services/Cr301_accountsService';
import type { Cr301_accountsBase, Cr301_accountscr301_accounttype } from '../generated/models/Cr301_accountsModel';
import { operationErrorMessage } from '../lib/errors';
import './Accounts.css';

const TYPE_INSURANCE: Cr301_accountscr301_accounttype = 983080002;

/**
 * Classification is the _cr301_accountclassification_value LOOKUP (→ cr301_category,
 * logical name; entity set "cr301_categories"). cr301_category was never added as a
 * connected data source for this Code App (no generated model/service exists for it),
 * so it's queried through the generic Dataverse client. If that table isn't reachable
 * at runtime, the picker degrades to a "load failed" option.
 */
const CLASS_LOOKUP_TABLE = 'cr301_categories';
const CLASS_VALUE_FIELD = '_cr301_accountclassification_value';
const CLASS_NAV = 'cr301_AccountClassification';
const ACCOUNTS_TABLE = 'cr301_accounts';

interface RawAccountType {
  cr301_categoryid: string;
  cr301_title?: string;
}

interface ClassOption {
  id: string;
  label: string;
}

interface AccountRow {
  id: string;
  name: string;
  code: string;
  clsId: string;
  cls: string;
  typeName: string;
  target: number;
  actual: number;
}

type CoverageFilter = '' | 'met' | 'partial' | 'none';

/**
 * The generated model advertises "cr301_accounttypename" / "cr301_accountclassificationname"
 * as if selectable, but Dataverse rejects them. The display text lives on the base
 * attribute's own `@OData.Community.Display.V1.FormattedValue` annotation — read it directly.
 */
function formattedValue(record: unknown, attribute: string): string {
  const raw = (record as Record<string, unknown>)[`${attribute}@OData.Community.Display.V1.FormattedValue`];
  return typeof raw === 'string' ? raw : '';
}

function coverageDot(actual: number, target: number): string {
  if (target > 0 && actual >= target) return 'acc-ok';
  if (actual > 0) return 'acc-warn';
  return 'acc-bad';
}

function classBadgeClass(cls: string): string {
  if (cls === 'A') return 'acc-cls-A';
  if (cls === 'B') return 'acc-cls-B';
  if (cls === 'C') return 'acc-cls-C';
  return 'acc-cls-none';
}

function escapeForCsv(value: string | number): string {
  return `"${String(value).replace(/"/g, '""')}"`;
}

export default function AccountsPage() {
  const [accounts, setAccounts] = useState<AccountRow[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [classes, setClasses] = useState<ClassOption[]>([]);
  const [classesLoadFailed, setClassesLoadFailed] = useState(false);

  const [query, setQuery] = useState('');
  const [clsFilter, setClsFilter] = useState('');
  const [coverFilter, setCoverFilter] = useState<CoverageFilter>('');
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(15);

  const [drawerOpen, setDrawerOpen] = useState(false);
  const [editId, setEditId] = useState<string | null>(null);
  const [formName, setFormName] = useState('');
  const [formCode, setFormCode] = useState('');
  const [formClsId, setFormClsId] = useState('');
  const [formTarget, setFormTarget] = useState('2');
  const [formError, setFormError] = useState('');
  const [saving, setSaving] = useState(false);

  const [selectedIds, setSelectedIds] = useState<Set<string>>(new Set());
  const [bulkDrawerOpen, setBulkDrawerOpen] = useState(false);
  const [bulkTarget, setBulkTarget] = useState('');
  const [bulkError, setBulkError] = useState('');
  const [bulkSaving, setBulkSaving] = useState(false);

  const [toastMsg, setToastMsg] = useState('');
  const [toastOn, setToastOn] = useState(false);
  const toastTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  function showToast(msg: string, durationMs = 2200) {
    setToastMsg(msg);
    setToastOn(true);
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToastOn(false), durationMs);
  }

  async function loadClasses() {
    try {
      const client = getClient(dataSourcesInfo);
      const result = await client.retrieveMultipleRecordsAsync<RawAccountType>(CLASS_LOOKUP_TABLE, {
        select: ['cr301_categoryid', 'cr301_title'],
      });
      if (!result.success) {
        throw new Error(operationErrorMessage(result.error, 'Could not load classifications.'));
      }
      const rows: ClassOption[] = (result.data || [])
        .map((c) => ({ id: c.cr301_categoryid, label: (c.cr301_title || '').trim() }))
        .filter((c) => c.id)
        .sort((a, b) => a.label.localeCompare(b.label));
      setClasses(rows);
      setClassesLoadFailed(false);
    } catch {
      setClasses([]);
      setClassesLoadFailed(true);
    }
  }

  async function load() {
    setLoading(true);
    setLoadError(null);
    try {
      const result = await Cr301_accountsService.getAll({
        select: [
          'cr301_accountid',
          'cr301_acccountname',
          'cr301_accountcode',
          'cr301_accounttype',
          'cfm_visittargetcollection',
          'cr301_monthactualvisits',
          '_cr301_accountclassification_value',
        ],
        filter: `cr301_accounttype eq ${TYPE_INSURANCE}`,
        orderBy: ['cr301_acccountname'],
      });
      if (!result.success) {
        throw new Error(operationErrorMessage(result.error, 'Could not load accounts.'));
      }
      const rows: AccountRow[] = (result.data || []).map((r) => ({
        id: r.cr301_accountid,
        name: r.cr301_acccountname || '—',
        code: r.cr301_accountcode || '—',
        clsId: r._cr301_accountclassification_value || '',
        cls: formattedValue(r, '_cr301_accountclassification_value').trim().toUpperCase().charAt(0),
        typeName: (formattedValue(r, 'cr301_accounttype') || 'Insurance').trim(),
        target: r.cfm_visittargetcollection ?? 0,
        actual: r.cr301_monthactualvisits ?? 0,
      }));
      setAccounts(rows);
    } catch (e) {
      setLoadError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    // Independent loads: a classification-load failure never blocks the grid.
    loadClasses();
    load();
  }, []);

  const view = useMemo(() => {
    const q = query.trim().toLowerCase();
    return accounts.filter((a) => {
      if (q && !`${a.name} ${a.code}`.toLowerCase().includes(q)) return false;
      if (clsFilter && a.cls !== clsFilter) return false;
      if (coverFilter === 'met' && !(a.target > 0 && a.actual >= a.target)) return false;
      if (coverFilter === 'partial' && !(a.actual > 0 && a.actual < a.target)) return false;
      if (coverFilter === 'none' && a.actual !== 0) return false;
      return true;
    });
  }, [accounts, query, clsFilter, coverFilter]);

  useEffect(() => {
    setPage(1);
  }, [query, clsFilter, coverFilter]);

  const totalPages = Math.max(1, Math.ceil(view.length / pageSize));
  const clampedPage = Math.min(page, totalPages);
  const pageRows = view.slice((clampedPage - 1) * pageSize, (clampedPage - 1) * pageSize + pageSize);

  const kpis = useMemo(() => {
    let targetSum = 0;
    let actualSum = 0;
    accounts.forEach((a) => {
      targetSum += a.target;
      actualSum += a.actual;
    });
    const coverage = targetSum > 0 ? Math.round((actualSum / targetSum) * 100) : 0;
    return { total: accounts.length, targetSum, actualSum, coverage };
  }, [accounts]);

  const allViewSelected = view.length > 0 && view.every((a) => selectedIds.has(a.id));

  function toggleSelect(id: string) {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (next.has(id)) next.delete(id);
      else next.add(id);
      return next;
    });
  }

  function toggleSelectAllView() {
    setSelectedIds((prev) => {
      const next = new Set(prev);
      if (allViewSelected) {
        view.forEach((a) => next.delete(a.id));
      } else {
        view.forEach((a) => next.add(a.id));
      }
      return next;
    });
  }

  function clearSelection() {
    setSelectedIds(new Set());
  }

  function openBulkDrawer() {
    setBulkError('');
    setBulkTarget('');
    setDrawerOpen(false);
    setBulkDrawerOpen(true);
  }

  function closeBulkDrawer() {
    setBulkDrawerOpen(false);
  }

  async function handleBulkSave() {
    const value = parseInt(bulkTarget, 10);
    if (bulkTarget.trim() === '' || Number.isNaN(value) || value < 0) {
      setBulkError('Enter a valid Target Visit number.');
      return;
    }
    setBulkSaving(true);
    setBulkError('');
    try {
      const ids = Array.from(selectedIds);
      const results = await Promise.all(
        ids.map((id) =>
          Cr301_accountsService.update(id, { cfm_visittargetcollection: value })
        )
      );
      const rejectedByApi = results.filter((r) => !r.success).length;
      const accepted = ids.length - rejectedByApi;

      // A 200/204 is not proof the value stuck (Field Security / plugin can silently
      // discard one attribute). Re-read the real server value before reporting success.
      let confirmedCount = accepted;
      let verifyFailed = false;
      if (accepted > 0) {
        try {
          const guidPattern = /^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/;
          const idList = ids.filter((id) => guidPattern.test(id)).join(',');
          const verify = await Cr301_accountsService.getAll({
            select: ['cr301_accountid', 'cfm_visittargetcollection'],
            filter: `Microsoft.Dynamics.CRM.In(PropertyName='cr301_accountid',PropertyValues=[${idList}])`,
          });
          if (verify.success) {
            confirmedCount = (verify.data || []).filter((r) => r.cfm_visittargetcollection === value).length;
          } else {
            verifyFailed = true;
          }
        } catch {
          verifyFailed = true;
        }
      }

      setBulkDrawerOpen(false);
      setSelectedIds(new Set());

      const silentlyRejected = accepted - confirmedCount;
      if (verifyFailed) {
        showToast(
          rejectedByApi === 0
            ? `Target Visit sent for ${accepted} account${accepted !== 1 ? 's' : ''} (could not verify)`
            : `${accepted} sent, ${rejectedByApi} failed (could not verify)`
        );
      } else if (silentlyRejected > 0) {
        showToast(
          `${confirmedCount} actually updated, ${silentlyRejected} rejected by the server` +
            (rejectedByApi > 0 ? `, ${rejectedByApi} failed` : '') +
            ' — check Field Security / a plugin on cfm_visittargetcollection',
          5000
        );
      } else {
        showToast(
          rejectedByApi === 0
            ? `Target Visit updated for ${confirmedCount} account${confirmedCount !== 1 ? 's' : ''}`
            : `${confirmedCount} updated, ${rejectedByApi} failed`
        );
      }
      await load();
    } catch (e) {
      setBulkError(e instanceof Error ? e.message : String(e));
    } finally {
      setBulkSaving(false);
    }
  }

  function openDrawer(row?: AccountRow) {
    setFormError('');
    setBulkDrawerOpen(false);
    if (row) {
      setEditId(row.id);
      setFormName(row.name === '—' ? '' : row.name);
      setFormCode(row.code === '—' ? '' : row.code);
      setFormClsId(row.clsId || '');
      setFormTarget(String(row.target || 0));
    } else {
      setEditId(null);
      setFormName('');
      setFormCode('');
      setFormClsId('');
      setFormTarget('2');
    }
    setDrawerOpen(true);
  }

  function closeDrawer() {
    setDrawerOpen(false);
  }

  /**
   * Bind the classification lookup via the generic Dataverse client using the
   * nav-property @odata.bind form. The generated service's typed update silently
   * dropped the whole PATCH when a bind key was mixed into its payload, so the
   * lookup is set separately through the raw client, which honors @odata.bind.
   * Returns true on success.
   */
  async function setClassificationViaClient(id: string, classId: string | null): Promise<boolean> {
    try {
      const client = getClient(dataSourcesInfo);
      const body: Record<string, unknown> = {};
      if (classId) {
        body[`${CLASS_NAV}@odata.bind`] = `/${CLASS_LOOKUP_TABLE}(${classId})`;
      } else {
        // Disassociate: set the single-valued nav property to null.
        body[`${CLASS_NAV}@odata.bind`] = null;
      }
      const res = await client.updateRecordAsync(ACCOUNTS_TABLE, id, body);
      return !!res.success;
    } catch {
      return false;
    }
  }

  /** Re-read one account's classification value to confirm a write actually persisted. */
  async function readClassification(id: string): Promise<string> {
    try {
      const res = await Cr301_accountsService.get(id, { select: [CLASS_VALUE_FIELD] });
      if (res.success && res.data) {
        return ((res.data as unknown as Record<string, unknown>)[CLASS_VALUE_FIELD] as string) || '';
      }
    } catch {
      /* ignore */
    }
    return '';
  }

  async function handleSave() {
    const name = formName.trim();
    if (!name) {
      setFormError('Account Name is required.');
      return;
    }
    setSaving(true);
    setFormError('');
    try {
      // 1) Write the plain (non-lookup) fields through the generated service.
      //    cr301_monthactualvisits is never written — it's auto-counted from the Visits Log.
      const body: Partial<Omit<Cr301_accountsBase, 'cr301_accountid'>> = {
        cr301_acccountname: name,
        cr301_accountcode: formCode.trim() || null,
        cr301_accounttype: TYPE_INSURANCE,
        cfm_visittargetcollection: parseInt(formTarget, 10) || 0,
      } as Partial<Omit<Cr301_accountsBase, 'cr301_accountid'>>;

      let recordId = editId;
      let result;
      if (editId) {
        result = await Cr301_accountsService.update(editId, body);
      } else {
        result = await Cr301_accountsService.create(body as unknown as Omit<Cr301_accountsBase, 'cr301_accountid'>);
      }
      if (!result.success) {
        throw new Error(operationErrorMessage(result.error, 'Save failed.'));
      }
      if (!recordId) {
        recordId = ((result.data as unknown as Record<string, unknown>)?.cr301_accountid as string) || null;
      }

      // 2) The lookup can't be written through the typed service — Dataverse rejects
      //    a direct write to the read-only "_value" field (and rejects the whole
      //    request, not just that field). Set it via the raw client's @odata.bind
      //    form instead, then verify it actually persisted.
      if (recordId) {
        const desired = formClsId || '';
        const ok = await setClassificationViaClient(recordId, formClsId || null);
        const stored = ok ? await readClassification(recordId) : '';
        if (stored !== desired) {
          // Save the record but be honest that the lookup didn't take.
          setDrawerOpen(false);
          showToast(
            editId
              ? 'Account saved, but Classification could not be set — check field permissions / a plugin'
              : 'Account created, but Classification could not be set — check field permissions / a plugin',
            5000
          );
          await load();
          return;
        }
      }

      setDrawerOpen(false);
      showToast(editId ? 'Account updated' : 'Account created');
      await load();
    } catch (e) {
      setFormError(e instanceof Error ? e.message : String(e));
    } finally {
      setSaving(false);
    }
  }

  function exportCsv() {
    const header = ['Account Name', 'Account Code', 'Classification', 'Type', 'Target Visit', 'Actual Visit'];
    const rows = view.map((a) => [a.name, a.code, a.cls, a.typeName, a.target, a.actual]);
    const csv = [header, ...rows].map((r) => r.map(escapeForCsv).join(',')).join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'accounts_insurance.csv';
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
            <h1 className="acc-title">Accounts</h1>
            <div className="acc-sub">Insurance account master — classification, type and visit coverage</div>
          </div>
          <div className="acc-actions">
            <button className="acc-btn" onClick={exportCsv}>
              <i className="fa-solid fa-download"></i> Export
            </button>
            <button className="acc-btn acc-btn-primary" onClick={() => openDrawer()}>
              <i className="fa-solid fa-plus"></i> Add Account
            </button>
          </div>
        </div>

        <div className="acc-kpis">
          <div className="acc-kpi">
            <div className="acc-lab">
              <i className="fa-solid fa-building-shield"></i> Total Accounts
            </div>
            <div className="acc-val">{loading ? '—' : kpis.total}</div>
          </div>
          <div className="acc-kpi">
            <div className="acc-lab">
              <i className="fa-solid fa-bullseye"></i> Target Visits
            </div>
            <div className="acc-val">{loading ? '—' : kpis.targetSum}</div>
          </div>
          <div className="acc-kpi">
            <div className="acc-lab">
              <i className="fa-solid fa-check-double"></i> Actual Visits
            </div>
            <div className="acc-val">{loading ? '—' : kpis.actualSum}</div>
          </div>
          <div className="acc-kpi">
            <div className="acc-lab">
              <i className="fa-solid fa-gauge-high"></i> Coverage
            </div>
            <div className="acc-val">
              {loading ? '—' : kpis.coverage}
              <small>%</small>
            </div>
          </div>
        </div>

        <div className="acc-bar">
          <div className="acc-search">
            <i className="fa-solid fa-magnifying-glass"></i>
            <input placeholder="Search name or code…" value={query} onChange={(e) => setQuery(e.target.value)} />
          </div>
          <select className="acc-f" value={clsFilter} onChange={(e) => setClsFilter(e.target.value)}>
            <option value="">All Classifications</option>
            <option value="A">Class A</option>
            <option value="B">Class B</option>
            <option value="C">Class C</option>
          </select>
          <select className="acc-f" value={coverFilter} onChange={(e) => setCoverFilter(e.target.value as CoverageFilter)}>
            <option value="">All Coverage</option>
            <option value="met">Target Met</option>
            <option value="partial">Partial</option>
            <option value="none">No Visits</option>
          </select>
        </div>

        {selectedIds.size > 0 && (
          <div className="acc-bulkbar">
            <span>
              <strong>{selectedIds.size}</strong> account{selectedIds.size !== 1 ? 's' : ''} selected
            </span>
            <div className="acc-bulkbar-actions">
              <button className="acc-btn" onClick={clearSelection}>
                Clear
              </button>
              <button className="acc-btn acc-btn-primary" onClick={openBulkDrawer}>
                <i className="fa-solid fa-bullseye"></i> Bulk Update Target
              </button>
            </div>
          </div>
        )}

        <div className="acc-tablewrap">
          <table>
            <thead>
              <tr>
                <th className="acc-check">
                  <input type="checkbox" checked={allViewSelected} onChange={toggleSelectAllView} aria-label="Select all" />
                </th>
                <th>Account Name</th>
                <th>Account Code</th>
                <th>Classification</th>
                <th>Type</th>
                <th className="acc-num">Target Visit</th>
                <th className="acc-num">Actual Visit</th>
              </tr>
            </thead>
            <tbody>
              {loading ? (
                Array.from({ length: 6 }).map((_, i) => (
                  <tr key={i}>
                    {Array.from({ length: 7 }).map((__, j) => (
                      <td key={j}>
                        <div className="acc-skel"></div>
                      </td>
                    ))}
                  </tr>
                ))
              ) : loadError ? (
                <tr>
                  <td colSpan={7}>
                    <div className="acc-state">
                      <i className="fa-solid fa-triangle-exclamation"></i>
                      Could not load accounts.
                      <br />
                      <small>{loadError}</small>
                    </div>
                  </td>
                </tr>
              ) : pageRows.length === 0 ? (
                <tr>
                  <td colSpan={7}>
                    <div className="acc-state">
                      <i className="fa-solid fa-inbox"></i>
                      No insurance accounts match your filters.
                    </div>
                  </td>
                </tr>
              ) : (
                pageRows.map((a) => (
                  <tr key={a.id} className="acc-row" onClick={() => openDrawer(a)}>
                    <td className="acc-check" onClick={(e) => e.stopPropagation()}>
                      <input type="checkbox" checked={selectedIds.has(a.id)} onChange={() => toggleSelect(a.id)} aria-label={`Select ${a.name}`} />
                    </td>
                    <td className="acc-name">{a.name}</td>
                    <td className="acc-code">{a.code}</td>
                    <td>
                      <span className={`acc-cls ${classBadgeClass(a.cls)}`}>{a.cls || '–'}</span>
                    </td>
                    <td>
                      <span className="acc-typ">
                        <i className="fa-solid fa-shield-heart" style={{ color: 'var(--gold)', fontSize: 12 }}></i>
                        {a.typeName}
                      </span>
                    </td>
                    <td className="acc-num">{a.target}</td>
                    <td className="acc-num">
                      <span className="acc-cover">
                        <span className={`acc-dot ${coverageDot(a.actual, a.target)}`}></span>
                        {a.actual}
                      </span>
                    </td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
          <div className="acc-foot">
            <span>
              {view.length} account{view.length !== 1 ? 's' : ''}
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
              <option value={15}>15</option>
              <option value={25}>25</option>
              <option value={50}>50</option>
              <option value={9999}>All</option>
            </select>
          </div>
        </div>
      </div>

      <div
        className={`acc-scrim${drawerOpen || bulkDrawerOpen ? ' on' : ''}`}
        onClick={() => {
          closeDrawer();
          closeBulkDrawer();
        }}
      ></div>
      <aside className={`acc-drawer${drawerOpen ? ' on' : ''}`}>
        <div className="acc-dhead">
          <h2>{editId ? 'Edit Account' : 'Add Account'}</h2>
          <button className="acc-dclose" onClick={closeDrawer}>
            <i className="fa-solid fa-xmark"></i>
          </button>
        </div>
        <div className="acc-dbody">
          <div className={`acc-err${formError ? ' on' : ''}`}>{formError}</div>
          <div className="acc-fld">
            <label>Account Name</label>
            <input placeholder="e.g. Tawuniya Insurance" value={formName} onChange={(e) => setFormName(e.target.value)} />
          </div>
          <div className="acc-fld">
            <label>Account Code</label>
            <input placeholder="e.g. INS-001" value={formCode} onChange={(e) => setFormCode(e.target.value)} />
          </div>
          <div className="acc-fld">
            <label>Account Classification</label>
            <select value={formClsId} onChange={(e) => setFormClsId(e.target.value)}>
              {classesLoadFailed ? (
                <option value="">— (load failed) —</option>
              ) : (
                <>
                  <option value="">—</option>
                  {classes.map((c) => (
                    <option key={c.id} value={c.id}>
                      {c.label || '(blank)'}
                    </option>
                  ))}
                </>
              )}
            </select>
          </div>
          <div className="acc-fld">
            <label>Account Type</label>
            <div className="acc-typebadge">
              <i className="fa-solid fa-shield-heart"></i> Insurance
            </div>
            <div className="acc-hint">All accounts created here are set as Insurance automatically.</div>
          </div>
          <div className="acc-fld">
            <label>Account Target Visit</label>
            <input type="number" min={0} value={formTarget} onChange={(e) => setFormTarget(e.target.value)} />
          </div>
          <div className="acc-hint" style={{ marginTop: -6 }}>
            <i className="fa-solid fa-circle-info"></i> Actual visits are counted automatically from the Visits Log.
          </div>
        </div>
        <div className="acc-dfoot">
          <button className="acc-btn acc-btn-ghost" onClick={closeDrawer}>
            Cancel
          </button>
          <button className="acc-btn acc-btn-primary" disabled={saving} onClick={handleSave}>
            {saving ? 'Saving…' : 'Save Account'}
          </button>
        </div>
      </aside>

      <aside className={`acc-drawer${bulkDrawerOpen ? ' on' : ''}`}>
        <div className="acc-dhead">
          <h2>
            <i className="fa-solid fa-bullseye" style={{ color: 'var(--gold)', fontSize: 16 }}></i> Bulk Update Target Visit
          </h2>
          <button className="acc-dclose" onClick={closeBulkDrawer}>
            <i className="fa-solid fa-xmark"></i>
          </button>
        </div>
        <div className="acc-dbody">
          <div className={`acc-err${bulkError ? ' on' : ''}`}>{bulkError}</div>
          <div className="acc-hint" style={{ marginBottom: 16 }}>
            This sets <b>Target Visit</b> to the same value for all <b>{selectedIds.size}</b> selected account
            {selectedIds.size !== 1 ? 's' : ''}. Actual Visit is never affected.
          </div>
          <div className="acc-fld">
            <label>Target Visit</label>
            <input type="number" min={0} placeholder="e.g. 4" value={bulkTarget} onChange={(e) => setBulkTarget(e.target.value)} />
          </div>
        </div>
        <div className="acc-dfoot">
          <button className="acc-btn acc-btn-ghost" onClick={closeBulkDrawer}>
            Cancel
          </button>
          <button className="acc-btn acc-btn-primary" disabled={bulkSaving} onClick={handleBulkSave}>
            {bulkSaving ? 'Updating…' : `Apply to ${selectedIds.size}`}
          </button>
        </div>
      </aside>

      <div className={`acc-toast${toastOn ? ' on' : ''}`}>{toastMsg}</div>
    </>
  );
}