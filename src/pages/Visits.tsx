import { useEffect, useMemo, useRef, useState } from 'react';
import { Cr301_visitsService } from '../generated/services/Cr301_visitsService';
import { Cr301_accountsService } from '../generated/services/Cr301_accountsService';
import { Cr301_agentsService } from '../generated/services/Cr301_agentsService';
import type { Cr301_visitsBase, Cr301_visitscr301_visittype } from '../generated/models/Cr301_visitsModel';
import { operationErrorMessage } from '../lib/errors';
import './Visits.css';

const TYPE_INSURANCE = 983080002; // cr301_accounttype — Insurance (same value used on the Accounts page)

const VISIT_TYPE_SINGLE: Cr301_visitscr301_visittype = 983080000; // 'Single Visit'
const VISIT_TYPE_DOUBLE: Cr301_visitscr301_visittype = 983080001; // 'Double Visit'

const CATEGORY_OPTIONS = [
  { value: 766340002, label: 'Open' },
  { value: 766340000, label: 'Partially Closed' },
  { value: 766340001, label: 'Closed' },
];

/**
 * cfm_category, cfm_tagename, cfm_revisitdate and cfm_agent2 (Task Owner 2) do not exist
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
  cfm_category?: number;
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
  isDouble: boolean;
  category: number | null;
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

function categoryChip(category: number | null): { cls: string; label: string } {
  if (category === 766340002) return { cls: 'vis-cat-open', label: 'Open' };
  if (category === 766340000) return { cls: 'vis-cat-partial', label: 'Partially Closed' };
  if (category === 766340001) return { cls: 'vis-cat-closed', label: 'Closed' };
  return { cls: 'vis-cat-none', label: '—' };
}

export default function VisitsPage() {
  const [visits, setVisits] = useState<VisitRow[]>([]);
  const [accounts, setAccounts] = useState<AccountOption[]>([]);
  const [agents, setAgents] = useState<AgentOption[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [query, setQuery] = useState('');
  const [catFilter, setCatFilter] = useState('');
  const [typeFilter, setTypeFilter] = useState<TypeFilter>('');
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(15);

  const [drawerOpen, setDrawerOpen] = useState(false);
  const [editId, setEditId] = useState<string | null>(null);
  const [formAccountId, setFormAccountId] = useState('');
  const [formAgentId, setFormAgentId] = useState('');
  const [formDate, setFormDate] = useState('');
  const [formCategory, setFormCategory] = useState('766340002');
  const [formVisitType, setFormVisitType] = useState<'single' | 'double'>('single');
  const [formAgent2Id, setFormAgent2Id] = useState('');
  const [formRevisit, setFormRevisit] = useState('');
  const [formNotes, setFormNotes] = useState('');
  const [formError, setFormError] = useState('');
  const [saving, setSaving] = useState(false);

  const [toastMsg, setToastMsg] = useState('');
  const [toastOn, setToastOn] = useState(false);
  const toastTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  function showToast(msg: string) {
    setToastMsg(msg);
    setToastOn(true);
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToastOn(false), 2200);
  }

  async function load() {
    setLoading(true);
    setLoadError(null);
    try {
      const [accResult, agentResult, visitResult] = await Promise.all([
        Cr301_accountsService.getAll({
          select: ['cr301_accountid', 'cr301_acccountname', 'cr301_accountcode'],
          filter: `cr301_accounttype eq ${TYPE_INSURANCE}`,
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
            'cr301_visittype',
            'cr301_notes',
            '_cr301_account_value',
            '_cr301_agent_value',
            // Not in the generated model — see the module comment above.
            'cfm_category',
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
          isDouble: r.cr301_visittype === VISIT_TYPE_DOUBLE,
          category: typeof extra.cfm_category === 'number' ? extra.cfm_category : null,
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

  const view = useMemo(() => {
    const q = query.trim().toLowerCase();
    return visits.filter((v) => {
      if (q && !`${v.owner} ${v.owner2} ${v.company} ${v.code} ${v.notes}`.toLowerCase().includes(q)) return false;
      if (catFilter && String(v.category) !== catFilter) return false;
      if (typeFilter === 'single' && v.isDouble) return false;
      if (typeFilter === 'double' && !v.isDouble) return false;
      return true;
    });
  }, [visits, query, catFilter, typeFilter]);

  useEffect(() => {
    setPage(1);
  }, [query, catFilter, typeFilter]);

  const totalPages = Math.max(1, Math.ceil(view.length / pageSize));
  const clampedPage = Math.min(page, totalPages);
  const pageRows = view.slice((clampedPage - 1) * pageSize, (clampedPage - 1) * pageSize + pageSize);

  const kpis = useMemo(() => {
    const uniqueCompanies = new Set(visits.filter((v) => v.accountId).map((v) => v.accountId)).size;
    const open = visits.filter((v) => v.category === 766340002).length;
    const closed = visits.filter((v) => v.category === 766340001).length;
    return { total: visits.length, uniqueCompanies, open, closed };
  }, [visits]);

  const selectedAccount = useMemo(() => accounts.find((a) => a.id === formAccountId) || null, [accounts, formAccountId]);

  function openDrawer(row?: VisitRow) {
    setFormError('');
    if (row) {
      setEditId(row.id);
      setFormAccountId(row.accountId);
      setFormAgentId(row.agentId);
      setFormDate(row.date ? row.date.substring(0, 10) : '');
      setFormCategory(row.category != null ? String(row.category) : '766340002');
      setFormVisitType(row.isDouble ? 'double' : 'single');
      setFormAgent2Id(row.agent2Id);
      setFormRevisit(row.revisit ? row.revisit.substring(0, 10) : '');
      setFormNotes(row.notes);
    } else {
      setEditId(null);
      setFormAccountId('');
      setFormAgentId('');
      setFormDate(new Date().toISOString().substring(0, 10));
      setFormCategory('766340002');
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

    setSaving(true);
    setFormError('');
    try {
      const body: Record<string, unknown> = {
        'cr301_Account@odata.bind': `/cr301_accounts(${formAccountId})`,
        'cr301_Agent@odata.bind': `/cr301_agents(${formAgentId})`,
        cr301_visittype: isDouble ? VISIT_TYPE_DOUBLE : VISIT_TYPE_SINGLE,
        cr301_visitdate: formDate ? new Date(formDate).toISOString() : null,
        cr301_notes: formNotes.trim() || null,
        // Not in the generated model — see the module comment above.
        cfm_category: parseInt(formCategory, 10),
        cfm_tagename: TAG_COLLECTION,
        cfm_revisitdate: formRevisit ? new Date(formRevisit).toISOString() : null,
      };
      if (isDouble && formAgent2Id) {
        body[`${AGENT2_BIND_NAV}@odata.bind`] = `/cr301_agents(${formAgent2Id})`;
      } else if (editId) {
        // Switched back to Single (or cleared) on an existing visit — unlink the second owner.
        body[`${AGENT2_BIND_NAV}@odata.bind`] = null;
      }
      if (!editId) {
        // cr301_visitstatus is required (non-optional) on Cr301_visitsBase, but Ref.html
        // never sets it — defaulting new visits to Open (983080000); see chat summary.
        body.cr301_visitstatus = 983080000;
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
      setDrawerOpen(false);
      showToast(editId ? 'Visit updated' : 'Visit recorded');
      await load();
    } catch (e) {
      setFormError(e instanceof Error ? e.message : String(e));
    } finally {
      setSaving(false);
    }
  }

  function exportCsv() {
    const escapeForCsv = (value: string) => `"${value.replace(/"/g, '""')}"`;
    const header = ['Task Owner', 'Task Owner 2', 'Company', 'Company Code', 'Visit Date', 'Category', 'Visit Type', 'Revisit Date', 'Comments'];
    const rows = view.map((v) => [
      v.owner,
      v.owner2,
      v.company,
      v.code,
      fmtDate(v.date),
      categoryChip(v.category).label,
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
            <div className="acc-sub">Recorded visits — Task Owner · Company · Visit Date · Category · Visit Type · Revisit · Comments</div>
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

        <div className="acc-kpis">
          <div className="acc-kpi">
            <div className="acc-lab">
              <i className="fa-solid fa-clipboard-check"></i> Total Visits
            </div>
            <div className="acc-val">{loading ? '—' : kpis.total}</div>
          </div>
          <div className="acc-kpi">
            <div className="acc-lab">
              <i className="fa-solid fa-building"></i> Unique Companies
            </div>
            <div className="acc-val">{loading ? '—' : kpis.uniqueCompanies}</div>
          </div>
          <div className="acc-kpi">
            <div className="acc-lab">
              <i className="fa-solid fa-door-open"></i> Open
            </div>
            <div className="acc-val">{loading ? '—' : kpis.open}</div>
          </div>
          <div className="acc-kpi">
            <div className="acc-lab">
              <i className="fa-solid fa-circle-check"></i> Closed
            </div>
            <div className="acc-val">{loading ? '—' : kpis.closed}</div>
          </div>
        </div>

        <div className="acc-bar">
          <div className="acc-search">
            <i className="fa-solid fa-magnifying-glass"></i>
            <input placeholder="Search owner, company or notes…" value={query} onChange={(e) => setQuery(e.target.value)} />
          </div>
          <select className="acc-f" value={catFilter} onChange={(e) => setCatFilter(e.target.value)}>
            <option value="">All Categories</option>
            <option value="766340002">Open</option>
            <option value="766340000">Partially Closed</option>
            <option value="766340001">Closed</option>
          </select>
          <select className="acc-f" value={typeFilter} onChange={(e) => setTypeFilter(e.target.value as TypeFilter)}>
            <option value="">All Types</option>
            <option value="single">Single visit</option>
            <option value="double">Double visit</option>
          </select>
        </div>

        <div className="acc-tablewrap">
          <table>
            <thead>
              <tr>
                <th>Task Owner</th>
                <th>Task Owner 2</th>
                <th>Company</th>
                <th>Visit Date</th>
                <th>Category</th>
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
                  const chip = categoryChip(v.category);
                  return (
                    <tr key={v.id} className="acc-row" onClick={() => openDrawer(v)}>
                      <td className="acc-name">{v.owner}</td>
                      <td>{v.owner2 ? <span className="acc-name">{v.owner2}</span> : <span className="vis-muted">—</span>}</td>
                      <td>
                        <div className="acc-name">{v.company}</div>
                        {v.code && <div className="acc-code">{v.code}</div>}
                      </td>
                      <td>{fmtDate(v.date)}</td>
                      <td>
                        <span className={`vis-chip ${chip.cls}`}>{chip.label}</span>
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
              <option value={15}>15</option>
              <option value={25}>25</option>
              <option value={50}>50</option>
              <option value={9999}>All</option>
            </select>
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
            <label>Category</label>
            <select value={formCategory} onChange={(e) => setFormCategory(e.target.value)}>
              {CATEGORY_OPTIONS.map((c) => (
                <option key={c.value} value={c.value}>
                  {c.label}
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
