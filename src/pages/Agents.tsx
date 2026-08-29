import { useEffect, useMemo, useRef, useState } from 'react';
import { Cr301_agentsService } from '../generated/services/Cr301_agentsService';
import { SystemusersService } from '../generated/services/SystemusersService';
import { Hr_employeesService } from '../generated/services/Hr_employeesService';
import type { Cr301_agentsBase } from '../generated/models/Cr301_agentsModel';
import { operationErrorMessage } from '../lib/errors';
import { escapeForCsv, downloadCsv } from '../lib/csvExport';
import './Agents.css';

/**
 * cfm_iscollection is a real Dataverse column on cr301_agent that keeps this
 * grid scoped to the collection-agents view (mirrors Ref.html's server-side
 * filter). It wasn't picked up when this table was wired up as a Code App
 * data source, so it's absent from Cr301_agentsModel — filter/body values
 * below are passed through untyped rather than via the generated fields.
 */
const COLLECTION_FILTER = 'cfm_iscollection eq true';

interface UserOption {
  id: string;
  displayName: string;
  title: string;
  domainName: string;
}

interface EmployeeOption {
  id: string;
  name: string;
}

interface AgentRow {
  id: string;
  name: string;
  userId: string;
  user: string;
  userFull: string;
  title: string;
}

function loginOf(domainName: string): string {
  return (domainName || '').split('@')[0] || '';
}

export default function AgentsPage() {
  const [agents, setAgents] = useState<AgentRow[]>([]);
  const [users, setUsers] = useState<UserOption[]>([]);
  const [employees, setEmployees] = useState<EmployeeOption[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [query, setQuery] = useState('');
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(15);

  const [drawerOpen, setDrawerOpen] = useState(false);
  const [editId, setEditId] = useState<string | null>(null);
  const [formName, setFormName] = useState('');
  const [formUserId, setFormUserId] = useState('');
  const [employeeQuery, setEmployeeQuery] = useState('');
  const [employeePickerOpen, setEmployeePickerOpen] = useState(false);
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

  async function fetchUsers(): Promise<UserOption[]> {
    const result = await SystemusersService.getAll({
      select: ['systemuserid', 'fullname', 'domainname', 'title'],
      filter: 'isdisabled eq false',
      orderBy: ['fullname'],
      top: 500,
    });
    if (!result.success) return [];
    return (result.data || []).map((u) => ({
      id: u.systemuserid,
      displayName: u.fullname || u.domainname || '(user)',
      title: (u.title || '').trim(),
      domainName: u.domainname || '',
    }));
  }

  async function fetchEmployees(): Promise<EmployeeOption[]> {
    const result = await Hr_employeesService.getAll({
      select: ['hr_employeeid', 'hr_fullname'],
      filter: 'statecode eq 0',
      orderBy: ['hr_fullname'],
      maxPageSize: 5000,
    });
    if (!result.success) return [];
    return (result.data || []).map((e) => ({ id: e.hr_employeeid, name: e.hr_fullname || '(employee)' }));
  }

  async function load() {
    setLoading(true);
    setLoadError(null);
    try {
      const [userOptions, employeeOptions, result] = await Promise.all([
        fetchUsers(),
        fetchEmployees(),
        Cr301_agentsService.getAll({
          select: ['cr301_agentid', 'cr301_newcolumn', '_cr301_agent_value'],
          filter: COLLECTION_FILTER,
          orderBy: ['createdon desc'],
        }),
      ]);
      if (!result.success) {
        throw new Error(operationErrorMessage(result.error, 'Could not load agents.'));
      }

      const userMap = new Map(userOptions.map((u) => [u.id, u]));
      const rows: AgentRow[] = (result.data || []).map((a) => {
        const userId = a._cr301_agent_value || '';
        const u = userId ? userMap.get(userId) : undefined;
        const userFull = u?.displayName || '';
        const login = u ? loginOf(u.domainName) : '';
        return {
          id: a.cr301_agentid,
          name: a.cr301_newcolumn || '—',
          userId,
          user: login || userFull || '—',
          userFull,
          title: u?.title || '',
        };
      });
      setUsers(userOptions);
      setEmployees(employeeOptions);
      setAgents(rows);
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
    if (!q) return agents;
    return agents.filter((a) => `${a.name} ${a.user} ${a.userFull} ${a.title}`.toLowerCase().includes(q));
  }, [agents, query]);

  useEffect(() => {
    setPage(1);
  }, [query]);

  const totalPages = Math.max(1, Math.ceil(view.length / pageSize));
  const clampedPage = Math.min(page, totalPages);
  const pageRows = view.slice((clampedPage - 1) * pageSize, (clampedPage - 1) * pageSize + pageSize);

  const kpis = useMemo(
    () => ({
      total: agents.length,
      linked: agents.filter((a) => a.userId).length,
    }),
    [agents]
  );

  function openDrawer(id?: string) {
    setFormError('');
    setEmployeePickerOpen(false);
    if (id) {
      const a = agents.find((x) => x.id === id);
      const name = a && a.name !== '—' ? a.name : '';
      setEditId(id);
      setFormName(name);
      setEmployeeQuery(name);
      setFormUserId(a?.userId || '');
    } else {
      setEditId(null);
      setFormName('');
      setEmployeeQuery('');
      setFormUserId('');
    }
    setDrawerOpen(true);
  }

  function closeDrawer() {
    setDrawerOpen(false);
  }

  const employeeMatches = useMemo(() => {
    const q = employeeQuery.trim().toLowerCase();
    const list = q ? employees.filter((e) => e.name.toLowerCase().includes(q)) : employees;
    return list.slice(0, 30);
  }, [employees, employeeQuery]);

  // First 3 words of a name, normalized — Arabic names commonly carry 3-4+ parts
  // (first + father's + grandfather's + family name) and HR vs. system-user
  // records don't always agree on how many are kept, so matching the full string
  // is too strict. Matching just the first 3 tolerates a trailing name differing
  // or being dropped, while still requiring real agreement on who the person is.
  function firstWords(name: string, n: number): string {
    return name.trim().toLowerCase().split(/\s+/).slice(0, n).join(' ');
  }

  // Employee doubles as the Agent Name field now (no separate free-text input) —
  // picking a suggestion sets the name AND searches the already-loaded system
  // users for a name match to auto-select Agent User. No match -> Agent User is
  // left untouched, per requirement (don't guess).
  function pickEmployee(employee: EmployeeOption) {
    setEmployeeQuery(employee.name);
    setFormName(employee.name);
    setEmployeePickerOpen(false);
    const key = firstWords(employee.name, 3);
    const match = key ? users.find((u) => firstWords(u.displayName, 3) === key) : undefined;
    if (match) {
      setFormUserId(match.id);
    } else {
      showToast(`No matching system user found for "${employee.name}" — select one manually.`);
    }
  }

  async function handleSave() {
    const name = formName.trim();
    if (!name) {
      setFormError('Employee name is required.');
      return;
    }
    if (!formUserId) {
      setFormError('Agent User is required.');
      return;
    }
    setSaving(true);
    setFormError('');
    try {
      const body: Record<string, unknown> = {
        cr301_newcolumn: name,
        cfm_iscollection: true,
        'cr301_Agent@odata.bind': `/systemusers(${formUserId})`,
      };

      // cr301_agents has a uniqueness constraint on the Agent User lookup — a
      // hidden/non-collection agent for this system user (cfm_iscollection false
      // or unset, so it doesn't show in this filtered list) still holds that key.
      // Reactivate/update it instead of racing straight into a duplicate-key error.
      let targetId = editId;
      if (!targetId) {
        const existing = await Cr301_agentsService.getAll({
          select: ['cr301_agentid'],
          filter: `_cr301_agent_value eq ${formUserId}`,
          top: 1,
        });
        if (existing.success && existing.data && existing.data.length > 0) {
          targetId = existing.data[0].cr301_agentid;
        }
      }

      const result = targetId
        ? await Cr301_agentsService.update(targetId, body as Partial<Omit<Cr301_agentsBase, 'cr301_agentid'>>)
        : await Cr301_agentsService.create(body as unknown as Omit<Cr301_agentsBase, 'cr301_agentid'>);
      if (!result.success) {
        throw new Error(operationErrorMessage(result.error, 'Save failed.'));
      }
      setDrawerOpen(false);
      showToast(editId ? 'Agent updated' : targetId ? 'Agent reactivated' : 'Agent created');
      await load();
    } catch (e) {
      setFormError(e instanceof Error ? e.message : String(e));
    } finally {
      setSaving(false);
    }
  }

  function exportCsv() {
    const header = ['Agent Name', 'Agent User', 'Title'];
    const rows = view.map((a) => [a.name, a.user, a.title]);
    downloadCsv('agents.csv', [header, ...rows].map((r) => r.map(escapeForCsv).join(',')));
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
            <h1 className="acc-title">Agents</h1>
            <div className="acc-sub">Collection agents, system users and team assignment</div>
          </div>
          <div className="acc-actions">
            <button className="acc-btn" onClick={exportCsv}>
              <i className="fa-solid fa-download"></i> Export
            </button>
            <button className="acc-btn acc-btn-primary" onClick={() => openDrawer()}>
              <i className="fa-solid fa-plus"></i> Add Agent
            </button>
          </div>
        </div>

        <div className="kpi-grid">
          <div className="kpi-card">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-gold">
                <i className="fa-solid fa-user-tie"></i>
              </span>
              <span className="kpi-label">Total agents</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.total}</div>
            <div className="kpi-context">&nbsp;</div>
          </div>
          <div className="kpi-card kpi-ok">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-ok">
                <i className="fa-solid fa-link"></i>
              </span>
              <span className="kpi-label">Linked users</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.linked}</div>
            <div className="kpi-context">
              {loading ? '' : kpis.total > 0 ? `${Math.round((kpis.linked / kpis.total) * 100)}% linked` : '—'}
            </div>
          </div>
        </div>

        <div className="tbl-shell">
          <div className="tbl-hdr">
            <div className="acc-search">
              <i className="fa-solid fa-magnifying-glass"></i>
              <input placeholder="Search agent, user or title…" value={query} onChange={(e) => setQuery(e.target.value)} />
            </div>
          </div>

          <div className="acc-tablewrap">
          <table>
            <thead>
              <tr>
                <th>Agent Name</th>
                <th>Agent User</th>
                <th>Title</th>
              </tr>
            </thead>
            <tbody>
              {loading ? (
                Array.from({ length: 6 }).map((_, i) => (
                  <tr key={i}>
                    {Array.from({ length: 3 }).map((__, j) => (
                      <td key={j}>
                        <div className="acc-skel"></div>
                      </td>
                    ))}
                  </tr>
                ))
              ) : loadError ? (
                <tr>
                  <td colSpan={3}>
                    <div className="acc-state">
                      <i className="fa-solid fa-triangle-exclamation"></i>
                      Could not load agents.
                      <br />
                      <small>{loadError}</small>
                    </div>
                  </td>
                </tr>
              ) : pageRows.length === 0 ? (
                <tr>
                  <td colSpan={3}>
                    <div className="acc-state">
                      <i className="fa-solid fa-inbox"></i>
                      No agents match your filters.
                    </div>
                  </td>
                </tr>
              ) : (
                pageRows.map((a) => (
                  <tr key={a.id} className="acc-row" onClick={() => openDrawer(a.id)}>
                    <td className="acc-name">{a.name}</td>
                    <td className="agt-user-cell">{a.user}</td>
                    <td className="agt-title-cell">{a.title ? a.title : <span className="agt-muted">—</span>}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
          <div className="acc-foot">
            <span>
              {view.length} agent{view.length !== 1 ? 's' : ''}
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
          <h2>{editId ? 'Edit Agent' : 'Add Agent'}</h2>
          <button className="acc-dclose" onClick={closeDrawer}>
            <i className="fa-solid fa-xmark"></i>
          </button>
        </div>
        <div className="acc-dbody">
          <div className={`acc-err${formError ? ' on' : ''}`}>{formError}</div>
          <div className="acc-combo" style={{ marginBottom: 16 }}>
            <div className="acc-fld" style={{ marginBottom: 0 }}>
              <label>
                Employee <span className="dm-req">*</span>
              </label>
              <input
                placeholder="Search employee or type a name…"
                autoComplete="off"
                autoFocus
                value={employeeQuery}
                onChange={(e) => {
                  setEmployeeQuery(e.target.value);
                  setFormName(e.target.value);
                  setEmployeePickerOpen(true);
                }}
                onFocus={() => setEmployeePickerOpen(true)}
                onBlur={() => setTimeout(() => setEmployeePickerOpen(false), 150)}
              />
            </div>
            {employeePickerOpen && (
              <div className="acc-combo-list">
                {employeeMatches.length === 0 ? (
                  <div className="acc-combo-empty">No matches.</div>
                ) : (
                  employeeMatches.map((e) => (
                    <div key={e.id} className="acc-combo-item" onMouseDown={(ev) => ev.preventDefault()} onClick={() => pickEmployee(e)}>
                      {e.name}
                    </div>
                  ))
                )}
              </div>
            )}
            <div className="acc-hint">Sets the Agent Name and auto-selects the matching Agent User below when found.</div>
          </div>
          <div className="acc-fld">
            <label>
              Agent User (System User) <span className="dm-req">*</span>
            </label>
            <select value={formUserId} onChange={(e) => setFormUserId(e.target.value)}>
              <option value="">— Select user —</option>
              {users.map((u) => (
                <option key={u.id} value={u.id}>
                  {u.displayName}
                  {u.title ? ` — ${u.title}` : ''}
                </option>
              ))}
            </select>
            <div className="acc-hint">Title and team are read from the linked system user.</div>
          </div>
        </div>
        <div className="acc-dfoot">
          <button className="acc-btn acc-btn-ghost" onClick={closeDrawer}>
            Cancel
          </button>
          <button className="acc-btn acc-btn-primary" disabled={saving} onClick={handleSave}>
            {saving ? (
              <>
                <i className="fa-solid fa-spinner fa-spin"></i> Saving…
              </>
            ) : (
              'Save Agent'
            )}
          </button>
        </div>
      </aside>

      <div className={`acc-toast${toastOn ? ' on' : ''}`}>{toastMsg}</div>
    </>
  );
}
