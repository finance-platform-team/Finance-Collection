import { useEffect, useMemo, useRef, useState } from 'react';
import { Cr301_agentsService } from '../generated/services/Cr301_agentsService';
import { SystemusersService } from '../generated/services/SystemusersService';
import type { Cr301_agentsBase } from '../generated/models/Cr301_agentsModel';
import { operationErrorMessage } from '../lib/errors';
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
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [query, setQuery] = useState('');
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(15);

  const [drawerOpen, setDrawerOpen] = useState(false);
  const [editId, setEditId] = useState<string | null>(null);
  const [formName, setFormName] = useState('');
  const [formUserId, setFormUserId] = useState('');
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

  async function load() {
    setLoading(true);
    setLoadError(null);
    try {
      const [userOptions, result] = await Promise.all([
        fetchUsers(),
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
    if (id) {
      const a = agents.find((x) => x.id === id);
      setEditId(id);
      setFormName(a && a.name !== '—' ? a.name : '');
      setFormUserId(a?.userId || '');
    } else {
      setEditId(null);
      setFormName('');
      setFormUserId('');
    }
    setDrawerOpen(true);
  }

  function closeDrawer() {
    setDrawerOpen(false);
  }

  async function handleSave() {
    const name = formName.trim();
    if (!name) {
      setFormError('Agent Name is required.');
      return;
    }
    setSaving(true);
    setFormError('');
    try {
      const body: Record<string, unknown> = {
        cr301_newcolumn: name,
        cfm_iscollection: true,
      };
      if (formUserId) {
        body['cr301_Agent@odata.bind'] = `/systemusers(${formUserId})`;
      } else if (editId) {
        body['cr301_Agent@odata.bind'] = null;
      }

      let result;
      if (editId) {
        result = await Cr301_agentsService.update(editId, body as Partial<Omit<Cr301_agentsBase, 'cr301_agentid'>>);
      } else {
        result = await Cr301_agentsService.create(body as unknown as Omit<Cr301_agentsBase, 'cr301_agentid'>);
      }
      if (!result.success) {
        throw new Error(operationErrorMessage(result.error, 'Save failed.'));
      }
      setDrawerOpen(false);
      showToast(editId ? 'Agent updated' : 'Agent created');
      await load();
    } catch (e) {
      setFormError(e instanceof Error ? e.message : String(e));
    } finally {
      setSaving(false);
    }
  }

  function exportCsv() {
    const escapeForCsv = (value: string) => `"${value.replace(/"/g, '""')}"`;
    const header = ['Agent Name', 'Agent User', 'Title'];
    const rows = view.map((a) => [a.name, a.user, a.title]);
    const csv = [header, ...rows].map((r) => r.map(escapeForCsv).join(',')).join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'agents.csv';
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

        <div className="agt-kpis">
          <div className="acc-kpi">
            <div className="acc-lab">
              <i className="fa-solid fa-user-tie"></i> Total Agents
            </div>
            <div className="acc-val">{loading ? '—' : kpis.total}</div>
          </div>
          <div className="acc-kpi">
            <div className="acc-lab">
              <i className="fa-solid fa-link"></i> Linked Users
            </div>
            <div className="acc-val">{loading ? '—' : kpis.linked}</div>
          </div>
        </div>

        <div className="acc-bar">
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
          <h2>{editId ? 'Edit Agent' : 'Add Agent'}</h2>
          <button className="acc-dclose" onClick={closeDrawer}>
            <i className="fa-solid fa-xmark"></i>
          </button>
        </div>
        <div className="acc-dbody">
          <div className={`acc-err${formError ? ' on' : ''}`}>{formError}</div>
          <div className="acc-fld">
            <label>Agent Name</label>
            <input placeholder="e.g. S. Al-Harbi" value={formName} onChange={(e) => setFormName(e.target.value)} />
          </div>
          <div className="acc-fld">
            <label>Agent User (System User)</label>
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
            {saving ? 'Saving…' : 'Save Agent'}
          </button>
        </div>
      </aside>

      <div className={`acc-toast${toastOn ? ' on' : ''}`}>{toastMsg}</div>
    </>
  );
}
