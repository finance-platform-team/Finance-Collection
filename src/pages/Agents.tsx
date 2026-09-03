import Icon from '../components/Icon';
import { useEffect, useMemo, useRef, useState } from 'react';
import { getClient } from '@microsoft/power-apps/data';
import { dataSourcesInfo } from '../../.power/schemas/appschemas/dataSourcesInfo';
import { Cr301_agentsService } from '../generated/services/Cr301_agentsService';
import { SystemusersService } from '../generated/services/SystemusersService';
import { Hr_employeesService } from '../generated/services/Hr_employeesService';
import type { Cr301_agentsBase } from '../generated/models/Cr301_agentsModel';
import { operationErrorMessage } from '../lib/errors';
import { escapeForCsv, downloadCsv } from '../lib/csvExport';
import type { Region } from '../lib/useRegion';
import './Agents.css';

/**
 * cfm_iscollection is a real Dataverse column on cr301_agent that keeps this
 * grid scoped to the collection-agents view (mirrors Ref.html's server-side
 * filter). It wasn't picked up when this table was wired up as a Code App
 * data source, so it's absent from Cr301_agentsModel — filter/body values
 * below are passed through untyped rather than via the generated fields.
 */
const COLLECTION_FILTER = 'cfm_iscollection eq true';

/**
 * cr18c_area is a real Dataverse choice field with 5 values (North, South,
 * Remote, Jeddah, KSA) — not a plain Egypt/KSA pair. Per explicit user
 * direction, the Region field in this UI only offers "Egypt" / "KSA": KSA
 * maps straight to the KSA choice, and Egypt maps to North (the closest
 * single-value stand-in for "not KSA" — adjust here if a different choice
 * should represent Egypt). Any agent whose stored value is South/Remote/
 * Jeddah (set outside this form) still displays as "Egypt" in the table,
 * so nothing looks unset just because it doesn't literally say "North".
 */
const AREA_KSA = 989230004;
const AREA_EGYPT = 989230000; // North — see comment above

const AGENTS_TABLE = 'cr301_agents';
const AGENT_NAV = 'cr301_Agent';
const AGENT_VALUE_FIELD = '_cr301_agent_value';

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
  region: 'EG' | 'KSA' | '';
}

function loginOf(domainName: string): string {
  return (domainName || '').split('@')[0] || '';
}

interface AgentsPageProps {
  region: Region | null;
}

export default function AgentsPage({ region }: AgentsPageProps) {
  const [agents, setAgents] = useState<AgentRow[]>([]);
  const [users, setUsers] = useState<UserOption[]>([]);
  const [employees, setEmployees] = useState<EmployeeOption[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [userLoadError, setUserLoadError] = useState<string | null>(null);
  const [employeeLoadError, setEmployeeLoadError] = useState<string | null>(null);

  const [query, setQuery] = useState('');
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(15);

  const [drawerOpen, setDrawerOpen] = useState(false);
  const [editId, setEditId] = useState<string | null>(null);
  const [formName, setFormName] = useState('');
  const [formUserId, setFormUserId] = useState('');
  const [formRegion, setFormRegion] = useState<'EG' | 'KSA' | ''>('');
  const [employeeQuery, setEmployeeQuery] = useState('');
  const [employeePickerOpen, setEmployeePickerOpen] = useState(false);
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

  /**
   * Both fetches used to swallow a failed result down to `[]`, which is
   * indistinguishable from "this org genuinely has zero active records" — the
   * Add Agent drawer would just show an empty, unexplained picker either way.
   * Throwing here instead lets load() capture and surface the real reason
   * (permissions, a renamed field, network) so it's actually diagnosable.
   */
  async function fetchUsers(): Promise<UserOption[]> {
    const result = await SystemusersService.getAll({
      select: ['systemuserid', 'fullname', 'domainname', 'title'],
      filter: 'isdisabled eq false',
      orderBy: ['fullname'],
      top: 500,
    });
    if (!result.success) {
      throw new Error(operationErrorMessage(result.error, 'Could not load system users.'));
    }
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
    if (!result.success) {
      throw new Error(operationErrorMessage(result.error, 'Could not load employees.'));
    }
    // Records with a blank/whitespace-only hr_fullname sort first (orderBy hr_fullname
    // ascending puts '' and ' ' before any real name) and render as invisible rows in
    // the picker — with the list capped to 30, they crowded out every real name from
    // the unfiltered view entirely. Drop them; they're not a usable pick anyway.
    return (result.data || [])
      .map((e) => ({ id: e.hr_employeeid, name: (e.hr_fullname || '').trim() }))
      .filter((e) => e.name);
  }

  async function load() {
    setLoading(true);
    setLoadError(null);
    setUserLoadError(null);
    setEmployeeLoadError(null);
    try {
      // Users/employees are independent, drawer-only concerns — a failure in either
      // (permissions, a renamed field) shouldn't take down the agents grid itself,
      // so each is caught on its own instead of failing the whole Promise.all.
      const [userSettled, employeeSettled, result] = await Promise.all([
        fetchUsers()
          .then((data) => ({ data, error: null }))
          .catch((e) => ({ data: [] as UserOption[], error: e instanceof Error ? e.message : String(e) })),
        fetchEmployees()
          .then((data) => ({ data, error: null }))
          .catch((e) => ({ data: [] as EmployeeOption[], error: e instanceof Error ? e.message : String(e) })),
        Cr301_agentsService.getAll({
          select: ['cr301_agentid', 'cr301_newcolumn', '_cr301_agent_value', 'cr18c_area'],
          filter: COLLECTION_FILTER,
          orderBy: ['createdon desc'],
        }),
      ]);
      if (!result.success) {
        throw new Error(operationErrorMessage(result.error, 'Could not load agents.'));
      }
      if (userSettled.error) setUserLoadError(userSettled.error);
      if (employeeSettled.error) setEmployeeLoadError(employeeSettled.error);

      const userOptions = userSettled.data;
      const employeeOptions = employeeSettled.data;
      const userMap = new Map(userOptions.map((u) => [u.id, u]));
      const rows: AgentRow[] = (result.data || []).map((a) => {
        const userId = a._cr301_agent_value || '';
        const u = userId ? userMap.get(userId) : undefined;
        const userFull = u?.displayName || '';
        const login = u ? loginOf(u.domainName) : '';
        const areaValue = a.cr18c_area == null ? null : Number(a.cr18c_area);
        const region: AgentRow['region'] = areaValue == null ? '' : areaValue === AREA_KSA ? 'KSA' : 'EG';
        return {
          id: a.cr301_agentid,
          name: a.cr301_newcolumn || '—',
          userId,
          user: login || userFull || '—',
          userFull,
          title: u?.title || '',
          region,
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

  // App-wide Region toggle (Layout header) — unset region on an agent counts
  // as Egypt, same convention used everywhere else this filter appears.
  const regionAgents = useMemo(
    () => (!region ? agents : agents.filter((a) => (region === 'KSA' ? a.region === 'KSA' : a.region !== 'KSA'))),
    [agents, region]
  );

  const view = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) return regionAgents;
    return regionAgents.filter((a) => `${a.name} ${a.user} ${a.userFull} ${a.title}`.toLowerCase().includes(q));
  }, [regionAgents, query]);

  useEffect(() => {
    setPage(1);
  }, [query, region]);

  const totalPages = Math.max(1, Math.ceil(view.length / pageSize));
  const clampedPage = Math.min(page, totalPages);
  const pageRows = view.slice((clampedPage - 1) * pageSize, (clampedPage - 1) * pageSize + pageSize);

  const kpis = useMemo(
    () => ({
      total: regionAgents.length,
      linked: regionAgents.filter((a) => a.userId).length,
    }),
    [regionAgents]
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
      setFormRegion(a?.region || '');
    } else {
      setEditId(null);
      setFormName('');
      setEmployeeQuery('');
      setFormUserId('');
      setFormRegion('');
    }
    setDrawerOpen(true);
  }

  function closeDrawer() {
    setDrawerOpen(false);
  }

  // Empty query -> no suggestions at all (don't dump the full roster on a bare
  // click) — the list only appears once the user starts typing.
  const employeeMatches = useMemo(() => {
    const q = employeeQuery.trim().toLowerCase();
    if (!q) return [];
    return employees.filter((e) => e.name.toLowerCase().includes(q)).slice(0, 30);
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

  /**
   * Bind the Agent User lookup via the generic Dataverse client's @odata.bind
   * form. Mixing an "@odata.bind" key into the same body as plain fields and
   * sending it through the generated typed service's create/update has been
   * observed (see Accounts.tsx's setClassificationViaClient) to make the
   * generated service silently drop the *whole* write instead of just
   * rejecting the bind key — so the lookup is set here, separately, through
   * the raw client, after the plain fields already landed. Returns true on
   * success (a successful call, not proof the value stuck — verify separately).
   */
  async function setAgentUserViaClient(id: string, userId: string): Promise<boolean> {
    try {
      const client = getClient(dataSourcesInfo);
      const res = await client.updateRecordAsync(AGENTS_TABLE, id, {
        [`${AGENT_NAV}@odata.bind`]: `/systemusers(${userId})`,
      });
      return !!res.success;
    } catch {
      return false;
    }
  }

  /** Re-read one agent's Agent User lookup to confirm a write actually persisted. */
  async function readAgentUser(id: string): Promise<string> {
    try {
      const res = await Cr301_agentsService.get(id, { select: [AGENT_VALUE_FIELD] });
      if (res.success && res.data) {
        return ((res.data as unknown as Record<string, unknown>)[AGENT_VALUE_FIELD] as string) || '';
      }
    } catch {
      /* ignore */
    }
    return '';
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
      // 1) Write the plain (non-lookup) fields through the generated service —
      //    see setAgentUserViaClient's comment for why the lookup is kept out
      //    of this call.
      const body: Partial<Omit<Cr301_agentsBase, 'cr301_agentid'>> = {
        cr301_newcolumn: name,
        cfm_iscollection: true,
        cr18c_area: formRegion === 'KSA' ? AREA_KSA : formRegion === 'EG' ? AREA_EGYPT : null,
      } as Partial<Omit<Cr301_agentsBase, 'cr301_agentid'>>;

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

      let recordId = targetId;
      const result = targetId
        ? await Cr301_agentsService.update(targetId, body)
        : await Cr301_agentsService.create(body as unknown as Omit<Cr301_agentsBase, 'cr301_agentid'>);
      if (!result.success) {
        throw new Error(operationErrorMessage(result.error, 'Save failed.'));
      }
      if (!recordId) {
        recordId = ((result.data as unknown as Record<string, unknown>)?.cr301_agentid as string) || null;
      }

      // 2) Set the Agent User lookup separately, then verify it actually persisted
      //    before reporting success — a 200/204 isn't proof the value stuck.
      if (recordId) {
        const ok = await setAgentUserViaClient(recordId, formUserId);
        const stored = ok ? await readAgentUser(recordId) : '';
        if (stored !== formUserId) {
          setDrawerOpen(false);
          showToast(
            editId
              ? 'Agent saved, but Agent User could not be set — check field permissions / a plugin'
              : 'Agent created, but Agent User could not be set — check field permissions / a plugin',
            5000
          );
          await load();
          return;
        }
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
    const header = ['Agent Name', 'Agent User', 'Title', 'Region'];
    const rows = view.map((a) => [a.name, a.user, a.title, a.region || '—']);
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
              <Icon className="fa-solid fa-download" /> Export
            </button>
            <button className="acc-btn acc-btn-primary" onClick={() => openDrawer()}>
              <Icon className="fa-solid fa-plus" /> Add Agent
            </button>
          </div>
        </div>

        <div className="kpi-grid">
          <div className="kpi-card">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-gold">
                <Icon className="fa-solid fa-user-tie" />
              </span>
              <span className="kpi-label">Total agents</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.total}</div>
            <div className="kpi-context">&nbsp;</div>
          </div>
          <div className="kpi-card kpi-ok">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-ok">
                <Icon className="fa-solid fa-link" />
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
              <Icon className="fa-solid fa-magnifying-glass" />
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
                <th>Region</th>
              </tr>
            </thead>
            <tbody>
              {loading ? (
                Array.from({ length: 6 }).map((_, i) => (
                  <tr key={i}>
                    {Array.from({ length: 4 }).map((__, j) => (
                      <td key={j}>
                        <div className="acc-skel"></div>
                      </td>
                    ))}
                  </tr>
                ))
              ) : loadError ? (
                <tr>
                  <td colSpan={4}>
                    <div className="acc-state">
                      <Icon className="fa-solid fa-triangle-exclamation" />
                      Could not load agents.
                      <br />
                      <small>{loadError}</small>
                    </div>
                  </td>
                </tr>
              ) : pageRows.length === 0 ? (
                <tr>
                  <td colSpan={4}>
                    <div className="acc-state">
                      <Icon className="fa-solid fa-inbox" />
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
                    <td className="agt-title-cell">{a.region ? (a.region === 'KSA' ? 'KSA' : 'Egypt') : <span className="agt-muted">—</span>}</td>
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
            <Icon className="fa-solid fa-xmark" />
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
                name="cn-employee-search-nofill"
                id="cn-employee-search-nofill"
                autoComplete="new-password"
                autoCorrect="off"
                autoCapitalize="off"
                spellCheck={false}
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
            {employeePickerOpen && (employeeQuery.trim() || employeeLoadError) && (
              <div className="acc-combo-list">
                {employeeLoadError ? (
                  <div className="acc-combo-empty">Could not load employees — {employeeLoadError}</div>
                ) : employeeMatches.length === 0 ? (
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
            <div className="acc-hint">
              {employeeLoadError
                ? 'Employee search is unavailable — type the name and pick the Agent User below manually.'
                : 'Sets the Agent Name and auto-selects the matching Agent User below when found.'}
            </div>
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
            <div className="acc-hint">
              {userLoadError ? `Could not load system users — ${userLoadError}` : 'Title and team are read from the linked system user.'}
            </div>
          </div>
          <div className="acc-fld">
            <label>Region</label>
            <select value={formRegion} onChange={(e) => setFormRegion(e.target.value as 'EG' | 'KSA' | '')}>
              <option value="">— Select region —</option>
              <option value="EG">Egypt</option>
              <option value="KSA">KSA</option>
            </select>
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
              'Save Agent'
            )}
          </button>
        </div>
      </aside>

      <div className={`acc-toast${toastOn ? ' on' : ''}`}>{toastMsg}</div>
    </>
  );
}
