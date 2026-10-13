import { useEffect, useMemo, useState } from 'react';
import { Cfm_slasService } from '../generated/services/Cfm_slasService';
import { SystemusersService } from '../generated/services/SystemusersService';
import { Cr301_agentsService } from '../generated/services/Cr301_agentsService';
import type { Cfm_slas, Cfm_slasBase } from '../generated/models/Cfm_slasModel';
import { operationErrorMessage } from '../lib/errors';
import { fetchAll } from '../lib/dataverse';
import type { Region } from '../lib/useRegion';
import Icon from '../components/Icon';
import './Sla.css';

/**
 * SLA rules (cfm_sla) — escalation/deduction definitions per department and
 * action, tagged Collection / AR / AP. Manager for Department1, Manager of
 * Department2 and Responsible are all systemuser lookups (confirmed live —
 * this table has no dedicated Ref.html section to copy from).
 */

const SLA_SELECT = [
  'cfm_slaid',
  'cfm_action',
  'cfm_tagename',
  'cfm_slatype',
  'cfm_slanoofhoursordate',
  'cfm_department1',
  '_cfm_managerfordepartment1_value',
  'cfm_department2',
  '_cfm_managerofdepartment2_value',
  '_cfm_responsible_value',
  'cfm_escalationrule',
  'cfm_deduction',
  'createdon',
];

const TAG = { COLLECTION: 766340000, AR: 766340001, AP: 766340002 } as const;
const TAG_LABEL: Record<number, string> = { 766340000: 'Collection', 766340001: 'AR', 766340002: 'AP' };

const SLA_TYPE = { DATE: 766340000, HOURS: 766340001 } as const;
const SLA_TYPE_LABEL: Record<number, string> = { 766340000: 'Date', 766340001: 'Hours' };
const SLA_TYPE_OPTIONS = [
  { value: SLA_TYPE.HOURS, label: 'Hours' },
  { value: SLA_TYPE.DATE, label: 'Date' },
];

// cfm_slas carries no company/BU field of its own — region is derived from
// Responsible (a systemuser lookup) via that same user's cr301_agents record's
// cr18c_area, the same source of truth Agents.tsx / Daily Collection use.
const AREA_KSA = 989230004;

function formattedValue(record: unknown, attribute: string): string {
  const raw = (record as Record<string, unknown>)[`${attribute}@OData.Community.Display.V1.FormattedValue`];
  return typeof raw === 'string' ? raw : '';
}

interface UserOption {
  id: string;
  name: string;
}

interface SlaRow {
  id: string;
  action: string;
  tag: number | null;
  tagLabel: string;
  slaType: number | null;
  slaTypeLabel: string;
  slaValue: string;
  department1: string;
  manager1Id: string;
  manager1Name: string;
  department2: string;
  manager2Id: string;
  manager2Name: string;
  responsibleId: string;
  responsibleName: string;
  escalationRule: string;
  deduction: string;
  createdOn: string;
  region: 'EG' | 'KSA' | '';
}

function normaliseSla(row: Cfm_slas, regionByUserId: Map<string, 'EG' | 'KSA'>): SlaRow {
  const tag = row.cfm_tagename == null ? null : Number(row.cfm_tagename);
  const slaType = row.cfm_slatype == null ? null : Number(row.cfm_slatype);
  return {
    id: row.cfm_slaid,
    action: row.cfm_action || '—',
    tag,
    tagLabel: formattedValue(row, 'cfm_tagename') || (tag != null ? TAG_LABEL[tag] : '—'),
    slaType,
    slaTypeLabel: formattedValue(row, 'cfm_slatype') || (slaType != null ? SLA_TYPE_LABEL[slaType] : '—'),
    slaValue: row.cfm_slanoofhoursordate || '',
    department1: row.cfm_department1 || '',
    manager1Id: row._cfm_managerfordepartment1_value || '',
    manager1Name: formattedValue(row, '_cfm_managerfordepartment1_value'),
    department2: row.cfm_department2 || '',
    manager2Id: row._cfm_managerofdepartment2_value || '',
    manager2Name: formattedValue(row, '_cfm_managerofdepartment2_value'),
    responsibleId: row._cfm_responsible_value || '',
    responsibleName: formattedValue(row, '_cfm_responsible_value'),
    escalationRule: row.cfm_escalationrule || '',
    deduction: row.cfm_deduction || '',
    createdOn: row.createdon || '',
    region: (row._cfm_responsible_value && regionByUserId.get(row._cfm_responsible_value)) || '',
  };
}

/* Searchable person picker — shared shape for Manager1 / Manager2 / Responsible.
 * Give each instance a `key` tied to the record being edited so it remounts
 * (and re-seeds from initialName) whenever the drawer switches records. */
function UserCombo({ users, initialName, onChange }: { users: UserOption[]; initialName: string; onChange: (u: UserOption | null) => void }) {
  const [query, setQuery] = useState(initialName);
  const [open, setOpen] = useState(false);
  const matches = useMemo(() => {
    const q = query.trim().toLowerCase();
    const list = q ? users.filter((u) => u.name.toLowerCase().includes(q)) : users;
    return list.slice(0, 30);
  }, [users, query]);
  return (
    <div className="acc-combo">
      <input
        placeholder="Search people…"
        value={query}
        onChange={(e) => {
          setQuery(e.target.value);
          setOpen(true);
        }}
        onFocus={() => setOpen(true)}
      />
      {open && (
        <div className="acc-combo-list">
          <div
            className="acc-combo-item"
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              setQuery('');
              setOpen(false);
              onChange(null);
            }}
          >
            <small>— None —</small>
          </div>
          {matches.length === 0 ? (
            <div className="acc-combo-empty">No matches.</div>
          ) : (
            matches.map((u) => (
              <div
                key={u.id}
                className="acc-combo-item"
                onMouseDown={(e) => e.preventDefault()}
                onClick={() => {
                  setQuery(u.name);
                  setOpen(false);
                  onChange(u);
                }}
              >
                <div>{u.name}</div>
              </div>
            ))
          )}
        </div>
      )}
    </div>
  );
}

interface SlaPageProps {
  region: Region | null;
}

export default function SlaPage({ region }: SlaPageProps) {
  const [rows, setRows] = useState<SlaRow[]>([]);
  const [users, setUsers] = useState<UserOption[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [query, setQuery] = useState('');
  const [page, setPage] = useState(1);
  const pageSize = 25;

  const [toastMsg, setToastMsg] = useState('');
  const [toastOn, setToastOn] = useState(false);

  const [drawerOpen, setDrawerOpen] = useState(false);
  const [mode, setMode] = useState<'create' | 'edit'>('create');
  const [editId, setEditId] = useState<string | null>(null);

  const [formAction, setFormAction] = useState('');
  const [formSlaType, setFormSlaType] = useState(String(SLA_TYPE.HOURS));
  const [formSlaValue, setFormSlaValue] = useState('');
  const [formDept1, setFormDept1] = useState('');
  const [formManager1, setFormManager1] = useState<UserOption | null>(null);
  const [formManager1Name, setFormManager1Name] = useState('');
  const [formDept2, setFormDept2] = useState('');
  const [formManager2, setFormManager2] = useState<UserOption | null>(null);
  const [formManager2Name, setFormManager2Name] = useState('');
  const [formResponsible, setFormResponsible] = useState<UserOption | null>(null);
  const [formResponsibleName, setFormResponsibleName] = useState('');
  const [formEscalation, setFormEscalation] = useState('');
  const [formDeduction, setFormDeduction] = useState('');
  const [saving, setSaving] = useState(false);
  const [formError, setFormError] = useState('');

  function showToast(msg: string, durationMs = 3200) {
    setToastMsg(msg);
    setToastOn(true);
    setTimeout(() => setToastOn(false), durationMs);
  }

  async function loadUsers(): Promise<UserOption[]> {
    const res = await SystemusersService.getAll({
      select: ['systemuserid', 'fullname'],
      filter: 'isdisabled eq false and applicationid eq null',
      orderBy: ['fullname'],
      maxPageSize: 5000,
    });
    if (!res.success) return [];
    return (res.data || []).filter((u) => u.fullname).map((u) => ({ id: u.systemuserid, name: u.fullname! }));
  }

  async function loadRegionByUserId(): Promise<Map<string, 'EG' | 'KSA'>> {
    const rows = await fetchAll<{ cr301_agentid: string; cr18c_area?: number; _cr301_agent_value?: string }>(
      (skipToken) =>
        Cr301_agentsService.getAll({
          select: ['cr301_agentid', 'cr18c_area', '_cr301_agent_value'],
          filter: 'cfm_iscollection eq true',
          maxPageSize: 5000,
          skipToken,
        }),
      'Could not load agents.'
    );
    const map = new Map<string, 'EG' | 'KSA'>();
    rows.forEach((a) => {
      if (a._cr301_agent_value && a.cr18c_area != null) {
        map.set(a._cr301_agent_value, Number(a.cr18c_area) === AREA_KSA ? 'KSA' : 'EG');
      }
    });
    return map;
  }

  async function load() {
    setLoading(true);
    setLoadError(null);
    try {
      const [userOptions, regionMap, slaRows] = await Promise.all([
        loadUsers(),
        loadRegionByUserId(),
        fetchAll<Cfm_slas>(
          (skipToken) => Cfm_slasService.getAll({ select: SLA_SELECT, filter: 'statecode eq 0', orderBy: ['cfm_action asc'], maxPageSize: 5000, skipToken }),
          'Could not load SLA rules.'
        ),
      ]);
      setUsers(userOptions);
      setRows(slaRows.map((r) => normaliseSla(r, regionMap)));
    } catch (e) {
      setLoadError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    load();
  }, []);

  // This page is Collection-only now (the Tag picker was removed from the create/edit
  // form and every rule saved through it is tagged Collection) — the list matches that,
  // filtering to Collection-tagged rows rather than showing legacy AR/AP/untagged ones.
  // Folded into regionRows (not just `filtered`) so the "Showing X of Y" count
  // below reflects Collection-only, region-scoped rows as its denominator —
  // not inflated by AR/AP/untagged rows the page never shows anyway.
  // App-wide Region toggle (Layout header) — unset region on a rule counts as
  // Egypt, same convention used everywhere else this filter appears.
  const regionRows = useMemo(
    () =>
      rows.filter((r) => {
        if (r.tag !== TAG.COLLECTION) return false;
        if (region && (region === 'KSA' ? r.region !== 'KSA' : r.region === 'KSA')) return false;
        return true;
      }),
    [rows, region]
  );

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return regionRows.filter((r) => {
      if (q && !`${r.action} ${r.department1} ${r.department2} ${r.responsibleName}`.toLowerCase().includes(q)) return false;
      return true;
    });
  }, [regionRows, query]);

  useEffect(() => {
    setPage(1);
  }, [query, region]);

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

  function openCreateDrawer() {
    setMode('create');
    setEditId(null);
    setFormAction('');
    setFormSlaType(String(SLA_TYPE.HOURS));
    setFormSlaValue('');
    setFormDept1('');
    setFormManager1(null);
    setFormManager1Name('');
    setFormDept2('');
    setFormManager2(null);
    setFormManager2Name('');
    setFormResponsible(null);
    setFormResponsibleName('');
    setFormEscalation('');
    setFormDeduction('');
    setFormError('');
    setDrawerOpen(true);
  }

  function openEditDrawer(row: SlaRow) {
    setMode('edit');
    setEditId(row.id);
    setFormAction(row.action === '—' ? '' : row.action);
    setFormSlaType(String(row.slaType ?? SLA_TYPE.HOURS));
    setFormSlaValue(row.slaValue);
    setFormDept1(row.department1);
    setFormManager1(row.manager1Id ? { id: row.manager1Id, name: row.manager1Name } : null);
    setFormManager1Name(row.manager1Name);
    setFormDept2(row.department2);
    setFormManager2(row.manager2Id ? { id: row.manager2Id, name: row.manager2Name } : null);
    setFormManager2Name(row.manager2Name);
    setFormResponsible(row.responsibleId ? { id: row.responsibleId, name: row.responsibleName } : null);
    setFormResponsibleName(row.responsibleName);
    setFormEscalation(row.escalationRule);
    setFormDeduction(row.deduction);
    setFormError('');
    setDrawerOpen(true);
  }

  function closeDrawer() {
    setDrawerOpen(false);
  }

  const canSave = !saving;

  async function saveSla() {
    if (saving) return;
    if (!formAction.trim()) {
      setFormError('Action is required.');
      return;
    }
    if (!formSlaValue.trim()) {
      setFormError(Number(formSlaType) === SLA_TYPE.DATE ? 'SLA date is required.' : 'SLA (number of hours) is required.');
      return;
    }
    if (!formDept1.trim()) {
      setFormError('Department 1 is required.');
      return;
    }
    if (!formManager1) {
      setFormError('Manager for Department 1 is required.');
      return;
    }
    if (!formDept2.trim()) {
      setFormError('Department 2 is required.');
      return;
    }
    if (!formManager2) {
      setFormError('Manager of Department 2 is required.');
      return;
    }
    if (!formResponsible) {
      setFormError('Responsible is required.');
      return;
    }
    if (!formEscalation.trim()) {
      setFormError('Escalation rule is required.');
      return;
    }
    if (!formDeduction.trim()) {
      setFormError('Deduction is required.');
      return;
    }
    setSaving(true);
    setFormError('');
    try {
      const body: Record<string, unknown> = {
        cfm_action: formAction.trim(),
        cfm_tagename: TAG.COLLECTION,
        cfm_slatype: parseInt(formSlaType, 10),
        cfm_slanoofhoursordate: formSlaValue.trim() || null,
        cfm_department1: formDept1.trim() || null,
        cfm_department2: formDept2.trim() || null,
        cfm_escalationrule: formEscalation.trim() || null,
        cfm_deduction: formDeduction.trim() || null,
      };
      body['cfm_ManagerforDepartment1@odata.bind'] = formManager1 ? `/systemusers(${formManager1.id})` : null;
      body['cfm_ManagerofDepartment2@odata.bind'] = formManager2 ? `/systemusers(${formManager2.id})` : null;
      body['cfm_Responsible@odata.bind'] = formResponsible ? `/systemusers(${formResponsible.id})` : null;

      if (mode === 'edit' && editId) {
        const result = await Cfm_slasService.update(editId, body as Partial<Omit<Cfm_slasBase, 'cfm_slaid'>>);
        if (!result.success) throw new Error(operationErrorMessage(result.error, 'Save failed.'));
        showToast('SLA rule updated.');
      } else {
        const result = await Cfm_slasService.create(body as unknown as Omit<Cfm_slasBase, 'cfm_slaid'>);
        if (!result.success) throw new Error(operationErrorMessage(result.error, 'Save failed.'));
        showToast('SLA rule created.');
      }
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
            <h1 className="acc-title">SLA</h1>
          </div>
          <div className="acc-actions">
            <button className="acc-btn acc-btn-primary" onClick={openCreateDrawer}>
              <Icon className="fa-solid fa-plus" /> Create SLA rule
            </button>
          </div>
        </div>

        <div className="tbl-shell">
          <div className="tbl-hdr">
            <div className="dm-filter-bar">
              <div className="acc-search dm-search">
                <Icon className="fa-solid fa-magnifying-glass" />
                <input placeholder="Search action, department, responsible…" value={query} onChange={(e) => setQuery(e.target.value)} />
              </div>
            </div>
            <div className="dm-count">
              <Icon className="fa-solid fa-table" /> Showing <strong>{filtered.length}</strong> of <strong>{regionRows.length}</strong>
            </div>
          </div>

          <div className="dm-tablewrap">
            <table>
              <thead>
                <tr>
                  <th>Action</th>
                  <th>Type</th>
                  <th>SLA</th>
                  <th>Department 1</th>
                  <th>Department 2</th>
                  <th>Responsible</th>
                  <th>Escalation rule</th>
                  <th>Deduction</th>
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
                        <Icon className="fa-solid fa-triangle-exclamation" />
                        Could not load SLA rules.
                        <br />
                        <small>{loadError}</small>
                      </div>
                    </td>
                  </tr>
                ) : filtered.length === 0 ? (
                  <tr>
                    <td colSpan={8}>
                      <div className="acc-state">
                        <Icon className="fa-solid fa-inbox" />
                        No SLA rules match your filters.
                      </div>
                    </td>
                  </tr>
                ) : (
                  pageRows.map((r) => (
                    <tr key={r.id} className="dm-row" onClick={() => openEditDrawer(r)}>
                      <td className="dm-company">{r.action}</td>
                      <td>{r.slaTypeLabel}</td>
                      <td>{r.slaValue || '—'}</td>
                      <td>
                        {r.department1 || '—'}
                        {r.manager1Name && <div className="sla-sub">{r.manager1Name}</div>}
                      </td>
                      <td>
                        {r.department2 || '—'}
                        {r.manager2Name && <div className="sla-sub">{r.manager2Name}</div>}
                      </td>
                      <td>{r.responsibleName || '—'}</td>
                      <td className="sla-wrap">{r.escalationRule || '—'}</td>
                      <td className="sla-wrap">{r.deduction || '—'}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
          {!loading && !loadError && filtered.length > 0 && (
            <div className="acc-foot">
              <span>
                {filtered.length} SLA rule{filtered.length !== 1 ? 's' : ''}
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
            </div>
          )}
        </div>
      </div>

      <div className={`acc-scrim${drawerOpen ? ' on' : ''}`} onClick={closeDrawer}></div>
      <aside className={`acc-drawer${drawerOpen ? ' on' : ''}`}>
        <div className="acc-dhead">
          <h2>{mode === 'edit' ? 'Edit · SLA Rule' : 'New · SLA Rule'}</h2>
          <button className="acc-dclose" onClick={closeDrawer}>
            <Icon className="fa-solid fa-xmark" />
          </button>
        </div>
        <div className="acc-dbody">
          <div className={`acc-err${formError ? ' on' : ''}`}>{formError}</div>

          <div className="acc-fld">
            <label>
              Action <span className="dm-req">*</span>
            </label>
            <input placeholder="e.g. Follow up on rejected collection" value={formAction} onChange={(e) => setFormAction(e.target.value)} />
          </div>

          <div className="acc-fld">
            <label>
              SLA type <span className="dm-req">*</span>
            </label>
            <select value={formSlaType} onChange={(e) => setFormSlaType(e.target.value)}>
              {SLA_TYPE_OPTIONS.map((o) => (
                <option key={o.value} value={o.value}>
                  {o.label}
                </option>
              ))}
            </select>
          </div>

          <div className="acc-fld">
            <label>
              {Number(formSlaType) === SLA_TYPE.DATE ? 'SLA date' : 'SLA (number of hours)'} <span className="dm-req">*</span>
            </label>
            <input
              placeholder={Number(formSlaType) === SLA_TYPE.DATE ? 'e.g. 2026-09-01' : 'e.g. 24'}
              value={formSlaValue}
              onChange={(e) => setFormSlaValue(e.target.value)}
            />
          </div>

          <div className="acc-fld">
            <label>
              Department 1 <span className="dm-req">*</span>
            </label>
            <input placeholder="Department name" value={formDept1} onChange={(e) => setFormDept1(e.target.value)} />
          </div>
          <div className="acc-fld">
            <label>
              Manager for Department 1 <span className="dm-req">*</span>
            </label>
            <UserCombo key={`${editId || 'new'}-m1`} users={users} initialName={formManager1Name} onChange={setFormManager1} />
          </div>

          <div className="acc-fld">
            <label>
              Department 2 <span className="dm-req">*</span>
            </label>
            <input placeholder="Department name" value={formDept2} onChange={(e) => setFormDept2(e.target.value)} />
          </div>
          <div className="acc-fld">
            <label>
              Manager of Department 2 <span className="dm-req">*</span>
            </label>
            <UserCombo key={`${editId || 'new'}-m2`} users={users} initialName={formManager2Name} onChange={setFormManager2} />
          </div>

          <div className="acc-fld">
            <label>
              Responsible <span className="dm-req">*</span>
            </label>
            <UserCombo key={`${editId || 'new'}-resp`} users={users} initialName={formResponsibleName} onChange={setFormResponsible} />
          </div>

          <div className="acc-fld">
            <label>
              Escalation rule <span className="dm-req">*</span>
            </label>
            <textarea rows={3} placeholder="What happens if the SLA is missed…" value={formEscalation} onChange={(e) => setFormEscalation(e.target.value)} />
          </div>

          <div className="acc-fld">
            <label>
              Deduction <span className="dm-req">*</span>
            </label>
            <textarea rows={3} placeholder="Penalty / deduction applied…" value={formDeduction} onChange={(e) => setFormDeduction(e.target.value)} />
          </div>
        </div>
        <div className="acc-dfoot">
          <div style={{ display: 'flex', gap: 8, marginLeft: 'auto' }}>
            <button className="acc-btn acc-btn-ghost" onClick={closeDrawer}>
              Cancel
            </button>
            <button className="acc-btn acc-btn-primary" disabled={!canSave} onClick={saveSla}>
              {saving ? 'Saving…' : mode === 'edit' ? 'Save changes' : 'Create SLA rule'}
            </button>
          </div>
        </div>
      </aside>

      <div className={`acc-toast${toastOn ? ' on' : ''}`}>{toastMsg}</div>
    </>
  );
}
