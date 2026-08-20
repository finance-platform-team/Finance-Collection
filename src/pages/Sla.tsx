import { useEffect, useMemo, useState } from 'react';
import { Cfm_slasService } from '../generated/services/Cfm_slasService';
import { SystemusersService } from '../generated/services/SystemusersService';
import type { Cfm_slas, Cfm_slasBase } from '../generated/models/Cfm_slasModel';
import { operationErrorMessage } from '../lib/errors';
import { fetchAll } from '../lib/dataverse';
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
const TAG_PILLS = [
  { value: null as number | null, label: 'All' },
  { value: TAG.COLLECTION as number, label: 'Collection' },
  { value: TAG.AR as number, label: 'AR' },
  { value: TAG.AP as number, label: 'AP' },
];

const SLA_TYPE = { DATE: 766340000, HOURS: 766340001 } as const;
const SLA_TYPE_LABEL: Record<number, string> = { 766340000: 'Date', 766340001: 'Hours' };
const SLA_TYPE_OPTIONS = [
  { value: SLA_TYPE.HOURS, label: 'Hours' },
  { value: SLA_TYPE.DATE, label: 'Date' },
];

function formattedValue(record: unknown, attribute: string): string {
  const raw = (record as Record<string, unknown>)[`${attribute}@OData.Community.Display.V1.FormattedValue`];
  return typeof raw === 'string' ? raw : '';
}

function tagPillClass(v: number | null): string {
  if (v === TAG.AR) return 'sla-tag-ar';
  if (v === TAG.AP) return 'sla-tag-ap';
  return 'sla-tag-collection';
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
}

function normaliseSla(row: Cfm_slas): SlaRow {
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

export default function SlaPage() {
  const [rows, setRows] = useState<SlaRow[]>([]);
  const [users, setUsers] = useState<UserOption[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);

  const [query, setQuery] = useState('');
  const [tagFilter, setTagFilter] = useState<number | null>(TAG.COLLECTION);

  const [toastMsg, setToastMsg] = useState('');
  const [toastOn, setToastOn] = useState(false);

  const [drawerOpen, setDrawerOpen] = useState(false);
  const [mode, setMode] = useState<'create' | 'edit'>('create');
  const [editId, setEditId] = useState<string | null>(null);

  const [formAction, setFormAction] = useState('');
  const [formTag, setFormTag] = useState(String(TAG.COLLECTION));
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

  async function load() {
    setLoading(true);
    setLoadError(null);
    try {
      const [userOptions, slaRows] = await Promise.all([
        loadUsers(),
        fetchAll<Cfm_slas>(
          (skipToken) => Cfm_slasService.getAll({ select: SLA_SELECT, filter: 'statecode eq 0', orderBy: ['cfm_action asc'], maxPageSize: 5000, skipToken }),
          'Could not load SLA rules.'
        ),
      ]);
      setUsers(userOptions);
      setRows(slaRows.map(normaliseSla));
    } catch (e) {
      setLoadError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    load();
  }, []);

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return rows.filter((r) => {
      if (tagFilter != null && r.tag !== tagFilter) return false;
      if (q && !`${r.action} ${r.department1} ${r.department2} ${r.responsibleName}`.toLowerCase().includes(q)) return false;
      return true;
    });
  }, [rows, query, tagFilter]);

  const kpis = useMemo(() => {
    const total = rows.length;
    const collection = rows.filter((r) => r.tag === TAG.COLLECTION).length;
    const ar = rows.filter((r) => r.tag === TAG.AR).length;
    const ap = rows.filter((r) => r.tag === TAG.AP).length;
    return { total, collection, ar, ap };
  }, [rows]);

  function openCreateDrawer() {
    setMode('create');
    setEditId(null);
    setFormAction('');
    setFormTag(String(TAG.COLLECTION));
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
    setFormTag(String(row.tag ?? TAG.COLLECTION));
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

  const canSave = !saving && formAction.trim() !== '';

  async function saveSla() {
    if (saving || !canSave) return;
    setSaving(true);
    setFormError('');
    try {
      const body: Record<string, unknown> = {
        cfm_action: formAction.trim(),
        cfm_tagename: parseInt(formTag, 10),
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
            <div className="acc-sub">Escalation and deduction rules per department and action — Collection, AR and AP.</div>
          </div>
          <div className="acc-actions">
            <button className="acc-btn acc-btn-primary" onClick={openCreateDrawer}>
              <i className="fa-solid fa-plus"></i> Create SLA rule
            </button>
          </div>
        </div>

        <div className="kpi-grid">
          <div className="kpi-card">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-gold">
                <i className="fa-solid fa-file-contract"></i>
              </span>
              <span className="kpi-label">Total SLA rules</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.total}</div>
          </div>
          <div className="kpi-card kpi-gold">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-gold">
                <i className="fa-solid fa-hand-holding-dollar"></i>
              </span>
              <span className="kpi-label">Collection</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.collection}</div>
          </div>
          <div className="kpi-card kpi-info">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-info">
                <i className="fa-solid fa-scale-balanced"></i>
              </span>
              <span className="kpi-label">AR</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.ar}</div>
          </div>
          <div className="kpi-card kpi-warn">
            <div className="kpi-top">
              <span className="kpi-icon kpi-icon-warn">
                <i className="fa-solid fa-file-invoice-dollar"></i>
              </span>
              <span className="kpi-label">AP</span>
            </div>
            <div className="kpi-value">{loading ? '—' : kpis.ap}</div>
          </div>
        </div>

        <div className="tbl-shell">
          <div className="tbl-hdr">
            <div className="dm-filter-bar">
              <div className="acc-search dm-search">
                <i className="fa-solid fa-magnifying-glass"></i>
                <input placeholder="Search action, department, responsible…" value={query} onChange={(e) => setQuery(e.target.value)} />
              </div>
            </div>
            <div className="dm-count">
              <i className="fa-solid fa-table"></i> Showing <strong>{filtered.length}</strong> of <strong>{rows.length}</strong>
            </div>
          </div>

          <div className="dm-category-bar">
            <span className="dm-category-label">Tag</span>
            {TAG_PILLS.map((p) => (
              <button key={String(p.value)} className={`dm-cat-pill${tagFilter === p.value ? ' active' : ''}`} onClick={() => setTagFilter(p.value)}>
                {p.label}
              </button>
            ))}
          </div>

          <div className="dm-tablewrap">
            <table>
              <thead>
                <tr>
                  <th>Action</th>
                  <th>Tag</th>
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
                        Could not load SLA rules.
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
                        No SLA rules match your filters.
                      </div>
                    </td>
                  </tr>
                ) : (
                  filtered.map((r) => (
                    <tr key={r.id} className="dm-row" onClick={() => openEditDrawer(r)}>
                      <td className="dm-company">{r.action}</td>
                      <td>
                        <span className={`sla-pill ${tagPillClass(r.tag)}`}>{r.tagLabel}</span>
                      </td>
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
        </div>
      </div>

      <div className={`acc-scrim${drawerOpen ? ' on' : ''}`} onClick={closeDrawer}></div>
      <aside className={`acc-drawer${drawerOpen ? ' on' : ''}`}>
        <div className="acc-dhead">
          <h2>{mode === 'edit' ? 'Edit · SLA Rule' : 'New · SLA Rule'}</h2>
          <button className="acc-dclose" onClick={closeDrawer}>
            <i className="fa-solid fa-xmark"></i>
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

          <div className="dm-fld-row">
            <div className="acc-fld">
              <label>Tag</label>
              <select value={formTag} onChange={(e) => setFormTag(e.target.value)}>
                {TAG_PILLS.filter((p) => p.value != null).map((p) => (
                  <option key={p.value} value={p.value as number}>
                    {p.label}
                  </option>
                ))}
              </select>
            </div>
            <div className="acc-fld">
              <label>SLA type</label>
              <select value={formSlaType} onChange={(e) => setFormSlaType(e.target.value)}>
                {SLA_TYPE_OPTIONS.map((o) => (
                  <option key={o.value} value={o.value}>
                    {o.label}
                  </option>
                ))}
              </select>
            </div>
          </div>

          <div className="acc-fld">
            <label>{Number(formSlaType) === SLA_TYPE.DATE ? 'SLA date' : 'SLA (number of hours)'}</label>
            <input
              placeholder={Number(formSlaType) === SLA_TYPE.DATE ? 'e.g. 2026-09-01' : 'e.g. 24'}
              value={formSlaValue}
              onChange={(e) => setFormSlaValue(e.target.value)}
            />
          </div>

          <div className="acc-fld">
            <label>Department 1</label>
            <input placeholder="Department name" value={formDept1} onChange={(e) => setFormDept1(e.target.value)} />
          </div>
          <div className="acc-fld">
            <label>Manager for Department 1</label>
            <UserCombo key={`${editId || 'new'}-m1`} users={users} initialName={formManager1Name} onChange={setFormManager1} />
          </div>

          <div className="acc-fld">
            <label>Department 2</label>
            <input placeholder="Department name" value={formDept2} onChange={(e) => setFormDept2(e.target.value)} />
          </div>
          <div className="acc-fld">
            <label>Manager of Department 2</label>
            <UserCombo key={`${editId || 'new'}-m2`} users={users} initialName={formManager2Name} onChange={setFormManager2} />
          </div>

          <div className="acc-fld">
            <label>Responsible</label>
            <UserCombo key={`${editId || 'new'}-resp`} users={users} initialName={formResponsibleName} onChange={setFormResponsible} />
          </div>

          <div className="acc-fld">
            <label>Escalation rule</label>
            <textarea rows={3} placeholder="What happens if the SLA is missed…" value={formEscalation} onChange={(e) => setFormEscalation(e.target.value)} />
          </div>

          <div className="acc-fld">
            <label>Deduction</label>
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
