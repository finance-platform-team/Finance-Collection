import Icon from '../components/Icon';
import { useEffect, useMemo, useRef, useState, type KeyboardEvent, type ReactNode } from 'react';
import { getClient } from '@microsoft/power-apps/data';
import { dataSourcesInfo } from '../../.power/schemas/appschemas/dataSourcesInfo';
import { Cr301_agentsService } from '../generated/services/Cr301_agentsService';
import { SystemusersService } from '../generated/services/SystemusersService';
import { Hr_employeesService } from '../generated/services/Hr_employeesService';
import { Cr301_salesteamsService } from '../generated/services/Cr301_salesteamsService';
import type { Cr301_agentsBase } from '../generated/models/Cr301_agentsModel';
import { operationErrorMessage } from '../lib/errors';
import { fetchAll } from '../lib/dataverse';
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
/** cr301_salesteam — a real lookup on cr301_agents (title "Sales Team") to the
 *  custom cr301_salesteams table (not the standard Dataverse `team` entity —
 *  confirmed live: the OData property is _cr301_salesteam_value, and there is
 *  no cr301_salesteamname shadow field despite the generated Cr301_agentsModel
 *  claiming one exists, so the team's display name is resolved via teamMap
 *  below instead of trusting that field). Its primary name column is
 *  cr301_newcolumn (see Cr301_salesteamsModel.ts), same quirky convention as
 *  cr301_agents' own name field.
 */
const TEAM_NAV = 'cr301_SalesTeam';
const TEAM_VALUE_FIELD = '_cr301_salesteam_value';
const TEAMS_ENTITY_SET = 'cr301_salesteams';
const USERS_ENTITY_SET = 'systemusers';
/** cfm_employee — a real lookup on cr301_agents to hr_employees (schema name
 *  "cfm_Employee", confirmed in agents.Schema.json). Never actually bound on
 *  save before this — VisitPlan.tsx/Visits.tsx had to fall back to fuzzy
 *  name-matching against hr_employees because this lookup was always empty.
 *  Binding it here lets those pages join on a real id instead. */
const EMPLOYEE_NAV = 'cfm_Employee';
const EMPLOYEE_VALUE_FIELD = '_cfm_employee_value';
const EMPLOYEES_ENTITY_SET = 'hr_employees';

interface UserOption {
  id: string;
  displayName: string;
  title: string;
  domainName: string;
  email: string;
  /** Application/service account (applicationid set) — kept in the list but
   *  ranked below real people, since they're the "# Something" noise that used
   *  to fill the top of the raw alphabetical dropdown. */
  isApp: boolean;
}

interface EmployeeOption {
  id: string;
  name: string;
  hrCode: string;
}

interface TeamOption {
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
  teamId: string;
  teamName: string;
  /** cr301_agents.cfm_employee (_cfm_employee_value) — empty for an agent
   *  saved before this field started being bound (see handleSave). */
  employeeId: string;
}

function loginOf(domainName: string): string {
  return (domainName || '').split('@')[0] || '';
}

/* -------------------------------------------------------------- search picker */

/** Rendered rows per keystroke. Rosters here run into the thousands; anything
 *  past this is unreachable by scrolling anyway — narrowing the query is the
 *  intended path, and the list footer says so. */
const SEARCH_RENDER_CAP = 60;

/** Tokenised query: "ahmed abd" must match ALL tokens (in any field, in any
 *  order) so multi-part Arabic names can be narrowed from either end. */
function queryTokens(q: string): string[] {
  return q.trim().toLowerCase().split(/\s+/).filter(Boolean);
}

/** Does any word in `hay` start with `token`? (i.e. "abd" hits "Ahmed AbdElWahed") */
function hasWordPrefix(hay: string, token: string): boolean {
  let i = hay.indexOf(token);
  while (i !== -1) {
    if (i === 0 || !/[a-z0-9]/i.test(hay[i - 1])) return true;
    i = hay.indexOf(token, i + 1);
  }
  return false;
}

/** A field to search, with its own weight — the primary name field should
 *  outweigh a secondary one like a login or job title. */
interface SearchField {
  text: string;
  weight: number;
}

/**
 * Relevance score for one option's fields against an already-lowercased token
 * set, or -1 the moment any token misses every field entirely (so it's
 * dropped from the results, not just sunk to the bottom). Each field scores a
 * token at its own weight for a prefix match, 70% for a match starting a
 * later word, 40% for a bare substring — so a full-name prefix always beats a
 * mid-name word, which beats a login/email hit, which beats a job-title hit,
 * as long as fields are passed in that priority order.
 */
function scoreOption(fields: SearchField[], tokens: string[]): number {
  let total = 0;
  for (const t of tokens) {
    let best = 0;
    for (const f of fields) {
      const v = f.text.toLowerCase();
      if (!v) continue;
      let s = 0;
      if (v.startsWith(t)) s = f.weight;
      else if (hasWordPrefix(v, t)) s = f.weight * 0.7;
      else if (v.includes(t)) s = f.weight * 0.4;
      if (s > best) best = s;
    }
    if (best === 0) return -1;
    total += best;
  }
  return total;
}

/** Splits `text` so every token match can be wrapped in <mark>. */
function highlight(text: string, tokens: string[]) {
  if (tokens.length === 0) return text;
  const lower = text.toLowerCase();
  const hits: [number, number][] = [];
  for (const t of tokens) {
    let i = lower.indexOf(t);
    while (i !== -1) {
      hits.push([i, i + t.length]);
      i = lower.indexOf(t, i + t.length);
    }
  }
  if (hits.length === 0) return text;
  hits.sort((a, b) => a[0] - b[0]);
  const merged: [number, number][] = [];
  for (const h of hits) {
    const last = merged[merged.length - 1];
    if (last && h[0] <= last[1]) last[1] = Math.max(last[1], h[1]);
    else merged.push([h[0], h[1]]);
  }
  const out: ReactNode[] = [];
  let cursor = 0;
  merged.forEach(([s, e], i) => {
    if (s > cursor) out.push(text.slice(cursor, s));
    out.push(
      <mark key={i} className="agt-hl">
        {text.slice(s, e)}
      </mark>
    );
    cursor = e;
  });
  if (cursor < text.length) out.push(text.slice(cursor));
  return out;
}

interface SearchPickerProps<T extends { id: string }> {
  options: T[];
  valueId: string;
  disabled?: boolean;
  disabledHint?: string;
  placeholder: string;
  /** Shown instead of the input's normal empty/placeholder state when
   *  `valueId` is set but doesn't match any option — e.g. Employee in edit
   *  mode, when openDrawer couldn't resolve a real employee id (no saved
   *  cfm_employee link and no name match) and left it empty while the Agent
   *  Name field still shows the on-file name. Ignored once a real match
   *  exists. */
  fallbackLabel?: string;
  loadErrorText?: string;
  listboxId: string;
  getPrimary: (t: T) => string;
  getSecondary?: (t: T) => string;
  getTag?: (t: T) => string | null | undefined;
  getFields: (t: T) => SearchField[];
  /** Extra amount subtracted only from ranking order (never from whether an
   *  option matches at all) — e.g. service accounts sinking below real people. */
  getPenalty?: (t: T) => number;
  onSelect: (option: T | null) => void;
}

/**
 * Searchable, keyboard-driven combobox shared by every lookup picker on this
 * page (Employee, Agent User, Sales Team). Searches every declared field,
 * ranks by relevance, highlights the match, and lists every option — no
 * arbitrary row cap on the underlying data, only on how many render per
 * keystroke (narrowing the query reaches the rest).
 */
function SearchPicker<T extends { id: string }>({
  options,
  valueId,
  disabled,
  disabledHint,
  placeholder,
  fallbackLabel,
  loadErrorText,
  listboxId,
  getPrimary,
  getSecondary,
  getTag,
  getFields,
  getPenalty,
  onSelect,
}: SearchPickerProps<T>) {
  const [query, setQuery] = useState('');
  const [open, setOpen] = useState(false);
  const [hi, setHi] = useState(0);
  const inputRef = useRef<HTMLInputElement | null>(null);
  const listRef = useRef<HTMLDivElement | null>(null);
  const blurTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  const selected = useMemo(() => options.find((o) => o.id === valueId) || null, [options, valueId]);
  const tokens = useMemo(() => queryTokens(query), [query]);

  const ranked = useMemo(() => {
    if (tokens.length === 0) {
      return [...options].sort((a, b) => (getPenalty?.(a) || 0) - (getPenalty?.(b) || 0) || getPrimary(a).localeCompare(getPrimary(b)));
    }
    return options
      .map((o) => {
        const fields = getFields(o);
        const raw = scoreOption(fields, tokens);
        if (raw < 0) return null;
        const primary = (fields[0]?.text || getPrimary(o)).toLowerCase();
        const wholeBonus = tokens.length > 1 && primary.startsWith(tokens.join(' ')) ? 60 : 0;
        const lengthPenalty = Math.min(primary.length, 60) / 60;
        return { o, s: raw + wholeBonus - lengthPenalty - (getPenalty?.(o) || 0) };
      })
      .filter((x): x is { o: T; s: number } => x !== null)
      .sort((a, b) => b.s - a.s || getPrimary(a.o).localeCompare(getPrimary(b.o)))
      .map((x) => x.o);
  }, [options, tokens, getFields, getPenalty, getPrimary]);

  const shown = ranked.slice(0, SEARCH_RENDER_CAP);

  useEffect(() => {
    setHi(0);
  }, [query, open]);

  // Keep the keyboard cursor visible while arrowing through a long list.
  useEffect(() => {
    if (!open) return;
    listRef.current?.querySelector<HTMLElement>(`[data-idx="${hi}"]`)?.scrollIntoView({ block: 'nearest' });
  }, [hi, open]);

  useEffect(() => () => clearTimeout(blurTimer.current), []);

  function commit(o: T) {
    onSelect(o);
    setQuery('');
    setOpen(false);
    inputRef.current?.blur();
  }

  function onKeyDown(e: KeyboardEvent<HTMLInputElement>) {
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      if (!open) {
        setOpen(true);
        return;
      }
      if (shown.length === 0) return;
      setHi((h) => (e.key === 'ArrowDown' ? (h + 1) % shown.length : (h - 1 + shown.length) % shown.length));
    } else if (e.key === 'Enter') {
      if (open && shown[hi]) {
        e.preventDefault();
        commit(shown[hi]);
      }
    } else if (e.key === 'Escape') {
      if (open) {
        e.preventDefault();
        setQuery('');
        setOpen(false);
      }
    } else if (e.key === 'Tab') {
      setOpen(false);
    }
  }

  const closedLabel = selected ? getPrimary(selected) : valueId && fallbackLabel ? fallbackLabel : '';

  return (
    <div className="acc-combo agt-picker">
      <div className={`agt-picker-field${disabled ? ' is-disabled' : ''}${open ? ' is-open' : ''}`}>
        <Icon className="fa-solid fa-magnifying-glass" />
        <input
          ref={inputRef}
          // Selection shows as the field's text once the picker is closed;
          // focusing starts a fresh search rather than making the user erase it.
          value={open ? query : closedLabel}
          placeholder={disabled ? disabledHint || '' : closedLabel || placeholder}
          disabled={disabled}
          title={disabled ? disabledHint : undefined}
          role="combobox"
          aria-expanded={open}
          aria-controls={listboxId}
          aria-activedescendant={open && shown[hi] ? `${listboxId}-opt-${shown[hi].id}` : undefined}
          autoComplete="new-password"
          autoCorrect="off"
          autoCapitalize="off"
          spellCheck={false}
          onChange={(e) => {
            setQuery(e.target.value);
            setOpen(true);
          }}
          onFocus={() => {
            clearTimeout(blurTimer.current);
            setQuery('');
            setOpen(true);
          }}
          onBlur={() => {
            blurTimer.current = setTimeout(() => setOpen(false), 150);
          }}
          onKeyDown={onKeyDown}
        />
        {selected && !disabled && (
          <button
            type="button"
            className="agt-picker-clear"
            title="Clear selection"
            aria-label="Clear selection"
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => {
              // Keep focus + the list open so clearing rolls straight into a
              // new search instead of forcing a re-click into the field.
              onSelect(null);
              setQuery('');
              setOpen(true);
              inputRef.current?.focus();
            }}
          >
            <Icon className="fa-solid fa-xmark" />
          </button>
        )}
      </div>

      {open && !disabled && (
        <div className="acc-combo-list agt-picker-list" id={listboxId} role="listbox" ref={listRef}>
          {loadErrorText ? (
            <div className="acc-combo-empty">{loadErrorText}</div>
          ) : shown.length === 0 ? (
            <div className="acc-combo-empty">{query.trim() ? `No match for “${query.trim()}”.` : 'No options available.'}</div>
          ) : (
            shown.map((o, i) => {
              const tag = getTag?.(o);
              const secondary = getSecondary?.(o);
              return (
                <div
                  key={o.id}
                  id={`${listboxId}-opt-${o.id}`}
                  data-idx={i}
                  role="option"
                  aria-selected={o.id === valueId}
                  className={`acc-combo-item agt-picker-item${i === hi ? ' hi' : ''}${o.id === valueId ? ' sel' : ''}`}
                  onMouseEnter={() => setHi(i)}
                  onMouseDown={(e) => e.preventDefault()}
                  onClick={() => commit(o)}
                >
                  <div className="agt-picker-row">
                    <span className="agt-picker-name">{highlight(getPrimary(o), tokens)}</span>
                    {tag && <span className="agt-picker-tag">{tag}</span>}
                    {o.id === valueId && <Icon className="fa-solid fa-check agt-picker-check" />}
                  </div>
                  {secondary != null && <small>{secondary ? highlight(secondary, tokens) : '—'}</small>}
                </div>
              );
            })
          )}
          {!loadErrorText && ranked.length > shown.length && (
            <div className="acc-combo-empty agt-picker-more">
              Showing {shown.length} of {ranked.length} — keep typing to narrow.
            </div>
          )}
        </div>
      )}
    </div>
  );
}

interface AgentsPageProps {
  region: Region | null;
}

export default function AgentsPage({ region }: AgentsPageProps) {
  const [agents, setAgents] = useState<AgentRow[]>([]);
  const [users, setUsers] = useState<UserOption[]>([]);
  const [employees, setEmployees] = useState<EmployeeOption[]>([]);
  const [teams, setTeams] = useState<TeamOption[]>([]);
  const [loading, setLoading] = useState(true);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [userLoadError, setUserLoadError] = useState<string | null>(null);
  const [employeeLoadError, setEmployeeLoadError] = useState<string | null>(null);
  const [teamLoadError, setTeamLoadError] = useState<string | null>(null);

  const [query, setQuery] = useState('');
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(15);

  const [drawerOpen, setDrawerOpen] = useState(false);
  const [editId, setEditId] = useState<string | null>(null);
  const [formName, setFormName] = useState('');
  // Set only by actually picking a suggestion (pickEmployee) — any manual edit
  // to the search text clears it again, so Save can require a real employee
  // was selected instead of accepting whatever text is currently typed.
  const [formEmployeeId, setFormEmployeeId] = useState('');
  const [formUserId, setFormUserId] = useState('');
  const [formRegion, setFormRegion] = useState<'EG' | 'KSA' | ''>('');
  const [formError, setFormError] = useState('');
  const [saving, setSaving] = useState(false);
  const [linkingEmployees, setLinkingEmployees] = useState(false);

  const [toastMsg, setToastMsg] = useState('');
  const [toastOn, setToastOn] = useState(false);
  const toastTimer = useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  function showToast(msg: string, durationMs = 2200) {
    setToastMsg(msg);
    setToastOn(true);
    clearTimeout(toastTimer.current);
    toastTimer.current = setTimeout(() => setToastOn(false), durationMs);
  }

  // Every agent managed on this page is a collection agent (cfm_iscollection
  // eq true) — Sales Team is no longer a per-agent choice, it's always
  // "Collection Team". Prefer an exact "Collection Team" name match, then a
  // bare "Collection", then fall back to any team whose name contains it —
  // tolerant of the record being named slightly differently than expected.
  const collectionTeam = useMemo(() => {
    const byExact = (label: string) => teams.find((t) => t.name.trim().toLowerCase() === label);
    return byExact('collection team') || byExact('collection') || teams.find((t) => t.name.toLowerCase().includes('collection')) || null;
  }, [teams]);

  // Only reflects a real pick from `employees` — see handleSave's comment on
  // why the edit-sentinel value of formEmployeeId must not match here.
  const selectedEmployee = useMemo(() => employees.find((e) => e.id === formEmployeeId) || null, [employees, formEmployeeId]);

  /**
   * These fetches used to swallow a failed result down to `[]`, which is
   * indistinguishable from "this org genuinely has zero active records" — the
   * Add Agent drawer would just show an empty, unexplained picker either way.
   * Throwing here instead lets load() capture and surface the real reason
   * (permissions, a renamed field, network) so it's actually diagnosable.
   */
  /**
   * Every active system user, not the first 500 — the old `top: 500` silently
   * truncated the roster, so anyone alphabetically past the cap simply couldn't
   * be picked. Drains all skip-token pages instead (fetchAll throws on a failed
   * page, which is the contract load() already expects here).
   */
  async function fetchUsers(): Promise<UserOption[]> {
    const rows = await fetchAll<{
      systemuserid: string;
      fullname?: string;
      domainname?: string;
      title?: string;
      internalemailaddress?: string;
      applicationid?: string;
    }>(
      (skipToken) =>
        SystemusersService.getAll({
          select: ['systemuserid', 'fullname', 'domainname', 'title', 'internalemailaddress', 'applicationid'],
          filter: 'isdisabled eq false',
          orderBy: ['fullname'],
          maxPageSize: 5000,
          skipToken,
        }),
      'Could not load system users.'
    );
    return rows.map((u) => ({
      id: u.systemuserid,
      displayName: (u.fullname || u.domainname || '(user)').trim(),
      title: (u.title || '').trim(),
      domainName: u.domainname || '',
      email: (u.internalemailaddress || '').trim(),
      isApp: !!u.applicationid,
    }));
  }

  /** Every active employee, fully drained the same way — see fetchUsers' comment. */
  async function fetchEmployees(): Promise<EmployeeOption[]> {
    const rows = await fetchAll<{ hr_employeeid: string; hr_fullname?: string; hr_currenthrcode?: string }>(
      (skipToken) =>
        Hr_employeesService.getAll({
          select: ['hr_employeeid', 'hr_fullname', 'hr_currenthrcode'],
          filter: 'statecode eq 0',
          orderBy: ['hr_fullname'],
          maxPageSize: 5000,
          skipToken,
        }),
      'Could not load employees.'
    );
    // Records with a blank/whitespace-only hr_fullname render as invisible,
    // unpickable rows — drop them; they're not a usable pick anyway.
    return rows
      .map((e) => ({ id: e.hr_employeeid, name: (e.hr_fullname || '').trim(), hrCode: (e.hr_currenthrcode || '').trim() }))
      .filter((e) => e.name);
  }

  async function fetchTeams(): Promise<TeamOption[]> {
    const rows = await fetchAll<{ cr301_salesteamid: string; cr301_newcolumn?: string }>(
      (skipToken) =>
        Cr301_salesteamsService.getAll({
          select: ['cr301_salesteamid', 'cr301_newcolumn'],
          filter: 'statecode eq 0',
          orderBy: ['cr301_newcolumn'],
          maxPageSize: 5000,
          skipToken,
        }),
      'Could not load teams.'
    );
    return rows.map((t) => ({ id: t.cr301_salesteamid, name: (t.cr301_newcolumn || '').trim() })).filter((t) => t.name);
  }

  async function load() {
    setLoading(true);
    setLoadError(null);
    setUserLoadError(null);
    setEmployeeLoadError(null);
    setTeamLoadError(null);
    try {
      // Users/employees/teams are independent, drawer-only concerns — a failure
      // in any one (permissions, a renamed field) shouldn't take down the
      // agents grid itself, so each is caught on its own instead of failing
      // the whole Promise.all.
      const [userSettled, employeeSettled, teamSettled, result] = await Promise.all([
        fetchUsers()
          .then((data) => ({ data, error: null }))
          .catch((e) => ({ data: [] as UserOption[], error: e instanceof Error ? e.message : String(e) })),
        fetchEmployees()
          .then((data) => ({ data, error: null }))
          .catch((e) => ({ data: [] as EmployeeOption[], error: e instanceof Error ? e.message : String(e) })),
        fetchTeams()
          .then((data) => ({ data, error: null }))
          .catch((e) => ({ data: [] as TeamOption[], error: e instanceof Error ? e.message : String(e) })),
        // cr301_salesteamname doesn't actually exist server-side (confirmed live —
        // see TEAM_VALUE_FIELD's comment) despite the generated model claiming it
        // does; only the plain _value field is selected, and the team's display
        // name is resolved below via teamMap, the same way userMap resolves the
        // Agent User's name.
        Cr301_agentsService.getAll({
          select: ['cr301_agentid', 'cr301_newcolumn', '_cr301_agent_value', 'cr18c_area', TEAM_VALUE_FIELD, EMPLOYEE_VALUE_FIELD],
          filter: COLLECTION_FILTER,
          orderBy: ['createdon desc'],
        }),
      ]);
      if (!result.success) {
        throw new Error(operationErrorMessage(result.error, 'Could not load agents.'));
      }
      if (userSettled.error) setUserLoadError(userSettled.error);
      if (employeeSettled.error) setEmployeeLoadError(employeeSettled.error);
      if (teamSettled.error) setTeamLoadError(teamSettled.error);

      const userOptions = userSettled.data;
      const employeeOptions = employeeSettled.data;
      const teamOptions = teamSettled.data;
      const userMap = new Map(userOptions.map((u) => [u.id, u]));
      const teamMap = new Map(teamOptions.map((t) => [t.id, t.name]));
      const rows: AgentRow[] = (result.data || []).map((a) => {
        const userId = a._cr301_agent_value || '';
        const u = userId ? userMap.get(userId) : undefined;
        const userFull = u?.displayName || '';
        const login = u ? loginOf(u.domainName) : '';
        const areaValue = a.cr18c_area == null ? null : Number(a.cr18c_area);
        const region: AgentRow['region'] = areaValue == null ? '' : areaValue === AREA_KSA ? 'KSA' : 'EG';
        const teamId = a._cr301_salesteam_value || '';
        return {
          id: a.cr301_agentid,
          name: a.cr301_newcolumn || '—',
          userId,
          user: login || userFull || '—',
          userFull,
          title: u?.title || '',
          region,
          employeeId: a._cfm_employee_value || '',
          teamId,
          teamName: (teamId && teamMap.get(teamId)) || '',
        };
      });
      setUsers(userOptions);
      setEmployees(employeeOptions);
      setTeams(teamOptions);
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
    return regionAgents.filter((a) => `${a.name} ${a.user} ${a.userFull} ${a.title} ${a.teamName}`.toLowerCase().includes(q));
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
    if (id) {
      const a = agents.find((x) => x.id === id);
      const name = a && a.name !== '—' ? a.name : '';
      setEditId(id);
      setFormName(name);
      // Prefer the agent's actual saved cfm_employee link. An agent saved
      // before that binding existed has none — fall back to the same
      // tolerant name match pickEmployee uses, so resaving a legacy agent
      // (even with no other changes) self-heals the missing link instead of
      // silently staying empty forever. If neither resolves, fall back to the
      // agent's own id purely so the picker has a truthy valueId to show
      // `name` via fallbackLabel — handleSave's employees.find lookup never
      // matches this id against a real hr_employees row, so it can't be
      // mistaken for (and bound as) a real employee link.
      let resolvedEmployeeId = a?.employeeId || '';
      if (!resolvedEmployeeId && name) {
        const key = firstWords(name, 3);
        const match = key ? employees.find((e) => firstWords(e.name, 3) === key) : undefined;
        if (match) resolvedEmployeeId = match.id;
      }
      setFormEmployeeId(resolvedEmployeeId || id);
      setFormUserId(a?.userId || '');
      setFormRegion(a?.region || '');
    } else {
      setEditId(null);
      setFormName('');
      setFormEmployeeId('');
      setFormUserId('');
      setFormRegion('');
    }
    setDrawerOpen(true);
  }

  function closeDrawer() {
    setDrawerOpen(false);
  }

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
  function pickEmployee(employee: EmployeeOption | null) {
    if (!employee) {
      setFormName('');
      setFormEmployeeId('');
      return;
    }
    setFormName(employee.name);
    setFormEmployeeId(employee.id);
    const key = firstWords(employee.name, 3);
    const match = key ? users.find((u) => firstWords(u.displayName, 3) === key) : undefined;
    if (match) {
      setFormUserId(match.id);
    } else {
      showToast(`No matching system user found for "${employee.name}" — select one manually.`);
    }
  }

  /**
   * Bind a lookup field via the generic Dataverse client's @odata.bind form.
   * Mixing an "@odata.bind" key into the same body as plain fields and sending
   * it through the generated typed service's create/update has been observed
   * (see Accounts.tsx's setClassificationViaClient) to make the generated
   * service silently drop the *whole* write instead of just rejecting the bind
   * key — so every lookup here is set separately, through the raw client,
   * after the plain fields already landed. `ok: true` means the call itself
   * succeeded (not proof the value stuck — verify separately); `message` is
   * only populated on failure, surfaced in the save toast so a real cause
   * (bad schema name, field security, a plugin) is visible instead of guessed at.
   */
  async function bindLookupViaClient(
    id: string,
    navProperty: string,
    targetEntitySet: string,
    targetId: string
  ): Promise<{ ok: boolean; message?: string }> {
    try {
      const client = getClient(dataSourcesInfo);
      const res = await client.updateRecordAsync(AGENTS_TABLE, id, {
        [`${navProperty}@odata.bind`]: `/${targetEntitySet}(${targetId})`,
      });
      if (res.success) return { ok: true };
      return { ok: false, message: operationErrorMessage(res.error, 'unknown error') };
    } catch (e) {
      return { ok: false, message: e instanceof Error ? e.message : String(e) };
    }
  }

  /** Re-read one agent's lookup value field to confirm a write actually persisted. */
  async function readLookupValue(id: string, valueField: string): Promise<string> {
    try {
      const res = await Cr301_agentsService.get(id, { select: [valueField] });
      if (res.success && res.data) {
        return ((res.data as unknown as Record<string, unknown>)[valueField] as string) || '';
      }
    } catch {
      /* ignore */
    }
    return '';
  }

  async function handleSave() {
    const name = formName.trim();
    // Only enforced when creating: an existing agent may be a legacy record
    // with no employee match found (openDrawer's fallback came up empty), and
    // that must not block re-saving its other fields.
    if (!editId && !formEmployeeId) {
      setFormError('Select an employee from the list — typed text alone can\'t be saved.');
      return;
    }
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
      //    see bindLookupViaClient's comment for why lookups are kept out of
      //    this call.
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

      // 2) Set lookups separately, then verify each actually persisted before
      //    reporting success — a 200/204 isn't proof the value stuck.
      if (recordId) {
        const userBind = await bindLookupViaClient(recordId, AGENT_NAV, USERS_ENTITY_SET, formUserId);
        const userStored = userBind.ok ? await readLookupValue(recordId, AGENT_VALUE_FIELD) : '';
        const userFailed = userStored !== formUserId;

        // Sales Team is always "Collection Team" now — only attempt (and verify)
        // the bind if that team actually resolved from cr301_salesteams.
        let teamFailed = false;
        let teamMessage = '';
        if (collectionTeam) {
          const teamBind = await bindLookupViaClient(recordId, TEAM_NAV, TEAMS_ENTITY_SET, collectionTeam.id);
          const teamStored = teamBind.ok ? await readLookupValue(recordId, TEAM_VALUE_FIELD) : '';
          teamFailed = teamStored !== collectionTeam.id;
          if (!teamBind.ok) teamMessage = teamBind.message || '';
        }

        // Only bind Employee when formEmployeeId is a real pick from the loaded
        // `employees` list — editing an existing agent can seed formEmployeeId
        // with the agent's own id as a display-only fallback (see openDrawer's
        // comment) when no real link or name match was found, and that must
        // never be sent to Dataverse as an hr_employees id.
        const pickedEmployee = employees.find((e) => e.id === formEmployeeId);
        let employeeFailed = false;
        let employeeMessage = '';
        if (pickedEmployee) {
          const employeeBind = await bindLookupViaClient(recordId, EMPLOYEE_NAV, EMPLOYEES_ENTITY_SET, pickedEmployee.id);
          const employeeStored = employeeBind.ok ? await readLookupValue(recordId, EMPLOYEE_VALUE_FIELD) : '';
          employeeFailed = employeeStored !== pickedEmployee.id;
          // A successful bind call whose read-back still doesn't match usually
          // means field-level security silently dropped it server-side — the
          // API returns success either way, so that case gets its own note.
          employeeMessage = !employeeBind.ok ? employeeBind.message || '' : employeeFailed ? 'write accepted but did not stick — check field security' : '';
        }

        if (userFailed || teamFailed || employeeFailed) {
          const failedFields = [
            userFailed && 'Agent User',
            teamFailed && `Sales Team${teamMessage ? ` (${teamMessage})` : ''}`,
            employeeFailed && `Employee${employeeMessage ? ` (${employeeMessage})` : ''}`,
          ]
            .filter(Boolean)
            .join(' and ');
          setDrawerOpen(false);
          showToast(`${editId ? 'Agent saved' : 'Agent created'}, but ${failedFields} could not be set`, 8000);
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
    const header = ['Agent Name', 'Agent User', 'Title', 'Sales Team', 'Region'];
    const rows = view.map((a) => [a.name, a.user, a.title, a.teamName || '—', a.region || '—']);
    downloadCsv('agents.csv', [header, ...rows].map((r) => r.map(escapeForCsv).join(',')));
  }

  /**
   * Bulk backfill for every agent missing cfm_employee (the common case right
   * after this link started being saved at all — see handleSave/openDrawer).
   * Only binds a name match confident enough to match on 3 words (same bar
   * pickEmployee/openDrawer use elsewhere) — anything less confident is left
   * alone and counted as unmatched rather than guessed at, so this can't
   * mis-link an agent to the wrong employee.
   */
  async function linkMissingEmployees() {
    setLinkingEmployees(true);
    try {
      const targets = agents.filter((a) => !a.employeeId && a.name && a.name !== '—');
      let linked = 0;
      let unmatched = 0;
      let failed = 0;
      for (const a of targets) {
        const key = firstWords(a.name, 3);
        const match = key ? employees.find((e) => firstWords(e.name, 3) === key) : undefined;
        if (!match) {
          unmatched++;
          continue;
        }
        const bind = await bindLookupViaClient(a.id, EMPLOYEE_NAV, EMPLOYEES_ENTITY_SET, match.id);
        const stored = bind.ok ? await readLookupValue(a.id, EMPLOYEE_VALUE_FIELD) : '';
        if (stored === match.id) linked++;
        else failed++;
      }
      const parts = [
        `Linked ${linked} agent${linked === 1 ? '' : 's'}`,
        unmatched ? `${unmatched} had no confident name match (link manually)` : '',
        failed ? `${failed} failed to save` : '',
      ].filter(Boolean);
      showToast(parts.join(' — '), 8000);
      await load();
    } catch (e) {
      showToast(e instanceof Error ? e.message : String(e), 6000);
    } finally {
      setLinkingEmployees(false);
    }
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
            <button className="acc-btn" disabled={linkingEmployees} onClick={linkMissingEmployees} title="Backfill the Employee link for agents that don't have one yet, by confident name match">
              <Icon className={`fa-solid ${linkingEmployees ? 'fa-spinner fa-spin' : 'fa-link'}`} /> {linkingEmployees ? 'Linking…' : 'Link Employees'}
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
              <input placeholder="Search agent, user, title or team…" value={query} onChange={(e) => setQuery(e.target.value)} />
            </div>
          </div>

          <div className="acc-tablewrap">
          <table>
            <thead>
              <tr>
                <th>Agent Name</th>
                <th>Agent User</th>
                <th>Title</th>
                <th>Sales Team</th>
                <th>Region</th>
              </tr>
            </thead>
            <tbody>
              {loading ? (
                Array.from({ length: 6 }).map((_, i) => (
                  <tr key={i}>
                    {Array.from({ length: 5 }).map((__, j) => (
                      <td key={j}>
                        <div className="acc-skel"></div>
                      </td>
                    ))}
                  </tr>
                ))
              ) : loadError ? (
                <tr>
                  <td colSpan={5}>
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
                  <td colSpan={5}>
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
                    <td className="agt-title-cell">{a.teamName ? a.teamName : <span className="agt-muted">—</span>}</td>
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
          <div className="acc-fld" style={{ marginBottom: 16 }}>
            <label>
              Employee <span className="dm-req">*</span>
            </label>
            <SearchPicker
              options={employees}
              valueId={formEmployeeId}
              fallbackLabel={formName}
              loadErrorText={employeeLoadError ? `Could not load employees — ${employeeLoadError}` : undefined}
              listboxId="agt-employee-listbox"
              placeholder="Search employee by name…"
              getPrimary={(e) => e.name}
              getFields={(e) => [{ text: e.name, weight: 100 }]}
              onSelect={pickEmployee}
            />
            <div className="acc-hint">
              {employeeLoadError
                ? `Could not load employees — ${employeeLoadError}`
                : `Search ${employees.length.toLocaleString()} active employees by name — ↑/↓ to move, Enter to pick. Sets the Agent Name and auto-selects the matching Agent User below when found.`}
            </div>
            {editId && !selectedEmployee && (
              <div style={{ background: 'var(--warnbg)', color: 'var(--warn)', padding: '9px 12px', borderRadius: 8, marginTop: 8, fontSize: 12 }}>
                <Icon className="fa-solid fa-triangle-exclamation" /> Not linked to a real Employee record yet — the name shown is text-only.
                Search and click it above (even if it looks the same) to actually link and save it.
              </div>
            )}
          </div>
          <div className="acc-fld">
            <div className="acc-ro-wrap">
              <div className="acc-ro-label">HR Code</div>
              <div className="acc-ro-value">{selectedEmployee?.hrCode || '—'}</div>
            </div>
          </div>
          <div className="acc-fld">
            <label>
              Agent User (System User) <span className="dm-req">*</span>
            </label>
            <SearchPicker
              options={users}
              valueId={formUserId}
              disabled={!editId && !formEmployeeId}
              disabledHint="Select an employee above first."
              loadErrorText={userLoadError ? `Could not load system users — ${userLoadError}` : undefined}
              listboxId="agt-user-listbox"
              placeholder="Search by name, login, email or title…"
              getPrimary={(u) => u.displayName}
              getSecondary={(u) => {
                const login = loginOf(u.domainName);
                return u.title ? `${login || '—'} · ${u.title}` : login;
              }}
              getTag={(u) => (u.isApp ? 'app' : undefined)}
              getFields={(u) => [
                { text: u.displayName, weight: 100 },
                { text: loginOf(u.domainName), weight: 60 },
                { text: u.email, weight: 60 },
                { text: u.title, weight: 15 },
              ]}
              getPenalty={(u) => (u.isApp ? 500 : 0)}
              onSelect={(u) => setFormUserId(u ? u.id : '')}
            />
            <div className="acc-hint">
              {userLoadError
                ? `Could not load system users — ${userLoadError}`
                : `Search ${users.length.toLocaleString()} active users by name, login, email or title — ↑/↓ to move, Enter to pick. Title and team are read from the linked system user.`}
            </div>
          </div>
          <div className="acc-fld">
            <label>Sales Team</label>
            <div className="acc-ro-wrap">
              <div className="acc-ro-label">Fixed — every collection agent belongs to it</div>
              <div className="acc-ro-value">
                {teamLoadError ? `Could not load — ${teamLoadError}` : collectionTeam ? collectionTeam.name : teams.length > 0 ? 'No "Collection" team found.' : 'Loading…'}
              </div>
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
