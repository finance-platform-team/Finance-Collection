import { Cr301_accountsService } from '../generated/services/Cr301_accountsService';
import { Strategy_kpisesService } from '../generated/services/Strategy_kpisesService';
import { Pm_kpiachievmentsService } from '../generated/services/Pm_kpiachievmentsService';
import { Stf_kpiachievmentbreakdownsService } from '../generated/services/Stf_kpiachievmentbreakdownsService';
import { Cr301_agentsService } from '../generated/services/Cr301_agentsService';
import { Hr_employeesService } from '../generated/services/Hr_employeesService';
import type { Region } from './useRegion';
import { fetchAll } from './dataverse';

// Same tolerant name match Agents.tsx / Visits.tsx / VisitPlan.tsx use for Agent <-> Employee.
function firstWords(name: string, n: number): string {
  return name.trim().toLowerCase().split(/\s+/).slice(0, n).join(' ');
}

/* Shared by Accounts.tsx and VisitPlan.tsx to flag when the company-wide
 * "Target Visit" total on Accounts drifts from the total planned Target on
 * Visit Plan — both region-scoped the same way each page already scopes its
 * own content, so the flag stays consistent no matter which page shows it. */

const TYPE_INSURANCE = 983080002;
const TYPE_DIRECT = 766340001;
const ACCOUNT_REGION_KSA = 983080001; // cr301_country's KSA value

const DEFAULT_KPI_NAME = 'No. of Visits';
// Same cr18c_area KSA value / "unset counts as Egypt" convention as
// Agents.tsx / VisitPlan.tsx.
const AREA_KSA = 989230004;

function matchesRegion(region: Region | null, isKsa: boolean): boolean {
  if (!region) return true;
  return region === 'KSA' ? isKsa : !isKsa;
}

export async function fetchAccountTargetSum(region: Region | null): Promise<number> {
  const rows = await fetchAll<{ cr301_country?: number; cfm_visittargetcollection?: number }>(
    (skipToken) =>
      Cr301_accountsService.getAll({
        select: ['cr301_accountid', 'cr301_country', 'cfm_visittargetcollection'],
        filter: `cr301_accounttype eq ${TYPE_INSURANCE} or cr301_accounttype eq ${TYPE_DIRECT}`,
        maxPageSize: 5000,
        skipToken,
      }),
    'Could not load account targets.'
  );
  return rows.filter((a) => matchesRegion(region, a.cr301_country === ACCOUNT_REGION_KSA)).reduce((s, a) => s + (a.cfm_visittargetcollection || 0), 0);
}

export async function fetchVisitPlanTargetSum(region: Region | null): Promise<number> {
  const kpiRes = await Strategy_kpisesService.getAll({
    select: ['strategy_kpisid', 'strategy_newcolumn'],
    filter: 'statecode eq 0',
    maxPageSize: 5000,
  });
  if (!kpiRes.success) throw new Error('Could not load KPIs.');
  const kpi = (kpiRes.data || []).find((k) => (k.strategy_newcolumn || '').trim().toLowerCase() === DEFAULT_KPI_NAME.toLowerCase());
  if (!kpi) return 0;

  const [achRows, agentRows, employeeRows] = await Promise.all([
    fetchAll<{ pm_kpiachievmentid: string }>(
      (skipToken) =>
        Pm_kpiachievmentsService.getAll({
          select: ['pm_kpiachievmentid'],
          filter: `_pm_kpi_value eq ${kpi.strategy_kpisid} and statecode eq 0`,
          maxPageSize: 5000,
          skipToken,
        }),
      'Could not load KPI achievements.'
    ),
    fetchAll<{ cr301_agentid: string; cr301_newcolumn?: string; cr18c_area?: number; _cfm_employee_value?: string }>(
      (skipToken) =>
        Cr301_agentsService.getAll({
          select: ['cr301_agentid', 'cr301_newcolumn', 'cr18c_area', '_cfm_employee_value'],
          filter: 'cfm_iscollection eq true',
          maxPageSize: 5000,
          skipToken,
        }),
      'Could not load agents.'
    ),
    fetchAll<{ hr_employeeid: string; hr_fullname?: string }>(
      (skipToken) =>
        Hr_employeesService.getAll({
          select: ['hr_employeeid', 'hr_fullname'],
          filter: 'statecode eq 0',
          maxPageSize: 5000,
          skipToken,
        }),
      'Could not load employees.'
    ),
  ]);
  if (achRows.length === 0) return 0;

  const employeeNameById = new Map(employeeRows.map((e) => [e.hr_employeeid, e.hr_fullname || '']));

  // Region-by-employee — try the real employee lookup first (_cfm_employee_value,
  // when set), falling back to the same tolerant name match VisitPlan.tsx's own
  // load() uses, since Agents.tsx doesn't currently populate that lookup on save.
  const agentByEmployeeId = new Map<string, boolean>();
  const agentByNameKey = new Map<string, boolean>();
  agentRows.forEach((a) => {
    if (a.cr18c_area == null) return;
    const isKsa = Number(a.cr18c_area) === AREA_KSA;
    if (a._cfm_employee_value) agentByEmployeeId.set(a._cfm_employee_value, isKsa);
    const key = firstWords(a.cr301_newcolumn || '', 3);
    if (key) agentByNameKey.set(key, isKsa);
  });
  function isKsaFor(employeeId: string): boolean | undefined {
    if (agentByEmployeeId.has(employeeId)) return agentByEmployeeId.get(employeeId);
    const name = employeeNameById.get(employeeId) || '';
    return agentByNameKey.get(firstWords(name, 3));
  }

  const achIdFilter = achRows.map((a) => `_stf_total_value eq ${a.pm_kpiachievmentid}`).join(' or ');
  const brkRows = await fetchAll<{ comp_breakdowntarget?: number; _stf_employee_value?: string; statecode?: number }>(
    (skipToken) =>
      Stf_kpiachievmentbreakdownsService.getAll({
        select: ['stf_kpiachievmentbreakdownid', 'comp_breakdowntarget', '_stf_employee_value', 'statecode'],
        filter: `(${achIdFilter}) and statecode eq 0`,
        maxPageSize: 5000,
        skipToken,
      }),
    'Could not load KPI achievement breakdowns.'
  );

  return brkRows
    .filter((b) => matchesRegion(region, isKsaFor(b._stf_employee_value || '') === true))
    .reduce((s, b) => s + (b.comp_breakdowntarget || 0), 0);
}
