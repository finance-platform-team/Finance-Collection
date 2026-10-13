import { Cr301_accountsService } from '../generated/services/Cr301_accountsService';
import { Strategy_kpisesService } from '../generated/services/Strategy_kpisesService';
import { Pm_kpiachievmentsService } from '../generated/services/Pm_kpiachievmentsService';
import { Stf_kpiachievmentbreakdownsService } from '../generated/services/Stf_kpiachievmentbreakdownsService';
import { Cfm_collectionplanweeksService } from '../generated/services/Cfm_collectionplanweeksService';
import { Cr301_agentsService } from '../generated/services/Cr301_agentsService';
import type { Region } from './useRegion';
import { fetchAll } from './dataverse';

/* Shared by Accounts.tsx and VisitPlan.tsx to flag when the company-wide
 * "Target Visit" total on Accounts drifts from Visit Plan's current-month
 * "Monthly planned" total — both region-scoped the same way each page already
 * scopes its own content, so the flag stays consistent no matter which page
 * shows it. Visit Plan's side is scoped to the current calendar month only
 * (not every achievement ever created), since Accounts' Target Visit field
 * isn't itself month-specific. */

const TYPE_INSURANCE = 983080002;
const TYPE_DIRECT = 766340001;
const ACCOUNT_REGION_EGYPT = 983080000; // cr301_country's Egypt value
const ACCOUNT_REGION_KSA = 983080001; // cr301_country's KSA value

const DEFAULT_KPI_NAME = 'No. of Visits';
// strategy_kpises can hold more than one KPI named "No. of Visits" across
// different Function/Department combinations — pin down the one VisitPlan.tsx
// means instead of taking whichever same-named record the server returns first.
// Function/Department are lookups (strategy_function / strategy_department) —
// Dataverse rejects filtering by a plain "...name" property on them ("Could
// not find a property named 'strategy_functionname'"), so the display text is
// read client-side from the @OData FormattedValue annotation instead (see
// formattedValue() below), same as everywhere else in this codebase.
const DEFAULT_KPI_FUNCTION = 'Collection/Finance';
const DEFAULT_KPI_DEPARTMENT = 'Finance';
// Same cr18c_area KSA value / "unset counts as Egypt" convention as
// Agents.tsx / VisitPlan.tsx.
const AREA_KSA = 989230004;

function matchesRegion(region: Region | null, isKsa: boolean): boolean {
  if (!region) return true;
  return region === 'KSA' ? isKsa : !isKsa;
}

function formattedValue(record: unknown, attribute: string): string {
  const raw = (record as Record<string, unknown>)[`${attribute}@OData.Community.Display.V1.FormattedValue`];
  return typeof raw === 'string' ? raw : '';
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
  // Explicit match only, same as Accounts.tsx's own Region filter — an account
  // with no cr301_country set (or some other legacy value) counts under
  // neither Egypt nor KSA here, unlike matchesRegion's "not KSA = Egypt"
  // convention used everywhere else in this file (agent-side matching).
  if (!region) return rows.reduce((s, a) => s + (a.cfm_visittargetcollection || 0), 0);
  const wantKsa = region === 'KSA';
  return rows
    .filter((a) => a.cr301_country === (wantKsa ? ACCOUNT_REGION_KSA : ACCOUNT_REGION_EGYPT))
    .reduce((s, a) => s + (a.cfm_visittargetcollection || 0), 0);
}

export async function fetchVisitPlanTargetSum(region: Region | null): Promise<number> {
  // orderBy must match VisitPlan.tsx's own KPI query exactly: if a duplicate
  // "No. of Visits" (Collection/Finance) strategy_kpises row exists — the create
  // wizard has been run twice before — an unordered query can resolve a
  // *different* duplicate than VisitPlan.tsx does, silently scoping this total
  // to the wrong KPI's (and so a different set of) achievements/breakdowns.
  const kpiRes = await Strategy_kpisesService.getAll({
    select: ['strategy_kpisid', '_strategy_function_value', '_strategy_department_value'],
    filter: `strategy_newcolumn eq '${DEFAULT_KPI_NAME}' and statecode eq 0`,
    orderBy: ['createdon desc'],
  });
  if (!kpiRes.success) throw new Error('Could not load KPIs.');
  const kpi = (kpiRes.data || []).find(
    (k) =>
      formattedValue(k, '_strategy_function_value').trim().toLowerCase() === DEFAULT_KPI_FUNCTION.toLowerCase() &&
      formattedValue(k, '_strategy_department_value').trim().toLowerCase() === DEFAULT_KPI_DEPARTMENT.toLowerCase()
  );
  if (!kpi) return 0;

  // Current calendar month/year only — same local-date convention Visits.tsx's
  // adjustVisitPlanActual uses — not every achievement ever created, so a past
  // month left active doesn't inflate this against Accounts' single, month-less
  // Target Visit field.
  const now = new Date();
  const month = now.getMonth() + 1;
  const year = now.getFullYear();

  const [achRows, agentRows] = await Promise.all([
    fetchAll<{ pm_kpiachievmentid: string }>(
      (skipToken) =>
        Pm_kpiachievmentsService.getAll({
          select: ['pm_kpiachievmentid'],
          filter: `_pm_kpi_value eq ${kpi.strategy_kpisid} and pm_month eq ${month} and pm_year eq ${year} and statecode eq 0`,
          maxPageSize: 5000,
          skipToken,
        }),
      'Could not load KPI achievements.'
    ),
    fetchAll<{ cr301_agentid: string; cr18c_area?: number; _cfm_employee_value?: string }>(
      (skipToken) =>
        Cr301_agentsService.getAll({
          select: ['cr301_agentid', 'cr18c_area', '_cfm_employee_value'],
          filter: 'cfm_iscollection eq true',
          maxPageSize: 5000,
          skipToken,
        }),
      'Could not load agents.'
    ),
  ]);
  if (achRows.length === 0) return 0;

  // Region-by-employee via the real cfm_employee lookup (a stable id) — no more
  // falling back to a fuzzy name match against hr_employees, same call VisitPlan.tsx
  // and Visits.tsx already made: an agent not yet linked just resolves no region
  // here, rather than risking a match against the wrong same-named person.
  const agentByEmployeeId = new Map<string, boolean>();
  agentRows.forEach((a) => {
    if (a.cr18c_area == null || !a._cfm_employee_value) return;
    agentByEmployeeId.set(a._cfm_employee_value, Number(a.cr18c_area) === AREA_KSA);
  });
  function isKsaFor(employeeId: string): boolean | undefined {
    return agentByEmployeeId.get(employeeId);
  }

  const achIdFilter = achRows.map((a) => `_stf_total_value eq ${a.pm_kpiachievmentid}`).join(' or ');
  const brkRows = await fetchAll<{ stf_kpiachievmentbreakdownid: string; _stf_employee_value?: string; statecode?: number }>(
    (skipToken) =>
      Stf_kpiachievmentbreakdownsService.getAll({
        select: ['stf_kpiachievmentbreakdownid', '_stf_employee_value', 'statecode'],
        filter: `(${achIdFilter}) and statecode eq 0`,
        maxPageSize: 5000,
        skipToken,
      }),
    'Could not load KPI achievement breakdowns.'
  );

  const regionBrkIds = brkRows
    .filter((b) => matchesRegion(region, isKsaFor(b._stf_employee_value || '') === true))
    .map((b) => b.stf_kpiachievmentbreakdownid);
  if (!regionBrkIds.length) return 0;

  // "Monthly planned" the same way VisitPlan.tsx's own summary tile computes it
  // (ownerTotals().planned) — the sum of each week's cfm_plan, not the
  // breakdown-level comp_breakdowntarget field, so this figure always agrees
  // with what that page shows for the same month/region.
  const weekIdFilter = regionBrkIds.map((id) => `_cfm_kpiachievementbreakdown_value eq ${id}`).join(' or ');
  const weekRows = await fetchAll<{ cfm_plan?: number }>(
    (skipToken) =>
      Cfm_collectionplanweeksService.getAll({
        select: ['cfm_plan'],
        filter: `(${weekIdFilter}) and statecode eq 0`,
        maxPageSize: 5000,
        skipToken,
      }),
    'Could not load collection plan weeks.'
  );
  return weekRows.reduce((s, w) => s + (w.cfm_plan || 0), 0);
}
