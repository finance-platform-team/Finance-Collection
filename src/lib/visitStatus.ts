/* Shared styling for cr301_visitstatus (Open / Completed / Cancelled / Partially
 * Closed) — used by both the Visits Log page and the Accounts page's per-account
 * visits panel. */
export function visitStatusClass(label: string): string {
  if (label === 'Completed') return 'acc-visit-status-ok';
  if (label === 'Cancelled') return 'acc-visit-status-bad';
  if (label === 'Open') return 'acc-visit-status-warn';
  if (label === 'Partially Closed') return 'acc-visit-status-info';
  return 'acc-visit-status-none';
}
