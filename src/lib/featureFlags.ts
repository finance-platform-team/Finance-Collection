
export const ENABLED_PAGES = [
  'accounts',
  'agents',
  'visits',
  'high-risk-companies',
  'visit-plan',
  'weekly-targets',
  'collection-plan',
  'daily-collection',
  'dispute-management',
  'sla',
] as const;

export function isPageEnabled(key: string): boolean {
  return (ENABLED_PAGES as readonly string[]).includes(key);
}


export const DEFAULT_PAGE = ENABLED_PAGES[0];
