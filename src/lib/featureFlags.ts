/**
 * Simple static feature flags — flip the value and redeploy (`pac code push`).
 * No runtime toggle, no per-user config; that's intentionally as far as this goes
 * unless a real need for more shows up.
 */

/** When true, the sidebar (Layout.tsx) shows only the "Visits" nav section —
 * Collection Plan and Management are hidden entirely, not just collapsed. */
export const VISITS_ONLY = true;
