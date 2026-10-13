import Icon from './Icon';

interface TargetFlagCardProps {
  loading: boolean;
  accountTarget: number | null;
  visitPlanTarget: number | null;
}

function fmt(n: number): string {
  return Math.round(n || 0).toLocaleString('en-US');
}

/** Company-wide (region-scoped) alert banner comparing Accounts' total Target
 * Visit against Visit Plan's total planned Target. Silent when the two match
 * (or while still loading) — it's an alert for a real mismatch, not a
 * permanent stat tile, so it only takes up space when there's something to
 * flag. */
export default function TargetFlagCard({ loading, accountTarget, visitPlanTarget }: TargetFlagCardProps) {
  if (loading || accountTarget == null || visitPlanTarget == null) return null;
  const diff = visitPlanTarget - accountTarget;
  if (diff === 0) return null;

  const detail =
    diff > 0
      ? `Visit Plan totals ${fmt(visitPlanTarget)}, Accounts totals ${fmt(accountTarget)} — Visit Plan is ${fmt(diff)} more.`
      : `Accounts totals ${fmt(accountTarget)}, Visit Plan totals ${fmt(visitPlanTarget)} — Accounts is ${fmt(-diff)} more.`;

  return (
    <div className="tf-banner tf-banner-warn">
      <span className="tf-banner-icon tf-banner-icon-warn">
        <Icon className="fa-solid fa-triangle-exclamation" />
      </span>
      <div className="tf-banner-body">
        <div className="tf-banner-title">Target mismatch between Accounts and Visit Plan</div>
        <div className="tf-banner-detail">{detail}</div>
      </div>
    </div>
  );
}
