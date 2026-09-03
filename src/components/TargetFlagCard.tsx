import Icon from './Icon';

interface TargetFlagCardProps {
  loading: boolean;
  accountTarget: number | null;
  visitPlanTarget: number | null;
}

function fmt(n: number): string {
  return Math.round(n || 0).toLocaleString('en-US');
}

/** Company-wide (region-scoped) flag comparing Accounts' total Target Visit
 * against Visit Plan's total planned Target — shown on both Accounts.tsx and
 * VisitPlan.tsx so a mismatch between the two is visible from either page. */
export default function TargetFlagCard({ loading, accountTarget, visitPlanTarget }: TargetFlagCardProps) {
  const ready = !loading && accountTarget != null && visitPlanTarget != null;
  const diff = ready ? visitPlanTarget! - accountTarget! : 0;
  const tone = !ready ? 'gold' : diff === 0 ? 'ok' : 'warn';
  const icon = !ready ? 'fa-spinner fa-spin' : diff === 0 ? 'fa-circle-check' : 'fa-triangle-exclamation';

  let context = 'Loading…';
  if (ready) {
    if (diff === 0) context = 'Matched';
    else if (diff > 0) context = `Visit Plan is ${fmt(diff)} more`;
    else context = `Account target is ${fmt(-diff)} more`;
  }

  return (
    <div className={`kpi-card kpi-${tone}`}>
      <div className="kpi-top">
        <span className={`kpi-icon kpi-icon-${tone}`}>
          <Icon className={`fa-solid ${icon}`} />
        </span>
        <span className="kpi-label">Account vs Visit Plan target</span>
      </div>
      <div className="kpi-value">{ready ? `${fmt(accountTarget!)} / ${fmt(visitPlanTarget!)}` : '—'}</div>
      <div className="kpi-context">{context}</div>
    </div>
  );
}
