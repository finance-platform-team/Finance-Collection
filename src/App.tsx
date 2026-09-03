import { useEffect, useState } from 'react';
import Layout, { type PageKey } from './components/Layout';
import AndalusiaLoader from './components/AndalusiaLoader';
import { useRegion } from './lib/useRegion';
import { DEFAULT_PAGE } from './lib/featureFlags';
import CollectionPlanPage from './pages/CollectionPlan';
import WeeklyTargetsPage from './pages/WeeklyTargets';
import DailyCollectionPage from './pages/DailyCollection';
import DisputeManagementPage from './pages/DisputeManagement';
import HighRiskCompaniesPage from './pages/HighRiskCompanies';
import SlaPage from './pages/Sla';
import AccountsPage from './pages/Accounts';
import AgentsPage from './pages/Agents';
import VisitsPage from './pages/Visits';
import VisitPlanPage from './pages/VisitPlan';
import VisitAuditPage from './pages/VisitAudit';

function App() {
  const [page, setPage] = useState<PageKey>(DEFAULT_PAGE);
  const [region, setRegion] = useRegion();
  const [booting, setBooting] = useState(true);

  useEffect(() => {
    const id = setTimeout(() => setBooting(false), 600);
    return () => clearTimeout(id);
  }, []);

  if (booting) return <AndalusiaLoader />;

  return (
    <Layout active={page} onNavigate={setPage} region={region} onRegionChange={setRegion}>
      {page === 'collection-plan' ? (
        <CollectionPlanPage region={region} />
      ) : page === 'weekly-targets' ? (
        <WeeklyTargetsPage region={region} />
      ) : page === 'daily-collection' ? (
        <DailyCollectionPage region={region} />
      ) : page === 'dispute-management' ? (
        <DisputeManagementPage region={region} />
      ) : page === 'high-risk-companies' ? (
        <HighRiskCompaniesPage region={region} />
      ) : page === 'sla' ? (
        <SlaPage region={region} />
      ) : page === 'accounts' ? (
        <AccountsPage region={region} />
      ) : page === 'agents' ? (
        <AgentsPage region={region} />
      ) : page === 'visit-plan' ? (
        <VisitPlanPage region={region} />
      ) : page === 'visit-audit' ? (
        <VisitAuditPage region={region} />
      ) : (
        <VisitsPage region={region} />
      )}
    </Layout>
  );
}

export default App;
