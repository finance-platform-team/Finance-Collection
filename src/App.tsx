import { useState } from 'react';
import Layout, { type PageKey } from './components/Layout';
import { useRegion } from './lib/useRegion';
import CollectionPlanPage from './pages/CollectionPlan';
import DailyCollectionPage from './pages/DailyCollection';
import DisputeManagementPage from './pages/DisputeManagement';
import SlaPage from './pages/Sla';
import AccountsPage from './pages/Accounts';
import AgentsPage from './pages/Agents';
import VisitsPage from './pages/Visits';

function App() {
  const [page, setPage] = useState<PageKey>('collection-plan');
  const [region, setRegion] = useRegion();

  return (
    <Layout active={page} onNavigate={setPage} region={region} onRegionChange={setRegion}>
      {page === 'collection-plan' ? (
        <CollectionPlanPage region={region} />
      ) : page === 'daily-collection' ? (
        <DailyCollectionPage />
      ) : page === 'dispute-management' ? (
        <DisputeManagementPage />
      ) : page === 'sla' ? (
        <SlaPage />
      ) : page === 'accounts' ? (
        <AccountsPage />
      ) : page === 'agents' ? (
        <AgentsPage />
      ) : (
        <VisitsPage />
      )}
    </Layout>
  );
}

export default App;
