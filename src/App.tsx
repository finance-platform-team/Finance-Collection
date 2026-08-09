import { useState } from 'react';
import Layout, { type PageKey } from './components/Layout';
import AccountsPage from './pages/Accounts';
import AgentsPage from './pages/Agents';
import VisitsPage from './pages/Visits';

const PAGE_TITLES: Record<PageKey, string> = {
  accounts: 'Accounts',
  agents: 'Agents',
  visits: 'Visits Log',
};

function App() {
  const [page, setPage] = useState<PageKey>('accounts');

  return (
    <Layout active={page} onNavigate={setPage} title={PAGE_TITLES[page]}>
      {page === 'accounts' ? <AccountsPage /> : page === 'agents' ? <AgentsPage /> : <VisitsPage />}
    </Layout>
  );
}

export default App;
