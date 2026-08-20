import { useState, type ReactNode } from 'react';
import { useTheme } from '../lib/useTheme';
import type { Region } from '../lib/useRegion';
import './Layout.css';

export type PageKey = 'collection-plan' | 'daily-collection' | 'dispute-management' | 'sla' | 'accounts' | 'agents' | 'visits';

const NAV_ITEMS: { key: PageKey; icon: string; label: string }[] = [
  { key: 'collection-plan', icon: 'fa-chart-line', label: 'Collection Plan' },
  { key: 'daily-collection', icon: 'fa-hand-holding-dollar', label: 'Daily Collection' },
  { key: 'dispute-management', icon: 'fa-scale-balanced', label: 'Dispute Management' },
  { key: 'sla', icon: 'fa-file-contract', label: 'SLA' },
  { key: 'accounts', icon: 'fa-building-shield', label: 'Accounts' },
  { key: 'agents', icon: 'fa-user-tie', label: 'Agents' },
  { key: 'visits', icon: 'fa-location-dot', label: 'Visits' },
];

const REGION_OPTIONS: { code: Region; name: string; hint: string }[] = [
  { code: 'EG', name: 'Egypt', hint: 'Everything outside the KSA business units' },
  { code: 'KSA', name: 'Saudi Arabia', hint: 'AFW, AHJ, AKW, ALW, LCH, MKR, ADC, SNB' },
];

interface LayoutProps {
  active: PageKey;
  onNavigate: (page: PageKey) => void;
  region: Region | null;
  onRegionChange: (region: Region) => void;
  children: ReactNode;
}

export default function Layout({ active, onNavigate, region, onRegionChange, children }: LayoutProps) {
  const [collapsed, setCollapsed] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  const [theme, setTheme] = useTheme();
  const [regionModalOpen, setRegionModalOpen] = useState(() => !region);

  function chooseRegion(code: Region) {
    onRegionChange(code);
    setRegionModalOpen(false);
  }

  function navigate(page: PageKey) {
    onNavigate(page);
    setMobileOpen(false);
  }

  return (
    <div className="acc-shell">
      <div className={`acc-scrim-mobile${mobileOpen ? ' on' : ''}`} onClick={() => setMobileOpen(false)}></div>

      <aside className={`acc-sidebar${collapsed ? ' collapsed' : ''}${mobileOpen ? ' mobile-open' : ''}`}>
        <div className="acc-sb-brand">
          <div className="acc-logo">
            <i className="fa-solid fa-network-wired"></i>
            <span>
              Collection
              <br />
              Networking
            </span>
          </div>
          <button className="acc-sb-toggle" title={collapsed ? 'Expand sidebar' : 'Collapse sidebar'} onClick={() => setCollapsed((v) => !v)}>
            <i className={`fa-solid ${collapsed ? 'fa-angles-right' : 'fa-bars'}`}></i>
          </button>
        </div>
        <ul className="acc-sb-nav">
          {NAV_ITEMS.map((item) => (
            <li key={item.key}>
              <a
                href="#"
                className={item.key === active ? 'active' : ''}
                title={item.label}
                onClick={(e) => {
                  e.preventDefault();
                  navigate(item.key);
                }}
              >
                <i className={`fa-solid ${item.icon}`}></i>
                <span className="acc-sb-label">{item.label}</span>
              </a>
            </li>
          ))}
        </ul>
        <div className="acc-sb-foot">
          <button className="acc-sb-region-btn" title="Change region" onClick={() => setRegionModalOpen(true)}>
            <i className="fa-solid fa-earth-americas"></i>
            <span className="acc-sb-label">{region ? `Region: ${region}` : 'Select region'}</span>
          </button>
          <div className="theme-switch" role="group" aria-label="Theme">
            <button className={`theme-btn${theme === 'light' ? ' active' : ''}`} title="Light theme" onClick={() => setTheme('light')}>
              <i className="fa-solid fa-sun"></i>
              <span className="theme-btn-label">Light</span>
            </button>
            <button className={`theme-btn${theme === 'dark' ? ' active' : ''}`} title="Dark theme" onClick={() => setTheme('dark')}>
              <i className="fa-solid fa-moon"></i>
              <span className="theme-btn-label">Dark</span>
            </button>
          </div>
        </div>
      </aside>

      <button className="acc-mobile-menu-btn" title="Open menu" onClick={() => setMobileOpen(true)}>
        <i className="fa-solid fa-bars"></i>
      </button>

      <div className="acc-content-col">
        <main className="acc-main">{children}</main>
      </div>

      {regionModalOpen && (
        <div className="region-scrim" onClick={() => region && setRegionModalOpen(false)}>
          <div className="region-modal" onClick={(e) => e.stopPropagation()}>
            <div className="region-modal-icon">
              <i className="fa-solid fa-earth-americas"></i>
            </div>
            <h2>Select Your Region</h2>
            <p>This sets your default region across the app. You can change it anytime from the sidebar.</p>
            <div className="region-options">
              {REGION_OPTIONS.map((opt) => (
                <button
                  key={opt.code}
                  className={`region-opt${region === opt.code ? ' sel' : ''}`}
                  onClick={() => chooseRegion(opt.code)}
                >
                  <span className="region-opt-code">{opt.code}</span>
                  <span className="region-opt-text">
                    <b>{opt.name}</b>
                    <small>{opt.hint}</small>
                  </span>
                </button>
              ))}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
