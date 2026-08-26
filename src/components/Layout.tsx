import { useState, type ReactNode } from 'react';
import { useTheme } from '../lib/useTheme';
import type { Region } from '../lib/useRegion';
import { isPageEnabled } from '../lib/featureFlags';
import './Layout.css';

export type PageKey =
  | 'collection-plan'
  | 'weekly-targets'
  | 'daily-collection'
  | 'dispute-management'
  | 'high-risk-companies'
  | 'sla'
  | 'accounts'
  | 'agents'
  | 'visits'
  | 'visit-plan';

interface NavItem {
  key: PageKey;
  icon: string;
  label: string;
}

interface NavGroup {
  title: string;
  items: NavItem[];
}

// Full set of pages — every page's route/component still exists regardless of the
// feature flags below. Each group's items are filtered by isPageEnabled() before
// rendering, so a disabled page just doesn't show a sidebar link; nothing here is
// removed. A group with zero visible items renders nothing.
const NAV_GROUPS: NavGroup[] = [
  {
    title: 'Collection Plan',
    items: [
      { key: 'collection-plan', icon: 'fa-chart-line', label: 'Collection Plan' },
      { key: 'daily-collection', icon: 'fa-hand-holding-dollar', label: 'Daily Collection' },
      { key: 'weekly-targets', icon: 'fa-calendar-week', label: 'Weekly Targets & Ach' },
    ],
  },
  {
    title: 'Visits',
    items: [
      { key: 'visit-plan', icon: 'fa-bullseye', label: 'Visits — Planned & Actual' },
      { key: 'visits', icon: 'fa-location-dot', label: 'Visits Log' },
      { key: 'accounts', icon: 'fa-building-shield', label: 'Accounts' },
      { key: 'agents', icon: 'fa-user-tie', label: 'Agents' },
    ],
  },
  {
    title: 'Management',
    items: [
      { key: 'dispute-management', icon: 'fa-scale-balanced', label: 'Dispute Management' },
      { key: 'sla', icon: 'fa-file-contract', label: 'SLA' },
      { key: 'high-risk-companies', icon: 'fa-triangle-exclamation', label: 'High Risk Companies' },
    ],
  },
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
  // Groups default to expanded — a title only lands in this set once the user
  // collapses it.
  const [collapsedGroups, setCollapsedGroups] = useState<Set<string>>(new Set());

  function toggleGroup(title: string) {
    setCollapsedGroups((prev) => {
      const next = new Set(prev);
      if (next.has(title)) next.delete(title);
      else next.add(title);
      return next;
    });
  }
  const [regionModalOpen, setRegionModalOpen] = useState(() => !region);

  function chooseRegion(code: Region) {
    onRegionChange(code);
    setRegionModalOpen(false);
  }

  function navigate(page: PageKey) {
    onNavigate(page);
    setMobileOpen(false);
  }

  // Below 760px the sidebar is an off-canvas overlay (opened via the floating
  // hamburger button) and the CSS forces it back to full width regardless of
  // `collapsed` — so toggling `collapsed` there is a no-op the user can't see.
  // Close the overlay instead; only actually collapse-to-rail on desktop widths.
  function toggleSidebarButton() {
    if (window.matchMedia('(max-width: 760px)').matches) {
      setMobileOpen(false);
    } else {
      setCollapsed((v) => !v);
    }
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
          <button className="acc-sb-toggle" title={collapsed ? 'Expand sidebar' : 'Collapse sidebar'} onClick={toggleSidebarButton}>
            <i className={`fa-solid ${collapsed ? 'fa-angles-right' : 'fa-bars'}`}></i>
          </button>
        </div>
        <nav className="acc-sb-nav">
          {NAV_GROUPS.map((group) => {
            const items = group.items.filter((item) => isPageEnabled(item.key));
            if (items.length === 0) return null;
            const isCollapsed = collapsedGroups.has(group.title);
            return (
              <div className="acc-sb-group" key={group.title}>
                <button type="button" className="acc-sb-group-hdr" onClick={() => toggleGroup(group.title)}>
                  <span className="acc-sb-label">{group.title}</span>
                  <i className={`fa-solid fa-chevron-down acc-sb-group-caret${isCollapsed ? '' : ' open'}`}></i>
                </button>
                {/* Icon-rail mode hides group headers entirely (no room, no toggle
                    affordance) — always show items there regardless of a group's
                    own collapsed state, so it never strands hidden icons. */}
                {(!isCollapsed || collapsed) && (
                  <ul className="acc-sb-group-items">
                    {items.map((item) => (
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
                )}
              </div>
            );
          })}
        </nav>
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
        <div className="region-scrim" onClick={() => setRegionModalOpen(false)}>
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
            <button className="region-skip" onClick={() => setRegionModalOpen(false)}>
              Skip — show everything, no region filter
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
