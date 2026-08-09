import { useEffect, useState, type ReactNode } from 'react';
import { getContext } from '@microsoft/power-apps/app';
import './Layout.css';

export type PageKey = 'accounts' | 'agents' | 'visits';

const NAV_ITEMS: { key: PageKey; icon: string; label: string }[] = [
  { key: 'accounts', icon: 'fa-building-shield', label: 'Accounts' },
  { key: 'agents', icon: 'fa-user-tie', label: 'Agents' },
  { key: 'visits', icon: 'fa-location-dot', label: 'Visits' },
];

interface LayoutProps {
  active: PageKey;
  onNavigate: (page: PageKey) => void;
  title: string;
  children: ReactNode;
}

function initialsOf(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return '—';
  return parts.slice(0, 2).map((p) => p[0]?.toUpperCase()).join('');
}

export default function Layout({ active, onNavigate, title, children }: LayoutProps) {
  const [collapsed, setCollapsed] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  const [userName, setUserName] = useState('');

  useEffect(() => {
    let alive = true;
    getContext()
      .then((ctx) => {
        if (alive) setUserName(ctx.user?.fullName || '');
      })
      .catch(() => {});
    return () => {
      alive = false;
    };
  }, []);

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
      </aside>

      <div className="acc-content-col">
        <div className="acc-topbar">
          <div className="acc-tb-left">
            <button className="acc-tb-burger" title="Open menu" onClick={() => setMobileOpen(true)}>
              <i className="fa-solid fa-bars"></i>
            </button>
            <span className="acc-tb-title">{title}</span>
          </div>
          <div className="acc-tb-right">
            <div className="acc-tb-user">
              <b>{userName || '—'}</b>
              <small>Signed in</small>
            </div>
            <div className="acc-tb-avatar">{initialsOf(userName)}</div>
          </div>
        </div>

        <main className="acc-main">{children}</main>
      </div>
    </div>
  );
}
