import { useEffect, useState } from 'react';
import brandLogo from '../assets/andalusia-logo-clean.png';
import { readStoredTheme } from '../lib/useTheme';

/**
 * Self-contained branded loading screen. Theme resolution for 'auto' reuses the
 * app's actual theme mechanism — useTheme.ts sets data-theme="light"|"dark" on
 * <html> (this Code App has no Fluent UI theme context; that's a plain custom
 * hook + localStorage) — a MutationObserver keeps it live if the user flips
 * the theme while this is showing.
 *
 * This loader renders before Layout (and therefore before useTheme's own effect)
 * ever mounts — App.tsx shows it during the boot gate, ahead of <Layout>, so
 * <html data-theme> hasn't been stamped yet on first paint. Falling back to
 * useTheme's own persisted preference (same localStorage key) means "auto"
 * still resolves correctly on that very first render instead of defaulting to
 * light until Layout mounts a moment later.
 */

const GOLD = '#a5845b';
const DARK_GREEN = '#2e372d';
const CANVAS = '#f4f1ea';

const PALETTE = {
  light: {
    bg: CANVAS,
    title: DARK_GREEN,
    subtitle: 'rgba(46,55,45,0.55)',
    status: 'rgba(46,55,45,0.50)',
    track: 'rgba(46,55,45,0.10)',
    ringBg: 'rgba(46,55,45,0.10)',
  },
  dark: {
    bg: DARK_GREEN,
    title: CANVAS,
    subtitle: 'rgba(244,241,234,0.50)',
    status: 'rgba(244,241,234,0.40)',
    track: 'rgba(244,241,234,0.10)',
    ringBg: 'rgba(244,241,234,0.08)',
  },
} as const;

const STATUS_MESSAGES = ['Connecting to Dataverse…', 'Authenticating session…', 'Loading collection data…', 'Preparing your workspace…'];

const STYLES = `
.al-root{transition:background .35s ease}
.al-grid{position:absolute;inset:0;pointer-events:none}
.al-ring{transform-origin:75px 75px;animation:al-spin 1.5s cubic-bezier(0.6,0.1,0.3,0.9) infinite}
.al-logo-wrap{animation:al-pulse 2s ease-in-out infinite}
.al-progress-fill{animation:al-progress 1.6s ease-in-out infinite}
@keyframes al-spin{to{transform:rotate(360deg)}}
@keyframes al-pulse{0%,100%{transform:scale(0.95);opacity:.85}50%{transform:scale(1.03);opacity:1}}
@keyframes al-progress{0%{left:-40%}100%{left:100%}}
@media (prefers-reduced-motion: reduce){
  .al-ring{animation:none!important}
  .al-logo-wrap{animation:none!important;transform:scale(1);opacity:1}
  .al-progress-fill{animation:none!important;left:0!important;width:100%!important}
}
`;

function resolveTheme(theme: 'light' | 'dark' | 'auto'): 'light' | 'dark' {
  if (theme === 'light' || theme === 'dark') return theme;
  if (typeof document === 'undefined') return 'light';
  const attr = document.documentElement.getAttribute('data-theme');
  if (attr === 'dark' || attr === 'light') return attr;
  return readStoredTheme();
}

function useResolvedTheme(theme: 'light' | 'dark' | 'auto'): 'light' | 'dark' {
  const [resolved, setResolved] = useState<'light' | 'dark'>(() => resolveTheme(theme));

  useEffect(() => {
    if (theme !== 'auto') {
      setResolved(theme);
      return;
    }
    setResolved(resolveTheme('auto'));
    const observer = new MutationObserver(() => setResolved(resolveTheme('auto')));
    observer.observe(document.documentElement, { attributes: true, attributeFilter: ['data-theme'] });
    return () => observer.disconnect();
  }, [theme]);

  return resolved;
}

interface AndalusiaLoaderProps {
  message?: string;
  theme?: 'light' | 'dark' | 'auto';
}

export default function AndalusiaLoader({ message, theme = 'auto' }: AndalusiaLoaderProps) {
  const resolved = useResolvedTheme(theme);
  const pal = PALETTE[resolved];
  const [msgIndex, setMsgIndex] = useState(0);

  useEffect(() => {
    if (message) return;
    const id = setInterval(() => setMsgIndex((i) => (i + 1) % STATUS_MESSAGES.length), 1400);
    return () => clearInterval(id);
  }, [message]);

  const statusText = message || STATUS_MESSAGES[msgIndex];
  const gridImage = `linear-gradient(${GOLD} .5px, transparent .5px), linear-gradient(90deg, ${GOLD} .5px, transparent .5px)`;

  return (
    <div
      role="status"
      aria-live="polite"
      aria-busy="true"
      aria-label="Loading Andalusia Finance Platform"
      className="al-root"
      style={{
        position: 'fixed',
        inset: 0,
        zIndex: 9999,
        display: 'flex',
        flexDirection: 'column',
        alignItems: 'center',
        justifyContent: 'center',
        overflow: 'hidden',
        background: pal.bg,
      }}
    >
      <style>{STYLES}</style>
      <div className="al-grid" aria-hidden="true" style={{ backgroundImage: gridImage, backgroundSize: '26px 26px', opacity: 0.04 }} />

      <div style={{ position: 'relative', width: 150, height: 150 }}>
        <svg width={150} height={150} viewBox="0 0 150 150" aria-hidden="true" focusable="false" style={{ position: 'absolute', inset: 0 }}>
          <circle cx={75} cy={75} r={66} fill="none" stroke={pal.ringBg} strokeWidth={2.5} />
          <circle
            className="al-ring"
            cx={75}
            cy={75}
            r={66}
            fill="none"
            stroke={GOLD}
            strokeWidth={2.5}
            strokeLinecap="round"
            strokeDasharray="140 275"
          />
        </svg>
        <div
          className="al-logo-wrap"
          style={{ position: 'absolute', inset: 0, display: 'flex', alignItems: 'center', justifyContent: 'center' }}
        >
          <img src={brandLogo} alt="" aria-hidden="true" width={74} style={{ height: 'auto', objectFit: 'contain' }} />
        </div>
      </div>

      <div style={{ textAlign: 'center', marginTop: '1.75rem' }}>
        <div style={{ fontSize: 20, fontWeight: 600, letterSpacing: '0.5px', marginBottom: 5, color: pal.title }}>Andalusia</div>
        <div style={{ fontSize: 12, letterSpacing: '2.5px', textTransform: 'uppercase', color: pal.subtitle }}>Finance Platform</div>
      </div>

      <div style={{ position: 'relative', width: 210, height: 3, borderRadius: 3, overflow: 'hidden', background: pal.track, marginTop: '2rem' }}>
        <div
          className="al-progress-fill"
          style={{ position: 'absolute', top: 0, left: '-40%', height: '100%', width: '40%', borderRadius: 3, background: GOLD }}
        />
      </div>

      <div
        style={{
          fontFamily: 'ui-monospace, SFMono-Regular, Menlo, Consolas, monospace',
          fontSize: 11,
          letterSpacing: '0.5px',
          color: pal.status,
          marginTop: '1rem',
        }}
      >
        {statusText}
      </div>
    </div>
  );
}
