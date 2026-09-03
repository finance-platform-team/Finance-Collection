import type { HTMLAttributes } from 'react';

// Power Apps code apps enforce a CSP with font-src 'self' by default, and the
// self-hosted Font Awesome webfont (@font-face in all.min.css) gets blocked in
// the deployed app — icons render as empty glyph boxes even though the CSS
// class rules (display, color, fa-spin animation, etc.) all still apply fine.
// Rather than relying on the icon *font* for the glyph shape, this renders the
// same Font Awesome artwork as an inline SVG child of the same <i className=
// "fa-solid fa-xxx">, so every existing CSS selector targeting `i` / `.fa-*`
// keeps working untouched — only the glyph now comes from bundled SVG source,
// which isn't subject to font-src at all.
const solidIcons = import.meta.glob(
  '../../node_modules/@fortawesome/fontawesome-free/svgs/solid/{angles-down,angles-right,angles-up,arrow-up-right-from-square,bars,building,building-columns,building-shield,bullseye,calendar-days,calendar-week,chart-bar,chart-line,chart-simple,check,chevron-down,chevron-right,chevron-up,circle-check,circle-exclamation,circle-info,circle-xmark,clipboard-check,clock,coins,door-open,download,earth-americas,file,file-arrow-down,file-circle-plus,file-contract,file-csv,file-excel,file-export,file-image,file-import,file-pdf,file-powerpoint,file-signature,file-word,file-zipper,filter-circle-xmark,floppy-disk,folder-open,gauge-high,hand-holding-dollar,hourglass-half,inbox,layer-group,link,list-check,location-dot,lock,magnifying-glass,magnifying-glass-chart,money-bill-trend-up,moon,network-wired,paper-plane,paperclip,pen,plus,rotate,rotate-left,sack-dollar,scale-balanced,shield-heart,spinner,square-poll-vertical,sun,table,toggle-off,toggle-on,trash,triangle-exclamation,upload,user,user-slash,user-tag,user-tie,users,xmark}.svg',
  { query: '?raw', import: 'default', eager: true },
) as Record<string, string>;

const iconsByName: Record<string, string> = {};
for (const path in solidIcons) {
  const match = path.match(/\/([a-z0-9-]+)\.svg$/);
  if (match) iconsByName[match[1]] = solidIcons[path];
}

const NON_NAME_CLASSES = new Set([
  'fa-solid', 'fa-regular', 'fa-brands', 'fa-spin',
  'fa-sm', 'fa-lg', 'fa-xs', 'fa-xl', 'fa-2x', 'fa-3x', 'fa-fw',
]);

function extractIconName(className: string): string | undefined {
  for (const cls of className.split(/\s+/)) {
    if (cls.startsWith('fa-') && !NON_NAME_CLASSES.has(cls)) return cls.slice(3);
  }
  return undefined;
}

interface IconProps extends HTMLAttributes<HTMLElement> {
  className: string;
}

// Drop-in replacement for `<i className="fa-solid fa-xxx" ...></i>` — same
// tag, same className/style/title/onClick, just with a real glyph inside.
export default function Icon({ className, ...rest }: IconProps) {
  const name = extractIconName(className);
  const raw = name ? iconsByName[name] : undefined;
  if (!raw) return <i className={className} {...rest} />;

  const viewBox = raw.match(/viewBox="([^"]+)"/)?.[1] ?? '0 0 512 512';
  const inner = raw.replace(/^<svg[^>]*>/, '').replace(/<\/svg>\s*$/, '');

  return (
    <i className={className} {...rest}>
      <svg
        className="fa-svg"
        viewBox={viewBox}
        aria-hidden="true"
        focusable="false"
        dangerouslySetInnerHTML={{ __html: inner }}
      />
    </i>
  );
}
