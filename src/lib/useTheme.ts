import { useEffect, useState } from 'react';

export type Theme = 'light' | 'dark';

const THEME_KEY = 'app-theme';

function readStoredTheme(): Theme {
  try {
    return window.localStorage.getItem(THEME_KEY) === 'dark' ? 'dark' : 'light';
  } catch {
    return 'light';
  }
}

/**
 * App-wide theme, applied as data-theme on <html>. Putting it on the document root
 * (rather than a page wrapper) means dark-mode CSS reaches everything descended from
 * it — including portal content appended straight to document.body (the searchable
 * user picker) and drawers that are DOM siblings of the page content, not just plain
 * descendants of whichever div happens to call this hook.
 */
export function useTheme(): [Theme, (theme: Theme) => void] {
  const [theme, setTheme] = useState<Theme>(readStoredTheme);

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme);
    try {
      window.localStorage.setItem(THEME_KEY, theme);
    } catch {
      /* private browsing / storage disabled — theme just won't persist */
    }
  }, [theme]);

  return [theme, setTheme];
}
