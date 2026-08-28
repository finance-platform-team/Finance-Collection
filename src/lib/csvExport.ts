/**
 * Shared CSV export helpers. Centralizing these two fixes an Excel mojibake bug that
 * hit every page's export: (1) placeholder tokens like '—' used for "no value" in the
 * UI were leaking into exported cells and, once re-encoded, rendering as garbled bytes
 * (e.g. "â€"") because (2) none of the CSV blobs carried a UTF-8 byte-order-mark, so
 * Excel guessed Windows-1252 and misdecoded both the placeholder dash and any Arabic
 * text alike.
 */
const PLACEHOLDER_TOKENS = new Set(['—', '–', 'N/A', 'n/a']);

export function escapeForCsv(value: string | number | null | undefined): string {
  const raw = value === null || value === undefined ? '' : String(value);
  const cleaned = PLACEHOLDER_TOKENS.has(raw.trim()) ? '' : raw;
  return `"${cleaned.replace(/"/g, '""')}"`;
}

export function downloadCsv(filename: string, lines: string[]): void {
  const blob = new Blob(['﻿' + lines.join('\n')], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}
