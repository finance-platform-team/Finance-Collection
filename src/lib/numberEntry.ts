/** Shared bound for money/amount entry fields app-wide (not count/percent fields, which
 * keep their own smaller, domain-specific caps — e.g. visit targets, week counts). */
export const MAX_ENTRY_VALUE = 100_000_000;

/**
 * Sanitizes a money-field's raw text input as the user types. Values that don't yet
 * parse (empty string, a lone "-", a trailing "12.") pass through untouched so normal
 * typing isn't fought keystroke-by-keystroke; only a value that actually parses out of
 * bounds gets clamped, paired with a message to show the user why.
 */
export function sanitizeAmountInput(raw: string): { value: string; message: string | null } {
  const n = Number(raw);
  if (raw.trim() === '' || !Number.isFinite(n)) return { value: raw, message: null };
  if (n < 0) return { value: '0', message: 'Value cannot be negative.' };
  if (n > MAX_ENTRY_VALUE) return { value: String(MAX_ENTRY_VALUE), message: `Maximum allowed value is ${MAX_ENTRY_VALUE.toLocaleString('en-US')}.` };
  return { value: raw, message: null };
}
